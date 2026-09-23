///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import type { BufferReader, BufferWriter } from "../codec/BufferCursor.js";
import { AuditAction } from "@rapidmx/restapi";
import type { RopContext, RopHandler } from "./RopHandler.js";

const ROP_ID_DELETE_MESSAGES = 0x1e;

/** Audits a message deleted through MAPI with the same `MESSAGE_DELETE` entry the REST message route records.
 * Charged to `context.budget` like the `findOne`/`delete` calls around every call site of this function - a real
 * DB `create()` whenever `context.audit` is configured (the case in every real deployment), so leaving it
 * uncharged undercounted this loop's real DB work by a third. No-op (and no charge - there's no DB work to
 * account for) when `context.audit` is absent, the same as an optional-repo item kind this pragmatic subset
 * already skips charging for. */
export async function auditMessageDelete(context: RopContext, message: { uid: string; mailboxUid: string; subject?: string; folderUid: string }): Promise<void> {
    if (!context.audit) {
        return;
    }
    context.budget?.chargeQueries();
    await context.audit({
        action: AuditAction.MESSAGE_DELETE,
        targetType: "Message",
        targetUid: message.uid,
        mailboxUid: message.mailboxUid,
        details: { subject: message.subject, folderUid: message.folderUid },
    });
}

/** The well-known MAPI HRESULT `MAPI_E_INVALID_OBJECT`, reused for "the referenced handle isn't a folder (or
 * doesn't exist)" - the same constant `RopGetContentsTableHandler` uses for its own analogous check. */
const ERROR_INVALID_OBJECT = 0x80070005;

/**
 * `RopDeleteMessages` (`[MS-OXCMSG]`/`[MS-OXCROPS]`, RopId `0x1E`): deletes one or more messages by `MessageId`
 * from an already-open folder handle. Reuses `RecoverableRepoUtils.delete()` exactly as-built in Phase 2 - the
 * same soft-delete-with-watermark-bump machinery EAS's own `Sync`/`FolderSync` incremental-delete detection
 * already relies on (see `BaseMapiEmsmdbRoute.ts`'s own doc comment for why `folderRepo`/`messageRepo`/
 * `calendarEventRepo` are built as `RecoverableRepoUtils`, not plain `RepoUtils`) - no new deletion-tracking
 * code, zero duplication.
 *
 * A `MessageId` resolves via `session.messageIds` (`MessageTarget.assignOrGetMid`'s own registry, already
 * generic over both `"message:<uid>"` and `"calendarEvent:<uid>"` target strings, since a Calendar item is
 * itself a Message object on the wire - see `CalendarEventTarget.ts`'s own doc comment) directly to the entity
 * to delete, the same "the session already knows what this small integer refers to" design every other MID-
 * consuming ROP in this pragmatic subset already uses (`RopOpenMessage`, `RopSaveChangesMessage`) - the
 * request's own `InputHandleIndex` (nominally the messages' shared parent folder) is validated to be a real,
 * open folder handle but not cross-checked against each message's actual `folderUid`, the same pragmatic
 * simplification `RopOpenMessageHandler`'s own doc comment documents for its analogous `FolderId` field.
 *
 * `WantAsynchronous` (`RopProgress`-reported async completion) and `NotifyNonRead` (non-read-receipt generation
 * to the original sender) are decoded to advance the reader correctly but not honored - this pragmatic subset
 * has no async-ROP or non-read-receipt machinery. `PartialCompletion` is reported `true` if any requested
 * `MessageId` didn't resolve to a real message/calendar-event target - the rest are still deleted, matching
 * this pragmatic subset's "don't fail the whole ROP over one stale reference" principle used throughout.
 *
 * @author Jean-Philippe Steinmetz
 */
export class RopDeleteMessagesHandler implements RopHandler {
    public readonly ropId = ROP_ID_DELETE_MESSAGES;

    public async handle(reader: BufferReader, writer: BufferWriter, context: RopContext): Promise<void> {
        reader.readUInt8(); // LogonId - this pragmatic subset doesn't track multiple concurrent logons per session
        const inputHandleIndex: number = reader.readUInt8();
        reader.readUInt8(); // WantAsynchronous - no RopProgress/async-ROP support in this pragmatic subset
        reader.readUInt8(); // NotifyNonRead - no non-read-receipt generation support
        const messageIdCount: number = reader.readUInt16LE();
        const messageIds: bigint[] = [];
        for (let i = 0; i < messageIdCount; i++) {
            messageIds.push(reader.readBigUInt64LE());
        }

        const handle = context.session.handles[inputHandleIndex];
        if (!handle || handle.type !== "folder") {
            writer.writeUInt8(ROP_ID_DELETE_MESSAGES);
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt32LE(ERROR_INVALID_OBJECT);
            return;
        }

        let partialCompletion = false;
        // The folders that lost a message - refreshed once below, after the loop, rather than once per message (a
        // client can name several `MessageId`s in one call, several of which are often the same folder).
        const affectedFolderUids = new Set<string>();
        for (const messageId of messageIds) {
            const target: string | undefined = context.session.messageIds[messageId.toString()];
            if (target?.startsWith("message:")) {
                const uid = target.slice("message:".length);
                // Always fetched now (not just when `context.audit` is set): `notifyFolderCounts` below needs to know
                // which folder this message was in, the same as `auditMessageDelete` already needed its subject.
                // Charged like `resolveContentsKind`'s own single-row lookup and `RopDeleteFolderHandler.deleteItems()`'s
                // own per-item delete - each is a real DB round trip, so a client can't repeat this ROP with a large
                // `messageIdCount` to run unbounded queries against its own mailbox within one `Execute`.
                context.budget?.chargeQueries();
                const message = await context.messageRepo.findOne(uid, { ignoreACL: true });
                context.budget?.chargeQueries();
                await context.messageRepo.delete(uid, { ignoreACL: true });
                if (message) {
                    affectedFolderUids.add(message.folderUid);
                    await auditMessageDelete(context, message);
                }
            } else if (target?.startsWith("calendarEvent:")) {
                context.budget?.chargeQueries();
                await context.calendarEventRepo.delete(target.slice("calendarEvent:".length), { ignoreACL: true });
            } else {
                partialCompletion = true;
            }
        }
        // Not `bumpSyncKey` - only adding a message to a folder does that (see `refreshFolderCounts()`'s own doc
        // comment); a deletion doesn't, matching restapi's own `ScanQueueJob.processRecall()` deletion call site.
        await context.notifyFolderCounts?.(affectedFolderUids);

        writer.writeUInt8(ROP_ID_DELETE_MESSAGES);
        writer.writeUInt8(inputHandleIndex);
        writer.writeUInt32LE(0); // ReturnValue - success
        writer.writeUInt8(partialCompletion ? 1 : 0); // PartialCompletion
    }
}
