///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import type { RepoUtils } from "@rapidrest/service-core";
import type { BufferReader, BufferWriter } from "../codec/BufferCursor.js";
import { auditMessageDelete } from "./RopDeleteMessagesHandler.js";
import { findAllCapped, findPage, MAX_COLLECTION_ROWS, type RepoSort } from "./RepoPaging.js";
import type { RopContext, RopHandler } from "./RopHandler.js";

const ROP_ID_DELETE_FOLDER = 0x1d;

/** The well-known MAPI HRESULT `MAPI_E_NOT_FOUND`, reused for an unrecognized `FolderId` - the same constant
 * `RopOpenFolderHandler` uses for its own analogous lookup miss. */
const ERROR_NOT_FOUND = 0x8004010f;

/** The well-known MAPI HRESULT `MAPI_E_INVALID_OBJECT`, reused for both "the referenced handle isn't a folder"
 * and "the folder has content/subfolders but the matching cascade flag wasn't set" - the same constant reused
 * throughout this pragmatic subset as a generic "can't do this" signal, not a claim of exact per-condition
 * `[MS-OXCFOLD]` return-value-table parity (the real spec's own `ecFolderNotEmpty` isn't separately modeled). */
const ERROR_INVALID_OBJECT = 0x80070005;

/** `MAPI_E_TOO_COMPLEX`, for a subfolder tree deeper than `MAX_FOLDER_DEPTH` or larger than `MAX_FOLDERS_PER_DELETE`. */
const ERROR_TOO_COMPLEX = 0x80040117;

/** `DeleteFolderFlags` bits (`[MS-OXCFOLD]` §2.2.1.3.1). `DELETE_HARD_DELETE` (`0x10`) is deliberately not
 * honored - see the class doc comment. */
const DEL_MESSAGES = 0x01;
const DEL_FOLDERS = 0x04;

/** The deepest subfolder nesting a recursive delete walks. */
export const MAX_FOLDER_DEPTH = 32;

/** The most folders one recursive delete removes. */
export const MAX_FOLDERS_PER_DELETE = 1000;

/** The most items (messages, events, contacts, tasks) one `RopDeleteFolder` deletes. Past this the ROP stops,
 * reports `PartialCompletion`, and leaves the folders that still hold items in place; the client can retry. */
export const MAX_ITEMS_PER_DELETE = MAX_COLLECTION_ROWS;

const UID_SORT: RepoSort = { uid: "ASC" };

/**
 * `RopDeleteFolder` (`[MS-OXCFOLD]`/`[MS-OXCROPS]`, RopId `0x1D`): deletes a folder by `FolderId`, via
 * `RecoverableRepoUtils.delete()` - see `RopDeleteMessagesHandler.ts`'s own doc comment for why that matters (EAS's
 * watermark-based incremental delete detection depends on it).
 *
 * **Real spec semantics honored**: `RopDeleteFolder` only operates on an empty folder by default - `DEL_MESSAGES`
 * must be set to also delete the folder's own messages/calendar events/contacts/tasks, and `DEL_FOLDERS` to also
 * delete (recursively) its subfolders; a non-empty folder deleted without the matching flag fails rather than
 * silently cascading.
 *
 * **Always a soft delete.** `DELETE_HARD_DELETE` is ignored. A permanent delete has to go through restapi's legal
 * hold check first (the REST routes call `LegalHoldUtils.assertNotOnLegalHold()` for `purge=true`), and that helper
 * isn't part of `@rapidmx/restapi`'s public exports. A soft-deleted item stays recoverable and discoverable, which is
 * exactly what a hold requires, so MAPI never purges. Each deleted message is audited as `MESSAGE_DELETE`, like the
 * REST message route.
 *
 * **Bounded.** The subfolder tree is walked iteratively with a visited set (a corrupt `parentFolderUid` cycle
 * can't loop), at most `MAX_FOLDER_DEPTH` deep and `MAX_FOLDERS_PER_DELETE` wide, and only through folders in the
 * caller's own mailbox. Items are fetched in explicit, paged queries (a bare `find()` stops at 100 rows, which used
 * to leave items behind in a folder that was then deleted), up to `MAX_ITEMS_PER_DELETE`. Every query and delete is
 * charged to the request's `ExecuteBudget` before it runs, so repeating this ROP can't multiply the walk unbounded.
 *
 * The request's own `InputHandleIndex` (nominally the *parent* folder of the one being deleted) is validated to
 * be a real, open folder handle but not cross-checked against the target folder's actual `parentFolderUid` -
 * the same pragmatic simplification `RopOpenMessageHandler`'s own doc comment documents for its analogous
 * `FolderId` field.
 *
 * @author Jean-Philippe Steinmetz
 */
export class RopDeleteFolderHandler implements RopHandler {
    public readonly ropId = ROP_ID_DELETE_FOLDER;

    public async handle(reader: BufferReader, writer: BufferWriter, context: RopContext): Promise<void> {
        reader.readUInt8(); // LogonId - this pragmatic subset doesn't track multiple concurrent logons per session
        const inputHandleIndex: number = reader.readUInt8();
        const deleteFolderFlags: number = reader.readUInt8();
        const folderId: bigint = reader.readBigUInt64LE();

        const handle = context.session.handles[inputHandleIndex];
        if (!handle || handle.type !== "folder") {
            this.writeFailure(writer, inputHandleIndex, ERROR_INVALID_OBJECT);
            return;
        }

        const target: string | undefined = context.session.folderIds[folderId.toString()];
        if (!target?.startsWith("folder:")) {
            this.writeFailure(writer, inputHandleIndex, ERROR_NOT_FOUND);
            return;
        }
        const uid = target.slice("folder:".length);

        if ((deleteFolderFlags & DEL_MESSAGES) === 0 && (await this.hasItems(uid, context))) {
            this.writeFailure(writer, inputHandleIndex, ERROR_INVALID_OBJECT);
            return;
        }
        const childQuery = { parentFolderUid: uid, mailboxUid: context.mailboxUid };
        if ((deleteFolderFlags & DEL_FOLDERS) === 0 && (await findPage(context.folderRepo, childQuery, UID_SORT, 0, 1, context.budget)).length > 0) {
            this.writeFailure(writer, inputHandleIndex, ERROR_INVALID_OBJECT);
            return;
        }

        const folders: string[] | undefined = await this.collectSubtree(uid, context);
        if (!folders) {
            this.writeFailure(writer, inputHandleIndex, ERROR_TOO_COMPLEX);
            return;
        }

        const budget = { remaining: MAX_ITEMS_PER_DELETE };
        let partialCompletion = false;
        for (const folderUid of folders) {
            if (!(await this.deleteItems(folderUid, context, budget))) {
                partialCompletion = true;
                // Every folder deleted below (this loop's earlier iterations) is gone outright - nothing to refresh.
                // This one is different: it survives (the budget ran out before `deleteItems` could empty it, so it's
                // never reached below), but `itemRepos()` always drains `context.messageRepo` for a folder before any
                // other repo, so some or all of its messages may already be gone - its stored counts must be
                // refreshed rather than left counting messages that no longer exist.
                await context.notifyFolderCounts?.([folderUid]);
                break;
            }
            context.budget?.chargeQueries();
            await context.folderRepo.delete(folderUid, { ignoreACL: true });
        }

        writer.writeUInt8(ROP_ID_DELETE_FOLDER);
        writer.writeUInt8(inputHandleIndex);
        writer.writeUInt32LE(0); // ReturnValue - success
        writer.writeUInt8(partialCompletion ? 1 : 0); // PartialCompletion
    }

    private writeFailure(writer: BufferWriter, inputHandleIndex: number, returnValue: number): void {
        writer.writeUInt8(ROP_ID_DELETE_FOLDER);
        writer.writeUInt8(inputHandleIndex);
        writer.writeUInt32LE(returnValue);
    }

    /** The repos a folder's items can live in, skipping the optional ones the context doesn't have. */
    private itemRepos(context: RopContext): RepoUtils<any>[] {
        return [context.messageRepo, context.calendarEventRepo, context.contactRepo, context.taskRepo].filter(
            (repo): repo is RepoUtils<any> => repo !== undefined,
        );
    }

    private async hasItems(folderUid: string, context: RopContext): Promise<boolean> {
        for (const repo of this.itemRepos(context)) {
            if ((await findPage(repo, { folderUid, mailboxUid: context.mailboxUid }, UID_SORT, 0, 1, context.budget)).length > 0) {
                return true;
            }
        }
        return false;
    }

    /**
     * `rootUid` and every folder below it in the caller's mailbox, children before their parents, or `undefined` when
     * the tree is deeper than `MAX_FOLDER_DEPTH` or has more than `MAX_FOLDERS_PER_DELETE` folders.
     */
    private async collectSubtree(rootUid: string, context: RopContext): Promise<string[] | undefined> {
        const visited = new Set<string>([rootUid]);
        const preOrder: string[] = [];
        const stack: { uid: string; depth: number }[] = [{ uid: rootUid, depth: 0 }];
        while (stack.length > 0) {
            const { uid, depth } = stack.pop()!;
            preOrder.push(uid);
            const { items: children, truncated } = await findAllCapped<{ uid: string }>(
                context.folderRepo,
                { parentFolderUid: uid, mailboxUid: context.mailboxUid },
                UID_SORT,
                MAX_FOLDERS_PER_DELETE,
                context.budget,
            );
            if (truncated) {
                return undefined;
            }
            for (const child of children) {
                if (visited.has(child.uid)) {
                    continue;
                }
                if (depth + 1 > MAX_FOLDER_DEPTH || visited.size >= MAX_FOLDERS_PER_DELETE) {
                    return undefined;
                }
                visited.add(child.uid);
                stack.push({ uid: child.uid, depth: depth + 1 });
            }
        }
        return preOrder.reverse();
    }

    /** Soft-deletes `folderUid`'s items within `budget`. Returns `false` if items remain once the budget is spent. */
    private async deleteItems(folderUid: string, context: RopContext, budget: { remaining: number }): Promise<boolean> {
        for (const repo of this.itemRepos(context)) {
            const { items, truncated } = await findAllCapped<any>(
                repo,
                { folderUid, mailboxUid: context.mailboxUid },
                UID_SORT,
                budget.remaining,
                context.budget,
            );
            for (const item of items) {
                context.budget?.chargeQueries();
                await repo.delete(item.uid, { ignoreACL: true });
                if (repo === context.messageRepo) {
                    await auditMessageDelete(context, item);
                }
            }
            budget.remaining -= items.length;
            if (truncated) {
                return false;
            }
        }
        return true;
    }
}
