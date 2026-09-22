///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import * as crypto from "crypto";
import MailComposer from "nodemailer/lib/mail-composer/index.js";
import { ApiErrors } from "@rapidrest/service-core";
import {
    findOrCreateWellKnownFolder,
    FolderType,
    MessageImportance,
    RecipientType,
    scanAndRelay,
    type CalendarEvent,
} from "@rapidmx/restapi";
import type { BufferReader, BufferWriter } from "../codec/BufferCursor.js";
import type { MapiObjectHandle } from "../MapiSessionManager.js";
import { MAX_RECIPIENTS_PER_MESSAGE, parseRecipientList, resolveRecipientList, type ResolvedRecipient } from "./AddressList.js";
import { handleDataOwners, writeStreamKey } from "./HandleDataCache.js";
import { submitMeetingResponse } from "./MeetingMessageClassHandler.js";
import { handleDataStoreOf, type RopContext, type RopHandler } from "./RopHandler.js";
import { readWriteStream } from "./RopWriteStreamHandler.js";
import { asEntity, isInvitePending } from "./RestapiRules.js";

const ROP_ID_SUBMIT_MESSAGE = 0x32;

/** The well-known MAPI HRESULT `MAPI_E_INVALID_OBJECT`, reused for both "the referenced handle isn't a draft
 * message" and "the draft has no resolvable recipients/sender" - the same constant reused throughout this
 * pragmatic subset as a generic "can't do this" signal, not a claim of exact per-condition `[MS-OXCRPC]`
 * return-value-table parity. */
const ERROR_INVALID_OBJECT = 0x80070005;

/** `MAPI_E_INVALID_PARAMETER`: the draft has no recipients, a recipient that isn't a usable address, or no sender. */
const ERROR_INVALID_PARAMETER = 0x80070057;

/** `MAPI_E_NOT_FOUND`: a recipient given only by display name matches no single contact of the caller's. */
const ERROR_NOT_FOUND = 0x8004010f;

/** `MAPI_E_TOO_BIG`: more than `MAX_RECIPIENTS_PER_MESSAGE` recipients. */
const ERROR_TOO_BIG = 0x80040305;

/** `MAPI_E_CALL_FAILED`: the body written through `RopWriteStream` can no longer be reassembled (its chunks expired). */
const ERROR_CALL_FAILED = 0x80004005;

// The same well-known property IDs RopSetPropertiesHandler tracks - duplicated here (rather than imported)
// since importing them would only save a handful of literals; see that file for the real documentation of what
// each means.
const PID_TAG_SUBJECT = 0x0037;
const PID_TAG_MESSAGE_CLASS = 0x001a;
const PID_TAG_DISPLAY_BCC = 0x0e02;
const PID_TAG_DISPLAY_CC = 0x0e03;
const PID_TAG_DISPLAY_TO = 0x0e04;
const PID_TAG_BODY = 0x1000;
const PID_TAG_READ_RECEIPT_REQUESTED = 0x0029;
const PID_TAG_DEFERRED_SEND_TIME = 0x3fef;

/** `PidTagMessageClass` values starting with this prefix (`IPM.Appointment`, `IPM.Appointment.*`) route through
 * `submitAppointment()` instead of the ordinary mail compose/send path below - see that method's own doc
 * comment. */
const MESSAGE_CLASS_APPOINTMENT_PREFIX = "IPM.Appointment";

/** `PidTagMessageClass` values starting with this prefix (`IPM.Schedule.Meeting.Resp.{Pos,Neg,Tent}`) route
 * through `MeetingMessageClassHandler.submitMeetingResponse()` instead of every other path here - see that
 * function's own doc comment. */
const MESSAGE_CLASS_MEETING_RESPONSE_PREFIX = "IPM.Schedule.Meeting.Resp.";

// Re-exported for existing importers; the implementation lives in AddressList.ts.
export { isPlainEmailAddress, parseAddressList } from "./AddressList.js";

function writeResult(writer: BufferWriter, inputHandleIndex: number, returnValue: number): void {
    writer.writeUInt8(ROP_ID_SUBMIT_MESSAGE);
    writer.writeUInt8(inputHandleIndex);
    writer.writeUInt32LE(returnValue);
}

/**
 * `RopSubmitMessage` (`[MS-OXOMSG]`/`[MS-OXCROPS]`): sends a composed message - the actual trigger point for
 * this pragmatic subset's compose/send pipeline, since `RopSaveChangesMessage` itself does no real persistence
 * (see its own doc comment). Builds a raw MIME buffer from whatever `RopSetProperties`/`RopWriteStream` calls
 * accumulated on the draft handle (`Subject`, the `PidTagBody` stream or an inline `PidTagBody` property, and
 * addressing - see below), via `nodemailer`'s `MailComposer` (already a dependency, used elsewhere in this
 * library only to *send*, not build, MIME - this is its first use purely as a RFC 5322 serializer), then calls
 * the exact same `scanAndRelay()` scan-then-relay pipeline `BaseMessageRoute.send()`/EAS's `ComposeMailCommand`
 * already share, and persists a Sent Items copy (mirroring `ComposeMailCommand`'s own always-committed Sent
 * Items behavior for MAPI submission specifically, since - unlike EAS's optional `SaveInSentItems` - a MAPI
 * `RopSubmitMessage` call has no "don't keep a copy" option to honor).
 *
 * **Known, deliberate limitation - no `RopModifyRecipients` support**: real recipient rows (`[MS-OXCMSG]`
 * §2.2.3.2, each with its own `PidTagEmailAddress`/`PidTagRecipientType`) are a substantial structure this
 * pragmatic subset's ROP coverage doesn't include. Instead, addressing is read from `PidTagDisplayTo`/
 * `DisplayCc`/`DisplayBcc` - the same semicolon-separated "cached recipient display string" properties real
 * Outlook *also* always sets via `RopSetProperties` alongside `RopModifyRecipients` (as a display-only
 * convenience cache). Entries are separated by `;` and parsed with nodemailer's `addressparser`, so both bare
 * addresses and resolved `Name <address>` forms work. A client that resolved a name against an address book may
 * write the display name alone; that is looked up among the caller's own contacts (see `resolveRecipientList`), and
 * a name matching no single contact fails the submit with `MAPI_E_NOT_FOUND`. It is still **not** a substitute for
 * real recipient rows (there is no recipient table to fall back on), so the whole send is refused rather than sent to
 * fewer people than the user addressed. A malformed recipient fails with `MAPI_E_INVALID_PARAMETER`.
 *
 * `SubmitFlags` (`PreprocessOnly`, ...) is decoded to advance past it correctly but not honored - this
 * pragmatic subset has no transport-agent preprocessing distinction to vary by flag.
 *
 * **`PidTagDeferredSendTime`** ("Do not deliver before"), if set to a future time via `RopSetProperties`,
 * parks the message in Outbox with `Message.scheduledSendTime` set instead of relaying immediately - mirroring
 * `BaseMessageRoute.send()`'s own identical branch - for `ScheduledSendJob` to relay later via this same
 * `scanAndRelay()` pipeline. **`PidTagReadReceiptRequested`**, if `true`, attaches a real
 * `Disposition-Notification-To` header (see `util/ReceiptUtils.ts`) and is recorded on the persisted message's
 * own `requestReceipt` field - honored only when the client explicitly sets it, not this mailbox's
 * `alwaysRequestReceiptInternal`/`External` defaults (those apply to the REST/webmail compose path only).
 *
 * **Once per draft.** A mail or meeting-response draft is marked `submitted` as soon as a submit gets past the handle
 * check, whatever the outcome, and a submitted draft can't be submitted again (`MAPI_E_INVALID_OBJECT`) or changed by
 * `RopSetProperties`/`RopWriteStream`; its write-stream chunks are deleted once read. Each `Execute` runs at most
 * `MAX_SUBMITS_PER_EXECUTE` submits, and a message may have at most `MAX_RECIPIENTS_PER_MESSAGE` recipients
 * (`MAPI_E_TOO_BIG`), counted before any display name is looked up.
 *
 * **Calendar branches**: a draft whose `PidTagMessageClass` starts with `"IPM.Appointment"` is routed to
 * `submitAppointment()` instead of the mail path below - see that method's own doc comment. One starting with
 * `"IPM.Schedule.Meeting.Resp."` (an attendee's own accept/decline/tentative response to a meeting this server
 * previously invited them to) is routed to `MeetingMessageClassHandler.submitMeetingResponse()` instead - see
 * that function's own doc comment.
 *
 * @author Jean-Philippe Steinmetz
 */
export class RopSubmitMessageHandler implements RopHandler {
    public readonly ropId = ROP_ID_SUBMIT_MESSAGE;

    public async handle(reader: BufferReader, writer: BufferWriter, context: RopContext): Promise<void> {
        reader.readUInt8(); // LogonId - this pragmatic subset doesn't track multiple concurrent logons per session
        const inputHandleIndex: number = reader.readUInt8();
        reader.readUInt8(); // SubmitFlags - see class doc comment

        const handle = context.session.handles[inputHandleIndex];
        if (!handle || handle.type !== "message" || handle.submitted) {
            writer.writeUInt8(ROP_ID_SUBMIT_MESSAGE);
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt32LE(ERROR_INVALID_OBJECT);
            return;
        }
        context.budget?.chargeSubmit();

        const properties = handle.draftProperties ?? {};
        const messageClass = properties[String(PID_TAG_MESSAGE_CLASS)] ?? "IPM.Note";
        if (messageClass.startsWith(MESSAGE_CLASS_APPOINTMENT_PREFIX)) {
            await this.submitAppointment(handle, context, writer, inputHandleIndex);
            return;
        }
        // From here on the draft is used up by this attempt, whatever its outcome, so it can't be sent again.
        handle.submitted = true;
        if (messageClass.startsWith(MESSAGE_CLASS_MEETING_RESPONSE_PREFIX)) {
            writeResult(writer, inputHandleIndex, await submitMeetingResponse(messageClass, properties, context));
            return;
        }

        const recipientCount = [PID_TAG_DISPLAY_TO, PID_TAG_DISPLAY_CC, PID_TAG_DISPLAY_BCC].reduce(
            (count, propertyId) => count + parseRecipientList(properties[String(propertyId)]).length,
            0,
        );
        if (recipientCount > MAX_RECIPIENTS_PER_MESSAGE) {
            writeResult(writer, inputHandleIndex, ERROR_TOO_BIG);
            return;
        }

        const resolutions = [
            await resolveRecipientList(properties[String(PID_TAG_DISPLAY_TO)], context),
            await resolveRecipientList(properties[String(PID_TAG_DISPLAY_CC)], context),
            await resolveRecipientList(properties[String(PID_TAG_DISPLAY_BCC)], context),
        ];
        const [to, cc, bcc] = resolutions.map((resolution) => resolution.recipients);
        const mailbox = await context.mailboxRepo.findOne(context.mailboxUid, { ignoreACL: true });
        const envelopeFrom: string | undefined = mailbox?.primarySmtpAddress;

        // Refuse the whole send over one malformed or unresolvable recipient rather than silently dropping it.
        if (!envelopeFrom || to.length + cc.length + bcc.length === 0 || resolutions.some((resolution) => resolution.invalid.length > 0)) {
            writeResult(writer, inputHandleIndex, ERROR_INVALID_PARAMETER);
            return;
        }
        if (resolutions.some((resolution) => resolution.unresolved.length > 0)) {
            writeResult(writer, inputHandleIndex, ERROR_NOT_FOUND);
            return;
        }

        const subject = properties[String(PID_TAG_SUBJECT)] ?? "";
        const bodyText = await resolveDraftBody(context, inputHandleIndex, properties);
        if (bodyText === undefined) {
            writeResult(writer, inputHandleIndex, ERROR_CALL_FAILED);
            return;
        }
        const requestReceipt = properties[String(PID_TAG_READ_RECEIPT_REQUESTED)] === "true";
        // "Do not deliver before" - a real Outlook compose option, tracked as a plain ISO-8601 string by
        // RopSetPropertiesHandler.stringifyValue()'s PtypTime encoding.
        const deferredSendTimeRaw = properties[String(PID_TAG_DEFERRED_SEND_TIME)];
        const deferredSendTime: Date | undefined = deferredSendTimeRaw ? new Date(deferredSendTimeRaw) : undefined;
        const isDeferred = deferredSendTime !== undefined && deferredSendTime.getTime() > Date.now();

        const raw: Buffer = await new MailComposer({
            from: envelopeFrom,
            to: to.map(mailAddress),
            cc: cc.map(mailAddress),
            bcc: bcc.map(mailAddress),
            subject,
            text: bodyText,
            // A real Disposition-Notification-To request, mirroring BaseMessageRoute.send()'s own explicit
            // Message.requestReceipt handling (see util/ReceiptUtils.ts) - honored here only when the client
            // itself asked for one via PidTagReadReceiptRequested, not this mailbox's own
            // alwaysRequestReceiptInternal/External defaults (those apply to the REST/webmail compose path,
            // which already goes through BaseMessageRoute.send() directly).
            ...(requestReceipt ? { headers: [{ key: "Disposition-Notification-To", value: envelopeFrom }] } : {}),
        })
            .compile()
            .build();
        const envelopeTo = [...to, ...cc, ...bcc].map((recipient) => recipient.address);

        const bodyBlobKey = `bodies/${crypto.randomUUID()}`;
        const recipients = [
            ...to.map((recipient) => storedRecipient(recipient, RecipientType.TO)),
            ...cc.map((recipient) => storedRecipient(recipient, RecipientType.CC)),
            ...bcc.map((recipient) => storedRecipient(recipient, RecipientType.BCC)),
        ];

        if (isDeferred) {
            // Mirrors BaseMessageRoute.send()'s own scheduledSendTime branch: park the message in Outbox,
            // unscanned/unrelayed, for the same ScheduledSendJob to pick up and relay (via this exact
            // scanAndRelay() pipeline) once due - see that job's own doc comment.
            await context.blobStore.put(bodyBlobKey, raw, { contentType: "message/rfc822" });
            const outbox = await findOrCreateWellKnownFolder(context.folderRepo, context.folderClass, context.mailboxUid, FolderType.OUTBOX);
            await context.messageRepo.create(
                new context.messageClass({
                    folderUid: outbox.uid,
                    mailboxUid: context.mailboxUid,
                    messageId: `${crypto.randomUUID()}@mapi`,
                    subject,
                    from: { address: envelopeFrom, type: RecipientType.TO },
                    recipients,
                    sentDate: new Date(),
                    receivedDate: new Date(),
                    bodyBlobKey,
                    bodyPreview: bodyText.slice(0, 200),
                    flags: { read: true, flagged: false, answered: false, forwarded: false },
                    importance: MessageImportance.NORMAL,
                    references: [],
                    hasAttachments: false,
                    scheduledSendTime: deferredSendTime,
                    requestReceipt,
                }),
                { ignoreACL: true },
            );
            // Adding a message to a folder always bumps its sync key too - see `refreshFolderCounts()`'s own doc comment.
            await context.notifyFolderCounts?.([outbox.uid], { bumpSyncKey: true });

            writer.writeUInt8(ROP_ID_SUBMIT_MESSAGE);
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt32LE(0); // ReturnValue - success
            return;
        }

        const relayed = await scanAndRelay(raw, envelopeFrom, envelopeTo, context.scanPipeline, context.mailTransport, context.blobStore);

        // The Sent Items copy stores exactly what was relayed: scanAndRelay may have added a Message-ID header, and
        // that header's value (plus the thread it derived) is what recipients' copies carry and what recall targets.
        await context.blobStore.put(bodyBlobKey, relayed.raw, { contentType: "message/rfc822" });
        const sentFolder = await findOrCreateWellKnownFolder(context.folderRepo, context.folderClass, context.mailboxUid, FolderType.SENT_ITEMS);
        await context.messageRepo.create(
            new context.messageClass({
                folderUid: sentFolder.uid,
                mailboxUid: context.mailboxUid,
                messageId: relayed.messageId,
                conversationId: relayed.conversationId,
                encrypted: relayed.encrypted,
                subject,
                from: { address: envelopeFrom, type: RecipientType.TO },
                recipients,
                sentDate: new Date(),
                receivedDate: new Date(),
                bodyBlobKey,
                sanitizedHtmlBlobKey: relayed.sanitizedHtmlBlobKey,
                bodyPreview: bodyText.slice(0, 200),
                flags: { read: true, flagged: false, answered: false, forwarded: false },
                importance: MessageImportance.NORMAL,
                references: [],
                hasAttachments: false,
                requestReceipt,
            }),
            { ignoreACL: true },
        );
        // Adding a message to a folder always bumps its sync key too - see `refreshFolderCounts()`'s own doc comment.
        await context.notifyFolderCounts?.([sentFolder.uid], { bumpSyncKey: true });

        writer.writeUInt8(ROP_ID_SUBMIT_MESSAGE);
        writer.writeUInt8(inputHandleIndex);
        writer.writeUInt32LE(0); // ReturnValue - success
    }

    /**
     * Submits a Calendar item. The appointment itself is already persisted (`RopSaveChangesMessageHandler` writes the
     * real `CalendarEvent` row at Save time, unlike a mail draft, which stays purely in-session until Submit), so there
     * is nothing left to store and nothing is sent from here.
     *
     * **Invites are restapi's `MeetingSchedulingJob`'s job.** That job sends an iTIP `REQUEST` for every
     * organizer-owned event whose `inviteSequenceSent` differs from its `sequence`, claiming each revision with a
     * versioned update before sending and building the invite with restapi's own `buildEventIcs` (so a recurring
     * meeting's invite carries its `RRULE` and exceptions). Saving stamps `inviteSequenceSent = sequence`, so a meeting
     * that is only saved ("Save & Close", or "save changes but don't send") invites no one. Submitting clears the stamp
     * (`inviteSequenceSent: null`) with a versioned update, which makes the job send the current revision exactly once.
     * - A revision that is already waiting for the job (submitted and not yet sent, or never stamped) is left alone, so
     * submitting it again before the job runs doesn't cause a second invite.
     * - The update is versioned (`asEntity`). The job only ever claims a waiting revision, so a conflict here comes from
     * some other edit (an attendee's reply updating the row, a REST edit); the row is read again and the check repeated,
     * up to `MAX_INVITE_REQUEST_ATTEMPTS` times, after which the ROP fails (`MAPI_E_CALL_FAILED`).
     * - Submitting a revision the job has already sent sends it again: the client asked to send.
     *
     * Only the organizer's own copy may be submitted (an attendee's copy of someone else's meeting answers
     * `MAPI_E_INVALID_OBJECT`, as before, and nothing is changed).
     */
    private async submitAppointment(handle: MapiObjectHandle, context: RopContext, writer: BufferWriter, inputHandleIndex: number): Promise<void> {
        const uid = handle.entityUid.startsWith("calendarEvent:") ? handle.entityUid.slice("calendarEvent:".length) : undefined;
        const event: CalendarEvent | undefined = uid ? await context.calendarEventRepo.findOne(uid, { ignoreACL: true }) : undefined;
        const mailbox = event?.mailboxUid === context.mailboxUid ? await context.mailboxRepo.findOne(context.mailboxUid, { ignoreACL: true }) : undefined;
        const callerAddresses: string[] = mailbox ? [mailbox.primarySmtpAddress, ...(mailbox.aliasAddresses ?? [])].map((a: string) => a.toLowerCase()) : [];
        const isOrganizer = !!event && callerAddresses.includes((event.organizer?.address ?? "").toLowerCase());
        if (!isOrganizer) {
            writeResult(writer, inputHandleIndex, ERROR_INVALID_OBJECT);
            return;
        }
        writeResult(writer, inputHandleIndex, (await requestInvites(context, uid!, event)) ? 0 : ERROR_NOT_FOUND);
    }
}

/** How many times `requestInvites` re-reads a row whose versioned update lost to another edit before giving up. */
export const MAX_INVITE_REQUEST_ATTEMPTS = 3;

/** Makes restapi's `MeetingSchedulingJob` send invites for the organizer copy `event`'s current revision (see
 * `submitAppointment`). `false` when the row disappeared meanwhile. */
async function requestInvites(context: RopContext, uid: string, event: CalendarEvent): Promise<boolean> {
    let row: CalendarEvent | undefined = event;
    for (let attempt = 0; attempt < MAX_INVITE_REQUEST_ATTEMPTS; attempt++) {
        if (!row) {
            return false;
        }
        if (isInvitePending(row)) {
            return true;
        }
        try {
            await context.calendarEventRepo.update(
                { uid: (row as any).uid, version: (row as any).version, inviteSequenceSent: null } as any,
                asEntity(context.calendarEventRepo, row),
                { ignoreACL: true },
            );
            return true;
        } catch (err: any) {
            if (err?.code !== ApiErrors.INVALID_OBJECT_VERSION) {
                throw err;
            }
            row = await context.calendarEventRepo.findOne(uid, { ignoreACL: true });
        }
    }
    throw new Error("RopSubmitMessage: the meeting kept changing while its invites were requested.");
}

/** A recipient as nodemailer's `MailComposer` takes it: a display name only when there is one. */
function mailAddress(recipient: ResolvedRecipient): string | { name: string; address: string } {
    return recipient.name ? { name: recipient.name, address: recipient.address } : recipient.address;
}

/** A recipient as the Sent Items copy stores it. */
function storedRecipient(recipient: ResolvedRecipient, type: RecipientType): { address: string; displayName?: string; type: RecipientType } {
    return { address: recipient.address, ...(recipient.name ? { displayName: recipient.name } : {}), type };
}

/** Prefers a `RopWriteStream`-accumulated body (the real path a large body takes) over an inline `PidTagBody`
 * set directly via `RopSetProperties` (only realistic for a short body small enough to set inline) - see
 * `RopOpenStreamHandler`'s write-mode branch for how `writeTargetHandleIndex` links a stream handle back to the
 * draft message handle it was opened against. The generation must match too, so a stream left over from an
 * earlier message at the same handle index never supplies this message's body. `undefined` when the stream's
 * chunks can no longer be reassembled - sending a truncated body would be worse than failing. */
async function resolveDraftBody(context: RopContext, messageHandleIndex: number, properties: Record<string, string>): Promise<string | undefined> {
    const generation = context.session.handles[messageHandleIndex].generation;
    for (const [index, candidate] of Object.entries(context.session.handles)) {
        if (
            candidate.type === "stream" &&
            candidate.writeTargetHandleIndex === messageHandleIndex &&
            candidate.writeTargetGeneration === generation &&
            candidate.writeSize
        ) {
            const raw = await readWriteStream(context, Number(index), candidate);
            // The chunks are only needed for this one submit; best effort, they expire anyway.
            await handleDataStoreOf(context)
                .delete(writeStreamKey(context.session.uid, Number(index), candidate.generation), handleDataOwners(context.session.uid, context.session.userUid))
                .catch(() => undefined);
            return raw?.toString("utf16le").replace(/\0+$/, "");
        }
    }
    return properties[String(PID_TAG_BODY)] ?? "";
}
