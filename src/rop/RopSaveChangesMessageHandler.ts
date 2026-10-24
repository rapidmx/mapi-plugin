///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import crypto from "crypto";
import {
    AttendeeResponseStatus,
    AttendeeRole,
    BusyStatus,
    RecipientType,
    type Attendee,
    type CalendarEvent,
    type RecurrenceRule,
} from "@rapidmx/restapi";
import { decodeAppointmentRecurrence } from "../codec/AppointmentRecurrence.js";
import { BufferReader, type BufferWriter } from "../codec/BufferCursor.js";
import { decodeTimeZoneStruct } from "../codec/MapiTimeZone.js";
import type { MapiObjectHandle, MapiSessionContext } from "../MapiSessionManager.js";
import {
    BUSY_STATUS_FROM_CODE,
    LID_APPOINTMENT_END_WHOLE,
    LID_APPOINTMENT_RECUR,
    LID_APPOINTMENT_START_WHOLE,
    LID_BUSY_STATUS,
    LID_LOCATION,
    LID_REMINDER_DELTA,
    LID_TIME_ZONE_STRUCT,
    PSETID_APPOINTMENT,
    PSETID_COMMON,
} from "./CalendarNamedProperties.js";
import { assignOrGetMid } from "./MessageTarget.js";
import { resolveNamedProperty } from "./NamedPropertyRegistry.js";
import { MAX_RECIPIENTS_PER_MESSAGE } from "./AddressList.js";
import { asEntity, isInvitePending } from "./RestapiRules.js";
import { parseAddressList } from "./RopSubmitMessageHandler.js";
import type { RopContext, RopHandler } from "./RopHandler.js";

const ROP_ID_SAVE_CHANGES_MESSAGE = 0x0c;

/** The well-known MAPI HRESULT `MAPI_E_INVALID_OBJECT`, reused for "the referenced handle isn't a message (or
 * doesn't exist)" - the same constant `RopGetPropertiesSpecificHandler` uses for its own analogous check. */
const ERROR_INVALID_OBJECT = 0x80070005;

/** `MAPI_E_TOO_BIG`: a meeting with more than `MAX_RECIPIENTS_PER_MESSAGE` attendees. */
const ERROR_TOO_BIG = 0x80040305;

// The same well-known property IDs RopSetPropertiesHandler tracks - duplicated here (rather than imported)
// since importing them would only save a handful of literals; see that file for the real documentation of what
// each means.
const PID_TAG_SUBJECT = 0x0037;
const PID_TAG_MESSAGE_CLASS = 0x001a;
const PID_TAG_DISPLAY_CC = 0x0e03;
const PID_TAG_DISPLAY_TO = 0x0e04;

/** `PidTagMessageClass` values starting with this prefix route through `saveAppointment()` instead of the
 * generic draft-MID-assignment path below - see that method's own doc comment. */
const MESSAGE_CLASS_APPOINTMENT_PREFIX = "IPM.Appointment";

/**
 * `RopSaveChangesMessage` (`[MS-OXCMSG]`/`[MS-OXCROPS]`): persists a message's pending changes, returning the
 * `MID` a client can reference it by afterwards. **Pragmatic subset**: this library has no separate Drafts-
 * folder persistence step (matching EAS's own `ComposeMailCommand`, which never creates a standalone draft
 * `Message` row either) - saving a `RopCreateMessage` draft does not yet write anything to a database. It only
 * assigns a session-scoped MID (`MessageTarget.assignOrGetMid`, the exact mechanism a `RopQueryRows` row's
 * `PidTagMid` column already uses) so the wire contract's `MessageId` field is well-formed, keyed by this
 * handle's own index (`` `draft:${inputHandleIndex}` ``) since an unsaved draft has no real `"message:<uid>"`
 * target yet. The real work - actually building and sending a MIME message from the properties/body
 * accumulated so far - happens at `RopSubmitMessage` time, not here; calling Save without ever calling Submit
 * silently discards the draft when the session expires, a real, deliberate limitation (this pragmatic subset
 * has no autosave-recovery story) rather than a silently-wrong one.
 *
 * Re-saving an already-real message (a `RopOpenMessage` handle, not a fresh `RopCreateMessage` draft) is
 * accepted but is a no-op beyond echoing that message's own already-assigned MID - editing an existing
 * message's properties is out of scope for this pragmatic subset's compose/send-only ROP coverage.
 *
 * `SaveFlags` is decoded to advance past it correctly but not honored - this pragmatic subset has no
 * conflict-resolution/force-save semantics to vary by flag.
 *
 * **Calendar branch**: unlike mail, a draft whose `PidTagMessageClass` starts with `"IPM.Appointment"` *is*
 * durably persisted here (`saveAppointment()`) - creating (or updating, if this handle already refers to a real
 * `"calendarEvent:<uid>"`) a real `CalendarEvent` row, so a subsequent `RopGetPropertiesSpecific` can read it
 * back before `RopSubmitMessage` is ever called (a calendar item must exist as soon as it's saved, the same way
 * a real Exchange server behaves - unlike a mail draft, which this pragmatic subset never persists at all until
 * Submit). Saving never invites anyone: a saved revision is stamped as already handled (`inviteSequenceSent =
 * sequence`), so restapi's `MeetingSchedulingJob` skips it until `RopSubmitMessageHandler.submitAppointment` asks for
 * invites. Attendees of a submitted meeting are then mailed by that job, which is why a meeting is capped at
 * `MAX_RECIPIENTS_PER_MESSAGE` attendees here. Client-supplied times and reminder minutes that don't parse are ignored (the field keeps its
 * existing or default value) rather than stored as `Invalid Date`/`NaN`, and an update is versioned (`asEntity`).
 *
 * @author Jean-Philippe Steinmetz
 */
export class RopSaveChangesMessageHandler implements RopHandler {
    public readonly ropId = ROP_ID_SAVE_CHANGES_MESSAGE;

    public async handle(reader: BufferReader, writer: BufferWriter, context: RopContext): Promise<void> {
        reader.readUInt8(); // LogonId - this pragmatic subset doesn't track multiple concurrent logons per session
        const responseHandleIndex: number = reader.readUInt8();
        const inputHandleIndex: number = reader.readUInt8();
        reader.readUInt8(); // SaveFlags - see class doc comment

        const handle = context.session.handles[inputHandleIndex];
        if (!handle || handle.type !== "message") {
            writer.writeUInt8(ROP_ID_SAVE_CHANGES_MESSAGE);
            writer.writeUInt8(responseHandleIndex);
            writer.writeUInt32LE(ERROR_INVALID_OBJECT);
            return;
        }

        const draftProperties = handle.draftProperties ?? {};
        const messageClass = draftProperties[String(PID_TAG_MESSAGE_CLASS)] ?? "IPM.Note";
        const mid = messageClass.startsWith(MESSAGE_CLASS_APPOINTMENT_PREFIX)
            ? await this.saveAppointment(handle, draftProperties, context)
            : BigInt(assignOrGetMid(context.session, handle.entityUid !== "" ? handle.entityUid : `draft:${inputHandleIndex}`));
        if (mid === undefined) {
            writer.writeUInt8(ROP_ID_SAVE_CHANGES_MESSAGE);
            writer.writeUInt8(responseHandleIndex);
            writer.writeUInt32LE(ERROR_TOO_BIG);
            return;
        }

        writer.writeUInt8(ROP_ID_SAVE_CHANGES_MESSAGE);
        writer.writeUInt8(responseHandleIndex);
        writer.writeUInt32LE(0); // ReturnValue - success
        writer.writeUInt8(inputHandleIndex);
        writer.writeBigUInt64LE(mid);
    }

    /**
     * Persists (creating, or updating an already-real one) a `CalendarEvent` from the draft's accumulated
     * `RopSetProperties` values, and returns the MID a client can reference it by. Every field this pragmatic
     * subset understands (per `CalendarNamedProperties.ts`) is decoded via `decodeCalendarFieldsFromDraft`
     * below - the write-side mirror of `PropertyResolvers.calendarEventValueFor`'s read-side switch. A field the
     * client never set (e.g. no `RopSetProperties` call touched `PidLidBusyStatus`) falls back to the existing
     * row's own value on an update, or the entity class's own default on a fresh create.
     *
     * Returns `undefined`, saving nothing, for more than `MAX_RECIPIENTS_PER_MESSAGE` attendees: restapi's
     * `MeetingSchedulingJob` mails every attendee of a saved meeting.
     */
    private async saveAppointment(handle: MapiObjectHandle, draftProperties: Record<string, string>, context: RopContext): Promise<bigint | undefined> {
        const decoded = decodeCalendarFieldsFromDraft(context.session, draftProperties);
        if (decoded.attendeeAddresses.length > MAX_RECIPIENTS_PER_MESSAGE) {
            return undefined;
        }
        const attendees: Attendee[] = decoded.attendeeAddresses.map((address) => ({
            address,
            role: AttendeeRole.REQUIRED,
            responseStatus: AttendeeResponseStatus.NEEDS_ACTION,
            isOrganizer: false,
        }));

        const mailbox = await context.mailboxRepo.findOne(context.mailboxUid, { ignoreACL: true });
        if (handle.entityUid.startsWith("calendarEvent:")) {
            const uid = handle.entityUid.slice("calendarEvent:".length);
            const existing: (CalendarEvent & { uid: string; version: number }) | undefined = await context.calendarEventRepo.findOne(uid, {
                ignoreACL: true,
            });
            if (existing) {
                const updated = {
                    uid: existing.uid,
                    version: existing.version,
                    title: decoded.title ?? existing.title,
                    location: decoded.location ?? existing.location,
                    startDate: decoded.startDate ?? existing.startDate,
                    endDate: decoded.endDate ?? existing.endDate,
                    busyStatus: decoded.busyStatus ?? existing.busyStatus,
                    recurrenceRule: decoded.recurrenceRule ?? existing.recurrenceRule,
                    timezone: decoded.timezone ?? existing.timezone,
                    reminderMinutesBeforeStart: decoded.reminderMinutesBeforeStart ?? existing.reminderMinutesBeforeStart,
                    // An attendee who is still invited keeps their recorded response.
                    attendees:
                        attendees.length > 0
                            ? attendees.map(
                                  (attendee) =>
                                      existing.attendees.find((prior) => prior.address.toLowerCase() === attendee.address.toLowerCase()) ??
                                      attendee,
                              )
                            : existing.attendees,
                };
                const sequence: number = (existing.sequence ?? 0) + (isSchedulingRelevantChange(existing, updated) ? 1 : 0);
                const patch: Record<string, any> = {
                    ...updated,
                    sequence,
                    // Saving sends nothing, so this revision counts as handled - unless a submitted revision is still
                    // waiting for the job, which then sends the meeting as it now stands rather than not at all.
                    inviteSequenceSent: isInvitePending(existing) ? (existing.inviteSequenceSent ?? null) : sequence,
                };
                // A protected calendar's event is sealed; the placeholders Outlook shows and sends back are ignored.
                await context.recordSealer?.sealWrite("calendar", mailbox, patch, existing);
                await context.calendarEventRepo.update(patch, asEntity(context.calendarEventRepo, existing), { ignoreACL: true });
            }
            return BigInt(assignOrGetMid(context.session, handle.entityUid));
        }

        // draftFolderUid is a "folder:<uid>"-prefixed target string (RopCreateMessageHandler stores the same
        // target FolderTarget.assignOrGetFid assigned), not a bare uid - strip the prefix before persisting.
        const draftFolderTarget = handle.draftFolderUid ?? "";
        const folderUid = draftFolderTarget.startsWith("folder:") ? draftFolderTarget.slice("folder:".length) : draftFolderTarget;
        const fields: Record<string, any> = {
                folderUid,
                mailboxUid: context.mailboxUid,
                title: decoded.title ?? "",
                location: decoded.location,
                startDate: decoded.startDate ?? new Date(),
                endDate: decoded.endDate ?? new Date(),
                timezone: decoded.timezone ?? "UTC",
                organizer: { address: mailbox?.primarySmtpAddress ?? "", displayName: mailbox?.displayName, type: RecipientType.TO },
                attendees,
                busyStatus: decoded.busyStatus ?? BusyStatus.BUSY,
                recurrenceRule: decoded.recurrenceRule,
                reminderMinutesBeforeStart: decoded.reminderMinutesBeforeStart,
                icalUid: `${crypto.randomUUID()}@mapi`,
                sequence: 0,
                // Handled until submitted: saving a meeting doesn't invite its attendees.
                inviteSequenceSent: 0,
        };
        await context.recordSealer?.sealWrite("calendar", mailbox, fields);
        const created: CalendarEvent & { uid: string } = await context.calendarEventRepo.create(new context.calendarEventClass(fields), { ignoreACL: true });
        handle.entityUid = `calendarEvent:${created.uid}`;
        return BigInt(assignOrGetMid(context.session, handle.entityUid));
    }
}

/**
 * `true` when an update changes what attendees were invited to: the start or end time, location, recurrence, or
 * who is invited. The same fields restapi's `BaseCalendarEventRoute.update()` bumps `sequence` for (a MAPI save
 * never sets `status`), so an edit made in Outlook is re-sent to attendees and supersedes their copy (RFC 5546
 * `SEQUENCE`) exactly like one made through REST.
 */
export function isSchedulingRelevantChange(
    existing: Pick<CalendarEvent, "startDate" | "endDate" | "location" | "recurrenceRule" | "attendees">,
    updated: Pick<CalendarEvent, "startDate" | "endDate" | "location" | "recurrenceRule" | "attendees">,
): boolean {
    const addresses = (attendees: Attendee[]): string =>
        JSON.stringify(attendees.map((attendee) => attendee.address.toLowerCase()).sort());
    return (
        new Date(updated.startDate).getTime() !== new Date(existing.startDate).getTime() ||
        new Date(updated.endDate).getTime() !== new Date(existing.endDate).getTime() ||
        (updated.location ?? "") !== (existing.location ?? "") ||
        JSON.stringify(updated.recurrenceRule ?? null) !== JSON.stringify(existing.recurrenceRule ?? null) ||
        addresses(updated.attendees) !== addresses(existing.attendees)
    );
}

/** Every field but `attendeeAddresses` is deliberately optional and left `undefined` when this particular
 * `RopSetProperties` call never touched it - `undefined` here specifically means "the client didn't set this
 * property this time," distinct from a real, deliberately-empty value (e.g. clearing `PidTagSubject` to `""`).
 * `saveAppointment`'s update branch relies on this distinction to fall back to the existing row's own value for
 * an untouched field instead of blanking it - see that method's own doc comment. */
interface DecodedCalendarFields {
    title?: string;
    location?: string;
    startDate?: Date;
    endDate?: Date;
    busyStatus?: BusyStatus;
    recurrenceRule?: RecurrenceRule;
    timezone?: string;
    reminderMinutesBeforeStart?: number;
    attendeeAddresses: string[];
}

/** Decodes a draft message handle's accumulated `RopSetProperties` values (`Record<string, string>`, keyed by
 * decimal `PropertyId`) back into typed `CalendarEvent` fields - the write-side mirror of
 * `PropertyResolvers.calendarEventValueFor`'s read-side switch, sharing the same `CalendarNamedProperties.ts`
 * LID table. `PidTagDisplayTo`/`Cc` (not `Bcc` - a meeting has no concept of a blind attendee) become
 * `attendeeAddresses`, matching `RopSubmitMessageHandler`'s own mail-side addressing pragmatic subset (no
 * `RopModifyRecipients` support - see that file's own doc comment).
 *
 * `title` used to default to `""` whenever `PidTagSubject` wasn't present in `properties` at all, rather than
 * staying `undefined` the way every other optional field here already did (the named-property loop below only
 * ever assigns a field when its key is actually present) - `saveAppointment`'s update branch then wrote that
 * `""` straight through with no existing-value fallback, silently blanking the title of any event edited via a
 * `RopSetProperties` call that only touched some other property (e.g. just `PidLidBusyStatus`). Checking `in`
 * explicitly (rather than `properties[key] ?? ""`) preserves the same "present but genuinely empty" vs.
 * "absent" distinction `PidTagSubject` deserves alongside every other field here. */
function decodeCalendarFieldsFromDraft(session: MapiSessionContext, properties: Record<string, string>): DecodedCalendarFields {
    const subjectKey = String(PID_TAG_SUBJECT);
    const decoded: DecodedCalendarFields = {
        title: subjectKey in properties ? properties[subjectKey] : undefined,
        attendeeAddresses: [
            ...parseAddressList(properties[String(PID_TAG_DISPLAY_TO)]),
            ...parseAddressList(properties[String(PID_TAG_DISPLAY_CC)]),
        ],
    };

    for (const [key, value] of Object.entries(properties)) {
        const propertyId = Number(key);
        if (propertyId < 0x8000) {
            continue;
        }
        const named = resolveNamedProperty(session, propertyId);
        if (!named || named.kind !== "lid") {
            continue;
        }
        const guid = named.guid.toLowerCase();
        if (guid === PSETID_APPOINTMENT) {
            switch (named.lid) {
                case LID_LOCATION:
                    decoded.location = value;
                    break;
                case LID_APPOINTMENT_START_WHOLE:
                    decoded.startDate = validDate(value);
                    break;
                case LID_APPOINTMENT_END_WHOLE:
                    decoded.endDate = validDate(value);
                    break;
                case LID_BUSY_STATUS:
                    decoded.busyStatus = BUSY_STATUS_FROM_CODE[Number(value)] ?? BusyStatus.BUSY;
                    break;
                case LID_APPOINTMENT_RECUR:
                    decoded.recurrenceRule = decodeAppointmentRecurrence(new BufferReader(Buffer.from(value, "base64")));
                    break;
                case LID_TIME_ZONE_STRUCT:
                    decoded.timezone = decodeTimeZoneStruct(new BufferReader(Buffer.from(value, "base64")));
                    break;
                default:
                    break;
            }
        } else if (guid === PSETID_COMMON && named.lid === LID_REMINDER_DELTA) {
            decoded.reminderMinutesBeforeStart = validMinutes(value);
        }
    }

    return decoded;
}

/** `value` as a date, or `undefined` (the field is left as it was) when it doesn't parse. A client picks the property
 * type it sends, so a "time" can arrive as any string. */
function validDate(value: string): Date | undefined {
    const date = new Date(value);
    return Number.isNaN(date.getTime()) ? undefined : date;
}

/** `value` as a whole, non-negative number of minutes, or `undefined` when it isn't one. */
function validMinutes(value: string): number | undefined {
    const minutes = Number(value);
    return value.trim() !== "" && Number.isSafeInteger(minutes) && minutes >= 0 ? minutes : undefined;
}
