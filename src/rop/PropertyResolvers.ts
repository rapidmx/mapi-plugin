///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { encodeAppointmentRecurrence } from "../codec/AppointmentRecurrence.js";
import { BufferWriter } from "../codec/BufferCursor.js";
import { encodeTimeZoneStruct } from "../codec/MapiTimeZone.js";
import { PropertyType, PropertyValueData, writePropertyValue } from "../codec/PropertyValue.js";
import type { MapiSessionContext } from "../MapiSessionManager.js";
import {
    BUSY_STATUS_CODES,
    LID_APPOINTMENT_END_WHOLE,
    LID_APPOINTMENT_RECUR,
    LID_APPOINTMENT_START_WHOLE,
    LID_BUSY_STATUS,
    LID_LOCATION,
    LID_RECURRING,
    LID_REMINDER_DELTA,
    LID_REMINDER_SET,
    LID_RESPONSE_STATUS,
    LID_TIME_ZONE_STRUCT,
    PSETID_APPOINTMENT,
    PSETID_COMMON,
    RESPONSE_STATUS_CODES,
    RESPONSE_STATUS_NONE,
    RESPONSE_STATUS_ORGANIZED,
} from "./CalendarNamedProperties.js";
import { CalendarEventTargetInfo, resolveCalendarEventInfo } from "./CalendarEventTarget.js";
import { ContactTargetInfo, resolveContactInfo } from "./ContactTarget.js";
import { assignOrGetFid, FolderResolutionCache, FolderTargetInfo, resolveFolderInfo } from "./FolderTarget.js";
import { assignOrGetMid, MessageTargetInfo, resolveMessageInfo } from "./MessageTarget.js";
import { resolveNamedProperty } from "./NamedPropertyRegistry.js";
import type { RopContext } from "./RopHandler.js";
import {
    LID_PERCENT_COMPLETE,
    LID_TASK_COMPLETE,
    LID_TASK_DUE_DATE,
    LID_TASK_STATUS,
    PSETID_TASK,
    TASK_STATUS_COMPLETE,
    TASK_STATUS_NOT_STARTED,
} from "./TaskNamedProperties.js";
import { resolveTaskInfo, TaskTargetInfo } from "./TaskTarget.js";

// Well-known folder property IDs this pragmatic subset supports - the small set a real client needs to render
// a folder-hierarchy view. Add more as a real need arises, not speculatively.
const PID_TAG_DISPLAY_NAME = 0x3001;
const PID_TAG_FOLDER_ID = 0x6748;
const PID_TAG_CONTENT_COUNT = 0x3602;
const PID_TAG_CONTENT_UNREAD_COUNT = 0x3603;
const PID_TAG_SUBFOLDERS = 0x360a;

// Well-known message property IDs this pragmatic subset supports - the small set a real client needs to render
// a message list or a single message's metadata.
const PID_TAG_SUBJECT = 0x0037;
const PID_TAG_MESSAGE_FLAGS = 0x0e07;
const PID_TAG_HAS_ATTACHMENTS = 0x0e1b;
const PID_TAG_MESSAGE_DELIVERY_TIME = 0x0e06;
const PID_TAG_MID = 0x674a;
/** `MSGFLAG_READ`, the one `PidTagMessageFlags` bit this pragmatic subset ever sets. */
const MSGFLAG_READ = 0x01;

/** `PidNameKeywords` ([MS-OXPROPS] "Keywords") - Outlook's Categories feature, a `Kind=name` named property
 * (unlike every other named property this pragmatic subset resolves, all `Kind=lid` - see
 * `NamedPropertyRegistry.ts`'s own `PropertyName` shape for why both kinds are supported). Maps directly onto
 * this data model's Gmail-style `Message.labelUids`: Outlook only understands free-text category names, not a
 * `Label.uid` reference, so `resolvePropertyValues` resolves each uid to its `Label.name` before this property
 * is ever read. `PT_MV_UNICODE` (`PtypMultipleString` in this codebase's own enum naming). */
const PSETID_PUBLIC_STRINGS = "00020329-0000-0000-c000-000000000046";
const NAME_KEYWORDS = "Keywords";

// Well-known Contact property IDs this pragmatic subset supports - all plain (non-named) PidTags, per
// [MS-OXOCNTC], since a Contact needs no PSETID/LID lookup the way Calendar/Task properties do.
const PID_TAG_GIVEN_NAME = 0x3a06;
const PID_TAG_SURNAME = 0x3a11;
const PID_TAG_EMAIL_ADDRESS = 0x3003;
const PID_TAG_BUSINESS_TELEPHONE_NUMBER = 0x3a08;
const PID_TAG_COMPANY_NAME = 0x3a16;
const PID_TAG_TITLE = 0x3a17;

// Calendar named-property identity (`(PropertySet GUID, LID)` pairs) this pragmatic subset supports - see
// CalendarNamedProperties.ts's own doc comment. A calendar item's property IDs `>= 0x8000` are resolved back
// to one of these via `NamedPropertyRegistry.resolveNamedProperty` before this switch can run, since MAPI has
// no fixed numeric ID for any of them the way `PidTagSubject` has one.

/** A type-appropriate zero/empty value for a requested property this handler has no real data for - keeps
 * `StandardPropertyRow` encoding valid (a real value of the right type, per `writePropertyValue`'s
 * expectations) without needing to model every property a client could ever ask for. Shared by
 * `RopQueryRowsHandler` (table rows) and `RopGetPropertiesSpecificHandler` (single-object property fetch). */
export function defaultValueForType(propertyType: PropertyType): PropertyValueData {
    switch (propertyType) {
        case PropertyType.PtypBoolean:
            return false;
        case PropertyType.PtypInteger16:
        case PropertyType.PtypInteger32:
        case PropertyType.PtypFloating32:
        case PropertyType.PtypFloating64:
            return 0;
        case PropertyType.PtypInteger64:
            return 0n;
        case PropertyType.PtypTime:
            return new Date(0);
        case PropertyType.PtypGuid:
            return "00000000-0000-0000-0000-000000000000";
        case PropertyType.PtypBinary:
            return Buffer.alloc(0);
        case PropertyType.PtypMultipleInteger32:
        case PropertyType.PtypMultipleString:
        case PropertyType.PtypMultipleString8:
        case PropertyType.PtypMultipleBinary:
            return [];
        case PropertyType.PtypString:
        case PropertyType.PtypString8:
        default:
            return "";
    }
}

/**
 * Encodes `value` as `propertyType` into `writer`, falling back to `defaultValueForType(propertyType)`'s own
 * value if that throws - `RopQueryRowsHandler`/`RopGetPropertiesSpecificHandler`'s "an unsupported property
 * falls back to a type-appropriate default, never an error" contract (see their own doc comments) was only true
 * for a `propertyId` this codebase has no data for; it wasn't actually true for a `propertyId` it resolves fine
 * whose value just doesn't fit the *client's own requested* `propertyType` for that column (`SetColumns` never
 * cross-checks a column's declared type against what its resolver actually returns) - e.g. `PidTagSubject`'s
 * plain string against a client-requested `PtypGuid` column throws inside `encodeGuid()`, well downstream of the
 * `xxxValueFor` switch that only ever looks at `propertyId`. That escaped all the way to `RopDispatcher` and
 * failed the *whole* ROP (discarding every row `RopQueryRows` had already built in that call) instead of just
 * this one column degrading, the same as an actually-unmodeled property already does.
 *
 * Writes into a scratch `BufferWriter` first and only copies its bytes into `writer` once a whole value's worth
 * encoded cleanly - several `PropertyType` cases (`PtypBinary`, every `PtypMultiple*`) write more than one field
 * per value, and a value of the wrong *shape* (not just the wrong type) can throw partway through one of those,
 * which would otherwise leave stray, wrongly-sized bytes in `writer` ahead of every column/row that follows -
 * silent wire-format corruption worse than the original failure.
 *
 * `defaultValueForType(propertyType)` is built to match `writePropertyValue`'s own cases exactly, so the
 * fallback only fails too when `propertyType` is a value neither function's `switch` recognizes at all (there's
 * no way to encode a wire type this codec doesn't implement, and `RopSetColumns` never validates `propertyType`
 * against the known enum) - that still throws, exactly as it already did before this function existed, since a
 * genuinely unknown type isn't the "known type, mismatched value" case this exists to fix.
 */
export function writePropertyValueSafely(writer: BufferWriter, propertyType: PropertyType, value: PropertyValueData): void {
    const scratch = new BufferWriter();
    try {
        writePropertyValue(scratch, propertyType, value);
    } catch {
        const fallback = new BufferWriter();
        writePropertyValue(fallback, propertyType, defaultValueForType(propertyType));
        writer.writeBytes(fallback.toBuffer());
        return;
    }
    writer.writeBytes(scratch.toBuffer());
}

/** Resolves one requested property's value for a `"folder:<uid>"`/`"virtual:<name>"` target. */
export function folderValueFor(
    session: MapiSessionContext,
    propertyId: number,
    propertyType: PropertyType,
    target: string,
    info: FolderTargetInfo,
): PropertyValueData {
    switch (propertyId) {
        case PID_TAG_DISPLAY_NAME:
            return info.displayName;
        case PID_TAG_FOLDER_ID:
            return BigInt(assignOrGetFid(session, target));
        case PID_TAG_CONTENT_COUNT:
            return info.totalCount;
        case PID_TAG_CONTENT_UNREAD_COUNT:
            return info.unreadCount;
        case PID_TAG_SUBFOLDERS:
            return info.hasChildren;
        default:
            return defaultValueForType(propertyType);
    }
}

/** Resolves one requested property's value for a `"message:<uid>"` target. `labelNames` is this message's own
 * `labelUids` already resolved to display names (see `resolvePropertyValues`'s own doc comment) - a plain
 * parameter rather than a repo lookup here, since this function (like every other `xxxValueFor`) is
 * synchronous and can't do its own async DB access. */
export function messageValueFor(
    session: MapiSessionContext,
    propertyId: number,
    propertyType: PropertyType,
    target: string,
    info: MessageTargetInfo,
    labelNames: string[] = [],
): PropertyValueData {
    switch (propertyId) {
        case PID_TAG_SUBJECT:
            return info.subject;
        case PID_TAG_MESSAGE_FLAGS:
            return info.read ? MSGFLAG_READ : 0;
        case PID_TAG_HAS_ATTACHMENTS:
            return info.hasAttachments;
        case PID_TAG_MESSAGE_DELIVERY_TIME:
            return info.receivedDate;
        case PID_TAG_MID:
            return BigInt(assignOrGetMid(session, target));
        default: {
            if (propertyId >= 0x8000) {
                const named = resolveNamedProperty(session, propertyId);
                if (named && named.kind === "name" && named.name === NAME_KEYWORDS && named.guid.toLowerCase() === PSETID_PUBLIC_STRINGS) {
                    return labelNames;
                }
            }
            return defaultValueForType(propertyType);
        }
    }
}

/** Resolves one requested property's value for a `"calendarEvent:<uid>"` target. Almost every Appointment
 * property is a *named* property (`PidLid*`, no fixed numeric ID) rather than a plain `PidTag*` - a property ID
 * `>= 0x8000` is resolved back to its `(PropertySet GUID, LID)` identity via `resolveNamedProperty` before this
 * can dispatch on it (see `NamedPropertyRegistry.ts`'s own doc comment and the architecture plan's "Calendar
 * support" section for the full named-property table this switch implements). */
export function calendarEventValueFor(
    session: MapiSessionContext,
    propertyId: number,
    propertyType: PropertyType,
    target: string,
    info: CalendarEventTargetInfo,
    callerAddress: string,
): PropertyValueData {
    if (propertyId === PID_TAG_SUBJECT) {
        return info.title;
    }
    if (propertyId === PID_TAG_MID) {
        return BigInt(assignOrGetMid(session, target));
    }
    if (propertyId < 0x8000) {
        return defaultValueForType(propertyType);
    }

    const namedProperty = resolveNamedProperty(session, propertyId);
    if (!namedProperty || namedProperty.kind !== "lid") {
        return defaultValueForType(propertyType);
    }
    const guid = namedProperty.guid.toLowerCase();
    const isAppointmentProperty = guid === PSETID_APPOINTMENT;
    const isCommonProperty = guid === PSETID_COMMON;

    switch (namedProperty.lid) {
        case LID_LOCATION:
            return isAppointmentProperty ? (info.location ?? "") : defaultValueForType(propertyType);
        case LID_APPOINTMENT_START_WHOLE:
            return isAppointmentProperty ? info.startDate : defaultValueForType(propertyType);
        case LID_APPOINTMENT_END_WHOLE:
            return isAppointmentProperty ? info.endDate : defaultValueForType(propertyType);
        case LID_BUSY_STATUS:
            return isAppointmentProperty ? BUSY_STATUS_CODES[info.busyStatus] : defaultValueForType(propertyType);
        case LID_RECURRING:
            return isAppointmentProperty ? info.recurrenceRule !== undefined : defaultValueForType(propertyType);
        case LID_APPOINTMENT_RECUR:
            return isAppointmentProperty && info.recurrenceRule
                ? encodeAppointmentRecurrence(info.recurrenceRule, info.startDate, info.endDate)
                : defaultValueForType(propertyType);
        case LID_TIME_ZONE_STRUCT:
            return isAppointmentProperty ? encodeTimeZoneStruct(info.timezone, info.startDate) : defaultValueForType(propertyType);
        case LID_RESPONSE_STATUS:
            return isAppointmentProperty ? responseStatusFor(info, callerAddress) : defaultValueForType(propertyType);
        case LID_REMINDER_SET:
            return isCommonProperty ? info.reminderMinutesBeforeStart != null : defaultValueForType(propertyType);
        case LID_REMINDER_DELTA:
            return isCommonProperty ? (info.reminderMinutesBeforeStart ?? 0) : defaultValueForType(propertyType);
        default:
            return defaultValueForType(propertyType);
    }
}

/** The caller's own `PidLidResponseStatus` value: `respOrganized` if the caller mailbox *is* this event's
 * organizer, else the matching `Attendee.responseStatus` (by address, case-insensitive), else `respNone` if the
 * caller isn't party to this event at all (e.g. a shared/delegate calendar view). */
function responseStatusFor(info: CalendarEventTargetInfo, callerAddress: string): number {
    if (callerAddress && info.organizerAddress.toLowerCase() === callerAddress.toLowerCase()) {
        return RESPONSE_STATUS_ORGANIZED;
    }
    const attendee = info.attendees.find((a) => a.address.toLowerCase() === callerAddress.toLowerCase());
    return attendee ? RESPONSE_STATUS_CODES[attendee.responseStatus] : RESPONSE_STATUS_NONE;
}

/** Resolves one requested property's value for a `"contact:<uid>"` target - a `Folder` of type `CONTACTS`'s
 * content-table rows/opened items. Every field here is a plain fixed `PidTag` (`[MS-OXOCNTC]`), unlike Calendar/
 * Task's named-property tables, since Contact's own well-known fields all have real numeric IDs. */
export function contactValueFor(
    session: MapiSessionContext,
    propertyId: number,
    propertyType: PropertyType,
    target: string,
    info: ContactTargetInfo,
): PropertyValueData {
    switch (propertyId) {
        case PID_TAG_SUBJECT:
            return info.displayName;
        case PID_TAG_MID:
            return BigInt(assignOrGetMid(session, target));
        case PID_TAG_GIVEN_NAME:
            return info.givenName ?? "";
        case PID_TAG_SURNAME:
            return info.surname ?? "";
        case PID_TAG_EMAIL_ADDRESS:
            return info.email ?? "";
        case PID_TAG_BUSINESS_TELEPHONE_NUMBER:
            return info.businessPhone ?? "";
        case PID_TAG_COMPANY_NAME:
            return info.companyName ?? "";
        case PID_TAG_TITLE:
            return info.jobTitle ?? "";
        default:
            return defaultValueForType(propertyType);
    }
}

/** Resolves one requested property's value for a `"task:<uid>"` target - a `Folder` of type `TASKS`'s
 * content-table rows/opened items. Almost every Task-specific property is a *named* property (`PidLid*`) under
 * `PSETID_Task`, resolved back to its `(PropertySet GUID, LID)` identity via `resolveNamedProperty` exactly the
 * same way `calendarEventValueFor` handles `PSETID_Appointment`/`PSETID_Common` - see `TaskNamedProperties.ts`'s
 * own doc comment for the LID table this switch implements. */
export function taskValueFor(
    session: MapiSessionContext,
    propertyId: number,
    propertyType: PropertyType,
    target: string,
    info: TaskTargetInfo,
): PropertyValueData {
    if (propertyId === PID_TAG_SUBJECT) {
        return info.title;
    }
    if (propertyId === PID_TAG_MID) {
        return BigInt(assignOrGetMid(session, target));
    }
    if (propertyId < 0x8000) {
        return defaultValueForType(propertyType);
    }

    const namedProperty = resolveNamedProperty(session, propertyId);
    if (!namedProperty || namedProperty.kind !== "lid" || namedProperty.guid.toLowerCase() !== PSETID_TASK) {
        return defaultValueForType(propertyType);
    }

    switch (namedProperty.lid) {
        case LID_TASK_STATUS:
            return info.completed ? TASK_STATUS_COMPLETE : TASK_STATUS_NOT_STARTED;
        case LID_PERCENT_COMPLETE:
            return info.completed ? 1.0 : 0.0;
        case LID_TASK_DUE_DATE:
            return info.dueDate ?? defaultValueForType(propertyType);
        case LID_TASK_COMPLETE:
            return info.completed;
        default:
            return defaultValueForType(propertyType);
    }
}

/** An optional, purely request-scoped (never persisted) memo shared across every `resolvePropertyValues` call
 * within one ROP handler invocation - e.g. one `RopQueryRows` call resolving many table rows, or the (rarer)
 * case of `RopGetPropertiesSpecific` resolving many columns for one object. Without it, both `allFolders` (via
 * `resolveFolderInfo`'s own `hasChildren` computation) and `callerAddress` (the calendar organizer/attendee
 * lookup below) would be re-fetched from the repo on every single row even though they can only ever have one
 * value for the whole batch (this mailbox's folder list, this mailbox's own primary address) - a real,
 * measurable cost for a hierarchy/calendar table with many rows, not just a theoretical one. A caller that
 * resolves only a single target (`RopGetPropertiesSpecificHandler`) can simply omit this - resolving without a
 * cache is still correct, just not deduplicated across calls it doesn't make anyway. */
export interface ResolutionCache extends FolderResolutionCache {
    callerAddress?: string;
    /** `Label.uid` -> `Label.name`, for the whole mailbox - built at most once per call the same way
     * `allFolders`/`callerAddress` are, the first time a message row actually has a non-empty `labelUids` to
     * resolve (most messages have none, so this stays unbuilt for a page with no labeled messages at all). */
    labelNamesByUid?: Map<string, string>;
}

/** Resolves every column in `columns` for a single `target` (a `"folder:"`/`"virtual:"`/`"message:"`/
 * `"calendarEvent:"`/`"contact:"`/`"task:"` target string), in order - the shared implementation behind both
 * `RopQueryRowsHandler` (one call per table row) and `RopGetPropertiesSpecificHandler` (one call for the single
 * object a handle refers to). */
export async function resolvePropertyValues(
    target: string,
    columns: { propertyId: number; propertyType: PropertyType }[],
    context: Pick<
        RopContext,
        | "mailboxUid"
        | "session"
        | "folderRepo"
        | "messageRepo"
        | "calendarEventRepo"
        | "mailboxRepo"
        | "contactRepo"
        | "taskRepo"
        | "labelRepo"
        | "budget"
    >,
    cache?: ResolutionCache,
): Promise<PropertyValueData[]> {
    // Every row resolved counts against the request's work budget - see ExecuteBudget.
    context.budget?.chargeRows();
    const isMessage = target.startsWith("message:");
    const isCalendarEvent = target.startsWith("calendarEvent:");
    const isContact = target.startsWith("contact:");
    const isTask = target.startsWith("task:");
    const folderInfo =
        isMessage || isCalendarEvent || isContact || isTask
            ? undefined
            : await resolveFolderInfo(context.mailboxUid, target, context.folderRepo, cache, context.budget);
    const messageInfo = isMessage ? await resolveMessageInfo(target, context.messageRepo) : undefined;
    const calendarEventInfo = isCalendarEvent ? await resolveCalendarEventInfo(target, context.calendarEventRepo) : undefined;
    const contactInfo = isContact && context.contactRepo ? await resolveContactInfo(target, context.contactRepo) : undefined;
    const taskInfo = isTask && context.taskRepo ? await resolveTaskInfo(target, context.taskRepo) : undefined;

    let labelNames: string[] = [];
    if (messageInfo && messageInfo.labelUids.length > 0 && context.labelRepo) {
        let labelNamesByUid: Map<string, string>;
        if (cache) {
            if (!cache.labelNamesByUid) {
                const labels = await context.labelRepo.find({ mailboxUid: context.mailboxUid }, { ignoreACL: true });
                cache.labelNamesByUid = new Map(labels.map((l: { uid: string; name: string }) => [l.uid, l.name]));
            }
            labelNamesByUid = cache.labelNamesByUid;
        } else {
            const labels = await context.labelRepo.find({ mailboxUid: context.mailboxUid }, { ignoreACL: true });
            labelNamesByUid = new Map(labels.map((l: { uid: string; name: string }) => [l.uid, l.name]));
        }
        labelNames = messageInfo.labelUids.map((uid) => labelNamesByUid.get(uid)).filter((name): name is string => name !== undefined);
    }

    let callerAddress = "";
    if (isCalendarEvent) {
        if (cache) {
            if (cache.callerAddress === undefined) {
                const mailbox = await context.mailboxRepo.findOne(context.mailboxUid, { ignoreACL: true });
                cache.callerAddress = mailbox?.primarySmtpAddress ?? "";
            }
            callerAddress = cache.callerAddress ?? "";
        } else {
            const mailbox = await context.mailboxRepo.findOne(context.mailboxUid, { ignoreACL: true });
            callerAddress = mailbox?.primarySmtpAddress ?? "";
        }
    }

    return columns.map((column) => {
        if (calendarEventInfo) {
            return calendarEventValueFor(context.session, column.propertyId, column.propertyType, target, calendarEventInfo, callerAddress);
        }
        if (contactInfo) {
            return contactValueFor(context.session, column.propertyId, column.propertyType, target, contactInfo);
        }
        if (taskInfo) {
            return taskValueFor(context.session, column.propertyId, column.propertyType, target, taskInfo);
        }
        if (messageInfo) {
            return messageValueFor(context.session, column.propertyId, column.propertyType, target, messageInfo, labelNames);
        }
        if (isContact || isTask) {
            // contactRepo/taskRepo absent (see RopHandler.ts's own doc comment on why they're optional) - no
            // real data to resolve against, so every column degrades to its type-appropriate default rather
            // than falling through to the folder-info branch below (folderInfo is never populated for these
            // target prefixes, see above - forcing that branch would throw on the `!` assertion).
            return defaultValueForType(column.propertyType);
        }
        return folderValueFor(context.session, column.propertyId, column.propertyType, target, folderInfo!);
    });
}
