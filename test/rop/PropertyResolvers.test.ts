///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { encodeAppointmentRecurrence } from "../../src/codec/AppointmentRecurrence.js";
import { BufferReader, BufferWriter } from "../../src/codec/BufferCursor.js";
import { encodeTimeZoneStruct } from "../../src/codec/MapiTimeZone.js";
import { PropertyType, readPropertyValue } from "../../src/codec/PropertyValue.js";
import { MapiSessionContext } from "../../src/MapiSessionManager.js";
import { assignOrGetNamedPropertyId } from "../../src/rop/NamedPropertyRegistry.js";
import { calendarEventValueFor, contactValueFor, defaultValueForType, messageValueFor, resolvePropertyValues, taskValueFor, writePropertyValueSafely } from "../../src/rop/PropertyResolvers.js";
import type { CalendarEventTargetInfo } from "../../src/rop/CalendarEventTarget.js";
import type { MessageTargetInfo } from "../../src/rop/MessageTarget.js";
import type { ContactTargetInfo } from "../../src/rop/ContactTarget.js";
import { LID_PERCENT_COMPLETE, LID_TASK_COMPLETE, LID_TASK_DUE_DATE, LID_TASK_STATUS, PSETID_TASK } from "../../src/rop/TaskNamedProperties.js";
import type { TaskTargetInfo } from "../../src/rop/TaskTarget.js";
import {
    AttendeeResponseStatus,
    AttendeeRole,
    BusyStatus,
    RecurrenceFrequency,
    type Attendee,
} from "@rapidmx/restapi";

const PSETID_APPOINTMENT = "00062002-0000-0000-c000-000000000046";
const PSETID_COMMON = "00062008-0000-0000-c000-000000000046";

const LID_LOCATION = 0x8208;
const LID_APPOINTMENT_START_WHOLE = 0x820d;
const LID_APPOINTMENT_END_WHOLE = 0x820e;
const LID_BUSY_STATUS = 0x8205;
const LID_RECURRING = 0x8223;
const LID_REMINDER_SET = 0x8503;
const LID_REMINDER_DELTA = 0x8501;
const LID_RESPONSE_STATUS = 0x8218;
const LID_APPOINTMENT_RECUR = 0x8216;
const LID_TIME_ZONE_STRUCT = 0x8233;

const PID_TAG_SUBJECT = 0x0037;
const PID_TAG_MID = 0x674a;
const PSETID_PUBLIC_STRINGS = "00020329-0000-0000-c000-000000000046";

function makeSession(): MapiSessionContext {
    return new MapiSessionContext({ mailboxUid: "mailbox-1", userUid: "user-1" });
}

function baseInfo(overrides: Partial<CalendarEventTargetInfo> = {}): CalendarEventTargetInfo {
    return {
        title: "Standup",
        location: "Room 1",
        startDate: new Date("2026-09-07T14:00:00.000Z"),
        endDate: new Date("2026-09-07T15:00:00.000Z"),
        timezone: "UTC",
        busyStatus: BusyStatus.BUSY,
        recurrenceRule: undefined,
        reminderMinutesBeforeStart: undefined,
        organizerAddress: "organizer@example.com",
        attendees: [],
        ...overrides,
    };
}

function namedId(session: MapiSessionContext, guid: string, lid: number): number {
    return assignOrGetNamedPropertyId(session, { guid, kind: "lid", lid });
}

describe("PropertyResolvers Tests", () => {
    describe("calendarEventValueFor", () => {
        it("Resolves PidTagSubject to the event's title.", () => {
            const session = makeSession();
            const value = calendarEventValueFor(session, PID_TAG_SUBJECT, PropertyType.PtypString, "calendarEvent:e1", baseInfo(), "");
            expect(value).toBe("Standup");
        });

        it("Resolves PidTagMid via assignOrGetMid.", () => {
            const session = makeSession();
            const value = calendarEventValueFor(session, PID_TAG_MID, PropertyType.PtypInteger64, "calendarEvent:e1", baseInfo(), "");
            expect(value).toBe(1n);
            expect(session.messageIds["1"]).toBe("calendarEvent:e1");
        });

        it("Falls back to a type-appropriate default for an unrecognized plain (< 0x8000) property tag.", () => {
            const session = makeSession();
            const value = calendarEventValueFor(session, 0x0e07, PropertyType.PtypInteger32, "calendarEvent:e1", baseInfo(), "");
            expect(value).toBe(0);
        });

        it("Falls back to a default when the property ID was never assigned by RopGetPropertyIdsFromNames.", () => {
            const session = makeSession();
            const value = calendarEventValueFor(session, 0x8000, PropertyType.PtypString, "calendarEvent:e1", baseInfo(), "");
            expect(value).toBe("");
        });

        it("Falls back to a default for a Kind=name named property (this pragmatic subset only maps Kind=LID Appointment/Common properties).", () => {
            const session = makeSession();
            const id = assignOrGetNamedPropertyId(session, { guid: PSETID_APPOINTMENT, kind: "name", name: "SomeCustomProp" });
            const value = calendarEventValueFor(session, id, PropertyType.PtypString, "calendarEvent:e1", baseInfo(), "");
            expect(value).toBe("");
        });

        it("Falls back to a default for a recognized LID under the wrong property-set GUID.", () => {
            const session = makeSession();
            const id = namedId(session, "11111111-0000-0000-c000-000000000046", LID_LOCATION);
            const value = calendarEventValueFor(session, id, PropertyType.PtypString, "calendarEvent:e1", baseInfo(), "");
            expect(value).toBe("");
        });

        it("Falls back to a type-appropriate default for an unrecognized LID under a known property set.", () => {
            const session = makeSession();
            const id = namedId(session, PSETID_APPOINTMENT, 0x9999);
            const value = calendarEventValueFor(session, id, PropertyType.PtypInteger32, "calendarEvent:e1", baseInfo(), "");
            expect(value).toBe(0);
        });

        it("Resolves PidLidLocation, defaulting to empty string when unset.", () => {
            const session = makeSession();
            const id = namedId(session, PSETID_APPOINTMENT, LID_LOCATION);
            expect(calendarEventValueFor(session, id, PropertyType.PtypString, "calendarEvent:e1", baseInfo({ location: "Room 9" }), "")).toBe("Room 9");
            expect(calendarEventValueFor(session, id, PropertyType.PtypString, "calendarEvent:e1", baseInfo({ location: undefined }), "")).toBe("");
        });

        it("Resolves PidLidAppointmentStartWhole/EndWhole.", () => {
            const session = makeSession();
            const info = baseInfo();
            const startId = namedId(session, PSETID_APPOINTMENT, LID_APPOINTMENT_START_WHOLE);
            const endId = namedId(session, PSETID_APPOINTMENT, LID_APPOINTMENT_END_WHOLE);
            expect(calendarEventValueFor(session, startId, PropertyType.PtypTime, "calendarEvent:e1", info, "")).toBe(info.startDate);
            expect(calendarEventValueFor(session, endId, PropertyType.PtypTime, "calendarEvent:e1", info, "")).toBe(info.endDate);
        });

        it("Resolves PidLidBusyStatus to the correct MS-OXOCAL wire value for each BusyStatus.", () => {
            const session = makeSession();
            const id = namedId(session, PSETID_APPOINTMENT, LID_BUSY_STATUS);
            const cases: [BusyStatus, number][] = [
                [BusyStatus.FREE, 0],
                [BusyStatus.TENTATIVE, 1],
                [BusyStatus.BUSY, 2],
                [BusyStatus.OUT_OF_OFFICE, 3],
            ];
            for (const [busyStatus, expected] of cases) {
                expect(calendarEventValueFor(session, id, PropertyType.PtypInteger32, "calendarEvent:e1", baseInfo({ busyStatus }), "")).toBe(expected);
            }
        });

        it("Resolves PidLidRecurring to whether recurrenceRule is set.", () => {
            const session = makeSession();
            const id = namedId(session, PSETID_APPOINTMENT, LID_RECURRING);
            const rule = { freq: RecurrenceFrequency.DAILY, interval: 1, exceptions: [] };
            expect(calendarEventValueFor(session, id, PropertyType.PtypBoolean, "calendarEvent:e1", baseInfo({ recurrenceRule: rule }), "")).toBe(true);
            expect(calendarEventValueFor(session, id, PropertyType.PtypBoolean, "calendarEvent:e1", baseInfo({ recurrenceRule: undefined }), "")).toBe(false);
        });

        it("Resolves PidLidAppointmentRecur to the encoded recurrence blob when a recurrenceRule is set.", () => {
            const session = makeSession();
            const id = namedId(session, PSETID_APPOINTMENT, LID_APPOINTMENT_RECUR);
            const rule = { freq: RecurrenceFrequency.DAILY, interval: 1, exceptions: [] };
            const info = baseInfo({ recurrenceRule: rule });
            const value = calendarEventValueFor(session, id, PropertyType.PtypBinary, "calendarEvent:e1", info, "");
            expect(value).toEqual(encodeAppointmentRecurrence(rule, info.startDate, info.endDate));
        });

        it("Falls back to a default for PidLidAppointmentRecur when no recurrenceRule is set.", () => {
            const session = makeSession();
            const id = namedId(session, PSETID_APPOINTMENT, LID_APPOINTMENT_RECUR);
            const value = calendarEventValueFor(session, id, PropertyType.PtypBinary, "calendarEvent:e1", baseInfo({ recurrenceRule: undefined }), "");
            expect(value).toEqual(Buffer.alloc(0));
        });

        it("Resolves PidLidTimeZoneStruct to the encoded timezone blob.", () => {
            const session = makeSession();
            const id = namedId(session, PSETID_APPOINTMENT, LID_TIME_ZONE_STRUCT);
            const info = baseInfo({ timezone: "UTC" });
            const value = calendarEventValueFor(session, id, PropertyType.PtypBinary, "calendarEvent:e1", info, "");
            expect(value).toEqual(encodeTimeZoneStruct("UTC", info.startDate));
        });

        it("Resolves PidLidReminderSet/ReminderDelta.", () => {
            const session = makeSession();
            const setId = namedId(session, PSETID_COMMON, LID_REMINDER_SET);
            const deltaId = namedId(session, PSETID_COMMON, LID_REMINDER_DELTA);
            const withReminder = baseInfo({ reminderMinutesBeforeStart: 15 });
            const withoutReminder = baseInfo({ reminderMinutesBeforeStart: undefined });

            expect(calendarEventValueFor(session, setId, PropertyType.PtypBoolean, "calendarEvent:e1", withReminder, "")).toBe(true);
            expect(calendarEventValueFor(session, setId, PropertyType.PtypBoolean, "calendarEvent:e1", withoutReminder, "")).toBe(false);
            expect(calendarEventValueFor(session, deltaId, PropertyType.PtypInteger32, "calendarEvent:e1", withReminder, "")).toBe(15);
            expect(calendarEventValueFor(session, deltaId, PropertyType.PtypInteger32, "calendarEvent:e1", withoutReminder, "")).toBe(0);
        });

        it("Falls back to a default for PidLidReminderSet/Delta under the wrong property set (PSETID_Appointment instead of PSETID_Common).", () => {
            const session = makeSession();
            const setId = namedId(session, PSETID_APPOINTMENT, LID_REMINDER_SET);
            const deltaId = namedId(session, PSETID_APPOINTMENT, LID_REMINDER_DELTA);
            expect(calendarEventValueFor(session, setId, PropertyType.PtypBoolean, "calendarEvent:e1", baseInfo(), "")).toBe(false);
            expect(calendarEventValueFor(session, deltaId, PropertyType.PtypInteger32, "calendarEvent:e1", baseInfo(), "")).toBe(0);
        });

        it("Falls back to a default for every PSETID_Appointment-only LID under the wrong property set (PSETID_Common instead).", () => {
            const session = makeSession();
            const startId = namedId(session, PSETID_COMMON, LID_APPOINTMENT_START_WHOLE);
            const endId = namedId(session, PSETID_COMMON, LID_APPOINTMENT_END_WHOLE);
            const busyId = namedId(session, PSETID_COMMON, LID_BUSY_STATUS);
            const recurringId = namedId(session, PSETID_COMMON, LID_RECURRING);
            const timeZoneId = namedId(session, PSETID_COMMON, LID_TIME_ZONE_STRUCT);
            const responseId = namedId(session, PSETID_COMMON, LID_RESPONSE_STATUS);
            const info = baseInfo();

            expect(calendarEventValueFor(session, startId, PropertyType.PtypTime, "calendarEvent:e1", info, "")).toEqual(new Date(0));
            expect(calendarEventValueFor(session, endId, PropertyType.PtypTime, "calendarEvent:e1", info, "")).toEqual(new Date(0));
            expect(calendarEventValueFor(session, busyId, PropertyType.PtypInteger32, "calendarEvent:e1", info, "")).toBe(0);
            expect(calendarEventValueFor(session, recurringId, PropertyType.PtypBoolean, "calendarEvent:e1", info, "")).toBe(false);
            expect(calendarEventValueFor(session, timeZoneId, PropertyType.PtypBinary, "calendarEvent:e1", info, "")).toEqual(Buffer.alloc(0));
            expect(calendarEventValueFor(session, responseId, PropertyType.PtypInteger32, "calendarEvent:e1", info, "")).toBe(0);
        });

        describe("PidLidResponseStatus", () => {
            const responseId = (session: MapiSessionContext) => namedId(session, PSETID_APPOINTMENT, LID_RESPONSE_STATUS);

            it("Returns respOrganized (1) when the caller is the organizer.", () => {
                const session = makeSession();
                const id = responseId(session);
                const info = baseInfo({ organizerAddress: "Owner@Example.com" });
                expect(calendarEventValueFor(session, id, PropertyType.PtypInteger32, "calendarEvent:e1", info, "owner@example.com")).toBe(1);
            });

            it("Returns the matching attendee's response status (by address, case-insensitive) when the caller is an attendee.", () => {
                const session = makeSession();
                const id = responseId(session);
                const attendees: Attendee[] = [
                    { address: "Attendee@Example.com", role: AttendeeRole.REQUIRED, responseStatus: AttendeeResponseStatus.ACCEPTED, isOrganizer: false },
                ];
                const info = baseInfo({ attendees });
                expect(calendarEventValueFor(session, id, PropertyType.PtypInteger32, "calendarEvent:e1", info, "attendee@example.com")).toBe(3);
            });

            it("Returns respNone (0) when the caller is neither the organizer nor a listed attendee.", () => {
                const session = makeSession();
                const id = responseId(session);
                expect(calendarEventValueFor(session, id, PropertyType.PtypInteger32, "calendarEvent:e1", baseInfo(), "nobody@example.com")).toBe(0);
            });

            it("Returns respNone (0) when callerAddress is empty (no mailbox address resolved).", () => {
                const session = makeSession();
                const id = responseId(session);
                expect(calendarEventValueFor(session, id, PropertyType.PtypInteger32, "calendarEvent:e1", baseInfo(), "")).toBe(0);
            });
        });
    });

    describe("resolvePropertyValues (calendarEvent dispatch)", () => {
        it("Resolves calendarEvent columns using the mailbox's own primarySmtpAddress as the caller address, without touching folderRepo/messageRepo.", async () => {
            const session = makeSession();
            const calendarEventRepo = {
                findOne: vi.fn().mockResolvedValue({
                    uid: "e1",
                    title: "Standup",
                    startDate: new Date("2026-09-07T14:00:00.000Z"),
                    endDate: new Date("2026-09-07T15:00:00.000Z"),
                    timezone: "UTC",
                    busyStatus: BusyStatus.BUSY,
                    organizer: { address: "organizer@example.com" },
                    attendees: [],
                }),
            };
            const mailboxRepo = { findOne: vi.fn().mockResolvedValue({ primarySmtpAddress: "organizer@example.com" }) };
            const folderRepo = { findOne: vi.fn(), find: vi.fn() };
            const messageRepo = { findOne: vi.fn(), find: vi.fn() };
            const responseId = namedId(session, PSETID_APPOINTMENT, LID_RESPONSE_STATUS);

            const values = await resolvePropertyValues(
                "calendarEvent:e1",
                [
                    { propertyId: PID_TAG_SUBJECT, propertyType: PropertyType.PtypString },
                    { propertyId: responseId, propertyType: PropertyType.PtypInteger32 },
                ],
                {
                    mailboxUid: "mailbox-1",
                    session,
                    folderRepo: folderRepo as any,
                    messageRepo: messageRepo as any,
                    calendarEventRepo: calendarEventRepo as any,
                    mailboxRepo: mailboxRepo as any,
                },
            );

            expect(values).toEqual(["Standup", 1]); // respOrganized, since the mailbox owns organizer@example.com
            expect(calendarEventRepo.findOne).toHaveBeenCalledWith("e1", { ignoreACL: true });
            expect(mailboxRepo.findOne).toHaveBeenCalledWith("mailbox-1", { ignoreACL: true });
            expect(folderRepo.findOne).not.toHaveBeenCalled();
            expect(messageRepo.findOne).not.toHaveBeenCalled();
        });

        it("Resolves callerAddress to empty string when the mailbox lookup itself comes back empty.", async () => {
            const session = makeSession();
            const calendarEventRepo = { findOne: vi.fn().mockResolvedValue(undefined) };
            const mailboxRepo = { findOne: vi.fn().mockResolvedValue(undefined) };
            const responseId = namedId(session, PSETID_APPOINTMENT, LID_RESPONSE_STATUS);

            const values = await resolvePropertyValues(
                "calendarEvent:gone",
                [{ propertyId: responseId, propertyType: PropertyType.PtypInteger32 }],
                {
                    mailboxUid: "mailbox-1",
                    session,
                    folderRepo: {} as any,
                    messageRepo: {} as any,
                    calendarEventRepo: calendarEventRepo as any,
                    mailboxRepo: mailboxRepo as any,
                },
            );

            expect(values).toEqual([0]); // respNone
        });

        it("Fetches the mailbox at most once across multiple calls sharing the same ResolutionCache.", async () => {
            const session = makeSession();
            const calendarEventRepo = { findOne: vi.fn().mockResolvedValue({ uid: "e1", organizer: { address: "organizer@example.com" }, attendees: [] }) };
            const mailboxRepo = { findOne: vi.fn().mockResolvedValue({ primarySmtpAddress: "owner@example.com" }) };
            const responseId = namedId(session, PSETID_APPOINTMENT, LID_RESPONSE_STATUS);
            const cache = {};

            const context = {
                mailboxUid: "mailbox-1",
                session,
                folderRepo: {} as any,
                messageRepo: {} as any,
                calendarEventRepo: calendarEventRepo as any,
                mailboxRepo: mailboxRepo as any,
            };
            await resolvePropertyValues("calendarEvent:e1", [{ propertyId: responseId, propertyType: PropertyType.PtypInteger32 }], context, cache);
            await resolvePropertyValues("calendarEvent:e1", [{ propertyId: responseId, propertyType: PropertyType.PtypInteger32 }], context, cache);

            expect(mailboxRepo.findOne).toHaveBeenCalledTimes(1);
            expect(cache).toEqual({ callerAddress: "owner@example.com" });
        });
    });

    describe("contactValueFor", () => {
        function baseContactInfo(overrides: Partial<ContactTargetInfo> = {}): ContactTargetInfo {
            return {
                displayName: "Jane Doe",
                givenName: "Jane",
                surname: "Doe",
                email: "jane@example.com",
                businessPhone: "555-1234",
                companyName: "Acme",
                jobTitle: "Engineer",
                ...overrides,
            };
        }

        it("Resolves PidTagSubject to the contact's displayName.", () => {
            const session = makeSession();
            expect(contactValueFor(session, PID_TAG_SUBJECT, PropertyType.PtypString, "contact:c1", baseContactInfo())).toBe("Jane Doe");
        });

        it("Resolves PidTagMid via assignOrGetMid.", () => {
            const session = makeSession();
            const value = contactValueFor(session, PID_TAG_MID, PropertyType.PtypInteger64, "contact:c1", baseContactInfo());
            expect(value).toBe(1n);
            expect(session.messageIds["1"]).toBe("contact:c1");
        });

        it("Resolves PidTagGivenName/Surname/EmailAddress/BusinessTelephoneNumber/CompanyName/Title.", () => {
            const session = makeSession();
            const info = baseContactInfo();
            expect(contactValueFor(session, 0x3a06, PropertyType.PtypString, "contact:c1", info)).toBe("Jane");
            expect(contactValueFor(session, 0x3a11, PropertyType.PtypString, "contact:c1", info)).toBe("Doe");
            expect(contactValueFor(session, 0x3003, PropertyType.PtypString, "contact:c1", info)).toBe("jane@example.com");
            expect(contactValueFor(session, 0x3a08, PropertyType.PtypString, "contact:c1", info)).toBe("555-1234");
            expect(contactValueFor(session, 0x3a16, PropertyType.PtypString, "contact:c1", info)).toBe("Acme");
            expect(contactValueFor(session, 0x3a17, PropertyType.PtypString, "contact:c1", info)).toBe("Engineer");
        });

        it("Defaults every optional field to empty string when unset.", () => {
            const session = makeSession();
            const info = baseContactInfo({ givenName: undefined, surname: undefined, email: undefined, businessPhone: undefined, companyName: undefined, jobTitle: undefined });
            expect(contactValueFor(session, 0x3a06, PropertyType.PtypString, "contact:c1", info)).toBe("");
            expect(contactValueFor(session, 0x3a11, PropertyType.PtypString, "contact:c1", info)).toBe("");
            expect(contactValueFor(session, 0x3003, PropertyType.PtypString, "contact:c1", info)).toBe("");
            expect(contactValueFor(session, 0x3a08, PropertyType.PtypString, "contact:c1", info)).toBe("");
            expect(contactValueFor(session, 0x3a16, PropertyType.PtypString, "contact:c1", info)).toBe("");
            expect(contactValueFor(session, 0x3a17, PropertyType.PtypString, "contact:c1", info)).toBe("");
        });

        it("Falls back to a type-appropriate default for an unrecognized property tag.", () => {
            const session = makeSession();
            expect(contactValueFor(session, 0x0e07, PropertyType.PtypInteger32, "contact:c1", baseContactInfo())).toBe(0);
        });
    });

    describe("taskValueFor", () => {
        function baseTaskInfo(overrides: Partial<TaskTargetInfo> = {}): TaskTargetInfo {
            return { title: "Ship it", completed: false, dueDate: new Date("2026-09-10T00:00:00.000Z"), ...overrides };
        }

        it("Resolves PidTagSubject to the task's title.", () => {
            const session = makeSession();
            expect(taskValueFor(session, PID_TAG_SUBJECT, PropertyType.PtypString, "task:t1", baseTaskInfo())).toBe("Ship it");
        });

        it("Resolves PidTagMid via assignOrGetMid.", () => {
            const session = makeSession();
            const value = taskValueFor(session, PID_TAG_MID, PropertyType.PtypInteger64, "task:t1", baseTaskInfo());
            expect(value).toBe(1n);
            expect(session.messageIds["1"]).toBe("task:t1");
        });

        it("Falls back to a default for an unrecognized plain (< 0x8000) property tag.", () => {
            const session = makeSession();
            expect(taskValueFor(session, 0x0e07, PropertyType.PtypInteger32, "task:t1", baseTaskInfo())).toBe(0);
        });

        it("Falls back to a default when the property ID was never assigned by RopGetPropertyIdsFromNames.", () => {
            const session = makeSession();
            expect(taskValueFor(session, 0x8000, PropertyType.PtypString, "task:t1", baseTaskInfo())).toBe("");
        });

        it("Falls back to a default for a recognized LID under the wrong property-set GUID.", () => {
            const session = makeSession();
            const id = namedId(session, "11111111-0000-0000-c000-000000000046", LID_TASK_STATUS);
            expect(taskValueFor(session, id, PropertyType.PtypInteger32, "task:t1", baseTaskInfo())).toBe(0);
        });

        it("Resolves PidLidTaskStatus/PercentComplete/Complete from Task.completed.", () => {
            const session = makeSession();
            const statusId = namedId(session, PSETID_TASK, LID_TASK_STATUS);
            const percentId = namedId(session, PSETID_TASK, LID_PERCENT_COMPLETE);
            const completeId = namedId(session, PSETID_TASK, LID_TASK_COMPLETE);

            expect(taskValueFor(session, statusId, PropertyType.PtypInteger32, "task:t1", baseTaskInfo({ completed: false }))).toBe(0);
            expect(taskValueFor(session, statusId, PropertyType.PtypInteger32, "task:t1", baseTaskInfo({ completed: true }))).toBe(2);
            expect(taskValueFor(session, percentId, PropertyType.PtypFloating64, "task:t1", baseTaskInfo({ completed: false }))).toBe(0.0);
            expect(taskValueFor(session, percentId, PropertyType.PtypFloating64, "task:t1", baseTaskInfo({ completed: true }))).toBe(1.0);
            expect(taskValueFor(session, completeId, PropertyType.PtypBoolean, "task:t1", baseTaskInfo({ completed: false }))).toBe(false);
            expect(taskValueFor(session, completeId, PropertyType.PtypBoolean, "task:t1", baseTaskInfo({ completed: true }))).toBe(true);
        });

        it("Resolves PidLidTaskDueDate, defaulting when unset.", () => {
            const session = makeSession();
            const dueId = namedId(session, PSETID_TASK, LID_TASK_DUE_DATE);
            const dueDate = new Date("2026-09-10T00:00:00.000Z");
            expect(taskValueFor(session, dueId, PropertyType.PtypTime, "task:t1", baseTaskInfo({ dueDate }))).toBe(dueDate);
            expect(taskValueFor(session, dueId, PropertyType.PtypTime, "task:t1", baseTaskInfo({ dueDate: undefined }))).toEqual(new Date(0));
        });

        it("Falls back to a type-appropriate default for an unrecognized LID under PSETID_Task.", () => {
            const session = makeSession();
            const id = namedId(session, PSETID_TASK, 0x9999);
            expect(taskValueFor(session, id, PropertyType.PtypInteger32, "task:t1", baseTaskInfo())).toBe(0);
        });
    });

    describe("resolvePropertyValues (contact/task dispatch)", () => {
        it("Resolves contact columns via contactRepo, without touching folderRepo/messageRepo.", async () => {
            const session = makeSession();
            const contactRepo = { findOne: vi.fn().mockResolvedValue({ uid: "c1", displayName: "Jane Doe", emails: [], phones: [] }) };
            const folderRepo = { findOne: vi.fn(), find: vi.fn() };
            const messageRepo = { findOne: vi.fn(), find: vi.fn() };

            const values = await resolvePropertyValues(
                "contact:c1",
                [{ propertyId: PID_TAG_SUBJECT, propertyType: PropertyType.PtypString }],
                {
                    mailboxUid: "mailbox-1",
                    session,
                    folderRepo: folderRepo as any,
                    messageRepo: messageRepo as any,
                    calendarEventRepo: {} as any,
                    mailboxRepo: {} as any,
                    contactRepo: contactRepo as any,
                },
            );

            expect(values).toEqual(["Jane Doe"]);
            expect(contactRepo.findOne).toHaveBeenCalledWith("c1", { ignoreACL: true });
            expect(folderRepo.findOne).not.toHaveBeenCalled();
            expect(messageRepo.findOne).not.toHaveBeenCalled();
        });

        it("Resolves task columns via taskRepo, without touching folderRepo/messageRepo.", async () => {
            const session = makeSession();
            const taskRepo = { findOne: vi.fn().mockResolvedValue({ uid: "t1", title: "Ship it", completed: true }) };
            const folderRepo = { findOne: vi.fn(), find: vi.fn() };
            const messageRepo = { findOne: vi.fn(), find: vi.fn() };

            const values = await resolvePropertyValues(
                "task:t1",
                [{ propertyId: PID_TAG_SUBJECT, propertyType: PropertyType.PtypString }],
                {
                    mailboxUid: "mailbox-1",
                    session,
                    folderRepo: folderRepo as any,
                    messageRepo: messageRepo as any,
                    calendarEventRepo: {} as any,
                    mailboxRepo: {} as any,
                    taskRepo: taskRepo as any,
                },
            );

            expect(values).toEqual(["Ship it"]);
            expect(taskRepo.findOne).toHaveBeenCalledWith("t1", { ignoreACL: true });
            expect(folderRepo.findOne).not.toHaveBeenCalled();
            expect(messageRepo.findOne).not.toHaveBeenCalled();
        });

        it("Degrades a contact target to type-appropriate defaults when contactRepo is absent from the context.", async () => {
            const session = makeSession();

            const values = await resolvePropertyValues(
                "contact:c1",
                [
                    { propertyId: PID_TAG_SUBJECT, propertyType: PropertyType.PtypString },
                    { propertyId: 0x3602, propertyType: PropertyType.PtypInteger32 },
                ],
                {
                    mailboxUid: "mailbox-1",
                    session,
                    folderRepo: {} as any,
                    messageRepo: {} as any,
                    calendarEventRepo: {} as any,
                    mailboxRepo: {} as any,
                },
            );

            expect(values).toEqual(["", 0]);
        });

        it("Degrades a task target to type-appropriate defaults when taskRepo is absent from the context.", async () => {
            const session = makeSession();

            const values = await resolvePropertyValues(
                "task:t1",
                [{ propertyId: PID_TAG_SUBJECT, propertyType: PropertyType.PtypString }],
                {
                    mailboxUid: "mailbox-1",
                    session,
                    folderRepo: {} as any,
                    messageRepo: {} as any,
                    calendarEventRepo: {} as any,
                    mailboxRepo: {} as any,
                },
            );

            expect(values).toEqual([""]);
        });
    });

    describe("messageValueFor", () => {
        function baseMessageInfo(overrides: Partial<MessageTargetInfo> = {}): MessageTargetInfo {
            return {
                subject: "Hello",
                read: true,
                hasAttachments: false,
                receivedDate: new Date("2026-01-01T00:00:00.000Z"),
                labelUids: [],
                ...overrides,
            };
        }

        it("Resolves PidNameKeywords (a Kind=name named property) to the already-resolved labelNames array.", () => {
            const session = makeSession();
            const id = assignOrGetNamedPropertyId(session, { guid: PSETID_PUBLIC_STRINGS, kind: "name", name: "Keywords" });
            const value = messageValueFor(session, id, PropertyType.PtypMultipleString, "message:m1", baseMessageInfo(), ["Important", "Work"]);
            expect(value).toEqual(["Important", "Work"]);
        });

        it("Defaults labelNames to an empty array when the caller omits it.", () => {
            const session = makeSession();
            const id = assignOrGetNamedPropertyId(session, { guid: PSETID_PUBLIC_STRINGS, kind: "name", name: "Keywords" });
            const value = messageValueFor(session, id, PropertyType.PtypMultipleString, "message:m1", baseMessageInfo());
            expect(value).toEqual([]);
        });

        it("Falls back to a default for a Kind=lid named property under the same GUID (Keywords is specifically Kind=name).", () => {
            const session = makeSession();
            const id = assignOrGetNamedPropertyId(session, { guid: PSETID_PUBLIC_STRINGS, kind: "lid", lid: 1 });
            const value = messageValueFor(session, id, PropertyType.PtypMultipleString, "message:m1", baseMessageInfo(), ["Important"]);
            expect(value).toEqual([]);
        });

        it("Falls back to a default for a Kind=name property under a different GUID with the same name.", () => {
            const session = makeSession();
            const id = assignOrGetNamedPropertyId(session, { guid: "11111111-0000-0000-c000-000000000046", kind: "name", name: "Keywords" });
            const value = messageValueFor(session, id, PropertyType.PtypMultipleString, "message:m1", baseMessageInfo(), ["Important"]);
            expect(value).toEqual([]);
        });

        it("Falls back to a default for an unassigned named property ID (>= 0x8000 but never resolved).", () => {
            const session = makeSession();
            const value = messageValueFor(session, 0x8000, PropertyType.PtypMultipleString, "message:m1", baseMessageInfo(), ["Important"]);
            expect(value).toEqual([]);
        });

        it("Falls back to a default for an unrecognized plain (< 0x8000) property tag.", () => {
            const session = makeSession();
            const value = messageValueFor(session, 0x3007, PropertyType.PtypTime, "message:m1", baseMessageInfo());
            expect(value).toEqual(new Date(0));
        });
    });

    describe("resolvePropertyValues (message label/Keywords resolution)", () => {
        function makeLabelRepo(labels: { uid: string; name: string }[]) {
            return { find: vi.fn().mockResolvedValue(labels) };
        }

        it("Resolves a message's labelUids to display names via labelRepo, fetching the mailbox's labels only once across a shared cache.", async () => {
            const session = makeSession();
            const keywordsId = assignOrGetNamedPropertyId(session, { guid: PSETID_PUBLIC_STRINGS, kind: "name", name: "Keywords" });
            const messageRepo = {
                findOne: vi.fn().mockResolvedValue({ uid: "m1", subject: "Hi", flags: { read: true }, labelUids: ["l1", "l2"] }),
            };
            const labelRepo = makeLabelRepo([
                { uid: "l1", name: "Important" },
                { uid: "l2", name: "Work" },
                { uid: "l3", name: "Unrelated" },
            ]);
            const context = {
                mailboxUid: "mailbox-1",
                session,
                folderRepo: {} as any,
                messageRepo: messageRepo as any,
                calendarEventRepo: {} as any,
                mailboxRepo: {} as any,
                labelRepo: labelRepo as any,
            };
            const cache = {};

            const values1 = await resolvePropertyValues("message:m1", [{ propertyId: keywordsId, propertyType: PropertyType.PtypMultipleString }], context, cache);
            const values2 = await resolvePropertyValues("message:m1", [{ propertyId: keywordsId, propertyType: PropertyType.PtypMultipleString }], context, cache);

            expect(values1).toEqual([["Important", "Work"]]);
            expect(values2).toEqual([["Important", "Work"]]);
            expect(labelRepo.find).toHaveBeenCalledTimes(1);
            expect(labelRepo.find).toHaveBeenCalledWith({ mailboxUid: "mailbox-1" }, { ignoreACL: true });
        });

        it("Fetches labels fresh each call when no cache is provided.", async () => {
            const session = makeSession();
            const keywordsId = assignOrGetNamedPropertyId(session, { guid: PSETID_PUBLIC_STRINGS, kind: "name", name: "Keywords" });
            const messageRepo = { findOne: vi.fn().mockResolvedValue({ uid: "m1", subject: "Hi", labelUids: ["l1"] }) };
            const labelRepo = makeLabelRepo([{ uid: "l1", name: "Important" }]);
            const context = {
                mailboxUid: "mailbox-1",
                session,
                folderRepo: {} as any,
                messageRepo: messageRepo as any,
                calendarEventRepo: {} as any,
                mailboxRepo: {} as any,
                labelRepo: labelRepo as any,
            };

            await resolvePropertyValues("message:m1", [{ propertyId: keywordsId, propertyType: PropertyType.PtypMultipleString }], context);
            await resolvePropertyValues("message:m1", [{ propertyId: keywordsId, propertyType: PropertyType.PtypMultipleString }], context);

            expect(labelRepo.find).toHaveBeenCalledTimes(2);
        });

        it("Skips fetching labels entirely when the message has no labelUids, even with labelRepo present.", async () => {
            const session = makeSession();
            const messageRepo = { findOne: vi.fn().mockResolvedValue({ uid: "m1", subject: "Hi", labelUids: [] }) };
            const labelRepo = makeLabelRepo([]);
            const context = {
                mailboxUid: "mailbox-1",
                session,
                folderRepo: {} as any,
                messageRepo: messageRepo as any,
                calendarEventRepo: {} as any,
                mailboxRepo: {} as any,
                labelRepo: labelRepo as any,
            };

            await resolvePropertyValues("message:m1", [{ propertyId: PID_TAG_SUBJECT, propertyType: PropertyType.PtypString }], context);

            expect(labelRepo.find).not.toHaveBeenCalled();
        });

        it("Degrades to an empty labelNames list when labelRepo is absent from the context, even for a message with labelUids.", async () => {
            const session = makeSession();
            const keywordsId = assignOrGetNamedPropertyId(session, { guid: PSETID_PUBLIC_STRINGS, kind: "name", name: "Keywords" });
            const messageRepo = { findOne: vi.fn().mockResolvedValue({ uid: "m1", subject: "Hi", labelUids: ["l1"] }) };
            const context = {
                mailboxUid: "mailbox-1",
                session,
                folderRepo: {} as any,
                messageRepo: messageRepo as any,
                calendarEventRepo: {} as any,
                mailboxRepo: {} as any,
            };

            const values = await resolvePropertyValues("message:m1", [{ propertyId: keywordsId, propertyType: PropertyType.PtypMultipleString }], context);

            expect(values).toEqual([[]]);
        });

        it("Drops a labelUid that no longer resolves to a real Label (deleted since the message was tagged).", async () => {
            const session = makeSession();
            const keywordsId = assignOrGetNamedPropertyId(session, { guid: PSETID_PUBLIC_STRINGS, kind: "name", name: "Keywords" });
            const messageRepo = { findOne: vi.fn().mockResolvedValue({ uid: "m1", subject: "Hi", labelUids: ["l1", "gone"] }) };
            const labelRepo = makeLabelRepo([{ uid: "l1", name: "Important" }]);
            const context = {
                mailboxUid: "mailbox-1",
                session,
                folderRepo: {} as any,
                messageRepo: messageRepo as any,
                calendarEventRepo: {} as any,
                mailboxRepo: {} as any,
                labelRepo: labelRepo as any,
            };

            const values = await resolvePropertyValues("message:m1", [{ propertyId: keywordsId, propertyType: PropertyType.PtypMultipleString }], context);

            expect(values).toEqual([["Important"]]);
        });
    });

    describe("writePropertyValueSafely Tests", () => {
        it("Writes the value as-is when it already fits the requested PropertyType.", () => {
            const writer = new BufferWriter();

            writePropertyValueSafely(writer, PropertyType.PtypString, "Hello");

            expect(readPropertyValue(new BufferReader(writer.toBuffer()), PropertyType.PtypString)).toBe("Hello");
        });

        it("Falls back to defaultValueForType() when the value doesn't fit the requested PropertyType, e.g. a plain string against PtypGuid.", () => {
            const writer = new BufferWriter();

            writePropertyValueSafely(writer, PropertyType.PtypGuid, "Not a GUID");

            expect(readPropertyValue(new BufferReader(writer.toBuffer()), PropertyType.PtypGuid)).toBe(defaultValueForType(PropertyType.PtypGuid));
        });

        it("Never leaves stray bytes in writer from a failed attempt - only the fallback's own bytes are written.", () => {
            const writer = new BufferWriter();
            writer.writeUInt8(0xaa); // a sentinel byte written before the call, to prove nothing extra sneaks in ahead of it

            writePropertyValueSafely(writer, PropertyType.PtypGuid, "Not a GUID");

            const reader = new BufferReader(writer.toBuffer());
            expect(reader.readUInt8()).toBe(0xaa);
            expect(readPropertyValue(reader, PropertyType.PtypGuid)).toBe(defaultValueForType(PropertyType.PtypGuid));
            expect(reader.hasMore()).toBe(false); // exactly 16 more bytes (one GUID), nothing left over from a partial first attempt
        });

        it("Still throws when even the type-appropriate default can't be encoded, e.g. a PropertyType neither switch recognizes at all.", () => {
            const writer = new BufferWriter();

            expect(() => writePropertyValueSafely(writer, 0x9999 as PropertyType, "anything")).toThrow();
        });
    });
});
