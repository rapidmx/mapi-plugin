///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { encodeAppointmentRecurrence } from "../../src/codec/AppointmentRecurrence.js";
import { BufferReader, BufferWriter } from "../../src/codec/BufferCursor.js";
import { encodeTimeZoneStruct } from "../../src/codec/MapiTimeZone.js";
import { assignOrGetNamedPropertyId } from "../../src/rop/NamedPropertyRegistry.js";
import { RopSaveChangesMessageHandler } from "../../src/rop/RopSaveChangesMessageHandler.js";
import type { RopContext } from "../../src/rop/RopHandler.js";
import { MapiSessionContext } from "../../src/MapiSessionManager.js";
import { BusyStatus, RecipientType, RecurrenceFrequency } from "@rapidmx/restapi";

const PSETID_APPOINTMENT = "00062002-0000-0000-c000-000000000046";
const PSETID_COMMON = "00062008-0000-0000-c000-000000000046";
const LID_LOCATION = 0x8208;
const LID_APPOINTMENT_START_WHOLE = 0x820d;
const LID_APPOINTMENT_END_WHOLE = 0x820e;
const LID_BUSY_STATUS = 0x8205;
const LID_REMINDER_DELTA = 0x8501;
const LID_APPOINTMENT_RECUR = 0x8216;
const LID_TIME_ZONE_STRUCT = 0x8233;

function buildRequest({ logonId = 0, responseHandleIndex = 7, inputHandleIndex = 5, saveFlags = 0 }): Buffer {
    const writer = new BufferWriter();
    writer.writeUInt8(logonId);
    writer.writeUInt8(responseHandleIndex);
    writer.writeUInt8(inputHandleIndex);
    writer.writeUInt8(saveFlags);
    return writer.toBuffer();
}

function makeContext(overrides: Partial<RopContext> = {}): RopContext {
    return {
        mailboxUid: "mailbox-1",
        userUid: "user-1",
        session: new MapiSessionContext({ mailboxUid: "mailbox-1", userUid: "user-1" }),
        folderRepo: {} as any,
        messageRepo: {} as any,
        calendarEventRepo: {} as any,
        mailboxRepo: { findOne: vi.fn().mockResolvedValue({ primarySmtpAddress: "owner@example.com", displayName: "Owner" }) } as any,
        folderClass: {} as any,
        messageClass: {} as any,
        calendarEventClass: class TestCalendarEvent {
            public constructor(data: any) {
                Object.assign(this, data);
            }
        },
        scanPipeline: {} as any,
        mailTransport: {} as any,
        blobStore: {} as any,
        ...overrides,
    };
}

describe("RopSaveChangesMessageHandler Tests", () => {
    it("Has RopId 0x0C.", () => {
        expect(new RopSaveChangesMessageHandler().ropId).toBe(0x0c);
    });

    it("Returns MAPI_E_INVALID_OBJECT when InputHandleIndex isn't a message handle.", async () => {
        const context = makeContext();
        const handler = new RopSaveChangesMessageHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({})), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0x80070005);
        expect(response.hasMore()).toBe(false);
    });

    it("Assigns a session-scoped MID to a fresh draft, stable across repeated saves of the same handle.", async () => {
        const context = makeContext();
        context.session.handles[5] = { type: "message", entityUid: "", draftFolderUid: "folder:f1", draftProperties: {} };
        const handler = new RopSaveChangesMessageHandler();

        const writer1 = new BufferWriter();
        await handler.handle(new BufferReader(buildRequest({})), writer1, context);
        const response1 = new BufferReader(writer1.toBuffer());
        expect(response1.readUInt8()).toBe(0x0c);
        expect(response1.readUInt8()).toBe(7); // ResponseHandleIndex
        expect(response1.readUInt32LE()).toBe(0); // ReturnValue
        expect(response1.readUInt8()).toBe(5); // InputHandleIndex, echoed
        const mid1 = response1.readBigUInt64LE();
        expect(mid1).toBeGreaterThan(0n);
        expect(response1.hasMore()).toBe(false);

        const writer2 = new BufferWriter();
        await handler.handle(new BufferReader(buildRequest({})), writer2, context);
        const response2 = new BufferReader(writer2.toBuffer());
        response2.readUInt8();
        response2.readUInt8();
        response2.readUInt32LE();
        response2.readUInt8();
        const mid2 = response2.readBigUInt64LE();
        expect(mid2).toBe(mid1);
    });

    it("Echoes an already-real message's own MID when re-saving an opened (not freshly created) message.", async () => {
        const context = makeContext();
        context.session.handles[5] = { type: "message", entityUid: "message:m1" };
        const handler = new RopSaveChangesMessageHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({})), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0);
        response.readUInt8();
        expect(response.readBigUInt64LE()).toBeGreaterThan(0n);
        expect(Object.values(context.session.messageIds)).toContain("message:m1");
    });

    describe("Calendar branch (PidTagMessageClass starts with IPM.Appointment)", () => {
        function assignNamed(session: MapiSessionContext, guid: string, lid: number): number {
            return assignOrGetNamedPropertyId(session, { guid, kind: "lid", lid });
        }

        it("Creates a new CalendarEvent from the draft's accumulated properties, mutating the handle to point at it.", async () => {
            const context = makeContext();
            const calendarEventRepo = { create: vi.fn().mockResolvedValue({ uid: "evt1" }) };
            context.calendarEventRepo = calendarEventRepo as any;

            const startId = assignNamed(context.session, PSETID_APPOINTMENT, LID_APPOINTMENT_START_WHOLE);
            const endId = assignNamed(context.session, PSETID_APPOINTMENT, LID_APPOINTMENT_END_WHOLE);
            const locationId = assignNamed(context.session, PSETID_APPOINTMENT, LID_LOCATION);
            const busyId = assignNamed(context.session, PSETID_APPOINTMENT, LID_BUSY_STATUS);
            const deltaId = assignNamed(context.session, PSETID_COMMON, LID_REMINDER_DELTA);
            const recurId = assignNamed(context.session, PSETID_APPOINTMENT, LID_APPOINTMENT_RECUR);
            const tzId = assignNamed(context.session, PSETID_APPOINTMENT, LID_TIME_ZONE_STRUCT);

            const startDate = new Date("2026-09-07T14:00:00.000Z");
            const endDate = new Date("2026-09-07T15:00:00.000Z");
            const rule = { freq: RecurrenceFrequency.DAILY, interval: 1, exceptions: [] };
            const recurBlob = encodeAppointmentRecurrence(rule, startDate, endDate);
            const tzBlob = encodeTimeZoneStruct("UTC", startDate);

            context.session.handles[5] = {
                type: "message",
                entityUid: "",
                draftFolderUid: "folder:cal1",
                draftProperties: {
                    "55": "Standup",
                    "26": "IPM.Appointment",
                    "3588": "attendee1@example.com; attendee2@example.com",
                    "3587": "attendee3@example.com",
                    [String(startId)]: startDate.toISOString(),
                    [String(endId)]: endDate.toISOString(),
                    [String(locationId)]: "Room 1",
                    [String(busyId)]: "2", // olBusy
                    [String(deltaId)]: "15",
                    [String(recurId)]: recurBlob.toString("base64"),
                    [String(tzId)]: tzBlob.toString("base64"),
                },
            };
            const handler = new RopSaveChangesMessageHandler();
            const writer = new BufferWriter();

            await handler.handle(new BufferReader(buildRequest({})), writer, context);

            const response = new BufferReader(writer.toBuffer());
            response.readUInt8();
            response.readUInt8();
            expect(response.readUInt32LE()).toBe(0); // ReturnValue
            response.readUInt8();
            expect(response.readBigUInt64LE()).toBeGreaterThan(0n);

            expect(context.session.handles[5]?.entityUid).toBe("calendarEvent:evt1");
            expect(calendarEventRepo.create).toHaveBeenCalledTimes(1);
            const [created, options] = calendarEventRepo.create.mock.calls[0];
            expect(created.folderUid).toBe("cal1");
            expect(created.mailboxUid).toBe("mailbox-1");
            expect(created.title).toBe("Standup");
            expect(created.location).toBe("Room 1");
            expect(created.startDate).toEqual(startDate);
            expect(created.endDate).toEqual(endDate);
            expect(created.busyStatus).toBe(BusyStatus.BUSY);
            expect(created.recurrenceRule).toEqual(rule);
            expect(created.timezone).toBe("UTC");
            expect(created.reminderMinutesBeforeStart).toBe(15);
            expect(created.organizer).toEqual({ address: "owner@example.com", displayName: "Owner", type: RecipientType.TO });
            expect(created.attendees).toEqual([
                { address: "attendee1@example.com", role: "required", responseStatus: "needsAction", isOrganizer: false },
                { address: "attendee2@example.com", role: "required", responseStatus: "needsAction", isOrganizer: false },
                { address: "attendee3@example.com", role: "required", responseStatus: "needsAction", isOrganizer: false },
            ]);
            expect(created.icalUid).toMatch(/@mapi$/);
            expect(created.sequence).toBe(0);
            expect(options).toEqual({ ignoreACL: true });
        });

        it("Seals what Outlook saves through the record sealer - with the mailbox, on a create and against the stored event on an update.", async () => {
            const sealWrite = vi.fn(async (_type: string, _mailbox: any, patch: any) => {
                patch.title = "Protected event";
                patch.sealedFields = { title: "sealed" };
            });
            const existing = { uid: "evt1", version: 3, title: "Protected event", attendees: [], sealedFields: { title: "sealed" } };
            const calendarEventRepo = { create: vi.fn().mockResolvedValue({ uid: "evt1" }), findOne: vi.fn().mockResolvedValue(existing), update: vi.fn() };
            const context = makeContext({ recordSealer: { sealWrite } as any, calendarEventRepo: calendarEventRepo as any });
            context.session.handles[5] = { type: "message", entityUid: "", draftFolderUid: "folder:cal1", draftProperties: { "55": "Board meeting", "26": "IPM.Appointment" } };
            const handler = new RopSaveChangesMessageHandler();

            await handler.handle(new BufferReader(buildRequest({})), new BufferWriter(), context);

            expect(sealWrite).toHaveBeenCalledWith("calendar", { primarySmtpAddress: "owner@example.com", displayName: "Owner" }, expect.objectContaining({ folderUid: "cal1" }));
            expect(calendarEventRepo.create.mock.calls[0][0]).toMatchObject({ title: "Protected event", sealedFields: { title: "sealed" } });

            context.session.handles[5].draftProperties = { "55": "Board meeting moved", "26": "IPM.Appointment" };
            await handler.handle(new BufferReader(buildRequest({})), new BufferWriter(), context);

            expect(sealWrite).toHaveBeenLastCalledWith("calendar", expect.anything(), expect.objectContaining({ uid: "evt1" }), existing);
            expect(calendarEventRepo.update.mock.calls[0][0]).toMatchObject({ title: "Protected event", sealedFields: { title: "sealed" } });
        });

        it("Falls back to defaults for fields the client never set on a fresh create.", async () => {
            const context = makeContext();
            const calendarEventRepo = { create: vi.fn().mockResolvedValue({ uid: "evt2" }) };
            context.calendarEventRepo = calendarEventRepo as any;
            context.session.handles[5] = {
                type: "message",
                entityUid: "",
                draftFolderUid: "folder:cal1",
                draftProperties: { "26": "IPM.Appointment" },
            };
            const handler = new RopSaveChangesMessageHandler();
            const writer = new BufferWriter();

            await handler.handle(new BufferReader(buildRequest({})), writer, context);

            const [created] = calendarEventRepo.create.mock.calls[0];
            expect(created.title).toBe("");
            expect(created.timezone).toBe("UTC");
            expect(created.busyStatus).toBe(BusyStatus.BUSY);
            expect(created.attendees).toEqual([]);
            expect(created.startDate).toBeInstanceOf(Date);
            expect(created.endDate).toBeInstanceOf(Date);
        });

        it("Updates an already-real CalendarEvent (a re-saved, previously-opened handle) rather than creating a new one.", async () => {
            const context = makeContext();
            const existing = {
                uid: "evt1",
                version: 3,
                startDate: new Date("2026-01-01T00:00:00.000Z"),
                endDate: new Date("2026-01-01T01:00:00.000Z"),
                busyStatus: BusyStatus.FREE,
                timezone: "America/Los_Angeles",
                attendees: [{ address: "existing@example.com", role: "required", responseStatus: "needsAction", isOrganizer: false }],
            };
            const calendarEventRepo = {
                findOne: vi.fn().mockResolvedValue(existing),
                update: vi.fn().mockResolvedValue(undefined),
            };
            context.calendarEventRepo = calendarEventRepo as any;
            context.session.handles[5] = {
                type: "message",
                entityUid: "calendarEvent:evt1",
                draftProperties: { "55": "Updated Title", "26": "IPM.Appointment" },
            };
            const handler = new RopSaveChangesMessageHandler();
            const writer = new BufferWriter();

            await handler.handle(new BufferReader(buildRequest({})), writer, context);

            expect(calendarEventRepo.findOne).toHaveBeenCalledWith("evt1", { ignoreACL: true });
            expect(calendarEventRepo.update).toHaveBeenCalledTimes(1);
            const [delta, existingArg, options] = calendarEventRepo.update.mock.calls[0];
            expect(delta.uid).toBe("evt1");
            expect(delta.version).toBe(3);
            expect(delta.title).toBe("Updated Title");
            expect(delta.startDate).toEqual(existing.startDate); // not re-set this call - falls back to existing
            expect(delta.busyStatus).toBe(existing.busyStatus);
            expect(delta.attendees).toEqual(existing.attendees); // no DisplayTo/Cc set this call
            expect(existingArg).toBe(existing);
            expect(options).toEqual({ ignoreACL: true });

            expect(context.session.handles[5]?.entityUid).toBe("calendarEvent:evt1");
        });

        it("Preserves title/location/recurrenceRule/reminder on an update that only touches an unrelated property (e.g. just PidLidBusyStatus), instead of blanking them.", async () => {
            const context = makeContext();
            const existing = {
                uid: "evt1",
                version: 3,
                title: "Weekly Standup",
                location: "Room 9",
                startDate: new Date("2026-01-01T00:00:00.000Z"),
                endDate: new Date("2026-01-01T01:00:00.000Z"),
                busyStatus: BusyStatus.FREE,
                timezone: "America/Los_Angeles",
                recurrenceRule: { freq: "weekly", interval: 1, exceptions: [] },
                reminderMinutesBeforeStart: 15,
                attendees: [{ address: "existing@example.com", role: "required", responseStatus: "needsAction", isOrganizer: false }],
            };
            const calendarEventRepo = {
                findOne: vi.fn().mockResolvedValue(existing),
                update: vi.fn().mockResolvedValue(undefined),
            };
            context.calendarEventRepo = calendarEventRepo as any;
            const busyStatusId = assignOrGetNamedPropertyId(context.session, { guid: PSETID_APPOINTMENT, kind: "lid", lid: LID_BUSY_STATUS });
            context.session.handles[5] = {
                type: "message",
                entityUid: "calendarEvent:evt1",
                // Only PidTagMessageClass + PidLidBusyStatus this call - Subject/Location/Recur/ReminderDelta
                // are all deliberately absent, exactly the "routine free/busy toggle" scenario this fix covers.
                draftProperties: { "26": "IPM.Appointment", [String(busyStatusId)]: "2" }, // 2 = BUSY
            };
            const handler = new RopSaveChangesMessageHandler();
            const writer = new BufferWriter();

            await handler.handle(new BufferReader(buildRequest({})), writer, context);

            const [delta] = calendarEventRepo.update.mock.calls[0];
            expect(delta.title).toBe(existing.title);
            expect(delta.location).toBe(existing.location);
            expect(delta.recurrenceRule).toEqual(existing.recurrenceRule);
            expect(delta.reminderMinutesBeforeStart).toBe(existing.reminderMinutesBeforeStart);
            expect(delta.busyStatus).toBe(BusyStatus.BUSY); // the one field this call actually changed
        });

        it("Overwrites attendees on an update when DisplayTo/Cc were re-set this call.", async () => {
            const context = makeContext();
            const existing = {
                uid: "evt1",
                version: 1,
                startDate: new Date(),
                endDate: new Date(),
                busyStatus: BusyStatus.FREE,
                timezone: "UTC",
                attendees: [{ address: "old@example.com" }],
            };
            const calendarEventRepo = { findOne: vi.fn().mockResolvedValue(existing), update: vi.fn().mockResolvedValue(undefined) };
            context.calendarEventRepo = calendarEventRepo as any;
            context.session.handles[5] = {
                type: "message",
                entityUid: "calendarEvent:evt1",
                draftProperties: { "26": "IPM.Appointment", "3588": "new@example.com" },
            };
            const handler = new RopSaveChangesMessageHandler();
            const writer = new BufferWriter();

            await handler.handle(new BufferReader(buildRequest({})), writer, context);

            const [delta] = calendarEventRepo.update.mock.calls[0];
            expect(delta.attendees).toEqual([
                { address: "new@example.com", role: "required", responseStatus: "needsAction", isOrganizer: false },
            ]);
        });

        describe("sequence (RFC 5546 SEQUENCE)", () => {
            const baseExisting = () => ({
                uid: "evt1",
                version: 1,
                sequence: 4,
                title: "Sync",
                location: "Room 1",
                startDate: new Date("2026-02-01T10:00:00.000Z"),
                endDate: new Date("2026-02-01T11:00:00.000Z"),
                busyStatus: BusyStatus.BUSY,
                timezone: "UTC",
                attendees: [{ address: "a@example.com", role: "required", responseStatus: "accepted", isOrganizer: false }],
            });

            async function saveWith(existing: any, properties: (session: MapiSessionContext) => Record<string, string>): Promise<any> {
                const context = makeContext();
                const calendarEventRepo = { findOne: vi.fn().mockResolvedValue(existing), update: vi.fn().mockResolvedValue(undefined) };
                context.calendarEventRepo = calendarEventRepo as any;
                context.session.handles[5] = { type: "message", entityUid: "calendarEvent:evt1", draftProperties: { "26": "IPM.Appointment", ...properties(context.session) } };
                await new RopSaveChangesMessageHandler().handle(new BufferReader(buildRequest({})), new BufferWriter(), context);
                return calendarEventRepo.update.mock.calls[0][0];
            }

            it("Keeps the sequence for an edit attendees don't need to hear about (title, busy status).", async () => {
                const delta = await saveWith(baseExisting(), (session) => ({
                    "55": "Renamed",
                    [String(assignOrGetNamedPropertyId(session, { guid: PSETID_APPOINTMENT, kind: "lid", lid: LID_BUSY_STATUS }))]: "0",
                }));
                expect(delta.sequence).toBe(4);
            });

            it("Bumps the sequence when the start time moves.", async () => {
                const delta = await saveWith(baseExisting(), (session) => ({
                    [String(assignOrGetNamedPropertyId(session, { guid: PSETID_APPOINTMENT, kind: "lid", lid: LID_APPOINTMENT_START_WHOLE }))]:
                        "2026-02-01T12:00:00.000Z",
                }));
                expect(delta.sequence).toBe(5);
            });

            it("Bumps the sequence when the location changes.", async () => {
                const delta = await saveWith(baseExisting(), (session) => ({
                    [String(assignOrGetNamedPropertyId(session, { guid: PSETID_APPOINTMENT, kind: "lid", lid: LID_LOCATION }))]: "Room 2",
                }));
                expect(delta.sequence).toBe(5);
            });

            it("Bumps the sequence when the invite list changes, but keeps a still-invited attendee's recorded response.", async () => {
                const delta = await saveWith(baseExisting(), () => ({ "3588": "A@example.com; b@example.com" }));
                expect(delta.sequence).toBe(5);
                expect(delta.attendees[0].responseStatus).toBe("accepted");
                expect(delta.attendees[1]).toEqual({ address: "b@example.com", role: "required", responseStatus: "needsAction", isOrganizer: false });
            });

            it("Keeps the sequence when the same attendees are re-sent, and treats a missing stored sequence as 0.", async () => {
                const { sequence: _sequence, ...withoutSequence } = baseExisting();
                const delta = await saveWith(withoutSequence, () => ({ "3588": "a@example.com" }));
                expect(delta.sequence).toBe(0);
            });

            it("Bumps the sequence when the end time or recurrence changes.", async () => {
                const endDelta = await saveWith(baseExisting(), (session) => ({
                    [String(assignOrGetNamedPropertyId(session, { guid: PSETID_APPOINTMENT, kind: "lid", lid: LID_APPOINTMENT_END_WHOLE }))]:
                        "2026-02-01T11:30:00.000Z",
                }));
                expect(endDelta.sequence).toBe(5);

                const recurrence = encodeAppointmentRecurrence(
                    { freq: RecurrenceFrequency.DAILY, interval: 1, exceptions: [] },
                    new Date("2026-02-01T10:00:00.000Z"),
                    new Date("2026-02-01T11:00:00.000Z"),
                );
                const recurDelta = await saveWith(baseExisting(), (session) => ({
                    [String(assignOrGetNamedPropertyId(session, { guid: PSETID_APPOINTMENT, kind: "lid", lid: LID_APPOINTMENT_RECUR }))]:
                        recurrence.toString("base64"),
                }));
                expect(recurDelta.sequence).toBe(5);
            });
        });

        it("Persists a bare folder uid (not 'folder:'-prefixed) as-is on create when draftFolderUid was never set.", async () => {
            const context = makeContext();
            const calendarEventRepo = { create: vi.fn().mockResolvedValue({ uid: "evt3" }) };
            context.calendarEventRepo = calendarEventRepo as any;
            context.session.handles[5] = { type: "message", entityUid: "", draftProperties: { "26": "IPM.Appointment" } }; // no draftFolderUid
            const handler = new RopSaveChangesMessageHandler();
            const writer = new BufferWriter();

            await handler.handle(new BufferReader(buildRequest({})), writer, context);

            const [created] = calendarEventRepo.create.mock.calls[0];
            expect(created.folderUid).toBe("");
        });

        it("Falls back to an empty organizer address/undefined displayName when the mailbox itself can't be resolved.", async () => {
            const context = makeContext({ mailboxRepo: { findOne: vi.fn().mockResolvedValue(undefined) } as any });
            const calendarEventRepo = { create: vi.fn().mockResolvedValue({ uid: "evt4" }) };
            context.calendarEventRepo = calendarEventRepo as any;
            context.session.handles[5] = {
                type: "message",
                entityUid: "",
                draftFolderUid: "folder:cal1",
                draftProperties: { "26": "IPM.Appointment" },
            };
            const handler = new RopSaveChangesMessageHandler();
            const writer = new BufferWriter();

            await handler.handle(new BufferReader(buildRequest({})), writer, context);

            const [created] = calendarEventRepo.create.mock.calls[0];
            expect(created.organizer).toEqual({ address: "", displayName: undefined, type: RecipientType.TO });
        });

        it("Ignores a named-property ID present in draftProperties that was never actually assigned by RopGetPropertyIdsFromNames.", async () => {
            const context = makeContext();
            const calendarEventRepo = { create: vi.fn().mockResolvedValue({ uid: "evt5" }) };
            context.calendarEventRepo = calendarEventRepo as any;
            context.session.handles[5] = {
                type: "message",
                entityUid: "",
                draftFolderUid: "folder:cal1",
                draftProperties: { "26": "IPM.Appointment", "33000": "orphaned-value" },
            };
            const handler = new RopSaveChangesMessageHandler();
            const writer = new BufferWriter();

            await handler.handle(new BufferReader(buildRequest({})), writer, context);

            expect(calendarEventRepo.create).toHaveBeenCalledTimes(1);
        });

        it("Ignores a Kind=name named property (only Kind=LID Calendar properties are decoded).", async () => {
            const context = makeContext();
            const calendarEventRepo = { create: vi.fn().mockResolvedValue({ uid: "evt6" }) };
            context.calendarEventRepo = calendarEventRepo as any;
            const nameId = assignOrGetNamedPropertyId(context.session, { guid: PSETID_APPOINTMENT, kind: "name", name: "Custom" });
            context.session.handles[5] = {
                type: "message",
                entityUid: "",
                draftFolderUid: "folder:cal1",
                draftProperties: { "26": "IPM.Appointment", [String(nameId)]: "ignored" },
            };
            const handler = new RopSaveChangesMessageHandler();
            const writer = new BufferWriter();

            await handler.handle(new BufferReader(buildRequest({})), writer, context);

            expect(calendarEventRepo.create).toHaveBeenCalledTimes(1);
        });

        it("Ignores a recognized PSETID_Appointment LID this pragmatic subset doesn't decode (e.g. PidLidRecurring, a purely-derived read-side flag).", async () => {
            const context = makeContext();
            const calendarEventRepo = { create: vi.fn().mockResolvedValue({ uid: "evt7" }) };
            context.calendarEventRepo = calendarEventRepo as any;
            const recurringId = assignOrGetNamedPropertyId(context.session, { guid: PSETID_APPOINTMENT, kind: "lid", lid: 0x8223 }); // PidLidRecurring
            context.session.handles[5] = {
                type: "message",
                entityUid: "",
                draftFolderUid: "folder:cal1",
                draftProperties: { "26": "IPM.Appointment", [String(recurringId)]: "true" },
            };
            const handler = new RopSaveChangesMessageHandler();
            const writer = new BufferWriter();

            await handler.handle(new BufferReader(buildRequest({})), writer, context);

            expect(calendarEventRepo.create).toHaveBeenCalledTimes(1);
        });

        it("Falls back to BusyStatus.BUSY for an unrecognized PidLidBusyStatus wire value.", async () => {
            const context = makeContext();
            const calendarEventRepo = { create: vi.fn().mockResolvedValue({ uid: "evt8" }) };
            context.calendarEventRepo = calendarEventRepo as any;
            const busyId = assignOrGetNamedPropertyId(context.session, { guid: PSETID_APPOINTMENT, kind: "lid", lid: LID_BUSY_STATUS });
            context.session.handles[5] = {
                type: "message",
                entityUid: "",
                draftFolderUid: "folder:cal1",
                draftProperties: { "26": "IPM.Appointment", [String(busyId)]: "99" },
            };
            const handler = new RopSaveChangesMessageHandler();
            const writer = new BufferWriter();

            await handler.handle(new BufferReader(buildRequest({})), writer, context);

            const [created] = calendarEventRepo.create.mock.calls[0];
            expect(created.busyStatus).toBe(BusyStatus.BUSY);
        });

        it("Ignores a PSETID_Common LID other than PidLidReminderDelta.", async () => {
            const context = makeContext();
            const calendarEventRepo = { create: vi.fn().mockResolvedValue({ uid: "evt9" }) };
            context.calendarEventRepo = calendarEventRepo as any;
            const setId = assignOrGetNamedPropertyId(context.session, { guid: PSETID_COMMON, kind: "lid", lid: 0x8503 }); // PidLidReminderSet, not Delta
            context.session.handles[5] = {
                type: "message",
                entityUid: "",
                draftFolderUid: "folder:cal1",
                draftProperties: { "26": "IPM.Appointment", [String(setId)]: "true" },
            };
            const handler = new RopSaveChangesMessageHandler();
            const writer = new BufferWriter();

            await handler.handle(new BufferReader(buildRequest({})), writer, context);

            const [created] = calendarEventRepo.create.mock.calls[0];
            expect(created.reminderMinutesBeforeStart).toBeUndefined();
        });

        it("Does nothing beyond assigning a MID if the referenced CalendarEvent has since vanished on an update.", async () => {
            const context = makeContext();
            const calendarEventRepo = { findOne: vi.fn().mockResolvedValue(undefined), update: vi.fn() };
            context.calendarEventRepo = calendarEventRepo as any;
            context.session.handles[5] = {
                type: "message",
                entityUid: "calendarEvent:gone",
                draftProperties: { "26": "IPM.Appointment" },
            };
            const handler = new RopSaveChangesMessageHandler();
            const writer = new BufferWriter();

            await handler.handle(new BufferReader(buildRequest({})), writer, context);

            const response = new BufferReader(writer.toBuffer());
            response.readUInt8();
            response.readUInt8();
            expect(response.readUInt32LE()).toBe(0);
            expect(calendarEventRepo.update).not.toHaveBeenCalled();
        });
    });
});
