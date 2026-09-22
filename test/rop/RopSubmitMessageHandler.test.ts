///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { simpleParser } from "mailparser";
import { BufferReader, BufferWriter } from "../../src/codec/BufferCursor.js";
import { RopSubmitMessageHandler } from "../../src/rop/RopSubmitMessageHandler.js";
import { handleDataStoreOf, type RopContext } from "../../src/rop/RopHandler.js";
import { MapiSessionContext } from "../../src/MapiSessionManager.js";
import { AvVerdict, FolderType, SpamVerdict } from "@rapidmx/restapi";
import { ModelUtils } from "@rapidrest/service-core";
import { InMemoryBlobStore } from "../testDoubles.js";
import { writeStreamKey } from "../../src/rop/RopWriteStreamHandler.js";
import { MAX_RECIPIENTS_PER_MESSAGE } from "../../src/rop/AddressList.js";
import { ExecuteBudget, WorkBudgetExceededError } from "../../src/rop/ExecuteBudget.js";

function buildRequest({ logonId = 0, inputHandleIndex = 5, submitFlags = 0 }): Buffer {
    const writer = new BufferWriter();
    writer.writeUInt8(logonId);
    writer.writeUInt8(inputHandleIndex);
    writer.writeUInt8(submitFlags);
    return writer.toBuffer();
}

function makeScanPipeline() {
    return {
        run: vi.fn().mockResolvedValue({
            spam: { verdict: SpamVerdict.CLEAN },
            av: { verdict: AvVerdict.CLEAN },
            attachments: [],
            // @rapidmx/restapi ^0.3.0's scanAndRelay() unconditionally derives a conversationId from these -
            // a real ScanPipeline.run() always populates them (see restapi's ScanPipelineResult), so this
            // mock must too, or deriveConversationId()'s references[0] throws on undefined.
            references: [],
            inReplyTo: undefined,
        }),
    };
}

function makeMailTransport() {
    return { send: vi.fn().mockResolvedValue({ accepted: ["placeholder"], rejected: [], messageId: "x" }) };
}

function makeContext(overrides: Partial<RopContext> = {}): RopContext {
    return {
        mailboxUid: "mailbox-1",
        userUid: "user-1",
        session: new MapiSessionContext({ mailboxUid: "mailbox-1", userUid: "user-1" }),
        folderRepo: { find: vi.fn().mockResolvedValue([{ uid: "sent-uid", type: FolderType.SENT_ITEMS }]) } as any,
        messageRepo: { create: vi.fn().mockResolvedValue(undefined) } as any,
        calendarEventRepo: {} as any,
        mailboxRepo: { findOne: vi.fn().mockResolvedValue({ uid: "mailbox-1", primarySmtpAddress: "owner@example.com" }) } as any,
        folderClass: class TestFolder {
            public constructor(data: any) {
                Object.assign(this, data);
            }
        },
        messageClass: class TestMessage {
            public constructor(data: any) {
                Object.assign(this, data);
            }
        },
        calendarEventClass: {} as any,
        scanPipeline: makeScanPipeline() as any,
        mailTransport: makeMailTransport() as any,
        blobStore: new InMemoryBlobStore(),
        ...overrides,
    };
}

describe("RopSubmitMessageHandler Tests", () => {
    it("Has RopId 0x32.", () => {
        expect(new RopSubmitMessageHandler().ropId).toBe(0x32);
    });

    it("Returns MAPI_E_INVALID_OBJECT when InputHandleIndex isn't a message handle.", async () => {
        const context = makeContext();
        const handler = new RopSubmitMessageHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({})), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0x80070005);
        expect(response.hasMore()).toBe(false);
    });

    it("Returns MAPI_E_INVALID_PARAMETER when the draft has no recipients at all.", async () => {
        const context = makeContext();
        context.session.handles[5] = {
            type: "message",
            entityUid: "",
            draftFolderUid: "folder:f1",
            draftProperties: { "55": "No Recipients" },
        };
        const handler = new RopSubmitMessageHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({})), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0x80070057);
        expect(context.mailTransport.send).not.toHaveBeenCalled();
    });

    it("Returns MAPI_E_INVALID_PARAMETER when the draft handle has no draftProperties at all.", async () => {
        const context = makeContext();
        context.session.handles[5] = { type: "message", entityUid: "", draftFolderUid: "folder:f1" };
        const handler = new RopSubmitMessageHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({})), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0x80070057);
    });

    it("Succeeds with an empty Subject and body when neither was ever set, only a recipient.", async () => {
        const context = makeContext();
        context.session.handles[5] = {
            type: "message",
            entityUid: "",
            draftFolderUid: "folder:f1",
            draftProperties: { "3588": "to@example.com" },
        };
        const handler = new RopSubmitMessageHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({})), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0); // ReturnValue - success

        const scanPipeline = context.scanPipeline as any;
        const [rawSent] = scanPipeline.run.mock.calls[0];
        const parsed = await simpleParser(rawSent as Buffer);
        expect(parsed.subject).toBeUndefined();
        expect((parsed.text ?? "").trim()).toBe("");
    });

    it("Returns MAPI_E_INVALID_PARAMETER when the mailbox has no resolvable primarySmtpAddress.", async () => {
        const context = makeContext({ mailboxRepo: { findOne: vi.fn().mockResolvedValue(undefined) } as any });
        context.session.handles[5] = {
            type: "message",
            entityUid: "",
            draftFolderUid: "folder:f1",
            draftProperties: { "3588": "to@example.com" },
        };
        const handler = new RopSubmitMessageHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({})), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0x80070057);
    });

    it("Builds and sends a real MIME message from Subject/DisplayTo/DisplayCc/inline Body, then saves a Sent Items copy.", async () => {
        const context = makeContext();
        context.session.handles[5] = {
            type: "message",
            entityUid: "",
            draftFolderUid: "folder:f1",
            draftProperties: {
                "55": "Hello From MAPI",
                "3588": "to@example.com",
                "3587": "cc@example.com",
                "3586": "bcc@example.com",
                "4096": "This is the message body.",
            },
        };
        const handler = new RopSubmitMessageHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({})), writer, context);

        const response = new BufferReader(writer.toBuffer());
        expect(response.readUInt8()).toBe(0x32);
        expect(response.readUInt8()).toBe(5);
        expect(response.readUInt32LE()).toBe(0); // ReturnValue - success
        expect(response.hasMore()).toBe(false);

        const scanPipeline = context.scanPipeline as any;
        expect(scanPipeline.run).toHaveBeenCalledTimes(1);
        const [rawSent, envelope] = scanPipeline.run.mock.calls[0];
        expect(envelope).toEqual({ from: "owner@example.com", to: ["to@example.com", "cc@example.com", "bcc@example.com"] });

        const parsed = await simpleParser(rawSent as Buffer);
        expect(parsed.subject).toBe("Hello From MAPI");
        expect(parsed.text?.trim()).toBe("This is the message body.");
        expect(parsed.from?.value[0]?.address).toBe("owner@example.com");

        expect(context.mailTransport.send).toHaveBeenCalledTimes(1);

        const messageRepo = context.messageRepo as any;
        expect(messageRepo.create).toHaveBeenCalledTimes(1);
        const [savedMessage, options] = messageRepo.create.mock.calls[0];
        expect(savedMessage.folderUid).toBe("sent-uid");
        expect(savedMessage.subject).toBe("Hello From MAPI");
        expect(savedMessage.recipients).toEqual([
            { address: "to@example.com", type: "to" },
            { address: "cc@example.com", type: "cc" },
            { address: "bcc@example.com", type: "bcc" },
        ]);
        expect(options).toEqual({ ignoreACL: true });
    });

    it("Stores the Sent Items copy with the Message-ID, conversation and raw MIME scanAndRelay actually relayed.", async () => {
        const context = makeContext();
        context.session.handles[5] = { type: "message", entityUid: "", draftProperties: { "55": "Tracked", "3588": "to@example.com" } };

        await new RopSubmitMessageHandler().handle(new BufferReader(buildRequest({})), new BufferWriter(), context);

        const relayedRaw: Buffer = context.mailTransport.send.mock.calls[0][0].raw;
        const relayedHeaders = await simpleParser(relayedRaw);
        const [savedMessage] = (context.messageRepo as any).create.mock.calls[0];
        expect(savedMessage.messageId).toBe(relayedHeaders.messageId!.replace(/^<|>$/g, ""));
        expect(savedMessage.messageId).not.toMatch(/@mapi$/);
        expect(typeof savedMessage.conversationId).toBe("string");
        expect(savedMessage).toHaveProperty("encrypted");
        const stored = await (context.blobStore as InMemoryBlobStore).get(savedMessage.bodyBlobKey);
        expect(stored.equals(relayedRaw)).toBe(true);
    });

    it.each([
        ["a CR/LF header injection", "to@example.com\r\nBcc: victim@example.com"],
        ["a display name with a malformed address", "Jane <jane@>"],
        ["an address-shaped value addressparser can't use", "<>"],
    ])("Returns MAPI_E_INVALID_PARAMETER without sending when a recipient is %s.", async (_label, badAddress) => {
        const context = makeContext();
        context.session.handles[5] = { type: "message", entityUid: "", draftProperties: { "3588": `ok@example.com; ${badAddress}` } };
        const writer = new BufferWriter();

        await new RopSubmitMessageHandler().handle(new BufferReader(buildRequest({})), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0x80070057);
        expect((context.scanPipeline as any).run).not.toHaveBeenCalled();
        expect(context.mailTransport.send).not.toHaveBeenCalled();
    });

    it("Sends to resolved 'Name <address>' recipients (split on ';' only) and to bare display names matching one contact.", async () => {
        const contactFind = vi.fn().mockImplementation((query: any) =>
            Promise.resolve(query.displayName?.value === "Ada Lovelace" ? [{ displayName: "Ada Lovelace", emails: [{ address: "not valid" }, { address: "ada@example.com" }] }] : []),
        );
        const context = makeContext({ contactRepo: { find: contactFind } as any });
        context.session.handles[5] = {
            type: "message",
            entityUid: "",
            draftProperties: { "3588": '"Doe, Jane" <jane@example.com>; Ada Lovelace', "3587": "Doe, John <john@example.com>" },
        };
        const writer = new BufferWriter();

        await new RopSubmitMessageHandler().handle(new BufferReader(buildRequest({})), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0);
        expect(contactFind).toHaveBeenCalledWith({ mailboxUid: "mailbox-1", displayName: ModelUtils.literal("Ada Lovelace"), limit: 2 }, { ignoreACL: true, limit: 2 });
        const [rawSent, envelope] = (context.scanPipeline as any).run.mock.calls[0];
        expect(envelope.to).toEqual(["jane@example.com", "ada@example.com", "john@example.com"]);
        const parsed = await simpleParser(rawSent as Buffer);
        expect((parsed.to as any).value).toEqual([
            { address: "jane@example.com", name: "Doe, Jane" },
            { address: "ada@example.com", name: "Ada Lovelace" },
        ]);
        const [savedMessage] = (context.messageRepo as any).create.mock.calls[0];
        expect(savedMessage.recipients).toEqual([
            { address: "jane@example.com", displayName: "Doe, Jane", type: "to" },
            { address: "ada@example.com", displayName: "Ada Lovelace", type: "to" },
            { address: "john@example.com", displayName: "Doe, John", type: "cc" },
        ]);
    });

    it.each([
        ["no contact", []],
        [
            "several contacts",
            [
                { displayName: "Someone Unknown", emails: [{ address: "a@example.com" }] },
                { displayName: "Someone Unknown", emails: [{ address: "b@example.com" }] },
            ],
        ],
        ["a contact without an email", [{ displayName: "Someone Unknown" }]],
        ["only a contact whose name isn't exactly it", [{ displayName: "someone unknown", emails: [{ address: "a@example.com" }] }]],
    ])("Returns MAPI_E_NOT_FOUND, not access denied, for a display name matching %s.", async (_label, contacts) => {
        const context = makeContext({ contactRepo: { find: vi.fn().mockResolvedValue(contacts) } as any });
        context.session.handles[5] = { type: "message", entityUid: "", draftProperties: { "3588": "ok@example.com; Someone Unknown" } };
        const writer = new BufferWriter();

        await new RopSubmitMessageHandler().handle(new BufferReader(buildRequest({})), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0x8004010f);
        expect(context.mailTransport.send).not.toHaveBeenCalled();
    });

    it("Can't resolve a display name without a contacts repo.", async () => {
        const context = makeContext();
        context.session.handles[5] = { type: "message", entityUid: "", draftProperties: { "3588": "to@example.com", "3586": "Someone" } };
        const writer = new BufferWriter();

        await new RopSubmitMessageHandler().handle(new BufferReader(buildRequest({})), writer, context);

        expect(writer.toBuffer().readUInt32LE(2)).toBe(0x8004010f);
    });

    it("Ignores a write stream left over from an earlier message that used the same handle index.", async () => {
        const context = makeContext();
        const staleBytes = Buffer.concat([Buffer.from("Stale body.", "utf16le"), Buffer.from([0, 0])]);
        context.session.handles[5] = {
            type: "message",
            entityUid: "",
            generation: "current",
            draftProperties: { "3588": "to@example.com", "4096": "Current inline body." },
        };
        context.session.handles[6] = { type: "stream", entityUid: "", generation: "stale-stream", writeTargetHandleIndex: 5, writeTargetGeneration: "earlier", writeSize: staleBytes.length };
        await handleDataStoreOf(context).putChunk(writeStreamKey(context.session.uid, 6, "stale-stream"), 0, staleBytes, 60);

        await new RopSubmitMessageHandler().handle(new BufferReader(buildRequest({})), new BufferWriter(), context);

        const [rawSent] = (context.scanPipeline as any).run.mock.calls[0];
        expect((await simpleParser(rawSent as Buffer)).text?.trim()).toBe("Current inline body.");
    });

    it("Prefers a RopWriteStream-accumulated body over an inline PidTagBody draftProperty.", async () => {
        const context = makeContext();
        const streamText = "Body written via RopWriteStream.";
        const streamBytes = Buffer.concat([Buffer.from(streamText, "utf16le"), Buffer.from([0, 0])]);
        context.session.handles[5] = {
            type: "message",
            entityUid: "",
            draftFolderUid: "folder:f1",
            draftProperties: { "55": "Subj", "3588": "to@example.com", "4096": "This inline body must be ignored." },
        };
        context.session.handles[6] = { type: "stream", entityUid: "", generation: "body-stream", writeTargetHandleIndex: 5, writeSize: streamBytes.length };
        await handleDataStoreOf(context).putChunk(writeStreamKey(context.session.uid, 6, "body-stream"), 0, streamBytes, 60);
        const handler = new RopSubmitMessageHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({})), writer, context);

        const scanPipeline = context.scanPipeline as any;
        const [rawSent] = scanPipeline.run.mock.calls[0];
        const parsed = await simpleParser(rawSent as Buffer);
        expect(parsed.text?.trim()).toBe(streamText);
    });

    it("Fails with MAPI_E_CALL_FAILED instead of sending a truncated body when the written stream's chunks are gone.", async () => {
        const context = makeContext();
        context.session.handles[5] = { type: "message", entityUid: "", draftProperties: { "3588": "to@example.com" } };
        context.session.handles[6] = { type: "stream", entityUid: "", generation: "expired-stream", writeTargetHandleIndex: 5, writeSize: 10 };
        const writer = new BufferWriter();

        await new RopSubmitMessageHandler().handle(new BufferReader(buildRequest({})), writer, context);

        expect(writer.toBuffer().readUInt32LE(2)).toBe(0x80004005);
        expect(context.mailTransport.send).not.toHaveBeenCalled();
    });

    it("Attaches a Disposition-Notification-To header and records requestReceipt when PidTagReadReceiptRequested is true.", async () => {
        const context = makeContext();
        context.session.handles[5] = {
            type: "message",
            entityUid: "",
            draftProperties: { "55": "Subj", "3588": "to@example.com", "41": "true" },
        };
        const handler = new RopSubmitMessageHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({})), writer, context);

        const scanPipeline = context.scanPipeline as any;
        const [rawSent] = scanPipeline.run.mock.calls[0];
        const parsed = await simpleParser(rawSent as Buffer);
        expect(parsed.headers.get("disposition-notification-to")).toMatchObject({ value: [{ address: "owner@example.com" }] });

        const messageRepo = context.messageRepo as any;
        const [savedMessage] = messageRepo.create.mock.calls[0];
        expect(savedMessage.requestReceipt).toBe(true);
    });

    it("Does not attach a Disposition-Notification-To header when PidTagReadReceiptRequested is absent.", async () => {
        const context = makeContext();
        context.session.handles[5] = {
            type: "message",
            entityUid: "",
            draftProperties: { "55": "Subj", "3588": "to@example.com" },
        };
        const handler = new RopSubmitMessageHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({})), writer, context);

        const scanPipeline = context.scanPipeline as any;
        const [rawSent] = scanPipeline.run.mock.calls[0];
        const parsed = await simpleParser(rawSent as Buffer);
        expect(parsed.headers.get("disposition-notification-to")).toBeUndefined();

        const messageRepo = context.messageRepo as any;
        const [savedMessage] = messageRepo.create.mock.calls[0];
        expect(savedMessage.requestReceipt).toBe(false);
    });

    it("Defers relay when PidTagDeferredSendTime is in the future, parking the message in Outbox instead of sending.", async () => {
        const context = makeContext();
        const future = new Date(Date.now() + 60 * 60 * 1000);
        context.session.handles[5] = {
            type: "message",
            entityUid: "",
            draftProperties: {
                "55": "Later",
                "3588": "to@example.com",
                "16367": future.toISOString(),
            },
        };
        const handler = new RopSubmitMessageHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({})), writer, context);

        const response = new BufferReader(writer.toBuffer());
        expect(response.readUInt8()).toBe(0x32);
        expect(response.readUInt8()).toBe(5);
        expect(response.readUInt32LE()).toBe(0); // ReturnValue - success

        expect((context.scanPipeline as any).run).not.toHaveBeenCalled();
        expect(context.mailTransport.send).not.toHaveBeenCalled();

        const messageRepo = context.messageRepo as any;
        expect(messageRepo.create).toHaveBeenCalledTimes(1);
        const [savedMessage] = messageRepo.create.mock.calls[0];
        expect(savedMessage.subject).toBe("Later");
        expect(savedMessage.scheduledSendTime).toEqual(future);
    });

    it("Refreshes the Outbox folder's counts, bumping its sync key, when deferred instead of sent.", async () => {
        const notifyFolderCounts = vi.fn().mockResolvedValue(undefined);
        const context = makeContext({
            notifyFolderCounts,
            folderRepo: { find: vi.fn().mockResolvedValue([{ uid: "outbox-uid", type: FolderType.OUTBOX }]) } as any,
        });
        const future = new Date(Date.now() + 60 * 60 * 1000);
        context.session.handles[5] = {
            type: "message",
            entityUid: "",
            draftProperties: { "55": "Later", "3588": "to@example.com", "16367": future.toISOString() },
        };
        const handler = new RopSubmitMessageHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({})), writer, context);

        expect(writer.toBuffer().readUInt32LE(2)).toBe(0); // ReturnValue - success
        expect(notifyFolderCounts).toHaveBeenCalledTimes(1);
        expect(notifyFolderCounts).toHaveBeenCalledWith(["outbox-uid"], { bumpSyncKey: true });
    });

    it("Refreshes the Sent Items folder's counts, bumping its sync key, after a successful send.", async () => {
        const notifyFolderCounts = vi.fn().mockResolvedValue(undefined);
        const context = makeContext({ notifyFolderCounts });
        context.session.handles[5] = { type: "message", entityUid: "", draftProperties: { "55": "Hi", "3588": "to@example.com" } };
        const handler = new RopSubmitMessageHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({})), writer, context);

        expect(writer.toBuffer().readUInt32LE(2)).toBe(0); // ReturnValue - success
        expect(notifyFolderCounts).toHaveBeenCalledTimes(1);
        expect(notifyFolderCounts).toHaveBeenCalledWith(["sent-uid"], { bumpSyncKey: true });
    });

    it("Sends immediately when PidTagDeferredSendTime is already in the past.", async () => {
        const context = makeContext();
        const past = new Date(Date.now() - 60 * 60 * 1000);
        context.session.handles[5] = {
            type: "message",
            entityUid: "",
            draftProperties: {
                "55": "Now",
                "3588": "to@example.com",
                "16367": past.toISOString(),
            },
        };
        const handler = new RopSubmitMessageHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({})), writer, context);

        expect((context.scanPipeline as any).run).toHaveBeenCalledTimes(1);
        expect(context.mailTransport.send).toHaveBeenCalledTimes(1);
    });

    describe("Calendar branch (PidTagMessageClass starts with IPM.Appointment)", () => {
        function makeCalendarEvent(overrides: Record<string, unknown> = {}) {
            return {
                uid: "evt1",
                mailboxUid: "mailbox-1",
                organizer: { address: "Owner@Example.com" },
                icalUid: "abc-123@mapi",
                sequence: 0,
                title: "Standup",
                location: "Room 1",
                startDate: new Date("2026-09-07T14:00:00.000Z"),
                endDate: new Date("2026-09-07T15:00:00.000Z"),
                attendees: [],
                ...overrides,
            };
        }

        it("Returns MAPI_E_INVALID_OBJECT when the handle's entityUid was never assigned a real calendarEvent (Save never succeeded as an Appointment).", async () => {
            const context = makeContext();
            context.session.handles[5] = { type: "message", entityUid: "", draftProperties: { "26": "IPM.Appointment" } };
            const handler = new RopSubmitMessageHandler();
            const writer = new BufferWriter();

            await handler.handle(new BufferReader(buildRequest({})), writer, context);

            const response = new BufferReader(writer.toBuffer());
            response.readUInt8();
            response.readUInt8();
            expect(response.readUInt32LE()).toBe(0x80070005);
        });

        it("Returns MAPI_E_INVALID_OBJECT when the referenced CalendarEvent has since vanished.", async () => {
            const context = makeContext({ calendarEventRepo: { findOne: vi.fn().mockResolvedValue(undefined) } as any });
            context.session.handles[5] = {
                type: "message",
                entityUid: "calendarEvent:gone",
                draftProperties: { "26": "IPM.Appointment" },
            };
            const handler = new RopSubmitMessageHandler();
            const writer = new BufferWriter();

            await handler.handle(new BufferReader(buildRequest({})), writer, context);

            const response = new BufferReader(writer.toBuffer());
            response.readUInt8();
            response.readUInt8();
            expect(response.readUInt32LE()).toBe(0x80070005);
        });

        it("Succeeds without sending anything when the event has no attendees (a private, non-meeting appointment).", async () => {
            const event = makeCalendarEvent({ attendees: [] });
            const context = makeContext({ calendarEventRepo: { findOne: vi.fn().mockResolvedValue(event) } as any });
            context.session.handles[5] = {
                type: "message",
                entityUid: "calendarEvent:evt1",
                draftProperties: { "26": "IPM.Appointment" },
            };
            const handler = new RopSubmitMessageHandler();
            const writer = new BufferWriter();

            await handler.handle(new BufferReader(buildRequest({})), writer, context);

            const response = new BufferReader(writer.toBuffer());
            response.readUInt8();
            response.readUInt8();
            expect(response.readUInt32LE()).toBe(0); // ReturnValue - success
            const scanPipeline = context.scanPipeline as any;
            expect(scanPipeline.run).not.toHaveBeenCalled();
        });

        it("Returns MAPI_E_INVALID_OBJECT without sending when the caller's mailbox can't be resolved to confirm they organize it.", async () => {
            const event = makeCalendarEvent({ attendees: [{ address: "attendee@example.com" }] });
            const context = makeContext({
                calendarEventRepo: { findOne: vi.fn().mockResolvedValue(event) } as any,
                mailboxRepo: { findOne: vi.fn().mockResolvedValue(undefined) } as any,
            });
            context.session.handles[5] = {
                type: "message",
                entityUid: "calendarEvent:evt1",
                draftProperties: { "26": "IPM.Appointment" },
            };
            const handler = new RopSubmitMessageHandler();
            const writer = new BufferWriter();

            await handler.handle(new BufferReader(buildRequest({})), writer, context);

            const response = new BufferReader(writer.toBuffer());
            response.readUInt8();
            response.readUInt8();
            expect(response.readUInt32LE()).toBe(0x80070005);
            const scanPipeline = context.scanPipeline as any;
            expect(scanPipeline.run).not.toHaveBeenCalled();
        });

        it("Refuses to send invites for an attendee's own copy of someone else's meeting.", async () => {
            const event = makeCalendarEvent({ organizer: { address: "boss@example.com" }, attendees: [{ address: "owner@example.com" }, { address: "x@example.com" }] });
            const context = makeContext({ calendarEventRepo: { findOne: vi.fn().mockResolvedValue(event) } as any });
            context.session.handles[5] = { type: "message", entityUid: "calendarEvent:evt1", draftProperties: { "26": "IPM.Appointment" } };
            const writer = new BufferWriter();

            await new RopSubmitMessageHandler().handle(new BufferReader(buildRequest({})), writer, context);

            const response = new BufferReader(writer.toBuffer());
            response.readUInt8();
            response.readUInt8();
            expect(response.readUInt32LE()).toBe(0x80070005);
            expect((context.scanPipeline as any).run).not.toHaveBeenCalled();
            expect(context.mailTransport.send).not.toHaveBeenCalled();
        });

        it("Refuses an event that belongs to another mailbox, even if the organizer address matches.", async () => {
            const event = makeCalendarEvent({ mailboxUid: "mailbox-2", attendees: [{ address: "x@example.com" }] });
            const mailboxRepo = { findOne: vi.fn().mockResolvedValue({ primarySmtpAddress: "owner@example.com" }) };
            const context = makeContext({ calendarEventRepo: { findOne: vi.fn().mockResolvedValue(event) } as any, mailboxRepo: mailboxRepo as any });
            context.session.handles[5] = { type: "message", entityUid: "calendarEvent:evt1", draftProperties: { "26": "IPM.Appointment" } };
            const writer = new BufferWriter();

            await new RopSubmitMessageHandler().handle(new BufferReader(buildRequest({})), writer, context);

            const response = new BufferReader(writer.toBuffer());
            response.readUInt8();
            response.readUInt8();
            expect(response.readUInt32LE()).toBe(0x80070005);
            expect(mailboxRepo.findOne).not.toHaveBeenCalled();
        });

        it("Accepts an organizer address that is one of the mailbox's aliases.", async () => {
            const event = makeCalendarEvent({ organizer: { address: "alias@example.com" }, attendees: [{ address: "good@example.com" }] });
            const mailboxRepo = { findOne: vi.fn().mockResolvedValue({ primarySmtpAddress: "owner@example.com", aliasAddresses: ["Alias@example.com"] }) };
            const context = makeContext({ calendarEventRepo: { findOne: vi.fn().mockResolvedValue(event) } as any, mailboxRepo: mailboxRepo as any });
            context.session.handles[5] = { type: "message", entityUid: "calendarEvent:evt1", draftProperties: { "26": "IPM.Appointment" } };
            const writer = new BufferWriter();

            await new RopSubmitMessageHandler().handle(new BufferReader(buildRequest({})), writer, context);

            expect(writer.toBuffer().readUInt32LE(2)).toBe(0);
        });

        it("Sends no invite itself, leaving it to restapi's MeetingSchedulingJob, and can be submitted again after an edit.", async () => {
            const event = makeCalendarEvent({ attendees: [{ address: "attendee1@example.com" }, { address: "attendee2@example.com" }] });
            const context = makeContext({ calendarEventRepo: { findOne: vi.fn().mockResolvedValue(event) } as any });
            context.session.handles[5] = { type: "message", entityUid: "calendarEvent:evt1", draftProperties: { "26": "IPM.Appointment" } };
            const handler = new RopSubmitMessageHandler();

            for (let i = 0; i < 2; i++) {
                const writer = new BufferWriter();
                await handler.handle(new BufferReader(buildRequest({})), writer, context);
                expect(writer.toBuffer()).toEqual(Buffer.from([0x32, 5, 0, 0, 0, 0]));
            }

            expect((context.scanPipeline as any).run).not.toHaveBeenCalled();
            expect(context.mailTransport.send).not.toHaveBeenCalled();
            expect((context.messageRepo as any).create).not.toHaveBeenCalled();
            expect(context.session.handles[5].submitted).toBeUndefined();
        });

        it("Ignores an unrecognized IPM.Appointment.* subclass exactly the same way (prefix match, not exact match).", async () => {
            const event = makeCalendarEvent({ attendees: [] });
            const context = makeContext({ calendarEventRepo: { findOne: vi.fn().mockResolvedValue(event) } as any });
            context.session.handles[5] = {
                type: "message",
                entityUid: "calendarEvent:evt1",
                draftProperties: { "26": "IPM.Appointment.SomeSubclass" },
            };
            const handler = new RopSubmitMessageHandler();
            const writer = new BufferWriter();

            await handler.handle(new BufferReader(buildRequest({})), writer, context);

            const response = new BufferReader(writer.toBuffer());
            response.readUInt8();
            response.readUInt8();
            expect(response.readUInt32LE()).toBe(0);
        });
    });

    describe("Meeting-response branch (PidTagMessageClass starts with IPM.Schedule.Meeting.Resp.)", () => {
        it("Dispatches to submitMeetingResponse and reports its ReturnValue, without touching the mail compose/send path.", async () => {
            const calendarEventRepo = { find: vi.fn().mockResolvedValue([]) };
            const context = makeContext({ calendarEventRepo: calendarEventRepo as any });
            context.session.handles[5] = {
                type: "message",
                entityUid: "",
                draftProperties: { "26": "IPM.Schedule.Meeting.Resp.Pos" },
            };
            const handler = new RopSubmitMessageHandler();
            const writer = new BufferWriter();

            await handler.handle(new BufferReader(buildRequest({})), writer, context);

            const response = new BufferReader(writer.toBuffer());
            expect(response.readUInt8()).toBe(0x32);
            expect(response.readUInt8()).toBe(5);
            expect(response.readUInt32LE()).toBe(0x80070057); // no PidLidGlobalObjectId: MAPI_E_INVALID_PARAMETER
            expect(response.hasMore()).toBe(false);

            expect(context.mailTransport.send).not.toHaveBeenCalled();
            expect((context.messageRepo as any).create).not.toHaveBeenCalled();
        });
    });
    describe("Round 5: once per draft", () => {
        it("Marks a mail draft submitted, refuses a second submit of it, and deletes its write-stream chunks.", async () => {
            const context = makeContext();
            const streamBytes = Buffer.concat([Buffer.from("Hello", "utf16le"), Buffer.from([0, 0])]);
            context.session.handles[5] = { type: "message", entityUid: "", draftProperties: { "3588": "to@example.com" } };
            context.session.handles[6] = { type: "stream", entityUid: "", generation: "g6", writeTargetHandleIndex: 5, writeSize: streamBytes.length };
            await handleDataStoreOf(context).putChunk(writeStreamKey(context.session.uid, 6, "g6"), 0, streamBytes, 60);
            const handler = new RopSubmitMessageHandler();

            const first = new BufferWriter();
            await handler.handle(new BufferReader(buildRequest({})), first, context);
            const second = new BufferWriter();
            await handler.handle(new BufferReader(buildRequest({})), second, context);

            expect(first.toBuffer().readUInt32LE(2)).toBe(0);
            expect(second.toBuffer().readUInt32LE(2)).toBe(0x80070005);
            expect(context.session.handles[5].submitted).toBe(true);
            expect(context.mailTransport.send).toHaveBeenCalledTimes(1);
            expect((context.messageRepo as any).create).toHaveBeenCalledTimes(1);
            expect(await handleDataStoreOf(context).readChunks(writeStreamKey(context.session.uid, 6, "g6"), streamBytes.length)).toBeUndefined();
        });

        it("Uses up the draft even when the attempt fails, and a failed chunk delete doesn't fail the submit.", async () => {
            const handleData = { readChunks: vi.fn().mockResolvedValue(undefined), delete: vi.fn().mockRejectedValue(new Error("redis down")) };
            const context = makeContext({ handleData: handleData as any });
            context.session.handles[5] = { type: "message", entityUid: "", draftProperties: { "3588": "to@example.com" } };
            context.session.handles[6] = { type: "stream", entityUid: "", generation: "g6", writeTargetHandleIndex: 5, writeSize: 4 };
            const writer = new BufferWriter();

            await new RopSubmitMessageHandler().handle(new BufferReader(buildRequest({})), writer, context);

            expect(writer.toBuffer().readUInt32LE(2)).toBe(0x80004005);
            expect(context.session.handles[5].submitted).toBe(true);
            expect(handleData.delete).toHaveBeenCalled();
        });

        it("Marks a meeting response submitted too.", async () => {
            const context = makeContext({ calendarEventRepo: { find: vi.fn().mockResolvedValue([]) } as any });
            context.session.handles[5] = { type: "message", entityUid: "", draftProperties: { "26": "IPM.Schedule.Meeting.Resp.Neg" } };

            await new RopSubmitMessageHandler().handle(new BufferReader(buildRequest({})), new BufferWriter(), context);

            expect(context.session.handles[5].submitted).toBe(true);
        });

        it("Refuses more than MAX_RECIPIENTS_PER_MESSAGE recipients with MAPI_E_TOO_BIG before looking any name up.", async () => {
            const contactFind = vi.fn().mockResolvedValue([]);
            const context = makeContext({ contactRepo: { find: contactFind } as any });
            const to = Array.from({ length: MAX_RECIPIENTS_PER_MESSAGE }, (_, i) => `r${i}@example.com`).join("; ");
            context.session.handles[5] = { type: "message", entityUid: "", draftProperties: { "3588": to, "3586": "Some Name" } };
            const writer = new BufferWriter();

            await new RopSubmitMessageHandler().handle(new BufferReader(buildRequest({})), writer, context);

            expect(writer.toBuffer().readUInt32LE(2)).toBe(0x80040305);
            expect(contactFind).not.toHaveBeenCalled();
            expect(context.mailTransport.send).not.toHaveBeenCalled();
        });

        it("Charges every submit to the Execute's budget.", async () => {
            const context = makeContext({ budget: new ExecuteBudget(undefined, undefined, { maxSubmits: 0 }) });
            context.session.handles[5] = { type: "message", entityUid: "", draftProperties: { "3588": "to@example.com" } };

            await expect(new RopSubmitMessageHandler().handle(new BufferReader(buildRequest({})), new BufferWriter(), context)).rejects.toBeInstanceOf(WorkBudgetExceededError);
            expect(context.mailTransport.send).not.toHaveBeenCalled();
        });
    });
});
