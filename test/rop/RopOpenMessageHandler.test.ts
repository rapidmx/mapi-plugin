///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { BufferReader, BufferWriter } from "../../src/codec/BufferCursor.js";
import { readTypedString } from "../../src/codec/TypedString.js";
import { RopOpenMessageHandler } from "../../src/rop/RopOpenMessageHandler.js";
import type { RopContext } from "../../src/rop/RopHandler.js";
import { MapiSessionContext } from "../../src/MapiSessionManager.js";

function buildRequest({
    logonId = 0,
    inputHandleIndex = 0,
    outputHandleIndex = 1,
    codePageId = 0,
    folderId = 0n,
    openModeFlags = 0,
    messageId = 1n,
}): Buffer {
    const writer = new BufferWriter();
    writer.writeUInt8(logonId);
    writer.writeUInt8(inputHandleIndex);
    writer.writeUInt8(outputHandleIndex);
    writer.writeUInt16LE(codePageId);
    writer.writeBigUInt64LE(folderId);
    writer.writeUInt8(openModeFlags);
    writer.writeBigUInt64LE(messageId);
    return writer.toBuffer();
}

function makeContext(messageIds: Record<string, string>, messageRepo: any = {}): RopContext {
    const session = new MapiSessionContext({ mailboxUid: "mailbox-1", userUid: "user-1" });
    session.messageIds = messageIds;
    return {
        mailboxUid: "mailbox-1",
        userUid: "user-1",
        session,
        folderRepo: {} as any,
        messageRepo,
        calendarEventRepo: {} as any,
        mailboxRepo: {} as any,
        folderClass: {} as any,
        messageClass: {} as any,
        calendarEventClass: {} as any,
        scanPipeline: {} as any,
        mailTransport: {} as any,
        blobStore: {} as any,
    };
}

describe("RopOpenMessageHandler Tests", () => {
    it("Has RopId 0x03.", () => {
        expect(new RopOpenMessageHandler().ropId).toBe(0x03);
    });

    it("Opens a known MID, creating a message handle and returning a well-formed success response.", async () => {
        const messageRepo = { findOne: vi.fn().mockResolvedValue({ uid: "msg1", subject: "Hello World" }) };
        const context = makeContext({ "1": "message:msg1" }, messageRepo);
        const handler = new RopOpenMessageHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({ outputHandleIndex: 5, messageId: 1n })), writer, context);

        const response = new BufferReader(writer.toBuffer());
        expect(response.readUInt8()).toBe(0x03); // RopId
        expect(response.readUInt8()).toBe(5); // OutputHandleIndex
        expect(response.readUInt32LE()).toBe(0); // ReturnValue
        expect(response.readUInt8()).toBe(0); // HasNamedProperties
        expect(readTypedString(response)).toBeUndefined(); // SubjectPrefix
        expect(readTypedString(response)).toBe("Hello World"); // NormalizedSubject
        expect(response.readUInt16LE()).toBe(0); // RecipientCount
        expect(response.readUInt16LE()).toBe(0); // ColumnCount
        expect(response.readUInt8()).toBe(0); // RowCount
        expect(response.hasMore()).toBe(false);

        expect(context.session.handles[5]).toMatchObject({ type: "message", entityUid: "message:msg1" });
        expect(messageRepo.findOne).toHaveBeenCalledWith("msg1", { ignoreACL: true });
    });

    it("Opens a known MID for a calendarEvent target, using the event's title as the subject.", async () => {
        const context = makeContext({ "1": "calendarEvent:evt1" });
        context.calendarEventRepo = { findOne: vi.fn().mockResolvedValue({ uid: "evt1", title: "Standup" }) } as any;
        const handler = new RopOpenMessageHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({ outputHandleIndex: 5, messageId: 1n })), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8(); // RopId
        response.readUInt8(); // OutputHandleIndex
        expect(response.readUInt32LE()).toBe(0); // ReturnValue
        response.readUInt8(); // HasNamedProperties
        expect(readTypedString(response)).toBeUndefined(); // SubjectPrefix
        expect(readTypedString(response)).toBe("Standup"); // NormalizedSubject

        expect(context.session.handles[5]).toMatchObject({ type: "message", entityUid: "calendarEvent:evt1" });
        expect((context.calendarEventRepo as any).findOne).toHaveBeenCalledWith("evt1", { ignoreACL: true });
    });

    it("Opens a known MID for a contact target, using the contact's displayName as the subject.", async () => {
        const context = makeContext({ "1": "contact:c1" });
        context.contactRepo = { findOne: vi.fn().mockResolvedValue({ uid: "c1", displayName: "Jane Doe" }) } as any;
        const handler = new RopOpenMessageHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({ outputHandleIndex: 5, messageId: 1n })), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8(); // RopId
        response.readUInt8(); // OutputHandleIndex
        expect(response.readUInt32LE()).toBe(0); // ReturnValue
        response.readUInt8(); // HasNamedProperties
        expect(readTypedString(response)).toBeUndefined(); // SubjectPrefix
        expect(readTypedString(response)).toBe("Jane Doe"); // NormalizedSubject

        expect(context.session.handles[5]).toMatchObject({ type: "message", entityUid: "contact:c1" });
        expect((context.contactRepo as any).findOne).toHaveBeenCalledWith("c1", { ignoreACL: true });
    });

    it("Opens a contact target as an empty subject when contactRepo is absent from the context.", async () => {
        const context = makeContext({ "1": "contact:c1" });
        const handler = new RopOpenMessageHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({ outputHandleIndex: 5, messageId: 1n })), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        response.readUInt32LE();
        response.readUInt8();
        expect(readTypedString(response)).toBeUndefined(); // SubjectPrefix
        expect(readTypedString(response)).toBe(""); // NormalizedSubject falls back to empty
    });

    it("Opens a known MID for a task target, using the task's title as the subject.", async () => {
        const context = makeContext({ "1": "task:t1" });
        context.taskRepo = { findOne: vi.fn().mockResolvedValue({ uid: "t1", title: "Ship it" }) } as any;
        const handler = new RopOpenMessageHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({ outputHandleIndex: 5, messageId: 1n })), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0);
        response.readUInt8();
        expect(readTypedString(response)).toBeUndefined();
        expect(readTypedString(response)).toBe("Ship it");

        expect(context.session.handles[5]).toMatchObject({ type: "message", entityUid: "task:t1" });
        expect((context.taskRepo as any).findOne).toHaveBeenCalledWith("t1", { ignoreACL: true });
    });

    it("Opens a task target as an empty subject when taskRepo is absent from the context.", async () => {
        const context = makeContext({ "1": "task:t1" });
        const handler = new RopOpenMessageHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({ outputHandleIndex: 5, messageId: 1n })), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        response.readUInt32LE();
        response.readUInt8();
        expect(readTypedString(response)).toBeUndefined(); // SubjectPrefix
        expect(readTypedString(response)).toBe(""); // NormalizedSubject falls back to empty
    });

    it("Returns MAPI_E_NOT_FOUND for an unrecognized MID, without creating a handle.", async () => {
        const context = makeContext({});
        const handler = new RopOpenMessageHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({ outputHandleIndex: 5, messageId: 999n })), writer, context);

        const response = new BufferReader(writer.toBuffer());
        expect(response.readUInt8()).toBe(0x03);
        expect(response.readUInt8()).toBe(5);
        expect(response.readUInt32LE()).toBe(0x8004010f);
        expect(response.hasMore()).toBe(false);
        expect(context.session.handles[5]).toBeUndefined();
    });
});
