///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { BufferReader, BufferWriter } from "../../src/codec/BufferCursor.js";
import { readWriteStream, RopWriteStreamHandler } from "../../src/rop/RopWriteStreamHandler.js";
import type { RopContext } from "../../src/rop/RopHandler.js";
import { MapiSessionContext } from "../../src/MapiSessionManager.js";

function buildRequest({ logonId = 0, inputHandleIndex = 6, data = Buffer.alloc(0) }): Buffer {
    const writer = new BufferWriter();
    writer.writeUInt8(logonId);
    writer.writeUInt8(inputHandleIndex);
    writer.writeUInt16LE(data.length);
    writer.writeBytes(data);
    return writer.toBuffer();
}

function makeContext(): RopContext {
    return {
        mailboxUid: "mailbox-1",
        userUid: "user-1",
        session: new MapiSessionContext({ mailboxUid: "mailbox-1", userUid: "user-1" }),
        folderRepo: {} as any,
        messageRepo: {} as any,
        mailboxRepo: {} as any,
        folderClass: {} as any,
        messageClass: {} as any,
        scanPipeline: {} as any,
        mailTransport: {} as any,
        blobStore: {} as any,
    };
}

describe("RopWriteStreamHandler Tests", () => {
    it("Has RopId 0x2D.", () => {
        expect(new RopWriteStreamHandler().ropId).toBe(0x2d);
    });

    it("Returns MAPI_E_INVALID_OBJECT (with WrittenSize 0) when InputHandleIndex isn't a write-mode stream.", async () => {
        const context = makeContext();
        context.session.handles[6] = { type: "stream", entityUid: "message:m1", streamPosition: 0 };
        const handler = new RopWriteStreamHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({ data: Buffer.from("hi") })), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0x80070005);
        expect(response.readUInt16LE()).toBe(0);
        expect(response.hasMore()).toBe(false);
    });

    it("Accumulates bytes across multiple calls, correctly reassembling regardless of chunk boundaries.", async () => {
        const context = makeContext();
        context.session.handles[6] = { type: "stream", entityUid: "", generation: "accumulate", writeTargetHandleIndex: 3 };
        const handler = new RopWriteStreamHandler();

        const fullText = "Hello, world! This is a body written in two chunks.";
        const fullBytes = Buffer.concat([Buffer.from(fullText, "utf16le"), Buffer.from([0, 0])]);
        const chunk1 = fullBytes.subarray(0, 7); // deliberately not aligned to a 2-byte UTF-16 boundary
        const chunk2 = fullBytes.subarray(7);

        const writer1 = new BufferWriter();
        await handler.handle(new BufferReader(buildRequest({ data: chunk1 })), writer1, context);
        const response1 = new BufferReader(writer1.toBuffer());
        response1.readUInt8();
        response1.readUInt8();
        expect(response1.readUInt32LE()).toBe(0);
        expect(response1.readUInt16LE()).toBe(chunk1.length);

        const writer2 = new BufferWriter();
        await handler.handle(new BufferReader(buildRequest({ data: chunk2 })), writer2, context);
        const response2 = new BufferReader(writer2.toBuffer());
        response2.readUInt8();
        response2.readUInt8();
        expect(response2.readUInt32LE()).toBe(0);
        expect(response2.readUInt16LE()).toBe(chunk2.length);

        const accumulated = (await readWriteStream(context, 6, context.session.handles[6]))!;
        expect(context.session.handles[6].writeSize).toBe(fullBytes.length);
        expect(accumulated).toEqual(fullBytes);
        expect(accumulated.toString("utf16le").replace(/\0+$/, "")).toBe(fullText);
    });

    it("Reads back an empty buffer for a write-stream handle nothing has been written to yet (writeSize defaults to 0).", async () => {
        const context = makeContext();
        context.session.handles[6] = { type: "stream", entityUid: "", generation: "never-written", writeTargetHandleIndex: 3 };

        const result = await readWriteStream(context, 6, context.session.handles[6]);

        expect(result).toEqual(Buffer.alloc(0));
    });
});
