///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { DecodeError } from "../src/codec/BufferCursor.js";
import { dispatchRops, ExecuteBufferTooSmallError } from "../src/RopDispatcher.js";
import { WorkBudgetExceededError } from "../src/rop/ExecuteBudget.js";
import { RopReleaseHandler } from "../src/rop/RopReleaseHandler.js";
import type { RopContext, RopHandler } from "../src/rop/RopHandler.js";

function u32(value: number): Buffer {
    const buffer = Buffer.alloc(4);
    buffer.writeUInt32LE(value);
    return buffer;
}
import { MapiSessionContext } from "../src/MapiSessionManager.js";

describe("RopDispatcher Tests", () => {
    const makeContext = function (): RopContext {
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
    };

    it("Dispatches the spec's own real captured RopRelease example (two ROPs) to the registered handler.", async () => {
        // From [MS-OXCROPS]'s "RopRelease ROP Request" scenario: RopsList = `01 00 00 01 00 01` (RopSize/
        // ServerObjectHandleTable are RopBuffer.ts's concern, not the dispatcher's - see that file's tests).
        const ropsList = Buffer.from("010000010001", "hex");
        const handlers = new Map([[0x01, new RopReleaseHandler()]]);
        const context = makeContext();
        context.session.handles[0] = { type: "logon", entityUid: "mailbox-1" };
        context.session.handles[1] = { type: "folder", entityUid: "folder-1" };

        const response = await dispatchRops(ropsList, handlers, context);

        expect(response.length).toBe(0); // RopRelease produces no response bytes
        expect(context.session.handles[0]).toBeUndefined();
        expect(context.session.handles[1]).toBeUndefined();
    });

    it("Throws a DecodeError (a 400) at an unrecognized RopId, since its byte layout can't be skipped safely.", async () => {
        const ropsList = Buffer.from([0x99, 0xaa, 0xbb, 0x01, 0x00, 0x00]); // unknown RopId 0x99, then a real RopRelease
        const handlers = new Map([[0x01, new RopReleaseHandler()]]);
        const context = makeContext();
        context.session.handles[0] = { type: "logon", entityUid: "mailbox-1" };

        await expect(dispatchRops(ropsList, handlers, context)).rejects.toBeInstanceOf(DecodeError);

        // The trailing RopRelease is never reached - the unknown ROP's bytes couldn't be skipped past.
        expect(context.session.handles[0]).toEqual({ type: "logon", entityUid: "mailbox-1" });
    });

    it("Throws a DecodeError for a truncated ROP.", async () => {
        const handlers = new Map([[0x01, new RopReleaseHandler()]]);
        await expect(dispatchRops(Buffer.from([0x01, 0x00]), handlers, makeContext())).rejects.toBeInstanceOf(DecodeError);
    });

    it("Fails only the ROP whose handler throws, with its handle index and MAPI_E_CALL_FAILED, and keeps running later ROPs.", async () => {
        const handlers = new Map<number, RopHandler>([
            [0x30, { ropId: 0x30, handle: (reader) => { reader.readUInt8(); reader.readUInt8(); throw new Error("database down"); } }],
            [0x31, { ropId: 0x31, responseHandleIndexOffset: 3, failureTailBytes: 2, handle: (reader, writer) => { reader.readBytes(3); writer.writeUInt8(0xee); throw new RangeError("bad stored blob"); } }],
            [0x32, { ropId: 0x32, handle: (reader, writer) => { reader.readUInt8(); reader.readUInt8(); writer.writeUInt8(0x32).writeUInt8(9).writeUInt32LE(0); } }],
            [0x33, { ropId: 0x33, hasNoResponse: true, handle: () => { throw new Error("no response either way"); } }],
            [0x34, { ropId: 0x34, handle: () => { throw new WorkBudgetExceededError("row"); } }],
        ]);

        const response = await dispatchRops(Buffer.from([0x30, 0x00, 0x04, 0x31, 0x00, 0x01, 0x07, 0x33, 0x32, 0x00, 0x09, 0x34]), handlers, makeContext());

        expect(response).toEqual(
            Buffer.concat([
                Buffer.from([0x30, 0x04]), u32(0x80004005),
                Buffer.from([0x31, 0x07]), u32(0x80004005), Buffer.alloc(2), // partial output dropped, tail zeroed
                Buffer.from([0x32, 0x09]), u32(0),
                Buffer.from([0x34, 0x00]), u32(0x80040117), // index byte past the end of the request reads as 0
            ]),
        );
    });

    it("Tells each handler the room left after the RopBufferTooSmall reserve, and answers RopBufferTooSmall, with the unprocessed request, when a response doesn't leave it.", async () => {
        const seen: number[] = [];
        const handlers = new Map<number, RopHandler>([
            [0x40, { ropId: 0x40, handle: (reader, writer, context) => { const size = reader.readUInt8(); seen.push(context.ropOutputRemaining!); writer.writeBytes(Buffer.alloc(size)); } }],
        ]);
        const context = makeContext();

        const response = await dispatchRops(Buffer.from([0x40, 6, 0x40, 8, 0x40, 1]), handlers, context, { maxOutputBytes: 13 });

        // 13 less a RopBufferTooSmall for the 4 bytes after the first ROP (7) leaves 6; then 7 less 5 leaves 2.
        expect(seen).toEqual([6, 2]);
        expect(response).toEqual(Buffer.concat([Buffer.alloc(6), Buffer.from([0xff, 8, 0, 0x40, 8, 0x40, 1])]));
        expect(context.ropOutputRemaining).toBeUndefined();

        // The last ROP needs no reserve after it.
        const last = await dispatchRops(Buffer.from([0x40, 6, 0x40, 2]), handlers, makeContext(), { maxOutputBytes: 13 });
        expect(last).toEqual(Buffer.alloc(8));

        // No room for a RopBufferTooSmall of the whole request: the Execute fails before any ROP runs.
        seen.length = 0;
        await expect(dispatchRops(Buffer.from([0x40, 6, 0x40, 8, 0x40, 1]), handlers, makeContext(), { maxOutputBytes: 8 })).rejects.toBeInstanceOf(ExecuteBufferTooSmallError);
        expect(seen).toEqual([]);
        // Never more than a 16-bit RopSize can describe.
        const huge = new Map<number, RopHandler>([[0x41, { ropId: 0x41, handle: (_reader, writer) => void writer.writeBytes(Buffer.alloc(70000)) }]]);
        expect(await dispatchRops(Buffer.from([0x41]), huge, makeContext(), { maxOutputBytes: 1 << 20 })).toEqual(Buffer.from([0xff, 0xff, 0xff, 0x41]));
    });

    it("Never fails an Execute after a ROP has run: a response that would eat the reserve becomes RopBufferTooSmall from that ROP.", async () => {
        const ran: number[] = [];
        const handlers = new Map<number, RopHandler>([
            [0x42, { ropId: 0x42, handle: (reader, writer) => { const size = reader.readUInt8(); ran.push(size); writer.writeBytes(Buffer.alloc(size)); } }],
        ]);
        // 20 bytes of room; the request is 8 bytes, so 11 must stay free before the first ROP runs.
        const request = Buffer.from([0x42, 4, 0x42, 10, 0x42, 1, 0x42, 1]);
        const response = await dispatchRops(request, handlers, makeContext(), { maxOutputBytes: 20 });
        // ROP 1 (4 bytes) leaves 16 >= 3 + 6. ROP 2 (10 bytes) would leave 6 < 3 + 4, so it is resent from there.
        expect(ran).toEqual([4, 10]);
        expect(response).toEqual(Buffer.concat([Buffer.alloc(4), Buffer.from([0xff, 10, 0]), request.subarray(2)]));
        expect(response.length).toBeLessThanOrEqual(20);
    });

    it("Logs a failing ROP's id, handler and error message, never the error object or the ROP's bytes.", async () => {
        const canary = "CANARY-subject";
        const warn = vi.spyOn(console, "warn").mockImplementation(() => undefined);
        const error = vi.spyOn(console, "error").mockImplementation(() => undefined);
        try {
            const failure: any = new Error("constraint failed");
            failure.parameters = [canary];
            const bare: any = "a bare string"; // a thrown non-Error
            const handlers = new Map<number, RopHandler>([
                [0x0a, { ropId: 0x0a, handle: (reader) => { reader.readBytes(2 + canary.length); throw failure; } }],
                [0x0b, { ropId: 0x0b, handle: (reader) => { reader.readUInt8(); reader.readUInt8(); throw bare; } }],
            ]);

            await dispatchRops(Buffer.concat([Buffer.from([0x0a, 0x00, 0x00]), Buffer.from(canary), Buffer.from([0x0b, 0x00, 0x00])]), handlers, makeContext());

            expect(error).not.toHaveBeenCalled();
            expect(warn.mock.calls).toEqual([
                ["RopDispatcher: ROP 0xa (Object) failed: Error: constraint failed"],
                ["RopDispatcher: ROP 0xb (Object) failed: Error: a bare string"],
            ]);
            expect(JSON.stringify(warn.mock.calls)).not.toContain(canary);
        } finally {
            warn.mockRestore();
            error.mockRestore();
        }
    });

    it("Returns an empty buffer for an empty ropsList.", async () => {
        const response = await dispatchRops(Buffer.alloc(0), new Map(), makeContext());
        expect(response.length).toBe(0);
    });
});
