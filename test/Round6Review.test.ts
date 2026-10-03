///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
// Round-6 review fixes below the route: O(1) write-stream and owner quota accounting, the RopBufferTooSmall reserve,
// meeting invites only for submitted revisions, and in-progress markers that never overwrite a stored response.
import { ApiErrors } from "@rapidrest/service-core";
import { BufferReader, BufferWriter } from "../src/codec/BufferCursor.js";
import { MapiSessionContext, MapiSessionManager } from "../src/MapiSessionManager.js";
import { dispatchRops, ExecuteBufferTooSmallError } from "../src/RopDispatcher.js";
import {
    DELETE_OWNED_SCRIPT,
    DELETE_WITH_OWNERS_SCRIPT,
    HandleDataCache,
    handleDataCache,
    handleDataKey,
    MAX_CHUNKS_PER_STREAM,
    MemoryHandleDataStore,
    OWNER_RECOUNT_INTERVAL_MS,
    ownerKey,
    PUT_CHUNK_WITH_QUOTA_SCRIPT,
    RedisHandleDataStore,
    SET_WITH_QUOTA_SCRIPT,
    type HandleDataStore,
} from "../src/rop/HandleDataCache.js";
import { isInvitePending } from "../src/rop/RestapiRules.js";
import type { RopContext, RopHandler } from "../src/rop/RopHandler.js";
import { RopReadStreamHandler } from "../src/rop/RopReadStreamHandler.js";
import { RopSaveChangesMessageHandler } from "../src/rop/RopSaveChangesMessageHandler.js";
import { MAX_INVITE_REQUEST_ATTEMPTS, RopSubmitMessageHandler } from "../src/rop/RopSubmitMessageHandler.js";
import { FakeRedisClient } from "./fakeRedis.js";
import { buildSessionManager } from "./managerFactory.js";

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
        calendarEventClass: class {
            public constructor(data: object) {
                Object.assign(this, data);
            }
        },
        scanPipeline: {} as any,
        mailTransport: {} as any,
        blobStore: {} as any,
        ...overrides,
    };
}

const owners = (max: number) => [{ index: "session.s1", maxBytes: max }];

describe.each([
    ["memory", () => ({ store: new MemoryHandleDataStore(new HandleDataCache()) as HandleDataStore, client: undefined as FakeRedisClient | undefined })],
    [
        "redis",
        () => {
            const client = new FakeRedisClient();
            return { store: new RedisHandleDataStore(client, new HandleDataCache()) as HandleDataStore, client };
        },
    ],
])("Write-stream accounting on %s", (_name, build) => {
    afterEach(() => {
        vi.useRealTimers();
    });

    it("Charges a stream its size, not a sum over its chunks, so a rewritten offset never counts twice.", async () => {
        const { store } = build();
        expect(await store.putChunk("w", 0, Buffer.alloc(6), 60, owners(10))).toBe(true);
        expect(await store.putChunk("w", 0, Buffer.alloc(6), 60, owners(10))).toBe(true); // a retried write
        expect(await store.putChunk("w", 6, Buffer.alloc(4), 60, owners(10))).toBe(true);
        expect(await store.putChunk("w", 10, Buffer.alloc(1), 60, owners(10))).toBe(false);
        expect(await store.readChunks("w", 10)).toEqual(Buffer.alloc(10));
    });

    it("Caps the chunks of one stream, still accepting a rewrite of a chunk it has.", async () => {
        const { store } = build();
        for (let offset = 0; offset < MAX_CHUNKS_PER_STREAM; offset++) {
            expect(await store.putChunk("w", offset, Buffer.from([offset & 0xff]), 60, owners(1 << 20))).toBe(true);
        }
        expect(await store.putChunk("w", MAX_CHUNKS_PER_STREAM, Buffer.from([1]), 60, owners(1 << 20))).toBe(false);
        expect(await store.putChunk("w", 5, Buffer.from([5]), 60, owners(1 << 20))).toBe(true);
        expect((await store.readChunks("w", MAX_CHUNKS_PER_STREAM))!.length).toBe(MAX_CHUNKS_PER_STREAM);
    });

    it("Recounts an index that looks full at most once per interval, so data gone without its owners frees up after it.", async () => {
        vi.useFakeTimers();
        const { store, client } = build();
        expect(await store.set("a", Buffer.alloc(8), 60, owners(10))).toBe(true);
        expect(await store.set("b", Buffer.alloc(4), 60, owners(10))).toBe(false); // recounted: "a" is still there
        await store.delete("a"); // no owners given: the total still counts it
        expect(await store.set("b", Buffer.alloc(4), 60, owners(10))).toBe(false);
        expect(await store.set("b", Buffer.alloc(4), 60, owners(10))).toBe(false);
        if (client) {
            // Once for the index's first write (no total yet), once for the first refusal.
            expect(client.recounts.get(ownerKey("session.s1"))).toBe(2);
        }
        vi.advanceTimersByTime(OWNER_RECOUNT_INTERVAL_MS);
        expect(await store.set("b", Buffer.alloc(4), 60, owners(10))).toBe(true);
    });

    it("Takes a deleted key off its owners' totals right away, and ignores a key an owner never recorded.", async () => {
        const { store } = build();
        expect(await store.putChunk("w", 0, Buffer.alloc(10), 60, owners(10))).toBe(true);
        expect(await store.set("x", Buffer.alloc(1), 60, owners(10))).toBe(false); // recounted now, so not again for a while
        await store.delete("w", owners(10));
        await store.delete("never-stored", owners(10));
        await store.delete("never-stored", [{ index: "session.unknown", maxBytes: 1 }]);
        expect(await store.putChunk("w2", 0, Buffer.alloc(10), 60, owners(10))).toBe(true);
    });
});

describe("Owner index scripts", () => {
    it("Keep no per-chunk or per-member walk on the write path.", () => {
        for (const script of [SET_WITH_QUOTA_SCRIPT, PUT_CHUNK_WITH_QUOTA_SCRIPT]) {
            expect(script).not.toContain("HKEYS");
            expect(script).not.toContain("HSTRLEN");
            // The only member walk is the rate-limited recount.
            expect(script.match(/ZRANGE/g)).toHaveLength(1);
            expect(script).toContain("'PX', recountMs, 'NX'");
        }
        expect(PUT_CHUNK_WITH_QUOTA_SCRIPT).toContain("HLEN");
    });

    it("Store a stream's size in its hash and a running total per index, and delete both with the index.", async () => {
        const client = new FakeRedisClient();
        const store = new RedisHandleDataStore(client, new HandleDataCache());
        await store.putChunk("w", 0, Buffer.from("abc"), 60, owners(100));
        await store.putChunk("w", 3, Buffer.from("de"), 60, owners(100));
        expect(client.hashes.get("mapi.handle.w")!.get("size")).toBe("5");
        expect(client.values.get(`${ownerKey("session.s1")}.total`)).toBe("5");
        expect(client.eval).toHaveBeenLastCalledWith(PUT_CHUNK_WITH_QUOTA_SCRIPT, {
            keys: ["mapi.handle.w", ownerKey("session.s1")],
            arguments: ["mapi.handle.", "w", String(24 * 60 * 60), String(OWNER_RECOUNT_INTERVAL_MS), "ZGU", "60", "2", "3", String(MAX_CHUNKS_PER_STREAM), "100"],
        });

        await store.delete("w", owners(100));
        expect(client.eval).toHaveBeenLastCalledWith(DELETE_WITH_OWNERS_SCRIPT, { keys: ["mapi.handle.w", ownerKey("session.s1")], arguments: ["w"] });
        expect(client.values.get(`${ownerKey("session.s1")}.total`)).toBe("0");

        await store.set("v", Buffer.from("x"), 60, owners(100));
        await store.deleteOwnedBy("session.s1");
        expect(client.eval).toHaveBeenLastCalledWith(DELETE_OWNED_SCRIPT, { keys: [ownerKey("session.s1")], arguments: ["mapi.handle."] });
        expect(client.values.has(`${ownerKey("session.s1")}.total`)).toBe(false);
    });

    it("Recount an index that has no running total yet (written before totals existed) instead of trusting zero.", async () => {
        const client = new FakeRedisClient();
        const store = new RedisHandleDataStore(client, new HandleDataCache());
        client.values.set("mapi.handle.old", "AAAA");
        client.sortedSets.set(ownerKey("session.s1"), new Map([["old", 8], ["gone", 50]]));
        expect(await store.set("new", Buffer.alloc(4), 60, owners(10))).toBe(false);
        expect(client.recounts.get(ownerKey("session.s1"))).toBe(1);
        expect(client.sortedSets.get(ownerKey("session.s1"))!.has("gone")).toBe(false);
        expect(await store.set("new", Buffer.alloc(2), 60, owners(10))).toBe(true);
        expect(client.values.get(`${ownerKey("session.s1")}.total`)).toBe("10");
    });

    it("Stop counting a memory stream whose first chunk was evicted.", async () => {
        const cache = new HandleDataCache(12, 60_000);
        const store = new MemoryHandleDataStore(cache);
        expect(await store.putChunk("w", 0, Buffer.alloc(6), 60, owners(12))).toBe(true);
        expect(await store.putChunk("w", 6, Buffer.alloc(6), 60, owners(12))).toBe(true);
        cache.set("other", Buffer.alloc(6)); // evicts the stream's first chunk
        expect(await store.set("v", Buffer.alloc(12), 60, owners(12))).toBe(true);
        expect(await store.readChunks("w", 12)).toBeUndefined();
    });
});

describe("RopBufferTooSmall reserve", () => {
    it("Turns a RopReadStream with no room into RopBufferTooSmall from that ROP, leaving the earlier ROP's response and the stream position.", async () => {
        const context = makeContext();
        context.session.handles[6] = { type: "stream", entityUid: "", streamPosition: 0, generation: "reserve" };
        handleDataCache.set(handleDataKey(context.session.uid, 6, "reserve"), Buffer.alloc(100, 1));
        const read = (count: number) => new BufferWriter().writeUInt8(0x2c).writeUInt8(0).writeUInt8(6).writeUInt16LE(count).toBuffer();
        const handlers = new Map<number, RopHandler>([[0x2c, new RopReadStreamHandler()]]);
        try {
            // 20 bytes of data and its 8-byte header, then exactly a RopBufferTooSmall (3 + 5) for the second read.
            const response = await dispatchRops(Buffer.concat([read(20), read(10)]), handlers, context, { maxOutputBytes: 28 + 8 });
            expect(response.length).toBe(36);
            expect(response.readUInt16LE(6)).toBe(20);
            expect(response[28]).toBe(0xff);
            expect(response.readUInt16LE(29)).toBe(18);
            expect(response.subarray(31)).toEqual(read(10));
            expect(context.session.handles[6].streamPosition).toBe(20);

            // A first read asking for more than fits returns what leaves the reserve, never failing the request.
            context.session.handles[6].streamPosition = 0;
            const partial = await dispatchRops(Buffer.concat([read(100), read(10)]), handlers, context, { maxOutputBytes: 40 });
            expect(partial.readUInt16LE(6)).toBe(40 - 8 - 8);
            expect(partial[32]).toBe(0xff);
            expect(context.session.handles[6].streamPosition).toBe(24);

            // Only a request that can't be resent whole fails, before anything is read.
            context.session.handles[6].streamPosition = 0;
            await expect(dispatchRops(Buffer.concat([read(1), read(1)]), handlers, context, { maxOutputBytes: 12 })).rejects.toBeInstanceOf(ExecuteBufferTooSmallError);
            expect(context.session.handles[6].streamPosition).toBe(0);
        } finally {
            handleDataCache.delete(handleDataKey(context.session.uid, 6, "reserve"));
        }
    });
});

describe("Meeting invites for submitted revisions only", () => {
    const saveRequest = () => new BufferReader(new BufferWriter().writeUInt8(0).writeUInt8(7).writeUInt8(5).writeUInt8(0).toBuffer());
    const submitRequest = () => new BufferReader(new BufferWriter().writeUInt8(0).writeUInt8(5).writeUInt8(0).toBuffer());
    const returnValue = (writer: BufferWriter) => writer.toBuffer().readUInt32LE(2);

    it("Matches the job's eligibility rule.", () => {
        expect(isInvitePending({ sequence: 2, inviteSequenceSent: 2 })).toBe(false);
        expect(isInvitePending({ sequence: 2, inviteSequenceSent: 1 })).toBe(true);
        expect(isInvitePending({ sequence: 0, inviteSequenceSent: null })).toBe(true);
        expect(isInvitePending({ sequence: 0 })).toBe(true);
    });

    it("Creates a saved meeting as already handled.", async () => {
        const create = vi.fn().mockResolvedValue({ uid: "evt1" });
        const context = makeContext({ calendarEventRepo: { create } as any });
        context.session.handles[5] = { type: "message", entityUid: "", draftProperties: { "26": "IPM.Appointment", "3588": "a@example.com" } };
        await new RopSaveChangesMessageHandler().handle(saveRequest(), new BufferWriter(), context);
        expect(create.mock.calls[0][0]).toMatchObject({ sequence: 0, inviteSequenceSent: 0 });
    });

    it("Stamps an edited revision as handled, unless a submitted revision is still waiting for the job.", async () => {
        const save = async (existing: Record<string, unknown>) => {
            const update = vi.fn().mockResolvedValue(undefined);
            const context = makeContext({ calendarEventRepo: { findOne: vi.fn().mockResolvedValue(existing), update } as any });
            context.session.handles[5] = { type: "message", entityUid: "calendarEvent:evt1", draftProperties: { "26": "IPM.Appointment", "3588": "a@example.com; b@example.com" } };
            await new RopSaveChangesMessageHandler().handle(saveRequest(), new BufferWriter(), context);
            return update.mock.calls[0][0];
        };
        const base = { uid: "evt1", version: 2, startDate: new Date(0), endDate: new Date(0), attendees: [{ address: "a@example.com" }] };

        // Invited (or saved) at sequence 4; adding an attendee bumps to 5 without sending.
        expect(await save({ ...base, sequence: 4, inviteSequenceSent: 4 })).toMatchObject({ sequence: 5, inviteSequenceSent: 5 });
        // Submitted at 4 and not yet sent: the job still sends, now the edited meeting.
        expect(await save({ ...base, sequence: 4, inviteSequenceSent: null })).toMatchObject({ sequence: 5, inviteSequenceSent: null });
        expect(await save({ ...base, sequence: 4, inviteSequenceSent: 3 })).toMatchObject({ sequence: 5, inviteSequenceSent: 3 });
        expect(await save({ ...base, sequence: 4 })).toMatchObject({ sequence: 5, inviteSequenceSent: null });
    });

    describe("RopSubmitMessage on an appointment", () => {
        const event = (overrides: Record<string, unknown> = {}) => ({
            uid: "evt1",
            version: 3,
            mailboxUid: "mailbox-1",
            organizer: { address: "owner@example.com" },
            sequence: 2,
            inviteSequenceSent: 2,
            attendees: [{ address: "a@example.com" }],
            ...overrides,
        });
        const submit = async (repo: Record<string, unknown>) => {
            const context = makeContext({ calendarEventRepo: repo as any });
            context.session.handles[5] = { type: "message", entityUid: "calendarEvent:evt1", draftProperties: { "26": "IPM.Appointment" } };
            const writer = new BufferWriter();
            await new RopSubmitMessageHandler().handle(submitRequest(), writer, context);
            return writer;
        };
        const conflict = () => Object.assign(new Error("version"), { code: ApiErrors.INVALID_OBJECT_VERSION });

        it("Clears the handled stamp with a versioned update so the job sends this revision.", async () => {
            const row = event();
            const update = vi.fn().mockResolvedValue(undefined);
            const writer = await submit({ findOne: vi.fn().mockResolvedValue(row), update });
            expect(returnValue(writer)).toBe(0);
            expect(update).toHaveBeenCalledTimes(1);
            expect(update.mock.calls[0][0]).toEqual({ uid: "evt1", version: 3, inviteSequenceSent: null });
            expect(update.mock.calls[0][1]).toMatchObject({ uid: "evt1", version: 3 });
            expect(update.mock.calls[0][2]).toEqual({ ignoreACL: true });
        });

        it("Changes nothing for a revision already waiting for the job, or for an attendee's copy.", async () => {
            const update = vi.fn();
            expect(returnValue(await submit({ findOne: vi.fn().mockResolvedValue(event({ inviteSequenceSent: null })), update }))).toBe(0);
            expect(returnValue(await submit({ findOne: vi.fn().mockResolvedValue(event({ organizer: { address: "someone@example.com" } })), update }))).toBe(0x80070005);
            expect(update).not.toHaveBeenCalled();
        });

        it("Re-reads a row another edit changed meanwhile, and gives up after a few conflicts.", async () => {
            // Conflict, then the re-read row (an attendee's reply landed) is still handled: update again.
            const findOne = vi.fn().mockResolvedValueOnce(event()).mockResolvedValueOnce(event({ version: 4 }));
            const update = vi.fn().mockRejectedValueOnce(conflict()).mockResolvedValueOnce(undefined);
            expect(returnValue(await submit({ findOne, update }))).toBe(0);
            expect(update.mock.calls[1][0]).toEqual({ uid: "evt1", version: 4, inviteSequenceSent: null });

            // Conflict, then the row is already waiting: done without another update.
            const waiting = vi.fn().mockResolvedValueOnce(event()).mockResolvedValueOnce(event({ inviteSequenceSent: null, version: 4 }));
            const once = vi.fn().mockRejectedValueOnce(conflict());
            expect(returnValue(await submit({ findOne: waiting, update: once }))).toBe(0);
            expect(once).toHaveBeenCalledTimes(1);

            // Conflict, then the row is gone.
            const gone = vi.fn().mockResolvedValueOnce(event()).mockResolvedValueOnce(undefined);
            expect(returnValue(await submit({ findOne: gone, update: vi.fn().mockRejectedValue(conflict()) }))).toBe(0x8004010f);

            // Always conflicting, or any other failure: the ROP fails.
            const always = vi.fn().mockRejectedValue(conflict());
            await expect(submit({ findOne: vi.fn().mockResolvedValue(event()), update: always })).rejects.toThrow(/kept changing/);
            expect(always).toHaveBeenCalledTimes(MAX_INVITE_REQUEST_ATTEMPTS);
            await expect(submit({ findOne: vi.fn().mockResolvedValue(event()), update: vi.fn().mockRejectedValue(new Error("db down")) })).rejects.toThrow("db down");
        });
    });
});

describe("In-progress markers", () => {
    async function redisManager(): Promise<MapiSessionManager> {
        return buildSessionManager(new FakeRedisClient());
    }

    async function memoryManager(): Promise<MapiSessionManager> {
        return buildSessionManager();
    }

    it.each([
        ["memory", memoryManager],
        ["redis", redisManager],
    ])("Renews only a still-present marker for the same request, never over a stored response (%s).", async (_name, build) => {
        const manager = await build();
        expect(await manager.renewInProgress("s1", "r1")).toBe(false); // nothing there: nothing written
        expect(await manager.storedResponse("s1", "r1")).toBeUndefined();

        await manager.markInProgress("s1", "r1");
        expect(await manager.renewInProgress("s1", "r2")).toBe(false);
        expect(await manager.renewInProgress("s1", "r1")).toBe(true);

        await manager.storeResponse("s1", "r1", Buffer.from([7]));
        expect(await manager.renewInProgress("s1", "r1")).toBe(false); // a late renewal
        await manager.clearInProgress("s1", "r1");
        expect(await manager.storedResponse("s1", "r1")).toEqual(Buffer.from([7]));
    });
});
