///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
// Round-5 review fixes below the route: the query/fetch/submit work budget, handle data quotas and cleanup, session
// lock renewal and least-recently-used eviction, submitted drafts, draft value validation and literal lookups.
import { BaseEntity, ModelUtils } from "@rapidrest/service-core";
import { FolderType } from "@rapidmx/restapi";
import { BufferReader, BufferWriter } from "../src/codec/BufferCursor.js";
import { PropertyType, writeTaggedPropertyValue } from "../src/codec/PropertyValue.js";
import {
    MapiSessionContext,
    MapiSessionManager,
    MAX_SESSIONS_PER_USER,
    MemoryMapiSessionStore,
    RedisMapiSessionStore,
    releaseHandle,
    RENEW_LOCK_SCRIPT,
    SESSION_LOCK_TTL_MS,
    takeReleasedHandleData,
} from "../src/MapiSessionManager.js";
import { dispatchRops } from "../src/RopDispatcher.js";
import { resolveRecipientList } from "../src/rop/AddressList.js";
import { LID_APPOINTMENT_END_WHOLE, LID_APPOINTMENT_START_WHOLE, LID_REMINDER_DELTA, PSETID_APPOINTMENT, PSETID_COMMON } from "../src/rop/CalendarNamedProperties.js";
import { resolveContentsKind } from "../src/rop/ContentsTable.js";
import { ExecuteBudget, MAX_QUERIES_PER_EXECUTE, MAX_ROWS_FETCHED_PER_EXECUTE, MAX_SUBMITS_PER_EXECUTE, WorkBudgetExceededError } from "../src/rop/ExecuteBudget.js";
import { loadFastTransferBuffer, openFastTransferHandle } from "../src/rop/FastTransferStream.js";
import {
    DELETE_OWNED_SCRIPT,
    HandleDataCache,
    handleDataKey,
    handleDataOwners,
    MAX_HANDLE_DATA_BYTES_PER_SESSION,
    MAX_HANDLE_DATA_BYTES_PER_USER,
    MemoryHandleDataStore,
    ownerKey,
    RedisHandleDataStore,
    sessionHandleDataIndex,
    writeStreamKey,
    type HandleDataStore,
} from "../src/rop/HandleDataCache.js";
import { loadStreamBody } from "../src/rop/MessageBodyStream.js";
import { assignOrGetNamedPropertyId } from "../src/rop/NamedPropertyRegistry.js";
import { findAllCapped, findPage, findWindow } from "../src/rop/RepoPaging.js";
import { asEntity, boundIndexedValue, literalQueryValue, MAX_INDEXED_VALUE_LENGTH } from "../src/rop/RestapiRules.js";
import { RopDeleteFolderHandler } from "../src/rop/RopDeleteFolderHandler.js";
import { RopFastTransferSourceCopyPropertiesHandler } from "../src/rop/RopFastTransferSourceCopyPropertiesHandler.js";
import { RopGetHierarchyTableHandler } from "../src/rop/RopGetHierarchyTableHandler.js";
import type { RopContext } from "../src/rop/RopHandler.js";
import { RopLogonHandler } from "../src/rop/RopLogonHandler.js";
import { RopSaveChangesMessageHandler } from "../src/rop/RopSaveChangesMessageHandler.js";
import { RopSetPropertiesHandler } from "../src/rop/RopSetPropertiesHandler.js";
import { RopWriteStreamHandler } from "../src/rop/RopWriteStreamHandler.js";
import { FakeRedisClient } from "./fakeRedis.js";
import { InMemoryBlobStore } from "./testDoubles.js";

function makeContext(overrides: Partial<RopContext> = {}): RopContext {
    return {
        mailboxUid: "mailbox-1",
        userUid: "user-1",
        session: new MapiSessionContext({ mailboxUid: "mailbox-1", userUid: "user-1" }),
        folderRepo: { findOne: vi.fn().mockResolvedValue(undefined), find: vi.fn().mockResolvedValue([]) } as any,
        messageRepo: { find: vi.fn().mockResolvedValue([]), findOne: vi.fn().mockResolvedValue(undefined) } as any,
        calendarEventRepo: {} as any,
        mailboxRepo: { findOne: vi.fn().mockResolvedValue({ primarySmtpAddress: "owner@example.com" }) } as any,
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

describe("ExecuteBudget limits", () => {
    it("Counts queries, fetched rows and submits, throwing once each is used up, and refuses unsized work at zero bytes.", () => {
        const budget = new ExecuteBudget(1, 1, { maxQueries: 1, maxFetchedRows: 2, maxSubmits: 1 });
        budget.chargeQueries();
        expect(() => budget.chargeQueries()).toThrow(WorkBudgetExceededError);
        budget.chargeFetchedRows(2);
        expect(() => budget.chargeFetchedRows(1)).toThrow(/fetched row/);
        budget.chargeSubmit();
        expect(() => budget.chargeSubmit()).toThrow(/submit/);
        budget.assertBytesLeft();
        budget.chargeBytes(1);
        expect(() => budget.assertBytesLeft()).toThrow(/byte/);

        const defaults = new ExecuteBudget();
        expect([defaults.queriesRemaining, defaults.fetchedRowsRemaining, defaults.submitsRemaining]).toEqual([
            MAX_QUERIES_PER_EXECUTE,
            MAX_ROWS_FETCHED_PER_EXECUTE,
            MAX_SUBMITS_PER_EXECUTE,
        ]);
    });
});

describe("Budgeted paging", () => {
    it("Charges each query before running it and the rows it returned after, stopping before the next query.", async () => {
        const repo = { find: vi.fn().mockResolvedValue(Array.from({ length: 1000 }, (_, i) => ({ uid: String(i) }))) };
        const budget = new ExecuteBudget(undefined, undefined, { maxQueries: 2 });

        await expect(findAllCapped(repo as any, {}, { uid: "ASC" }, 5000, budget)).rejects.toBeInstanceOf(WorkBudgetExceededError);
        expect(repo.find).toHaveBeenCalledTimes(2);
        expect(budget.fetchedRowsRemaining).toBe(MAX_ROWS_FETCHED_PER_EXECUTE - 2000);

        const fetchLimited = new ExecuteBudget(undefined, undefined, { maxFetchedRows: 10 });
        await expect(findWindow(repo as any, {}, { uid: "ASC" }, 0, 5, fetchLimited)).rejects.toThrow(/fetched row/);
        expect((await findPage(repo as any, {}, { uid: "ASC" }, 0)).length).toBe(1000);
    });

    it("Stops repeated RopGetHierarchyTable ROPs once the Execute's query budget is spent.", async () => {
        const folderRepo = { find: vi.fn().mockResolvedValue([{ uid: "f1", name: "Inbox" }]) };
        const context = makeContext({ folderRepo: folderRepo as any, budget: new ExecuteBudget(undefined, undefined, { maxQueries: 3 }) });
        context.session.handles[1] = { type: "folder", entityUid: "virtual:root" };
        const rop = Buffer.from([0x04, 0, 1, 2, 0]);

        const response = await dispatchRops(Buffer.concat(Array.from({ length: 10 }, () => rop)), new Map([[0x04, new RopGetHierarchyTableHandler()]]), context);

        expect(folderRepo.find).toHaveBeenCalledTimes(3);
        // The fourth and later ROPs fail with MAPI_E_TOO_COMPLEX.
        expect(response.readUInt32LE(6 * 3 + 2)).toBe(0x80040117);
        expect(response.readUInt32LE(6 * 9 + 2)).toBe(0x80040117);
    });

    it("Charges RopDeleteFolder's walk and each delete, and resolveContentsKind's lookup.", async () => {
        const folderRepo = { find: vi.fn().mockResolvedValue([]), delete: vi.fn(), findOne: vi.fn().mockResolvedValue({ type: FolderType.CONTACTS }) };
        const messageRepo = { find: vi.fn().mockResolvedValueOnce([{ uid: "m1", mailboxUid: "mailbox-1", folderUid: "f1" }]).mockResolvedValue([]), delete: vi.fn() };
        const budget = new ExecuteBudget();
        const calendarEventRepo = { find: vi.fn().mockResolvedValue([]), delete: vi.fn() };
        const context = makeContext({ folderRepo: folderRepo as any, messageRepo: messageRepo as any, calendarEventRepo: calendarEventRepo as any, budget });
        context.session.handles[1] = { type: "folder", entityUid: "folder:parent" };
        context.session.folderIds["20"] = "folder:f1";
        const request = new BufferWriter().writeUInt8(0).writeUInt8(1).writeUInt8(0x05).writeBigUInt64LE(20n).toBuffer();

        await new RopDeleteFolderHandler().handle(new BufferReader(request), new BufferWriter(), context);

        expect(messageRepo.delete).toHaveBeenCalledTimes(1);
        expect(folderRepo.delete).toHaveBeenCalledTimes(1);
        // collectSubtree (1 query), the message page (1), the message delete (1), the calendar page (1), the folder delete (1).
        expect(budget.queriesRemaining).toBe(MAX_QUERIES_PER_EXECUTE - 5);

        const spent = makeContext({ folderRepo: folderRepo as any, messageRepo: messageRepo as any, calendarEventRepo: calendarEventRepo as any, budget: new ExecuteBudget(undefined, undefined, { maxQueries: 2 }) });
        spent.session.handles[1] = context.session.handles[1];
        spent.session.folderIds["20"] = "folder:f1";
        messageRepo.find.mockResolvedValueOnce([{ uid: "m2", mailboxUid: "mailbox-1", folderUid: "f1" }]);
        await expect(new RopDeleteFolderHandler().handle(new BufferReader(request), new BufferWriter(), spent)).rejects.toBeInstanceOf(WorkBudgetExceededError);
        expect(messageRepo.delete).toHaveBeenCalledTimes(1); // the second delete never ran

        const kindBudget = new ExecuteBudget(undefined, undefined, { maxQueries: 0 });
        await expect(resolveContentsKind("f1", makeContext({ folderRepo: folderRepo as any, budget: kindBudget }))).rejects.toBeInstanceOf(WorkBudgetExceededError);
    });

    it("Charges RopDeleteFolder's audit write too, once context.audit is configured, undercounting real DB work otherwise.", async () => {
        const folderRepo = { find: vi.fn().mockResolvedValue([]), delete: vi.fn(), findOne: vi.fn().mockResolvedValue({ type: FolderType.CONTACTS }) };
        const messageRepo = { find: vi.fn().mockResolvedValueOnce([{ uid: "m1", mailboxUid: "mailbox-1", folderUid: "f1" }]).mockResolvedValue([]), delete: vi.fn() };
        const calendarEventRepo = { find: vi.fn().mockResolvedValue([]), delete: vi.fn() };
        const audit = vi.fn().mockResolvedValue(undefined);
        const budget = new ExecuteBudget();
        const context = makeContext({ folderRepo: folderRepo as any, messageRepo: messageRepo as any, calendarEventRepo: calendarEventRepo as any, budget, audit });
        context.session.handles[1] = { type: "folder", entityUid: "folder:parent" };
        context.session.folderIds["20"] = "folder:f1";
        const request = new BufferWriter().writeUInt8(0).writeUInt8(1).writeUInt8(0x05).writeBigUInt64LE(20n).toBuffer();

        await new RopDeleteFolderHandler().handle(new BufferReader(request), new BufferWriter(), context);

        expect(audit).toHaveBeenCalledTimes(1);
        // The same 5 charges as the sibling test above, plus one more for the audit write m1's deletion now triggers.
        expect(budget.queriesRemaining).toBe(MAX_QUERIES_PER_EXECUTE - 6);

        // A 3-query budget covers collectSubtree, the message page and the message delete for m2, but not its
        // audit write - proving the audit charge, not just the delete, can be what a tight budget runs out on.
        const spent = makeContext({ folderRepo: folderRepo as any, messageRepo: messageRepo as any, calendarEventRepo: calendarEventRepo as any, audit, budget: new ExecuteBudget(undefined, undefined, { maxQueries: 3 }) });
        spent.session.handles[1] = context.session.handles[1];
        spent.session.folderIds["20"] = "folder:f1";
        messageRepo.find.mockResolvedValueOnce([{ uid: "m2", mailboxUid: "mailbox-1", folderUid: "f1" }]);
        await expect(new RopDeleteFolderHandler().handle(new BufferReader(request), new BufferWriter(), spent)).rejects.toBeInstanceOf(WorkBudgetExceededError);
        expect(messageRepo.delete).toHaveBeenCalledTimes(2); // m1 (above) and m2 (here) both got deleted
        expect(audit).toHaveBeenCalledTimes(1); // m2's audit write never ran - its charge is what threw
    });

    it("Doesn't fetch a message body once the byte budget is gone, and charges the raw message before parsing it.", async () => {
        const blobStore = new InMemoryBlobStore();
        await blobStore.put("bodies/m1", Buffer.from(`Subject: Hi\r\n\r\n${"x".repeat(100)}`));
        const findOne = vi.fn().mockResolvedValue({ uid: "m1", bodyBlobKey: "bodies/m1" });

        const empty = new ExecuteBudget(undefined, 0);
        await expect(loadStreamBody(makeContext({ blobStore, budget: empty, messageRepo: { findOne } as any }), 6, { type: "stream", entityUid: "message:m1", generation: "z" })).rejects.toBeInstanceOf(
            WorkBudgetExceededError,
        );
        expect(findOne).not.toHaveBeenCalled();

        const small = new ExecuteBudget(undefined, 50);
        await expect(loadStreamBody(makeContext({ blobStore, budget: small, messageRepo: { findOne } as any }), 6, { type: "stream", entityUid: "message:m1", generation: "y" })).rejects.toThrow(/byte/);
        expect(small.bytesRemaining).toBeLessThan(0);
    });
});

describe.each([
    ["memory", () => ({ store: new MemoryHandleDataStore(new HandleDataCache()) as HandleDataStore, client: undefined as FakeRedisClient | undefined })],
    [
        "redis",
        () => {
            const client = new FakeRedisClient();
            return { store: new RedisHandleDataStore(client, new HandleDataCache()) as HandleDataStore, client };
        },
    ],
])("HandleDataStore quotas on %s", (_name, build) => {
    const owners = (sessionMax: number, userMax: number) => [
        { index: "session.s1", maxBytes: sessionMax },
        { index: "user.u1", maxBytes: userMax },
    ];

    it("Refuses a value past any owner's quota, counting a replaced key once and a deleted key not at all.", async () => {
        const { store } = build();
        expect(await store.set("a", Buffer.alloc(6), 60, owners(10, 100))).toBe(true);
        expect(await store.set("a", Buffer.alloc(8), 60, owners(10, 100))).toBe(true); // replaces its own 6 bytes
        expect(await store.set("b", Buffer.alloc(4), 60, owners(10, 100))).toBe(false);
        expect(await store.get("b")).toBeUndefined();
        expect(await store.set("b", Buffer.alloc(4), 60, [{ index: "session.s2", maxBytes: 10 }, { index: "user.u1", maxBytes: 11 }])).toBe(false);

        await store.delete("a", owners(10, 100));
        expect(await store.set("b", Buffer.alloc(10), 60, owners(10, 100))).toBe(true);
        expect(await store.set("c", Buffer.alloc(1), 60)).toBe(true); // no owners: no quota
    });

    it("Counts write-stream chunks, a rewritten offset replacing its chunk, and refuses past the quota.", async () => {
        const { store } = build();
        expect(await store.putChunk("w", 0, Buffer.alloc(4), 60, owners(10, 100))).toBe(true);
        expect(await store.putChunk("w", 4, Buffer.alloc(4), 60, owners(10, 100))).toBe(true);
        expect(await store.putChunk("w", 4, Buffer.alloc(6), 60, owners(10, 100))).toBe(true); // 4 + 6
        expect(await store.putChunk("w", 10, Buffer.alloc(1), 60, owners(10, 100))).toBe(false);
        expect(await store.set("v", Buffer.alloc(1), 60, owners(10, 100))).toBe(false);
        expect(await store.readChunks("w", 10)).toEqual(Buffer.alloc(10));
        expect(await store.putChunk("x", 0, Buffer.alloc(1), 60)).toBe(true);

        await store.delete("w", owners(10, 100));
        expect(await store.readChunks("w", 10)).toBeUndefined();
        expect(await store.set("v", Buffer.alloc(10), 60, owners(10, 100))).toBe(true);
    });

    it("Deletes everything an owner index recorded.", async () => {
        const { store } = build();
        await store.set("a", Buffer.from("aa"), 60, owners(100, 100));
        await store.putChunk("w", 0, Buffer.from("ww"), 60, owners(100, 100));
        await store.set("other", Buffer.from("oo"), 60, [{ index: "session.s2", maxBytes: 100 }]);

        await store.deleteOwnedBy("session.s1");
        await store.deleteOwnedBy("session.never-used");

        expect(await store.get("a")).toBeUndefined();
        expect(await store.readChunks("w", 2)).toBeUndefined();
        expect(await store.get("other")).toEqual(Buffer.from("oo"));
    });
});

describe("HandleDataStore quota details", () => {
    it("Stops counting an entry the memory cache evicted or expired.", async () => {
        vi.useFakeTimers();
        try {
            const cache = new HandleDataCache(100, 1000);
            const store = new MemoryHandleDataStore(cache);
            await store.set("a", Buffer.alloc(8), 60, [{ index: "session.s", maxBytes: 10 }]);
            expect(cache.sizeOf("a")).toBe(8);
            vi.advanceTimersByTime(1001);
            expect(cache.sizeOf("a")).toBeUndefined();
            expect(await store.set("b", Buffer.alloc(10), 60, [{ index: "session.s", maxBytes: 10 }])).toBe(true);
        } finally {
            vi.useRealTimers();
        }
    });

    it("Runs the Redis quota and delete scripts with the owner indexes as keys, and keeps a refused value out of the local cache.", async () => {
        const client = new FakeRedisClient();
        const local = new HandleDataCache();
        const store = new RedisHandleDataStore(client, local);
        await store.set("k", Buffer.from("abc"), 600, handleDataOwners("s1", "u1"));
        expect(client.eval).toHaveBeenLastCalledWith(expect.stringContaining("SETEX"), {
            keys: ["mapi.handle.k", ownerKey("session.s1"), ownerKey("user.u1")],
            arguments: ["mapi.handle.", "k", String(24 * 60 * 60), "1000", "YWJj", "600", "3", String(MAX_HANDLE_DATA_BYTES_PER_SESSION), String(MAX_HANDLE_DATA_BYTES_PER_USER)],
        });
        expect(client.ttls.get(ownerKey("user.u1"))).toBe(24 * 60 * 60);

        expect(await store.set("big", Buffer.alloc(4), 60, [{ index: "session.s1", maxBytes: 5 }])).toBe(false);
        expect(local.get("big")).toBeUndefined();

        await store.deleteOwnedBy(sessionHandleDataIndex("s1"));
        expect(client.eval).toHaveBeenLastCalledWith(DELETE_OWNED_SCRIPT, { keys: [ownerKey("session.s1")], arguments: ["mapi.handle."] });
        expect(client.values.has("mapi.handle.k")).toBe(false);
        expect(local.get("k")).toBeUndefined();

        const silent = new FakeRedisClient();
        silent.eval.mockResolvedValueOnce(null);
        await new RedisHandleDataStore(silent, new HandleDataCache()).deleteOwnedBy("session.s1");
    });
});

describe("Session lifecycle", () => {
    function redisManager(client: FakeRedisClient): MapiSessionManager {
        const manager = new MapiSessionManager();
        (manager as any).redisClient = client;
        manager.init();
        return manager;
    }

    function memoryManager(): MapiSessionManager {
        const manager = new MapiSessionManager();
        manager.init();
        return manager;
    }

    describe.each([
        ["memory", memoryManager],
        ["redis", () => redisManager(new FakeRedisClient())],
    ])("on the %s store", (_name, build) => {
        it("Renews a lock only for the token holding it.", async () => {
            const manager = build();
            const token = (await manager.acquireLock("s1"))!;
            expect(await manager.renewLock("s1", "not-it")).toBe(false);
            expect(await manager.renewLock("s1", token)).toBe(true);
            await manager.releaseLock("s1", token);
            expect(await manager.renewLock("s1", token)).toBe(false);
        });

        it("Ends the least recently used session at the cap, not the oldest one.", async () => {
            vi.useFakeTimers();
            try {
                const manager = build();
                const sessions: MapiSessionContext[] = [];
                for (let i = 0; i < MAX_SESSIONS_PER_USER; i++) {
                    vi.advanceTimersByTime(1000);
                    sessions.push(await manager.create("mailbox-1", "user-1"));
                }
                vi.advanceTimersByTime(1000);
                await manager.touch(sessions[0]); // the oldest session is in use
                await manager.touch({ ...sessions[1], uid: "not-in-the-index" });

                vi.advanceTimersByTime(1000);
                await manager.create("mailbox-1", "user-1");

                expect(await manager.load(sessions[0].uid)).toBeDefined();
                expect(await manager.load(sessions[1].uid)).toBeUndefined();
            } finally {
                vi.useRealTimers();
            }
        });

        it("Marks a request in progress, then replaces the marker with its response; a cleared marker lets it run.", async () => {
            const manager = build();
            await manager.markInProgress("s1", "r1");
            expect(await manager.storedResponse("s1", "r1")).toBe("inProgress");
            expect(await manager.storedResponse("s1", "r2")).toBeUndefined();

            await manager.clearInProgress("s1", "r2"); // someone else's id: untouched
            expect(await manager.storedResponse("s1", "r1")).toBe("inProgress");
            await manager.clearInProgress("s1", "r1");
            expect(await manager.storedResponse("s1", "r1")).toBeUndefined();

            await manager.markInProgress("s1", "r3");
            await manager.storeResponse("s1", "r3", Buffer.from([9]));
            await manager.clearInProgress("s1", "r3"); // a stored response is kept
            expect(await manager.storedResponse("s1", "r3")).toEqual(Buffer.from([9]));
        });

        it("Deletes a session's handle data when it is destroyed or evicted, ignoring a store failure.", async () => {
            const manager = build();
            const deleteOwnedBy = vi.spyOn(manager.handleDataStore, "deleteOwnedBy");
            const sessions: MapiSessionContext[] = [];
            for (let i = 0; i < MAX_SESSIONS_PER_USER; i++) {
                sessions.push(await manager.create("mailbox-1", "user-1"));
            }
            await manager.create("mailbox-1", "user-1");
            expect(deleteOwnedBy).toHaveBeenCalledWith(sessionHandleDataIndex(sessions[0].uid));

            deleteOwnedBy.mockRejectedValueOnce(new Error("redis down"));
            await manager.destroy(sessions[1].uid);
            expect(deleteOwnedBy).toHaveBeenLastCalledWith(sessionHandleDataIndex(sessions[1].uid));
            expect(await manager.load(sessions[1].uid)).toBeUndefined();
        });
    });

    it("Renews through its Redis script with the lock TTL.", async () => {
        const client = new FakeRedisClient();
        const manager = redisManager(client);
        const token = (await manager.acquireLock("s9"))!;
        await manager.renewLock("s9", token);
        expect(client.eval).toHaveBeenLastCalledWith(RENEW_LOCK_SCRIPT, { keys: ["mapi.session.{s9}.lock"], arguments: [token, String(SESSION_LOCK_TTL_MS)] });
        expect(await new RedisMapiSessionStore(client).renewLock("mapi.session.{s9}.lock", "other", 5)).toBe(false);
        expect(await new MemoryMapiSessionStore().renewLock("x", "t", 5)).toBe(false);
    });

    it("Remembers the shared-store keys of released stream and FastTransfer handles, once.", () => {
        const session = new MapiSessionContext({ mailboxUid: "m", userUid: "u" });
        session.handles[1] = { type: "fastTransfer", entityUid: "folder:f", generation: "g1" };
        session.handles[2] = { type: "stream", entityUid: "", generation: "g2", writeTargetHandleIndex: 3, writeTargetGeneration: "g3" };
        session.handles[3] = { type: "message", entityUid: "", generation: "g3" };
        session.handles[4] = { type: "folder", entityUid: "folder:f", generation: "g4" };

        releaseHandle(session, 1);
        releaseHandle(session, 3); // takes its write stream (2) with it
        releaseHandle(session, 4);

        expect(takeReleasedHandleData(session)).toEqual([handleDataKey(session.uid, 1, "g1"), writeStreamKey(session.uid, 2, "g2")]);
        expect(takeReleasedHandleData(session)).toEqual([]);
    });
});

describe("FastTransfer quota", () => {
    const refusing = (): HandleDataStore => ({
        get: vi.fn().mockResolvedValue(undefined),
        set: vi.fn().mockResolvedValue(false),
        putChunk: vi.fn(),
        readChunks: vi.fn(),
        delete: vi.fn(),
        deleteOwnedBy: vi.fn(),
    });

    it("Doesn't open a transfer the quota has no room for, releasing the handle it assigned.", async () => {
        const context = makeContext({ handleData: refusing() });
        expect(await openFastTransferHandle(context, 9, { type: "message", entityUid: "message:m1" }, {})).toBe(false);
        expect(context.session.handles[9]).toBeUndefined();
        expect((context.handleData!.set as any).mock.calls[0][3]).toEqual(handleDataOwners(context.session.uid, "user-1"));
    });

    it("Reports a rebuilt stream the quota refuses as tooBig.", async () => {
        const context = makeContext({ handleData: refusing() });
        const transfer = { type: "fastTransfer" as const, entityUid: "message:m1", transferSourceType: "message" as const, transferPosition: 0, generation: "g" };
        expect(await loadFastTransferBuffer(context, 9, transfer)).toBe("tooBig");
    });

    it("Builds one column per property id however often RopFastTransferSourceCopyProperties repeats it.", async () => {
        const context = makeContext({ handleData: new MemoryHandleDataStore(new HandleDataCache()) });
        context.session.handles[1] = { type: "message", entityUid: "message:m1" };
        const request = new BufferWriter().writeUInt8(0).writeUInt8(1).writeUInt8(2).writeUInt8(0).writeUInt8(0).writeUInt8(0).writeUInt16LE(3);
        for (const type of [PropertyType.PtypString, PropertyType.PtypString, PropertyType.PtypInteger32]) {
            request.writeUInt16LE(type).writeUInt16LE(0x0037);
        }

        await new RopFastTransferSourceCopyPropertiesHandler().handle(new BufferReader(request.toBuffer()), new BufferWriter(), context);

        const stream = await context.handleData!.get(handleDataKey(context.session.uid, 2, context.session.handles[2].generation));
        // One PtypString PidTagSubject: tag (4) + empty UTF-16 string terminator (2).
        expect(stream!.length).toBe(6);
    });
});

describe("Submitted drafts", () => {
    it("RopSetProperties and RopWriteStream refuse a draft that was already submitted.", async () => {
        const context = makeContext();
        context.session.handles[5] = { type: "message", entityUid: "", generation: "g5", draftProperties: {}, submitted: true };
        context.session.handles[6] = { type: "stream", entityUid: "", generation: "g6", writeTargetHandleIndex: 5, writeTargetGeneration: "g5", writeSize: 0 };
        const values = new BufferWriter();
        writeTaggedPropertyValue(values, { propertyId: 0x0037, propertyType: PropertyType.PtypString, value: "Again" });
        const setRequest = new BufferWriter().writeUInt8(0).writeUInt8(5).writeUInt16LE(2 + values.length).writeUInt16LE(1).writeBytes(values.toBuffer()).toBuffer();
        const setWriter = new BufferWriter();

        new RopSetPropertiesHandler().handle(new BufferReader(setRequest), setWriter, context);
        const writeWriter = new BufferWriter();
        await new RopWriteStreamHandler().handle(new BufferReader(new BufferWriter().writeUInt8(0).writeUInt8(6).writeUInt16LE(2).writeBytes(Buffer.from("hi")).toBuffer()), writeWriter, context);

        expect(setWriter.toBuffer().readUInt32LE(2)).toBe(0x80070005);
        expect(context.session.handles[5].draftProperties).toEqual({});
        expect(writeWriter.toBuffer().readUInt32LE(2)).toBe(0x80070005);
        expect(context.session.handles[6].writeSize).toBe(0);
    });

    it("RopWriteStream answers MAPI_E_TOO_BIG when the handle data quota refuses a chunk.", async () => {
        const putChunk = vi.fn().mockResolvedValue(false);
        const context = makeContext({ handleData: { putChunk } as any });
        context.session.handles[5] = { type: "message", entityUid: "", generation: "g5" };
        context.session.handles[6] = { type: "stream", entityUid: "", generation: "g6", writeTargetHandleIndex: 5, writeTargetGeneration: "g5", writeSize: 0 };
        const writer = new BufferWriter();

        await new RopWriteStreamHandler().handle(new BufferReader(new BufferWriter().writeUInt8(0).writeUInt8(6).writeUInt16LE(2).writeBytes(Buffer.from("hi")).toBuffer()), writer, context);

        expect(writer.toBuffer().readUInt32LE(2)).toBe(0x80040305);
        expect(putChunk.mock.calls[0][4]).toEqual(handleDataOwners(context.session.uid, "user-1"));
        expect(context.session.handles[6].writeSize).toBe(0);
    });
});

describe("RopSaveChangesMessage validation", () => {
    const saveRequest = () => new BufferReader(Buffer.from([0, 7, 5, 0]));

    it("Ignores a start, end or reminder that doesn't parse instead of storing Invalid Date or NaN.", async () => {
        const create = vi.fn().mockImplementation(async (event: any) => ({ ...event, uid: "evt1" }));
        const context = makeContext({ calendarEventRepo: { create } as any });
        const start = assignOrGetNamedPropertyId(context.session, { guid: PSETID_APPOINTMENT, kind: "lid", lid: LID_APPOINTMENT_START_WHOLE });
        const end = assignOrGetNamedPropertyId(context.session, { guid: PSETID_APPOINTMENT, kind: "lid", lid: LID_APPOINTMENT_END_WHOLE });
        const reminder = assignOrGetNamedPropertyId(context.session, { guid: PSETID_COMMON, kind: "lid", lid: LID_REMINDER_DELTA });
        context.session.handles[5] = {
            type: "message",
            entityUid: "",
            draftFolderUid: "folder:cal",
            draftProperties: { "26": "IPM.Appointment", [String(start)]: "not a date", [String(end)]: "2026-10-01T11:00:00.000Z", [String(reminder)]: "soon" },
        };

        await new RopSaveChangesMessageHandler().handle(saveRequest(), new BufferWriter(), context);

        const saved = create.mock.calls[0][0];
        expect(Number.isNaN(new Date(saved.startDate).getTime())).toBe(false);
        expect(saved.endDate).toEqual(new Date("2026-10-01T11:00:00.000Z"));
        expect(saved.reminderMinutesBeforeStart).toBeUndefined();

        for (const [value, expected] of [["15", 15], ["-5", undefined], ["1.5", undefined], [" ", undefined]] as const) {
            create.mockClear();
            context.session.handles[5] = { type: "message", entityUid: "", draftFolderUid: "folder:cal", draftProperties: { "26": "IPM.Appointment", [String(reminder)]: value } };
            await new RopSaveChangesMessageHandler().handle(saveRequest(), new BufferWriter(), context);
            expect(create.mock.calls[0][0].reminderMinutesBeforeStart).toBe(expected);
        }
    });

    it("Refuses a meeting with more than MAX_RECIPIENTS_PER_MESSAGE attendees with MAPI_E_TOO_BIG, saving nothing.", async () => {
        const create = vi.fn();
        const context = makeContext({ calendarEventRepo: { create } as any });
        const to = Array.from({ length: 501 }, (_, i) => `a${i}@example.com`).join("; ");
        context.session.handles[5] = { type: "message", entityUid: "", draftFolderUid: "folder:cal", draftProperties: { "26": "IPM.Appointment", "3588": to } };
        const writer = new BufferWriter();

        await new RopSaveChangesMessageHandler().handle(saveRequest(), writer, context);

        expect(writer.toBuffer()).toEqual(Buffer.from([0x0c, 7, 0x05, 0x03, 0x04, 0x80]));
        expect(create).not.toHaveBeenCalled();
    });

    it("Updates an existing appointment through an instance of the repo's model class.", async () => {
        class Model {
            public constructor(other: object) {
                Object.assign(this, other);
            }
        }
        const existing = { uid: "evt1", version: 2, title: "Old", startDate: new Date(0), endDate: new Date(0), attendees: [] };
        const update = vi.fn().mockResolvedValue(undefined);
        const context = makeContext({ calendarEventRepo: { findOne: vi.fn().mockResolvedValue(existing), update, modelClass: Model } as any });
        context.session.handles[5] = { type: "message", entityUid: "calendarEvent:evt1", draftProperties: { "26": "IPM.Appointment", "55": "New" } };

        await new RopSaveChangesMessageHandler().handle(saveRequest(), new BufferWriter(), context);

        expect(update.mock.calls[0][1]).toBeInstanceOf(Model);
        expect(update.mock.calls[0][0]).toMatchObject({ uid: "evt1", version: 2, title: "New" });
    });
});

describe("Literal lookups and restapi rules", () => {
    it("Picks the oldest folder of each well-known type at logon.", async () => {
        const find = vi.fn().mockResolvedValue([]);
        const context = makeContext({ folderRepo: { find } as any, mailboxUid: "00000000-0000-0000-0000-00000000000a" });

        await new RopLogonHandler().handle(new BufferReader(new BufferWriter().writeUInt8(0).writeUInt8(0).writeUInt8(1).writeUInt32LE(0).writeUInt32LE(0).writeUInt16LE(0).toBuffer()), new BufferWriter(), context);

        expect(find).toHaveBeenCalledWith(
            { mailboxUid: "00000000-0000-0000-0000-00000000000a", type: FolderType.INBOX, sort: { dateCreated: "ASC", uid: "ASC" }, limit: 1 },
            { ignoreACL: true, limit: 1 },
        );
    });

    it("Looks a display name up literally, keeping only exact matches, and treats a rejected query as unresolved.", async () => {
        const find = vi.fn().mockResolvedValueOnce([{ displayName: "Other", emails: [{ address: "o@example.com" }] }]).mockRejectedValueOnce(new Error("bad operand"));

        const first = await resolveRecipientList("ne(x)", { mailboxUid: "m", contactRepo: { find } as any });
        const second = await resolveRecipientList("me", { mailboxUid: "m", contactRepo: { find } as any });

        expect(find.mock.calls[0][0].displayName).toEqual(ModelUtils.literal("ne(x)"));
        expect(first).toEqual({ recipients: [], unresolved: ["ne(x)"], invalid: [] });
        expect(second.unresolved).toEqual(["me"]);
    });

    it("Copies restapi's boundIndexedValue and asEntity.", () => {
        expect(boundIndexedValue("short")).toBe("short");
        expect(boundIndexedValue(undefined)).toBeUndefined();
        const long = "x".repeat(MAX_INDEXED_VALUE_LENGTH + 1);
        expect(boundIndexedValue(long)).toMatch(/^sha256:[0-9a-f]{64}$/);
        expect(boundIndexedValue(boundIndexedValue(long))).toBe(boundIndexedValue(long));
        expect(literalQueryValue("a(b)")).toEqual(ModelUtils.literal("a(b)"));

        class Entity extends BaseEntity {}
        const instance = new Entity({ uid: "u" });
        expect(asEntity({ modelClass: Entity } as any, instance)).toBe(instance);
        const plain = { uid: "p" };
        expect(asEntity({} as any, plain)).toBe(plain);
        expect(asEntity({ modelClass: Entity } as any, plain)).toBeInstanceOf(Entity);
    });
});
