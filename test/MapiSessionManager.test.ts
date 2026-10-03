///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import {
    ACQUIRE_LOCK_SCRIPT,
    COMPARE_AND_SET_SCRIPT,
    MapiSessionContext,
    MapiSessionManager,
    MAX_SESSION_BYTES,
    MAX_SESSION_LIFETIME_MS,
    MAX_SESSIONS_PER_USER,
    MemoryMapiSessionStore,
    RedisMapiSessionStore,
    SESSION_LOCK_TTL_MS,
    SESSION_TTL_SECONDS,
    sessionKey,
} from "../src/MapiSessionManager.js";
import { MemoryHandleDataStore, RedisHandleDataStore } from "../src/rop/HandleDataCache.js";
import { FakeRedisClient } from "./fakeRedis.js";
import { buildSessionManager } from "./managerFactory.js";

async function memoryManager(): Promise<MapiSessionManager> {
    return buildSessionManager();
}

async function redisManager(client: FakeRedisClient): Promise<MapiSessionManager> {
    return buildSessionManager(client);
}

describe("MapiSessionManager Tests", () => {
    afterEach(() => {
        vi.useRealTimers();
    });

    describe.each([
        ["memory", async () => ({ manager: await memoryManager() })],
        ["redis", async () => ({ manager: await redisManager(new FakeRedisClient()) })],
    ])("on the %s store", (_name, build) => {
        it("Round-trips a session, handing out an independent copy on every load.", async () => {
            const { manager } = await build();
            const created = await manager.create("mailbox-1", "user-1");

            const first = (await manager.load(created.uid))!;
            const second = (await manager.load(created.uid))!;
            first.handles[1] = { type: "logon", entityUid: "mailbox-1" };

            expect(first.mailboxUid).toBe("mailbox-1");
            expect(first.userUid).toBe("user-1");
            expect(second.handles[1]).toBeUndefined();
            expect(await manager.load("unknown")).toBeUndefined();
        });

        it("Saves compare-and-set on version: the second of two overlapping saves conflicts and keeps nothing.", async () => {
            const { manager } = await build();
            const created = await manager.create("mailbox-1", "user-1");
            const requestA = (await manager.load(created.uid))!;
            const requestB = (await manager.load(created.uid))!;
            requestA.handles[1] = { type: "logon", entityUid: "from-a" };
            requestB.handles[1] = { type: "logon", entityUid: "from-b" };

            expect(await manager.save(requestA)).toBe("saved");
            expect(requestA.version).toBe(1);
            expect(await manager.save(requestB)).toBe("conflict");
            expect(requestB.version).toBe(0); // restored, so the object still describes what it was loaded as

            const stored = (await manager.load(created.uid))!;
            expect(stored.handles[1].entityUid).toBe("from-a");
            expect(stored.version).toBe(1);
            expect(await manager.save(stored)).toBe("saved");
        });

        it("Compares the separate version entry, not the version inside the session JSON.", async () => {
            const { manager } = await build();
            const created = await manager.create("mailbox-1", "user-1");
            const loaded = (await manager.load(created.uid))!;
            // A client can't forge its way past a conflict by carrying a different version in the JSON it saves.
            loaded.version = 5;
            expect(await manager.save(loaded)).toBe("conflict");
            loaded.version = 0;
            expect(await manager.save(loaded)).toBe("saved");
        });

        it("Reports a save of a destroyed session as missing instead of recreating it.", async () => {
            const { manager } = await build();
            const created = await manager.create("mailbox-1", "user-1");
            const loaded = (await manager.load(created.uid))!;
            await manager.destroy(created.uid);

            expect(await manager.save(loaded)).toBe("missing");
            expect(await manager.load(created.uid)).toBeUndefined();
        });

        it("Refuses to save a session larger than MAX_SESSION_BYTES, keeping the stored one.", async () => {
            const { manager } = await build();
            const created = await manager.create("mailbox-1", "user-1");
            const loaded = (await manager.load(created.uid))!;
            loaded.handles[1] = { type: "table", entityUid: "folder:x", rows: ["x".repeat(MAX_SESSION_BYTES)] };

            expect(await manager.save(loaded)).toBe("tooBig");
            expect(loaded.version).toBe(0);
            expect((await manager.load(created.uid))!.handles[1]).toBeUndefined();
        });

        it("Ends a session older than MAX_SESSION_LIFETIME_MS on load, however recently it was used.", async () => {
            const { manager } = await build();
            const created = await manager.create("mailbox-1", "user-1");
            const loaded = (await manager.load(created.uid))!;
            loaded.createdAt = new Date(Date.now() - MAX_SESSION_LIFETIME_MS - 1000).toISOString();
            expect(await manager.save(loaded)).toBe("saved");

            expect(await manager.load(created.uid)).toBeUndefined();
            expect(await manager.save(loaded)).toBe("missing");
        });

        it("Holds at most MAX_SESSIONS_PER_USER sessions per user, ending the oldest first, without touching other users.", async () => {
            const { manager } = await build();
            const other = await manager.create("mailbox-2", "user-2");
            const sessions = [];
            for (let i = 0; i < MAX_SESSIONS_PER_USER; i++) {
                sessions.push(await manager.create("mailbox-1", "user-1"));
            }
            // A session the user already ended doesn't count against the limit.
            await manager.destroy(sessions[5].uid);
            const replacement = await manager.create("mailbox-1", "user-1");
            expect(await manager.load(sessions[0].uid)).toBeDefined();

            const overLimit = await manager.create("mailbox-1", "user-1");

            expect(await manager.load(sessions[0].uid)).toBeUndefined();
            expect(await manager.load(sessions[1].uid)).toBeDefined();
            expect(await manager.load(replacement.uid)).toBeDefined();
            expect(await manager.load(overLimit.uid)).toBeDefined();
            expect(await manager.load(other.uid)).toBeDefined();
        });

        it("Keeps the per-user cap when many Connects run concurrently.", async () => {
            const { manager } = await build();
            const created = await Promise.all(Array.from({ length: MAX_SESSIONS_PER_USER * 2 }, () => manager.create("mailbox-1", "user-1")));

            const live = (await Promise.all(created.map((session) => manager.load(session.uid)))).filter(Boolean);
            expect(live).toHaveLength(MAX_SESSIONS_PER_USER);
        });

        it("Lets one request at a time hold a session's lock, released only by its own token.", async () => {
            const { manager } = await build();
            const token = await manager.acquireLock("session-1");
            expect(token).toEqual(expect.any(String));
            expect(await manager.acquireLock("session-1")).toBeUndefined();
            expect(await manager.acquireLock("session-2")).toBeDefined();

            await manager.releaseLock("session-1", "someone-else");
            expect(await manager.acquireLock("session-1")).toBeUndefined();
            await manager.releaseLock("session-1", token);
            expect(await manager.acquireLock("session-1")).toBeDefined();
        });

        it("Remembers the last response for its X-RequestId only.", async () => {
            const { manager } = await build();
            expect(await manager.storedResponse("session-1", "req-1")).toBeUndefined();

            await manager.storeResponse("session-1", "req-1", Buffer.from([1, 2, 3]));
            expect(await manager.storedResponse("session-1", "req-1")).toEqual(Buffer.from([1, 2, 3]));
            expect(await manager.storedResponse("session-1", "req-2")).toBeUndefined();
            expect(await manager.storedResponse("session-2", "req-1")).toBeUndefined();

            await manager.storeResponse("session-1", "req-2", Buffer.from([4]));
            expect(await manager.storedResponse("session-1", "req-1")).toBeUndefined();
        });
    });

    it("Uses shared handle data storage on Redis, and process memory otherwise.", async () => {
        expect((await memoryManager()).handleDataStore).toBeInstanceOf(MemoryHandleDataStore);
        expect((await redisManager(new FakeRedisClient())).handleDataStore).toBeInstanceOf(RedisHandleDataStore);
    });

    it("Redis store: never serves a stale per-process copy - a change saved through another replica is seen on the next load.", async () => {
        const client = new FakeRedisClient();
        const podA = await redisManager(client);
        const podB = await redisManager(client);
        const created = await podA.create("mailbox-1", "user-1");
        await podA.load(created.uid); // would have primed a local cache before

        const onB = (await podB.load(created.uid))!;
        onB.handles[3] = { type: "folder", entityUid: "folder:x" };
        expect(await podB.save(onB)).toBe("saved");

        expect((await podA.load(created.uid))!.handles[3]).toEqual({ type: "folder", entityUid: "folder:x" });
    });

    it("Redis store: saves through the compare-and-set script against the version key, with the session TTL.", async () => {
        const client = new FakeRedisClient();
        const manager = await redisManager(client);
        const created = await manager.create("mailbox-1", "user-1");
        const loaded = (await manager.load(created.uid))!;

        await manager.save(loaded);

        const key = `mapi.session.{${created.uid}}`;
        expect(sessionKey(created.uid)).toBe(key);
        expect(client.eval).toHaveBeenCalledWith(COMPARE_AND_SET_SCRIPT, {
            keys: [key, `${key}.version`],
            arguments: ["0", expect.any(String), String(SESSION_TTL_SECONDS)],
        });
        expect(COMPARE_AND_SET_SCRIPT).not.toContain("cjson");
        expect(client.ttls.get(key)).toBe(SESSION_TTL_SECONDS);
        expect(client.values.get(`${key}.version`)).toBe("1");
    });

    it("Redis store: takes the lock with SET NX and a TTL.", async () => {
        const client = new FakeRedisClient();
        await (await redisManager(client)).acquireLock("s1");
        expect(client.eval).toHaveBeenCalledWith(ACQUIRE_LOCK_SCRIPT, { keys: ["mapi.session.{s1}.lock"], arguments: [expect.any(String), String(SESSION_LOCK_TTL_MS)] });
    });

    it("RedisMapiSessionStore maps a missing key to undefined.", async () => {
        const store = new RedisMapiSessionStore(new FakeRedisClient());
        expect(await store.get("nope")).toBeUndefined();
        expect(await store.getValue("nope")).toBeUndefined();
    });

    it("MemoryMapiSessionStore expires entries after their TTL, for reads, saves, locks and later writes.", async () => {
        vi.useFakeTimers();
        const store = new MemoryMapiSessionStore();
        await store.create("a", JSON.stringify({ version: 0 }), 1);
        await store.create("b", JSON.stringify({ version: 0 }), 1);
        expect(await store.acquireLock("lock", "t", 500)).toBe(true);
        expect(await store.get("a")).toBeDefined();

        vi.advanceTimersByTime(1001);

        expect(await store.get("a")).toBeUndefined();
        expect(await store.compareAndSet("b", 0, "{}", 1)).toBe("missing");
        expect(await store.acquireLock("lock", "t2", 500)).toBe(true);
        await store.create("c", "{}", 10); // sweeps the expired "b"
        expect((store as any).entries.has("b")).toBe(false);
        expect(await store.get("c")).toBe("{}");
    });

    it("MapiSessionContext starts MIDs and handles at 1.", () => {
        const session = new MapiSessionContext({ mailboxUid: "m", userUid: "u" });
        expect(session.firstMessageId).toBe(1);
        expect(session.nextMessageId).toBe(1);
    });
});
