///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import * as crypto from "crypto";
import { ObjectDecorators } from "@rapidrest/core";
import { DatabaseDecorators, ObjectFactory, SimpleEntity } from "@rapidrest/service-core";
import {
    handleDataCache,
    handleDataKey,
    MemoryHandleDataStore,
    RedisHandleDataStore,
    sessionHandleDataIndex,
    writeStreamKey,
    type HandleDataStore,
} from "./rop/HandleDataCache.js";
const { Init } = ObjectDecorators;
const { Redis } = DatabaseDecorators;

/** A pragmatic idle-session lifetime. The spec leaves session lifetime to server policy, not a fixed wire
 * value - a real client re-establishes (`Connect`) transparently whenever its session context has expired. */
export const SESSION_TTL_SECONDS = 15 * 60;

/** The longest a session may live, however active it stays. A client simply `Connect`s again afterwards. */
export const MAX_SESSION_LIFETIME_MS = 24 * 60 * 60 * 1000;

/** The most sessions one user may hold at once. `Connect` past this ends that user's least recently used session (each
 * `Execute` refreshes its session's place - see `MapiSessionManager.touch`). */
export const MAX_SESSIONS_PER_USER = 20;

/** The largest a serialized session may grow. Everything that grows with use is capped on its own (handles by their
 * one-byte index, named properties, MIDs; write streams live outside the session); this is the backstop, and a save
 * past it fails with `MAPI_E_TOO_BIG` rather than storing megabytes every later request has to load. */
export const MAX_SESSION_BYTES = 4 * 1024 * 1024;

/** How long an `Execute`'s session lock lives without being renewed. A running `Execute` renews it every
 * `SESSION_LOCK_RENEW_MS`, so only a crashed pod's lock expires. The session save is still compare-and-set, so an
 * expired lock can't corrupt state. */
export const SESSION_LOCK_TTL_MS = 120 * 1000;

/** How often a running `Execute` renews its session lock and in-progress marker. */
export const SESSION_LOCK_RENEW_MS = SESSION_LOCK_TTL_MS / 3;

/** Tags what a ROP-assigned integer handle (the `ServerObjectHandleTable` index space) refers to.
 * `entityUid` for a `"folder"` handle is one of `session.folderIds`' own value strings (`"virtual:<name>"` or
 * `"folder:<uid>"`), not a bare UID - the same format throughout avoids a second parallel encoding.
 *
 * A `"table"` handle's `columns`/`cursor` hold `RopSetColumns`/`RopQueryRows` state. A hierarchy table carries its
 * resolved child-folder targets in `rows`. A contents table carries no rows at all: `contentsKind` says which repo
 * its folder's items live in and `RopQueryRows` reads just the requested window from the database (see
 * `ContentsTable.ts`), so a large folder never lands in session state.
 *
 * A `"stream"` handle's `entityUid` is the `"message:<uid>"` target its content was opened from,
 * `propertyId`/`propertyType` the `PropertyTag` `RopOpenStream` opened (this pragmatic subset only ever supports
 * `PidTagBody`/`PtypString`, see `MessageBodyStream.ts`), and `streamPosition` how many bytes `RopReadStream` has
 * already returned. A read stream's decoded body lives in `HandleDataCache`, not here. A write stream (opened
 * `ReadWrite`/`Create` against a `RopCreateMessage` draft's `PidTagBody`) records the draft it belongs to in
 * `writeTargetHandleIndex`/`writeTargetGeneration`, and how many bytes have been written in `writeSize`. The bytes
 * themselves are chunks in the `HandleDataStore` (see `RopWriteStreamHandler`), not session state.
 *
 * A `"message"` handle from `RopCreateMessage` (a draft not yet `RopSaveChangesMessage`d) has `entityUid: ""`
 * and instead carries `draftFolderUid` (the folder it will belong to) and `draftProperties` (the small,
 * well-known set of properties this pragmatic subset's `RopSetProperties` tracks - Subject/DisplayTo/
 * DisplayCc/DisplayBcc/an inline `PidTagBody`, each coerced to a plain string, keyed by decimal `PropertyId` -
 * a string key because a JSON-object key is always a string regardless of how it's written). Buffers are stored
 * as base64 and dates as ISO strings since sessions round-trip through `JSON.stringify`/`JSON.parse`.
 *
 * A `"fastTransfer"` handle (`RopFastTransferSourceCopyTo`/`CopyProperties`'s output handle) keeps only what is
 * needed to rebuild its stream (`transferSourceType`, `transferColumns`/`transferExcludeIds`) plus the paging
 * cursor `transferPosition`; the built stream itself lives in `HandleDataCache`.
 *
 * `generation` is a random nonce per assignment (see `assignHandle`), so data keyed by handle index can tell a handle
 * apart from a later one that reuses the same index - including one assigned by a request whose session changes were
 * never saved, which a per-session counter would hand out again. */
export interface MapiObjectHandle {
    type: "logon" | "folder" | "message" | "table" | "stream" | "fastTransfer";
    entityUid: string;
    generation?: string;
    rows?: string[];
    contentsKind?: "message" | "calendarEvent" | "contact" | "task";
    columns?: { propertyId: number; propertyType: number }[];
    cursor?: number;
    propertyId?: number;
    propertyType?: number;
    streamPosition?: number;
    draftFolderUid?: string;
    draftProperties?: Record<string, string>;
    writeTargetHandleIndex?: number;
    writeTargetGeneration?: string;
    writeSize?: number;
    transferSourceType?: "folder" | "message";
    transferColumns?: { propertyId: number; propertyType: number }[];
    transferExcludeIds?: number[];
    transferPosition?: number;
    /** Set on a draft message once `RopSubmitMessage` has tried to send it, whatever the outcome. A submitted draft
     * can't be submitted again or changed. */
    submitted?: boolean;
}

/**
 * The MAPI/HTTP `Session Context` (`[MS-OXCMAPIHTTP]` §3.1.1.1): everything a `Connect`-established session
 * needs across subsequent `Execute` requests. Never persisted to a real database - purely an ephemeral,
 * TTL-bound store entry (see `MapiSessionManager` below), so this deliberately extends `SimpleEntity` (just a
 * `uid`) rather than this library's own `BaseEntity`.
 *
 * `createdAt` is a plain ISO-8601 string, not a `Date`, since sessions round-trip through JSON. `version` is
 * bumped on every successful save and compared on the next one - see `MapiSessionManager.save()`.
 */
export class MapiSessionContext extends SimpleEntity {
    public mailboxUid: string;
    public userUid: string;
    public version = 0;
    public handles: Record<number, MapiObjectHandle> = {};
    public nextHandleIndex = 1;
    public createdAt: string = new Date().toISOString();
    /** This session's FID assignments for the 13 `RopLogon` special folders, keyed by FID (decimal string),
     * valued `"virtual:<name>"` or `"folder:<uid>"` - see `RopLogonHandler`'s own doc comment. Populated by
     * `RopLogon`, read back by a later `RopOpenFolder`. */
    public folderIds: Record<string, string> = {};

    /** The reverse of `folderIds` (target -> FID) plus a monotonic counter, maintained alongside it by
     * `FolderTarget.assignOrGetFid`/`RopLogonHandler` so "does this target already have a FID" is an O(1)
     * lookup instead of a linear scan of `folderIds` repeated once per table row. */
    public folderTargetIds: Record<string, number> = {};
    public nextFolderId = 1;

    /** This session's MID assignments, keyed by MID (decimal string), valued `"message:<uid>"` - the message
     * analog of `folderIds` above. A MID only ever comes into existence lazily, the first time a `RopQueryRows`
     * row exposes a message's `PidTagMid` column (see `MessageTarget.assignOrGetMid`), read back by a later
     * `RopOpenMessage`. */
    public messageIds: Record<string, string> = {};

    /** The reverse of `messageIds` (target -> MID) plus a monotonic counter. */
    public messageTargetIds: Record<string, number> = {};
    public nextMessageId = 1;
    /** The oldest MID still mapped. Past `MAX_MESSAGE_IDS` mapped MIDs the oldest are dropped - see
     * `MessageTarget.assignOrGetMid`. */
    public firstMessageId = 1;

    /** This session's `RopGetPropertyIdsFromNames` mapping table (`[MS-OXCPRPT]` §2.2.12), keyed by a JSON
     * string encoding of the `{guid, kind, lid|name}` `PropertyName` the numeric ID was assigned to - see
     * `NamedPropertyRegistry.ts`. */
    public namedProperties: Record<string, number> = {};

    /** The reverse of `namedProperties` (numeric ID -> `PropertyName` JSON key) plus a monotonic counter. */
    public namedPropertyIds: Record<number, string> = {};
    public nextNamedPropertyId = 0x8000;

    public constructor(other: Partial<SimpleEntity> & { mailboxUid: string; userUid: string }) {
        super(other);
        this.mailboxUid = other.mailboxUid;
        this.userUid = other.userUid;
    }
}

/**
 * Stores `handle` at `index` with a fresh generation. Whatever previously held that index is released first
 * (see `releaseHandle`), so a stream or cached data belonging to the old handle never attaches to the new one.
 */
export function assignHandle(session: MapiSessionContext, index: number, handle: MapiObjectHandle): MapiObjectHandle {
    releaseHandle(session, index);
    handle.generation = crypto.randomUUID();
    session.handles[index] = handle;
    return handle;
}

/**
 * Releases the handle at `index`: drops its cached data, and releases every write stream opened against it
 * (a stream is meaningless once its draft message is gone, and would otherwise attach to whatever reuses the
 * index next).
 */
export function releaseHandle(session: MapiSessionContext, index: number): void {
    const handle = session.handles[index];
    if (!handle) {
        return;
    }
    delete session.handles[index];
    handleDataCache.delete(handleDataKey(session.uid, index, handle.generation));
    if (handle.type === "fastTransfer" || handle.type === "stream") {
        const released = releasedHandleData.get(session) ?? [];
        released.push(handle.type === "fastTransfer" ? handleDataKey(session.uid, index, handle.generation) : writeStreamKey(session.uid, index, handle.generation));
        releasedHandleData.set(session, released);
    }
    for (const [key, candidate] of Object.entries(session.handles)) {
        if (candidate.writeTargetHandleIndex === index && candidate.writeTargetGeneration === handle.generation) {
            releaseHandle(session, Number(key));
        }
    }
}

/** Session object -> the `HandleDataStore` keys of the handles released on it since the last `takeReleasedHandleData`.
 * Kept beside the session, never in it, so it isn't saved. */
const releasedHandleData = new WeakMap<object, string[]>();

/** The shared-store keys (FastTransfer streams, write-stream chunks) of every handle released on `session` since the
 * last call, for the route to delete once the session's changes are saved. */
export function takeReleasedHandleData(session: MapiSessionContext): string[] {
    const released = releasedHandleData.get(session) ?? [];
    releasedHandleData.delete(session);
    return released;
}

/** The outcome of `MapiSessionManager.save()`. `"conflict"` means another request saved the session first;
 * `"missing"` that it expired or was ended; `"tooBig"` that it grew past `MAX_SESSION_BYTES` and was not saved. */
export type SessionSaveResult = "saved" | "conflict" | "missing" | "tooBig";

/**
 * The storage a `MapiSessionManager` runs on: Redis when a `cache` datastore is configured, otherwise memory.
 *
 * A session is two entries: its JSON (`sessionKey`) and its version (`<sessionKey>.version`). A save compares the
 * version entry only, so Redis never has to parse the session JSON.
 */
export interface MapiSessionStore {
    get(key: string): Promise<string | undefined>;
    create(key: string, value: string, ttlSeconds: number): Promise<void>;
    compareAndSet(key: string, expectedVersion: number, value: string, ttlSeconds: number): Promise<"saved" | "conflict" | "missing">;
    delete(key: string): Promise<void>;
    /** Atomically records `sessionId` in the user's index (ordered by last use, starting at `createdAtMs`): drops ids
     * whose session is gone, ends the least recently used sessions until fewer than `max` remain, then adds it. Returns
     * the ids it ended. */
    addToUserIndex(indexKey: string, sessionId: string, createdAtMs: number, max: number, ttlSeconds: number): Promise<string[]>;
    /** `SET NX` with a TTL: `true` when the lock was free and is now held with `token`. */
    acquireLock(key: string, token: string, ttlMs: number): Promise<boolean>;
    /** Releases the lock only if it is still held with `token`. */
    releaseLock(key: string, token: string): Promise<void>;
    /** Extends the lock to `ttlMs` from now, only while it is still held with `token`. */
    renewLock(key: string, token: string, ttlMs: number): Promise<boolean>;
    /** Moves `sessionId` to `score` in the user's index when it is still there. */
    touchUserIndex(indexKey: string, sessionId: string, score: number): Promise<void>;
    getValue(key: string): Promise<string | undefined>;
    putValue(key: string, value: string, ttlSeconds: number): Promise<void>;
}

/** Creates a session at version 0. KEYS: session, version. ARGV: json, ttl. */
export const CREATE_SCRIPT = `
redis.call('SETEX', KEYS[1], ARGV[2], ARGV[1])
redis.call('SETEX', KEYS[2], ARGV[2], '0')
return 1
`;

/** Replaces the session only when its version entry still equals the expected one, without decoding the session.
 * KEYS: session, version. ARGV: expected version, json, ttl. Returns 1 saved, 0 on a version mismatch, -1 when gone. */
export const COMPARE_AND_SET_SCRIPT = `
local current = redis.call('GET', KEYS[2])
if not current or redis.call('EXISTS', KEYS[1]) == 0 then return -1 end
if tonumber(current) ~= tonumber(ARGV[1]) then return 0 end
redis.call('SETEX', KEYS[1], ARGV[3], ARGV[2])
redis.call('SETEX', KEYS[2], ARGV[3], tostring(tonumber(ARGV[1]) + 1))
return 1
`;

/** Adds a session to its user's sorted-set index in one step, so concurrent Connects can't each see room under the
 * cap. KEYS: index. ARGV: session id, created-at score, max, index ttl, session key prefix, session key suffix. The
 * session keys it checks and deletes are built from the ids, so they aren't declared in KEYS (fine on one Redis node,
 * not on a cluster). Returns the ids it ended. */
export const ADD_TO_USER_INDEX_SCRIPT = `
local function sessionKey(id) return ARGV[5] .. id .. ARGV[6] end
for _, id in ipairs(redis.call('ZRANGE', KEYS[1], 0, -1)) do
  if redis.call('EXISTS', sessionKey(id)) == 0 then redis.call('ZREM', KEYS[1], id) end
end
local ended = {}
while redis.call('ZCARD', KEYS[1]) >= tonumber(ARGV[3]) do
  local oldest = redis.call('ZRANGE', KEYS[1], 0, 0)[1]
  redis.call('ZREM', KEYS[1], oldest)
  redis.call('DEL', sessionKey(oldest), sessionKey(oldest) .. '.version')
  table.insert(ended, oldest)
end
redis.call('ZADD', KEYS[1], ARGV[2], ARGV[1])
redis.call('EXPIRE', KEYS[1], ARGV[4])
return ended
`;

/** KEYS: lock. ARGV: token, ttl ms. Returns 1 when acquired. */
export const ACQUIRE_LOCK_SCRIPT = `
if redis.call('SET', KEYS[1], ARGV[1], 'NX', 'PX', ARGV[2]) then return 1 end
return 0
`;

/** KEYS: lock. ARGV: token. Deletes the lock only while this token still holds it. */
export const RELEASE_LOCK_SCRIPT = `
if redis.call('GET', KEYS[1]) == ARGV[1] then redis.call('DEL', KEYS[1]) end
return 1
`;

/** KEYS: lock. ARGV: token, ttl ms. Returns 1 when the token still held the lock and it was extended. */
export const RENEW_LOCK_SCRIPT = `
if redis.call('GET', KEYS[1]) == ARGV[1] then redis.call('PEXPIRE', KEYS[1], ARGV[2]) return 1 end
return 0
`;

/** KEYS: user index. ARGV: score, session id. Updates the score of a member that is still present. */
export const TOUCH_USER_INDEX_SCRIPT = `
redis.call('ZADD', KEYS[1], 'XX', ARGV[1], ARGV[2])
return 1
`;

/** A `MapiSessionStore` on a node-redis client. Every read goes to Redis, never a per-process copy, so each
 * replica sees the same session. */
export class RedisMapiSessionStore implements MapiSessionStore {
    public constructor(private readonly client: any) {}

    public async get(key: string): Promise<string | undefined> {
        return (await this.client.get(key)) ?? undefined;
    }

    public async create(key: string, value: string, ttlSeconds: number): Promise<void> {
        await this.client.eval(CREATE_SCRIPT, { keys: [key, versionKey(key)], arguments: [value, String(ttlSeconds)] });
    }

    public async compareAndSet(key: string, expectedVersion: number, value: string, ttlSeconds: number): Promise<"saved" | "conflict" | "missing"> {
        const result = Number(
            await this.client.eval(COMPARE_AND_SET_SCRIPT, {
                keys: [key, versionKey(key)],
                arguments: [String(expectedVersion), value, String(ttlSeconds)],
            }),
        );
        return result === 1 ? "saved" : result === 0 ? "conflict" : "missing";
    }

    public async delete(key: string): Promise<void> {
        await this.client.del([key, versionKey(key)]);
    }

    public async addToUserIndex(indexKey: string, sessionId: string, createdAtMs: number, max: number, ttlSeconds: number): Promise<string[]> {
        const ended: unknown[] = await this.client.eval(ADD_TO_USER_INDEX_SCRIPT, {
            keys: [indexKey],
            arguments: [sessionId, String(createdAtMs), String(max), String(ttlSeconds), SESSION_KEY_PREFIX, SESSION_KEY_SUFFIX],
        });
        return ended.map(String);
    }

    public async acquireLock(key: string, token: string, ttlMs: number): Promise<boolean> {
        return Number(await this.client.eval(ACQUIRE_LOCK_SCRIPT, { keys: [key], arguments: [token, String(ttlMs)] })) === 1;
    }

    public async releaseLock(key: string, token: string): Promise<void> {
        await this.client.eval(RELEASE_LOCK_SCRIPT, { keys: [key], arguments: [token] });
    }

    public async renewLock(key: string, token: string, ttlMs: number): Promise<boolean> {
        return Number(await this.client.eval(RENEW_LOCK_SCRIPT, { keys: [key], arguments: [token, String(ttlMs)] })) === 1;
    }

    public async touchUserIndex(indexKey: string, sessionId: string, score: number): Promise<void> {
        await this.client.eval(TOUCH_USER_INDEX_SCRIPT, { keys: [indexKey], arguments: [String(score), sessionId] });
    }

    public async getValue(key: string): Promise<string | undefined> {
        return (await this.client.get(key)) ?? undefined;
    }

    public async putValue(key: string, value: string, ttlSeconds: number): Promise<void> {
        await this.client.setEx(key, ttlSeconds, value);
    }
}

/** A single-process `MapiSessionStore`. Values are stored as strings so every load hands out an independent copy,
 * exactly like the Redis store. No method awaits between reading and writing, so each one is atomic. */
export class MemoryMapiSessionStore implements MapiSessionStore {
    private readonly entries = new Map<string, { value: string; expiresAt: number }>();
    private readonly indexes = new Map<string, Map<string, number>>();

    public async get(key: string): Promise<string | undefined> {
        return this.read(key);
    }

    public async create(key: string, value: string, ttlSeconds: number): Promise<void> {
        this.sweep();
        this.write(key, value, ttlSeconds);
        this.write(versionKey(key), "0", ttlSeconds);
    }

    public async compareAndSet(key: string, expectedVersion: number, value: string, ttlSeconds: number): Promise<"saved" | "conflict" | "missing"> {
        const version = this.read(versionKey(key));
        if (version === undefined || this.read(key) === undefined) {
            return "missing";
        }
        if (Number(version) !== expectedVersion) {
            return "conflict";
        }
        this.write(key, value, ttlSeconds);
        this.write(versionKey(key), String(expectedVersion + 1), ttlSeconds);
        return "saved";
    }

    public async delete(key: string): Promise<void> {
        this.entries.delete(key);
        this.entries.delete(versionKey(key));
    }

    public async addToUserIndex(indexKey: string, sessionId: string, createdAtMs: number, max: number): Promise<string[]> {
        const members = this.indexes.get(indexKey) ?? new Map<string, number>();
        this.indexes.set(indexKey, members);
        for (const id of [...members.keys()]) {
            if (this.read(SESSION_KEY_PREFIX + id + SESSION_KEY_SUFFIX) === undefined) {
                members.delete(id);
            }
        }
        const oldestFirst = [...members.entries()].sort((a, b) => a[1] - b[1]).map(([id]) => id);
        const ended: string[] = [];
        while (members.size >= max) {
            const oldest = oldestFirst.shift()!;
            members.delete(oldest);
            this.entries.delete(SESSION_KEY_PREFIX + oldest + SESSION_KEY_SUFFIX);
            this.entries.delete(versionKey(SESSION_KEY_PREFIX + oldest + SESSION_KEY_SUFFIX));
            ended.push(oldest);
        }
        members.set(sessionId, createdAtMs);
        return ended;
    }

    public async acquireLock(key: string, token: string, ttlMs: number): Promise<boolean> {
        if (this.read(key) !== undefined) {
            return false;
        }
        this.entries.set(key, { value: token, expiresAt: Date.now() + ttlMs });
        return true;
    }

    public async releaseLock(key: string, token: string): Promise<void> {
        if (this.read(key) === token) {
            this.entries.delete(key);
        }
    }

    public async renewLock(key: string, token: string, ttlMs: number): Promise<boolean> {
        if (this.read(key) !== token) {
            return false;
        }
        this.entries.set(key, { value: token, expiresAt: Date.now() + ttlMs });
        return true;
    }

    public async touchUserIndex(indexKey: string, sessionId: string, score: number): Promise<void> {
        const members = this.indexes.get(indexKey);
        if (members?.has(sessionId)) {
            members.set(sessionId, score);
        }
    }

    public async getValue(key: string): Promise<string | undefined> {
        return this.read(key);
    }

    public async putValue(key: string, value: string, ttlSeconds: number): Promise<void> {
        this.write(key, value, ttlSeconds);
    }

    private read(key: string): string | undefined {
        const entry = this.entries.get(key);
        if (entry && entry.expiresAt <= Date.now()) {
            this.entries.delete(key);
            return undefined;
        }
        return entry?.value;
    }

    private write(key: string, value: string, ttlSeconds: number): void {
        this.entries.set(key, { value, expiresAt: Date.now() + ttlSeconds * 1000 });
    }

    private sweep(): void {
        const now = Date.now();
        for (const [key, entry] of this.entries) {
            if (entry.expiresAt <= now) {
                this.entries.delete(key);
            }
        }
    }
}

// A session's keys put its id in a hash tag, so its JSON, version, lock and last response share one cluster slot.
const SESSION_KEY_PREFIX = "mapi.session.{";
const SESSION_KEY_SUFFIX = "}";

/** The store key of a session's JSON. */
export function sessionKey(sessionId: string): string {
    return SESSION_KEY_PREFIX + encodeURIComponent(sessionId) + SESSION_KEY_SUFFIX;
}

/** The `.last` value marking `requestId` as running. Always serialized the same way, so it can be compared as a string. */
function inProgressMarker(requestId: string): string {
    return JSON.stringify({ requestId });
}

function versionKey(key: string): string {
    return `${key}.version`;
}

function userIndexKey(userUid: string): string {
    return `mapi.session.user.${encodeURIComponent(userUid)}`;
}

/**
 * Stores/loads `MapiSessionContext`s, keyed by the opaque session id that also becomes the `MapiContext` cookie
 * value.
 *
 * Every load reads the shared store (Redis when a `cache` datastore is configured). An earlier version used
 * `RedisCache`, whose per-process copy served a session without checking Redis, so with several replicas one pod
 * could run ROPs against handles another pod had already changed. Saves are compare-and-set on `version`: when
 * two requests on the same session overlap, the second save reports `"conflict"` instead of overwriting the
 * first request's handles. `Execute` also holds the session's lock (`acquireLock`) while it runs, so an overlapping
 * request is turned away before it has any side effects, and keeps its last response (`storeResponse`) so a retry of
 * the same request is answered again instead of being run twice.
 */
export class MapiSessionManager {
    // Automatically injected by ObjectFactory on instantiation
    private _objectFactory?: ObjectFactory;

    @Redis("cache", false)
    private redisClient?: any;

    private store?: MapiSessionStore;

    /** Where handlers keep FastTransfer streams and write-stream chunks: in Redis next to the sessions when configured,
     * so any replica can continue a transfer. */
    public handleDataStore?: HandleDataStore;

    /** Builds the session store and the handle-data store once, here: Redis-backed when a `cache` datasource is
     * configured, otherwise in-process. Each is skipped when already set. */
    @Init
    public async init(): Promise<void> {
        if (!this._objectFactory) {
            throw new Error("objectFactory is not set.");
        }
        if (!this.store) {
            this.store = this.redisClient
                ? await this._objectFactory.newInstance(RedisMapiSessionStore, {
                      name: "RedisMapiSessionStore",
                      args: [this.redisClient],
                  })
                : await this._objectFactory.newInstance(MemoryMapiSessionStore, { name: "MemoryMapiSessionStore" });
        }
        if (!this.handleDataStore) {
            this.handleDataStore = this.redisClient
                ? await this._objectFactory.newInstance(RedisHandleDataStore, {
                      name: "RedisHandleDataStore",
                      args: [this.redisClient],
                  })
                : await this._objectFactory.newInstance(MemoryHandleDataStore, { name: "MemoryHandleDataStore" });
        }
    }

    /** Creates a session for `userUid`, ending that user's least recently used sessions (and their stored handle data)
     * if they already hold `MAX_SESSIONS_PER_USER`. */
    public async create(mailboxUid: string, userUid: string): Promise<MapiSessionContext> {
        const context = new MapiSessionContext({ mailboxUid, userUid });
        await this.store!.create(sessionKey(context.uid), JSON.stringify(context), SESSION_TTL_SECONDS);
        const ended = await this.store!.addToUserIndex(
            userIndexKey(userUid),
            encodeURIComponent(context.uid),
            Date.now(),
            MAX_SESSIONS_PER_USER,
            MAX_SESSION_LIFETIME_MS / 1000,
        );
        for (const id of ended) {
            await this.deleteHandleData(decodeURIComponent(id));
        }
        return context;
    }

    /** Records that `session` was just used, so the per-user cap ends the least recently used session rather than the
     * oldest one. */
    public async touch(session: MapiSessionContext): Promise<void> {
        await this.store!.touchUserIndex(userIndexKey(session.userUid), encodeURIComponent(session.uid), Date.now());
    }

    /** Loads a session straight from the store. A session past `MAX_SESSION_LIFETIME_MS` is ended and treated as
     * missing. */
    public async load(sessionId: string): Promise<MapiSessionContext | undefined> {
        const json = await this.store!.get(sessionKey(sessionId));
        if (!json) {
            return undefined;
        }
        const session: MapiSessionContext = JSON.parse(json);
        if (Date.now() - Date.parse(session.createdAt) > MAX_SESSION_LIFETIME_MS) {
            await this.destroy(sessionId);
            return undefined;
        }
        return session;
    }

    /** Saves `context` only if nobody else saved it since it was loaded and it is within `MAX_SESSION_BYTES`,
     * bumping its `version` on success. */
    public async save(context: MapiSessionContext): Promise<SessionSaveResult> {
        const expected = context.version;
        context.version = expected + 1;
        const json = JSON.stringify(context);
        const result: SessionSaveResult =
            Buffer.byteLength(json) > MAX_SESSION_BYTES
                ? "tooBig"
                : await this.store!.compareAndSet(sessionKey(context.uid), expected, json, SESSION_TTL_SECONDS);
        if (result !== "saved") {
            context.version = expected;
        }
        return result;
    }

    /** Ends a session, deleting its stored handle data too. */
    public async destroy(sessionId: string): Promise<void> {
        await this.store!.delete(sessionKey(sessionId));
        await this.deleteHandleData(sessionId);
    }

    /** Deletes every FastTransfer stream and write-stream chunk the session stored. Best effort: they expire anyway. */
    private async deleteHandleData(sessionId: string): Promise<void> {
        await this.handleDataStore!.deleteOwnedBy(sessionHandleDataIndex(sessionId)).catch(() => undefined);
    }

    /** Takes the session's `Execute` lock. Returns the token to release it with, or `undefined` while another
     * request holds it. */
    public async acquireLock(sessionId: string): Promise<string | undefined> {
        const token = crypto.randomUUID();
        return (await this.store!.acquireLock(`${sessionKey(sessionId)}.lock`, token, SESSION_LOCK_TTL_MS)) ? token : undefined;
    }

    public async releaseLock(sessionId: string, token: string): Promise<void> {
        await this.store!.releaseLock(`${sessionKey(sessionId)}.lock`, token);
    }

    /** Extends the session's lock by another `SESSION_LOCK_TTL_MS`. `false` when `token` no longer holds it. */
    public async renewLock(sessionId: string, token: string): Promise<boolean> {
        return this.store!.renewLock(`${sessionKey(sessionId)}.lock`, token, SESSION_LOCK_TTL_MS);
    }

    /** The response last sent on this session if it answered `requestId`; `"inProgress"` while a request with that id
     * is still running (or ran on a pod that died less than `SESSION_LOCK_TTL_MS` ago); otherwise `undefined`. */
    public async storedResponse(sessionId: string, requestId: string): Promise<Buffer | "inProgress" | undefined> {
        const json = await this.store!.getValue(`${sessionKey(sessionId)}.last`);
        const stored: { requestId: string; body?: string } | undefined = json ? JSON.parse(json) : undefined;
        if (stored?.requestId !== requestId) {
            return undefined;
        }
        return stored.body === undefined ? "inProgress" : Buffer.from(stored.body, "base64");
    }

    /** Records that `requestId` is running on this session, for `SESSION_LOCK_TTL_MS` (renewed along with the lock). A
     * retry of it is then answered as busy instead of running a second time. */
    public async markInProgress(sessionId: string, requestId: string): Promise<void> {
        await this.store!.putValue(`${sessionKey(sessionId)}.last`, inProgressMarker(requestId), SESSION_LOCK_TTL_MS / 1000);
    }

    /** Extends the in-progress marker of `requestId` by another `SESSION_LOCK_TTL_MS`, only while it is still that marker:
     * once the request's response (or anything else) has replaced it, nothing is written. `false` when it wasn't there. */
    public async renewInProgress(sessionId: string, requestId: string): Promise<boolean> {
        // The lock scripts are exactly "extend/delete only while the value is still this one".
        return this.store!.renewLock(`${sessionKey(sessionId)}.last`, inProgressMarker(requestId), SESSION_LOCK_TTL_MS);
    }

    /** Drops the in-progress marker of `requestId` when it is still there (the request ended without storing a
     * response), so a retry runs. Atomic, so it never deletes a response stored meanwhile. */
    public async clearInProgress(sessionId: string, requestId: string): Promise<void> {
        await this.store!.releaseLock(`${sessionKey(sessionId)}.last`, inProgressMarker(requestId));
    }

    /** Remembers `body` as this session's answer to `requestId`. Only the latest request is kept. */
    public async storeResponse(sessionId: string, requestId: string, body: Buffer): Promise<void> {
        await this.store!.putValue(`${sessionKey(sessionId)}.last`, JSON.stringify({ requestId, body: body.toString("base64") }), SESSION_TTL_SECONDS);
    }
}
