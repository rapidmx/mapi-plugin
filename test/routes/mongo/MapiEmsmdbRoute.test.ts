///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
// These tests prove BaseMapiEmsmdbRoute's transport plumbing (JWT auth, mailbox resolution, session-cookie
// establishment/teardown, the Execute envelope + RopBuffer framing round trip) AND the RopLogon/RopRelease/
// RopOpenFolder/RopGetHierarchyTable/RopGetContentsTable/RopSetColumns/RopQueryRows/RopOpenMessage/
// RopGetPropertiesSpecific/RopOpenStream/RopReadStream/RopGetPropertyIdsFromNames ROPs' real behavior, all over
// a real HTTP round trip - see the architecture plan's "Phase 3" section, build-order steps 3-9a.
//
// Uses `mapiTestClient.ts`'s raw-socket `mapiRequest()` instead of the shared `@rapidrest/service-core/test`
// `request()` helper - see that file's doc comment for why (the shared helper's axios `responseType: "text"`
// reading corrupts binary response bytes >= 0x80, which MAPI/HTTP's raw integers routinely contain).
import config from "../../config.js";
import { MongoConnection, MongoRepository, Server, ObjectFactory, ConnectionManager, ACLAction } from "@rapidrest/service-core";
import { JWTUtils, Logger } from "@rapidrest/core";
import * as uuid from "uuid";
import { ContactMongo, FolderMongo, LabelMongo, MailboxMongo, MeetingSchedulingJobMongo, MessageMongo, TaskMongo } from "@rapidmx/restapi/mongo";
import { FolderType, MessageImportance, RecipientType } from "@rapidmx/restapi";
import { MongoMemoryServer } from "mongodb-memory-server";
import { InMemoryBlobStore, RecordingMailTransport, registerTestDoubles } from "../../testDoubles.js";
import { cookieHeaderFrom, mapiRequest } from "../../mapiTestClient.js";
import { BufferReader, BufferWriter } from "../../../src/codec/BufferCursor.js";
import { decodeGuid, encodeGuid } from "../../../src/codec/MapiGuid.js";
import { PropertyType, readPropertyValue, writePropertyTag, writeTaggedPropertyValue, type TaggedPropertyValue } from "../../../src/codec/PropertyValue.js";
import { decodeRopBuffer, encodeRopBuffer } from "../../../src/codec/RopBuffer.js";
import { MapiSessionManager, SESSION_LOCK_RENEW_MS } from "../../../src/MapiSessionManager.js";
import { handleDataCache, handleDataKey, MemoryHandleDataStore } from "../../../src/rop/HandleDataCache.js";

const mongod: MongoMemoryServer = new MongoMemoryServer({
    instance: {
        port: 9999,
        dbName: "rrst-test",
    },
});

describe("Route:MapiEmsmdbRouteMongo Tests", () => {
    const logger = Logger();
    const objectFactory: ObjectFactory = new ObjectFactory(config, logger);
    const server: Server = new Server({ config, basePath: "./test/server-mongo", logger, objectFactory });
    const baseUrl = "/mongo/mapi/emsmdb";
    let mailboxRepo: MongoRepository<MailboxMongo>;
    let folderRepo: MongoRepository<FolderMongo>;
    let messageRepo: MongoRepository<MessageMongo>;
    let contactRepo: MongoRepository<ContactMongo>;
    let taskRepo: MongoRepository<TaskMongo>;
    let labelRepo: MongoRepository<LabelMongo>;
    let aclRepo: MongoRepository<any>;

    const owner: any = { uid: uuid.v4(), roles: [], elevated: Date.now() };
    const ownerToken = JWTUtils.createTokenSync(config.get("auth"), owner);

    const createMailbox = async function (ownerUid: string, data?: Partial<MailboxMongo>): Promise<MailboxMongo> {
        const obj: MailboxMongo = new MailboxMongo({
            ownerUserUid: ownerUid,
            primarySmtpAddress: `${uuid.v4()}@example.com`,
            aliasAddresses: [],
            displayName: "Test Mailbox",
            timezone: "UTC",
            quotaBytes: 1_000_000_000,
            usedBytes: 0,
            ...data,
        });
        const result: MailboxMongo = await mailboxRepo.save(obj);
        await aclRepo.save({
            uid: result.uid,
            dateCreated: new Date(),
            dateModified: new Date(),
            version: 0,
            records: [{ userOrRoleId: ownerUid, actions: [ACLAction.FULL] }],
            parentUid: "Mailbox",
        });
        return result;
    };

    const createFolder = async function (mailboxUid: string, type: FolderType, name: string, parentFolderUid?: string): Promise<FolderMongo> {
        return await folderRepo.save(
            new FolderMongo({ mailboxUid, name, type, parentFolderUid, unreadCount: 0, totalCount: 0, syncKeyVersion: 0 }),
        );
    };

    const createMessage = async function (mailboxUid: string, folderUid: string, data?: Partial<MessageMongo>): Promise<MessageMongo> {
        return await messageRepo.save(
            new MessageMongo({
                mailboxUid,
                folderUid,
                messageId: `${uuid.v4()}@example.com`,
                subject: "Test Subject",
                from: { address: "sender@example.com", displayName: "Sender", type: RecipientType.TO },
                recipients: [{ address: "owner@example.com", type: RecipientType.TO }],
                sentDate: new Date(),
                receivedDate: new Date("2026-03-15T09:00:00.000Z"),
                bodyBlobKey: `bodies/${uuid.v4()}`,
                bodyPreview: "Hello world",
                flags: { read: true, flagged: false, answered: false, forwarded: false },
                importance: MessageImportance.NORMAL,
                references: [],
                hasAttachments: false,
                ...data,
            }),
        );
    };

    const createContact = async function (mailboxUid: string, folderUid: string, data?: Partial<ContactMongo>): Promise<ContactMongo> {
        return await contactRepo.save(
            new ContactMongo({ mailboxUid, folderUid, displayName: "Test Contact", emails: [], phones: [], addresses: [], ...data }),
        );
    };

    const createTask = async function (mailboxUid: string, folderUid: string, data?: Partial<TaskMongo>): Promise<TaskMongo> {
        return await taskRepo.save(new TaskMongo({ mailboxUid, folderUid, title: "Test Task", ...data }));
    };

    const createLabel = async function (mailboxUid: string, data?: Partial<LabelMongo>): Promise<LabelMongo> {
        return await labelRepo.save(new LabelMongo({ mailboxUid, name: "Test Label", ...data }));
    };

    const blobStore = function (): InMemoryBlobStore {
        return objectFactory.getInstance<InMemoryBlobStore>("BlobStore")!;
    };

    const mailTransport = function (): RecordingMailTransport {
        return objectFactory.getInstance<RecordingMailTransport>("MailTransport")!;
    };

    const connect = async function (headers: Record<string, string> = {}) {
        return mapiRequest(
            server.getApplication(),
            baseUrl,
            {
                Authorization: "jwt " + ownerToken,
                "X-RequestType": "Connect",
                "Content-Type": "application/mapi-http",
                ...headers,
            },
            Buffer.alloc(0),
        );
    };

    const execute = async function (cookie: string, ropBuffer: Buffer) {
        const body = new BufferWriter();
        body.writeUInt32LE(0); // Flags
        body.writeUInt32LE(ropBuffer.length);
        body.writeBytes(ropBuffer);
        body.writeUInt32LE(256 * 1024); // MaxRopOut
        body.writeUInt32LE(0); // AuxiliaryBufferSize

        return mapiRequest(
            server.getApplication(),
            baseUrl,
            {
                Authorization: "jwt " + ownerToken,
                "X-RequestType": "Execute",
                "Content-Type": "application/mapi-http",
                ...(cookie ? { Cookie: cookie } : {}),
            },
            body.toBuffer(),
        );
    };

    beforeAll(async () => {
        await mongod.start();
        registerTestDoubles(objectFactory);
        await server.start();

        const connMgr: ConnectionManager | undefined = objectFactory.getInstance(ConnectionManager);
        const conn: any = connMgr?.connections.get("mongo");
        if (conn instanceof MongoConnection) {
            mailboxRepo = conn.getMongoRepository("MailboxMongo");
            folderRepo = conn.getMongoRepository("FolderMongo");
            messageRepo = conn.getMongoRepository("MessageMongo");
            contactRepo = conn.getMongoRepository("ContactMongo");
            taskRepo = conn.getMongoRepository("TaskMongo");
            labelRepo = conn.getMongoRepository("LabelMongo");
        } else {
            throw new Error("Could not find mongo connection");
        }
        const aclConn: any = connMgr?.connections.get("acl");
        if (aclConn instanceof MongoConnection) {
            aclRepo = aclConn.getMongoRepository("AccessControlListMongo");
        } else {
            throw new Error("Could not find mongo acl connection");
        }
    });

    afterAll(async () => {
        await server.stop();
        await mongod.stop();
        await objectFactory.destroy();
    });

    beforeEach(async () => {
        for (const repo of [mailboxRepo, folderRepo, messageRepo, contactRepo, taskRepo, labelRepo, aclRepo]) {
            try {
                await repo.clear();
            } catch (err: any) {
                if (err.message !== "ns not found") {
                    throw err;
                }
            }
        }
    });

    describe("Connect", () => {
        it("Requires authentication.", async () => {
            const result = await mapiRequest(
                server.getApplication(),
                baseUrl,
                { "X-RequestType": "Connect", "Content-Type": "application/mapi-http" },
                Buffer.alloc(0),
            );
            expect(result.status).toBe(401);
        });

        it("Returns 404 when the caller owns no mailbox.", async () => {
            const result = await connect();
            expect(result.status).toBe(404);
        });

        it("Establishes a session, setting both MapiContext and MapiSequence cookies.", async () => {
            await createMailbox(owner.uid);
            const result = await connect();
            expect(result.status).toBe(200);
            const cookie = cookieHeaderFrom(result.headers["set-cookie"]);
            expect(cookie).toContain("MapiContext=");
            expect(cookie).toContain("MapiSequence=");
        });

        it("Returns a well-formed success response body (StatusCode/ErrorCode both zero) with the DisplayName echoed.", async () => {
            await createMailbox(owner.uid, { displayName: "Ada Lovelace" });
            const result = await connect();

            const reader = new BufferReader(result.body);
            expect(reader.readUInt32LE()).toBe(0); // StatusCode
            expect(reader.readUInt32LE()).toBe(0); // ErrorCode
            expect(reader.readUInt32LE()).toBe(60000); // PollsMax
            expect(reader.readUInt32LE()).toBe(3); // RetryCount
            expect(reader.readUInt32LE()).toBe(5000); // RetryDelay
            reader.readNullTerminatedString8(); // DnPrefix
            expect(reader.readNullTerminatedUtf16LE()).toBe("Ada Lovelace");
            expect(reader.readUInt32LE()).toBe(0); // AuxiliaryBufferSize
            expect(reader.hasMore()).toBe(false);
        });

        it("Sets the common MAPI/HTTP response headers.", async () => {
            await createMailbox(owner.uid);
            const result = await connect({ "X-RequestId": "req-123" });

            expect(result.headers["content-type"]).toContain("application/mapi-http");
            expect(result.headers["x-requesttype"]).toBe("Connect");
            expect(result.headers["x-requestid"]).toBe("req-123");
            expect(result.headers["x-responsecode"]).toBe("0");
        });
    });

    describe("Execute", () => {
        it("Returns X-ResponseCode 10 (Context Not Found), with the HRESULT only in the body, when no MapiContext cookie is presented.", async () => {
            await createMailbox(owner.uid);
            const emptyRop = encodeRopBuffer({ ropsList: Buffer.alloc(0), handleTable: [] });
            const result = await execute("", emptyRop);

            expect(result.status).toBe(200);
            expect(result.headers["x-responsecode"]).toBe("10");
            const reader = new BufferReader(result.body);
            expect(reader.readUInt32LE()).toBe(0); // StatusCode
            expect(reader.readUInt32LE()).toBe(0x80040111); // ErrorCode
        });

        it("Echoes back an empty but well-formed ROP buffer, preserving the handle table, for a valid session.", async () => {
            await createMailbox(owner.uid);
            const connectResult = await connect();
            const cookie = cookieHeaderFrom(connectResult.headers["set-cookie"]);

            const inputRop = encodeRopBuffer({ ropsList: Buffer.alloc(0), handleTable: [7, 9] });
            const result = await execute(cookie, inputRop);

            expect(result.status).toBe(200);
            expect(result.headers["x-responsecode"]).toBe("0");
            const reader = new BufferReader(result.body);
            expect(reader.readUInt32LE()).toBe(0); // StatusCode
            expect(reader.readUInt32LE()).toBe(0); // ErrorCode
            reader.readUInt32LE(); // Flags
            const responseRopBufferSize = reader.readUInt32LE();
            const responseRopBuffer = reader.readBytes(responseRopBufferSize);
            // RopSize field (2 bytes) + empty ropsList = 2, then the two preserved handle-table entries.
            expect(responseRopBuffer.readUInt16LE(0)).toBe(2);
            expect(responseRopBuffer.readUInt32LE(2)).toBe(7);
            expect(responseRopBuffer.readUInt32LE(6)).toBe(9);
            expect(reader.readUInt32LE()).toBe(0); // AuxiliaryBufferSize
        });

        it("Returns 400 for a RopBufferSize larger than the ROP buffer limit or the body itself.", async () => {
            await createMailbox(owner.uid);
            const cookie = cookieHeaderFrom((await connect()).headers["set-cookie"]);
            const oversized = new BufferWriter().writeUInt32LE(0).writeUInt32LE(0x7fffffff).writeBytes(Buffer.alloc(16)).toBuffer();

            const result = await mapiRequest(
                server.getApplication(),
                baseUrl,
                { Authorization: "jwt " + ownerToken, "X-RequestType": "Execute", "Content-Type": "application/mapi-http", Cookie: cookie },
                oversized,
            );

            expect(result.status).toBe(400);
        });

        it("Ends the session when the caller no longer owns the session's mailbox.", async () => {
            const mailbox = await createMailbox(owner.uid);
            const cookie = cookieHeaderFrom((await connect()).headers["set-cookie"]);
            await mailboxRepo.updateOne({ uid: mailbox.uid } as any, { $set: { ownerUserUid: uuid.v4() } });

            const result = await execute(cookie, encodeRopBuffer({ ropsList: Buffer.alloc(0), handleTable: [] }));

            expect(result.status).toBe(200);
            expect(result.headers["x-responsecode"]).toBe("10");
        });

        it("Fails a request whose session was saved by an overlapping request first, with X-ResponseCode 15, and a vanished session as not-found.", async () => {
            await createMailbox(owner.uid);
            const cookie = cookieHeaderFrom((await connect()).headers["set-cookie"]);
            const save = vi.spyOn(MapiSessionManager.prototype, "save").mockResolvedValueOnce("conflict").mockResolvedValueOnce("missing");
            try {
                const conflict = await execute(cookie, encodeRopBuffer({ ropsList: Buffer.alloc(0), handleTable: [] }));
                expect(conflict.headers["x-responsecode"]).toBe("15");
                const conflictBody = new BufferReader(conflict.body);
                conflictBody.readUInt32LE();
                expect(conflictBody.readUInt32LE()).toBe(15);

                const missing = await execute(cookie, encodeRopBuffer({ ropsList: Buffer.alloc(0), handleTable: [] }));
                expect(missing.headers["x-responsecode"]).toBe("10");
                expect(new BufferReader(missing.body).readBytes(8).readUInt32LE(4)).toBe(0x80040111);
            } finally {
                save.mockRestore();
            }

            const ok = await execute(cookie, encodeRopBuffer({ ropsList: Buffer.alloc(0), handleTable: [] }));
            expect(ok.headers["x-responsecode"]).toBe("0");
        });

        it("Turns away a request that overlaps another one on the same session before running any ROP.", async () => {
            await createMailbox(owner.uid);
            const cookie = cookieHeaderFrom((await connect()).headers["set-cookie"]);
            const lock = vi.spyOn(MapiSessionManager.prototype, "acquireLock").mockResolvedValueOnce(undefined);
            const save = vi.spyOn(MapiSessionManager.prototype, "save");
            try {
                const overlapping = await execute(cookie, encodeRopBuffer({ ropsList: Buffer.from([0x01, 0x00, 0x00]), handleTable: [0] }));
                expect(overlapping.headers["x-responsecode"]).toBe("15");
                expect(save).not.toHaveBeenCalled();
            } finally {
                lock.mockRestore();
                save.mockRestore();
            }
            // The lock is released after every request, so the next one runs.
            const next = await execute(cookie, encodeRopBuffer({ ropsList: Buffer.alloc(0), handleTable: [] }));
            expect(next.headers["x-responsecode"]).toBe("0");
        });

        it("Answers a retried X-RequestId with the stored response instead of running its ROPs again.", async () => {
            await createMailbox(owner.uid);
            const cookie = cookieHeaderFrom((await connect()).headers["set-cookie"]);
            const logon = new BufferWriter().writeUInt8(0xfe).writeUInt8(0).writeUInt8(0).writeUInt8(0x01).writeUInt32LE(0).writeUInt32LE(0).writeUInt16LE(0).toBuffer();
            const body = new BufferWriter();
            const ropBuffer = encodeRopBuffer({ ropsList: logon, handleTable: [0xffffffff] });
            body.writeUInt32LE(0).writeUInt32LE(ropBuffer.length).writeBytes(ropBuffer).writeUInt32LE(0x10008).writeUInt32LE(0);
            const send = (requestId: string) =>
                mapiRequest(
                    server.getApplication(),
                    baseUrl,
                    { Authorization: "jwt " + ownerToken, "X-RequestType": "Execute", "Content-Type": "application/mapi-http", Cookie: cookie, "X-RequestId": requestId },
                    body.toBuffer(),
                );

            const first = await send("{guid}:1");
            const save = vi.spyOn(MapiSessionManager.prototype, "save");
            try {
                const retry = await send("{guid}:1");
                expect(retry.headers["x-responsecode"]).toBe("0");
                expect(retry.body.equals(first.body)).toBe(true);
                expect(save).not.toHaveBeenCalled();

                await send("{guid}:2");
                expect(save).toHaveBeenCalledTimes(1);
            } finally {
                save.mockRestore();
            }
        });

        it("Answers 400 for a ROP it can't decode, after saving the handles the ROPs before it created.", async () => {
            await createMailbox(owner.uid);
            const cookie = cookieHeaderFrom((await connect()).headers["set-cookie"]);
            const logon = new BufferWriter().writeUInt8(0xfe).writeUInt8(0).writeUInt8(0).writeUInt8(0x01).writeUInt32LE(0).writeUInt32LE(0).writeUInt16LE(0).toBuffer();
            const save = vi.spyOn(MapiSessionManager.prototype, "save");
            try {
                const result = await execute(cookie, encodeRopBuffer({ ropsList: Buffer.concat([logon, Buffer.from([0x99, 0x00])]), handleTable: [0xffffffff] }));
                expect(result.status).toBe(400);
                expect(save).toHaveBeenCalledTimes(1);
                expect((await save.mock.results[0].value)).toBe("saved");
                expect(save.mock.calls[0][0].handles[0]).toMatchObject({ type: "logon" });
            } finally {
                save.mockRestore();
            }
        });

        it("Holds ROP responses to MaxRopOut, answering RopBufferTooSmall for one that doesn't fit.", async () => {
            await createMailbox(owner.uid);
            const cookie = cookieHeaderFrom((await connect()).headers["set-cookie"]);
            const logon = new BufferWriter().writeUInt8(0xfe).writeUInt8(0).writeUInt8(0).writeUInt8(0x01).writeUInt32LE(0).writeUInt32LE(0).writeUInt16LE(0).toBuffer();
            const ropBuffer = encodeRopBuffer({ ropsList: logon, handleTable: [0xffffffff] });
            const body = new BufferWriter().writeUInt32LE(0).writeUInt32LE(ropBuffer.length).writeBytes(ropBuffer).writeUInt32LE(40).writeUInt32LE(0).toBuffer();

            const result = await mapiRequest(
                server.getApplication(),
                baseUrl,
                { Authorization: "jwt " + ownerToken, "X-RequestType": "Execute", "Content-Type": "application/mapi-http", Cookie: cookie },
                body,
            );

            const reader = new BufferReader(result.body);
            reader.readBytes(12);
            const { ropsList } = decodeRopBuffer(reader.readBytes(reader.readUInt32LE()));
            // 40 - RopSize (2) - one handle table entry (4) leaves 34 bytes: too few for RopLogon's response.
            expect(ropsList[0]).toBe(0xff);
            expect(ropsList.readUInt16LE(1)).toBeGreaterThan(34);
            expect(ropsList.subarray(3)).toEqual(logon);
        });

        it("Fails the whole Execute with ecBufferTooSmall when not even RopBufferTooSmall fits in MaxRopOut.", async () => {
            await createMailbox(owner.uid);
            const cookie = cookieHeaderFrom((await connect()).headers["set-cookie"]);
            const logon = new BufferWriter().writeUInt8(0xfe).writeUInt8(0).writeUInt8(0).writeUInt8(0x01).writeUInt32LE(0).writeUInt32LE(0).writeUInt16LE(0).toBuffer();
            const ropBuffer = encodeRopBuffer({ ropsList: logon, handleTable: [0xffffffff] });
            // 16 - RopSize (2) - one handle (4) leaves 10 bytes: less than RopBufferTooSmall's 3 plus the 14-byte request.
            const body = new BufferWriter().writeUInt32LE(0).writeUInt32LE(ropBuffer.length).writeBytes(ropBuffer).writeUInt32LE(16).writeUInt32LE(0).toBuffer();

            const result = await mapiRequest(
                server.getApplication(),
                baseUrl,
                { Authorization: "jwt " + ownerToken, "X-RequestType": "Execute", "Content-Type": "application/mapi-http", Cookie: cookie },
                body,
            );

            expect(result.headers["x-responsecode"]).toBe("0");
            expect(result.body.readUInt32LE(4)).toBe(0x47d);
            expect(result.body.readUInt32LE(12)).toBe(0); // no RopBuffer
        });

        it("Checks MaxRopOut before running any ROP, so an ecBufferTooSmall answer leaves nothing changed or stored.", async () => {
            await createMailbox(owner.uid);
            const cookie = cookieHeaderFrom((await connect()).headers["set-cookie"]);
            const logon = new BufferWriter().writeUInt8(0xfe).writeUInt8(0).writeUInt8(0).writeUInt8(0x01).writeUInt32LE(0).writeUInt32LE(0).writeUInt16LE(0).toBuffer();
            const ropBuffer = encodeRopBuffer({ ropsList: logon, handleTable: [0xffffffff] });
            const send = (maxRopOut: number) =>
                mapiRequest(
                    server.getApplication(),
                    baseUrl,
                    { Authorization: "jwt " + ownerToken, "X-RequestType": "Execute", "Content-Type": "application/mapi-http", Cookie: cookie, "X-RequestId": "{guid}:21" },
                    new BufferWriter().writeUInt32LE(0).writeUInt32LE(ropBuffer.length).writeBytes(ropBuffer).writeUInt32LE(maxRopOut).writeUInt32LE(0).toBuffer(),
                );
            const save = vi.spyOn(MapiSessionManager.prototype, "save");
            const storeResponse = vi.spyOn(MapiSessionManager.prototype, "storeResponse");
            try {
                const tooSmall = await send(16);
                expect(tooSmall.body.readUInt32LE(4)).toBe(0x47d);
                expect(save.mock.calls[0][0].handles[0]).toBeUndefined(); // RopLogon never ran
                expect(storeResponse).not.toHaveBeenCalled();

                // A retry of the same request id with room runs normally.
                const retry = await send(0x10008);
                expect(retry.body.readUInt32LE(4)).toBe(0);
                expect(save.mock.calls[1][0].handles[0]).toMatchObject({ type: "logon" });
                expect(storeResponse).toHaveBeenCalledTimes(1);
            } finally {
                save.mockRestore();
                storeResponse.mockRestore();
            }
        });

        it("Answers RopBufferTooSmall for a RopReadStream with no room, storing that response for a retry and keeping the stream position.", async () => {
            await createMailbox(owner.uid);
            const cookie = cookieHeaderFrom((await connect()).headers["set-cookie"]);
            const sessionId = decodeURIComponent(/MapiContext=([^;]+)/.exec(cookie)![1]);
            const readStream = new BufferWriter().writeUInt8(0x2c).writeUInt8(0).writeUInt8(6).writeUInt16LE(10).toBuffer();
            const loadSession = MapiSessionManager.prototype.load;
            const load = vi.spyOn(MapiSessionManager.prototype, "load").mockImplementation(async function (this: MapiSessionManager, id: string) {
                const session = await loadSession.call(this, id);
                if (session && !session.handles[6]) {
                    session.handles[6] = { type: "stream", entityUid: "", streamPosition: 0, generation: "g" };
                }
                return session;
            });
            handleDataCache.set(handleDataKey(sessionId, 6, "g"), Buffer.from("stream body bytes"));
            const ropBuffer = encodeRopBuffer({ ropsList: readStream, handleTable: [0xffffffff] });
            // RopSize (2) and one handle (4) leave 8 bytes: exactly a RopBufferTooSmall for the 5-byte request, and no data.
            const body = new BufferWriter().writeUInt32LE(0).writeUInt32LE(ropBuffer.length).writeBytes(ropBuffer).writeUInt32LE(14).writeUInt32LE(0).toBuffer();
            const send = () =>
                mapiRequest(
                    server.getApplication(),
                    baseUrl,
                    { Authorization: "jwt " + ownerToken, "X-RequestType": "Execute", "Content-Type": "application/mapi-http", Cookie: cookie, "X-RequestId": "{guid}:22" },
                    body,
                );
            const save = vi.spyOn(MapiSessionManager.prototype, "save");
            try {
                const result = await send();
                expect(result.headers["x-responsecode"]).toBe("0");
                expect(result.body.readUInt32LE(4)).toBe(0);
                const reader = new BufferReader(result.body);
                reader.readBytes(12);
                const responseRops = decodeRopBuffer(reader.readBytes(reader.readUInt32LE())).ropsList;
                expect(responseRops[0]).toBe(0xff);
                expect(responseRops.readUInt16LE(1)).toBe(8 + 10); // SizeNeeded: the header plus the 10 bytes asked for
                expect(responseRops.subarray(3)).toEqual(readStream);
                expect(save.mock.calls[0][0].handles[6].streamPosition).toBe(0);

                const retry = await send();
                expect(retry.body.equals(result.body)).toBe(true);
                expect(save).toHaveBeenCalledTimes(1);
            } finally {
                load.mockRestore();
                save.mockRestore();
                handleDataCache.delete(handleDataKey(sessionId, 6, "g"));
            }
        });

        it("Checks session ownership before taking the lock, and refreshes the session's place in the per-user cap.", async () => {
            await createMailbox(owner.uid);
            const cookie = cookieHeaderFrom((await connect()).headers["set-cookie"]);
            const stranger: any = { uid: uuid.v4(), roles: [], elevated: Date.now() };
            const lock = vi.spyOn(MapiSessionManager.prototype, "acquireLock");
            const touch = vi.spyOn(MapiSessionManager.prototype, "touch").mockRejectedValueOnce(new Error("redis down"));
            try {
                const foreign = await mapiRequest(
                    server.getApplication(),
                    baseUrl,
                    {
                        Authorization: "jwt " + JWTUtils.createTokenSync(config.get("auth"), stranger),
                        "X-RequestType": "Execute",
                        "Content-Type": "application/mapi-http",
                        Cookie: cookie,
                    },
                    new BufferWriter().writeUInt32LE(0).writeUInt32LE(0).toBuffer(),
                );
                expect(foreign.headers["x-responsecode"]).toBe("10");
                expect(lock).not.toHaveBeenCalled();

                // A failing touch doesn't fail the request.
                const own = await execute(cookie, encodeRopBuffer({ ropsList: Buffer.alloc(0), handleTable: [] }));
                expect(own.headers["x-responsecode"]).toBe("0");
                expect(touch).toHaveBeenCalledTimes(1);
            } finally {
                lock.mockRestore();
                touch.mockRestore();
            }
        });

        it("Answers a retry of a request that is still running with X-ResponseCode 15, and drops the marker of a request that failed.", async () => {
            await createMailbox(owner.uid);
            const cookie = cookieHeaderFrom((await connect()).headers["set-cookie"]);
            const sessionId = /MapiContext=([^;]+)/.exec(cookie)![1];
            const empty = encodeRopBuffer({ ropsList: Buffer.alloc(0), handleTable: [] });
            const body = new BufferWriter().writeUInt32LE(0).writeUInt32LE(empty.length).writeBytes(empty).writeUInt32LE(0x10008).writeUInt32LE(0).toBuffer();
            const send = (requestId: string) =>
                mapiRequest(
                    server.getApplication(),
                    baseUrl,
                    { Authorization: "jwt " + ownerToken, "X-RequestType": "Execute", "Content-Type": "application/mapi-http", Cookie: cookie, "X-RequestId": requestId },
                    body,
                );

            const stored = vi.spyOn(MapiSessionManager.prototype, "storedResponse").mockResolvedValueOnce("inProgress");
            const save = vi.spyOn(MapiSessionManager.prototype, "save");
            try {
                const retry = await send("{guid}:7");
                expect(retry.headers["x-responsecode"]).toBe("15");
                expect(save).not.toHaveBeenCalled();
            } finally {
                stored.mockRestore();
            }

            // A session that ends while the request waits for its lock is not found.
            const originalLoad = MapiSessionManager.prototype.load;
            const load = vi
                .spyOn(MapiSessionManager.prototype, "load")
                .mockImplementationOnce(function (this: MapiSessionManager, id: string) {
                    return originalLoad.call(this, id);
                })
                .mockResolvedValueOnce(undefined);
            try {
                expect((await send("{guid}:6")).headers["x-responsecode"]).toBe("10");
            } finally {
                load.mockRestore();
            }

            // A request that fails after marking itself in progress (a conflicting save) leaves no marker behind.
            save.mockResolvedValueOnce("conflict");
            const markInProgress = vi.spyOn(MapiSessionManager.prototype, "markInProgress");
            const clearInProgress = vi.spyOn(MapiSessionManager.prototype, "clearInProgress");
            try {
                expect((await send("{guid}:8")).headers["x-responsecode"]).toBe("15");
                expect(markInProgress).toHaveBeenCalledWith(sessionId, "{guid}:8");
                expect(clearInProgress).toHaveBeenCalledWith(sessionId, "{guid}:8");
                const again = await send("{guid}:8");
                expect(again.headers["x-responsecode"]).toBe("0");
                // Dropping the marker is best effort.
                clearInProgress.mockRejectedValueOnce(new Error("redis down"));
                save.mockResolvedValueOnce("conflict");
                expect((await send("{guid}:10")).headers["x-responsecode"]).toBe("15");
                expect(save).toHaveBeenCalledTimes(3);
            } finally {
                save.mockRestore();
                markInProgress.mockRestore();
                clearInProgress.mockRestore();
            }
        });

        it("Renews the session lock and the in-progress marker while a request runs, best effort, one renewal at a time.", async () => {
            await createMailbox(owner.uid);
            const cookie = cookieHeaderFrom((await connect()).headers["set-cookie"]);
            const empty = encodeRopBuffer({ ropsList: Buffer.alloc(0), handleTable: [] });
            const body = new BufferWriter().writeUInt32LE(0).writeUInt32LE(empty.length).writeBytes(empty).writeUInt32LE(0x10008).writeUInt32LE(0).toBuffer();
            const realSetInterval = global.setInterval;
            let tick: (() => void) | undefined;
            const interval = vi.spyOn(global, "setInterval").mockImplementation(((callback: () => void, ms: number) => {
                if (ms === SESSION_LOCK_RENEW_MS) {
                    tick = callback;
                }
                return realSetInterval(() => undefined, 1 << 30);
            }) as any);
            const renewLock = vi.spyOn(MapiSessionManager.prototype, "renewLock").mockRejectedValueOnce(new Error("redis down"));
            const markInProgress = vi.spyOn(MapiSessionManager.prototype, "markInProgress");
            const renewInProgress = vi.spyOn(MapiSessionManager.prototype, "renewInProgress");
            const originalSave = MapiSessionManager.prototype.save;
            // The renewal fires while the request is still running (during its save): once failing, once succeeding.
            const save = vi.spyOn(MapiSessionManager.prototype, "save").mockImplementationOnce(async function (this: MapiSessionManager, session: any) {
                renewInProgress.mockRejectedValueOnce(new Error("redis down"));
                tick!();
                tick!();
                await new Promise((resolve) => setTimeout(resolve, 20));
                return originalSave.call(this, session);
            });
            const send = () =>
                mapiRequest(
                    server.getApplication(),
                    baseUrl,
                    { Authorization: "jwt " + ownerToken, "X-RequestType": "Execute", "Content-Type": "application/mapi-http", Cookie: cookie, "X-RequestId": "{guid}:9" },
                    body,
                );
            try {
                const result = await send();
                expect(result.headers["x-responsecode"]).toBe("0");
                expect(renewLock).toHaveBeenCalledTimes(2);
                expect(markInProgress).toHaveBeenCalledTimes(1); // only at the start
                expect(renewInProgress).toHaveBeenCalledTimes(2);

                // A renewal landing after the response was stored finds the lock released and writes nothing, so a retry
                // still gets the stored response instead of an in-progress answer.
                tick!();
                await new Promise((resolve) => setTimeout(resolve, 20));
                expect(renewLock).toHaveBeenCalledTimes(3);
                expect(renewInProgress).toHaveBeenCalledTimes(2);
                const retry = await send();
                expect(retry.headers["x-responsecode"]).toBe("0");
                expect(retry.body.equals(result.body)).toBe(true);
            } finally {
                interval.mockRestore();
                renewLock.mockRestore();
                markInProgress.mockRestore();
                renewInProgress.mockRestore();
                save.mockRestore();
            }
        });

        it("Stops renewing once the lock is lost.", async () => {
            await createMailbox(owner.uid);
            const cookie = cookieHeaderFrom((await connect()).headers["set-cookie"]);
            const empty = encodeRopBuffer({ ropsList: Buffer.alloc(0), handleTable: [] });
            const body = new BufferWriter().writeUInt32LE(0).writeUInt32LE(empty.length).writeBytes(empty).writeUInt32LE(0x10008).writeUInt32LE(0).toBuffer();
            const realSetInterval = global.setInterval;
            let tick: (() => void) | undefined;
            let renewalTimer: any;
            const interval = vi.spyOn(global, "setInterval").mockImplementation(((callback: () => void, ms: number) => {
                const timer = realSetInterval(() => undefined, 1 << 30);
                if (ms === SESSION_LOCK_RENEW_MS) {
                    tick = callback;
                    renewalTimer = timer;
                }
                return timer;
            }) as any);
            const clear = vi.spyOn(global, "clearInterval");
            const renewLock = vi.spyOn(MapiSessionManager.prototype, "renewLock").mockResolvedValueOnce(false);
            const renewInProgress = vi.spyOn(MapiSessionManager.prototype, "renewInProgress");
            const originalSave = MapiSessionManager.prototype.save;
            let stoppedEarly = false;
            const save = vi.spyOn(MapiSessionManager.prototype, "save").mockImplementationOnce(async function (this: MapiSessionManager, session: any) {
                tick!();
                await new Promise((resolve) => setTimeout(resolve, 20));
                // Stopped as soon as the renewal saw the lock gone, before the request finished.
                stoppedEarly = clear.mock.calls.some(([timer]) => timer === renewalTimer);
                return originalSave.call(this, session);
            });
            try {
                const result = await mapiRequest(
                    server.getApplication(),
                    baseUrl,
                    { Authorization: "jwt " + ownerToken, "X-RequestType": "Execute", "Content-Type": "application/mapi-http", Cookie: cookie, "X-RequestId": "{guid}:11" },
                    body,
                );
                expect(result.headers["x-responsecode"]).toBe("0");
                expect(stoppedEarly).toBe(true);
                expect(renewLock).toHaveBeenCalledTimes(1);
                expect(renewInProgress).not.toHaveBeenCalled();
            } finally {
                interval.mockRestore();
                clear.mockRestore();
                renewLock.mockRestore();
                renewInProgress.mockRestore();
                save.mockRestore();
            }
        });

        it("Fails a request that would grow the session past its size cap with MAPI_E_TOO_BIG.", async () => {
            await createMailbox(owner.uid);
            const cookie = cookieHeaderFrom((await connect()).headers["set-cookie"]);
            const save = vi.spyOn(MapiSessionManager.prototype, "save").mockResolvedValueOnce("tooBig");
            try {
                const result = await execute(cookie, encodeRopBuffer({ ropsList: Buffer.alloc(0), handleTable: [] }));
                expect(result.headers["x-responsecode"]).toBe("0");
                expect(result.body.readUInt32LE(4)).toBe(0x80040305);
                expect(result.body.readUInt32LE(12)).toBe(0); // no RopBuffer
            } finally {
                save.mockRestore();
            }
        });
    });

    describe("RopLogon / RopRelease", () => {
        const buildRopLogonRops = function (outputHandleIndex: number): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0xfe); // RopId
            writer.writeUInt8(0); // LogonId
            writer.writeUInt8(outputHandleIndex);
            writer.writeUInt8(0x01); // LogonFlags (Private)
            writer.writeUInt32LE(0); // OpenFlags
            writer.writeUInt32LE(0); // StoreState
            writer.writeUInt16LE(0); // EssdnSize
            return writer.toBuffer();
        };

        const buildRopReleaseRops = function (inputHandleIndex: number): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0x01); // RopId
            writer.writeUInt8(0); // LogonId
            writer.writeUInt8(inputHandleIndex);
            return writer.toBuffer();
        };

        it("Logs on to a real mailbox and returns 13 well-formed, distinct FIDs.", async () => {
            const mailbox = await createMailbox(owner.uid, { displayName: "Ada Lovelace" });
            const connectResult = await connect();
            const cookie = cookieHeaderFrom(connectResult.headers["set-cookie"]);

            const inputRop = encodeRopBuffer({ ropsList: buildRopLogonRops(0), handleTable: [0xffffffff] });
            const result = await execute(cookie, inputRop);

            expect(result.status).toBe(200);
            expect(result.headers["x-responsecode"]).toBe("0");
            const reader = new BufferReader(result.body);
            reader.readUInt32LE(); // StatusCode
            reader.readUInt32LE(); // ErrorCode
            reader.readUInt32LE(); // Flags
            const ropBufferSize = reader.readUInt32LE();
            const { ropsList, handleTable } = decodeRopBuffer(reader.readBytes(ropBufferSize));
            expect(handleTable).toEqual([0xffffffff]);

            const ropsReader = new BufferReader(ropsList);
            expect(ropsReader.readUInt8()).toBe(0xfe); // RopId
            expect(ropsReader.readUInt8()).toBe(0); // OutputHandleIndex
            expect(ropsReader.readUInt32LE()).toBe(0); // ReturnValue
            expect(ropsReader.readUInt8()).toBe(0x01); // LogonFlags, echoed

            const fids: bigint[] = [];
            for (let i = 0; i < 13; i++) {
                fids.push(ropsReader.readBigUInt64LE());
            }
            expect(new Set(fids).size).toBe(13); // all 13 FIDs are distinct

            ropsReader.readUInt8(); // ResponseFlags
            expect(decodeGuid(ropsReader)).toBe(mailbox.uid); // MailboxGuid
            expect(ropsReader.hasMore()).toBe(true); // ReplId/ReplGuid/LogonTime/GwartTime/StoreState follow
            // The FID<->Folder mapping itself (Inbox/Outbox/Sent/Deleted resolving to a real Folder when one
            // exists) is verified directly against session.folderIds in RopLogonHandler's own unit tests
            // (test/mapi/rop/RopLogonHandler.test.ts) - not observable from outside this black-box HTTP test,
            // since FIDs are opaque numbers to the client until a later RopOpenFolder (a future build step).
        });

        it("Releases a logon handle with no response bytes for that ROP, matching the spec's own captured example.", async () => {
            await createMailbox(owner.uid);
            const connectResult = await connect();
            const cookie = cookieHeaderFrom(connectResult.headers["set-cookie"]);

            const logonRop = encodeRopBuffer({ ropsList: buildRopLogonRops(0), handleTable: [0xffffffff] });
            await execute(cookie, logonRop);

            const releaseRop = encodeRopBuffer({ ropsList: buildRopReleaseRops(0), handleTable: [0] });
            const result = await execute(cookie, releaseRop);

            expect(result.headers["x-responsecode"]).toBe("0");
            const reader = new BufferReader(result.body);
            reader.readUInt32LE(); // StatusCode
            reader.readUInt32LE(); // ErrorCode
            reader.readUInt32LE(); // Flags
            const ropBufferSize = reader.readUInt32LE();
            const { ropsList } = decodeRopBuffer(reader.readBytes(ropBufferSize));
            // RopRelease produces no response entry - the entire ropsList is empty.
            expect(ropsList.length).toBe(0);
        });
    });

    describe("Folder hierarchy browsing (RopOpenFolder / RopGetHierarchyTable / RopSetColumns / RopQueryRows)", () => {
        const buildLogonRops = function (outputHandleIndex: number): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0xfe);
            writer.writeUInt8(0);
            writer.writeUInt8(outputHandleIndex);
            writer.writeUInt8(0x01);
            writer.writeUInt32LE(0);
            writer.writeUInt32LE(0);
            writer.writeUInt16LE(0);
            return writer.toBuffer();
        };

        const buildOpenFolderRops = function (inputHandleIndex: number, outputHandleIndex: number, folderId: bigint): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0x02);
            writer.writeUInt8(0); // LogonId
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt8(outputHandleIndex);
            writer.writeUInt8(0); // OpenModeFlags
            writer.writeBigUInt64LE(folderId);
            return writer.toBuffer();
        };

        const buildGetHierarchyTableRops = function (inputHandleIndex: number, outputHandleIndex: number): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0x04);
            writer.writeUInt8(0); // LogonId
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt8(outputHandleIndex);
            writer.writeUInt8(0); // TableFlags - Standard
            return writer.toBuffer();
        };

        const buildSetColumnsRops = function (inputHandleIndex: number, tags: { propertyId: number; propertyType: PropertyType }[]): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0x12);
            writer.writeUInt8(0); // LogonId
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt8(0); // SetColumnsFlags
            writer.writeUInt16LE(tags.length);
            for (const tag of tags) {
                writePropertyTag(writer, tag);
            }
            return writer.toBuffer();
        };

        const buildQueryRowsRops = function (inputHandleIndex: number, rowCount: number): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0x15);
            writer.writeUInt8(0); // LogonId
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt8(0); // QueryRowsFlags
            writer.writeUInt8(1); // ForwardRead
            writer.writeUInt16LE(rowCount);
            return writer.toBuffer();
        };

        it("Browses the top-level folder list end to end: Logon, OpenFolder(root), GetHierarchyTable, SetColumns, QueryRows.", async () => {
            const mailbox = await createMailbox(owner.uid);
            await createFolder(mailbox.uid, FolderType.INBOX, "Inbox");
            // FolderType.ARCHIVE is a real folder type (restapi ^0.7) with no special-cased handling anywhere
            // in this codebase - proving it appears in generic hierarchy browsing exactly like a plain
            // FolderType.USER folder confirms no dedicated support is actually needed for it: only CALENDAR/
            // CONTACTS/TASKS get routed to a different repo (see RopGetContentsTableHandler.ts's own doc
            // comment), every other type - INBOX/SENT_ITEMS/DRAFTS/DELETED_ITEMS/OUTBOX/JUNK/ARCHIVE/NOTES/USER
            // - already falls through to the same generic Message-backed path.
            await createFolder(mailbox.uid, FolderType.ARCHIVE, "Archive");
            const connectResult = await connect();
            const cookie = cookieHeaderFrom(connectResult.headers["set-cookie"]);

            // Logon (handle 0) - discover the "root" FID.
            const logonInput = encodeRopBuffer({ ropsList: buildLogonRops(0), handleTable: [0xffffffff] });
            const logonResult = await execute(cookie, logonInput);
            const logonReader = new BufferReader(logonResult.body);
            logonReader.readUInt32LE();
            logonReader.readUInt32LE();
            logonReader.readUInt32LE();
            const logonRopBufferSize = logonReader.readUInt32LE();
            const { ropsList: logonRopsList } = decodeRopBuffer(logonReader.readBytes(logonRopBufferSize));
            const logonRopsReader = new BufferReader(logonRopsList);
            logonRopsReader.readUInt8(); // RopId
            logonRopsReader.readUInt8(); // OutputHandleIndex
            logonRopsReader.readUInt32LE(); // ReturnValue
            logonRopsReader.readUInt8(); // LogonFlags
            const rootFid = logonRopsReader.readBigUInt64LE(); // FolderIds[0] = Root

            // OpenFolder(root) at handle 1, using logon handle 0 as input.
            const openFolderInput = encodeRopBuffer({ ropsList: buildOpenFolderRops(0, 1, rootFid), handleTable: [0xffffffff, 0xffffffff] });
            const openFolderResult = await execute(cookie, openFolderInput);
            const openFolderReader = new BufferReader(openFolderResult.body);
            openFolderReader.readUInt32LE();
            openFolderReader.readUInt32LE();
            openFolderReader.readUInt32LE();
            const openFolderRopBufferSize = openFolderReader.readUInt32LE();
            const { ropsList: openFolderRopsList } = decodeRopBuffer(openFolderReader.readBytes(openFolderRopBufferSize));
            const openFolderRopsReader = new BufferReader(openFolderRopsList);
            openFolderRopsReader.readUInt8();
            openFolderRopsReader.readUInt8();
            expect(openFolderRopsReader.readUInt32LE()).toBe(0); // ReturnValue - success

            // GetHierarchyTable at handle 2, using folder handle 1 as input.
            const tableInput = encodeRopBuffer({ ropsList: buildGetHierarchyTableRops(1, 2), handleTable: [0xffffffff, 0xffffffff, 0xffffffff] });
            const tableResult = await execute(cookie, tableInput);
            const tableReader = new BufferReader(tableResult.body);
            tableReader.readUInt32LE();
            tableReader.readUInt32LE();
            tableReader.readUInt32LE();
            const tableRopBufferSize = tableReader.readUInt32LE();
            const { ropsList: tableRopsList } = decodeRopBuffer(tableReader.readBytes(tableRopBufferSize));
            const tableRopsReader = new BufferReader(tableRopsList);
            tableRopsReader.readUInt8();
            tableRopsReader.readUInt8();
            expect(tableRopsReader.readUInt32LE()).toBe(0); // ReturnValue - success

            // SetColumns on table handle 2: DisplayName + ContentCount + ContentUnreadCount.
            const columns = [
                { propertyId: 0x3001, propertyType: PropertyType.PtypString },
                { propertyId: 0x3602, propertyType: PropertyType.PtypInteger32 },
                { propertyId: 0x3603, propertyType: PropertyType.PtypInteger32 },
            ];
            const setColumnsInput = encodeRopBuffer({ ropsList: buildSetColumnsRops(2, columns), handleTable: [0xffffffff] });
            const setColumnsResult = await execute(cookie, setColumnsInput);
            expect(setColumnsResult.headers["x-responsecode"]).toBe("0");

            // QueryRows on table handle 2.
            const queryRowsInput = encodeRopBuffer({ ropsList: buildQueryRowsRops(2, 10), handleTable: [0xffffffff] });
            const queryRowsResult = await execute(cookie, queryRowsInput);
            const queryRowsReader = new BufferReader(queryRowsResult.body);
            queryRowsReader.readUInt32LE();
            queryRowsReader.readUInt32LE();
            queryRowsReader.readUInt32LE();
            const queryRowsBufferSize = queryRowsReader.readUInt32LE();
            const { ropsList: queryRowsList } = decodeRopBuffer(queryRowsReader.readBytes(queryRowsBufferSize));
            const rowsReader = new BufferReader(queryRowsList);
            rowsReader.readUInt8(); // RopId
            rowsReader.readUInt8(); // InputHandleIndex
            expect(rowsReader.readUInt32LE()).toBe(0); // ReturnValue
            rowsReader.readUInt8(); // Origin
            const rowCount = rowsReader.readUInt16LE();
            expect(rowCount).toBe(2);

            const names: string[] = [];
            for (let i = 0; i < rowCount; i++) {
                rowsReader.readUInt8(); // PropertyRow Flags
                names.push(readPropertyValue(rowsReader, PropertyType.PtypString) as string);
                readPropertyValue(rowsReader, PropertyType.PtypInteger32); // ContentCount
                readPropertyValue(rowsReader, PropertyType.PtypInteger32); // ContentUnreadCount
            }
            expect(names.sort()).toEqual(["Archive", "Inbox"]);
            expect(rowsReader.hasMore()).toBe(false);
        });
    });

    describe("Message listing (RopGetContentsTable / RopSetColumns / RopQueryRows)", () => {
        const buildLogonRops = function (outputHandleIndex: number): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0xfe);
            writer.writeUInt8(0);
            writer.writeUInt8(outputHandleIndex);
            writer.writeUInt8(0x01);
            writer.writeUInt32LE(0);
            writer.writeUInt32LE(0);
            writer.writeUInt16LE(0);
            return writer.toBuffer();
        };

        const buildOpenFolderRops = function (inputHandleIndex: number, outputHandleIndex: number, folderId: bigint): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0x02);
            writer.writeUInt8(0); // LogonId
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt8(outputHandleIndex);
            writer.writeUInt8(0); // OpenModeFlags
            writer.writeBigUInt64LE(folderId);
            return writer.toBuffer();
        };

        const buildGetContentsTableRops = function (inputHandleIndex: number, outputHandleIndex: number): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0x05);
            writer.writeUInt8(0); // LogonId
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt8(outputHandleIndex);
            writer.writeUInt8(0); // TableFlags - Standard
            return writer.toBuffer();
        };

        const buildSetColumnsRops = function (inputHandleIndex: number, tags: { propertyId: number; propertyType: PropertyType }[]): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0x12);
            writer.writeUInt8(0); // LogonId
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt8(0); // SetColumnsFlags
            writer.writeUInt16LE(tags.length);
            for (const tag of tags) {
                writePropertyTag(writer, tag);
            }
            return writer.toBuffer();
        };

        const buildQueryRowsRops = function (inputHandleIndex: number, rowCount: number): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0x15);
            writer.writeUInt8(0); // LogonId
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt8(0); // QueryRowsFlags
            writer.writeUInt8(1); // ForwardRead
            writer.writeUInt16LE(rowCount);
            return writer.toBuffer();
        };

        it("Lists a folder's messages end to end: Logon, OpenFolder(inbox), GetContentsTable, SetColumns, QueryRows.", async () => {
            const mailbox = await createMailbox(owner.uid);
            const inbox = await createFolder(mailbox.uid, FolderType.INBOX, "Inbox");
            await createMessage(mailbox.uid, inbox.uid, { subject: "Hello World" });
            await createMessage(mailbox.uid, inbox.uid, { subject: "Second Message" });
            const connectResult = await connect();
            const cookie = cookieHeaderFrom(connectResult.headers["set-cookie"]);

            // Logon (handle 0) - discover the Inbox FID.
            const logonInput = encodeRopBuffer({ ropsList: buildLogonRops(0), handleTable: [0xffffffff] });
            const logonResult = await execute(cookie, logonInput);
            const logonReader = new BufferReader(logonResult.body);
            logonReader.readUInt32LE();
            logonReader.readUInt32LE();
            logonReader.readUInt32LE();
            const logonRopBufferSize = logonReader.readUInt32LE();
            const { ropsList: logonRopsList } = decodeRopBuffer(logonReader.readBytes(logonRopBufferSize));
            const logonRopsReader = new BufferReader(logonRopsList);
            logonRopsReader.readUInt8(); // RopId
            logonRopsReader.readUInt8(); // OutputHandleIndex
            logonRopsReader.readUInt32LE(); // ReturnValue
            logonRopsReader.readUInt8(); // LogonFlags
            logonRopsReader.readBigUInt64LE(); // FolderIds[0] = Root
            logonRopsReader.readBigUInt64LE(); // FolderIds[1] = Deferred Action
            logonRopsReader.readBigUInt64LE(); // FolderIds[2] = Spooler Queue
            logonRopsReader.readBigUInt64LE(); // FolderIds[3] = IPM Subtree
            const inboxFid = logonRopsReader.readBigUInt64LE(); // FolderIds[4] = Inbox

            // OpenFolder(inbox) at handle 1, using logon handle 0 as input.
            const openFolderInput = encodeRopBuffer({ ropsList: buildOpenFolderRops(0, 1, inboxFid), handleTable: [0xffffffff, 0xffffffff] });
            const openFolderResult = await execute(cookie, openFolderInput);
            const openFolderReader = new BufferReader(openFolderResult.body);
            openFolderReader.readUInt32LE();
            openFolderReader.readUInt32LE();
            openFolderReader.readUInt32LE();
            const openFolderRopBufferSize = openFolderReader.readUInt32LE();
            const { ropsList: openFolderRopsList } = decodeRopBuffer(openFolderReader.readBytes(openFolderRopBufferSize));
            const openFolderRopsReader = new BufferReader(openFolderRopsList);
            openFolderRopsReader.readUInt8();
            openFolderRopsReader.readUInt8();
            expect(openFolderRopsReader.readUInt32LE()).toBe(0); // ReturnValue - success

            // GetContentsTable at handle 2, using folder handle 1 as input.
            const tableInput = encodeRopBuffer({ ropsList: buildGetContentsTableRops(1, 2), handleTable: [0xffffffff, 0xffffffff, 0xffffffff] });
            const tableResult = await execute(cookie, tableInput);
            const tableReader = new BufferReader(tableResult.body);
            tableReader.readUInt32LE();
            tableReader.readUInt32LE();
            tableReader.readUInt32LE();
            const tableRopBufferSize = tableReader.readUInt32LE();
            const { ropsList: tableRopsList } = decodeRopBuffer(tableReader.readBytes(tableRopBufferSize));
            const tableRopsReader = new BufferReader(tableRopsList);
            tableRopsReader.readUInt8();
            tableRopsReader.readUInt8();
            expect(tableRopsReader.readUInt32LE()).toBe(0); // ReturnValue - success

            // SetColumns on table handle 2: Subject.
            const columns = [{ propertyId: 0x0037, propertyType: PropertyType.PtypString }];
            const setColumnsInput = encodeRopBuffer({ ropsList: buildSetColumnsRops(2, columns), handleTable: [0xffffffff] });
            const setColumnsResult = await execute(cookie, setColumnsInput);
            expect(setColumnsResult.headers["x-responsecode"]).toBe("0");

            // QueryRows on table handle 2.
            const queryRowsInput = encodeRopBuffer({ ropsList: buildQueryRowsRops(2, 10), handleTable: [0xffffffff] });
            const queryRowsResult = await execute(cookie, queryRowsInput);
            const queryRowsReader = new BufferReader(queryRowsResult.body);
            queryRowsReader.readUInt32LE();
            queryRowsReader.readUInt32LE();
            queryRowsReader.readUInt32LE();
            const queryRowsBufferSize = queryRowsReader.readUInt32LE();
            const { ropsList: queryRowsList } = decodeRopBuffer(queryRowsReader.readBytes(queryRowsBufferSize));
            const rowsReader = new BufferReader(queryRowsList);
            rowsReader.readUInt8(); // RopId
            rowsReader.readUInt8(); // InputHandleIndex
            expect(rowsReader.readUInt32LE()).toBe(0); // ReturnValue
            rowsReader.readUInt8(); // Origin
            const rowCount = rowsReader.readUInt16LE();
            expect(rowCount).toBe(2);

            const subjects: string[] = [];
            for (let i = 0; i < rowCount; i++) {
                rowsReader.readUInt8(); // PropertyRow Flags
                subjects.push(readPropertyValue(rowsReader, PropertyType.PtypString) as string);
            }
            expect(subjects.sort()).toEqual(["Hello World", "Second Message"]);
            expect(rowsReader.hasMore()).toBe(false);
        });
    });

    describe("Contacts/Tasks folder browsing (RopGetContentsTable over CONTACTS/TASKS folders)", () => {
        const buildLogonRops = function (outputHandleIndex: number): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0xfe);
            writer.writeUInt8(0);
            writer.writeUInt8(outputHandleIndex);
            writer.writeUInt8(0x01);
            writer.writeUInt32LE(0);
            writer.writeUInt32LE(0);
            writer.writeUInt16LE(0);
            return writer.toBuffer();
        };

        const buildOpenFolderRops = function (inputHandleIndex: number, outputHandleIndex: number, folderId: bigint): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0x02);
            writer.writeUInt8(0); // LogonId
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt8(outputHandleIndex);
            writer.writeUInt8(0); // OpenModeFlags
            writer.writeBigUInt64LE(folderId);
            return writer.toBuffer();
        };

        const buildGetTableRops = function (ropId: number, inputHandleIndex: number, outputHandleIndex: number): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(ropId);
            writer.writeUInt8(0); // LogonId
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt8(outputHandleIndex);
            writer.writeUInt8(0); // TableFlags - Standard
            return writer.toBuffer();
        };

        const buildSetColumnsRops = function (inputHandleIndex: number, tags: { propertyId: number; propertyType: PropertyType }[]): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0x12);
            writer.writeUInt8(0); // LogonId
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt8(0); // SetColumnsFlags
            writer.writeUInt16LE(tags.length);
            for (const tag of tags) {
                writePropertyTag(writer, tag);
            }
            return writer.toBuffer();
        };

        const buildQueryRowsRops = function (inputHandleIndex: number, rowCount: number): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0x15);
            writer.writeUInt8(0); // LogonId
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt8(0); // QueryRowsFlags
            writer.writeUInt8(1); // ForwardRead
            writer.writeUInt16LE(rowCount);
            return writer.toBuffer();
        };

        /** Discovers `wellKnownFolderName`'s real FID via the root folder's own hierarchy table - Contacts/
         * Tasks aren't part of RopLogon's fixed 13-folder `FolderIds` list (see RopLogonHandler's own doc
         * comment), so a real client (and this test) finds them the same way it finds any other top-level
         * folder: OpenFolder(root) -> GetHierarchyTable -> SetColumns([DisplayName, FolderId]) -> QueryRows. */
        async function discoverFolderId(cookie: string, wellKnownFolderName: string): Promise<bigint> {
            const logonInput = encodeRopBuffer({ ropsList: buildLogonRops(0), handleTable: [0xffffffff] });
            const logonResult = await execute(cookie, logonInput);
            const logonReader = new BufferReader(logonResult.body);
            logonReader.readUInt32LE();
            logonReader.readUInt32LE();
            logonReader.readUInt32LE();
            const logonRopBufferSize = logonReader.readUInt32LE();
            const { ropsList: logonRopsList } = decodeRopBuffer(logonReader.readBytes(logonRopBufferSize));
            const logonRopsReader = new BufferReader(logonRopsList);
            logonRopsReader.readUInt8();
            logonRopsReader.readUInt8();
            logonRopsReader.readUInt32LE();
            logonRopsReader.readUInt8();
            const rootFid = logonRopsReader.readBigUInt64LE();

            const openFolderInput = encodeRopBuffer({ ropsList: buildOpenFolderRops(0, 1, rootFid), handleTable: [0xffffffff, 0xffffffff] });
            await execute(cookie, openFolderInput);

            const tableInput = encodeRopBuffer({ ropsList: buildGetTableRops(0x04, 1, 2), handleTable: [0xffffffff, 0xffffffff, 0xffffffff] });
            await execute(cookie, tableInput);

            const columns = [
                { propertyId: 0x3001, propertyType: PropertyType.PtypString }, // DisplayName
                { propertyId: 0x6748, propertyType: PropertyType.PtypInteger64 }, // FolderId
            ];
            const setColumnsInput = encodeRopBuffer({ ropsList: buildSetColumnsRops(2, columns), handleTable: [0xffffffff] });
            await execute(cookie, setColumnsInput);

            const queryRowsInput = encodeRopBuffer({ ropsList: buildQueryRowsRops(2, 10), handleTable: [0xffffffff] });
            const queryRowsResult = await execute(cookie, queryRowsInput);
            const queryRowsReader = new BufferReader(queryRowsResult.body);
            queryRowsReader.readUInt32LE();
            queryRowsReader.readUInt32LE();
            queryRowsReader.readUInt32LE();
            const queryRowsBufferSize = queryRowsReader.readUInt32LE();
            const { ropsList: queryRowsList } = decodeRopBuffer(queryRowsReader.readBytes(queryRowsBufferSize));
            const rowsReader = new BufferReader(queryRowsList);
            rowsReader.readUInt8();
            rowsReader.readUInt8();
            rowsReader.readUInt32LE();
            rowsReader.readUInt8();
            const rowCount = rowsReader.readUInt16LE();
            for (let i = 0; i < rowCount; i++) {
                rowsReader.readUInt8(); // PropertyRow Flags
                const name = readPropertyValue(rowsReader, PropertyType.PtypString) as string;
                const fid = readPropertyValue(rowsReader, PropertyType.PtypInteger64) as bigint;
                if (name === wellKnownFolderName) {
                    return fid;
                }
            }
            throw new Error(`Folder named "${wellKnownFolderName}" was not found in the root hierarchy table.`);
        }

        it("Lists a CONTACTS folder's contacts end to end, resolved by displayName via ContactTarget.", async () => {
            const mailbox = await createMailbox(owner.uid);
            const contactsFolder = await createFolder(mailbox.uid, FolderType.CONTACTS, "Contacts");
            await createContact(mailbox.uid, contactsFolder.uid, { displayName: "Jane Doe" });
            await createContact(mailbox.uid, contactsFolder.uid, { displayName: "John Smith" });
            const connectResult = await connect();
            const cookie = cookieHeaderFrom(connectResult.headers["set-cookie"]);

            const contactsFid = await discoverFolderId(cookie, "Contacts");

            const openFolderInput = encodeRopBuffer({ ropsList: buildOpenFolderRops(0, 3, contactsFid), handleTable: [0xffffffff, 0xffffffff] });
            const openFolderResult = await execute(cookie, openFolderInput);
            const openFolderReader = new BufferReader(openFolderResult.body);
            openFolderReader.readUInt32LE();
            openFolderReader.readUInt32LE();
            openFolderReader.readUInt32LE();
            const openFolderRopBufferSize = openFolderReader.readUInt32LE();
            const { ropsList: openFolderRopsList } = decodeRopBuffer(openFolderReader.readBytes(openFolderRopBufferSize));
            const openFolderRopsReader = new BufferReader(openFolderRopsList);
            openFolderRopsReader.readUInt8();
            openFolderRopsReader.readUInt8();
            expect(openFolderRopsReader.readUInt32LE()).toBe(0); // ReturnValue - success

            const tableInput = encodeRopBuffer({ ropsList: buildGetTableRops(0x05, 3, 4), handleTable: [0xffffffff, 0xffffffff, 0xffffffff] });
            const tableResult = await execute(cookie, tableInput);
            const tableReader = new BufferReader(tableResult.body);
            tableReader.readUInt32LE();
            tableReader.readUInt32LE();
            tableReader.readUInt32LE();
            const tableRopBufferSize = tableReader.readUInt32LE();
            const { ropsList: tableRopsList } = decodeRopBuffer(tableReader.readBytes(tableRopBufferSize));
            const tableRopsReader = new BufferReader(tableRopsList);
            tableRopsReader.readUInt8();
            tableRopsReader.readUInt8();
            expect(tableRopsReader.readUInt32LE()).toBe(0); // ReturnValue - success

            const columns = [{ propertyId: 0x0037, propertyType: PropertyType.PtypString }]; // PidTagSubject -> displayName
            const setColumnsInput = encodeRopBuffer({ ropsList: buildSetColumnsRops(4, columns), handleTable: [0xffffffff] });
            await execute(cookie, setColumnsInput);

            const queryRowsInput = encodeRopBuffer({ ropsList: buildQueryRowsRops(4, 10), handleTable: [0xffffffff] });
            const queryRowsResult = await execute(cookie, queryRowsInput);
            const queryRowsReader = new BufferReader(queryRowsResult.body);
            queryRowsReader.readUInt32LE();
            queryRowsReader.readUInt32LE();
            queryRowsReader.readUInt32LE();
            const queryRowsBufferSize = queryRowsReader.readUInt32LE();
            const { ropsList: queryRowsList } = decodeRopBuffer(queryRowsReader.readBytes(queryRowsBufferSize));
            const rowsReader = new BufferReader(queryRowsList);
            rowsReader.readUInt8();
            rowsReader.readUInt8();
            expect(rowsReader.readUInt32LE()).toBe(0);
            rowsReader.readUInt8();
            const rowCount = rowsReader.readUInt16LE();
            expect(rowCount).toBe(2);

            const names: string[] = [];
            for (let i = 0; i < rowCount; i++) {
                rowsReader.readUInt8();
                names.push(readPropertyValue(rowsReader, PropertyType.PtypString) as string);
            }
            expect(names.sort()).toEqual(["Jane Doe", "John Smith"]);
        });

        it("Lists a TASKS folder's tasks end to end, resolved by title via TaskTarget.", async () => {
            const mailbox = await createMailbox(owner.uid);
            const tasksFolder = await createFolder(mailbox.uid, FolderType.TASKS, "Tasks");
            await createTask(mailbox.uid, tasksFolder.uid, { title: "Ship it" });
            await createTask(mailbox.uid, tasksFolder.uid, { title: "Write docs" });
            const connectResult = await connect();
            const cookie = cookieHeaderFrom(connectResult.headers["set-cookie"]);

            const tasksFid = await discoverFolderId(cookie, "Tasks");

            const openFolderInput = encodeRopBuffer({ ropsList: buildOpenFolderRops(0, 3, tasksFid), handleTable: [0xffffffff, 0xffffffff] });
            await execute(cookie, openFolderInput);

            const tableInput = encodeRopBuffer({ ropsList: buildGetTableRops(0x05, 3, 4), handleTable: [0xffffffff, 0xffffffff, 0xffffffff] });
            await execute(cookie, tableInput);

            const columns = [{ propertyId: 0x0037, propertyType: PropertyType.PtypString }]; // PidTagSubject -> title
            const setColumnsInput = encodeRopBuffer({ ropsList: buildSetColumnsRops(4, columns), handleTable: [0xffffffff] });
            await execute(cookie, setColumnsInput);

            const queryRowsInput = encodeRopBuffer({ ropsList: buildQueryRowsRops(4, 10), handleTable: [0xffffffff] });
            const queryRowsResult = await execute(cookie, queryRowsInput);
            const queryRowsReader = new BufferReader(queryRowsResult.body);
            queryRowsReader.readUInt32LE();
            queryRowsReader.readUInt32LE();
            queryRowsReader.readUInt32LE();
            const queryRowsBufferSize = queryRowsReader.readUInt32LE();
            const { ropsList: queryRowsList } = decodeRopBuffer(queryRowsReader.readBytes(queryRowsBufferSize));
            const rowsReader = new BufferReader(queryRowsList);
            rowsReader.readUInt8();
            rowsReader.readUInt8();
            expect(rowsReader.readUInt32LE()).toBe(0);
            rowsReader.readUInt8();
            const rowCount = rowsReader.readUInt16LE();
            expect(rowCount).toBe(2);

            const titles: string[] = [];
            for (let i = 0; i < rowCount; i++) {
                rowsReader.readUInt8();
                titles.push(readPropertyValue(rowsReader, PropertyType.PtypString) as string);
            }
            expect(titles.sort()).toEqual(["Ship it", "Write docs"]);
        });
    });

    describe("Message reading (RopOpenMessage / RopGetPropertiesSpecific / RopOpenStream / RopReadStream)", () => {
        const buildLogonRops = function (outputHandleIndex: number): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0xfe);
            writer.writeUInt8(0);
            writer.writeUInt8(outputHandleIndex);
            writer.writeUInt8(0x01);
            writer.writeUInt32LE(0);
            writer.writeUInt32LE(0);
            writer.writeUInt16LE(0);
            return writer.toBuffer();
        };

        const buildOpenFolderRops = function (inputHandleIndex: number, outputHandleIndex: number, folderId: bigint): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0x02);
            writer.writeUInt8(0);
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt8(outputHandleIndex);
            writer.writeUInt8(0);
            writer.writeBigUInt64LE(folderId);
            return writer.toBuffer();
        };

        const buildGetContentsTableRops = function (inputHandleIndex: number, outputHandleIndex: number): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0x05);
            writer.writeUInt8(0);
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt8(outputHandleIndex);
            writer.writeUInt8(0);
            return writer.toBuffer();
        };

        const buildSetColumnsRops = function (inputHandleIndex: number, tags: { propertyId: number; propertyType: PropertyType }[]): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0x12);
            writer.writeUInt8(0);
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt8(0);
            writer.writeUInt16LE(tags.length);
            for (const tag of tags) {
                writePropertyTag(writer, tag);
            }
            return writer.toBuffer();
        };

        const buildQueryRowsRops = function (inputHandleIndex: number, rowCount: number): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0x15);
            writer.writeUInt8(0);
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt8(0);
            writer.writeUInt8(1);
            writer.writeUInt16LE(rowCount);
            return writer.toBuffer();
        };

        const buildOpenMessageRops = function (
            inputHandleIndex: number,
            outputHandleIndex: number,
            folderId: bigint,
            messageId: bigint,
        ): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0x03);
            writer.writeUInt8(0); // LogonId
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt8(outputHandleIndex);
            writer.writeUInt16LE(0); // CodePageId
            writer.writeBigUInt64LE(folderId);
            writer.writeUInt8(0); // OpenModeFlags
            writer.writeBigUInt64LE(messageId);
            return writer.toBuffer();
        };

        const buildGetPropertiesSpecificRops = function (
            inputHandleIndex: number,
            tags: { propertyId: number; propertyType: PropertyType }[],
        ): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0x07);
            writer.writeUInt8(0); // LogonId
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt16LE(0); // PropertySizeLimit
            writer.writeUInt16LE(0); // WantUnicode
            writer.writeUInt16LE(tags.length);
            for (const tag of tags) {
                writePropertyTag(writer, tag);
            }
            return writer.toBuffer();
        };

        const buildOpenStreamRops = function (
            inputHandleIndex: number,
            outputHandleIndex: number,
            propertyTag: { propertyId: number; propertyType: PropertyType },
        ): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0x2b);
            writer.writeUInt8(0); // LogonId
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt8(outputHandleIndex);
            writePropertyTag(writer, propertyTag);
            writer.writeUInt8(0); // OpenModeFlags
            return writer.toBuffer();
        };

        const buildReadStreamRops = function (inputHandleIndex: number, byteCount: number): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0x2c);
            writer.writeUInt8(0); // LogonId
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt16LE(byteCount);
            return writer.toBuffer();
        };

        it("Reads a real message end to end: OpenMessage, GetPropertiesSpecific(Subject), OpenStream+ReadStream(Body).", async () => {
            const mailbox = await createMailbox(owner.uid);
            const inbox = await createFolder(mailbox.uid, FolderType.INBOX, "Inbox");
            const message = await createMessage(mailbox.uid, inbox.uid, { subject: "Read Me" });
            await blobStore().put(
                message.bodyBlobKey,
                Buffer.from("From: sender@example.com\r\nTo: owner@example.com\r\nSubject: Read Me\r\n\r\nActual body text."),
            );
            const connectResult = await connect();
            const cookie = cookieHeaderFrom(connectResult.headers["set-cookie"]);

            // Logon (handle 0) - discover the Inbox FID.
            const logonResult = await execute(cookie, encodeRopBuffer({ ropsList: buildLogonRops(0), handleTable: [0xffffffff] }));
            const logonReader = new BufferReader(logonResult.body);
            logonReader.readUInt32LE();
            logonReader.readUInt32LE();
            logonReader.readUInt32LE();
            const { ropsList: logonRopsList } = decodeRopBuffer(logonReader.readBytes(logonReader.readUInt32LE()));
            const logonRopsReader = new BufferReader(logonRopsList);
            logonRopsReader.readUInt8();
            logonRopsReader.readUInt8();
            logonRopsReader.readUInt32LE();
            logonRopsReader.readUInt8();
            logonRopsReader.readBigUInt64LE(); // Root
            logonRopsReader.readBigUInt64LE(); // Deferred Action
            logonRopsReader.readBigUInt64LE(); // Spooler Queue
            logonRopsReader.readBigUInt64LE(); // IPM Subtree
            const inboxFid = logonRopsReader.readBigUInt64LE(); // Inbox

            // OpenFolder(inbox) at handle 1.
            await execute(cookie, encodeRopBuffer({ ropsList: buildOpenFolderRops(0, 1, inboxFid), handleTable: [0xffffffff, 0xffffffff] }));

            // GetContentsTable at handle 2, SetColumns(Subject, Mid), QueryRows - to learn the message's MID.
            await execute(cookie, encodeRopBuffer({ ropsList: buildGetContentsTableRops(1, 2), handleTable: [0xffffffff, 0xffffffff, 0xffffffff] }));
            const columns = [
                { propertyId: 0x0037, propertyType: PropertyType.PtypString },
                { propertyId: 0x674a, propertyType: PropertyType.PtypInteger64 },
            ];
            await execute(cookie, encodeRopBuffer({ ropsList: buildSetColumnsRops(2, columns), handleTable: [0xffffffff] }));
            const queryRowsResult = await execute(cookie, encodeRopBuffer({ ropsList: buildQueryRowsRops(2, 10), handleTable: [0xffffffff] }));
            const queryRowsReader = new BufferReader(queryRowsResult.body);
            queryRowsReader.readUInt32LE();
            queryRowsReader.readUInt32LE();
            queryRowsReader.readUInt32LE();
            const { ropsList: queryRowsList } = decodeRopBuffer(queryRowsReader.readBytes(queryRowsReader.readUInt32LE()));
            const rowsReader = new BufferReader(queryRowsList);
            rowsReader.readUInt8();
            rowsReader.readUInt8();
            rowsReader.readUInt32LE();
            rowsReader.readUInt8();
            expect(rowsReader.readUInt16LE()).toBe(1); // RowCount
            rowsReader.readUInt8(); // PropertyRow Flags
            expect(readPropertyValue(rowsReader, PropertyType.PtypString)).toBe("Read Me");
            const mid = readPropertyValue(rowsReader, PropertyType.PtypInteger64) as bigint;

            // OpenMessage at handle 3, using the MID just learned.
            const openMessageResult = await execute(
                cookie,
                encodeRopBuffer({ ropsList: buildOpenMessageRops(1, 3, inboxFid, mid), handleTable: [0xffffffff, 0xffffffff] }),
            );
            const openMessageReader = new BufferReader(openMessageResult.body);
            openMessageReader.readUInt32LE();
            openMessageReader.readUInt32LE();
            openMessageReader.readUInt32LE();
            const { ropsList: openMessageRopsList } = decodeRopBuffer(openMessageReader.readBytes(openMessageReader.readUInt32LE()));
            const openMessageRopsReader = new BufferReader(openMessageRopsList);
            openMessageRopsReader.readUInt8();
            openMessageRopsReader.readUInt8();
            expect(openMessageRopsReader.readUInt32LE()).toBe(0); // ReturnValue - success

            // GetPropertiesSpecific(Subject) on message handle 3.
            const getPropsInput = encodeRopBuffer({
                ropsList: buildGetPropertiesSpecificRops(3, [{ propertyId: 0x0037, propertyType: PropertyType.PtypString }]),
                handleTable: [0xffffffff],
            });
            const getPropsResult = await execute(cookie, getPropsInput);
            const getPropsReader = new BufferReader(getPropsResult.body);
            getPropsReader.readUInt32LE();
            getPropsReader.readUInt32LE();
            getPropsReader.readUInt32LE();
            const { ropsList: getPropsRopsList } = decodeRopBuffer(getPropsReader.readBytes(getPropsReader.readUInt32LE()));
            const getPropsRopsReader = new BufferReader(getPropsRopsList);
            getPropsRopsReader.readUInt8();
            getPropsRopsReader.readUInt8();
            expect(getPropsRopsReader.readUInt32LE()).toBe(0); // ReturnValue
            getPropsRopsReader.readUInt8(); // PropertyRow Flags
            expect(readPropertyValue(getPropsRopsReader, PropertyType.PtypString)).toBe("Read Me");

            // OpenStream(PidTagBody) at handle 4, then ReadStream to fetch the actual body text.
            const openStreamInput = encodeRopBuffer({
                ropsList: buildOpenStreamRops(3, 4, { propertyId: 0x1000, propertyType: PropertyType.PtypString }),
                handleTable: [0xffffffff, 0xffffffff],
            });
            const openStreamResult = await execute(cookie, openStreamInput);
            const openStreamReader = new BufferReader(openStreamResult.body);
            openStreamReader.readUInt32LE();
            openStreamReader.readUInt32LE();
            openStreamReader.readUInt32LE();
            const { ropsList: openStreamRopsList } = decodeRopBuffer(openStreamReader.readBytes(openStreamReader.readUInt32LE()));
            const openStreamRopsReader = new BufferReader(openStreamRopsList);
            openStreamRopsReader.readUInt8();
            openStreamRopsReader.readUInt8();
            expect(openStreamRopsReader.readUInt32LE()).toBe(0); // ReturnValue
            const streamSize = openStreamRopsReader.readUInt32LE();
            expect(streamSize).toBeGreaterThan(0);

            const readStreamInput = encodeRopBuffer({ ropsList: buildReadStreamRops(4, 1000), handleTable: [0xffffffff] });
            const readStreamResult = await execute(cookie, readStreamInput);
            const readStreamReader = new BufferReader(readStreamResult.body);
            readStreamReader.readUInt32LE();
            readStreamReader.readUInt32LE();
            readStreamReader.readUInt32LE();
            const { ropsList: readStreamRopsList } = decodeRopBuffer(readStreamReader.readBytes(readStreamReader.readUInt32LE()));
            const readStreamRopsReader = new BufferReader(readStreamRopsList);
            readStreamRopsReader.readUInt8();
            readStreamRopsReader.readUInt8();
            expect(readStreamRopsReader.readUInt32LE()).toBe(0); // ReturnValue
            const dataSize = readStreamRopsReader.readUInt16LE();
            expect(dataSize).toBe(streamSize);
            expect(readPropertyValue(readStreamRopsReader, PropertyType.PtypString)).toBe("Actual body text.");
        });

        it("Deletes a message by MID with RopDeleteMessages as a soft delete, recording a MESSAGE_DELETE audit entry.", async () => {
            const mailbox = await createMailbox(owner.uid);
            const inbox = await createFolder(mailbox.uid, FolderType.INBOX, "Inbox");
            const message = await createMessage(mailbox.uid, inbox.uid, { subject: "Delete Me" });
            const connectResult = await connect();
            const cookie = cookieHeaderFrom(connectResult.headers["set-cookie"]);

            const logonResult = await execute(cookie, encodeRopBuffer({ ropsList: buildLogonRops(0), handleTable: [0xffffffff] }));
            const logonReader = new BufferReader(logonResult.body);
            logonReader.readBytes(12);
            const logonRops = new BufferReader(decodeRopBuffer(logonReader.readBytes(logonReader.readUInt32LE())).ropsList);
            logonRops.readBytes(7);
            logonRops.readBytes(8 * 4); // Root, Deferred Action, Spooler Queue, IPM Subtree
            const inboxFid = logonRops.readBigUInt64LE();

            await execute(cookie, encodeRopBuffer({ ropsList: buildOpenFolderRops(0, 1, inboxFid), handleTable: [0xffffffff, 0xffffffff] }));
            await execute(cookie, encodeRopBuffer({ ropsList: buildGetContentsTableRops(1, 2), handleTable: [0xffffffff, 0xffffffff, 0xffffffff] }));
            await execute(cookie, encodeRopBuffer({ ropsList: buildSetColumnsRops(2, [{ propertyId: 0x674a, propertyType: PropertyType.PtypInteger64 }]), handleTable: [0xffffffff] }));
            const rowsResult = await execute(cookie, encodeRopBuffer({ ropsList: buildQueryRowsRops(2, 10), handleTable: [0xffffffff] }));
            const rowsReader = new BufferReader(rowsResult.body);
            rowsReader.readBytes(12);
            const rows = new BufferReader(decodeRopBuffer(rowsReader.readBytes(rowsReader.readUInt32LE())).ropsList);
            rows.readBytes(7);
            expect(rows.readUInt16LE()).toBe(1);
            rows.readUInt8();
            const mid = readPropertyValue(rows, PropertyType.PtypInteger64) as bigint;

            const deleteRops = new BufferWriter().writeUInt8(0x1e).writeUInt8(0).writeUInt8(1).writeUInt8(0).writeUInt8(0).writeUInt16LE(1).writeBigUInt64LE(mid).toBuffer();
            const deleteResult = await execute(cookie, encodeRopBuffer({ ropsList: deleteRops, handleTable: [0xffffffff] }));
            const deleteReader = new BufferReader(deleteResult.body);
            deleteReader.readBytes(12);
            const deleteResponse = new BufferReader(decodeRopBuffer(deleteReader.readBytes(deleteReader.readUInt32LE())).ropsList);
            deleteResponse.readBytes(2);
            expect(deleteResponse.readUInt32LE()).toBe(0);
            expect(deleteResponse.readUInt8()).toBe(0); // PartialCompletion

            const stored: any = await messageRepo.findOne({ uid: message.uid } as any);
            expect(stored?.deleted).toBe(true); // soft-deleted, still recoverable
            const connMgr: ConnectionManager | undefined = objectFactory.getInstance(ConnectionManager);
            const auditRepo = (connMgr?.connections.get("mongo") as MongoConnection).getMongoRepository("AuditLogEntryMongo");
            const entries: any[] = await auditRepo.find({ targetUid: message.uid }).toArray();
            expect(entries).toHaveLength(1);
            expect(entries[0]).toMatchObject({ action: "message.delete", targetType: "Message", mailboxUid: mailbox.uid, actorUserUid: owner.uid });
        });

        it("Refreshes the source folder's stored unreadCount/totalCount after RopDeleteMessages, and publishes a Folder update event.", async () => {
            const mailbox = await createMailbox(owner.uid);
            const inbox = await createFolder(mailbox.uid, FolderType.INBOX, "Inbox");
            const unreadFlags = { read: false, flagged: false, answered: false, forwarded: false };
            await createMessage(mailbox.uid, inbox.uid, { subject: "Keep Unread", flags: unreadFlags });
            await createMessage(mailbox.uid, inbox.uid, { subject: "Delete Me Too", flags: unreadFlags });
            const connectResult = await connect();
            const cookie = cookieHeaderFrom(connectResult.headers["set-cookie"]);

            const logonResult = await execute(cookie, encodeRopBuffer({ ropsList: buildLogonRops(0), handleTable: [0xffffffff] }));
            const logonReader = new BufferReader(logonResult.body);
            logonReader.readBytes(12);
            const logonRops = new BufferReader(decodeRopBuffer(logonReader.readBytes(logonReader.readUInt32LE())).ropsList);
            logonRops.readBytes(7);
            logonRops.readBytes(8 * 4); // Root, Deferred Action, Spooler Queue, IPM Subtree
            const inboxFid = logonRops.readBigUInt64LE();

            await execute(cookie, encodeRopBuffer({ ropsList: buildOpenFolderRops(0, 1, inboxFid), handleTable: [0xffffffff, 0xffffffff] }));
            await execute(cookie, encodeRopBuffer({ ropsList: buildGetContentsTableRops(1, 2), handleTable: [0xffffffff, 0xffffffff, 0xffffffff] }));
            await execute(cookie, encodeRopBuffer({ ropsList: buildSetColumnsRops(2, [{ propertyId: 0x674a, propertyType: PropertyType.PtypInteger64 }]), handleTable: [0xffffffff] }));
            const rowsResult = await execute(cookie, encodeRopBuffer({ ropsList: buildQueryRowsRops(2, 10), handleTable: [0xffffffff] }));
            const rowsReader = new BufferReader(rowsResult.body);
            rowsReader.readBytes(12);
            const rows = new BufferReader(decodeRopBuffer(rowsReader.readBytes(rowsReader.readUInt32LE())).ropsList);
            rows.readBytes(7);
            expect(rows.readUInt16LE()).toBe(2); // both messages
            rows.readUInt8();
            const firstMid = readPropertyValue(rows, PropertyType.PtypInteger64) as bigint;

            // Only one of the two unread messages is deleted - the folder still holds the other, still unread.
            const deleteRops = new BufferWriter().writeUInt8(0x1e).writeUInt8(0).writeUInt8(1).writeUInt8(0).writeUInt8(0).writeUInt16LE(1).writeBigUInt64LE(firstMid).toBuffer();
            const deleteResult = await execute(cookie, encodeRopBuffer({ ropsList: deleteRops, handleTable: [0xffffffff] }));
            const deleteReader = new BufferReader(deleteResult.body);
            deleteReader.readBytes(12);
            const deleteResponse = new BufferReader(decodeRopBuffer(deleteReader.readBytes(deleteReader.readUInt32LE())).ropsList);
            deleteResponse.readBytes(2);
            expect(deleteResponse.readUInt32LE()).toBe(0);
            expect(deleteResponse.readUInt8()).toBe(0); // PartialCompletion

            // Without RopDeleteMessagesHandler refreshing the folder's cached counts, these would stay at
            // createFolder()'s placeholder (0/0) forever - never reflecting either the 2 messages filed or the 1 left.
            const stored: any = await folderRepo.findOne({ uid: inbox.uid } as any);
            expect(stored?.totalCount).toBe(1);
            expect(stored?.unreadCount).toBe(1);
        });
    });

    describe("Message categories (PidNameKeywords / Outlook Categories, backed by Message.labelUids)", () => {
        const PSETID_PUBLIC_STRINGS = "00020329-0000-0000-c000-000000000046";

        const buildLogonRops = function (outputHandleIndex: number): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0xfe);
            writer.writeUInt8(0);
            writer.writeUInt8(outputHandleIndex);
            writer.writeUInt8(0x01);
            writer.writeUInt32LE(0);
            writer.writeUInt32LE(0);
            writer.writeUInt16LE(0);
            return writer.toBuffer();
        };

        const buildOpenFolderRops = function (inputHandleIndex: number, outputHandleIndex: number, folderId: bigint): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0x02);
            writer.writeUInt8(0);
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt8(outputHandleIndex);
            writer.writeUInt8(0);
            writer.writeBigUInt64LE(folderId);
            return writer.toBuffer();
        };

        const buildGetContentsTableRops = function (inputHandleIndex: number, outputHandleIndex: number): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0x05);
            writer.writeUInt8(0);
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt8(outputHandleIndex);
            writer.writeUInt8(0);
            return writer.toBuffer();
        };

        const buildSetColumnsRops = function (inputHandleIndex: number, tags: { propertyId: number; propertyType: PropertyType }[]): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0x12);
            writer.writeUInt8(0);
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt8(0);
            writer.writeUInt16LE(tags.length);
            for (const tag of tags) {
                writePropertyTag(writer, tag);
            }
            return writer.toBuffer();
        };

        const buildQueryRowsRops = function (inputHandleIndex: number, rowCount: number): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0x15);
            writer.writeUInt8(0);
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt8(0);
            writer.writeUInt8(1);
            writer.writeUInt16LE(rowCount);
            return writer.toBuffer();
        };

        const buildOpenMessageRops = function (inputHandleIndex: number, outputHandleIndex: number, folderId: bigint, messageId: bigint): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0x03);
            writer.writeUInt8(0);
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt8(outputHandleIndex);
            writer.writeUInt16LE(0);
            writer.writeBigUInt64LE(folderId);
            writer.writeUInt8(0);
            writer.writeBigUInt64LE(messageId);
            return writer.toBuffer();
        };

        /** `RopGetPropertyIdsFromNames`'s `Kind=Name` variant (`0x01`) - unlike the existing
         * `RopGetPropertyIdsFromNames` describe block's own helper (`Kind=LID` only), `PidNameKeywords` is
         * specifically a named-by-string property (see `PropertyResolvers.ts`'s own doc comment). `NameSize`
         * is inclusive of the trailing UTF-16LE null terminator, per `RopGetPropertyIdsFromNamesHandler
         * .readPropertyName`'s own documented convention. */
        const buildGetPropertyIdsFromNamesForNameRops = function (inputHandleIndex: number, guid: string, name: string): Buffer {
            const nameBytes = Buffer.concat([Buffer.from(name, "utf16le"), Buffer.from([0x00, 0x00])]);
            const writer = new BufferWriter();
            writer.writeUInt8(0x56);
            writer.writeUInt8(0);
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt8(0x02); // Flags - assign new IDs
            writer.writeUInt16LE(1); // PropertyNameCount
            writer.writeUInt8(0x01); // Kind - Name
            writer.writeBytes(encodeGuid(guid));
            writer.writeUInt8(nameBytes.length);
            writer.writeBytes(nameBytes);
            return writer.toBuffer();
        };

        const buildGetPropertiesSpecificRops = function (inputHandleIndex: number, tags: { propertyId: number; propertyType: PropertyType }[]): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0x07);
            writer.writeUInt8(0);
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt16LE(0);
            writer.writeUInt16LE(0);
            writer.writeUInt16LE(tags.length);
            for (const tag of tags) {
                writePropertyTag(writer, tag);
            }
            return writer.toBuffer();
        };

        it("Resolves a real message's labelUids to display names via PidNameKeywords, end to end.", async () => {
            const mailbox = await createMailbox(owner.uid);
            const inbox = await createFolder(mailbox.uid, FolderType.INBOX, "Inbox");
            const important = await createLabel(mailbox.uid, { name: "Important" });
            const work = await createLabel(mailbox.uid, { name: "Work" });
            await createMessage(mailbox.uid, inbox.uid, { subject: "Categorized", labelUids: [important.uid, work.uid] });
            const connectResult = await connect();
            const cookie = cookieHeaderFrom(connectResult.headers["set-cookie"]);

            // Logon -> discover the Inbox FID.
            const logonResult = await execute(cookie, encodeRopBuffer({ ropsList: buildLogonRops(0), handleTable: [0xffffffff] }));
            const logonReader = new BufferReader(logonResult.body);
            logonReader.readUInt32LE();
            logonReader.readUInt32LE();
            logonReader.readUInt32LE();
            const { ropsList: logonRopsList } = decodeRopBuffer(logonReader.readBytes(logonReader.readUInt32LE()));
            const logonRopsReader = new BufferReader(logonRopsList);
            logonRopsReader.readUInt8();
            logonRopsReader.readUInt8();
            logonRopsReader.readUInt32LE();
            logonRopsReader.readUInt8();
            logonRopsReader.readBigUInt64LE();
            logonRopsReader.readBigUInt64LE();
            logonRopsReader.readBigUInt64LE();
            logonRopsReader.readBigUInt64LE();
            const inboxFid = logonRopsReader.readBigUInt64LE();

            // OpenFolder(inbox) -> GetContentsTable -> SetColumns(Subject, Mid) -> QueryRows, to learn the MID.
            await execute(cookie, encodeRopBuffer({ ropsList: buildOpenFolderRops(0, 1, inboxFid), handleTable: [0xffffffff, 0xffffffff] }));
            await execute(cookie, encodeRopBuffer({ ropsList: buildGetContentsTableRops(1, 2), handleTable: [0xffffffff, 0xffffffff, 0xffffffff] }));
            const listColumns = [
                { propertyId: 0x0037, propertyType: PropertyType.PtypString },
                { propertyId: 0x674a, propertyType: PropertyType.PtypInteger64 },
            ];
            await execute(cookie, encodeRopBuffer({ ropsList: buildSetColumnsRops(2, listColumns), handleTable: [0xffffffff] }));
            const queryRowsResult = await execute(cookie, encodeRopBuffer({ ropsList: buildQueryRowsRops(2, 10), handleTable: [0xffffffff] }));
            const queryRowsReader = new BufferReader(queryRowsResult.body);
            queryRowsReader.readUInt32LE();
            queryRowsReader.readUInt32LE();
            queryRowsReader.readUInt32LE();
            const { ropsList: queryRowsList } = decodeRopBuffer(queryRowsReader.readBytes(queryRowsReader.readUInt32LE()));
            const rowsReader = new BufferReader(queryRowsList);
            rowsReader.readUInt8();
            rowsReader.readUInt8();
            rowsReader.readUInt32LE();
            rowsReader.readUInt8();
            rowsReader.readUInt16LE(); // RowCount
            rowsReader.readUInt8(); // PropertyRow Flags
            readPropertyValue(rowsReader, PropertyType.PtypString); // Subject
            const mid = readPropertyValue(rowsReader, PropertyType.PtypInteger64) as bigint;

            // OpenMessage, then RopGetPropertyIdsFromNames(Kind=Name "Keywords") to learn its session-scoped ID.
            await execute(cookie, encodeRopBuffer({ ropsList: buildOpenMessageRops(1, 3, inboxFid, mid), handleTable: [0xffffffff, 0xffffffff] }));
            const namesResult = await execute(
                cookie,
                encodeRopBuffer({ ropsList: buildGetPropertyIdsFromNamesForNameRops(0, PSETID_PUBLIC_STRINGS, "Keywords"), handleTable: [0xffffffff] }),
            );
            const namesReader = new BufferReader(namesResult.body);
            namesReader.readUInt32LE();
            namesReader.readUInt32LE();
            namesReader.readUInt32LE();
            const { ropsList: namesRopsList } = decodeRopBuffer(namesReader.readBytes(namesReader.readUInt32LE()));
            const namesRopsReader = new BufferReader(namesRopsList);
            namesRopsReader.readUInt8();
            namesRopsReader.readUInt8();
            expect(namesRopsReader.readUInt32LE()).toBe(0); // ReturnValue
            expect(namesRopsReader.readUInt16LE()).toBe(1); // PropertyIdCount
            const keywordsId = namesRopsReader.readUInt16LE();
            expect(keywordsId).toBeGreaterThanOrEqual(0x8000);

            // GetPropertiesSpecific(Keywords) on the opened message handle.
            const getPropsInput = encodeRopBuffer({
                ropsList: buildGetPropertiesSpecificRops(3, [{ propertyId: keywordsId, propertyType: PropertyType.PtypMultipleString }]),
                handleTable: [0xffffffff],
            });
            const getPropsResult = await execute(cookie, getPropsInput);
            const getPropsReader = new BufferReader(getPropsResult.body);
            getPropsReader.readUInt32LE();
            getPropsReader.readUInt32LE();
            getPropsReader.readUInt32LE();
            const { ropsList: getPropsRopsList } = decodeRopBuffer(getPropsReader.readBytes(getPropsReader.readUInt32LE()));
            const getPropsRopsReader = new BufferReader(getPropsRopsList);
            getPropsRopsReader.readUInt8();
            getPropsRopsReader.readUInt8();
            expect(getPropsRopsReader.readUInt32LE()).toBe(0); // ReturnValue
            getPropsRopsReader.readUInt8(); // PropertyRow Flags
            const categories = readPropertyValue(getPropsRopsReader, PropertyType.PtypMultipleString) as string[];
            expect(categories.sort()).toEqual(["Important", "Work"]);
        });
    });

    describe("Compose/Send (RopCreateMessage / RopSetProperties / RopOpenStream / RopWriteStream / RopSaveChangesMessage / RopSubmitMessage)", () => {
        const buildLogonRops = function (outputHandleIndex: number): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0xfe);
            writer.writeUInt8(0);
            writer.writeUInt8(outputHandleIndex);
            writer.writeUInt8(0x01);
            writer.writeUInt32LE(0);
            writer.writeUInt32LE(0);
            writer.writeUInt16LE(0);
            return writer.toBuffer();
        };

        const buildCreateMessageRops = function (inputHandleIndex: number, outputHandleIndex: number, folderId: bigint): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0x06);
            writer.writeUInt8(0); // LogonId
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt8(outputHandleIndex);
            writer.writeUInt16LE(0); // CodePageId
            writer.writeBigUInt64LE(folderId);
            writer.writeUInt8(0); // AssociatedFlag
            return writer.toBuffer();
        };

        const buildSetPropertiesRops = function (inputHandleIndex: number, values: TaggedPropertyValue[]): Buffer {
            const valuesWriter = new BufferWriter();
            for (const value of values) {
                writeTaggedPropertyValue(valuesWriter, value);
            }
            const valuesBytes = valuesWriter.toBuffer();
            const writer = new BufferWriter();
            writer.writeUInt8(0x0a);
            writer.writeUInt8(0); // LogonId
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt16LE(2 + valuesBytes.length);
            writer.writeUInt16LE(values.length);
            writer.writeBytes(valuesBytes);
            return writer.toBuffer();
        };

        const buildOpenStreamRops = function (
            inputHandleIndex: number,
            outputHandleIndex: number,
            propertyTag: { propertyId: number; propertyType: PropertyType },
            openModeFlags: number,
        ): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0x2b);
            writer.writeUInt8(0); // LogonId
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt8(outputHandleIndex);
            writePropertyTag(writer, propertyTag);
            writer.writeUInt8(openModeFlags);
            return writer.toBuffer();
        };

        const buildWriteStreamRops = function (inputHandleIndex: number, data: Buffer): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0x2d);
            writer.writeUInt8(0); // LogonId
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt16LE(data.length);
            writer.writeBytes(data);
            return writer.toBuffer();
        };

        const buildSaveChangesMessageRops = function (responseHandleIndex: number, inputHandleIndex: number): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0x0c);
            writer.writeUInt8(0); // LogonId
            writer.writeUInt8(responseHandleIndex);
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt8(0); // SaveFlags
            return writer.toBuffer();
        };

        const buildSubmitMessageRops = function (inputHandleIndex: number): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0x32);
            writer.writeUInt8(0); // LogonId
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt8(0); // SubmitFlags
            return writer.toBuffer();
        };

        it("Deletes a released write stream's chunks from the shared store once the session is saved.", async () => {
            await createMailbox(owner.uid);
            const cookie = cookieHeaderFrom((await connect()).headers["set-cookie"]);
            const sessionId = /MapiContext=([^;]+)/.exec(cookie)![1];
            // Best effort: a failing delete doesn't fail the request.
            const del = vi.spyOn(MemoryHandleDataStore.prototype, "delete").mockRejectedValueOnce(new Error("redis down"));
            try {
                await execute(cookie, encodeRopBuffer({ ropsList: buildLogonRops(0), handleTable: [0xffffffff] }));
                const rops = Buffer.concat([
                    buildCreateMessageRops(0, 3, 5n),
                    buildOpenStreamRops(3, 4, { propertyId: 0x1000, propertyType: PropertyType.PtypString }, 0x02),
                    buildWriteStreamRops(4, Buffer.from("hi", "utf16le")),
                    Buffer.from([0x01, 0x00, 0x04]), // RopRelease(4)
                ]);
                const result = await execute(cookie, encodeRopBuffer({ ropsList: rops, handleTable: [0xffffffff] }));
                expect(result.headers["x-responsecode"]).toBe("0");
                expect(del.mock.calls.some(([key]) => key.startsWith(`${sessionId}:4:`) && key.endsWith(":write"))).toBe(true);
            } finally {
                del.mockRestore();
            }
        });

        it("Invites a meeting's attendees only for revisions the client submitted, once each (restapi's MeetingSchedulingJob).", async () => {
            mailTransport().sent = [];
            await createMailbox(owner.uid);
            const cookie = cookieHeaderFrom((await connect()).headers["set-cookie"]);
            const job: any = await objectFactory.newInstance(MeetingSchedulingJobMongo, { name: "MeetingSchedulingJobMongo" });
            const first = `${uuid.v4()}@example.com`;
            const second = `${uuid.v4()}@example.com`;
            const invitesTo = (address: string) => mailTransport().sent.filter((message) => message.envelopeTo.includes(address)).length;
            const run = async (rops: Buffer): Promise<number> => {
                const result = await execute(cookie, encodeRopBuffer({ ropsList: rops, handleTable: [0xffffffff, 0xffffffff] }));
                const reader = new BufferReader(result.body);
                reader.readBytes(12);
                return decodeRopBuffer(reader.readBytes(reader.readUInt32LE())).ropsList.readUInt32LE(2);
            };
            const attendees = (list: string) => buildSetPropertiesRops(3, [
                { propertyId: 0x001a, propertyType: PropertyType.PtypString, value: "IPM.Appointment" },
                { propertyId: 0x0037, propertyType: PropertyType.PtypString, value: "Planning" },
                { propertyId: 0x0e04, propertyType: PropertyType.PtypString, value: list },
            ]);

            await execute(cookie, encodeRopBuffer({ ropsList: buildLogonRops(0), handleTable: [0xffffffff] }));
            expect(await run(buildCreateMessageRops(0, 3, 5n))).toBe(0);
            expect(await run(attendees(first))).toBe(0);

            // Save & Close without sending: nobody is invited.
            expect(await run(buildSaveChangesMessageRops(7, 3))).toBe(0);
            await job.run();
            expect(invitesTo(first)).toBe(0);

            // Submitting the saved meeting invites once, however often the job runs or the client repeats the submit
            // before it does.
            expect(await run(buildSubmitMessageRops(3))).toBe(0);
            expect(await run(buildSubmitMessageRops(3))).toBe(0);
            await job.run();
            await job.run();
            expect(invitesTo(first)).toBe(1);

            // An edit saved without sending (a new attendee bumps the sequence) invites nobody.
            expect(await run(attendees(`${first}; ${second}`))).toBe(0);
            expect(await run(buildSaveChangesMessageRops(7, 3))).toBe(0);
            await job.run();
            expect(invitesTo(first)).toBe(1);
            expect(invitesTo(second)).toBe(0);

            // Sending that edit later sends the edited meeting once.
            expect(await run(buildSubmitMessageRops(3))).toBe(0);
            await job.run();
            await job.run();
            expect(invitesTo(first)).toBe(2);
            expect(invitesTo(second)).toBe(1);
        });

        it("Composes and sends a real message end to end, saving a Sent Items copy.", async () => {
            mailTransport().sent = [];
            const mailbox = await createMailbox(owner.uid);
            const connectResult = await connect();
            const cookie = cookieHeaderFrom(connectResult.headers["set-cookie"]);

            // Logon (handle 0) - discover the Inbox FID (used only as RopCreateMessage's target folder; the
            // actual saved copy always lands in Sent Items regardless, matching RopSubmitMessageHandler's own
            // documented behavior).
            const logonResult = await execute(cookie, encodeRopBuffer({ ropsList: buildLogonRops(0), handleTable: [0xffffffff] }));
            const logonReader = new BufferReader(logonResult.body);
            logonReader.readUInt32LE();
            logonReader.readUInt32LE();
            logonReader.readUInt32LE();
            const { ropsList: logonRopsList } = decodeRopBuffer(logonReader.readBytes(logonReader.readUInt32LE()));
            const logonRopsReader = new BufferReader(logonRopsList);
            logonRopsReader.readUInt8();
            logonRopsReader.readUInt8();
            logonRopsReader.readUInt32LE();
            logonRopsReader.readUInt8();
            logonRopsReader.readBigUInt64LE(); // Root
            logonRopsReader.readBigUInt64LE(); // Deferred Action
            logonRopsReader.readBigUInt64LE(); // Spooler Queue
            logonRopsReader.readBigUInt64LE(); // IPM Subtree
            const inboxFid = logonRopsReader.readBigUInt64LE(); // Inbox

            // RopCreateMessage at handle 3.
            const createMessageResult = await execute(
                cookie,
                encodeRopBuffer({ ropsList: buildCreateMessageRops(0, 3, inboxFid), handleTable: [0xffffffff, 0xffffffff] }),
            );
            const createMessageReader = new BufferReader(createMessageResult.body);
            createMessageReader.readUInt32LE();
            createMessageReader.readUInt32LE();
            createMessageReader.readUInt32LE();
            const { ropsList: createMessageRopsList } = decodeRopBuffer(createMessageReader.readBytes(createMessageReader.readUInt32LE()));
            const createMessageRopsReader = new BufferReader(createMessageRopsList);
            createMessageRopsReader.readUInt8();
            createMessageRopsReader.readUInt8();
            expect(createMessageRopsReader.readUInt32LE()).toBe(0); // ReturnValue - success

            // RopSetProperties(Subject, DisplayTo) on message handle 3.
            const setPropertiesInput = encodeRopBuffer({
                ropsList: buildSetPropertiesRops(3, [
                    { propertyId: 0x0037, propertyType: PropertyType.PtypString, value: "Hello From MAPI" },
                    { propertyId: 0x0e04, propertyType: PropertyType.PtypString, value: "recipient@example.com" },
                ]),
                handleTable: [0xffffffff],
            });
            const setPropertiesResult = await execute(cookie, setPropertiesInput);
            expect(setPropertiesResult.headers["x-responsecode"]).toBe("0");

            // RopOpenStream(PidTagBody, Create) at handle 4, then RopWriteStream to set the body text.
            const openStreamInput = encodeRopBuffer({
                ropsList: buildOpenStreamRops(3, 4, { propertyId: 0x1000, propertyType: PropertyType.PtypString }, 0x02),
                handleTable: [0xffffffff, 0xffffffff],
            });
            await execute(cookie, openStreamInput);

            const bodyText = "This message was composed entirely via MAPI ROPs.";
            const bodyBytes = Buffer.concat([Buffer.from(bodyText, "utf16le"), Buffer.from([0, 0])]);
            const writeStreamInput = encodeRopBuffer({ ropsList: buildWriteStreamRops(4, bodyBytes), handleTable: [0xffffffff] });
            const writeStreamResult = await execute(cookie, writeStreamInput);
            const writeStreamReader = new BufferReader(writeStreamResult.body);
            writeStreamReader.readUInt32LE();
            writeStreamReader.readUInt32LE();
            writeStreamReader.readUInt32LE();
            const { ropsList: writeStreamRopsList } = decodeRopBuffer(writeStreamReader.readBytes(writeStreamReader.readUInt32LE()));
            const writeStreamRopsReader = new BufferReader(writeStreamRopsList);
            writeStreamRopsReader.readUInt8();
            writeStreamRopsReader.readUInt8();
            expect(writeStreamRopsReader.readUInt32LE()).toBe(0); // ReturnValue
            expect(writeStreamRopsReader.readUInt16LE()).toBe(bodyBytes.length); // WrittenSize

            // RopSaveChangesMessage at response handle 7, on message handle 3.
            const saveChangesInput = encodeRopBuffer({ ropsList: buildSaveChangesMessageRops(7, 3), handleTable: [0xffffffff] });
            const saveChangesResult = await execute(cookie, saveChangesInput);
            const saveChangesReader = new BufferReader(saveChangesResult.body);
            saveChangesReader.readUInt32LE();
            saveChangesReader.readUInt32LE();
            saveChangesReader.readUInt32LE();
            const { ropsList: saveChangesRopsList } = decodeRopBuffer(saveChangesReader.readBytes(saveChangesReader.readUInt32LE()));
            const saveChangesRopsReader = new BufferReader(saveChangesRopsList);
            saveChangesRopsReader.readUInt8();
            saveChangesRopsReader.readUInt8();
            expect(saveChangesRopsReader.readUInt32LE()).toBe(0); // ReturnValue

            // RopSubmitMessage on message handle 3 - the real send trigger.
            const submitInput = encodeRopBuffer({ ropsList: buildSubmitMessageRops(3), handleTable: [0xffffffff] });
            const submitResult = await execute(cookie, submitInput);
            const submitReader = new BufferReader(submitResult.body);
            submitReader.readUInt32LE();
            submitReader.readUInt32LE();
            submitReader.readUInt32LE();
            const { ropsList: submitRopsList } = decodeRopBuffer(submitReader.readBytes(submitReader.readUInt32LE()));
            const submitRopsReader = new BufferReader(submitRopsList);
            submitRopsReader.readUInt8();
            submitRopsReader.readUInt8();
            expect(submitRopsReader.readUInt32LE()).toBe(0); // ReturnValue - success

            expect(mailTransport().sent.length).toBe(1);
            expect(mailTransport().sent[0].envelopeFrom).toBe(mailbox.primarySmtpAddress);
            expect(mailTransport().sent[0].envelopeTo).toEqual(["recipient@example.com"]);
            expect(mailTransport().sent[0].raw.toString("utf-8")).toContain("Hello From MAPI");

            const sentMessages = await messageRepo.find({ subject: "Hello From MAPI" }).toArray();
            expect(sentMessages.length).toBe(1);
            expect(sentMessages[0].recipients).toEqual([{ address: "recipient@example.com", type: RecipientType.TO }]);
            const savedRaw = await blobStore().get(sentMessages[0].bodyBlobKey);
            expect(savedRaw.toString("utf-8")).toContain(bodyText);

            // RopSubmitMessageHandler must refresh the Sent Items folder's cached counts after filing the copy above -
            // otherwise Outlook's own folder pane, which reads Folder.totalCount/unreadCount directly, never learns
            // this mailbox's Sent Items now holds a message.
            const sentFolder: any = await folderRepo.findOne({ uid: sentMessages[0].folderUid } as any);
            expect(sentFolder?.type).toBe(FolderType.SENT_ITEMS);
            expect(sentFolder?.totalCount).toBe(1);
            expect(sentFolder?.unreadCount).toBe(0); // a sender's own Sent Items copy is always already "read"
        });
    });

    describe("RopGetPropertyIdsFromNames", () => {
        const PSETID_APPOINTMENT = "00062002-0000-0000-c000-000000000046";

        const buildLogonRops = function (outputHandleIndex: number): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0xfe);
            writer.writeUInt8(0);
            writer.writeUInt8(outputHandleIndex);
            writer.writeUInt8(0x01);
            writer.writeUInt32LE(0);
            writer.writeUInt32LE(0);
            writer.writeUInt16LE(0);
            return writer.toBuffer();
        };

        const buildGetPropertyIdsFromNamesRops = function (
            inputHandleIndex: number,
            propertyNames: { guid: string; lid: number }[],
        ): Buffer {
            const namesWriter = new BufferWriter();
            for (const name of propertyNames) {
                namesWriter.writeUInt8(0x00); // Kind - LID
                namesWriter.writeBytes(encodeGuid(name.guid));
                namesWriter.writeUInt32LE(name.lid);
            }
            const namesBytes = namesWriter.toBuffer();

            const writer = new BufferWriter();
            writer.writeUInt8(0x56);
            writer.writeUInt8(0); // LogonId
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt8(0x02); // Flags - assign new IDs
            writer.writeUInt16LE(propertyNames.length);
            writer.writeBytes(namesBytes);
            return writer.toBuffer();
        };

        it("Resolves named properties into stable, session-scoped numeric IDs starting at 0x8000.", async () => {
            await createMailbox(owner.uid);
            const connectResult = await connect();
            const cookie = cookieHeaderFrom(connectResult.headers["set-cookie"]);

            // Logon (handle 0) - any real handle works as RopGetPropertyIdsFromNames's InputHandleIndex.
            await execute(cookie, encodeRopBuffer({ ropsList: buildLogonRops(0), handleTable: [0xffffffff] }));

            // Arbitrary LIDs within PSETID_Appointment - this test only exercises the ROP's own mechanics
            // (resolving a PropertyName to a stable session-scoped numeric ID), not any specific real Calendar
            // property's exact LID value (those are pinned against MS-OXPROPS individually where they're
            // actually consumed - see PropertyResolvers.ts's own Calendar property mapping, a later step).
            const propertyNames = [
                { guid: PSETID_APPOINTMENT, lid: 0x8208 },
                { guid: PSETID_APPOINTMENT, lid: 0x820d },
            ];
            const result = await execute(cookie, encodeRopBuffer({ ropsList: buildGetPropertyIdsFromNamesRops(0, propertyNames), handleTable: [0xffffffff] }));

            const reader = new BufferReader(result.body);
            reader.readUInt32LE();
            reader.readUInt32LE();
            reader.readUInt32LE();
            const { ropsList } = decodeRopBuffer(reader.readBytes(reader.readUInt32LE()));
            const ropsReader = new BufferReader(ropsList);
            expect(ropsReader.readUInt8()).toBe(0x56); // RopId
            expect(ropsReader.readUInt8()).toBe(0); // InputHandleIndex, echoed
            expect(ropsReader.readUInt32LE()).toBe(0); // ReturnValue
            expect(ropsReader.readUInt16LE()).toBe(2); // PropertyIdCount
            const id1 = ropsReader.readUInt16LE();
            const id2 = ropsReader.readUInt16LE();
            expect(id1).toBe(0x8000);
            expect(id2).toBe(0x8001);
            expect(ropsReader.hasMore()).toBe(false);

            // Resolving the exact same PropertyName again in a later Execute call reuses the same ID - the
            // registry is session-scoped, not re-assigned per call.
            const secondResult = await execute(
                cookie,
                encodeRopBuffer({ ropsList: buildGetPropertyIdsFromNamesRops(0, [propertyNames[0]]), handleTable: [0xffffffff] }),
            );
            const secondReader = new BufferReader(secondResult.body);
            secondReader.readUInt32LE();
            secondReader.readUInt32LE();
            secondReader.readUInt32LE();
            const { ropsList: secondRopsList } = decodeRopBuffer(secondReader.readBytes(secondReader.readUInt32LE()));
            const secondRopsReader = new BufferReader(secondRopsList);
            secondRopsReader.readUInt8();
            secondRopsReader.readUInt8();
            secondRopsReader.readUInt32LE();
            secondRopsReader.readUInt16LE();
            expect(secondRopsReader.readUInt16LE()).toBe(id1);
        });
    });

    describe("Disconnect", () => {
        it("Releases the session such that a subsequent Execute reports session-not-found.", async () => {
            await createMailbox(owner.uid);
            const connectResult = await connect();
            const cookie = cookieHeaderFrom(connectResult.headers["set-cookie"]);

            const disconnectResult = await mapiRequest(
                server.getApplication(),
                baseUrl,
                {
                    Authorization: "jwt " + ownerToken,
                    "X-RequestType": "Disconnect",
                    "Content-Type": "application/mapi-http",
                    Cookie: cookie,
                },
                Buffer.alloc(0),
            );
            expect(disconnectResult.status).toBe(200);

            const emptyRop = encodeRopBuffer({ ropsList: Buffer.alloc(0), handleTable: [] });
            const executeResult = await execute(cookie, emptyRop);
            expect(executeResult.headers["x-responsecode"]).not.toBe("0");
        });

        it("Treats another user's MapiContext session as not found for both Execute and Disconnect.", async () => {
            await createMailbox(owner.uid);
            const connectResult = await connect();
            const cookie = cookieHeaderFrom(connectResult.headers["set-cookie"]);
            const other: any = { uid: uuid.v4(), roles: [], elevated: Date.now() };
            await createMailbox(other.uid);
            const otherToken = JWTUtils.createTokenSync(config.get("auth"), other);
            const emptyRop = encodeRopBuffer({ ropsList: Buffer.alloc(0), handleTable: [] });
            const asOther = (requestType: string, body: Buffer) =>
                mapiRequest(
                    server.getApplication(),
                    baseUrl,
                    { Authorization: "jwt " + otherToken, "X-RequestType": requestType, "Content-Type": "application/mapi-http", Cookie: cookie },
                    body,
                );

            const executeBody = new BufferWriter();
            executeBody.writeUInt32LE(0); // Flags
            executeBody.writeUInt32LE(emptyRop.length);
            executeBody.writeBytes(emptyRop);
            executeBody.writeUInt32LE(256 * 1024); // MaxRopOut
            executeBody.writeUInt32LE(0); // AuxiliaryBufferSize
            const stolenExecute = await asOther("Execute", executeBody.toBuffer());
            expect(stolenExecute.status).toBe(200);
            expect(stolenExecute.headers["x-responsecode"]).not.toBe("0");
            const reader = new BufferReader(stolenExecute.body);
            expect(reader.readUInt32LE()).toBe(0); // StatusCode
            expect(reader.readUInt32LE()).not.toBe(0); // ErrorCode

            const stolenDisconnect = await asOther("Disconnect", Buffer.alloc(0));
            expect(stolenDisconnect.status).toBe(200);

            // The owner's session survived the other user's Disconnect.
            const ownerExecute = await execute(cookie, emptyRop);
            expect(ownerExecute.headers["x-responsecode"]).toBe("0");
        });
    });

    describe("Unrecognized/deferred request types", () => {
        it("Returns 400 for a missing/unrecognized X-RequestType.", async () => {
            const result = await mapiRequest(
                server.getApplication(),
                baseUrl,
                { Authorization: "jwt " + ownerToken, "Content-Type": "application/mapi-http" },
                Buffer.alloc(0),
            );
            expect(result.status).toBe(400);
        });

        it("Returns 501 for the recognized-but-deferred NotificationWait request type.", async () => {
            const result = await mapiRequest(
                server.getApplication(),
                baseUrl,
                {
                    Authorization: "jwt " + ownerToken,
                    "X-RequestType": "NotificationWait",
                    "Content-Type": "application/mapi-http",
                },
                Buffer.alloc(0),
            );
            expect(result.status).toBe(501);
        });
    });
});
