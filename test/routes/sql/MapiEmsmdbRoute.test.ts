///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
// See the identical file header in test/routes/mongo/MapiEmsmdbRoute.test.ts for the full rationale - this
// verifies the same BaseMapiEmsmdbRoute transport plumbing on the SQL-backed variant.
import { testZeroAccessKey } from "../../zeroAccessTestKey.js";
import config from "../../config.sql.js";
import { Server, ObjectFactory, ConnectionManager, isSqlDataSource, ACLAction, AccessControlListSQL, RepoUtils } from "@rapidrest/service-core";
import { JWTUtils, Logger } from "@rapidrest/core";
import * as uuid from "uuid";
import { inspect } from "node:util";
import { Repository } from "typeorm";
import { CalendarEventSQL, ContactSQL, FolderSQL, LabelSQL, MailboxSQL, MeetingSchedulingJobSQL, MessageSQL, TaskSQL, ZeroAccessIndexEntrySQL } from "@rapidmx/restapi/sql";
import { FolderType, MessageImportance, RecipientType } from "@rapidmx/restapi";
import { InMemoryBlobStore, RecordingMailTransport, registerTestDoubles } from "../../testDoubles.js";
import { cookieHeaderFrom, mapiRequest } from "../../mapiTestClient.js";
import { BufferReader, BufferWriter } from "../../../src/codec/BufferCursor.js";
import { decodeGuid, encodeGuid } from "../../../src/codec/MapiGuid.js";
import { PropertyType, readPropertyValue, writePropertyTag, writeTaggedPropertyValue } from "../../../src/codec/PropertyValue.js";
import { decodeRopBuffer, encodeRopBuffer } from "../../../src/codec/RopBuffer.js";
import { MapiEmsmdbRoute } from "../../server-sql/routes/MapiEmsmdbRoute.js";

describe("Route:MapiEmsmdbRouteSQL Tests", () => {
    const logger = Logger();
    const objectFactory: ObjectFactory = new ObjectFactory(config, logger);
    const server: Server = new Server({ config, basePath: "./test/server-sql", logger, objectFactory });
    const baseUrl = "/sql/mapi/emsmdb";
    let mailboxRepo: Repository<MailboxSQL>;
    let folderRepo: Repository<FolderSQL>;
    let messageRepo: Repository<MessageSQL>;
    let contactRepo: Repository<ContactSQL>;
    let taskRepo: Repository<TaskSQL>;
    let labelRepo: Repository<LabelSQL>;
    let aclRepo: Repository<AccessControlListSQL>;

    const owner: any = { uid: uuid.v4(), roles: [], elevated: Date.now() };
    const ownerToken = JWTUtils.createTokenSync(config.get("auth"), owner);

    const createMailbox = async function (ownerUid: string, data?: Partial<MailboxSQL>): Promise<MailboxSQL> {
        const obj: MailboxSQL = new MailboxSQL({
            ownerUserUid: ownerUid,
            primarySmtpAddress: `${uuid.v4()}@example.com`,
            aliasAddresses: [],
            displayName: "Test Mailbox",
            timezone: "UTC",
            quotaBytes: 1_000_000_000,
            usedBytes: 0,
            ...data,
        });
        const result: MailboxSQL = await mailboxRepo.save(obj);
        await aclRepo.save({
            uid: result.uid,
            dateCreated: new Date(),
            dateModified: new Date(),
            version: 0,
            records: [{ userOrRoleId: ownerUid, actions: [ACLAction.FULL] }],
            parentUid: "Mailbox",
        } as any);
        return result;
    };

    const createFolder = async function (mailboxUid: string, type: FolderType, name: string): Promise<FolderSQL> {
        return await folderRepo.save(
            new FolderSQL({ mailboxUid, name, type, unreadCount: 0, totalCount: 0, syncKeyVersion: 0 } as any),
        );
    };

    const createMessage = async function (mailboxUid: string, folderUid: string, data?: Partial<MessageSQL>): Promise<MessageSQL> {
        return await messageRepo.save(
            new MessageSQL({
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
            } as any),
        );
    };

    const createContact = async function (mailboxUid: string, folderUid: string, data?: Partial<ContactSQL>): Promise<ContactSQL> {
        return await contactRepo.save(
            new ContactSQL({ mailboxUid, folderUid, displayName: "Test Contact", emails: [], phones: [], addresses: [], ...data } as any),
        );
    };

    const createTask = async function (mailboxUid: string, folderUid: string, data?: Partial<TaskSQL>): Promise<TaskSQL> {
        return await taskRepo.save(new TaskSQL({ mailboxUid, folderUid, title: "Test Task", ...data } as any));
    };

    const createLabel = async function (mailboxUid: string, data?: Partial<LabelSQL>): Promise<LabelSQL> {
        return await labelRepo.save(new LabelSQL({ mailboxUid, name: "Test Label", ...data } as any));
    };

    const execute = async function (cookie: string, ropBuffer: Buffer) {
        const body = new BufferWriter();
        body.writeUInt32LE(0);
        body.writeUInt32LE(ropBuffer.length);
        body.writeBytes(ropBuffer);
        body.writeUInt32LE(256 * 1024);
        body.writeUInt32LE(0);
        return mapiRequest(
            server.getApplication(),
            baseUrl,
            {
                Authorization: "jwt " + ownerToken,
                "X-RequestType": "Execute",
                "Content-Type": "application/mapi-http",
                Cookie: cookie,
            },
            body.toBuffer(),
        );
    };

    /** The mailbox's messages - the backend's own query shape. */
    const messagesOf = async (mailboxUid: string): Promise<any[]> => await messageRepo.find({ where: { mailboxUid } });

    const blobStore = function (): InMemoryBlobStore {
        return objectFactory.getInstance<InMemoryBlobStore>("BlobStore")!;
    };

    const mailTransport = function (): RecordingMailTransport {
        return objectFactory.getInstance<RecordingMailTransport>("MailTransport")!;
    };

    const connect = async function () {
        return mapiRequest(
            server.getApplication(),
            baseUrl,
            {
                Authorization: "jwt " + ownerToken,
                "X-RequestType": "Connect",
                "Content-Type": "application/mapi-http",
            },
            Buffer.alloc(0),
        );
    };

    beforeAll(async () => {
        registerTestDoubles(objectFactory);
        await server.start();

        const connMgr: ConnectionManager | undefined = objectFactory.getInstance(ConnectionManager);
        const conn: any = connMgr?.connections.get("sql");
        if (isSqlDataSource(conn)) {
            mailboxRepo = conn.getRepository(MailboxSQL);
            folderRepo = conn.getRepository(FolderSQL);
            messageRepo = conn.getRepository(MessageSQL);
            contactRepo = conn.getRepository(ContactSQL);
            taskRepo = conn.getRepository(TaskSQL);
            labelRepo = conn.getRepository(LabelSQL);
        } else {
            throw new Error("Could not find sql connection");
        }
        const aclConn: any = connMgr?.connections.get("acl");
        if (isSqlDataSource(aclConn)) {
            aclRepo = aclConn.getRepository(AccessControlListSQL);
        } else {
            throw new Error("Could not find sql acl connection");
        }
    });

    afterAll(async () => {
        await server.stop();
        await objectFactory.destroy();
    });

    beforeEach(async () => {
        await mailboxRepo.clear();
        await folderRepo.clear();
        await messageRepo.clear();
        await contactRepo.clear();
        await taskRepo.clear();
        await labelRepo.clear();
        await aclRepo.clear();
    });

    it("Establishes a session and returns a well-formed Connect success body.", async () => {
        await createMailbox(owner.uid, { displayName: "Ada Lovelace" });
        const result = await connect();

        expect(result.status).toBe(200);
        const cookie = cookieHeaderFrom(result.headers["set-cookie"]);
        expect(cookie).toContain("MapiContext=");

        const reader = new BufferReader(result.body);
        expect(reader.readUInt32LE()).toBe(0); // StatusCode
        expect(reader.readUInt32LE()).toBe(0); // ErrorCode
        reader.readUInt32LE(); // PollsMax
        reader.readUInt32LE(); // RetryCount
        reader.readUInt32LE(); // RetryDelay
        reader.readNullTerminatedString8(); // DnPrefix
        expect(reader.readNullTerminatedUtf16LE()).toBe("Ada Lovelace");
    });

    it("Records on the mailbox that Outlook syncs it, refreshing that at most every few hours.", async () => {
        const mailbox = await createMailbox(owner.uid, { syncClients: { activesync: { devices: [] } } });
        expect((await connect()).status).toBe(200);
        const recorded: any = await mailboxRepo.findOne({ where: { uid: mailbox.uid } });
        expect(recorded?.syncClients).toEqual({ activesync: { devices: [] }, mapi: { lastSeenAt: expect.any(Number) } });
        expect((await connect()).status).toBe(200);
        expect(((await mailboxRepo.findOne({ where: { uid: mailbox.uid } })) as any)?.version).toBe(recorded.version);
    });

    it("Echoes back an empty but well-formed ROP buffer for a valid session.", async () => {
        await createMailbox(owner.uid);
        const connectResult = await connect();
        const cookie = cookieHeaderFrom(connectResult.headers["set-cookie"]);

        const inputRop = encodeRopBuffer({ ropsList: Buffer.alloc(0), handleTable: [42] });
        const requestBody = new BufferWriter();
        requestBody.writeUInt32LE(0);
        requestBody.writeUInt32LE(inputRop.length);
        requestBody.writeBytes(inputRop);
        requestBody.writeUInt32LE(256 * 1024);
        requestBody.writeUInt32LE(0);

        const result = await mapiRequest(
            server.getApplication(),
            baseUrl,
            {
                Authorization: "jwt " + ownerToken,
                "X-RequestType": "Execute",
                "Content-Type": "application/mapi-http",
                Cookie: cookie,
            },
            requestBody.toBuffer(),
        );

        expect(result.status).toBe(200);
        expect(result.headers["x-responsecode"]).toBe("0");
        const reader = new BufferReader(result.body);
        expect(reader.readUInt32LE()).toBe(0); // StatusCode
        expect(reader.readUInt32LE()).toBe(0); // ErrorCode
        reader.readUInt32LE(); // Flags
        const responseRopBufferSize = reader.readUInt32LE();
        const responseRopBuffer = reader.readBytes(responseRopBufferSize);
        expect(responseRopBuffer.readUInt16LE(0)).toBe(2);
        expect(responseRopBuffer.readUInt32LE(2)).toBe(42);
    });

    it("Logs on to a real mailbox and returns 13 well-formed, distinct FIDs.", async () => {
        const mailbox = await createMailbox(owner.uid);
        const connectResult = await connect();
        const cookie = cookieHeaderFrom(connectResult.headers["set-cookie"]);

        const logonRops = new BufferWriter();
        logonRops.writeUInt8(0xfe); // RopId
        logonRops.writeUInt8(0); // LogonId
        logonRops.writeUInt8(0); // OutputHandleIndex
        logonRops.writeUInt8(0x01); // LogonFlags (Private)
        logonRops.writeUInt32LE(0); // OpenFlags
        logonRops.writeUInt32LE(0); // StoreState
        logonRops.writeUInt16LE(0); // EssdnSize
        const inputRop = encodeRopBuffer({ ropsList: logonRops.toBuffer(), handleTable: [0xffffffff] });

        const requestBody = new BufferWriter();
        requestBody.writeUInt32LE(0);
        requestBody.writeUInt32LE(inputRop.length);
        requestBody.writeBytes(inputRop);
        requestBody.writeUInt32LE(256 * 1024);
        requestBody.writeUInt32LE(0);

        const result = await mapiRequest(
            server.getApplication(),
            baseUrl,
            {
                Authorization: "jwt " + ownerToken,
                "X-RequestType": "Execute",
                "Content-Type": "application/mapi-http",
                Cookie: cookie,
            },
            requestBody.toBuffer(),
        );

        expect(result.status).toBe(200);
        expect(result.headers["x-responsecode"]).toBe("0");
        const reader = new BufferReader(result.body);
        reader.readUInt32LE(); // StatusCode
        reader.readUInt32LE(); // ErrorCode
        reader.readUInt32LE(); // Flags
        const ropBufferSize = reader.readUInt32LE();
        const { ropsList } = decodeRopBuffer(reader.readBytes(ropBufferSize));

        const ropsReader = new BufferReader(ropsList);
        expect(ropsReader.readUInt8()).toBe(0xfe); // RopId
        expect(ropsReader.readUInt8()).toBe(0); // OutputHandleIndex
        expect(ropsReader.readUInt32LE()).toBe(0); // ReturnValue
        expect(ropsReader.readUInt8()).toBe(0x01); // LogonFlags, echoed

        const fids: bigint[] = [];
        for (let i = 0; i < 13; i++) {
            fids.push(ropsReader.readBigUInt64LE());
        }
        expect(new Set(fids).size).toBe(13);

        ropsReader.readUInt8(); // ResponseFlags
        expect(decodeGuid(ropsReader)).toBe(mailbox.uid); // MailboxGuid
    });

    it("Browses the top-level folder list end to end: Logon, OpenFolder(root), GetHierarchyTable, SetColumns, QueryRows.", async () => {
        const mailbox = await createMailbox(owner.uid);
        await createFolder(mailbox.uid, FolderType.INBOX, "Inbox");
        await createFolder(mailbox.uid, FolderType.USER, "Archive");
        const connectResult = await connect();
        const cookie = cookieHeaderFrom(connectResult.headers["set-cookie"]);

        const logonRops = new BufferWriter();
        logonRops.writeUInt8(0xfe);
        logonRops.writeUInt8(0);
        logonRops.writeUInt8(0);
        logonRops.writeUInt8(0x01);
        logonRops.writeUInt32LE(0);
        logonRops.writeUInt32LE(0);
        logonRops.writeUInt16LE(0);
        const logonResult = await execute(cookie, encodeRopBuffer({ ropsList: logonRops.toBuffer(), handleTable: [0xffffffff] }));
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
        const rootFid = logonRopsReader.readBigUInt64LE();

        const openFolderRops = new BufferWriter();
        openFolderRops.writeUInt8(0x02);
        openFolderRops.writeUInt8(0);
        openFolderRops.writeUInt8(0); // InputHandleIndex (logon)
        openFolderRops.writeUInt8(1); // OutputHandleIndex
        openFolderRops.writeUInt8(0); // OpenModeFlags
        openFolderRops.writeBigUInt64LE(rootFid);
        await execute(cookie, encodeRopBuffer({ ropsList: openFolderRops.toBuffer(), handleTable: [0xffffffff, 0xffffffff] }));

        const tableRops = new BufferWriter();
        tableRops.writeUInt8(0x04);
        tableRops.writeUInt8(0);
        tableRops.writeUInt8(1); // InputHandleIndex (folder)
        tableRops.writeUInt8(2); // OutputHandleIndex
        tableRops.writeUInt8(0); // TableFlags
        await execute(cookie, encodeRopBuffer({ ropsList: tableRops.toBuffer(), handleTable: [0xffffffff, 0xffffffff, 0xffffffff] }));

        const setColumnsRops = new BufferWriter();
        setColumnsRops.writeUInt8(0x12);
        setColumnsRops.writeUInt8(0);
        setColumnsRops.writeUInt8(2); // InputHandleIndex (table)
        setColumnsRops.writeUInt8(0); // SetColumnsFlags
        setColumnsRops.writeUInt16LE(1);
        writePropertyTag(setColumnsRops, { propertyId: 0x3001, propertyType: PropertyType.PtypString });
        await execute(cookie, encodeRopBuffer({ ropsList: setColumnsRops.toBuffer(), handleTable: [0xffffffff] }));

        const queryRowsRops = new BufferWriter();
        queryRowsRops.writeUInt8(0x15);
        queryRowsRops.writeUInt8(0);
        queryRowsRops.writeUInt8(2); // InputHandleIndex (table)
        queryRowsRops.writeUInt8(0); // QueryRowsFlags
        queryRowsRops.writeUInt8(1); // ForwardRead
        queryRowsRops.writeUInt16LE(10);
        const queryRowsResult = await execute(cookie, encodeRopBuffer({ ropsList: queryRowsRops.toBuffer(), handleTable: [0xffffffff] }));

        const queryRowsReader = new BufferReader(queryRowsResult.body);
        queryRowsReader.readUInt32LE();
        queryRowsReader.readUInt32LE();
        queryRowsReader.readUInt32LE();
        const { ropsList: queryRowsList } = decodeRopBuffer(queryRowsReader.readBytes(queryRowsReader.readUInt32LE()));
        const rowsReader = new BufferReader(queryRowsList);
        rowsReader.readUInt8();
        rowsReader.readUInt8();
        expect(rowsReader.readUInt32LE()).toBe(0); // ReturnValue
        rowsReader.readUInt8(); // Origin
        const rowCount = rowsReader.readUInt16LE();
        expect(rowCount).toBe(2);

        const names: string[] = [];
        for (let i = 0; i < rowCount; i++) {
            rowsReader.readUInt8(); // PropertyRow Flags
            names.push(readPropertyValue(rowsReader, PropertyType.PtypString) as string);
        }
        expect(names.sort()).toEqual(["Archive", "Inbox"]);
    });

    it("Lists a folder's messages end to end: Logon, OpenFolder(inbox), GetContentsTable, SetColumns, QueryRows.", async () => {
        const mailbox = await createMailbox(owner.uid);
        const inbox = await createFolder(mailbox.uid, FolderType.INBOX, "Inbox");
        await createMessage(mailbox.uid, inbox.uid, { subject: "Hello World" });
        await createMessage(mailbox.uid, inbox.uid, { subject: "Second Message" });
        const connectResult = await connect();
        const cookie = cookieHeaderFrom(connectResult.headers["set-cookie"]);

        const logonRops = new BufferWriter();
        logonRops.writeUInt8(0xfe);
        logonRops.writeUInt8(0);
        logonRops.writeUInt8(0);
        logonRops.writeUInt8(0x01);
        logonRops.writeUInt32LE(0);
        logonRops.writeUInt32LE(0);
        logonRops.writeUInt16LE(0);
        const logonResult = await execute(cookie, encodeRopBuffer({ ropsList: logonRops.toBuffer(), handleTable: [0xffffffff] }));
        const logonReader = new BufferReader(logonResult.body);
        logonReader.readUInt32LE();
        logonReader.readUInt32LE();
        logonReader.readUInt32LE();
        const { ropsList: logonRopsList } = decodeRopBuffer(logonReader.readBytes(logonReader.readUInt32LE()));
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

        const openFolderRops = new BufferWriter();
        openFolderRops.writeUInt8(0x02);
        openFolderRops.writeUInt8(0);
        openFolderRops.writeUInt8(0); // InputHandleIndex (logon)
        openFolderRops.writeUInt8(1); // OutputHandleIndex
        openFolderRops.writeUInt8(0); // OpenModeFlags
        openFolderRops.writeBigUInt64LE(inboxFid);
        await execute(cookie, encodeRopBuffer({ ropsList: openFolderRops.toBuffer(), handleTable: [0xffffffff, 0xffffffff] }));

        const tableRops = new BufferWriter();
        tableRops.writeUInt8(0x05);
        tableRops.writeUInt8(0);
        tableRops.writeUInt8(1); // InputHandleIndex (folder)
        tableRops.writeUInt8(2); // OutputHandleIndex
        tableRops.writeUInt8(0); // TableFlags
        await execute(cookie, encodeRopBuffer({ ropsList: tableRops.toBuffer(), handleTable: [0xffffffff, 0xffffffff, 0xffffffff] }));

        const setColumnsRops = new BufferWriter();
        setColumnsRops.writeUInt8(0x12);
        setColumnsRops.writeUInt8(0);
        setColumnsRops.writeUInt8(2); // InputHandleIndex (table)
        setColumnsRops.writeUInt8(0); // SetColumnsFlags
        setColumnsRops.writeUInt16LE(1);
        writePropertyTag(setColumnsRops, { propertyId: 0x0037, propertyType: PropertyType.PtypString });
        await execute(cookie, encodeRopBuffer({ ropsList: setColumnsRops.toBuffer(), handleTable: [0xffffffff] }));

        const queryRowsRops = new BufferWriter();
        queryRowsRops.writeUInt8(0x15);
        queryRowsRops.writeUInt8(0);
        queryRowsRops.writeUInt8(2); // InputHandleIndex (table)
        queryRowsRops.writeUInt8(0); // QueryRowsFlags
        queryRowsRops.writeUInt8(1); // ForwardRead
        queryRowsRops.writeUInt16LE(10);
        const queryRowsResult = await execute(cookie, encodeRopBuffer({ ropsList: queryRowsRops.toBuffer(), handleTable: [0xffffffff] }));

        const queryRowsReader = new BufferReader(queryRowsResult.body);
        queryRowsReader.readUInt32LE();
        queryRowsReader.readUInt32LE();
        queryRowsReader.readUInt32LE();
        const { ropsList: queryRowsList } = decodeRopBuffer(queryRowsReader.readBytes(queryRowsReader.readUInt32LE()));
        const rowsReader = new BufferReader(queryRowsList);
        rowsReader.readUInt8();
        rowsReader.readUInt8();
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
    });

    /** Discovers `wellKnownFolderName`'s real FID via the root folder's own hierarchy table - Contacts/Tasks
     * aren't part of RopLogon's fixed 13-folder `FolderIds` list (see RopLogonHandler's own doc comment), so a
     * real client (and this test) finds them the same way it finds any other top-level folder: OpenFolder(root)
     * -> GetHierarchyTable -> SetColumns([DisplayName, FolderId]) -> QueryRows. */
    async function discoverFolderId(cookie: string, wellKnownFolderName: string): Promise<bigint> {
        const logonRops = new BufferWriter();
        logonRops.writeUInt8(0xfe);
        logonRops.writeUInt8(0);
        logonRops.writeUInt8(0);
        logonRops.writeUInt8(0x01);
        logonRops.writeUInt32LE(0);
        logonRops.writeUInt32LE(0);
        logonRops.writeUInt16LE(0);
        const logonResult = await execute(cookie, encodeRopBuffer({ ropsList: logonRops.toBuffer(), handleTable: [0xffffffff] }));
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
        const rootFid = logonRopsReader.readBigUInt64LE();

        const openFolderRops = new BufferWriter();
        openFolderRops.writeUInt8(0x02);
        openFolderRops.writeUInt8(0);
        openFolderRops.writeUInt8(0); // InputHandleIndex (logon)
        openFolderRops.writeUInt8(1); // OutputHandleIndex
        openFolderRops.writeUInt8(0); // OpenModeFlags
        openFolderRops.writeBigUInt64LE(rootFid);
        await execute(cookie, encodeRopBuffer({ ropsList: openFolderRops.toBuffer(), handleTable: [0xffffffff, 0xffffffff] }));

        const tableRops = new BufferWriter();
        tableRops.writeUInt8(0x04);
        tableRops.writeUInt8(0);
        tableRops.writeUInt8(1); // InputHandleIndex (folder)
        tableRops.writeUInt8(2); // OutputHandleIndex
        tableRops.writeUInt8(0); // TableFlags
        await execute(cookie, encodeRopBuffer({ ropsList: tableRops.toBuffer(), handleTable: [0xffffffff, 0xffffffff, 0xffffffff] }));

        const setColumnsRops = new BufferWriter();
        setColumnsRops.writeUInt8(0x12);
        setColumnsRops.writeUInt8(0);
        setColumnsRops.writeUInt8(2); // InputHandleIndex (table)
        setColumnsRops.writeUInt8(0); // SetColumnsFlags
        setColumnsRops.writeUInt16LE(2);
        writePropertyTag(setColumnsRops, { propertyId: 0x3001, propertyType: PropertyType.PtypString }); // DisplayName
        writePropertyTag(setColumnsRops, { propertyId: 0x6748, propertyType: PropertyType.PtypInteger64 }); // FolderId
        await execute(cookie, encodeRopBuffer({ ropsList: setColumnsRops.toBuffer(), handleTable: [0xffffffff] }));

        const queryRowsRops = new BufferWriter();
        queryRowsRops.writeUInt8(0x15);
        queryRowsRops.writeUInt8(0);
        queryRowsRops.writeUInt8(2); // InputHandleIndex (table)
        queryRowsRops.writeUInt8(0); // QueryRowsFlags
        queryRowsRops.writeUInt8(1); // ForwardRead
        queryRowsRops.writeUInt16LE(10);
        const queryRowsResult = await execute(cookie, encodeRopBuffer({ ropsList: queryRowsRops.toBuffer(), handleTable: [0xffffffff] }));

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

        const openFolderRops = new BufferWriter();
        openFolderRops.writeUInt8(0x02);
        openFolderRops.writeUInt8(0);
        openFolderRops.writeUInt8(0); // InputHandleIndex (logon)
        openFolderRops.writeUInt8(3); // OutputHandleIndex
        openFolderRops.writeUInt8(0); // OpenModeFlags
        openFolderRops.writeBigUInt64LE(contactsFid);
        await execute(cookie, encodeRopBuffer({ ropsList: openFolderRops.toBuffer(), handleTable: [0xffffffff, 0xffffffff] }));

        const tableRops = new BufferWriter();
        tableRops.writeUInt8(0x05);
        tableRops.writeUInt8(0);
        tableRops.writeUInt8(3); // InputHandleIndex (folder)
        tableRops.writeUInt8(4); // OutputHandleIndex
        tableRops.writeUInt8(0); // TableFlags
        await execute(cookie, encodeRopBuffer({ ropsList: tableRops.toBuffer(), handleTable: [0xffffffff, 0xffffffff, 0xffffffff] }));

        const setColumnsRops = new BufferWriter();
        setColumnsRops.writeUInt8(0x12);
        setColumnsRops.writeUInt8(0);
        setColumnsRops.writeUInt8(4); // InputHandleIndex (table)
        setColumnsRops.writeUInt8(0); // SetColumnsFlags
        setColumnsRops.writeUInt16LE(1);
        writePropertyTag(setColumnsRops, { propertyId: 0x0037, propertyType: PropertyType.PtypString }); // PidTagSubject -> displayName
        await execute(cookie, encodeRopBuffer({ ropsList: setColumnsRops.toBuffer(), handleTable: [0xffffffff] }));

        const queryRowsRops = new BufferWriter();
        queryRowsRops.writeUInt8(0x15);
        queryRowsRops.writeUInt8(0);
        queryRowsRops.writeUInt8(4); // InputHandleIndex (table)
        queryRowsRops.writeUInt8(0); // QueryRowsFlags
        queryRowsRops.writeUInt8(1); // ForwardRead
        queryRowsRops.writeUInt16LE(10);
        const queryRowsResult = await execute(cookie, encodeRopBuffer({ ropsList: queryRowsRops.toBuffer(), handleTable: [0xffffffff] }));

        const queryRowsReader = new BufferReader(queryRowsResult.body);
        queryRowsReader.readUInt32LE();
        queryRowsReader.readUInt32LE();
        queryRowsReader.readUInt32LE();
        const { ropsList: queryRowsList } = decodeRopBuffer(queryRowsReader.readBytes(queryRowsReader.readUInt32LE()));
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

        const openFolderRops = new BufferWriter();
        openFolderRops.writeUInt8(0x02);
        openFolderRops.writeUInt8(0);
        openFolderRops.writeUInt8(0); // InputHandleIndex (logon)
        openFolderRops.writeUInt8(3); // OutputHandleIndex
        openFolderRops.writeUInt8(0); // OpenModeFlags
        openFolderRops.writeBigUInt64LE(tasksFid);
        await execute(cookie, encodeRopBuffer({ ropsList: openFolderRops.toBuffer(), handleTable: [0xffffffff, 0xffffffff] }));

        const tableRops = new BufferWriter();
        tableRops.writeUInt8(0x05);
        tableRops.writeUInt8(0);
        tableRops.writeUInt8(3); // InputHandleIndex (folder)
        tableRops.writeUInt8(4); // OutputHandleIndex
        tableRops.writeUInt8(0); // TableFlags
        await execute(cookie, encodeRopBuffer({ ropsList: tableRops.toBuffer(), handleTable: [0xffffffff, 0xffffffff, 0xffffffff] }));

        const setColumnsRops = new BufferWriter();
        setColumnsRops.writeUInt8(0x12);
        setColumnsRops.writeUInt8(0);
        setColumnsRops.writeUInt8(4); // InputHandleIndex (table)
        setColumnsRops.writeUInt8(0); // SetColumnsFlags
        setColumnsRops.writeUInt16LE(1);
        writePropertyTag(setColumnsRops, { propertyId: 0x0037, propertyType: PropertyType.PtypString }); // PidTagSubject -> title
        await execute(cookie, encodeRopBuffer({ ropsList: setColumnsRops.toBuffer(), handleTable: [0xffffffff] }));

        const queryRowsRops = new BufferWriter();
        queryRowsRops.writeUInt8(0x15);
        queryRowsRops.writeUInt8(0);
        queryRowsRops.writeUInt8(4); // InputHandleIndex (table)
        queryRowsRops.writeUInt8(0); // QueryRowsFlags
        queryRowsRops.writeUInt8(1); // ForwardRead
        queryRowsRops.writeUInt16LE(10);
        const queryRowsResult = await execute(cookie, encodeRopBuffer({ ropsList: queryRowsRops.toBuffer(), handleTable: [0xffffffff] }));

        const queryRowsReader = new BufferReader(queryRowsResult.body);
        queryRowsReader.readUInt32LE();
        queryRowsReader.readUInt32LE();
        queryRowsReader.readUInt32LE();
        const { ropsList: queryRowsList } = decodeRopBuffer(queryRowsReader.readBytes(queryRowsReader.readUInt32LE()));
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

        const logonRops = new BufferWriter();
        logonRops.writeUInt8(0xfe);
        logonRops.writeUInt8(0);
        logonRops.writeUInt8(0);
        logonRops.writeUInt8(0x01);
        logonRops.writeUInt32LE(0);
        logonRops.writeUInt32LE(0);
        logonRops.writeUInt16LE(0);
        const logonResult = await execute(cookie, encodeRopBuffer({ ropsList: logonRops.toBuffer(), handleTable: [0xffffffff] }));
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

        const openFolderRops = new BufferWriter();
        openFolderRops.writeUInt8(0x02);
        openFolderRops.writeUInt8(0);
        openFolderRops.writeUInt8(0); // InputHandleIndex (logon)
        openFolderRops.writeUInt8(1); // OutputHandleIndex
        openFolderRops.writeUInt8(0); // OpenModeFlags
        openFolderRops.writeBigUInt64LE(inboxFid);
        await execute(cookie, encodeRopBuffer({ ropsList: openFolderRops.toBuffer(), handleTable: [0xffffffff, 0xffffffff] }));

        const tableRops = new BufferWriter();
        tableRops.writeUInt8(0x05);
        tableRops.writeUInt8(0);
        tableRops.writeUInt8(1); // InputHandleIndex (folder)
        tableRops.writeUInt8(2); // OutputHandleIndex
        tableRops.writeUInt8(0); // TableFlags
        await execute(cookie, encodeRopBuffer({ ropsList: tableRops.toBuffer(), handleTable: [0xffffffff, 0xffffffff, 0xffffffff] }));

        const setColumnsRops = new BufferWriter();
        setColumnsRops.writeUInt8(0x12);
        setColumnsRops.writeUInt8(0);
        setColumnsRops.writeUInt8(2); // InputHandleIndex (table)
        setColumnsRops.writeUInt8(0); // SetColumnsFlags
        setColumnsRops.writeUInt16LE(2);
        writePropertyTag(setColumnsRops, { propertyId: 0x0037, propertyType: PropertyType.PtypString });
        writePropertyTag(setColumnsRops, { propertyId: 0x674a, propertyType: PropertyType.PtypInteger64 });
        await execute(cookie, encodeRopBuffer({ ropsList: setColumnsRops.toBuffer(), handleTable: [0xffffffff] }));

        const queryRowsRops = new BufferWriter();
        queryRowsRops.writeUInt8(0x15);
        queryRowsRops.writeUInt8(0);
        queryRowsRops.writeUInt8(2); // InputHandleIndex (table)
        queryRowsRops.writeUInt8(0); // QueryRowsFlags
        queryRowsRops.writeUInt8(1); // ForwardRead
        queryRowsRops.writeUInt16LE(10);
        const queryRowsResult = await execute(cookie, encodeRopBuffer({ ropsList: queryRowsRops.toBuffer(), handleTable: [0xffffffff] }));

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

        const openMessageRops = new BufferWriter();
        openMessageRops.writeUInt8(0x03);
        openMessageRops.writeUInt8(0); // LogonId
        openMessageRops.writeUInt8(1); // InputHandleIndex (folder)
        openMessageRops.writeUInt8(3); // OutputHandleIndex
        openMessageRops.writeUInt16LE(0); // CodePageId
        openMessageRops.writeBigUInt64LE(inboxFid);
        openMessageRops.writeUInt8(0); // OpenModeFlags
        openMessageRops.writeBigUInt64LE(mid);
        const openMessageResult = await execute(cookie, encodeRopBuffer({ ropsList: openMessageRops.toBuffer(), handleTable: [0xffffffff, 0xffffffff] }));
        const openMessageReader = new BufferReader(openMessageResult.body);
        openMessageReader.readUInt32LE();
        openMessageReader.readUInt32LE();
        openMessageReader.readUInt32LE();
        const { ropsList: openMessageRopsList } = decodeRopBuffer(openMessageReader.readBytes(openMessageReader.readUInt32LE()));
        const openMessageRopsReader = new BufferReader(openMessageRopsList);
        openMessageRopsReader.readUInt8();
        openMessageRopsReader.readUInt8();
        expect(openMessageRopsReader.readUInt32LE()).toBe(0); // ReturnValue - success

        const getPropsRops = new BufferWriter();
        getPropsRops.writeUInt8(0x07);
        getPropsRops.writeUInt8(0); // LogonId
        getPropsRops.writeUInt8(3); // InputHandleIndex (message)
        getPropsRops.writeUInt16LE(0); // PropertySizeLimit
        getPropsRops.writeUInt16LE(0); // WantUnicode
        getPropsRops.writeUInt16LE(1);
        writePropertyTag(getPropsRops, { propertyId: 0x0037, propertyType: PropertyType.PtypString });
        const getPropsResult = await execute(cookie, encodeRopBuffer({ ropsList: getPropsRops.toBuffer(), handleTable: [0xffffffff] }));
        const getPropsReader = new BufferReader(getPropsResult.body);
        getPropsReader.readUInt32LE();
        getPropsReader.readUInt32LE();
        getPropsReader.readUInt32LE();
        const { ropsList: getPropsRopsList } = decodeRopBuffer(getPropsReader.readBytes(getPropsReader.readUInt32LE()));
        const getPropsRopsReader = new BufferReader(getPropsRopsList);
        getPropsRopsReader.readUInt8();
        getPropsRopsReader.readUInt8();
        expect(getPropsRopsReader.readUInt32LE()).toBe(0);
        getPropsRopsReader.readUInt8(); // PropertyRow Flags
        expect(readPropertyValue(getPropsRopsReader, PropertyType.PtypString)).toBe("Read Me");

        const openStreamRops = new BufferWriter();
        openStreamRops.writeUInt8(0x2b);
        openStreamRops.writeUInt8(0); // LogonId
        openStreamRops.writeUInt8(3); // InputHandleIndex (message)
        openStreamRops.writeUInt8(4); // OutputHandleIndex
        writePropertyTag(openStreamRops, { propertyId: 0x1000, propertyType: PropertyType.PtypString });
        openStreamRops.writeUInt8(0); // OpenModeFlags
        const openStreamResult = await execute(cookie, encodeRopBuffer({ ropsList: openStreamRops.toBuffer(), handleTable: [0xffffffff, 0xffffffff] }));
        const openStreamReader = new BufferReader(openStreamResult.body);
        openStreamReader.readUInt32LE();
        openStreamReader.readUInt32LE();
        openStreamReader.readUInt32LE();
        const { ropsList: openStreamRopsList } = decodeRopBuffer(openStreamReader.readBytes(openStreamReader.readUInt32LE()));
        const openStreamRopsReader = new BufferReader(openStreamRopsList);
        openStreamRopsReader.readUInt8();
        openStreamRopsReader.readUInt8();
        expect(openStreamRopsReader.readUInt32LE()).toBe(0);
        const streamSize = openStreamRopsReader.readUInt32LE();
        expect(streamSize).toBeGreaterThan(0);

        const readStreamRops = new BufferWriter();
        readStreamRops.writeUInt8(0x2c);
        readStreamRops.writeUInt8(0); // LogonId
        readStreamRops.writeUInt8(4); // InputHandleIndex (stream)
        readStreamRops.writeUInt16LE(1000);
        const readStreamResult = await execute(cookie, encodeRopBuffer({ ropsList: readStreamRops.toBuffer(), handleTable: [0xffffffff] }));
        const readStreamReader = new BufferReader(readStreamResult.body);
        readStreamReader.readUInt32LE();
        readStreamReader.readUInt32LE();
        readStreamReader.readUInt32LE();
        const { ropsList: readStreamRopsList } = decodeRopBuffer(readStreamReader.readBytes(readStreamReader.readUInt32LE()));
        const readStreamRopsReader = new BufferReader(readStreamRopsList);
        readStreamRopsReader.readUInt8();
        readStreamRopsReader.readUInt8();
        expect(readStreamRopsReader.readUInt32LE()).toBe(0);
        const dataSize = readStreamRopsReader.readUInt16LE();
        expect(dataSize).toBe(streamSize);
        expect(readPropertyValue(readStreamRopsReader, PropertyType.PtypString)).toBe("Actual body text.");
    });

    it("Resolves a real message's labelUids to display names via PidNameKeywords, end to end.", async () => {
        const PSETID_PUBLIC_STRINGS = "00020329-0000-0000-c000-000000000046";
        const mailbox = await createMailbox(owner.uid);
        const inbox = await createFolder(mailbox.uid, FolderType.INBOX, "Inbox");
        const important = await createLabel(mailbox.uid, { name: "Important" });
        const work = await createLabel(mailbox.uid, { name: "Work" });
        await createMessage(mailbox.uid, inbox.uid, { subject: "Categorized", labelUids: [important.uid, work.uid] });
        const connectResult = await connect();
        const cookie = cookieHeaderFrom(connectResult.headers["set-cookie"]);

        const logonRops = new BufferWriter();
        logonRops.writeUInt8(0xfe);
        logonRops.writeUInt8(0);
        logonRops.writeUInt8(0);
        logonRops.writeUInt8(0x01);
        logonRops.writeUInt32LE(0);
        logonRops.writeUInt32LE(0);
        logonRops.writeUInt16LE(0);
        const logonResult = await execute(cookie, encodeRopBuffer({ ropsList: logonRops.toBuffer(), handleTable: [0xffffffff] }));
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

        const openFolderRops = new BufferWriter();
        openFolderRops.writeUInt8(0x02);
        openFolderRops.writeUInt8(0);
        openFolderRops.writeUInt8(0);
        openFolderRops.writeUInt8(1);
        openFolderRops.writeUInt8(0);
        openFolderRops.writeBigUInt64LE(inboxFid);
        await execute(cookie, encodeRopBuffer({ ropsList: openFolderRops.toBuffer(), handleTable: [0xffffffff, 0xffffffff] }));

        const tableRops = new BufferWriter();
        tableRops.writeUInt8(0x05);
        tableRops.writeUInt8(0);
        tableRops.writeUInt8(1);
        tableRops.writeUInt8(2);
        tableRops.writeUInt8(0);
        await execute(cookie, encodeRopBuffer({ ropsList: tableRops.toBuffer(), handleTable: [0xffffffff, 0xffffffff, 0xffffffff] }));

        const setColumnsRops = new BufferWriter();
        setColumnsRops.writeUInt8(0x12);
        setColumnsRops.writeUInt8(0);
        setColumnsRops.writeUInt8(2);
        setColumnsRops.writeUInt8(0);
        setColumnsRops.writeUInt16LE(2);
        writePropertyTag(setColumnsRops, { propertyId: 0x0037, propertyType: PropertyType.PtypString });
        writePropertyTag(setColumnsRops, { propertyId: 0x674a, propertyType: PropertyType.PtypInteger64 });
        await execute(cookie, encodeRopBuffer({ ropsList: setColumnsRops.toBuffer(), handleTable: [0xffffffff] }));

        const queryRowsRops = new BufferWriter();
        queryRowsRops.writeUInt8(0x15);
        queryRowsRops.writeUInt8(0);
        queryRowsRops.writeUInt8(2);
        queryRowsRops.writeUInt8(0);
        queryRowsRops.writeUInt8(1);
        queryRowsRops.writeUInt16LE(10);
        const queryRowsResult = await execute(cookie, encodeRopBuffer({ ropsList: queryRowsRops.toBuffer(), handleTable: [0xffffffff] }));
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
        rowsReader.readUInt16LE();
        rowsReader.readUInt8();
        readPropertyValue(rowsReader, PropertyType.PtypString);
        const mid = readPropertyValue(rowsReader, PropertyType.PtypInteger64) as bigint;

        const openMessageRops = new BufferWriter();
        openMessageRops.writeUInt8(0x03);
        openMessageRops.writeUInt8(0);
        openMessageRops.writeUInt8(1);
        openMessageRops.writeUInt8(3);
        openMessageRops.writeUInt16LE(0);
        openMessageRops.writeBigUInt64LE(inboxFid);
        openMessageRops.writeUInt8(0);
        openMessageRops.writeBigUInt64LE(mid);
        await execute(cookie, encodeRopBuffer({ ropsList: openMessageRops.toBuffer(), handleTable: [0xffffffff, 0xffffffff] }));

        // RopGetPropertyIdsFromNames, Kind=Name ("Keywords") - NameSize is inclusive of the trailing UTF-16LE
        // null terminator, per RopGetPropertyIdsFromNamesHandler.readPropertyName's own documented convention.
        const nameBytes = Buffer.concat([Buffer.from("Keywords", "utf16le"), Buffer.from([0x00, 0x00])]);
        const namesRops = new BufferWriter();
        namesRops.writeUInt8(0x56);
        namesRops.writeUInt8(0);
        namesRops.writeUInt8(0);
        namesRops.writeUInt8(0x02);
        namesRops.writeUInt16LE(1);
        namesRops.writeUInt8(0x01); // Kind - Name
        namesRops.writeBytes(encodeGuid(PSETID_PUBLIC_STRINGS));
        namesRops.writeUInt8(nameBytes.length);
        namesRops.writeBytes(nameBytes);
        const namesResult = await execute(cookie, encodeRopBuffer({ ropsList: namesRops.toBuffer(), handleTable: [0xffffffff] }));
        const namesReader = new BufferReader(namesResult.body);
        namesReader.readUInt32LE();
        namesReader.readUInt32LE();
        namesReader.readUInt32LE();
        const { ropsList: namesRopsList } = decodeRopBuffer(namesReader.readBytes(namesReader.readUInt32LE()));
        const namesRopsReader = new BufferReader(namesRopsList);
        namesRopsReader.readUInt8();
        namesRopsReader.readUInt8();
        expect(namesRopsReader.readUInt32LE()).toBe(0);
        expect(namesRopsReader.readUInt16LE()).toBe(1);
        const keywordsId = namesRopsReader.readUInt16LE();
        expect(keywordsId).toBeGreaterThanOrEqual(0x8000);

        const getPropsRops = new BufferWriter();
        getPropsRops.writeUInt8(0x07);
        getPropsRops.writeUInt8(0);
        getPropsRops.writeUInt8(3);
        getPropsRops.writeUInt16LE(0);
        getPropsRops.writeUInt16LE(0);
        getPropsRops.writeUInt16LE(1);
        writePropertyTag(getPropsRops, { propertyId: keywordsId, propertyType: PropertyType.PtypMultipleString });
        const getPropsResult = await execute(cookie, encodeRopBuffer({ ropsList: getPropsRops.toBuffer(), handleTable: [0xffffffff] }));
        const getPropsReader = new BufferReader(getPropsResult.body);
        getPropsReader.readUInt32LE();
        getPropsReader.readUInt32LE();
        getPropsReader.readUInt32LE();
        const { ropsList: getPropsRopsList } = decodeRopBuffer(getPropsReader.readBytes(getPropsReader.readUInt32LE()));
        const getPropsRopsReader = new BufferReader(getPropsRopsList);
        getPropsRopsReader.readUInt8();
        getPropsRopsReader.readUInt8();
        expect(getPropsRopsReader.readUInt32LE()).toBe(0);
        getPropsRopsReader.readUInt8();
        const categories = readPropertyValue(getPropsRopsReader, PropertyType.PtypMultipleString) as string[];
        expect(categories.sort()).toEqual(["Important", "Work"]);
    });

    it("Composes and sends a real message end to end, saving a Sent Items copy.", async () => {
        mailTransport().sent = [];
        const mailbox = await createMailbox(owner.uid);
        const connectResult = await connect();
        const cookie = cookieHeaderFrom(connectResult.headers["set-cookie"]);

        const logonRops = new BufferWriter();
        logonRops.writeUInt8(0xfe);
        logonRops.writeUInt8(0);
        logonRops.writeUInt8(0);
        logonRops.writeUInt8(0x01);
        logonRops.writeUInt32LE(0);
        logonRops.writeUInt32LE(0);
        logonRops.writeUInt16LE(0);
        const logonResult = await execute(cookie, encodeRopBuffer({ ropsList: logonRops.toBuffer(), handleTable: [0xffffffff] }));
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

        const createMessageRops = new BufferWriter();
        createMessageRops.writeUInt8(0x06);
        createMessageRops.writeUInt8(0); // LogonId
        createMessageRops.writeUInt8(0); // InputHandleIndex (logon)
        createMessageRops.writeUInt8(3); // OutputHandleIndex
        createMessageRops.writeUInt16LE(0); // CodePageId
        createMessageRops.writeBigUInt64LE(inboxFid);
        createMessageRops.writeUInt8(0); // AssociatedFlag
        const createMessageResult = await execute(cookie, encodeRopBuffer({ ropsList: createMessageRops.toBuffer(), handleTable: [0xffffffff, 0xffffffff] }));
        const createMessageReader = new BufferReader(createMessageResult.body);
        createMessageReader.readUInt32LE();
        createMessageReader.readUInt32LE();
        createMessageReader.readUInt32LE();
        const { ropsList: createMessageRopsList } = decodeRopBuffer(createMessageReader.readBytes(createMessageReader.readUInt32LE()));
        const createMessageRopsReader = new BufferReader(createMessageRopsList);
        createMessageRopsReader.readUInt8();
        createMessageRopsReader.readUInt8();
        expect(createMessageRopsReader.readUInt32LE()).toBe(0); // ReturnValue - success

        const propertyValues = new BufferWriter();
        writeTaggedPropertyValue(propertyValues, { propertyId: 0x0037, propertyType: PropertyType.PtypString, value: "Hello From MAPI" });
        writeTaggedPropertyValue(propertyValues, { propertyId: 0x0e04, propertyType: PropertyType.PtypString, value: "recipient@example.com" });
        const propertyValuesBytes = propertyValues.toBuffer();
        const setPropertiesRops = new BufferWriter();
        setPropertiesRops.writeUInt8(0x0a);
        setPropertiesRops.writeUInt8(0); // LogonId
        setPropertiesRops.writeUInt8(3); // InputHandleIndex (message)
        setPropertiesRops.writeUInt16LE(2 + propertyValuesBytes.length);
        setPropertiesRops.writeUInt16LE(2); // PropertyValueCount
        setPropertiesRops.writeBytes(propertyValuesBytes);
        const setPropertiesResult = await execute(cookie, encodeRopBuffer({ ropsList: setPropertiesRops.toBuffer(), handleTable: [0xffffffff] }));
        expect(setPropertiesResult.headers["x-responsecode"]).toBe("0");

        const openStreamRops = new BufferWriter();
        openStreamRops.writeUInt8(0x2b);
        openStreamRops.writeUInt8(0); // LogonId
        openStreamRops.writeUInt8(3); // InputHandleIndex (message)
        openStreamRops.writeUInt8(4); // OutputHandleIndex
        writePropertyTag(openStreamRops, { propertyId: 0x1000, propertyType: PropertyType.PtypString });
        openStreamRops.writeUInt8(0x02); // OpenModeFlags - Create
        await execute(cookie, encodeRopBuffer({ ropsList: openStreamRops.toBuffer(), handleTable: [0xffffffff, 0xffffffff] }));

        const bodyText = "This message was composed entirely via MAPI ROPs.";
        const bodyBytes = Buffer.concat([Buffer.from(bodyText, "utf16le"), Buffer.from([0, 0])]);
        const writeStreamRops = new BufferWriter();
        writeStreamRops.writeUInt8(0x2d);
        writeStreamRops.writeUInt8(0); // LogonId
        writeStreamRops.writeUInt8(4); // InputHandleIndex (stream)
        writeStreamRops.writeUInt16LE(bodyBytes.length);
        writeStreamRops.writeBytes(bodyBytes);
        const writeStreamResult = await execute(cookie, encodeRopBuffer({ ropsList: writeStreamRops.toBuffer(), handleTable: [0xffffffff] }));
        const writeStreamReader = new BufferReader(writeStreamResult.body);
        writeStreamReader.readUInt32LE();
        writeStreamReader.readUInt32LE();
        writeStreamReader.readUInt32LE();
        const { ropsList: writeStreamRopsList } = decodeRopBuffer(writeStreamReader.readBytes(writeStreamReader.readUInt32LE()));
        const writeStreamRopsReader = new BufferReader(writeStreamRopsList);
        writeStreamRopsReader.readUInt8();
        writeStreamRopsReader.readUInt8();
        expect(writeStreamRopsReader.readUInt32LE()).toBe(0);
        expect(writeStreamRopsReader.readUInt16LE()).toBe(bodyBytes.length);

        const saveChangesRops = new BufferWriter();
        saveChangesRops.writeUInt8(0x0c);
        saveChangesRops.writeUInt8(0); // LogonId
        saveChangesRops.writeUInt8(7); // ResponseHandleIndex
        saveChangesRops.writeUInt8(3); // InputHandleIndex (message)
        saveChangesRops.writeUInt8(0); // SaveFlags
        const saveChangesResult = await execute(cookie, encodeRopBuffer({ ropsList: saveChangesRops.toBuffer(), handleTable: [0xffffffff] }));
        const saveChangesReader = new BufferReader(saveChangesResult.body);
        saveChangesReader.readUInt32LE();
        saveChangesReader.readUInt32LE();
        saveChangesReader.readUInt32LE();
        const { ropsList: saveChangesRopsList } = decodeRopBuffer(saveChangesReader.readBytes(saveChangesReader.readUInt32LE()));
        const saveChangesRopsReader = new BufferReader(saveChangesRopsList);
        saveChangesRopsReader.readUInt8();
        saveChangesRopsReader.readUInt8();
        expect(saveChangesRopsReader.readUInt32LE()).toBe(0);

        const submitRops = new BufferWriter();
        submitRops.writeUInt8(0x32);
        submitRops.writeUInt8(0); // LogonId
        submitRops.writeUInt8(3); // InputHandleIndex (message)
        submitRops.writeUInt8(0); // SubmitFlags
        const submitResult = await execute(cookie, encodeRopBuffer({ ropsList: submitRops.toBuffer(), handleTable: [0xffffffff] }));
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

        const sentMessage = await messageRepo.findOne({ where: { subject: "Hello From MAPI" } });
        expect(sentMessage).not.toBeNull();
        expect(sentMessage?.recipients).toEqual([{ address: "recipient@example.com", type: RecipientType.TO }]);
        const savedRaw = await blobStore().get(sentMessage!.bodyBlobKey);
        expect(savedRaw.toString("utf-8")).toContain(bodyText);
    });

    it("Invites a meeting's attendees only for revisions the client submitted, once each (restapi's MeetingSchedulingJob).", async () => {
        mailTransport().sent = [];
        await createMailbox(owner.uid);
        const cookie = cookieHeaderFrom((await connect()).headers["set-cookie"]);
        const job: any = await objectFactory.newInstance(MeetingSchedulingJobSQL, { name: "MeetingSchedulingJobSQL" });
        const attendee = `${uuid.v4()}@example.com`;
        const invites = () => mailTransport().sent.filter((message) => message.envelopeTo.includes(attendee)).length;
        const run = async (rops: Buffer): Promise<number> => {
            const result = await execute(cookie, encodeRopBuffer({ ropsList: rops, handleTable: [0xffffffff, 0xffffffff] }));
            const reader = new BufferReader(result.body);
            reader.readBytes(12);
            return decodeRopBuffer(reader.readBytes(reader.readUInt32LE())).ropsList.readUInt32LE(2);
        };
        const logon = new BufferWriter().writeUInt8(0xfe).writeUInt8(0).writeUInt8(0).writeUInt8(0x01).writeUInt32LE(0).writeUInt32LE(0).writeUInt16LE(0).toBuffer();
        const create = new BufferWriter().writeUInt8(0x06).writeUInt8(0).writeUInt8(0).writeUInt8(3).writeUInt16LE(0).writeBigUInt64LE(5n).writeUInt8(0).toBuffer();
        const values = new BufferWriter();
        writeTaggedPropertyValue(values, { propertyId: 0x001a, propertyType: PropertyType.PtypString, value: "IPM.Appointment" });
        writeTaggedPropertyValue(values, { propertyId: 0x0037, propertyType: PropertyType.PtypString, value: "Planning" });
        writeTaggedPropertyValue(values, { propertyId: 0x0e04, propertyType: PropertyType.PtypString, value: attendee });
        const valueBytes = values.toBuffer();
        const setProperties = new BufferWriter().writeUInt8(0x0a).writeUInt8(0).writeUInt8(3).writeUInt16LE(2 + valueBytes.length).writeUInt16LE(3).writeBytes(valueBytes).toBuffer();
        const save = new BufferWriter().writeUInt8(0x0c).writeUInt8(0).writeUInt8(7).writeUInt8(3).writeUInt8(0).toBuffer();
        const submit = new BufferWriter().writeUInt8(0x32).writeUInt8(0).writeUInt8(3).writeUInt8(0).toBuffer();

        await execute(cookie, encodeRopBuffer({ ropsList: logon, handleTable: [0xffffffff] }));
        expect(await run(create)).toBe(0);
        expect(await run(setProperties)).toBe(0);
        expect(await run(save)).toBe(0);
        await job.run();
        expect(invites()).toBe(0); // saved, never sent

        expect(await run(submit)).toBe(0);
        await job.run();
        await job.run();
        expect(invites()).toBe(1);

        expect(await run(save)).toBe(0); // saved again without sending
        await job.run();
        expect(invites()).toBe(1);
    });

    it("Resolves named properties into stable, session-scoped numeric IDs starting at 0x8000 (RopGetPropertyIdsFromNames).", async () => {
        const PSETID_APPOINTMENT = "00062002-0000-0000-c000-000000000046";
        await createMailbox(owner.uid);
        const connectResult = await connect();
        const cookie = cookieHeaderFrom(connectResult.headers["set-cookie"]);

        const logonRops = new BufferWriter();
        logonRops.writeUInt8(0xfe);
        logonRops.writeUInt8(0);
        logonRops.writeUInt8(0);
        logonRops.writeUInt8(0x01);
        logonRops.writeUInt32LE(0);
        logonRops.writeUInt32LE(0);
        logonRops.writeUInt16LE(0);
        await execute(cookie, encodeRopBuffer({ ropsList: logonRops.toBuffer(), handleTable: [0xffffffff] }));

        const buildRequest = function (lids: number[]): Buffer {
            const namesWriter = new BufferWriter();
            for (const lid of lids) {
                namesWriter.writeUInt8(0x00); // Kind - LID
                namesWriter.writeBytes(encodeGuid(PSETID_APPOINTMENT));
                namesWriter.writeUInt32LE(lid);
            }
            const namesBytes = namesWriter.toBuffer();
            const writer = new BufferWriter();
            writer.writeUInt8(0x56);
            writer.writeUInt8(0); // LogonId
            writer.writeUInt8(0); // InputHandleIndex (logon)
            writer.writeUInt8(0x02); // Flags - assign new IDs
            writer.writeUInt16LE(lids.length);
            writer.writeBytes(namesBytes);
            return writer.toBuffer();
        };

        // Arbitrary LIDs within PSETID_Appointment - this test only exercises the ROP's own mechanics
        // (resolving a PropertyName to a stable session-scoped numeric ID), not any specific real Calendar
        // property's exact LID value.
        const result = await execute(cookie, encodeRopBuffer({ ropsList: buildRequest([0x8208, 0x820d]), handleTable: [0xffffffff] }));
        const reader = new BufferReader(result.body);
        reader.readUInt32LE();
        reader.readUInt32LE();
        reader.readUInt32LE();
        const { ropsList } = decodeRopBuffer(reader.readBytes(reader.readUInt32LE()));
        const ropsReader = new BufferReader(ropsList);
        expect(ropsReader.readUInt8()).toBe(0x56);
        expect(ropsReader.readUInt8()).toBe(0);
        expect(ropsReader.readUInt32LE()).toBe(0);
        expect(ropsReader.readUInt16LE()).toBe(2);
        const id1 = ropsReader.readUInt16LE();
        const id2 = ropsReader.readUInt16LE();
        expect(id1).toBe(0x8000);
        expect(id2).toBe(0x8001);

        const secondResult = await execute(cookie, encodeRopBuffer({ ropsList: buildRequest([0x8208]), handleTable: [0xffffffff] }));
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

    describe("Zero-access writes (RopSubmitMessage, RopSaveChangesMessage)", () => {
        /** One ROP request over `cookie`'s session, answering the first ROP's ReturnValue. */
        const run = async (cookie: string, rops: Buffer): Promise<number> => {
            const result = await execute(cookie, encodeRopBuffer({ ropsList: rops, handleTable: [0xffffffff, 0xffffffff] }));
            const reader = new BufferReader(result.body);
            reader.readBytes(12);
            return decodeRopBuffer(reader.readBytes(reader.readUInt32LE())).ropsList.readUInt32LE(2);
        };
        const logon = (): Buffer => Buffer.from([0xfe, 0, 0, 0x01, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]);
        const createMessage = (folderId: bigint): Buffer => {
            const writer = new BufferWriter();
            writer.writeUInt8(0x06);
            writer.writeUInt8(0); // LogonId
            writer.writeUInt8(0); // InputHandleIndex
            writer.writeUInt8(3); // OutputHandleIndex
            writer.writeUInt16LE(0); // CodePageId
            writer.writeBigUInt64LE(folderId);
            writer.writeUInt8(0); // AssociatedFlag
            return writer.toBuffer();
        };
        const setProperties = (values: { propertyId: number; value: string }[]): Buffer => {
            const valuesWriter = new BufferWriter();
            for (const { propertyId, value } of values) {
                writeTaggedPropertyValue(valuesWriter, { propertyId, propertyType: PropertyType.PtypString, value });
            }
            const valuesBytes = valuesWriter.toBuffer();
            const writer = new BufferWriter();
            writer.writeUInt8(0x0a);
            writer.writeUInt8(0); // LogonId
            writer.writeUInt8(3); // InputHandleIndex
            writer.writeUInt16LE(2 + valuesBytes.length);
            writer.writeUInt16LE(values.length);
            writer.writeBytes(valuesBytes);
            return writer.toBuffer();
        };
        const submit = (): Buffer => Buffer.from([0x32, 0, 3, 0]);

        it("Seals the Sent Items copy of a message Outlook sends when zero-access covers the mailbox's mail, storing nothing readable.", async () => {
            mailTransport().sent = [];
            const canary = `CANARY-${uuid.v4()}`;
            const mailbox: any = await createMailbox(owner.uid, { zeroAccessKeys: [await testZeroAccessKey()], zeroAccessTypes: { mail: true } });
            const cookie = cookieHeaderFrom((await connect()).headers["set-cookie"]);

            await execute(cookie, encodeRopBuffer({ ropsList: logon(), handleTable: [0xffffffff] }));
            expect(await run(cookie, createMessage(5n))).toBe(0);
            expect(
                await run(
                    cookie,
                    setProperties([
                        { propertyId: 0x0037, value: `${canary} subject` },
                        { propertyId: 0x0e04, value: "recipient@example.com" },
                        { propertyId: 0x1000, value: `${canary} body` },
                    ]),
                ),
            ).toBe(0);
            expect(await run(cookie, submit())).toBe(0);

            expect(mailTransport().sent).toHaveLength(1);
            expect(mailTransport().sent[0].raw.toString("utf-8")).toContain(`${canary} body`);
            const copies: any[] = await messagesOf(mailbox.uid);
            expect(copies).toHaveLength(1);
            expect(copies[0].zeroAccess).toBeTruthy();
            expect(copies[0].subject).toBe("[...]");
            expect(copies[0].bodyPreview).toBe("");
            expect(JSON.stringify(copies[0])).not.toContain(canary);
            expect(await blobStore().get(copies[0].bodyBlobKey)).toBeTruthy();
            expect([...(blobStore() as any).store.values()].some((blob: Buffer) => blob.toString("utf-8").includes(canary))).toBe(false);
        });

        it("Logs nothing an Outlook write carries: no ROP bytes, request or response body, or session cookie, only sizes and codes.", async () => {
            mailTransport().sent = [];
            const canary = `CANARY-${uuid.v4()}`;
            await createMailbox(owner.uid, { zeroAccessKeys: [await testZeroAccessKey()], zeroAccessTypes: { mail: true } });
            const route: any = objectFactory.getInstance(MapiEmsmdbRoute);
            const lines: string[] = [];
            const record = (...args: any[]): void => {
                lines.push(args.map((arg) => (typeof arg === "string" ? arg : inspect(arg, { depth: 6, maxArrayLength: null, maxStringLength: null }))).join(" "));
            };
            const spies: any[] = [];
            for (const target of new Set([route.logger, logger, console])) {
                for (const level of ["error", "warn", "info", "debug", "verbose", "log"]) {
                    if (typeof target?.[level] === "function") {
                        spies.push(vi.spyOn(target, level).mockImplementation(record as any));
                    }
                }
            }
            try {
                const cookie = cookieHeaderFrom((await connect()).headers["set-cookie"]);
                await execute(cookie, encodeRopBuffer({ ropsList: logon(), handleTable: [0xffffffff] }));
                expect(await run(cookie, createMessage(5n))).toBe(0);
                const subject = setProperties([
                    { propertyId: 0x0037, value: `${canary} subject` },
                    { propertyId: 0x0e04, value: "recipient@example.com" },
                    { propertyId: 0x1000, value: `${canary} body` },
                ]);
                expect(await run(cookie, subject)).toBe(0);
                // A ROP whose handler throws (here a failed query holding what it wrote) is logged by name and message.
                route.ropHandlers.set(0x77, {
                    ropId: 0x77,
                    handle: async (reader: any) => {
                        reader.readBytes(1 + Buffer.byteLength(canary));
                        throw Object.assign(new Error("query failed"), { parameters: [`${canary} subject`] });
                    },
                });
                expect(await run(cookie, Buffer.concat([Buffer.from([0x77, 0]), Buffer.from(canary)]))).toBe(0x80004005);
                expect(await run(cookie, submit())).toBe(0);
                // An Execute without a session answers a failure code.
                expect((await execute("", encodeRopBuffer({ ropsList: logon(), handleTable: [0xffffffff] }))).headers["x-responsecode"]).toBe("10");

                const sessionUid: string = /MapiContext=([^;]+)/.exec(cookie)![1];
                const all: string = lines.join("\n");
                expect(all).toContain("MAPI Execute IN ropBytes=");
                expect(all).toContain("MAPI Execute OUT ropBytes=");
                expect(all).toContain("MAPI Execute failed responseCode=10");
                expect(all).toContain("RopDispatcher: ROP 0x77 (Object) failed: Error: query failed");
                for (const secret of [canary, Buffer.from(canary, "utf16le").toString("hex"), Buffer.from(canary).toString("hex"), sessionUid]) {
                    expect(all).not.toContain(secret);
                }
                expect(all).not.toMatch(/Hex=|MAPI_DEBUG/);
            } finally {
                route.ropHandlers.delete(0x77);
                spies.forEach((spy) => spy.mockRestore());
            }
        });

        it("Leaves no sealed blob, index entry or copy when sealing the Sent Items copy fails writing its index entry.", async () => {
            mailTransport().sent = [];
            const canary = `CANARY-${uuid.v4()}`;
            const mailbox: any = await createMailbox(owner.uid, { zeroAccessKeys: [await testZeroAccessKey()], zeroAccessTypes: { mail: true } });
            const cookie = cookieHeaderFrom((await connect()).headers["set-cookie"]);

            await execute(cookie, encodeRopBuffer({ ropsList: logon(), handleTable: [0xffffffff] }));
            expect(await run(cookie, createMessage(5n))).toBe(0);
            expect(
                await run(
                    cookie,
                    setProperties([
                        { propertyId: 0x0037, value: `${canary} subject` },
                        { propertyId: 0x0e04, value: "recipient@example.com" },
                        { propertyId: 0x1000, value: `${canary} body` },
                    ]),
                ),
            ).toBe(0);
            const create = RepoUtils.prototype.create;
            const createSpy = vi.spyOn(RepoUtils.prototype, "create").mockImplementation(async function (this: any, ...args: any[]) {
                // The index entry sealRaw() writes after storing the sealed blob.
                if (args[0]?.keyFingerprint && args[0]?.sealed) {
                    throw new Error("index down");
                }
                return await (create as any).apply(this, args);
            });
            try {
                expect(await run(cookie, submit())).not.toBe(0);
                expect(createSpy.mock.calls.some((call: any[]) => call[0]?.keyFingerprint && call[0]?.sealed)).toBe(true);
            } finally {
                createSpy.mockRestore();
            }

            expect(mailTransport().sent).toHaveLength(1);
            // The copy's uid was fresh to the send, so restapi's seal() removes the blob it stored under it.
            expect([...(blobStore() as any).store.keys()].some((key: string) => key.startsWith(`za/${mailbox.uid}/`))).toBe(false);
            expect([...(blobStore() as any).store.values()].some((blob: Buffer) => blob.toString("utf-8").includes(canary))).toBe(false);
            expect(await (objectFactory.getInstance(ConnectionManager)!.connections.get("sql") as any).getRepository(ZeroAccessIndexEntrySQL).count({ where: { mailboxUid: mailbox.uid } })).toBe(0);
            expect((await messagesOf(mailbox.uid)).filter((message: any) => message.zeroAccess)).toHaveLength(0);
        });

        it("Keeps a protected appointment's sealed location when Outlook saves it again with the empty location it was shown.", async () => {
            const canary = `CANARY-${uuid.v4()}`;
            const mailbox: any = await createMailbox(owner.uid, { zeroAccessKeys: [await testZeroAccessKey()], zeroAccessTypes: { calendar: true } });
            const cookie = cookieHeaderFrom((await connect()).headers["set-cookie"]);
            const save = (): Buffer => Buffer.from([0x0c, 0, 1, 3, 0]);
            const eventsOf = async (): Promise<any[]> => (await (objectFactory.getInstance(ConnectionManager)?.connections.get("sql") as any).getRepository(CalendarEventSQL).find({ where: { mailboxUid: mailbox.uid } })) as any[];

            await execute(cookie, encodeRopBuffer({ ropsList: logon(), handleTable: [0xffffffff] }));
            // RopGetPropertyIdsFromNames: PidLidLocation's session-scoped ID.
            const names = new BufferWriter();
            names.writeUInt8(0x56);
            names.writeUInt8(0); // LogonId
            names.writeUInt8(0); // InputHandleIndex
            names.writeUInt8(0x02); // Flags - assign new IDs
            names.writeUInt16LE(1);
            names.writeUInt8(0x00); // Kind - LID
            names.writeBytes(encodeGuid("00062002-0000-0000-c000-000000000046"));
            names.writeUInt32LE(0x8208);
            const namesResult = new BufferReader((await execute(cookie, encodeRopBuffer({ ropsList: names.toBuffer(), handleTable: [0xffffffff] }))).body);
            namesResult.readBytes(12);
            const locationId = decodeRopBuffer(namesResult.readBytes(namesResult.readUInt32LE())).ropsList.readUInt16LE(8);

            expect(await run(cookie, createMessage(5n))).toBe(0);
            expect(
                await run(
                    cookie,
                    setProperties([
                        { propertyId: 0x001a, value: "IPM.Appointment" },
                        { propertyId: 0x0037, value: "Checkup" },
                        { propertyId: locationId, value: `${canary} clinic` },
                    ]),
                ),
            ).toBe(0);
            expect(await run(cookie, save())).toBe(0);
            const [created] = await eventsOf();
            expect(created.title).toBe("Protected event");
            expect(created.location ?? null).toBeNull();
            expect(created.sealedFields.location).toEqual(expect.any(String));
            expect(JSON.stringify(created)).not.toContain(canary);

            // Outlook re-sends the placeholders it was shown: the title, and "" for the location.
            expect(await run(cookie, setProperties([{ propertyId: 0x0037, value: "Protected event" }, { propertyId: locationId, value: "" }]))).toBe(0);
            expect(await run(cookie, save())).toBe(0);
            const events = await eventsOf();
            expect(events).toHaveLength(1);
            expect(events[0].sealedFields.location).toBe(created.sealedFields.location);
            expect(events[0].sealedFields.title).toBe(created.sealedFields.title);
            expect(events[0].location ?? null).toBeNull();
        });

        it("Leaves an event made from a zero-access invitation (sealedContent) to its own sealing when Outlook saves it: the placeholders it sends back are kept and a new title or location is refused.", async () => {
            const mailbox: any = await createMailbox(owner.uid, { zeroAccessKeys: [await testZeroAccessKey()], zeroAccessTypes: { mail: true, calendar: true } });
            const cookie = cookieHeaderFrom((await connect()).headers["set-cookie"]);
            const save = (): Buffer => Buffer.from([0x0c, 0, 1, 3, 0]);
            const repo = (): any => (objectFactory.getInstance(ConnectionManager)?.connections.get("sql") as any).getRepository(CalendarEventSQL);
            const eventsOf = async (): Promise<any[]> => (await repo().find({ where: { mailboxUid: mailbox.uid } })) as any[];
            const fileAsInvitation = async (fields: Record<string, any>): Promise<void> => {
                await repo().update({ mailboxUid: mailbox.uid }, fields);
            };
            await execute(cookie, encodeRopBuffer({ ropsList: logon(), handleTable: [0xffffffff] }));
            // RopGetPropertyIdsFromNames: PidLidLocation's session-scoped ID.
            const names = new BufferWriter();
            names.writeUInt8(0x56);
            names.writeUInt8(0); // LogonId
            names.writeUInt8(0); // InputHandleIndex
            names.writeUInt8(0x02); // Flags - assign new IDs
            names.writeUInt16LE(1);
            names.writeUInt8(0x00); // Kind - LID
            names.writeBytes(encodeGuid("00062002-0000-0000-c000-000000000046"));
            names.writeUInt32LE(0x8208);
            const namesResult = new BufferReader((await execute(cookie, encodeRopBuffer({ ropsList: names.toBuffer(), handleTable: [0xffffffff] }))).body);
            namesResult.readBytes(12);
            const locationId = decodeRopBuffer(namesResult.readBytes(namesResult.readUInt32LE())).ropsList.readUInt16LE(8);

            expect(await run(cookie, createMessage(5n))).toBe(0);
            expect(await run(cookie, setProperties([{ propertyId: 0x001a, value: "IPM.Appointment" }, { propertyId: 0x0037, value: "Invitation" }]))).toBe(0);
            expect(await run(cookie, save())).toBe(0);
            // As ScanQueueJob files an invitation that reached a zero-access mailbox: its content sealed whole, placeholders in the columns.
            const sealedContent = { keyFingerprint: "fp", sealed: "sealed-content" };
            await fileAsInvitation({ title: "[...]", location: null, sealedFields: null, sealedContent });
            const [filed] = await eventsOf();

            // Outlook re-sends the placeholders it was shown: the title, and "" for the location.
            expect(await run(cookie, setProperties([{ propertyId: 0x0037, value: "[...]" }, { propertyId: locationId, value: "" }]))).toBe(0);
            expect(await run(cookie, save())).toBe(0);
            const [kept] = await eventsOf();
            expect(kept.version).toBeGreaterThan(filed.version);
            expect(kept.title).toBe("[...]");
            expect(kept.location ?? null).toBeNull();
            expect(kept.sealedFields ?? null).toBeNull();
            expect(kept.sealedContent).toEqual(sealedContent);

            // A new title, or a new location, can't be sealed into sealedContent: refused (MAPI_E_NO_ACCESS), nothing saved.
            expect(await run(cookie, setProperties([{ propertyId: 0x0037, value: "Q3 review (moved)" }, { propertyId: locationId, value: "" }]))).toBe(0);
            expect(await run(cookie, save())).toBe(0x80070005);
            expect(await run(cookie, setProperties([{ propertyId: 0x0037, value: "[...]" }, { propertyId: locationId, value: "Room 9" }]))).toBe(0);
            expect(await run(cookie, save())).toBe(0x80070005);
            const [unchanged] = await eventsOf();
            expect(unchanged.version).toBe(kept.version);
            expect(unchanged.title).toBe("[...]");
            expect(unchanged.sealedFields ?? null).toBeNull();
        });
    });
});
