///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
// See the identical file header in test/routes/mongo/MapiNspiRoute.test.ts for the full rationale - this
// verifies the same BaseMapiNspiRoute transport plumbing on the SQL-backed variant.
import config from "../../config.sql.js";
import { Server, ObjectFactory, ConnectionManager, isSqlDataSource, ACLAction, AccessControlListSQL } from "@rapidrest/service-core";
import { JWTUtils, Logger } from "@rapidrest/core";
import * as uuid from "uuid";
import { Repository } from "typeorm";
import { ContactSQL, MailboxSQL } from "@rapidmx/restapi/sql";
import { ContactAddressKind } from "@rapidmx/restapi";
import { registerTestDoubles } from "../../testDoubles.js";
import { mapiRequest } from "../../mapiTestClient.js";
import { BufferReader, BufferWriter } from "../../../src/codec/BufferCursor.js";
import { decodeGuid } from "../../../src/codec/MapiGuid.js";
import { PropertyType, readPropertyValue } from "../../../src/codec/PropertyValue.js";

describe("Route:MapiNspiRouteSQL Tests", () => {
    const logger = Logger();
    const objectFactory: ObjectFactory = new ObjectFactory(config, logger);
    const server: Server = new Server({ config, basePath: "./test/server-sql", logger, objectFactory });
    const baseUrl = "/sql/mapi/nspi";
    let mailboxRepo: Repository<MailboxSQL>;
    let contactRepo: Repository<ContactSQL>;
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

    const createContact = async function (mailboxUid: string, data?: Partial<ContactSQL>): Promise<ContactSQL> {
        return await contactRepo.save(
            new ContactSQL({
                mailboxUid,
                displayName: "Jane Doe",
                emails: [{ address: "jane@example.com", type: ContactAddressKind.WORK }],
                phones: [],
                addresses: [],
                ...data,
            } as any),
        );
    };

    const bind = async function (headers: Record<string, string> = {}) {
        return mapiRequest(
            server.getApplication(),
            baseUrl,
            {
                Authorization: "jwt " + ownerToken,
                "X-RequestType": "Bind",
                "Content-Type": "application/mapi-http",
                ...headers,
            },
            (() => {
                const body = new BufferWriter();
                body.writeUInt32LE(0); // Flags
                body.writeUInt8(0); // HasState
                body.writeUInt32LE(0); // AuxiliaryBufferSize
                return body.toBuffer();
            })(),
        );
    };

    const getMatches = async function (options: { searchTerm?: string; rowCount?: number } = {}) {
        const body = new BufferWriter();
        body.writeUInt32LE(0); // Reserved
        body.writeUInt8(0); // HasState
        body.writeUInt8(0); // HasMinimalIds
        body.writeUInt32LE(0); // InterfaceOptionFlags
        if (options.searchTerm !== undefined) {
            body.writeUInt8(1); // HasFilter
            body.writeUInt8(0x03); // RestrictType - ContentRestriction
            body.writeUInt16LE(0x0001); // FuzzyLevelLow - FL_SUBSTRING
            body.writeUInt16LE(0x0001); // FuzzyLevelHigh - FL_IGNORECASE
            body.writeUInt16LE(PropertyType.PtypString); // PropertyTag (restriction's own target column): PropertyType
            body.writeUInt16LE(0x3001); // PropertyTag: PropertyId (PidTagDisplayName)
            body.writeUInt16LE(PropertyType.PtypString); // TaggedValue's own embedded PropertyTag: PropertyType
            body.writeUInt16LE(0x3001); // TaggedValue's own embedded PropertyTag: PropertyId
            body.writeNullTerminatedUtf16LE(options.searchTerm);
        } else {
            body.writeUInt8(0); // HasFilter
        }
        body.writeUInt8(0); // HasPropertyName
        body.writeUInt32LE(options.rowCount ?? 50); // RowCount
        body.writeUInt8(0); // HasColumns - use the pragmatic default column set
        body.writeUInt32LE(0); // AuxiliaryBufferSize

        return mapiRequest(
            server.getApplication(),
            baseUrl,
            {
                Authorization: "jwt " + ownerToken,
                "X-RequestType": "GetMatches",
                "Content-Type": "application/mapi-http",
            },
            body.toBuffer(),
        );
    };

    beforeAll(async () => {
        registerTestDoubles(objectFactory);
        await server.start();

        const connMgr: ConnectionManager | undefined = objectFactory.getInstance(ConnectionManager);
        const conn: any = connMgr?.connections.get("sql");
        if (isSqlDataSource(conn)) {
            mailboxRepo = conn.getRepository(MailboxSQL);
            contactRepo = conn.getRepository(ContactSQL);
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
        await contactRepo.clear();
        await aclRepo.clear();
    });

    describe("Bind", () => {
        it("Requires authentication.", async () => {
            const result = await mapiRequest(
                server.getApplication(),
                baseUrl,
                { "X-RequestType": "Bind", "Content-Type": "application/mapi-http" },
                Buffer.alloc(0),
            );
            expect(result.status).toBe(401);
        });

        it("Reports a nonzero StatusCode when the caller owns no mailbox.", async () => {
            const result = await bind();
            expect(result.status).toBe(200);
            const reader = new BufferReader(result.body);
            expect(reader.readUInt32LE()).not.toBe(0); // StatusCode
        });

        it("Succeeds, setting an NspiContext cookie and returning a real ServerGuid.", async () => {
            await createMailbox(owner.uid);
            const result = await bind();
            expect(result.status).toBe(200);
            expect(String(result.headers["set-cookie"])).toContain("NspiContext=");

            const reader = new BufferReader(result.body);
            expect(reader.readUInt32LE()).toBe(0); // StatusCode
            expect(reader.readUInt32LE()).toBe(0); // ErrorCode
            const serverGuid = decodeGuid(reader);
            expect(serverGuid).toMatch(/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/);
            expect(reader.readUInt32LE()).toBe(0); // AuxiliaryBufferSize
            expect(reader.hasMore()).toBe(false);
        });

        it("Never sends the framework's own CSRF cookie, only NspiContext.", async () => {
            await createMailbox(owner.uid);
            const result = await bind();
            expect(result.status).toBe(200);
            expect(String(result.headers["set-cookie"])).not.toContain("csrf=");
        });

        it("Echoes the caller's own X-RequestId back on the response, per [MS-OXCMAPIHTTP]'s common response format.", async () => {
            await createMailbox(owner.uid);
            const result = await bind({ "X-RequestId": "{11111111-2222-3333-4444-555555555555}:1" });
            expect(result.status).toBe(200);
            expect(result.headers["x-requestid"]).toBe("{11111111-2222-3333-4444-555555555555}:1");
        });
    });

    describe("Unbind", () => {
        it("Always succeeds.", async () => {
            const result = await mapiRequest(
                server.getApplication(),
                baseUrl,
                { Authorization: "jwt " + ownerToken, "X-RequestType": "Unbind", "Content-Type": "application/mapi-http" },
                (() => {
                    const body = new BufferWriter();
                    body.writeUInt32LE(0);
                    body.writeUInt32LE(0);
                    return body.toBuffer();
                })(),
            );
            expect(result.status).toBe(200);
            const reader = new BufferReader(result.body);
            expect(reader.readUInt32LE()).toBe(0); // StatusCode
            expect(reader.readUInt32LE()).toBe(0); // ErrorCode
            expect(reader.readUInt32LE()).toBe(0); // AuxiliaryBufferSize
        });
    });

    describe("GetMatches", () => {
        it("Returns 404 when the caller owns no mailbox.", async () => {
            const result = await getMatches();
            expect(result.status).toBe(404);
        });

        it("Returns every contact (default columns: DisplayName/EmailAddress) when no filter is given.", async () => {
            const mailbox = await createMailbox(owner.uid);
            await createContact(mailbox.uid, { displayName: "Jane Doe", emails: [{ address: "jane@example.com", type: ContactAddressKind.WORK }] });
            await createContact(mailbox.uid, { displayName: "John Smith", emails: [{ address: "john@example.com", type: ContactAddressKind.WORK }] });

            const result = await getMatches();
            expect(result.status).toBe(200);
            const reader = new BufferReader(result.body);
            expect(reader.readUInt32LE()).toBe(0); // StatusCode
            expect(reader.readUInt32LE()).toBe(0); // ErrorCode
            expect(reader.readUInt8()).toBe(1); // HasState
            reader.readBytes(36); // State (STAT) - not asserted field-by-field, see NspiCodec.ts
            expect(reader.readUInt8()).toBe(1); // HasMinimalIds
            const minimalIdCount = reader.readUInt32LE();
            expect(minimalIdCount).toBe(2);
            for (let i = 0; i < minimalIdCount; i++) {
                reader.readUInt32LE();
            }
            expect(reader.readUInt8()).toBe(1); // HasColumnsAndRows
            const columnCount = reader.readUInt32LE();
            expect(columnCount).toBe(2);
            const columns = [];
            for (let i = 0; i < columnCount; i++) {
                columns.push({ propertyType: reader.readUInt16LE(), propertyId: reader.readUInt16LE() });
            }
            const rowCount = reader.readUInt32LE();
            expect(rowCount).toBe(2);
            const displayNames: string[] = [];
            for (let i = 0; i < rowCount; i++) {
                reader.readUInt8(); // Flags
                for (const column of columns) {
                    reader.readUInt8(); // HasValue - PtypString always has a value in this pragmatic subset
                    const value = readPropertyValue(reader, column.propertyType);
                    if (column.propertyId === 0x3001) {
                        displayNames.push(value as string);
                    }
                }
            }
            expect(displayNames.sort()).toEqual(["Jane Doe", "John Smith"]);
        });

        it("Filters by a ContentRestriction search term, matching only Jane.", async () => {
            const mailbox = await createMailbox(owner.uid);
            await createContact(mailbox.uid, { displayName: "Jane Doe" });
            await createContact(mailbox.uid, { displayName: "John Smith" });

            const result = await getMatches({ searchTerm: "jane" });
            const reader = new BufferReader(result.body);
            reader.readUInt32LE(); // StatusCode
            reader.readUInt32LE(); // ErrorCode
            reader.readUInt8(); // HasState
            reader.readBytes(36); // State
            reader.readUInt8(); // HasMinimalIds
            const minimalIdCount = reader.readUInt32LE();
            expect(minimalIdCount).toBe(1);
        });

        it("Limits results to RowCount even when more matches exist.", async () => {
            const mailbox = await createMailbox(owner.uid);
            await createContact(mailbox.uid, { displayName: "Jane Doe" });
            await createContact(mailbox.uid, { displayName: "John Smith" });

            const result = await getMatches({ rowCount: 1 });
            const reader = new BufferReader(result.body);
            reader.readUInt32LE();
            reader.readUInt32LE();
            reader.readUInt8();
            reader.readBytes(36);
            reader.readUInt8();
            const minimalIdCount = reader.readUInt32LE();
            expect(minimalIdCount).toBe(1); // capped by RowCount, even though 2 contacts exist
        });
    });
});
