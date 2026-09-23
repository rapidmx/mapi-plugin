///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { BufferReader, BufferWriter } from "../../src/codec/BufferCursor.js";
import { PropertyType, readPropertyValue } from "../../src/codec/PropertyValue.js";
import { RopQueryRowsHandler } from "../../src/rop/RopQueryRowsHandler.js";
import type { RopContext } from "../../src/rop/RopHandler.js";
import { MapiSessionContext } from "../../src/MapiSessionManager.js";

function buildRequest({ logonId = 0, inputHandleIndex = 7, queryRowsFlags = 0, forwardRead = 1, rowCount = 10 }): Buffer {
    const writer = new BufferWriter();
    writer.writeUInt8(logonId);
    writer.writeUInt8(inputHandleIndex);
    writer.writeUInt8(queryRowsFlags);
    writer.writeUInt8(forwardRead);
    writer.writeUInt16LE(rowCount);
    return writer.toBuffer();
}

function makeContext(folderRepo: any, messageRepo: any = {}): RopContext {
    const session = new MapiSessionContext({ mailboxUid: "mailbox-1", userUid: "user-1" });
    return { mailboxUid: "mailbox-1", userUid: "user-1", session, folderRepo, messageRepo, mailboxRepo: {} as any, folderClass: {} as any, messageClass: {} as any, scanPipeline: {} as any, mailTransport: {} as any, blobStore: {} as any };
}

describe("RopQueryRowsHandler Tests", () => {
    it("Has RopId 0x15.", () => {
        expect(new RopQueryRowsHandler().ropId).toBe(0x15);
    });

    it("Returns MAPI_E_INVALID_OBJECT when InputHandleIndex isn't a table handle.", async () => {
        const context = makeContext({});
        const handler = new RopQueryRowsHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({})), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0x80070005);
    });

    it("Returns a well-formed empty response for a table with no rows (rows/cursor/columns all defaulted).", async () => {
        const context = makeContext({});
        context.session.handles[7] = { type: "table", entityUid: "virtual:root" };
        const handler = new RopQueryRowsHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({})), writer, context);

        const response = new BufferReader(writer.toBuffer());
        expect(response.readUInt8()).toBe(0x15);
        expect(response.readUInt8()).toBe(7);
        expect(response.readUInt32LE()).toBe(0);
        expect(response.readUInt8()).toBe(0x02); // Origin - BOOKMARK_END
        expect(response.readUInt16LE()).toBe(0); // RowCount
        expect(response.hasMore()).toBe(false);
    });

    it("Defaults cursor/columns to empty when a table handle was constructed without them.", async () => {
        const context = makeContext({ findOne: vi.fn().mockResolvedValue(undefined), find: vi.fn().mockResolvedValue([]) });
        context.session.handles[7] = { type: "table", entityUid: "virtual:root", rows: ["virtual:root"] };
        const handler = new RopQueryRowsHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({})), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0);
        response.readUInt8();
        expect(response.readUInt16LE()).toBe(1); // RowCount - the one row this table was constructed with
        response.readUInt8(); // PropertyRow Flags - no columns configured, so the row has no property values
        expect(response.hasMore()).toBe(false);
        expect(context.session.handles[7]?.cursor).toBe(1);
    });

    it("Fetches the mailbox's full folder list at most once across many rows when Subfolders is a requested column, not once per row.", async () => {
        const folderRepo = {
            findOne: vi.fn().mockImplementation(async (uid: string) => ({ uid, name: uid, unreadCount: 0, totalCount: 0 })),
            find: vi.fn().mockResolvedValue([]),
        };
        const context = makeContext(folderRepo);
        const rows = Array.from({ length: 5 }, (_, i) => `folder:f${i}`);
        context.session.handles[7] = {
            type: "table",
            entityUid: "virtual:root",
            rows,
            columns: [{ propertyId: 0x360a, propertyType: PropertyType.PtypBoolean }], // PidTagSubfolders
        };
        const handler = new RopQueryRowsHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({ rowCount: 10 })), writer, context);

        expect(folderRepo.findOne).toHaveBeenCalledTimes(5); // once per distinct folder, as expected
        expect(folderRepo.find).toHaveBeenCalledTimes(1); // the whole-mailbox list for hasChildren, shared across all 5 rows
    });

    it("Builds rows with DisplayName/FolderId/ContentCount/UnreadCount/Subfolders columns and advances the cursor.", async () => {
        const folderRepo = {
            findOne: vi.fn().mockResolvedValue({ uid: "f1", name: "Inbox", unreadCount: 3, totalCount: 10 }),
            find: vi.fn().mockResolvedValue([{ uid: "child", parentFolderUid: "f1", mailboxUid: "mailbox-1" }]),
        };
        const context = makeContext(folderRepo);
        context.session.handles[7] = {
            type: "table",
            entityUid: "folder:top",
            rows: ["folder:f1"],
            cursor: 0,
            columns: [
                { propertyId: 0x3001, propertyType: PropertyType.PtypString },
                { propertyId: 0x6748, propertyType: PropertyType.PtypInteger64 },
                { propertyId: 0x3602, propertyType: PropertyType.PtypInteger32 },
                { propertyId: 0x3603, propertyType: PropertyType.PtypInteger32 },
                { propertyId: 0x360a, propertyType: PropertyType.PtypBoolean },
            ],
        };
        const handler = new RopQueryRowsHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({})), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8(); // RopId
        response.readUInt8(); // InputHandleIndex
        expect(response.readUInt32LE()).toBe(0); // ReturnValue
        response.readUInt8(); // Origin
        expect(response.readUInt16LE()).toBe(1); // RowCount

        expect(response.readUInt8()).toBe(0x00); // PropertyRow Flags - StandardPropertyRow
        expect(readPropertyValue(response, PropertyType.PtypString)).toBe("Inbox");
        const fid = readPropertyValue(response, PropertyType.PtypInteger64) as bigint;
        expect(fid).toBeGreaterThan(0n);
        expect(readPropertyValue(response, PropertyType.PtypInteger32)).toBe(10); // ContentCount
        expect(readPropertyValue(response, PropertyType.PtypInteger32)).toBe(3); // UnreadCount
        expect(readPropertyValue(response, PropertyType.PtypBoolean)).toBe(true); // Subfolders
        expect(response.hasMore()).toBe(false);

        expect(context.session.handles[7]?.cursor).toBe(1);
        // The FID assigned to resolve PidTagFolderId is now remembered for a later RopOpenFolder.
        expect(Object.values(context.session.folderIds)).toContain("folder:f1");
    });

    // Every PropertyType this codec supports (PropertyValue.ts) must have a type-appropriate default for an
    // unsupported property column, so an arbitrary client column request never crashes writePropertyValue()
    // regardless of which type it names - not just the one or two types this handler's own supported PidTags
    // happen to use.
    const allPropertyTypes = Object.values(PropertyType).filter((v) => typeof v === "number") as PropertyType[];

    it.each(allPropertyTypes)("Falls back to a valid default value for an unsupported property of type 0x%s", async (propertyType) => {
        const folderRepo = {
            findOne: vi.fn().mockResolvedValue({ uid: "f1", name: "Inbox", unreadCount: 0, totalCount: 0 }),
            find: vi.fn().mockResolvedValue([]),
        };
        const context = makeContext(folderRepo);
        context.session.handles[7] = {
            type: "table",
            entityUid: "folder:top",
            rows: ["folder:f1"],
            cursor: 0,
            columns: [{ propertyId: 0x9999, propertyType }],
        };
        const handler = new RopQueryRowsHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({})), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        response.readUInt32LE();
        response.readUInt8();
        response.readUInt16LE();
        response.readUInt8(); // PropertyRow Flags
        // Decoding without throwing is the real assertion - confirms writePropertyValue() accepted the
        // default as a genuinely valid value for this type, not just "some" value.
        expect(() => readPropertyValue(response, propertyType)).not.toThrow();
    });

    it.each(allPropertyTypes)("Falls back to a valid default value for an unsupported message property of type 0x%s", async (propertyType) => {
        const messageRepo = {
            findOne: vi.fn().mockResolvedValue({ uid: "m1", subject: "Hi", flags: { read: false }, hasAttachments: false, receivedDate: new Date() }),
        };
        const context = makeContext({}, messageRepo);
        context.session.handles[7] = {
            type: "table",
            entityUid: "folder:top",
            rows: ["message:m1"],
            cursor: 0,
            columns: [{ propertyId: 0x9999, propertyType }],
        };
        const handler = new RopQueryRowsHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({})), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        response.readUInt32LE();
        response.readUInt8();
        response.readUInt16LE();
        response.readUInt8(); // PropertyRow Flags
        expect(() => readPropertyValue(response, propertyType)).not.toThrow();
    });

    it("Builds message rows with Subject/MessageFlags/HasAttachments/DeliveryTime columns.", async () => {
        const receivedDate = new Date("2026-03-15T09:00:00.000Z");
        const messageRepo = {
            findOne: vi.fn().mockResolvedValue({
                uid: "m1",
                subject: "Hello",
                flags: { read: true },
                hasAttachments: true,
                receivedDate,
            }),
        };
        const context = makeContext({}, messageRepo);
        context.session.handles[7] = {
            type: "table",
            entityUid: "folder:top",
            rows: ["message:m1"],
            cursor: 0,
            columns: [
                { propertyId: 0x0037, propertyType: PropertyType.PtypString },
                { propertyId: 0x0e07, propertyType: PropertyType.PtypInteger32 },
                { propertyId: 0x0e1b, propertyType: PropertyType.PtypBoolean },
                { propertyId: 0x0e06, propertyType: PropertyType.PtypTime },
            ],
        };
        const handler = new RopQueryRowsHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({})), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0);
        response.readUInt8();
        expect(response.readUInt16LE()).toBe(1);

        response.readUInt8(); // PropertyRow Flags
        expect(readPropertyValue(response, PropertyType.PtypString)).toBe("Hello");
        expect(readPropertyValue(response, PropertyType.PtypInteger32)).toBe(1); // MSGFLAG_READ
        expect(readPropertyValue(response, PropertyType.PtypBoolean)).toBe(true);
        expect(readPropertyValue(response, PropertyType.PtypTime)).toEqual(receivedDate);
        expect(messageRepo.findOne).toHaveBeenCalledWith("m1", { ignoreACL: true });
    });

    it("Resolves PidTagMid to a session-assigned MID, remembered for a later RopOpenMessage.", async () => {
        const messageRepo = {
            findOne: vi.fn().mockResolvedValue({ uid: "m1", subject: "Hi", flags: { read: false }, hasAttachments: false, receivedDate: new Date() }),
        };
        const context = makeContext({}, messageRepo);
        context.session.handles[7] = {
            type: "table",
            entityUid: "folder:top",
            rows: ["message:m1"],
            cursor: 0,
            columns: [{ propertyId: 0x674a, propertyType: PropertyType.PtypInteger64 }],
        };
        const handler = new RopQueryRowsHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({})), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        response.readUInt32LE();
        response.readUInt8();
        response.readUInt16LE();
        response.readUInt8();
        const mid = readPropertyValue(response, PropertyType.PtypInteger64) as bigint;
        expect(mid).toBeGreaterThan(0n);
        expect(Object.values(context.session.messageIds)).toContain("message:m1");
    });

    it("Reports MessageFlags 0 for an unread message.", async () => {
        const messageRepo = {
            findOne: vi.fn().mockResolvedValue({ uid: "m1", subject: "Hi", flags: { read: false }, hasAttachments: false, receivedDate: new Date() }),
        };
        const context = makeContext({}, messageRepo);
        context.session.handles[7] = {
            type: "table",
            entityUid: "folder:top",
            rows: ["message:m1"],
            cursor: 0,
            columns: [{ propertyId: 0x0e07, propertyType: PropertyType.PtypInteger32 }],
        };
        const handler = new RopQueryRowsHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({})), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        response.readUInt32LE();
        response.readUInt8();
        response.readUInt16LE();
        response.readUInt8();
        expect(readPropertyValue(response, PropertyType.PtypInteger32)).toBe(0);
    });

    it("Degrades to empty values for a message row whose real Message has since vanished.", async () => {
        const messageRepo = { findOne: vi.fn().mockResolvedValue(undefined) };
        const context = makeContext({}, messageRepo);
        context.session.handles[7] = {
            type: "table",
            entityUid: "folder:top",
            rows: ["message:gone"],
            cursor: 0,
            columns: [{ propertyId: 0x0037, propertyType: PropertyType.PtypString }],
        };
        const handler = new RopQueryRowsHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({})), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        response.readUInt32LE();
        response.readUInt8();
        response.readUInt16LE();
        response.readUInt8();
        expect(readPropertyValue(response, PropertyType.PtypString)).toBe("");
    });

    it("Falls back to a type-appropriate default, not MAPI_E_CALL_FAILED, when a column's requested PropertyType doesn't match what its propertyId actually resolves to.", async () => {
        // The repro a round-3 review gave: SetColumns([{propertyId: 0x0037 (PidTagSubject), propertyType: 0x0048
        // (PtypGuid)}]) then QueryRows against a real message - messageValueFor's own switch resolves
        // PidTagSubject to a plain string regardless of the client's requested type, and encodeGuid() throws on
        // a non-GUID string. Before writePropertyValueSafely, that escaped RopQueryRowsHandler entirely and
        // RopDispatcher failed the whole ROP with MAPI_E_CALL_FAILED, discarding every row already built.
        const messageRepo = {
            findOne: vi.fn().mockResolvedValue({ uid: "m1", subject: "Not a GUID", flags: { read: false }, hasAttachments: false, receivedDate: new Date() }),
        };
        const context = makeContext({}, messageRepo);
        context.session.handles[7] = {
            type: "table",
            entityUid: "folder:top",
            rows: ["message:m1"],
            cursor: 0,
            columns: [{ propertyId: 0x0037, propertyType: PropertyType.PtypGuid }], // PidTagSubject requested as PtypGuid
        };
        const handler = new RopQueryRowsHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({})), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0); // ReturnValue - success, not MAPI_E_CALL_FAILED
        response.readUInt8();
        expect(response.readUInt16LE()).toBe(1); // RowCount - the row was still built and returned
        response.readUInt8(); // PropertyRow Flags
        expect(readPropertyValue(response, PropertyType.PtypGuid)).toBe("00000000-0000-0000-0000-000000000000");
        expect(response.hasMore()).toBe(false);
    });

    it("Respects the requested RowCount, leaving remaining rows for a subsequent call.", async () => {
        const folderRepo = {
            findOne: vi.fn().mockImplementation(async (uid: string) => ({ uid, name: uid, unreadCount: 0, totalCount: 0 })),
            find: vi.fn().mockResolvedValue([]),
        };
        const context = makeContext(folderRepo);
        context.session.handles[7] = {
            type: "table",
            entityUid: "folder:top",
            rows: ["folder:a", "folder:b", "folder:c"],
            cursor: 0,
            columns: [],
        };
        const handler = new RopQueryRowsHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({ rowCount: 2 })), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        response.readUInt32LE();
        response.readUInt8();
        expect(response.readUInt16LE()).toBe(2);
        expect(context.session.handles[7]?.cursor).toBe(2);
    });
});
