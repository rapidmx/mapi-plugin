///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { BufferReader, BufferWriter } from "../../src/codec/BufferCursor.js";
import { PropertyType, readPropertyValue, writePropertyTag } from "../../src/codec/PropertyValue.js";
import { RopGetPropertiesSpecificHandler } from "../../src/rop/RopGetPropertiesSpecificHandler.js";
import type { RopContext } from "../../src/rop/RopHandler.js";
import { MapiSessionContext } from "../../src/MapiSessionManager.js";

function buildRequest({
    logonId = 0,
    inputHandleIndex = 5,
    propertySizeLimit = 0,
    wantUnicode = 0,
    tags = [] as { propertyId: number; propertyType: PropertyType }[],
}): Buffer {
    const writer = new BufferWriter();
    writer.writeUInt8(logonId);
    writer.writeUInt8(inputHandleIndex);
    writer.writeUInt16LE(propertySizeLimit);
    writer.writeUInt16LE(wantUnicode);
    writer.writeUInt16LE(tags.length);
    for (const tag of tags) {
        writePropertyTag(writer, tag);
    }
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

describe("RopGetPropertiesSpecificHandler Tests", () => {
    it("Has RopId 0x07.", () => {
        expect(new RopGetPropertiesSpecificHandler().ropId).toBe(0x07);
    });

    it("Returns MAPI_E_INVALID_OBJECT when InputHandleIndex isn't a folder or message handle.", async () => {
        const context = makeContext();
        const handler = new RopGetPropertiesSpecificHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({})), writer, context);

        const response = new BufferReader(writer.toBuffer());
        expect(response.readUInt8()).toBe(0x07);
        expect(response.readUInt8()).toBe(5);
        expect(response.readUInt32LE()).toBe(0x80070005);
        expect(response.hasMore()).toBe(false);
    });

    it("Fetches Subject/MessageFlags for an open message handle.", async () => {
        const context = makeContext();
        context.session.handles[5] = { type: "message", entityUid: "message:m1" };
        context.messageRepo = {
            findOne: vi.fn().mockResolvedValue({ uid: "m1", subject: "Hello", flags: { read: true } }),
        } as any;
        const handler = new RopGetPropertiesSpecificHandler();
        const writer = new BufferWriter();

        const tags = [
            { propertyId: 0x0037, propertyType: PropertyType.PtypString },
            { propertyId: 0x0e07, propertyType: PropertyType.PtypInteger32 },
        ];
        await handler.handle(new BufferReader(buildRequest({ tags })), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8(); // RopId
        response.readUInt8(); // InputHandleIndex
        expect(response.readUInt32LE()).toBe(0); // ReturnValue
        response.readUInt8(); // PropertyRow Flags
        expect(readPropertyValue(response, PropertyType.PtypString)).toBe("Hello");
        expect(readPropertyValue(response, PropertyType.PtypInteger32)).toBe(1);
        expect(response.hasMore()).toBe(false);
    });

    it("Falls back to a type-appropriate default, not MAPI_E_CALL_FAILED, when a tag's requested PropertyType doesn't match what its propertyId actually resolves to.", async () => {
        // The same round-3 review repro as RopQueryRowsHandler's own mirrored test: PidTagSubject requested as
        // PtypGuid. messageValueFor resolves it to a plain string regardless of the client's requested type, and
        // encodeGuid() throws on a non-GUID string - previously that escaped this handler entirely and failed
        // the whole ROP with MAPI_E_CALL_FAILED instead of degrading just this one property.
        const context = makeContext();
        context.session.handles[5] = { type: "message", entityUid: "message:m1" };
        context.messageRepo = {
            findOne: vi.fn().mockResolvedValue({ uid: "m1", subject: "Not a GUID", flags: { read: true } }),
        } as any;
        const handler = new RopGetPropertiesSpecificHandler();
        const writer = new BufferWriter();

        const tags = [{ propertyId: 0x0037, propertyType: PropertyType.PtypGuid }]; // PidTagSubject requested as PtypGuid
        await handler.handle(new BufferReader(buildRequest({ tags })), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0); // ReturnValue - success, not MAPI_E_CALL_FAILED
        response.readUInt8(); // PropertyRow Flags
        expect(readPropertyValue(response, PropertyType.PtypGuid)).toBe("00000000-0000-0000-0000-000000000000");
        expect(response.hasMore()).toBe(false);
    });

    it("Fetches DisplayName for an open folder handle.", async () => {
        const context = makeContext();
        context.session.handles[5] = { type: "folder", entityUid: "folder:f1" };
        context.folderRepo = {
            findOne: vi.fn().mockResolvedValue({ uid: "f1", name: "Inbox" }),
            find: vi.fn().mockResolvedValue([]),
        } as any;
        const handler = new RopGetPropertiesSpecificHandler();
        const writer = new BufferWriter();

        const tags = [{ propertyId: 0x3001, propertyType: PropertyType.PtypString }];
        await handler.handle(new BufferReader(buildRequest({ tags })), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0);
        response.readUInt8();
        expect(readPropertyValue(response, PropertyType.PtypString)).toBe("Inbox");
    });
});
