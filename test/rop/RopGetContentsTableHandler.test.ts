///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { BufferReader, BufferWriter } from "../../src/codec/BufferCursor.js";
import { resolveContentsWindow } from "../../src/rop/ContentsTable.js";
import { RopGetContentsTableHandler } from "../../src/rop/RopGetContentsTableHandler.js";
import type { RopContext } from "../../src/rop/RopHandler.js";
import { MapiSessionContext } from "../../src/MapiSessionManager.js";

function buildRequest({ logonId = 0, inputHandleIndex = 5, outputHandleIndex = 6, tableFlags = 0 }): Buffer {
    const writer = new BufferWriter();
    writer.writeUInt8(logonId);
    writer.writeUInt8(inputHandleIndex);
    writer.writeUInt8(outputHandleIndex);
    writer.writeUInt8(tableFlags);
    return writer.toBuffer();
}

function makeContext(folderEntityUid: string, overrides: Partial<RopContext> = {}): RopContext {
    const session = new MapiSessionContext({ mailboxUid: "mailbox-1", userUid: "user-1" });
    session.handles[5] = { type: "folder", entityUid: folderEntityUid };
    return {
        mailboxUid: "mailbox-1",
        userUid: "user-1",
        session,
        folderRepo: {} as any,
        messageRepo: {} as any,
        calendarEventRepo: {} as any,
        mailboxRepo: {} as any,
        folderClass: {} as any,
        messageClass: {} as any,
        calendarEventClass: {} as any,
        scanPipeline: {} as any,
        mailTransport: {} as any,
        blobStore: {} as any,
        ...overrides,
    };
}

/** A repo whose `find` returns `total` rows `<prefix>0..`, honoring `page`/`limit` like the real backends do. */
function pagedRepo(prefix: string, total: number) {
    return {
        find: vi.fn().mockImplementation((_query: any, options: { page: number; limit: number }) => {
            const start = options.page * options.limit;
            return Promise.resolve(
                Array.from({ length: Math.max(0, Math.min(options.limit, total - start)) }, (_, i) => ({ uid: `${prefix}${start + i}` })),
            );
        }),
    };
}

describe("RopGetContentsTableHandler Tests", () => {
    it("Has RopId 0x05.", () => {
        expect(new RopGetContentsTableHandler().ropId).toBe(0x05);
    });

    it("Creates a row-less message table handle for an ordinary folder, without listing its messages up front.", async () => {
        const messageRepo = { find: vi.fn() };
        const folderRepo = { findOne: vi.fn().mockResolvedValue({ uid: "top1", type: "inbox" }) };
        const context = makeContext("folder:top1", { folderRepo: folderRepo as any, messageRepo: messageRepo as any });

        const writer = new BufferWriter();
        await new RopGetContentsTableHandler().handle(new BufferReader(buildRequest({})), writer, context);

        const response = new BufferReader(writer.toBuffer());
        expect(response.readUInt8()).toBe(0x05);
        expect(response.readUInt8()).toBe(6);
        expect(response.readUInt32LE()).toBe(0);
        expect(response.hasMore()).toBe(false);

        expect(context.session.handles[6]).toEqual({ type: "table", entityUid: "folder:top1", contentsKind: "message", cursor: 0, generation: expect.any(String) });
        expect(folderRepo.findOne).toHaveBeenCalledWith("top1", { ignoreACL: true });
        expect(messageRepo.find).not.toHaveBeenCalled();
    });

    it.each([
        ["calendar", "calendarEvent"],
        ["contacts", "contact"],
        ["suggested_contacts", "contact"],
        ["tasks", "task"],
    ])("Records a %s folder's contents as %s rows.", async (folderType, contentsKind) => {
        const folderRepo = { findOne: vi.fn().mockResolvedValue({ uid: "f1", type: folderType }) };
        const context = makeContext("folder:f1", { folderRepo: folderRepo as any });

        await new RopGetContentsTableHandler().handle(new BufferReader(buildRequest({})), new BufferWriter(), context);

        expect(context.session.handles[6]?.contentsKind).toBe(contentsKind);
        expect(context.session.handles[6]?.rows).toBeUndefined();
    });

    it("Returns an empty table for a virtual folder, without querying any repo.", async () => {
        const messageRepo = { find: vi.fn() };
        const context = makeContext("virtual:root", { messageRepo: messageRepo as any });

        await new RopGetContentsTableHandler().handle(new BufferReader(buildRequest({})), new BufferWriter(), context);

        expect(context.session.handles[6]?.rows).toEqual([]);
        expect(context.session.handles[6]?.contentsKind).toBeUndefined();
        expect(messageRepo.find).not.toHaveBeenCalled();
    });

    it("Returns MAPI_E_INVALID_OBJECT when InputHandleIndex isn't a folder handle.", async () => {
        const context = makeContext("folder:f1");
        const writer = new BufferWriter();

        await new RopGetContentsTableHandler().handle(new BufferReader(buildRequest({ inputHandleIndex: 99 })), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0x80070005);
        expect(context.session.handles[6]).toBeUndefined();
    });
});

describe("ContentsTable.resolveContentsWindow Tests", () => {
    it("Reads a window from the message repo scoped to the folder and mailbox, in a stable newest-first order.", async () => {
        const messageRepo = pagedRepo("msg", 3);
        const context = makeContext("folder:top1", { messageRepo: messageRepo as any });

        const rows = await resolveContentsWindow(context, { type: "table", entityUid: "folder:top1", contentsKind: "message" }, 1, 5);

        expect(rows).toEqual(["message:msg1", "message:msg2"]);
        expect(messageRepo.find).toHaveBeenCalledWith(
            { folderUid: "top1", mailboxUid: "mailbox-1", sort: { receivedDate: "DESC", uid: "ASC" }, limit: 1000, page: 0 },
            { ignoreACL: true, limit: 1000, page: 0 },
        );
    });

    it("Pages past the repo's 100-row default and across its 1000-row page boundary.", async () => {
        const messageRepo = pagedRepo("msg", 2500);
        const context = makeContext("folder:top1", { messageRepo: messageRepo as any });
        const table = { type: "table" as const, entityUid: "folder:top1", contentsKind: "message" as const };

        const rows = await resolveContentsWindow(context, table, 990, 20);

        expect(rows.length).toBe(20);
        expect(rows[0]).toBe("message:msg990");
        expect(rows[19]).toBe("message:msg1009");
        expect(messageRepo.find).toHaveBeenCalledTimes(2);
        expect(await resolveContentsWindow(context, table, 2490, 50)).toHaveLength(10);
    });

    it.each([
        ["calendarEvent", "calendarEventRepo", { startDate: "DESC", uid: "ASC" }],
        ["contact", "contactRepo", { displayName: "ASC", uid: "ASC" }],
        ["task", "taskRepo", { title: "ASC", uid: "ASC" }],
    ])("Reads %s rows from %s with its own sort.", async (kind, repoName, sort) => {
        const repo = pagedRepo("x", 2);
        const overrides: Partial<RopContext> = { [repoName]: repo };
        const context = makeContext("folder:f1", overrides);
        const contentsKind = kind as "calendarEvent" | "contact" | "task";

        const rows = await resolveContentsWindow(context, { type: "table", entityUid: "folder:f1", contentsKind }, 0, 10);

        expect(rows).toEqual([`${kind}:x0`, `${kind}:x1`]);
        expect(repo.find).toHaveBeenCalledWith(expect.objectContaining({ sort }), expect.anything());
    });

    it("Yields no rows for a contacts or tasks table when the context has no such repo.", async () => {
        const context = makeContext("folder:f1");
        expect(await resolveContentsWindow(context, { type: "table", entityUid: "folder:f1", contentsKind: "contact" }, 0, 10)).toEqual([]);
        expect(await resolveContentsWindow(context, { type: "table", entityUid: "folder:f1", contentsKind: "task" }, 0, 10)).toEqual([]);
    });
});
