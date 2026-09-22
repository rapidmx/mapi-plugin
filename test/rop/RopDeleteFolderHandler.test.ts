///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { BufferReader, BufferWriter } from "../../src/codec/BufferCursor.js";
import { MAX_FOLDER_DEPTH, MAX_FOLDERS_PER_DELETE, MAX_ITEMS_PER_DELETE, RopDeleteFolderHandler } from "../../src/rop/RopDeleteFolderHandler.js";
import { AuditAction } from "@rapidmx/restapi";
import type { RopContext } from "../../src/rop/RopHandler.js";
import { MapiSessionContext } from "../../src/MapiSessionManager.js";

const DEL_MESSAGES = 0x01;
const DEL_FOLDERS = 0x04;
const DELETE_HARD_DELETE = 0x10;

function buildRequest({ logonId = 0, inputHandleIndex = 5, deleteFolderFlags = 0, folderId = 1n }): Buffer {
    const writer = new BufferWriter();
    writer.writeUInt8(logonId);
    writer.writeUInt8(inputHandleIndex);
    writer.writeUInt8(deleteFolderFlags);
    writer.writeBigUInt64LE(folderId);
    return writer.toBuffer();
}

function makeContext(overrides: Partial<RopContext> = {}): RopContext {
    return {
        mailboxUid: "mailbox-1",
        userUid: "user-1",
        session: new MapiSessionContext({ mailboxUid: "mailbox-1", userUid: "user-1" }),
        folderRepo: { find: vi.fn().mockResolvedValue([]), delete: vi.fn().mockResolvedValue(undefined) } as any,
        messageRepo: { find: vi.fn().mockResolvedValue([]), delete: vi.fn().mockResolvedValue(undefined) } as any,
        calendarEventRepo: { find: vi.fn().mockResolvedValue([]), delete: vi.fn().mockResolvedValue(undefined) } as any,
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

describe("RopDeleteFolderHandler Tests", () => {
    it("Has RopId 0x1D.", () => {
        expect(new RopDeleteFolderHandler().ropId).toBe(0x1d);
    });

    it("Returns MAPI_E_INVALID_OBJECT when InputHandleIndex isn't a folder handle.", async () => {
        const context = makeContext();
        const handler = new RopDeleteFolderHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({})), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0x80070005);
    });

    it("Returns MAPI_E_NOT_FOUND for an unrecognized FolderId.", async () => {
        const context = makeContext();
        context.session.handles[5] = { type: "folder", entityUid: "folder:parent" };
        const handler = new RopDeleteFolderHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({ folderId: 999n })), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0x8004010f);
    });

    it("Deletes an empty folder with no flags set.", async () => {
        const context = makeContext();
        context.session.handles[5] = { type: "folder", entityUid: "folder:parent" };
        context.session.folderIds = { "1": "folder:target" };
        const handler = new RopDeleteFolderHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({})), writer, context);

        const response = new BufferReader(writer.toBuffer());
        expect(response.readUInt8()).toBe(0x1d);
        expect(response.readUInt8()).toBe(5);
        expect(response.readUInt32LE()).toBe(0);
        expect(response.readUInt8()).toBe(0); // PartialCompletion
        expect(response.hasMore()).toBe(false);

        expect((context.folderRepo as any).delete).toHaveBeenCalledWith("target", { ignoreACL: true });
    });

    it("Returns MAPI_E_INVALID_OBJECT for a non-empty folder (has messages) when DEL_MESSAGES isn't set, without deleting anything.", async () => {
        const context = makeContext({
            messageRepo: { find: vi.fn().mockResolvedValue([{ uid: "m1" }]), delete: vi.fn() } as any,
        });
        context.session.handles[5] = { type: "folder", entityUid: "folder:parent" };
        context.session.folderIds = { "1": "folder:target" };
        const handler = new RopDeleteFolderHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({})), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0x80070005);
        expect((context.folderRepo as any).delete).not.toHaveBeenCalled();
        expect((context.messageRepo as any).delete).not.toHaveBeenCalled();
    });

    it("Returns MAPI_E_INVALID_OBJECT for a folder with calendar events when DEL_MESSAGES isn't set.", async () => {
        const context = makeContext({
            calendarEventRepo: { find: vi.fn().mockResolvedValue([{ uid: "evt1" }]), delete: vi.fn() } as any,
        });
        context.session.handles[5] = { type: "folder", entityUid: "folder:parent" };
        context.session.folderIds = { "1": "folder:target" };
        const handler = new RopDeleteFolderHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({})), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0x80070005);
    });

    it("Deletes messages and calendar events, then the folder, when DEL_MESSAGES is set.", async () => {
        const context = makeContext({
            messageRepo: { find: vi.fn().mockResolvedValue([{ uid: "m1" }, { uid: "m2" }]), delete: vi.fn().mockResolvedValue(undefined) } as any,
            calendarEventRepo: { find: vi.fn().mockResolvedValue([{ uid: "evt1" }]), delete: vi.fn().mockResolvedValue(undefined) } as any,
        });
        context.session.handles[5] = { type: "folder", entityUid: "folder:parent" };
        context.session.folderIds = { "1": "folder:target" };
        const handler = new RopDeleteFolderHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({ deleteFolderFlags: DEL_MESSAGES })), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0);

        expect((context.messageRepo as any).delete).toHaveBeenCalledWith("m1", { ignoreACL: true });
        expect((context.messageRepo as any).delete).toHaveBeenCalledWith("m2", { ignoreACL: true });
        expect((context.calendarEventRepo as any).delete).toHaveBeenCalledWith("evt1", { ignoreACL: true });
        expect((context.folderRepo as any).delete).toHaveBeenCalledWith("target", { ignoreACL: true });
    });

    it("Returns MAPI_E_INVALID_OBJECT for a Contacts folder with contacts when DEL_MESSAGES isn't set, without deleting anything.", async () => {
        const context = makeContext({
            contactRepo: { find: vi.fn().mockResolvedValue([{ uid: "c1" }]), delete: vi.fn() } as any,
        });
        context.session.handles[5] = { type: "folder", entityUid: "folder:parent" };
        context.session.folderIds = { "1": "folder:target" };
        const handler = new RopDeleteFolderHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({})), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0x80070005);
        expect((context.folderRepo as any).delete).not.toHaveBeenCalled();
        expect((context.contactRepo as any).delete).not.toHaveBeenCalled();
    });

    it("Returns MAPI_E_INVALID_OBJECT for a Tasks folder with tasks when DEL_MESSAGES isn't set.", async () => {
        const context = makeContext({
            taskRepo: { find: vi.fn().mockResolvedValue([{ uid: "t1" }]), delete: vi.fn() } as any,
        });
        context.session.handles[5] = { type: "folder", entityUid: "folder:parent" };
        context.session.folderIds = { "1": "folder:target" };
        const handler = new RopDeleteFolderHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({})), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0x80070005);
        expect((context.taskRepo as any).delete).not.toHaveBeenCalled();
    });

    it("Deletes contacts and tasks, then the folder, when DEL_MESSAGES is set.", async () => {
        const context = makeContext({
            contactRepo: { find: vi.fn().mockResolvedValue([{ uid: "c1" }]), delete: vi.fn().mockResolvedValue(undefined) } as any,
            taskRepo: { find: vi.fn().mockResolvedValue([{ uid: "t1" }]), delete: vi.fn().mockResolvedValue(undefined) } as any,
        });
        context.session.handles[5] = { type: "folder", entityUid: "folder:parent" };
        context.session.folderIds = { "1": "folder:target" };
        const handler = new RopDeleteFolderHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({ deleteFolderFlags: DEL_MESSAGES })), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0);

        expect((context.contactRepo as any).delete).toHaveBeenCalledWith("c1", { ignoreACL: true });
        expect((context.taskRepo as any).delete).toHaveBeenCalledWith("t1", { ignoreACL: true });
        expect((context.folderRepo as any).delete).toHaveBeenCalledWith("target", { ignoreACL: true });
    });

    it("Treats a folder as empty for contacts/tasks purposes when contactRepo/taskRepo are absent from the context.", async () => {
        const context = makeContext();
        context.session.handles[5] = { type: "folder", entityUid: "folder:parent" };
        context.session.folderIds = { "1": "folder:target" };
        const handler = new RopDeleteFolderHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({})), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0);
        expect((context.folderRepo as any).delete).toHaveBeenCalledWith("target", { ignoreACL: true });
    });

    it("Cascades contact/task deletion into recursively-deleted subfolders too.", async () => {
        const folderFind = vi.fn().mockImplementation(({ parentFolderUid }: { parentFolderUid: string }) => {
            if (parentFolderUid === "target") return Promise.resolve([{ uid: "child1" }]);
            return Promise.resolve([]);
        });
        const contactFind = vi.fn().mockImplementation(({ folderUid }: { folderUid: string }) => {
            if (folderUid === "child1") return Promise.resolve([{ uid: "child-contact" }]);
            return Promise.resolve([]);
        });
        const taskFind = vi.fn().mockImplementation(({ folderUid }: { folderUid: string }) => {
            if (folderUid === "child1") return Promise.resolve([{ uid: "child-task" }]);
            return Promise.resolve([]);
        });
        const context = makeContext({
            folderRepo: { find: folderFind, delete: vi.fn().mockResolvedValue(undefined) } as any,
            contactRepo: { find: contactFind, delete: vi.fn().mockResolvedValue(undefined) } as any,
            taskRepo: { find: taskFind, delete: vi.fn().mockResolvedValue(undefined) } as any,
        });
        context.session.handles[5] = { type: "folder", entityUid: "folder:parent" };
        context.session.folderIds = { "1": "folder:target" };
        const handler = new RopDeleteFolderHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({ deleteFolderFlags: DEL_FOLDERS })), writer, context);

        expect((context.contactRepo as any).delete).toHaveBeenCalledWith("child-contact", { ignoreACL: true });
        expect((context.taskRepo as any).delete).toHaveBeenCalledWith("child-task", { ignoreACL: true });
        expect((context.folderRepo as any).delete).toHaveBeenCalledWith("child1", { ignoreACL: true });
    });

    it("Returns MAPI_E_INVALID_OBJECT for a folder with subfolders when DEL_FOLDERS isn't set.", async () => {
        const context = makeContext({
            folderRepo: {
                find: vi.fn().mockResolvedValue([{ uid: "child1" }]),
                delete: vi.fn(),
            } as any,
        });
        context.session.handles[5] = { type: "folder", entityUid: "folder:parent" };
        context.session.folderIds = { "1": "folder:target" };
        const handler = new RopDeleteFolderHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({})), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0x80070005);
        expect((context.folderRepo as any).delete).not.toHaveBeenCalled();
    });

    it("Recursively deletes subfolders (and their own messages/events) when DEL_FOLDERS is set.", async () => {
        const folderFind = vi.fn().mockImplementation(({ parentFolderUid }: { parentFolderUid: string }) => {
            if (parentFolderUid === "target") return Promise.resolve([{ uid: "child1" }]);
            if (parentFolderUid === "child1") return Promise.resolve([{ uid: "grandchild1" }]);
            return Promise.resolve([]);
        });
        const messageFind = vi.fn().mockImplementation(({ folderUid }: { folderUid: string }) => {
            if (folderUid === "child1") return Promise.resolve([{ uid: "child-msg" }]);
            return Promise.resolve([]);
        });
        const eventFind = vi.fn().mockImplementation(({ folderUid }: { folderUid: string }) => {
            if (folderUid === "grandchild1") return Promise.resolve([{ uid: "grandchild-evt" }]);
            return Promise.resolve([]);
        });
        const context = makeContext({
            folderRepo: { find: folderFind, delete: vi.fn().mockResolvedValue(undefined) } as any,
            messageRepo: { find: messageFind, delete: vi.fn().mockResolvedValue(undefined) } as any,
            calendarEventRepo: { find: eventFind, delete: vi.fn().mockResolvedValue(undefined) } as any,
        });
        context.session.handles[5] = { type: "folder", entityUid: "folder:parent" };
        context.session.folderIds = { "1": "folder:target" };
        const handler = new RopDeleteFolderHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({ deleteFolderFlags: DEL_FOLDERS })), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0);

        expect((context.messageRepo as any).delete).toHaveBeenCalledWith("child-msg", { ignoreACL: true });
        expect((context.calendarEventRepo as any).delete).toHaveBeenCalledWith("grandchild-evt", { ignoreACL: true });
        expect((context.folderRepo as any).delete).toHaveBeenCalledWith("grandchild1", { ignoreACL: true });
        expect((context.folderRepo as any).delete).toHaveBeenCalledWith("child1", { ignoreACL: true });
        expect((context.folderRepo as any).delete).toHaveBeenCalledWith("target", { ignoreACL: true });
    });

    it("Never purges, even with DELETE_HARD_DELETE set, so legal holds can't be bypassed through MAPI.", async () => {
        const context = makeContext({
            messageRepo: { find: vi.fn().mockResolvedValue([{ uid: "m1" }]), delete: vi.fn().mockResolvedValue(undefined) } as any,
        });
        context.session.handles[5] = { type: "folder", entityUid: "folder:parent" };
        context.session.folderIds = { "1": "folder:target" };

        await new RopDeleteFolderHandler().handle(
            new BufferReader(buildRequest({ deleteFolderFlags: DELETE_HARD_DELETE | DEL_MESSAGES })),
            new BufferWriter(),
            context,
        );

        expect((context.messageRepo as any).delete).toHaveBeenCalledWith("m1", { ignoreACL: true });
        expect((context.folderRepo as any).delete).toHaveBeenCalledWith("target", { ignoreACL: true });
        for (const repo of [context.messageRepo, context.folderRepo] as any[]) {
            for (const call of repo.delete.mock.calls) {
                expect(call[1].purge).toBeUndefined();
            }
        }
    });

    it("Audits every deleted message as MESSAGE_DELETE, like the REST message route.", async () => {
        const audit = vi.fn().mockResolvedValue(undefined);
        const context = makeContext({
            audit,
            messageRepo: {
                find: vi.fn().mockResolvedValue([{ uid: "m1", mailboxUid: "mailbox-1", subject: "Hi", folderUid: "target" }]),
                delete: vi.fn().mockResolvedValue(undefined),
            } as any,
            calendarEventRepo: { find: vi.fn().mockResolvedValue([{ uid: "evt1" }]), delete: vi.fn().mockResolvedValue(undefined) } as any,
        });
        context.session.handles[5] = { type: "folder", entityUid: "folder:parent" };
        context.session.folderIds = { "1": "folder:target" };

        await new RopDeleteFolderHandler().handle(new BufferReader(buildRequest({ deleteFolderFlags: DEL_MESSAGES })), new BufferWriter(), context);

        expect(audit).toHaveBeenCalledTimes(1);
        expect(audit).toHaveBeenCalledWith({
            action: AuditAction.MESSAGE_DELETE,
            targetType: "Message",
            targetUid: "m1",
            mailboxUid: "mailbox-1",
            details: { subject: "Hi", folderUid: "target" },
        });
    });

    it("Scopes every item and child-folder lookup to the caller's mailbox, with explicit paging and sort.", async () => {
        const context = makeContext();
        context.session.handles[5] = { type: "folder", entityUid: "folder:parent" };
        context.session.folderIds = { "1": "folder:target" };

        await new RopDeleteFolderHandler().handle(
            new BufferReader(buildRequest({ deleteFolderFlags: DEL_MESSAGES | DEL_FOLDERS })),
            new BufferWriter(),
            context,
        );

        expect((context.folderRepo as any).find).toHaveBeenCalledWith(
            { parentFolderUid: "target", mailboxUid: "mailbox-1", sort: { uid: "ASC" }, limit: 1000, page: 0 },
            { ignoreACL: true, limit: 1000, page: 0 },
        );
        expect((context.messageRepo as any).find).toHaveBeenCalledWith(
            { folderUid: "target", mailboxUid: "mailbox-1", sort: { uid: "ASC" }, limit: 1000, page: 0 },
            { ignoreACL: true, limit: 1000, page: 0 },
        );
    });

    it("Survives a parentFolderUid cycle instead of recursing forever.", async () => {
        const folderFind = vi.fn().mockImplementation(({ parentFolderUid }: { parentFolderUid: string }) => {
            if (parentFolderUid === "target") return Promise.resolve([{ uid: "child1" }]);
            if (parentFolderUid === "child1") return Promise.resolve([{ uid: "target" }, { uid: "child1" }]);
            return Promise.resolve([]);
        });
        const context = makeContext({ folderRepo: { find: folderFind, delete: vi.fn().mockResolvedValue(undefined) } as any });
        context.session.handles[5] = { type: "folder", entityUid: "folder:parent" };
        context.session.folderIds = { "1": "folder:target" };
        const writer = new BufferWriter();

        await new RopDeleteFolderHandler().handle(new BufferReader(buildRequest({ deleteFolderFlags: DEL_FOLDERS })), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0);
        expect((context.folderRepo as any).delete.mock.calls.map((call: any[]) => call[0])).toEqual(["child1", "target"]);
    });

    it("Returns MAPI_E_TOO_COMPLEX without deleting anything for a tree deeper than MAX_FOLDER_DEPTH.", async () => {
        const folderFind = vi.fn().mockImplementation(({ parentFolderUid }: { parentFolderUid: string }) => {
            const depth = parentFolderUid === "target" ? 0 : Number(parentFolderUid.slice("f".length));
            return Promise.resolve([{ uid: `f${depth + 1}` }]);
        });
        const context = makeContext({ folderRepo: { find: folderFind, delete: vi.fn() } as any });
        context.session.handles[5] = { type: "folder", entityUid: "folder:parent" };
        context.session.folderIds = { "1": "folder:target" };
        const writer = new BufferWriter();

        await new RopDeleteFolderHandler().handle(new BufferReader(buildRequest({ deleteFolderFlags: DEL_FOLDERS })), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0x80040117);
        expect(folderFind.mock.calls.length).toBe(MAX_FOLDER_DEPTH + 1); // one child lookup per level walked
        expect((context.folderRepo as any).delete).not.toHaveBeenCalled();
    });

    it("Returns MAPI_E_TOO_COMPLEX for a tree with more than MAX_FOLDERS_PER_DELETE folders.", async () => {
        // 999 children directly under the target fill the budget (with the target itself); a grandchild tips it over.
        const wide = Array.from({ length: MAX_FOLDERS_PER_DELETE - 1 }, (_, i) => ({ uid: `child${i}` }));
        const folderFind = vi.fn().mockImplementation(({ parentFolderUid }: { parentFolderUid: string }) =>
            Promise.resolve(parentFolderUid === "target" ? wide : parentFolderUid === "child998" ? [{ uid: "grandchild" }] : []),
        );
        const context = makeContext({ folderRepo: { find: folderFind, delete: vi.fn() } as any });
        context.session.handles[5] = { type: "folder", entityUid: "folder:parent" };
        context.session.folderIds = { "1": "folder:target" };
        const writer = new BufferWriter();

        await new RopDeleteFolderHandler().handle(new BufferReader(buildRequest({ deleteFolderFlags: DEL_FOLDERS })), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0x80040117);
        expect((context.folderRepo as any).delete).not.toHaveBeenCalled();
    });

    it("Returns MAPI_E_TOO_COMPLEX when one folder's child list is itself past the cap.", async () => {
        const fullPage = Array.from({ length: 1000 }, (_, i) => ({ uid: `child${i}` }));
        const folderFind = vi.fn().mockImplementation(({ parentFolderUid }: { parentFolderUid: string }) =>
            Promise.resolve(parentFolderUid === "target" ? [...fullPage, { uid: "one-more" }] : []),
        );
        const context = makeContext({ folderRepo: { find: folderFind, delete: vi.fn() } as any });
        context.session.handles[5] = { type: "folder", entityUid: "folder:parent" };
        context.session.folderIds = { "1": "folder:target" };
        const writer = new BufferWriter();

        await new RopDeleteFolderHandler().handle(new BufferReader(buildRequest({ deleteFolderFlags: DEL_FOLDERS })), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0x80040117);
    });

    it("Pages through more than 100 items, and stops with PartialCompletion once MAX_ITEMS_PER_DELETE is spent, keeping the folder.", async () => {
        const page = (offset: number) => Array.from({ length: 1000 }, (_, i) => ({ uid: `m${offset + i}` }));
        const messageFind = vi.fn().mockImplementation((query: { page: number }) => Promise.resolve(page(query.page * 1000)));
        const notifyFolderCounts = vi.fn().mockResolvedValue(undefined);
        const context = makeContext({
            notifyFolderCounts,
            messageRepo: { find: messageFind, delete: vi.fn().mockResolvedValue(undefined) } as any,
        });
        context.session.handles[5] = { type: "folder", entityUid: "folder:parent" };
        context.session.folderIds = { "1": "folder:target" };
        const writer = new BufferWriter();

        await new RopDeleteFolderHandler().handle(new BufferReader(buildRequest({ deleteFolderFlags: DEL_MESSAGES })), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0);
        expect(response.readUInt8()).toBe(1); // PartialCompletion
        expect((context.messageRepo as any).delete).toHaveBeenCalledTimes(MAX_ITEMS_PER_DELETE);
        expect((context.folderRepo as any).delete).not.toHaveBeenCalled();
        // The folder itself survives the partial delete (never reached the `folderRepo.delete` call above), so unlike
        // a folder that's fully deleted below, its now-stale message count must be refreshed instead of left behind.
        expect(notifyFolderCounts).toHaveBeenCalledTimes(1);
        expect(notifyFolderCounts).toHaveBeenCalledWith(["target"]);
    });

    it("Never refreshes a fully-deleted folder's counts - it no longer exists to refresh.", async () => {
        const notifyFolderCounts = vi.fn().mockResolvedValue(undefined);
        const context = makeContext({
            notifyFolderCounts,
            messageRepo: { find: vi.fn().mockResolvedValue([{ uid: "m1" }]), delete: vi.fn().mockResolvedValue(undefined) } as any,
        });
        context.session.handles[5] = { type: "folder", entityUid: "folder:parent" };
        context.session.folderIds = { "1": "folder:target" };
        const writer = new BufferWriter();

        await new RopDeleteFolderHandler().handle(new BufferReader(buildRequest({ deleteFolderFlags: DEL_MESSAGES })), writer, context);

        expect((context.folderRepo as any).delete).toHaveBeenCalledWith("target", { ignoreACL: true });
        expect(notifyFolderCounts).not.toHaveBeenCalled();
    });
});
