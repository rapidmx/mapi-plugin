///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
// Unit tests for the single `@Init` hook of the MAPI routes and of `MapiSessionManager`, which build each repository
// and service once through the ObjectFactory instead of lazily.
import { RepoUtils } from "@rapidrest/service-core";
import { MapiSessionManager, MemoryMapiSessionStore, RedisMapiSessionStore } from "../src/MapiSessionManager.js";
import { BaseMapiEmsmdbRoute } from "../src/BaseMapiEmsmdbRoute.js";
import { BaseMapiNspiRoute } from "../src/BaseMapiNspiRoute.js";
import { MemoryHandleDataStore, RedisHandleDataStore } from "../src/rop/HandleDataCache.js";
import { AuditLogUtils, RecoverableRepoUtils } from "@rapidmx/restapi";

/** A uniquely named stand-in model class. */
function model(name: string): any {
    return { [name]: class {} }[name];
}

/** A fake ObjectFactory recording every `newInstance` call and answering with a tagged object. */
function fakeFactory(): { factory: any; calls: any[] } {
    const calls: any[] = [];
    const factory = {
        newInstance: vi.fn(async (type: any, options?: any) => {
            calls.push({ type, options });
            return { built: type, options, ropId: type.ropId };
        }),
    };
    return { factory, calls };
}

class TestEmsmdbRoute extends BaseMapiEmsmdbRoute<any> {
    protected mailboxClass: any = model("MailboxModel");
    protected folderClass: any = model("FolderModel");
    protected messageClass: any = model("MessageModel");
    protected calendarEventClass: any = model("CalendarEventModel");
    protected contactClass: any = model("ContactModel");
    protected taskClass: any = model("TaskModel");
    protected labelClass: any = model("LabelModel");
    protected auditLogClass: any = model("AuditModel");
    protected ropHandlerClasses: any[] = [];
}

class TestNspiRoute extends BaseMapiNspiRoute<any> {
    protected mailboxClass: any = model("NspiMailboxModel");
    protected contactClass: any = model("NspiContactModel");
}

describe("BaseMapiEmsmdbRoute initialize", () => {
    it("Refuses to build its repos and audit service when no ObjectFactory was injected.", async () => {
        const route: any = new TestEmsmdbRoute();
        await expect(route.initialize()).rejects.toThrow("objectFactory is not set.");
    });

    it("Builds each repo, the audit service, the session manager and the ROP handlers with exactly { name, args }.", async () => {
        class Handler {
            public static ropId = 0x42;
        }
        const route: any = new TestEmsmdbRoute();
        route.ropHandlerClasses = [Handler];
        const { factory, calls } = fakeFactory();
        route._objectFactory = factory;

        await route.initialize();

        const byName = (name: string) => calls.find((c) => c.options?.name === name && c.type !== AuditLogUtils);
        expect(byName("MailboxModel").type).toBe(RepoUtils);
        expect(byName("MailboxModel").options).toEqual({ name: "MailboxModel", args: [route.mailboxClass] });
        const recoverable: [string, string][] = [
            ["FolderModel", "folderClass"],
            ["MessageModel", "messageClass"],
            ["CalendarEventModel", "calendarEventClass"],
            ["ContactModel", "contactClass"],
            ["TaskModel", "taskClass"],
        ];
        for (const [name, field] of recoverable) {
            expect(byName(name).type).toBe(RecoverableRepoUtils);
            expect(byName(name).options).toEqual({ name, args: [route[field]] });
        }
        expect(byName("LabelModel").type).toBe(RepoUtils);
        expect(byName("AuditModel").type).toBe(RepoUtils);
        const audit = calls.find((c) => c.type === AuditLogUtils);
        expect(audit.options.name).toBe("AuditModel");
        expect(audit.options.args).toHaveLength(1);
        expect(calls.some((c) => c.type === MapiSessionManager && c.options === undefined)).toBe(true);
        expect(calls.some((c) => c.type === Handler)).toBe(true);
        expect(route.ropHandlers.get(0x42)).toBeDefined();
    });

    it("Does not rebuild a pre-set repo, service, session manager or handler.", async () => {
        const route: any = new TestEmsmdbRoute();
        const { factory, calls } = fakeFactory();
        route._objectFactory = factory;
        for (const field of [
            "mailboxRepo",
            "folderRepo",
            "messageRepo",
            "calendarEventRepo",
            "contactRepo",
            "taskRepo",
            "labelRepo",
            "auditLogUtils",
            "sessionManager",
        ]) {
            route[field] = { preset: field };
        }
        route.ropHandlers.set(1, {});
        route.ropHandlerClasses = [class {}];

        await route.initialize();

        expect(calls).toEqual([]);
        expect(route.mailboxRepo).toEqual({ preset: "mailboxRepo" });
    });

    it("Skips the audit service when no audit log class is set, and any repo whose class is unset.", async () => {
        const route: any = new TestEmsmdbRoute();
        route.auditLogClass = undefined;
        route.labelClass = undefined;
        const { factory, calls } = fakeFactory();
        route._objectFactory = factory;

        await route.initialize();

        expect(calls.some((c) => c.type === AuditLogUtils)).toBe(false);
        expect(route.auditLogUtils).toBeUndefined();
        expect(route.labelRepo).toBeUndefined();
        expect(route.mailboxRepo).toBeDefined();
    });

    it("Skips every repo whose model class is unset.", async () => {
        const route: any = new TestEmsmdbRoute();
        for (const field of ["mailbox", "folder", "message", "calendarEvent", "contact", "task", "label", "auditLog"]) {
            route[`${field}Class`] = undefined;
        }
        const { factory, calls } = fakeFactory();
        route._objectFactory = factory;

        await route.initialize();

        expect(calls.map((c) => c.type)).toEqual([MapiSessionManager]);
    });
});

describe("BaseMapiNspiRoute initialize", () => {
    it("Refuses to build its repos when no ObjectFactory was injected.", async () => {
        const route: any = new TestNspiRoute();
        await expect(route.initialize()).rejects.toThrow("objectFactory is not set.");
    });

    it("Builds the mailbox and contact repos with exactly { name, args }.", async () => {
        const route: any = new TestNspiRoute();
        const { factory, calls } = fakeFactory();
        route._objectFactory = factory;

        await route.initialize();

        expect(calls).toEqual([
            { type: RepoUtils, options: { name: "NspiMailboxModel", args: [route.mailboxClass] } },
            { type: RepoUtils, options: { name: "NspiContactModel", args: [route.contactClass] } },
        ]);
    });

    it("Does not rebuild a pre-set repo, and skips an unset class.", async () => {
        const route: any = new TestNspiRoute();
        const { factory, calls } = fakeFactory();
        route._objectFactory = factory;
        route.mailboxRepo = { preset: true };
        route.contactClass = undefined;

        await route.initialize();

        expect(calls).toEqual([]);
        expect(route.mailboxRepo).toEqual({ preset: true });
        expect(route.contactRepo).toBeUndefined();
    });
});

describe("MapiSessionManager init", () => {
    it("Refuses to build its stores when no ObjectFactory was injected.", async () => {
        await expect(new MapiSessionManager().init()).rejects.toThrow("objectFactory is not set.");
    });

    it("Builds the in-process stores through the factory, named, when no Redis client is configured.", async () => {
        const manager: any = new MapiSessionManager();
        const { factory, calls } = fakeFactory();
        manager._objectFactory = factory;

        await manager.init();

        expect(calls).toEqual([
            { type: MemoryMapiSessionStore, options: { name: "MemoryMapiSessionStore" } },
            { type: MemoryHandleDataStore, options: { name: "MemoryHandleDataStore" } },
        ]);
    });

    it("Builds the Redis stores with exactly { name, args } when a Redis client is configured.", async () => {
        const manager: any = new MapiSessionManager();
        const { factory, calls } = fakeFactory();
        manager._objectFactory = factory;
        manager.redisClient = { client: true };

        await manager.init();

        expect(calls).toEqual([
            { type: RedisMapiSessionStore, options: { name: "RedisMapiSessionStore", args: [manager.redisClient] } },
            { type: RedisHandleDataStore, options: { name: "RedisHandleDataStore", args: [manager.redisClient] } },
        ]);
    });

    it("Does not rebuild pre-set stores.", async () => {
        const manager: any = new MapiSessionManager();
        const { factory, calls } = fakeFactory();
        manager._objectFactory = factory;
        manager.store = {};
        manager.handleDataStore = {};

        await manager.init();

        expect(calls).toEqual([]);
    });
});
