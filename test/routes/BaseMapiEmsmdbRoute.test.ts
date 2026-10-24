///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
// Isolated unit tests for BaseMapiEmsmdbRoute, reserved for the defensive guard branches a real wired server
// can never exercise (`!this.mailboxRepo || !this.folderRepo || !this.messageRepo || !this.calendarEventRepo ||
// !this.contactRepo || !this.taskRepo || !this.labelRepo || !this.sessionManager || !this.blobStore ||
// !this.scanPipeline || !this.mailTransport` - DI always populates all eleven before a request can reach a
// route - and `!user`, a second defensive check behind `@Auth(["jwt"])` itself) - the same rationale
// test/routes/BaseEasRoute.test.ts already uses for its own identical guards. Every other behavior is
// exercised via real HTTP+DB requests in test/routes/mongo/MapiEmsmdbRoute.test.ts (and its sql/ counterpart).
import config from "../config.js";
import { ObjectFactory } from "@rapidrest/service-core";
import { Logger } from "@rapidrest/core";
import { BaseMapiEmsmdbRoute } from "../../src/BaseMapiEmsmdbRoute.js";

class TestMapiEmsmdbRoute extends BaseMapiEmsmdbRoute<any> {
    protected mailboxClass: any = { name: "TestMailbox" };
    protected folderClass: any = { name: "TestFolder" };
    protected messageClass: any = { name: "TestMessage" };
    protected calendarEventClass: any = { name: "TestCalendarEvent" };
    protected contactClass: any = { name: "TestContact" };
    protected taskClass: any = { name: "TestTask" };
    protected labelClass: any = { name: "TestLabel" };
}

function makeReq(): any {
    return { headers: { "x-requesttype": "Connect" }, cookies: {}, rawBody: undefined };
}

function makeRes(): any {
    return {
        status: vi.fn().mockReturnThis(),
        setHeader: vi.fn().mockReturnThis(),
        appendHeader: vi.fn().mockReturnThis(),
        send: vi.fn().mockReturnThis(),
    };
}

describe("BaseMapiEmsmdbRoute Tests (guard clauses only)", () => {
    const objectFactory: ObjectFactory = new ObjectFactory(config, Logger());

    afterEach(() => {
        vi.restoreAllMocks();
    });

    it("dispatch() throws INTERNAL_ERROR when mailboxRepo/folderRepo/messageRepo/calendarEventRepo/contactRepo/taskRepo/labelRepo/sessionManager/blobStore/scanPipeline/mailTransport are not set.", async () => {
        // `initialize: false` skips `@Init`, leaving all eleven genuinely `undefined` - exactly what this
        // guard clause exists to catch.
        const route = objectFactory.newInstance<TestMapiEmsmdbRoute>(TestMapiEmsmdbRoute, { initialize: false });

        await expect(route.dispatch(makeReq(), makeRes(), { uid: "user-1" } as any)).rejects.toThrow(
            /internal error/i,
        );
    });

    it("dispatch() throws AUTH_PERMISSION_FAILURE when no authenticated user is present.", async () => {
        const route = objectFactory.newInstance<TestMapiEmsmdbRoute>(TestMapiEmsmdbRoute, { initialize: false });
        // Poking the private fields directly (TypeScript `private` is compile-time only) isolates this guard
        // from the one above, which would otherwise fire first.
        (route as any).mailboxRepo = {};
        (route as any).folderRepo = {};
        (route as any).messageRepo = {};
        (route as any).calendarEventRepo = {};
        (route as any).contactRepo = {};
        (route as any).taskRepo = {};
        (route as any).labelRepo = {};
        (route as any).sessionManager = {};
        (route as any).blobStore = {};
        (route as any).scanPipeline = {};
        (route as any).mailTransport = {};

        await expect(route.dispatch(makeReq(), makeRes(), undefined)).rejects.toThrow(/permission/i);
    });
});

describe("BaseMapiEmsmdbRoute.recordOutlook (isolated)", () => {
    it("records Outlook on the mailbox, skips a recent record, and only logs a write that fails", async () => {
        const route = new TestMapiEmsmdbRoute() as any;
        route.logger = { warn: vi.fn() };
        route.mailboxRepo = { update: vi.fn().mockResolvedValue({ version: 2 }) };
        const mailbox: any = { uid: "mb", version: 1, syncClients: { activesync: { devices: [] } } };

        await route.recordOutlook(mailbox, 1_000);
        expect(route.mailboxRepo.update.mock.calls[0][0]).toEqual({ uid: "mb", version: 1, syncClients: { activesync: { devices: [] }, mapi: { lastSeenAt: 1_000 } } });

        await route.recordOutlook({ uid: "mb", syncClients: { mapi: { lastSeenAt: 1_000 } } }, 2_000);
        await route.recordOutlook(undefined, 2_000);
        expect(route.mailboxRepo.update).toHaveBeenCalledTimes(1);

        route.mailboxRepo.update.mockRejectedValueOnce(new Error("conflict"));
        await route.recordOutlook({ uid: "mb2" });
        expect(route.logger.warn).toHaveBeenCalledWith(expect.stringContaining("failed to record Outlook on mailbox mb2: conflict"));
    });
});

