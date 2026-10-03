///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { ApiError, ObjectDecorators, type JWTUser } from "@rapidrest/core";
import { ApiErrorMessages, ApiErrors, HttpRequest, HttpResponse, ObjectFactory, RepoUtils, RouteDecorators } from "@rapidrest/service-core";
import { handleNspiBind, handleNspiUnbind } from "./nspi/NspiBindHandler.js";
import { handleNspiGetMatches } from "./nspi/NspiGetMatchesHandler.js";
import { Mailbox, resolveCallerMailboxUid } from "@rapidmx/restapi";
const { Init, Logger } = ObjectDecorators;
const { Auth, Post, Request, Response, User: AuthUser } = RouteDecorators;

function firstHeader(req: HttpRequest, name: string): string | undefined {
    const value: string | string[] | undefined = req.headers[name];
    return Array.isArray(value) ? value[0] : value;
}

/**
 * Abstract base for the single fixed NSPI endpoint (`POST /mapi/nspi` by `[MS-OXCMAPIHTTP]` convention, though
 * the concrete path is left to the consuming application to mount via `@Route(...)`, the identical
 * undecorated-base-class pattern `BaseMapiEmsmdbRoute`/`BaseEasRoute` already establish). Dispatches on the
 * `X-RequestType` header, the same multiplexing mechanism EMSMDB uses - but unlike EMSMDB, there is no ROP
 * buffer framing here: each `X-RequestType` value (`Bind`/`Unbind`/`GetMatches`) is its own flat, self-
 * contained request/response body (`[MS-OXCMAPIHTTP]` §2.2.5), so this route dispatches directly to a small
 * handler function per operation instead of decoding a multiplexed stream.
 *
 * **Pragmatic subset scope**: only `Bind`/`Unbind`/`GetMatches` are implemented - real NSPI's other dozen-plus
 * request types (`QueryRows`, `ResolveNames`, `GetProps`, `ModProps`, ...) are genuine full address-book/
 * directory-replication operations this deployment's own minimal GAL-lookup use case doesn't need; an
 * unrecognized `X-RequestType` gets a `501`, the same "recognized-but-deferred vs malformed" distinction
 * `BaseMapiEmsmdbRoute`'s own `NotificationWait` case already draws. See `NspiBindHandler.ts`'s own doc comment
 * for why no real NSPI session state is created or tracked between calls.
 *
 * **Auth**: `@Auth(["jwt"])`, unchanged from every other route in this app - the same reasoning
 * `BaseMapiEmsmdbRoute`'s own doc comment already gives for why Bearer-token auth on an address-book endpoint
 * is real, current Exchange behavior, not a deviation this library invents.
 *
 * `mailboxClass`/`contactClass` are supplied by the Mongo/SQL concrete subclasses. GAL search itself (see
 * `NspiGetMatchesHandler.ts`) uses `RepoUtils`'s `regex(...)` operator, which compiles identically on both
 * backends - no per-backend hook needed here.
 *
 * @author Jean-Philippe Steinmetz
 */
export abstract class BaseMapiNspiRoute<M extends Mailbox> {
    protected abstract mailboxClass: any;
    protected abstract contactClass: any;

    // Automatically injected by ObjectFactory on instantiation
    private _objectFactory?: ObjectFactory;

    private mailboxRepo?: RepoUtils<M>;
    private contactRepo?: RepoUtils<any>;

    // TEMPORARY DIAGNOSTIC LOGGING - see NOTES.md's 2026-09-28 "diagnostic release" entry. Remove once the
    // Outlook desktop "set of folders cannot be opened" root cause is confirmed.
    @Logger
    private logger: any;

    /** Builds the mailbox and contact repositories once, here, rather than lazily in `dispatch()`. */
    @Init
    protected async initialize(): Promise<void> {
        if (!this._objectFactory) {
            throw new Error("objectFactory is not set.");
        }
        if (!this.mailboxRepo && this.mailboxClass) {
            this.mailboxRepo = await this._objectFactory.newInstance(RepoUtils, {
                name: this.mailboxClass.name,
                args: [this.mailboxClass],
            });
        }
        if (!this.contactRepo && this.contactClass) {
            this.contactRepo = await this._objectFactory.newInstance(RepoUtils, {
                name: this.contactClass.name,
                args: [this.contactClass],
            });
        }
    }

    @Auth(["jwt"])
    @Post()
    public async dispatch(@Request req: HttpRequest, @Response res: HttpResponse, @AuthUser user?: JWTUser): Promise<void> {
        if (!this.mailboxRepo || !this.contactRepo) {
            throw new ApiError(ApiErrors.INTERNAL_ERROR, 500, ApiErrorMessages.INTERNAL_ERROR);
        }
        if (!user) {
            throw new ApiError(ApiErrors.AUTH_PERMISSION_FAILURE, 403, ApiErrorMessages.AUTH_PERMISSION_FAILURE);
        }

        const requestType: string | undefined = firstHeader(req, "x-requesttype");
        // [MS-OXCMAPIHTTP] 2.2.3.3.2/"Common Response Format": every response (success or failure, every
        // request type, this endpoint included) MUST carry X-RequestId back with the exact value the client
        // sent on the request - the server-assigned session-affinity GUID:counter the client uses to detect a
        // response for the wrong request. This route never echoed it at all - confirmed against both the spec
        // text and Gromox's own `commonHeader()` (github.com/grommunio/gromox, exch/mh/mh_common.cpp), which
        // sets X-RequestId/X-ClientInfo unconditionally from one shared function for exactly this reason - a
        // per-route header list, like this one used to be, is exactly how an endpoint can silently miss it.
        // Mirrors BaseMapiEmsmdbRoute.dispatch()'s own identical header set.
        const clientInfo: string = firstHeader(req, "x-clientinfo") ?? "";
        const requestId: string = firstHeader(req, "x-requestid") ?? "";
        res.setHeader("Content-Type", "application/mapi-http")
            .setHeader("X-RequestType", requestType ?? "")
            .setHeader("X-RequestId", requestId)
            .setHeader("X-ResponseCode", "0")
            .setHeader("X-ClientInfo", clientInfo)
            .setHeader("X-ServerApplication", "RapidREST-Mail");
        // Strip the framework's own CSRF cookie before this route's own Set-Cookie (Bind's NspiContext) gets
        // appended - see the matching comment in BaseMapiEmsmdbRoute.dispatch() for the full reasoning.
        res.setHeader("Set-Cookie", []);

        // TEMPORARY DIAGNOSTIC LOGGING - see the matching comment above. Wraps res.send() to log the exact
        // outgoing bytes regardless of which case below builds them, and logs the raw incoming body too (the
        // request body is otherwise never decoded/logged at all for Bind).
        this.logger?.warn(
            `MAPI_DEBUG NSPI dispatch requestType=${requestType} user=${user.uid} clientInfo=${clientInfo} requestId=${requestId} rawBodyHex=${req.rawBody ? req.rawBody.toString("hex") : "<empty>"}`,
        );
        const originalSend = res.send.bind(res);
        res.send = (body?: any): any => {
            this.logger?.warn(`MAPI_DEBUG NSPI ${requestType} OUT responseHex=${Buffer.isBuffer(body) ? body.toString("hex") : String(body)}`);
            return originalSend(body);
        };
        try {
            switch (requestType) {
                case "Bind":
                    await handleNspiBind(res, user, this.mailboxRepo);
                    this.logger?.warn(`MAPI_DEBUG NSPI Bind OK user=${user.uid}`);
                    return;
                case "Unbind":
                    handleNspiUnbind(res);
                    return;
                case "GetMatches": {
                    const mailboxUid = await resolveCallerMailboxUid(this.mailboxRepo, user);
                    if (!mailboxUid) {
                        throw new ApiError(ApiErrors.NOT_FOUND, 404, ApiErrorMessages.NOT_FOUND);
                    }
                    await handleNspiGetMatches(req, res, mailboxUid, this.contactRepo);
                    this.logger?.warn(`MAPI_DEBUG NSPI GetMatches OK mailboxUid=${mailboxUid}`);
                    return;
                }
                default:
                    this.logger?.warn(`MAPI_DEBUG NSPI unrecognized requestType=${requestType}, answering 501`);
                    res.status(501).send();
                    return;
            }
        } catch (err) {
            // TEMPORARY DIAGNOSTIC LOGGING - the real error here would otherwise only ever surface as a generic
            // 500/ApiError with no server-side trace of why.
            this.logger?.warn(`MAPI_DEBUG NSPI ${requestType} threw:`, err);
            throw err;
        }
    }
}
