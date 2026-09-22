///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { ApiError, ObjectDecorators, type JWTUser } from "@rapidrest/core";
import {
    ApiErrorMessages,
    ApiErrors,
    HttpRequest,
    HttpResponse,
    NotificationUtils,
    ObjectFactory,
    RepoUtils,
    RouteDecorators,
} from "@rapidrest/service-core";
import { BufferReader, BufferWriter, DecodeError } from "./codec/BufferCursor.js";
import { decodeRopBuffer, encodeRopBuffer, type RopBuffer } from "./codec/RopBuffer.js";
import { MapiSessionContext, MapiSessionManager, SESSION_LOCK_RENEW_MS, takeReleasedHandleData } from "./MapiSessionManager.js";
import { dispatchRops, ExecuteBufferTooSmallError, MAX_ROPS_LIST_BYTES } from "./RopDispatcher.js";
import { handleDataOwners, type HandleDataStore } from "./rop/HandleDataCache.js";
import { handleDataStoreOf, type RopContext, type RopHandler } from "./rop/RopHandler.js";
import { ScanPipeline } from "@rapidmx/restapi/scan";
import {
    Folder,
    Mailbox,
    RecoverableRepoUtils,
    recordAuditLog,
    refreshFolderCounts,
    resolveCallerMailboxUid,
    type BlobStore,
    type FolderCountsContext,
} from "@rapidmx/restapi";
const { Config, Init, Inject, Logger } = ObjectDecorators;
const { Auth, Post, Request, Response, User: AuthUser } = RouteDecorators;

/** `[MS-OXCMAPIHTTP]` `X-ResponseCode` 10, "Context Not Found": the session context named by the `MapiContext`
 * cookie doesn't exist (expired, ended, or not the caller's). The client `Connect`s again. `X-ResponseCode` only takes
 * the spec's small status values; the MAPI HRESULT goes in the body's `ErrorCode` (`ERROR_SESSION_NOT_FOUND`). */
const RESPONSE_CODE_CONTEXT_NOT_FOUND = 10;

/** The well-known MAPI HRESULT `MAPI_E_LOGON_FAILED`, reused as the body `ErrorCode` for "no such session -
 * reconnect". Not a claim of exact `[MS-OXCRPC]` return-value-table parity for this specific condition. */
const ERROR_SESSION_NOT_FOUND = 0x80040111;

/** `MAPI_E_TOO_BIG`, the body `ErrorCode` when a request would grow the session past `MAX_SESSION_BYTES`. */
const ERROR_SESSION_TOO_BIG = 0x80040305;

/** `[MS-OXCMAPIHTTP]` `X-ResponseCode` 15, "Invalid Sequence": the request overlapped another one on the same
 * session context. Sent when the session changed underneath this request (see `MapiSessionManager.save()`). */
const RESPONSE_CODE_INVALID_SEQUENCE = 15;

/** `ecBufferTooSmall` ([MS-OXCRPC]), the body `ErrorCode` when the client's `MaxRopOut` can't even hold a
 * `RopBufferTooSmall` carrying the whole request. No ROP has run by then. */
const ERROR_BUFFER_TOO_SMALL = 0x0000047d;

/** The largest `RopBufferSize` accepted - the `[MS-OXCRPC]` ROP input buffer limit. */
export const MAX_ROP_BUFFER_SIZE = 32767;

/** A decoded `Execute` request: the ROP buffer framing plus the client's `MaxRopOut`. */
export interface ExecuteRequest extends RopBuffer {
    /** The most bytes the client accepts in the response's `RopBuffer`. */
    maxRopOut: number;
}

/**
 * Decodes an `Execute` request body's `Flags`/`RopBufferSize`/`RopBuffer`/`MaxRopOut` and the ROP buffer framing
 * inside it. Throws a 400 `ApiError` for a malformed body - a `RopBufferSize` over `MAX_ROP_BUFFER_SIZE` or past the
 * end of the body, or a ROP buffer `decodeRopBuffer` rejects - instead of letting a bare `RangeError` surface as a
 * 500. A body that ends before `MaxRopOut` gets the largest response a `RopBuffer` can carry.
 * `AuxiliaryBufferSize`/`AuxiliaryBuffer` are left unread - no auxiliary-payload support in this pragmatic subset.
 */
export function decodeExecuteRequest(rawBody: Buffer): ExecuteRequest {
    try {
        const reader = new BufferReader(rawBody);
        reader.readUInt32LE(); // Flags - unused by this pragmatic subset (no client ROP-response hints honored)
        const ropBufferSize: number = reader.readUInt32LE();
        if (ropBufferSize > MAX_ROP_BUFFER_SIZE || ropBufferSize > reader.remaining) {
            throw new RangeError(`Execute: invalid RopBufferSize ${ropBufferSize}.`);
        }
        const ropBuffer = decodeRopBuffer(reader.readBytes(ropBufferSize));
        const maxRopOut = reader.remaining >= 4 ? reader.readUInt32LE() : MAX_ROPS_LIST_BYTES + 2;
        return { ...ropBuffer, maxRopOut };
    } catch {
        throw new ApiError(ApiErrors.INVALID_REQUEST, 400, ApiErrorMessages.INVALID_REQUEST);
    }
}

function firstHeader(req: HttpRequest, name: string): string | undefined {
    const value: string | string[] | undefined = req.headers[name];
    return Array.isArray(value) ? value[0] : value;
}

/**
 * Abstract base for the single fixed EMSMDB endpoint (`POST /mapi/emsmdb` by `[MS-OXCMAPIHTTP]` convention,
 * though the concrete path is left to the consuming application to mount via `@Route(...)` - see
 * `BaseEasRoute.ts` for the identical undecorated-base-class pattern this follows). Like EAS, MAPI/HTTP
 * multiplexes several request types against one URL - here via the `X-RequestType` **header**
 * (`Connect`/`Execute`/`Disconnect`/`NotificationWait`) rather than a query parameter - so there is exactly
 * one `@Post()` method, not one per request type.
 *
 * **Auth**: `@Auth(["jwt"])` - unchanged from every other route in this app. Real, modern Exchange Server's
 * own MAPI virtual directory genuinely supports `OAuth` as a configured `IISAuthenticationMethods` value
 * alongside `NTLM`/`Negotiate`, so Bearer-token auth on this exact endpoint is real current Exchange
 * behavior, not a deviation this library invents - see the architecture plan's "Auth" section.
 *
 * **Connect/Disconnect are complete**: `Connect` establishes a real `MapiSessionContext` (via
 * `MapiSessionManager`) and sets the two spec-fixed session cookies (`MapiContext`/`MapiSequence`);
 * `Disconnect` releases the session. **`Execute` dispatches real ROPs** (starting with `RopLogon`/`RopRelease`
 * - build-order step 4) via `RopDispatcher`, using whichever `RopHandler`s `ropHandlerClasses` registers;
 * an unrecognized `RopId` simply stops processing (see `RopDispatcher`'s own doc comment for why), the same
 * "real, functional, no commands implemented yet" stance `BaseEasRoute`'s own skeleton step took for its
 * empty `commandHandlerClasses` before `ProvisionCommand` landed - here scoped per-ROP instead of per-request.
 *
 * **Always non-chunked**: every response here uses `Content-Length` (via `res.status(200).send(buffer)`),
 * never `Transfer-Encoding: chunked` - both are equally spec-valid per `[MS-OXCMAPIHTTP]`'s own "Common
 * Response Format", and skipping the `PROCESSING`/`PENDING`/`DONE` keep-alive meta-tag streaming matches
 * every other route in this library's own response idiom. See the architecture plan's "Protocol facts"
 * section for the full reasoning and its documented limitation.
 *
 * `mailboxClass`/`folderClass` are supplied by the Mongo/SQL concrete subclasses, following the exact
 * one-line-per-backend pattern used throughout this library's other routes/jobs. `folderRepo` is built once
 * here (not per-handler, unlike EAS's per-command repo pattern) and threaded through `RopContext` to every
 * `RopHandler` - see `RopHandler.ts`'s own doc comment for why.
 *
 * @author Jean-Philippe Steinmetz
 */
export abstract class BaseMapiEmsmdbRoute<M extends Mailbox> {
    protected abstract mailboxClass: any;
    protected abstract folderClass: any;
    protected abstract messageClass: any;
    protected abstract calendarEventClass: any;
    protected abstract contactClass: any;
    protected abstract taskClass: any;
    protected abstract labelClass: any;
    /** The backend's `AuditLogEntry` class. When set, messages deleted through MAPI are audited like REST deletes. */
    protected auditLogClass?: any;

    /** ROP handler classes to instantiate (one each) in `@Init`, keyed by their own `ropId`. Empty until a
     * concrete `RopHandler` lands - every ROP is then simply left unprocessed (see `RopDispatcher`'s own doc
     * comment), the correct, honest behavior for a transport skeleton with no ROPs implemented yet. */
    protected ropHandlerClasses: any[] = [];

    // Automatically injected by ObjectFactory on instantiation
    private _objectFactory?: ObjectFactory;

    private mailboxRepo?: RepoUtils<M>;
    private folderRepo?: RepoUtils<Folder>;
    private messageRepo?: RepoUtils<any>;
    private calendarEventRepo?: RepoUtils<any>;
    private contactRepo?: RepoUtils<any>;
    private taskRepo?: RepoUtils<any>;
    private labelRepo?: RepoUtils<any>;
    private sessionManager?: MapiSessionManager;
    private readonly ropHandlers = new Map<number, RopHandler>();

    @Inject("BlobStore")
    private blobStore?: BlobStore;

    @Inject(ScanPipeline)
    private scanPipeline?: ScanPipeline;

    @Inject("MailTransport")
    private mailTransport?: any;

    /** Publishes the live `Folder` counts-update event `refreshFolderCounts()` sends after a ROP handler changes what
     * a folder holds through `context.messageRepo` directly - see `folderCountsContext()`/`RopContext.notifyFolderCounts`. */
    @Inject(NotificationUtils)
    private notificationUtils?: NotificationUtils;

    @Logger
    private logger: any;

    @Config()
    private config?: any;

    @Init
    public async init(): Promise<void> {
        this.mailboxRepo = await this._objectFactory!.newInstance(RepoUtils, {
            name: this.mailboxClass.name,
            args: [this.mailboxClass],
        });
        this.folderRepo = await this._objectFactory!.newInstance(RecoverableRepoUtils, {
            name: this.folderClass.name,
            args: [this.folderClass],
        });
        this.messageRepo = await this._objectFactory!.newInstance(RecoverableRepoUtils, {
            name: this.messageClass.name,
            args: [this.messageClass],
        });
        this.calendarEventRepo = await this._objectFactory!.newInstance(RecoverableRepoUtils, {
            name: this.calendarEventClass.name,
            args: [this.calendarEventClass],
        });
        this.contactRepo = await this._objectFactory!.newInstance(RecoverableRepoUtils, {
            name: this.contactClass.name,
            args: [this.contactClass],
        });
        this.taskRepo = await this._objectFactory!.newInstance(RecoverableRepoUtils, {
            name: this.taskClass.name,
            args: [this.taskClass],
        });
        // Label extends plain BaseEntity, not RecoverableBaseEntity (see restapi's own model doc comment) - no
        // soft-delete support to preserve, so a plain RepoUtils is correct here, matching mailboxRepo's own
        // choice for the identical reason.
        this.labelRepo = await this._objectFactory!.newInstance(RepoUtils, {
            name: this.labelClass.name,
            args: [this.labelClass],
        });
        this.sessionManager = await this._objectFactory!.newInstance(MapiSessionManager);
        for (const HandlerClass of this.ropHandlerClasses) {
            const handler: RopHandler = await this._objectFactory!.newInstance(HandlerClass);
            this.ropHandlers.set(handler.ropId, handler);
        }
    }

    @Auth(["jwt"])
    @Post()
    public async dispatch(
        @Request req: HttpRequest,
        @Response res: HttpResponse,
        @AuthUser user?: JWTUser,
    ): Promise<void> {
        if (
            !this.mailboxRepo ||
            !this.folderRepo ||
            !this.messageRepo ||
            !this.calendarEventRepo ||
            !this.contactRepo ||
            !this.taskRepo ||
            !this.labelRepo ||
            !this.sessionManager ||
            !this.blobStore ||
            !this.scanPipeline ||
            !this.mailTransport
        ) {
            throw new ApiError(ApiErrors.INTERNAL_ERROR, 500, ApiErrorMessages.INTERNAL_ERROR);
        }
        if (!user) {
            throw new ApiError(ApiErrors.AUTH_PERMISSION_FAILURE, 403, ApiErrorMessages.AUTH_PERMISSION_FAILURE);
        }

        const requestType: string | undefined = firstHeader(req, "x-requesttype");
        const clientInfo: string = firstHeader(req, "x-clientinfo") ?? "";
        const requestId: string = firstHeader(req, "x-requestid") ?? "";

        res.setHeader("Content-Type", "application/mapi-http")
            .setHeader("X-RequestType", requestType ?? "")
            .setHeader("X-RequestId", requestId)
            .setHeader("X-ResponseCode", "0")
            .setHeader("X-ClientInfo", clientInfo)
            .setHeader("X-ServerApplication", "RapidREST-Mail");

        switch (requestType) {
            case "Connect":
                await this.handleConnect(req, res, user);
                return;
            case "Execute":
                await this.handleExecute(req, res, user);
                return;
            case "Disconnect":
                await this.handleDisconnect(req, res, user);
                return;
            case "NotificationWait":
                // Real, spec-defined request type, deliberately deferred - no push-notification support in
                // this pass (see the architecture plan's ROP scope table). Distinct from an unrecognized
                // X-RequestType value (400 below), the same distinction BaseEasRoute draws between 501
                // ("recognized but deferred") and 400 ("malformed request").
                res.status(501).send();
                return;
            default:
                throw new ApiError(ApiErrors.INVALID_REQUEST, 400, ApiErrorMessages.INVALID_REQUEST);
        }
    }

    /**
     * `Connect` establishes a new Session Context - never trusts the request body's own `UserDn` field for
     * mailbox identity (the same "never a client-supplied mailbox, always resolved from the authenticated
     * JWT" principle `BaseEasRoute`/`BaseSearchRoute` already apply via `resolveCallerMailboxUid`) - so the
     * request body is never decoded at all here.
     */
    private async handleConnect(req: HttpRequest, res: HttpResponse, user: JWTUser): Promise<void> {
        const mailboxUid: string | undefined = await resolveCallerMailboxUid(this.mailboxRepo!, user);
        if (!mailboxUid) {
            throw new ApiError(ApiErrors.NOT_FOUND, 404, ApiErrorMessages.NOT_FOUND);
        }
        const mailbox: M | undefined = await this.mailboxRepo!.findOne(mailboxUid, { ignoreACL: true });

        const session: MapiSessionContext = await this.sessionManager!.create(mailboxUid, user.uid);
        res.appendHeader("Set-Cookie", `MapiContext=${session.uid}`);
        res.appendHeader("Set-Cookie", `MapiSequence=0`);

        const body = new BufferWriter();
        body.writeUInt32LE(0); // StatusCode
        body.writeUInt32LE(0); // ErrorCode
        body.writeUInt32LE(60000); // PollsMax (ms) - pragmatic constant, not a tuned server policy value
        body.writeUInt32LE(3); // RetryCount - pragmatic constant
        body.writeUInt32LE(5000); // RetryDelay (ms) - pragmatic constant
        // DnPrefix: this library addresses mailboxes by UID/SMTP-address, never real X.500 DNs, so this is a
        // fixed, unused placeholder rather than a real DN prefix a client could build recipients from.
        body.writeNullTerminatedString8("/");
        body.writeNullTerminatedUtf16LE(mailbox?.displayName ?? "");
        body.writeUInt32LE(0); // AuxiliaryBufferSize
        res.status(200).send(body.toBuffer());
    }

    /**
     * `Execute` decodes the outer envelope (`Flags`/`RopBufferSize`/`RopBuffer`/`MaxRopOut`/...) and the inner ROP
     * buffer framing (`RopBuffer.ts`), dispatches every contained ROP via `RopDispatcher`, and re-encodes the
     * collected responses - preserving the incoming `handleTable` unchanged (this pragmatic subset never
     * allocates/frees table-wide handle slots at the framing level; individual `RopHandler`s manage their own
     * entries within `session.handles` instead).
     *
     * - **One request at a time per session.** Only the caller's own session (whose mailbox is still the caller's) is
     * locked. The lock is taken before any ROP runs and renewed every `SESSION_LOCK_RENEW_MS` while the request runs,
     * so a slow request keeps it; an overlapping request gets `X-ResponseCode` 15 (Invalid Sequence) without any side
     * effects. The save stays compare-and-set. Each `Execute` also refreshes the session's place in the per-user cap.
     * - **Retries.** While a request with an `X-RequestId` runs, an in-progress marker for that id is kept (renewed with
     * the lock); a retry of it is answered with code 15 instead of running again. Its response then replaces the
     * marker, and a request repeating the last id (a client retrying after a lost response) gets that response again
     * instead of running its ROPs twice. A request that ends without a stored response drops its marker.
     * `MapiSequence` isn't validated - the lock and the replay cache cover what it guards against here.
     * - **Output size.** ROP responses are held to the client's `MaxRopOut`, less the `RopSize` field and handle
     * table, and to what a 16-bit `RopSize` can describe (see `dispatchRops`). Room for a `RopBufferTooSmall` covering
     * the rest of the request is kept free as ROPs run, so a response that doesn't fit is always answered that way (a
     * normal, stored response). Only a `MaxRopOut` too small for a `RopBufferTooSmall` of the whole request answers
     * `ecBufferTooSmall` in the body `ErrorCode`, before any ROP runs, so nothing is stored and a retry changes nothing.
     * - **Released handles.** Once the session is saved, the FastTransfer streams and write-stream chunks of handles
     * the request released are deleted from the shared store; `Disconnect` and session eviction delete a session's
     * whole set.
     * - **Malformed ROPs.** A ROP that can't be decoded answers 400; the ROPs before it already ran, so the session is
     * still saved first.
     */
    private async handleExecute(req: HttpRequest, res: HttpResponse, user: JWTUser): Promise<void> {
        const sessionId: string | undefined = req.cookies["MapiContext"];
        if (!sessionId) {
            this.sendExecuteFailure(res, RESPONSE_CODE_CONTEXT_NOT_FOUND, ERROR_SESSION_NOT_FOUND);
            return;
        }
        // Ownership first: someone else's session (or a mailbox reassigned since Connect) is never locked by this caller.
        const owned: MapiSessionContext | undefined = await this.loadOwnSession(req, user);
        if (!owned || (await resolveCallerMailboxUid(this.mailboxRepo!, user)) !== owned.mailboxUid) {
            this.sendExecuteFailure(res, RESPONSE_CODE_CONTEXT_NOT_FOUND, ERROR_SESSION_NOT_FOUND);
            return;
        }
        const lockToken: string | undefined = await this.sessionManager!.acquireLock(sessionId);
        if (!lockToken) {
            this.sendExecuteFailure(res, RESPONSE_CODE_INVALID_SEQUENCE, RESPONSE_CODE_INVALID_SEQUENCE);
            return;
        }
        const running: { requestId?: string } = {};
        // Renewed while the request runs, so a slow Execute never loses its lock (or its in-progress marker) to a retry.
        // Renewals run one after another, and stop once the lock is lost. The marker is only extended while it is still
        // this request's marker, so a renewal landing after the response was stored never turns it back into a marker.
        let renewing: Promise<void> = Promise.resolve();
        const renewal = setInterval(() => {
            renewing = renewing.then(async () => {
                const held: boolean = await this.sessionManager!.renewLock(sessionId, lockToken).catch(() => true);
                if (!held) {
                    clearInterval(renewal);
                    return;
                }
                if (running.requestId) {
                    await this.sessionManager!.renewInProgress(sessionId, running.requestId).catch(() => undefined);
                }
            });
        }, SESSION_LOCK_RENEW_MS);
        renewal.unref();
        try {
            await this.executeLocked(req, res, user, running);
        } finally {
            clearInterval(renewal);
            await renewing;
            if (running.requestId) {
                await this.sessionManager!.clearInProgress(sessionId, running.requestId).catch(() => undefined);
            }
            await this.sessionManager!.releaseLock(sessionId, lockToken);
        }
    }

    /** Runs an `Execute` under its session's lock. `running.requestId` is set while this request's in-progress marker is
     * stored, and cleared once its response is stored instead. */
    private async executeLocked(req: HttpRequest, res: HttpResponse, user: JWTUser, running: { requestId?: string }): Promise<void> {
        // Loaded again under the lock: the copy read before locking may already be out of date.
        const session: MapiSessionContext | undefined = await this.loadOwnSession(req, user);
        if (!session) {
            this.sendExecuteFailure(res, RESPONSE_CODE_CONTEXT_NOT_FOUND, ERROR_SESSION_NOT_FOUND);
            return;
        }

        const requestId: string = firstHeader(req, "x-requestid") ?? "";
        const replay = requestId ? await this.sessionManager!.storedResponse(session.uid, requestId) : undefined;
        if (replay === "inProgress") {
            // The same request is still running elsewhere (its lock lapsed, or its pod died moments ago).
            this.sendExecuteFailure(res, RESPONSE_CODE_INVALID_SEQUENCE, RESPONSE_CODE_INVALID_SEQUENCE);
            return;
        }
        if (replay) {
            res.status(200).send(replay);
            return;
        }
        if (requestId) {
            await this.sessionManager!.markInProgress(session.uid, requestId);
            running.requestId = requestId;
        }
        await this.sessionManager!.touch(session).catch(() => undefined);

        const { ropsList, handleTable, maxRopOut } = decodeExecuteRequest(req.rawBody ?? Buffer.alloc(0));
        const context: RopContext = {
            mailboxUid: session.mailboxUid,
            userUid: session.userUid,
            session,
            mailboxRepo: this.mailboxRepo!,
            folderRepo: this.folderRepo!,
            messageRepo: this.messageRepo!,
            calendarEventRepo: this.calendarEventRepo!,
            contactRepo: this.contactRepo!,
            taskRepo: this.taskRepo!,
            labelRepo: this.labelRepo!,
            folderClass: this.folderClass,
            messageClass: this.messageClass,
            calendarEventClass: this.calendarEventClass,
            contactClass: this.contactClass,
            taskClass: this.taskClass,
            labelClass: this.labelClass,
            blobStore: this.blobStore!,
            scanPipeline: this.scanPipeline!,
            mailTransport: this.mailTransport!,
            handleData: this.sessionManager!.handleDataStore,
            audit: this.auditLogClass
                ? (params) =>
                      recordAuditLog(this._objectFactory!, this.auditLogClass, { config: this.config, req, user, logger: this.logger }, params)
                : undefined,
            notifyFolderCounts: (folderUids, options) => refreshFolderCounts(this.folderCountsContext(), folderUids, options),
        };
        // RopSize (2 bytes) and the echoed handle table share MaxRopOut with the ROP responses. dispatchRops only throws
        // for a request it can't decode (a failing ROP gets a failure response instead).
        const dispatched: Buffer | { error: unknown } = await dispatchRops(ropsList, this.ropHandlers, context, {
            maxOutputBytes: maxRopOut - 2 - handleTable.length * 4,
        }).catch((error: unknown) => ({ error }));

        const saved = await this.sessionManager!.save(session);
        const released: string[] = takeReleasedHandleData(session);
        if (saved === "saved") {
            // Streams and write chunks of handles this request released; best effort, they expire anyway.
            const store: HandleDataStore = handleDataStoreOf(context);
            const owners = handleDataOwners(session.uid, session.userUid);
            await Promise.all(released.map((key) => store.delete(key, owners).catch(() => undefined)));
        }
        if (!Buffer.isBuffer(dispatched) && !(dispatched.error instanceof ExecuteBufferTooSmallError)) {
            throw dispatched.error instanceof DecodeError ? new ApiError(ApiErrors.INVALID_REQUEST, 400, ApiErrorMessages.INVALID_REQUEST) : dispatched.error;
        }
        if (saved !== "saved") {
            // conflict: another request saved first (possible once a lock has expired), so this request's handle
            // changes can't be kept and the client retries. missing: the session ended meanwhile. tooBig: see
            // MAX_SESSION_BYTES.
            const [responseCode, errorCode] =
                saved === "conflict"
                    ? [RESPONSE_CODE_INVALID_SEQUENCE, RESPONSE_CODE_INVALID_SEQUENCE]
                    : saved === "missing"
                      ? [RESPONSE_CODE_CONTEXT_NOT_FOUND, ERROR_SESSION_NOT_FOUND]
                      : [0, ERROR_SESSION_TOO_BIG];
            this.sendExecuteFailure(res, responseCode, errorCode);
            return;
        }
        if (!Buffer.isBuffer(dispatched)) {
            // [MS-OXCRPC] ecBufferTooSmall: not even a RopBufferTooSmall of the whole request fits, so no ROP ran.
            this.sendExecuteFailure(res, 0, ERROR_BUFFER_TOO_SMALL);
            return;
        }
        const responseRopsList: Buffer = dispatched;

        const responseRopBuffer: Buffer = encodeRopBuffer({ ropsList: responseRopsList, handleTable });
        const body = new BufferWriter();
        body.writeUInt32LE(0); // StatusCode
        body.writeUInt32LE(0); // ErrorCode
        body.writeUInt32LE(0); // Flags
        body.writeUInt32LE(responseRopBuffer.length);
        body.writeBytes(responseRopBuffer);
        body.writeUInt32LE(0); // AuxiliaryBufferSize
        const responseBody: Buffer = body.toBuffer();
        if (requestId) {
            await this.sessionManager!.storeResponse(session.uid, requestId, responseBody);
            delete running.requestId;
        }
        res.status(200).send(responseBody);
    }

    /** What `refreshFolderCounts()` (`@rapidmx/restapi`) needs to recompute, re-cache and publish a folder's counts -
     * see `RopContext.notifyFolderCounts`'s own doc comment for why a ROP handler must call it after adding, deleting
     * or moving a `Message` row through `context.messageRepo` itself. */
    private folderCountsContext(): FolderCountsContext {
        return {
            messageRepo: this.messageRepo!,
            folderRepo: this.folderRepo!,
            folderClass: this.folderClass,
            notificationUtils: this.notificationUtils,
            logger: this.logger,
        };
    }

    /** Writes an `Execute` failure body (no ROP buffer) with the given `X-ResponseCode` and `ErrorCode`. */
    private sendExecuteFailure(res: HttpResponse, responseCode: number, errorCode: number): void {
        res.setHeader("X-ResponseCode", String(responseCode));
        const body = new BufferWriter();
        body.writeUInt32LE(0); // StatusCode
        body.writeUInt32LE(errorCode); // ErrorCode
        body.writeUInt32LE(0); // Flags
        body.writeUInt32LE(0); // RopBufferSize
        body.writeUInt32LE(0); // AuxiliaryBufferSize
        res.status(200).send(body.toBuffer());
    }

    /**
     * Loads the session named by the `MapiContext` cookie, but only if the authenticated caller is the user who
     * `Connect`ed it. The cookie value alone is a bearer token for someone else's mailbox otherwise, so a session
     * belonging to another user is treated exactly like one that doesn't exist.
     */
    private async loadOwnSession(req: HttpRequest, user: JWTUser): Promise<MapiSessionContext | undefined> {
        const sessionId: string | undefined = req.cookies["MapiContext"];
        const session: MapiSessionContext | undefined = sessionId ? await this.sessionManager!.load(sessionId) : undefined;
        return session && session.userUid === user.uid ? session : undefined;
    }

    private async handleDisconnect(req: HttpRequest, res: HttpResponse, user: JWTUser): Promise<void> {
        // Another user's session is left alone, with the same success response as for an unknown session.
        const session: MapiSessionContext | undefined = await this.loadOwnSession(req, user);
        if (session) {
            await this.sessionManager!.destroy(session.uid);
        }
        const body = new BufferWriter();
        body.writeUInt32LE(0); // StatusCode
        body.writeUInt32LE(0); // ErrorCode
        body.writeUInt32LE(0); // AuxiliaryBufferSize
        res.status(200).send(body.toBuffer());
    }
}
