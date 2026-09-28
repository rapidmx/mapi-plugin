///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import crypto from "crypto";
import type { JWTUser } from "@rapidrest/core";
import type { HttpResponse, RepoUtils } from "@rapidrest/service-core";
import { BufferWriter } from "../codec/BufferCursor.js";
import { encodeGuid } from "../codec/MapiGuid.js";
import { resolveCallerMailboxUid } from "@rapidmx/restapi";

/** `[MS-OXNSPI]` §2.2.1.2's own `MAPI_E_NOT_FOUND` error code, reused for "the caller has no mailbox in this
 * deployment" - the one real failure mode `Bind` can hit in this pragmatic subset. */
const ERROR_NOT_FOUND = 0x8004010f;

/**
 * `Bind`/`Unbind` request types (`[MS-OXCMAPIHTTP]` §2.2.5.1/§2.2.5.2, backed by `[MS-OXNSPI]`'s own `NspiBind`/
 * `NspiUnbind` methods): establishes/tears down an NSPI Session Context, the address-book endpoint's own
 * analog of EMSMDB's `Connect`/`Disconnect`.
 *
 * **No real session state is created or tracked** - a deliberate, documented simplification, not an oversight:
 * every subsequent NSPI operation in this pragmatic subset (`GetMatches`) is independently `@Auth(["jwt"])`-
 * protected and re-resolves the caller's own mailbox from that same JWT, exactly the way `RopLogon`/every ROP
 * handler already does for EMSMDB - the real spec's `MapiContext`-equivalent session cookie exists mainly for
 * multi-server-farm request affinity, not additional authentication, so this pragmatic subset mints an opaque,
 * unvalidated cookie value purely for wire-format conformance and never checks it again. `HasState`/`State`
 * (an optional `STAT` scoping the bind to one address-book container) is accepted but not decoded at all -
 * this pragmatic subset has exactly one "container" (a mailbox's own `Contact` list) regardless.
 *
 * @author Jean-Philippe Steinmetz
 */
export async function handleNspiBind(res: HttpResponse, user: JWTUser, mailboxRepo: RepoUtils<any>): Promise<void> {
    const mailboxUid = await resolveCallerMailboxUid(mailboxRepo, user);
    if (!mailboxUid) {
        const body = new BufferWriter();
        body.writeUInt32LE(ERROR_NOT_FOUND); // StatusCode - MUST NOT be 0 on failure
        body.writeUInt32LE(0); // AuxiliaryBufferSize
        res.status(200).send(body.toBuffer());
        return;
    }

    // Path/Secure/SameSite=None/HttpOnly mirror BaseMapiEmsmdbRoute.handleConnect()'s identical fix - see its
    // own comment for why (a bare cookie is never sent back by Outlook; this decorated shape is what a real
    // Exchange Online capture showed actually working).
    res.appendHeader("Set-Cookie", `NspiContext=${crypto.randomUUID()}; Path=/mapi/nspi; Secure; HttpOnly; SameSite=None`); // opaque, unvalidated - see class doc comment

    const body = new BufferWriter();
    body.writeUInt32LE(0); // StatusCode - success
    body.writeUInt32LE(0); // ErrorCode - success
    body.writeBytes(encodeGuid(crypto.randomUUID())); // ServerGuid - a fresh value per Bind is harmless, see class doc comment
    body.writeUInt32LE(0); // AuxiliaryBufferSize
    res.status(200).send(body.toBuffer());
}

/** `Unbind` never fails and has nothing to tear down - see `handleNspiBind`'s own doc comment for why no real
 * session exists to release. */
export function handleNspiUnbind(res: HttpResponse): void {
    const body = new BufferWriter();
    body.writeUInt32LE(0); // StatusCode
    body.writeUInt32LE(0); // ErrorCode
    body.writeUInt32LE(0); // AuxiliaryBufferSize
    res.status(200).send(body.toBuffer());
}
