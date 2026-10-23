///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { simpleParser } from "mailparser";
import type { RepoUtils } from "@rapidrest/service-core";
import { Message, type BlobStore } from "@rapidmx/restapi";
import { BufferWriter } from "../codec/BufferCursor.js";
import type { MapiObjectHandle } from "../MapiSessionManager.js";
import type { ExecuteBudget } from "./ExecuteBudget.js";
import { handleDataCache, handleDataKey } from "./HandleDataCache.js";
import type { RopContext } from "./RopHandler.js";

/** `PidTagBody` (`[MS-OXPROPS]`): property ID `0x1000`, `PtypString`. The only streamable property this
 * pragmatic subset supports via `RopOpenStream`/`RopReadStream` - `PidTagHtml`/`PidTagRtfCompressed` are a
 * documented gap (see `RopOpenStreamHandler`'s own doc comment). */
export const PID_TAG_BODY = 0x1000;

/** The body Outlook shows for a zero-access message (@rapidmx/restapi `specs/search.md` §3): mail the server encrypted at rest on
 * arrival and can't read. Serving it as S/MIME needs MAPI attachment objects, which this server doesn't implement yet, so Outlook over
 * MAPI gets this explanation instead (ActiveSync serves the S/MIME message itself). */
export const ZERO_ACCESS_PLACEHOLDER_TEXT =
    "This message is encrypted. Read it in RapidMX, or in a mail app with your RapidMX encryption keys installed.";

/**
 * Resolves a `"message:<uid>"` target's plain-text body as the exact byte sequence `RopReadStream` must serve
 * for a streamed `PtypString` property - confirmed via `[MS-OXCPRPT]`'s `RopOpenStream` page ("a string of
 * Unicode characters in UTF-16LE format encoding with terminating null character"), the identical encoding
 * `writePropertyValue()` already uses for an inline `PtypString` value, so this reuses
 * `BufferWriter.writeNullTerminatedUtf16LE` rather than reimplementing it.
 *
 * Always derived from the message's raw MIME source (`bodyBlobKey`, via `mailparser` - the same fallback path
 * `ItemOperationsCommand.fetchMessage()` already uses for EAS's own plain-text body delivery) rather than
 * `sanitizedHtmlBlobKey` - `PidTagBody` is specifically the plain-text body per `[MS-OXPROPS]`, unlike EAS's
 * `Body` element, which can carry either format tagged by its own `Type` field.
 *
 * Does no caching itself - `loadStreamBody` below is what the stream ROPs call. With a `budget`, nothing is fetched
 * once its bytes are used up, and the raw message's size is charged before it is parsed.
 */
export async function resolveMessageBodyBytes(target: string, messageRepo: RepoUtils<any>, blobStore: BlobStore, budget?: ExecuteBudget): Promise<Buffer> {
    budget?.assertBytesLeft();
    const uid = target.slice("message:".length);
    const message: Message | undefined = await messageRepo.findOne(uid, { ignoreACL: true });
    if (!message) {
        return Buffer.alloc(0);
    }
    // Read structurally: `Message.zeroAccess` is newer than the @rapidmx/restapi this builds against.
    if ((message as { zeroAccess?: unknown }).zeroAccess) {
        const writer = new BufferWriter();
        writer.writeNullTerminatedUtf16LE(ZERO_ACCESS_PLACEHOLDER_TEXT);
        return writer.toBuffer();
    }
    const raw = await blobStore.get(message.bodyBlobKey);
    budget?.chargeBytes(raw.length);
    const parsed = await simpleParser(raw);
    const writer = new BufferWriter();
    writer.writeNullTerminatedUtf16LE(parsed.text ?? "");
    return writer.toBuffer();
}

/**
 * The body bytes for the read stream at `handleIndex`. Parsed once and kept in `HandleDataCache` (outside the
 * session JSON) for the life of the handle, so reading a large body in many `RopReadStream` chunks costs one
 * fetch and one MIME parse instead of one per chunk. A cache miss (eviction, or a request served by another
 * replica) just resolves the body again: a stored message's body never changes, so the bytes are the same.
 *
 * Within one `Execute`, streams opened on the same message share a single parse (`ExecuteBudget.bodies`), and each
 * parse's bytes count against the request's byte budget.
 */
export async function loadStreamBody(context: RopContext, handleIndex: number, stream: MapiObjectHandle): Promise<Buffer> {
    const key = handleDataKey(context.session.uid, handleIndex, stream.generation);
    const cached = handleDataCache.get(key);
    if (cached) {
        return cached;
    }
    const budget = context.budget;
    let pending = budget?.bodies.get(stream.entityUid);
    if (!pending) {
        budget?.assertBytesLeft();
        pending = resolveMessageBodyBytes(stream.entityUid, context.messageRepo, context.blobStore, budget);
        budget?.bodies.set(stream.entityUid, pending);
        budget?.chargeBytes((await pending).length);
    }
    const bytes = await pending;
    handleDataCache.set(key, bytes);
    return bytes;
}
