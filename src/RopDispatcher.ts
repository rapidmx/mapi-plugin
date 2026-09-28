///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { BufferReader, BufferWriter, DecodeError } from "./codec/BufferCursor.js";
import { ExecuteBudget, WorkBudgetExceededError } from "./rop/ExecuteBudget.js";
import type { RopContext, RopHandler } from "./rop/RopHandler.js";

/** The most ROPs processed from one `Execute` request. A real client batches far fewer; anything past this is
 * left unprocessed. */
export const MAX_ROPS_PER_EXECUTE = 1024;

/** The largest `RopsList` a response can carry: `RopSize` is a 16-bit count that includes its own 2 bytes. */
export const MAX_ROPS_LIST_BYTES = 0xffff - 2;

/** `RopBufferTooSmall`'s `RopId` (`[MS-OXCROPS]` §2.2.15.1). */
export const ROP_ID_BUFFER_TOO_SMALL = 0xff;

/** `MAPI_E_CALL_FAILED`: a ROP failed while running (a database error, undecodable stored data, ...). */
export const ERROR_CALL_FAILED = 0x80004005;

/** `MAPI_E_TOO_COMPLEX`: the ROP would have gone past this request's `ExecuteBudget`. */
export const ERROR_TOO_COMPLEX = 0x80040117;

/** `RopBufferTooSmall`'s fixed fields: `RopId` and `SizeNeeded`. */
const BUFFER_TOO_SMALL_HEADER_BYTES = 3;

/** Thrown by `dispatchRops`, before any ROP runs, when the room for responses can't even hold a `RopBufferTooSmall`
 * carrying the whole request. The route answers the `Execute` with `ecBufferTooSmall` ([MS-OXCRPC]). No ROP has run. */
export class ExecuteBufferTooSmallError extends Error {
    public constructor() {
        super("Execute: MaxRopOut is too small for even a RopBufferTooSmall response.");
        this.name = "ExecuteBufferTooSmallError";
    }
}

export interface DispatchOptions {
    /** Stop after this many ROPs. Defaults to `MAX_ROPS_PER_EXECUTE`. */
    maxRops?: number;
    /** The room for response ROPs, in bytes. Defaults to `MAX_ROPS_LIST_BYTES`; the route passes what the client's
     * `MaxRopOut` leaves after the `RopSize` field and handle table. */
    maxOutputBytes?: number;
}

/** Wraps `reader` so that running out of request bytes raises a `DecodeError` (a malformed request) rather than a
 * plain `RangeError`, which a handler may also raise while working on data it read from somewhere else. */
function decodingReader(reader: BufferReader): BufferReader {
    return new Proxy(reader, {
        get(target, property) {
            const value = Reflect.get(target, property, target);
            if (typeof value !== "function") {
                return value;
            }
            return (...args: unknown[]) => {
                try {
                    return value.apply(target, args);
                } catch (err) {
                    throw err instanceof RangeError && !(err instanceof DecodeError) ? new DecodeError(err.message) : err;
                }
            };
        },
    });
}

/**
 * Walks a decoded `RopBuffer`'s `ropsList` blob, dispatching each contained ROP - identified by its own leading
 * `RopId` byte - to the matching registered handler in turn, collecting every response into one output buffer.
 * Deliberately ROP-agnostic: it knows nothing about any specific ROP's own fields, only how to find the *next* one,
 * via the handler's own advanced `reader` position once `handle()` returns - see `RopBuffer.ts`'s doc comment for
 * why a fully generic split isn't possible for this wire format.
 *
 * - **Malformed requests.** An unrecognized `RopId`, or a ROP whose bytes run out or can't be sized (a
 * `DecodeError`), throws a `DecodeError`: nothing after it can be located, so the route answers 400.
 * - **A failing ROP.** Any other error thrown by a handler fails just that ROP: its partial output is dropped, a
 * failure response (`RopId`, the handle index it echoes, `ReturnValue` `MAPI_E_CALL_FAILED` or, past the
 * `ExecuteBudget`, `MAPI_E_TOO_COMPLEX`) takes its place, and the following ROPs still run. The handler has
 * already read its whole request by the time it does work that can fail, so the next ROP is still where the
 * reader points.
 * - **Output space.** Room for a `RopBufferTooSmall` covering the next ROP and everything after it (3 bytes plus
 * those request bytes) is always kept free. `context.ropOutputRemaining` tells each handler the room left *after*
 * that reserve (computed from where the reader stands, so once a handler has read its request it excludes the
 * handler's own bytes). A response that doesn't leave the reserve is replaced by `RopBufferTooSmall` (`SizeNeeded`
 * plus the unprocessed request bytes, from this ROP on, for the client to resend) and processing stops; that always
 * fits, because the reserve for this ROP was kept by the one before it. Handlers that change state by how much they
 * return (`RopQueryRows`, `RopReadStream`, `RopFastTransferSourceGetBuffer`) size their response to the room, so they
 * are never replaced after moving a cursor. Only when the request as a whole can't be covered by a
 * `RopBufferTooSmall` is `ExecuteBufferTooSmallError` thrown, before anything runs, and the route answers the
 * `Execute` with `ecBufferTooSmall` ([MS-OXCRPC]). So a failed `Execute` never follows state already changed and
 * saved, which a retry would skip past.
 *
 * @author Jean-Philippe Steinmetz
 */
export async function dispatchRops(
    ropsList: Buffer,
    handlers: Map<number, RopHandler>,
    context: RopContext,
    options: DispatchOptions = {},
): Promise<Buffer> {
    const maxRops = options.maxRops ?? MAX_ROPS_PER_EXECUTE;
    const maxOutputBytes = Math.min(options.maxOutputBytes ?? MAX_ROPS_LIST_BYTES, MAX_ROPS_LIST_BYTES);
    const reader = decodingReader(new BufferReader(ropsList));
    const writer = new BufferWriter();
    context.budget ??= new ExecuteBudget();
    /** Room for a RopBufferTooSmall resending the request from `offset` on (nothing when there is nothing left). */
    const reserveFrom = (offset: number): number => (offset < ropsList.length ? BUFFER_TOO_SMALL_HEADER_BYTES + ropsList.length - offset : 0);
    if (reserveFrom(0) > maxOutputBytes) {
        throw new ExecuteBufferTooSmallError();
    }
    let remaining = maxOutputBytes;
    Object.defineProperty(context, "ropOutputRemaining", {
        configurable: true,
        enumerable: true,
        get: () => Math.max(0, remaining - reserveFrom(reader.position)),
    });
    let processed = 0;
    try {
        // Stops after `maxRops`, so one Execute can't queue unbounded work.
        while (reader.hasMore() && processed++ < maxRops) {
            const start = reader.position;
            const ropId = reader.readUInt8();
            const handler = handlers.get(ropId);
            if (!handler) {
                throw new DecodeError(`RopDispatcher: unsupported RopId 0x${ropId.toString(16)}.`);
            }
            remaining = maxOutputBytes - writer.length;

            let response: Buffer;
            try {
                const ropWriter = new BufferWriter();
                await handler.handle(reader, ropWriter, context);
                response = ropWriter.toBuffer();
            } catch (err) {
                if (err instanceof DecodeError) {
                    throw err;
                }
                // TEMPORARY DIAGNOSTIC LOGGING - see NOTES.md's 2026-09-28 "diagnostic release" entry. A ROP
                // handler's real error is otherwise completely swallowed - the client only ever sees a bare
                // MAPI_E_CALL_FAILED/MAPI_E_TOO_COMPLEX, with nothing server-side recording why. Remove once the
                // Outlook desktop "set of folders cannot be opened" root cause is confirmed.
                console.error(`MAPI_DEBUG ROP 0x${ropId.toString(16)} (${handler.constructor.name}) threw:`, err);
                response = failureResponse(handler, ropsList, start, err instanceof WorkBudgetExceededError ? ERROR_TOO_COMPLEX : ERROR_CALL_FAILED);
            }

            if (response.length + reserveFrom(reader.position) > remaining) {
                // The reserve kept for this ROP guarantees this fits.
                writer
                    .writeUInt8(ROP_ID_BUFFER_TOO_SMALL)
                    .writeUInt16LE(Math.min(response.length, 0xffff)) // SizeNeeded
                    .writeBytes(ropsList.subarray(start)); // RequestBuffers
                break;
            }
            writer.writeBytes(response);
        }
    } finally {
        delete context.ropOutputRemaining;
    }
    return writer.toBuffer();
}

/** The response a ROP gets when its handler throws: `RopId`, the handle index its responses echo, `ReturnValue`,
 * then any always-present fields zeroed. Nothing at all for a ROP that never responds. */
function failureResponse(handler: RopHandler, ropsList: Buffer, start: number, returnValue: number): Buffer {
    if (handler.hasNoResponse) {
        return Buffer.alloc(0);
    }
    return new BufferWriter()
        .writeUInt8(handler.ropId)
        .writeUInt8(ropsList[start + (handler.responseHandleIndexOffset ?? 2)] ?? 0)
        .writeUInt32LE(returnValue)
        .writeBytes(Buffer.alloc(handler.failureTailBytes ?? 0))
        .toBuffer();
}
