///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { BufferReader, BufferWriter } from "../codec/BufferCursor.js";
import { resolveContentsWindow } from "./ContentsTable.js";
import { resolvePropertyValues, ResolutionCache, writePropertyValueSafely } from "./PropertyResolvers.js";
import type { RopContext, RopHandler } from "./RopHandler.js";

const ROP_ID_QUERY_ROWS = 0x15;

/** The well-known MAPI HRESULT `MAPI_E_INVALID_OBJECT`, reused for "the referenced handle isn't a table (or
 * doesn't exist)" - the same constant `RopGetHierarchyTableHandler` uses for its own analogous check. */
const ERROR_INVALID_OBJECT = 0x80070005;

/** `BOOKMARK_END` (`[MS-OXCTABL]`'s `Origin` field) - this pragmatic subset has no separate seek/bookmark ROP
 * support, so every response reports the cursor as having landed at the table's current end. */
const ORIGIN_BOOKMARK_END = 0x02;

/** The most rows one `RopQueryRows` returns. Fewer rows than requested is normal table paging: the client asks
 * again from the new cursor position. */
export const MAX_ROWS_PER_QUERY = 500;

/** `ecBufferTooSmall`: not even one row fits in the room left in the ROP output buffer. */
const ERROR_BUFFER_TOO_SMALL = 0x0000047d;

/** RopId, InputHandleIndex, ReturnValue, Origin and RowCount. */
const RESPONSE_HEADER_BYTES = 9;

/**
 * `RopQueryRows` (`[MS-OXCTABL]`/`[MS-OXCROPS]`): fetches up to `RowCount` rows from an already-configured
 * table (`RopGetHierarchyTable` + `RopSetColumns`), advancing the table's cursor. Confirmed field-by-field
 * against `[MS-OXCTABL]`'s own real captured request/response byte example - including the response's
 * `RowData` encoding (`PropertyRow`: a leading `Flags` byte, `0x00` for a `StandardPropertyRow` where every
 * column is a plain `PropertyValue`, `0x01` for a `FlaggedPropertyRow` where each column gets its own
 * `FlaggedPropertyValue` flag+optional-value). This pragmatic subset always emits `StandardPropertyRow`
 * (`Flags = 0x00`) - every configured column always resolves to *some* value here (falling back to a
 * type-appropriate zero/empty default for an unsupported property, never an error), so `FlaggedPropertyRow`'s
 * per-column error signaling is never needed. That "never an error" promise covers a client-requested
 * `propertyType` that doesn't match what a column's own `propertyId` resolves to, not just an entirely
 * unmodeled `propertyId` - `buildRow` writes each column through `PropertyResolvers.writePropertyValueSafely`,
 * which falls back to the same type-appropriate default there too, so one row's mismatched column degrades
 * instead of failing this whole call (and discarding every row already built in it).
 *
 * Backward reads (`ForwardRead = FALSE`) and the `NoAdvance` flag are decoded (to advance the reader
 * correctly) but not honored - this pragmatic subset's tables are simple forward-only cursors. Works generically
 * against either a `RopGetHierarchyTable` (folder rows, `"folder:"`/`"virtual:"` targets) or a
 * `RopGetContentsTable` (message rows, `"message:"` targets) handle - dispatching to `FolderTarget`'s or
 * `MessageTarget`'s resolver and property-value mapping by row-target prefix, since a single table handle only
 * ever holds one kind of row.
 *
 * **Known, deliberate performance limitation**: `resolvePropertyValues` still does one `repo.findOne()` per
 * row (message/calendarEvent/contact/task) rather than a single batched fetch for the whole page - a genuine
 * N+1 for a large `RowCount`. A real fix needs a batched "fetch all these uids" query, which this codebase's
 * `RepoUtils<T>` abstraction doesn't expose a backend-agnostic way to express (Mongo's `$in` and TypeORM's
 * `In()` aren't interchangeable at this call site, which has no idea which backend it's running against) - see
 * `ResolutionCache` for the narrower, safe win this pass *did* make (the redundant *identical* whole-mailbox-
 * folder-list and calendar-mailbox fetches that were happening once per row regardless of which distinct
 * entity each row was, now happen at most once per call).
 *
 * @author Jean-Philippe Steinmetz
 */
export class RopQueryRowsHandler implements RopHandler {
    public readonly ropId = ROP_ID_QUERY_ROWS;

    public async handle(reader: BufferReader, writer: BufferWriter, context: RopContext): Promise<void> {
        reader.readUInt8(); // LogonId - this pragmatic subset doesn't track multiple concurrent logons per session
        const inputHandleIndex: number = reader.readUInt8();
        reader.readUInt8(); // QueryRowsFlags - NoAdvance/EnablePackedBuffers not supported; always advances
        reader.readUInt8(); // ForwardRead - backward reads not supported; always reads forward
        const requestedCount: number = reader.readUInt16LE();

        const table = context.session.handles[inputHandleIndex];
        if (!table || table.type !== "table") {
            writer.writeUInt8(ROP_ID_QUERY_ROWS);
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt32LE(ERROR_INVALID_OBJECT);
            return;
        }

        const cursor: number = table.cursor ?? 0;
        const count: number = Math.min(requestedCount, MAX_ROWS_PER_QUERY);
        // A hierarchy table carries its rows; a contents table reads just this window from the database.
        const slice: string[] = table.contentsKind
            ? await resolveContentsWindow(context, table, cursor, count)
            : (table.rows ?? []).slice(cursor, cursor + count);

        // Shared across every row this call resolves - see ResolutionCache's own doc comment for why: without
        // it, a hierarchy table's `hasChildren` column (or a calendar table's organizer/attendee resolution)
        // would re-fetch the same whole-mailbox folder list/mailbox record once per row instead of once per
        // call.
        const cache: ResolutionCache = {};
        const rowBuffers: Buffer[] = [];
        // Only as many rows as fit in the room left in this request's ROP output buffer; the rest stay for the next
        // call, since the cursor only advances past rows actually returned.
        let room = (context.ropOutputRemaining ?? Infinity) - RESPONSE_HEADER_BYTES;
        for (const target of slice) {
            const row = await this.buildRow(context, target, table.columns ?? [], cache);
            if (row.length > room) {
                break;
            }
            room -= row.length;
            rowBuffers.push(row);
        }
        if (rowBuffers.length === 0 && slice.length > 0) {
            // Not even one row fits: [MS-OXCTABL] ecBufferTooSmall, and the cursor stays put.
            writer.writeUInt8(ROP_ID_QUERY_ROWS);
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt32LE(ERROR_BUFFER_TOO_SMALL);
            return;
        }
        table.cursor = cursor + rowBuffers.length;

        writer.writeUInt8(ROP_ID_QUERY_ROWS);
        writer.writeUInt8(inputHandleIndex);
        writer.writeUInt32LE(0); // ReturnValue - success
        writer.writeUInt8(ORIGIN_BOOKMARK_END);
        writer.writeUInt16LE(rowBuffers.length);
        for (const rowBuffer of rowBuffers) {
            writer.writeBytes(rowBuffer);
        }
    }

    private async buildRow(
        context: RopContext,
        target: string,
        columns: { propertyId: number; propertyType: number }[],
        cache: ResolutionCache,
    ): Promise<Buffer> {
        const writer = new BufferWriter();
        writer.writeUInt8(0x00); // Flags - StandardPropertyRow, see class doc comment
        const values = await resolvePropertyValues(target, columns, context, cache);
        columns.forEach((column, index) => writePropertyValueSafely(writer, column.propertyType, values[index]));
        return writer.toBuffer();
    }
}
