///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import type { BufferReader, BufferWriter } from "../codec/BufferCursor.js";
import { readPropertyTagArray } from "../codec/PropertyValue.js";
import { resolvePropertyValues, writePropertyValueSafely } from "./PropertyResolvers.js";
import type { RopContext, RopHandler } from "./RopHandler.js";

const ROP_ID_GET_PROPERTIES_SPECIFIC = 0x07;

/** The well-known MAPI HRESULT `MAPI_E_INVALID_OBJECT`, reused for "the referenced handle isn't a
 * folder/message (or doesn't exist)" - the same constant `RopGetHierarchyTableHandler`/`RopGetContentsTableHandler`
 * use for their own analogous checks. Attachment/Logon objects (also spec-valid targets for this ROP) aren't
 * supported in this pragmatic subset. */
const ERROR_INVALID_OBJECT = 0x80070005;

/** `MAPI_E_TOO_BIG`, for more than `MAX_PROPERTY_TAG_COUNT` requested tags. */
const ERROR_TOO_BIG = 0x80040305;

/**
 * `RopGetPropertiesSpecific` (`[MS-OXCPRPT]`/`[MS-OXCROPS]`): fetches named property values for a single
 * already-opened Folder or Message object (`RopOpenFolder`/`RopOpenMessage`), by an explicit `PropertyTags`
 * list - the single-object analog of `RopQueryRows`' per-row column fetch, reusing the exact same
 * `PropertyResolvers` value-resolution logic (a folder/message handle's `entityUid` is already the same
 * `"folder:"`/`"virtual:"`/`"message:"` target-string format a table row uses). Like `RopQueryRows`, this
 * pragmatic subset always emits a `StandardPropertyRow` (`Flags = 0x00`, confirmed as the response's `RowData`
 * format via `[MS-OXCROPS]`'s own "Success Response Buffer" page, which names `[MS-OXCDATA]` §2.8's
 * `PropertyRow` structure directly) - an unsupported property falls back to a type-appropriate default rather
 * than a `FlaggedPropertyRow`'s per-column error signaling, exactly as `RopQueryRows` already does. That includes
 * a client-requested `propertyType` that doesn't match what a tag's own `propertyId` actually resolves to
 * (`PropertyResolvers.writePropertyValueSafely`, shared with `RopQueryRowsHandler`) - not just a `propertyId`
 * this codebase has no data for.
 *
 * `PropertySizeLimit`/`WantUnicode` are decoded (to advance the reader correctly) but not honored - this
 * pragmatic subset never truncates a property value and always encodes strings the same way regardless of the
 * client's Unicode preference (see `PropertyValue.ts`'s own string-encoding doc comments).
 *
 * @author Jean-Philippe Steinmetz
 */
export class RopGetPropertiesSpecificHandler implements RopHandler {
    public readonly ropId = ROP_ID_GET_PROPERTIES_SPECIFIC;

    public async handle(reader: BufferReader, writer: BufferWriter, context: RopContext): Promise<void> {
        reader.readUInt8(); // LogonId - this pragmatic subset doesn't track multiple concurrent logons per session
        const inputHandleIndex: number = reader.readUInt8();
        reader.readUInt16LE(); // PropertySizeLimit - no truncation support in this pragmatic subset
        reader.readUInt16LE(); // WantUnicode - this pragmatic subset always encodes strings the same way regardless
        const propertyTags = readPropertyTagArray(reader, reader.readUInt16LE());

        const handle = context.session.handles[inputHandleIndex];
        if (!handle || (handle.type !== "folder" && handle.type !== "message") || !propertyTags) {
            writer.writeUInt8(ROP_ID_GET_PROPERTIES_SPECIFIC);
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt32LE(propertyTags ? ERROR_INVALID_OBJECT : ERROR_TOO_BIG);
            return;
        }

        const values = await resolvePropertyValues(handle.entityUid, propertyTags, context);

        writer.writeUInt8(ROP_ID_GET_PROPERTIES_SPECIFIC);
        writer.writeUInt8(inputHandleIndex);
        writer.writeUInt32LE(0); // ReturnValue - success
        writer.writeUInt8(0x00); // PropertyRow Flags - StandardPropertyRow, see class doc comment
        propertyTags.forEach((tag, index) => writePropertyValueSafely(writer, tag.propertyType, values[index]));
    }
}
