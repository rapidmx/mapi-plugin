///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
// MapiTimeZone is pure binary-format logic with no DI/DB dependency, tested directly here - same precedent as
// test/mapi/codec/AppointmentRecurrence.test.ts.
import { BufferReader } from "../../src/codec/BufferCursor.js";
import { decodeTimeZoneStruct, encodeTimeZoneStruct } from "../../src/codec/MapiTimeZone.js";

/** A 48-byte TimeZoneStruct with the given lBias and no DST. */
function structWithBias(bias: number): Buffer {
    const buffer = Buffer.alloc(48);
    buffer.writeInt32LE(bias, 0);
    return buffer;
}

describe("MapiTimeZone Tests", () => {
    describe("UTC", () => {
        it("Encodes UTC as a 48-byte, all-zero-except-nothing struct.", () => {
            const encoded = encodeTimeZoneStruct("UTC", new Date("2026-01-15T00:00:00.000Z"));
            expect(encoded.length).toBe(48);
            expect(encoded.every((byte) => byte === 0)).toBe(true);
        });

        it("Round-trips UTC back to the string 'UTC'.", () => {
            const encoded = encodeTimeZoneStruct("UTC", new Date("2026-01-15T00:00:00.000Z"));
            expect(decodeTimeZoneStruct(new BufferReader(encoded))).toBe("UTC");
        });
    });

    describe("Fixed-offset IANA zones", () => {
        it("Encodes America/Los_Angeles in January (PST, UTC-8) with lBias=480.", () => {
            const encoded = encodeTimeZoneStruct("America/Los_Angeles", new Date("2026-01-15T00:00:00.000Z"));
            expect(encoded.readInt32LE(0)).toBe(480);
            expect(encoded.length).toBe(48);
        });

        it("Encodes America/Los_Angeles in July (PDT, UTC-7) with lBias=420, reflecting the reference instant's own offset.", () => {
            const encoded = encodeTimeZoneStruct("America/Los_Angeles", new Date("2026-07-15T00:00:00.000Z"));
            expect(encoded.readInt32LE(0)).toBe(420);
        });

        it("Encodes Asia/Shanghai (UTC+8) with a negative lBias.", () => {
            const encoded = encodeTimeZoneStruct("Asia/Shanghai", new Date("2026-01-15T00:00:00.000Z"));
            expect(encoded.readInt32LE(0)).toBe(-480);
        });

        it("Writes zero lStandardBias/lDaylightBias and all-zero transition blocks (no DST modeled).", () => {
            const encoded = encodeTimeZoneStruct("America/Los_Angeles", new Date("2026-01-15T00:00:00.000Z"));
            expect(encoded.readInt32LE(4)).toBe(0); // lStandardBias
            expect(encoded.readInt32LE(8)).toBe(0); // lDaylightBias
            expect(encoded.subarray(12).every((byte) => byte === 0)).toBe(true); // both transition blocks
        });

        it("Decodes an America/Los_Angeles-derived struct back to the Etc/GMT+8 approximation.", () => {
            const encoded = encodeTimeZoneStruct("America/Los_Angeles", new Date("2026-01-15T00:00:00.000Z"));
            expect(decodeTimeZoneStruct(new BufferReader(encoded))).toBe("Etc/GMT+8");
        });

        it("Decodes an Asia/Shanghai-derived struct back to the Etc/GMT-8 approximation.", () => {
            const encoded = encodeTimeZoneStruct("Asia/Shanghai", new Date("2026-01-15T00:00:00.000Z"));
            expect(decodeTimeZoneStruct(new BufferReader(encoded))).toBe("Etc/GMT-8");
        });

        it("Keeps a half-hour offset (Asia/Kolkata, UTC+5:30) on decode instead of rounding it to a whole hour.", () => {
            const encoded = encodeTimeZoneStruct("Asia/Kolkata", new Date("2026-01-15T00:00:00.000Z"));
            expect(encoded.readInt32LE(0)).toBe(-330);
            expect(decodeTimeZoneStruct(new BufferReader(encoded))).toBe("Asia/Kolkata");
        });

        it("Keeps a quarter-hour offset (UTC+5:45) as Asia/Kathmandu.", () => {
            expect(decodeTimeZoneStruct(new BufferReader(structWithBias(-345)))).toBe("Asia/Kathmandu");
        });

        it("Decodes a minute offset with no DST-less zone to a fixed-offset identifier Intl accepts.", () => {
            const zone = decodeTimeZoneStruct(new BufferReader(structWithBias(210))); // UTC-3:30
            expect(zone).toBe("-03:30");
            // Round-trips through encode without losing the minutes.
            expect(encodeTimeZoneStruct(zone, new Date("2026-01-15T00:00:00.000Z")).readInt32LE(0)).toBe(210);
            expect(decodeTimeZoneStruct(new BufferReader(structWithBias(-765)))).toBe("+12:45");
        });

        it("Decodes the extreme whole-hour offsets UTC-12 and UTC+14 to their Etc/GMT zones.", () => {
            expect(decodeTimeZoneStruct(new BufferReader(structWithBias(720)))).toBe("Etc/GMT+12");
            expect(decodeTimeZoneStruct(new BufferReader(structWithBias(-840)))).toBe("Etc/GMT-14");
        });

        it("Decodes an out-of-range bias (beyond UTC-12/UTC+14) as UTC instead of an invalid zone name.", () => {
            expect(decodeTimeZoneStruct(new BufferReader(structWithBias(721)))).toBe("UTC");
            expect(decodeTimeZoneStruct(new BufferReader(structWithBias(-841)))).toBe("UTC");
            expect(decodeTimeZoneStruct(new BufferReader(structWithBias(-600000)))).toBe("UTC");
        });

        it("Encodes a zero-offset zone that isn't literally named 'UTC' (Africa/Abidjan) with lBias=0, exercising the bare-'GMT' offset-string branch.", () => {
            const encoded = encodeTimeZoneStruct("Africa/Abidjan", new Date("2026-01-15T00:00:00.000Z"));
            expect(encoded.readInt32LE(0)).toBe(0);
        });
    });

    describe("Intl output shapes", () => {
        afterEach(() => {
            vi.restoreAllMocks();
        });

        it("Treats a bare 'GMT' offset string as a zero offset regardless of the runtime's own ICU data.", () => {
            // Whether a real zero-offset zone formats as "GMT" or "GMT+0" varies across ICU/CLDR versions, so pin
            // the bare-"GMT" shape with a stub rather than relying on the host runtime to produce it.
            vi.spyOn(Intl, "DateTimeFormat").mockImplementation(function () {
                return {
                    formatToParts: () => [{ type: "timeZoneName", value: "GMT" }],
                } as unknown as Intl.DateTimeFormat;
            });

            const encoded = encodeTimeZoneStruct("Europe/London", new Date("2026-01-15T00:00:00.000Z"));
            expect(encoded.readInt32LE(0)).toBe(0);
        });
    });

    describe("Error handling", () => {
        it("Encodes an unrecognized zone identifier as UTC instead of throwing.", () => {
            const encoded = encodeTimeZoneStruct("Not/AZone", new Date("2026-01-15T00:00:00.000Z"));
            expect(encoded.length).toBe(48);
            expect(encoded.readInt32LE(0)).toBe(0);
        });
    });
});
