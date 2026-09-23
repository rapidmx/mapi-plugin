///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { sendOrThrow } from "../../src/rop/TransportSend.js";

describe("sendOrThrow Tests", () => {
    it("Throws when the transport reports no result at all.", async () => {
        const transport = { send: vi.fn().mockResolvedValue(undefined) } as any;

        await expect(sendOrThrow(transport, {} as any)).rejects.toThrow(/did not accept the message \(rejected: 0\)/);
    });

    it("Throws when nothing was accepted, even with no rejections either.", async () => {
        const transport = { send: vi.fn().mockResolvedValue({ accepted: [], rejected: [] }) } as any;

        await expect(sendOrThrow(transport, {} as any)).rejects.toThrow(/rejected: 0/);
    });

    it("Throws when at least one recipient was rejected, reporting the rejected count.", async () => {
        const transport = { send: vi.fn().mockResolvedValue({ accepted: ["a@example.com"], rejected: ["b@example.com", "c@example.com"] }) } as any;

        await expect(sendOrThrow(transport, {} as any)).rejects.toThrow(/rejected: 2/);
    });

    it("Resolves with the transport's result when at least one recipient was accepted and none were rejected.", async () => {
        const result = { accepted: ["a@example.com"], rejected: [] };
        const transport = { send: vi.fn().mockResolvedValue(result) } as any;

        await expect(sendOrThrow(transport, {} as any)).resolves.toBe(result);
    });

    it("Treats a result with neither field present as nothing accepted, defaulting both to an empty array.", async () => {
        const transport = { send: vi.fn().mockResolvedValue({}) } as any;

        await expect(sendOrThrow(transport, {} as any)).rejects.toThrow(/rejected: 0/);
    });

    it("Resolves when accepted is present but rejected is entirely absent (defaults to an empty array).", async () => {
        const result = { accepted: ["a@example.com"] };
        const transport = { send: vi.fn().mockResolvedValue(result) } as any;

        await expect(sendOrThrow(transport, {} as any)).resolves.toBe(result);
    });
});
