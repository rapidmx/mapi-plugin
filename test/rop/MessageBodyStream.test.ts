///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { ENCRYPTED_PLACEHOLDER_TEXT, PID_TAG_BODY, resolveMessageBodyBytes, ZERO_ACCESS_PLACEHOLDER_TEXT } from "../../src/rop/MessageBodyStream.js";
import { InMemoryBlobStore } from "../testDoubles.js";

describe("MessageBodyStream Tests", () => {
    it("Exposes PidTagBody as 0x1000.", () => {
        expect(PID_TAG_BODY).toBe(0x1000);
    });

    it("Resolves a real message's plain-text body as null-terminated UTF-16LE bytes.", async () => {
        const blobStore = new InMemoryBlobStore();
        await blobStore.put(
            "bodies/m1",
            Buffer.from("From: sender@example.com\r\nTo: owner@example.com\r\nSubject: Hi\r\n\r\nPlain body text."),
        );
        const messageRepo = { findOne: vi.fn().mockResolvedValue({ uid: "m1", bodyBlobKey: "bodies/m1" }) };

        const bytes = await resolveMessageBodyBytes("message:m1", messageRepo as any, blobStore);

        const expected = Buffer.concat([Buffer.from("Plain body text.", "utf16le"), Buffer.from([0x00, 0x00])]);
        expect(bytes).toEqual(expected);
        expect(messageRepo.findOne).toHaveBeenCalledWith("m1", { ignoreACL: true });
    });

    it("Degrades to an empty buffer for a message target whose real Message has since vanished.", async () => {
        const messageRepo = { findOne: vi.fn().mockResolvedValue(undefined) };
        const bytes = await resolveMessageBodyBytes("message:gone", messageRepo as any, {} as any);
        expect(bytes.length).toBe(0);
    });

    it("Falls back to an empty body when the raw MIME has no body content at all for mailparser to extract.", async () => {
        const blobStore = new InMemoryBlobStore();
        await blobStore.put("bodies/m2", Buffer.from("Subject: No Body\r\n\r\n"));
        const messageRepo = { findOne: vi.fn().mockResolvedValue({ uid: "m2", bodyBlobKey: "bodies/m2" }) };

        const bytes = await resolveMessageBodyBytes("message:m2", messageRepo as any, blobStore);

        const expected = Buffer.concat([Buffer.from("", "utf16le"), Buffer.from([0x00, 0x00])]);
        expect(bytes).toEqual(expected);
    });

    it("Serves the explanation for a zero-access message, never reading its sealed body.", async () => {
        const blobStore = { get: vi.fn() };
        const messageRepo = { findOne: vi.fn().mockResolvedValue({ uid: "m3", bodyBlobKey: "za/mbx/m3", zeroAccess: { keyFingerprint: "fp", sealedSummary: "x" } }) };

        const bytes = await resolveMessageBodyBytes("message:m3", messageRepo as any, blobStore as any);

        expect(bytes).toEqual(Buffer.concat([Buffer.from(ZERO_ACCESS_PLACEHOLDER_TEXT, "utf16le"), Buffer.from([0x00, 0x00])]));
        expect(blobStore.get).not.toHaveBeenCalled();
    });

    it("Serves the explanation for a message its sender encrypted end to end, never parsing the ciphertext.", async () => {
        const blobStore = { get: vi.fn() };
        const messageRepo = { findOne: vi.fn().mockResolvedValue({ uid: "m4", bodyBlobKey: "bodies/m4", encrypted: true }) };

        const bytes = await resolveMessageBodyBytes("message:m4", messageRepo as any, blobStore as any);

        expect(bytes).toEqual(Buffer.concat([Buffer.from(ENCRYPTED_PLACEHOLDER_TEXT, "utf16le"), Buffer.from([0x00, 0x00])]));
        expect(blobStore.get).not.toHaveBeenCalled();
    });
});
