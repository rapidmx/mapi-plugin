///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { describe, expect, it } from "vitest";
import { MapiEmsmdbRouteMongo } from "../src/mongo/MapiEmsmdbRouteMongo.js";

describe("BaseMapiEmsmdbRoute init", () => {
    it("Refuses to build its repos and audit service when no ObjectFactory was injected.", async () => {
        const route = new MapiEmsmdbRouteMongo();
        await expect(route.init()).rejects.toThrow("objectFactory is not set.");
    });
});
