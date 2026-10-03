///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
// Builds a `MapiSessionManager` the way the framework does: with an `ObjectFactory` of its own (a fresh one per manager,
// since the factory returns the first instance built under a name) and its `@Init` hook run.
import { Logger } from "@rapidrest/core";
import { ObjectFactory } from "@rapidrest/service-core";
import { MapiSessionManager } from "../src/MapiSessionManager.js";
import config from "./config.js";

/** A manager on `client`'s Redis stores, or on the in-process stores when no client is given. */
export async function buildSessionManager(client?: any): Promise<MapiSessionManager> {
    const manager = new MapiSessionManager();
    (manager as any)._objectFactory = new ObjectFactory(config, Logger());
    if (client) {
        (manager as any).redisClient = client;
    }
    await manager.init();
    return manager;
}
