///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { RouteDecorators } from "@rapidrest/service-core";
import { MapiEmsmdbRouteMongo } from "../../../src/mongo/MapiEmsmdbRouteMongo.js";
const { Route } = RouteDecorators;

/**
 * A second, otherwise-identical mount of `MapiEmsmdbRouteMongo` with no `auditLogClass` configured - exercises
 * `BaseMapiEmsmdbRoute.executeLocked()`'s `this.auditLogClass ? ... : undefined` ternary's false side (real
 * deployments always configure one; this route exists purely so that branch has a real-HTTP test instead of only
 * ever running the true side). See `test/routes/mongo/MapiEmsmdbRoute.test.ts`'s "no audit log class" test.
 */
@Route("/mongo/mapi/emsmdb-noaudit")
export class MapiEmsmdbRouteNoAudit extends MapiEmsmdbRouteMongo {
    protected auditLogClass: any = undefined;
}
