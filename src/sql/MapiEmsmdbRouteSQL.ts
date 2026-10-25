///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { AuditLogEntrySQL, CalendarEventSQL, ContactSQL, EncryptionPolicySQL, FolderSQL, LabelSQL, MailboxSQL, MessageSQL, TaskSQL, ZeroAccessIndexEntrySQL } from "@rapidmx/restapi/sql";
import { BaseMapiEmsmdbRoute } from "../BaseMapiEmsmdbRoute.js";
import { RopLogonHandler } from "../rop/RopLogonHandler.js";
import { RopReleaseHandler } from "../rop/RopReleaseHandler.js";
import { RopOpenFolderHandler } from "../rop/RopOpenFolderHandler.js";
import { RopGetHierarchyTableHandler } from "../rop/RopGetHierarchyTableHandler.js";
import { RopGetContentsTableHandler } from "../rop/RopGetContentsTableHandler.js";
import { RopSetColumnsHandler } from "../rop/RopSetColumnsHandler.js";
import { RopQueryRowsHandler } from "../rop/RopQueryRowsHandler.js";
import { RopOpenMessageHandler } from "../rop/RopOpenMessageHandler.js";
import { RopGetPropertiesSpecificHandler } from "../rop/RopGetPropertiesSpecificHandler.js";
import { RopOpenStreamHandler } from "../rop/RopOpenStreamHandler.js";
import { RopReadStreamHandler } from "../rop/RopReadStreamHandler.js";
import { RopCreateMessageHandler } from "../rop/RopCreateMessageHandler.js";
import { RopSetPropertiesHandler } from "../rop/RopSetPropertiesHandler.js";
import { RopWriteStreamHandler } from "../rop/RopWriteStreamHandler.js";
import { RopSaveChangesMessageHandler } from "../rop/RopSaveChangesMessageHandler.js";
import { RopSubmitMessageHandler } from "../rop/RopSubmitMessageHandler.js";
import { RopGetPropertyIdsFromNamesHandler } from "../rop/RopGetPropertyIdsFromNamesHandler.js";
import { RopDeleteMessagesHandler } from "../rop/RopDeleteMessagesHandler.js";
import { RopDeleteFolderHandler } from "../rop/RopDeleteFolderHandler.js";
import { RopFastTransferSourceCopyToHandler } from "../rop/RopFastTransferSourceCopyToHandler.js";
import { RopFastTransferSourceCopyPropertiesHandler } from "../rop/RopFastTransferSourceCopyPropertiesHandler.js";
import { RopFastTransferSourceGetBufferHandler } from "../rop/RopFastTransferSourceGetBufferHandler.js";
import { RouteDecorators } from "@rapidrest/service-core";
const { Route } = RouteDecorators;

/**
 * SQL-backed concrete `BaseMapiEmsmdbRoute`, mounted at `/mapi/emsmdb`. Exported from this plugin's
 * `./sql` entry point, so the server host mounts it without a wrapper class of its
 * own.
 *
 * @author Jean-Philippe Steinmetz
 */
@Route("/mapi/emsmdb")
export class MapiEmsmdbRouteSQL extends BaseMapiEmsmdbRoute<MailboxSQL> {
    protected mailboxClass: any = MailboxSQL;
    protected folderClass: any = FolderSQL;
    protected messageClass: any = MessageSQL;
    protected calendarEventClass: any = CalendarEventSQL;
    protected encryptionPolicyClass: any = EncryptionPolicySQL;
    protected zeroAccessIndexEntryClass: any = ZeroAccessIndexEntrySQL;
    protected contactClass: any = ContactSQL;
    protected taskClass: any = TaskSQL;
    protected labelClass: any = LabelSQL;
    protected auditLogClass: any = AuditLogEntrySQL;
    protected ropHandlerClasses: any[] = [
        RopLogonHandler,
        RopReleaseHandler,
        RopOpenFolderHandler,
        RopGetHierarchyTableHandler,
        RopGetContentsTableHandler,
        RopSetColumnsHandler,
        RopQueryRowsHandler,
        RopOpenMessageHandler,
        RopGetPropertiesSpecificHandler,
        RopOpenStreamHandler,
        RopReadStreamHandler,
        RopCreateMessageHandler,
        RopSetPropertiesHandler,
        RopWriteStreamHandler,
        RopSaveChangesMessageHandler,
        RopSubmitMessageHandler,
        RopGetPropertyIdsFromNamesHandler,
        RopDeleteMessagesHandler,
        RopDeleteFolderHandler,
        RopFastTransferSourceCopyToHandler,
        RopFastTransferSourceCopyPropertiesHandler,
        RopFastTransferSourceGetBufferHandler,
    ];
}
