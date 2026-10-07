///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import type { RepoUtils } from "@rapidrest/service-core";
import { FolderType, type Folder } from "@rapidmx/restapi";
import type { MapiObjectHandle } from "../MapiSessionManager.js";
import { CALENDAR_EVENT_SORT } from "./CalendarEventTarget.js";
import { CONTACT_SORT } from "./ContactTarget.js";
import { MESSAGE_SORT } from "./MessageTarget.js";
import { findWindow, type RepoSort } from "./RepoPaging.js";
import type { RopContext } from "./RopHandler.js";
import { TASK_SORT } from "./TaskTarget.js";

export type ContentsKind = NonNullable<MapiObjectHandle["contentsKind"]>;

const SORT_BY_KIND: Record<ContentsKind, RepoSort> = {
    message: MESSAGE_SORT,
    calendarEvent: CALENDAR_EVENT_SORT,
    contact: CONTACT_SORT,
    task: TASK_SORT,
};

/** Which repo a folder's items live in, decided by the folder's own `type`: a `Contact`/`Task`/`CalendarEvent` is
 * its own entity, never a `Message` row. */
export async function resolveContentsKind(folderUid: string, context: RopContext): Promise<ContentsKind> {
    context.budget?.chargeQueries();
    const folder: Folder | undefined = await context.folderRepo.findOne(folderUid, { ignoreACL: true });
    switch (folder?.type) {
        case FolderType.CALENDAR:
            return "calendarEvent";
        case FolderType.CONTACTS:
        case FolderType.SUGGESTED_CONTACTS:
            return "contact";
        case FolderType.TASKS:
            return "task";
        default:
            return "message";
    }
}

function repoFor(kind: ContentsKind, context: RopContext): RepoUtils<any> | undefined {
    switch (kind) {
        case "calendarEvent":
            return context.calendarEventRepo;
        case "contact":
            return context.contactRepo;
        case "task":
            return context.taskRepo;
        default:
            return context.messageRepo;
    }
}

/**
 * The `"<kind>:<uid>"` row targets at positions `[start, start + count)` of a contents table (a `"folder:<uid>"`
 * table with `contentsKind` set), read straight from the database in a fixed, stable order. A context without the
 * matching optional repo yields no rows.
 *
 * Reading a window per `RopQueryRows` instead of snapshotting every row at `RopGetContentsTable` time keeps session
 * state small and makes no folder too large to page through. The trade-off is that a message added or removed
 * between two `RopQueryRows` calls can shift the rows that follow.
 */
export async function resolveContentsWindow(context: RopContext, table: MapiObjectHandle, start: number, count: number): Promise<string[]> {
    const kind: ContentsKind = table.contentsKind!;
    const repo = repoFor(kind, context);
    if (!repo) {
        return [];
    }
    const folderUid = table.entityUid.slice("folder:".length);
    const rows: { uid: string }[] = await findWindow(repo, { folderUid, mailboxUid: context.mailboxUid }, SORT_BY_KIND[kind], start, count, context.budget);
    return rows.map((row) => `${kind}:${row.uid}`);
}
