# Release Notes

## Unreleased

### Added

- **Zero-access encryption for calendars** (`@rapidmx/restapi` 0.35). When zero-access encryption protects a mailbox's calendar, appointments saved from Outlook have their title, location and description sealed, as web-client and ActiveSync saves are. Outlook can't open sealed fields, so it shows their placeholders. mapi doesn't save contacts or tasks, so the new contact and task protection doesn't apply here.
- **Outlook is recorded on the mailbox.** Each Outlook Connect sets `Mailbox.syncClients.mapi.lastSeenAt`, at most every six hours. When a user turns on zero-access, restapi uses this to decide which data types to protect by default and which to warn about. Outlook can't read protected calendars, contacts or tasks, so those types default to off for a mailbox that syncs with Outlook. Mail stays on once the user has exported their keys. If the write fails, it is logged and the Connect still succeeds.
- **End-to-end encrypted mail has an explanatory body.** Outlook used to show an empty or garbled body for a message its sender encrypted end to end (S/MIME, OpenPGP/MIME or inline OpenPGP, from Proton Mail and other OpenPGP apps). It now shows a short explanation: open the message in RapidMX, or in a mail app that has the user's encryption keys. This is the same approach as the zero-access placeholder.

### Changed

- **Requires `@rapidmx/restapi` 0.35 or later.** The plugin relies on restapi's new zero-access writer and calendar sealing. Upgrade restapi (the server) before or together with this plugin. restapi 0.35 also brings zero-access encryption for contacts and tasks, per-meeting keys for meetings with external attendees, and OpenPGP interoperability (Web Key Directory, Autocrypt, per-recipient choice of format).
- **Custom route subclasses must name two more models.** `BaseMapiEmsmdbRoute` takes the optional `encryptionPolicyClass` and `zeroAccessIndexEntryClass`. `MapiEmsmdbRouteMongo` and `MapiEmsmdbRouteSQL` already set both. A subclass of your own that leaves `encryptionPolicyClass` out seals nothing. One that leaves `zeroAccessIndexEntryClass` out stores Sent Items copies readable.
- **Refused edits to events from zero-access invitations.** An event filed from a zero-access invitation holds its title, location and description in one sealed copy, and Outlook only ever sees placeholders for them. Outlook can still move such an event or change its other details, and those saves keep the sealed copy as it is. A new title or location is refused with `MAPI_E_NO_ACCESS` and nothing is saved, because resealing the content would lose the location and description Outlook never received. Make those edits in RapidMX.
- **Deferred sends stay readable until relayed.** A message Outlook sends with "Do not deliver before" stays readable in the Outbox until it is relayed, the same as a scheduled send through REST. restapi 0.35 seals the Sent Items copy, and it keeps a deleted scheduled message readable for `mail:jobs:zero_access_backfill:deleted_scheduled_send_days` (default 30) so restoring it resumes the send.
- **MAPI logging is quieter.** The temporary `MAPI_DEBUG` warn-level logging is gone. Requests now produce debug-level lines with the request type, user, client info, request id, body size and ROP byte counts. A ROP handler that throws is still logged at warn, by ROP id, handler name and error message.

### Fixed

- **Sealed locations survive an Outlook re-save.** A protected appointment saved back with the empty location Outlook was shown no longer deletes its sealed location. Clearing the location of an event that has no sealed location still works.
- **Events from zero-access invitations keep their sealing.** Any save from Outlook, even a drag to a new time, used to seal the `[...]` placeholder over the real title and rename the event to "Protected event" in other clients. These events no longer pick up a second set of sealed fields.
- **A failed send leaves nothing behind.** If the Sent Items copy of an Outlook send fails to seal, find its folder or save, its body blob, sealed index entry and the relay scan's sanitized HTML are removed. Before, they were left with no message naming them, and a stray index entry could show up in the owner's search.

### Security

- **Outlook's Sent Items copies are sealed.** The Sent Items copy of a message Outlook sends from a zero-access mailbox used to be stored readable. It is now sealed with restapi's zero-access writer, with a placeholder subject and preview, and the readable HTML the relay scan produced is deleted. Messages the sender encrypted end to end are stored as they are.
- **Message content no longer reaches the logs.** The temporary diagnostic logging wrote whole `Execute` ROP buffers and responses, `Connect` and NSPI bodies, and the `MapiContext` session cookie to the log at warn level. That included the subject, location and body of everything Outlook saved or sent, even for mailboxes the server then seals, and address book names and addresses. Logs now contain only sizes, ids and error codes. Delete any old logs that captured `MAPI_DEBUG` lines.
- **Dependency audit.** `resolutions` now force fixed versions of follow-redirects, nodemailer, source-map-js, handlebars, brace-expansion, tar and undici, and `nodemailer` is raised to `^10.1.4`. `.yarnrc.yml` ignores only advisory 1240912 (GHSA-86w9-cpqp-85rv, node-forge through `acme-client`), which has no fixed release yet. Every other node-forge advisory still fails the audit, and the exception will be removed once a fix ships.

## v1.3.0

### Added

- **Zero-access encrypted mail has an explanatory body.** This server can't read a message stored with zero-access encryption (`@rapidmx/restapi` 0.34), so Outlook over MAPI shows a short explanation instead of its body: open it in RapidMX, or install the exported RapidMX keys.

## v1.2.0

## v1.1.1

## v1.1.0

## v1.0.0

Published by mistake as a release-tooling error (`release patch` on a `1.0.0-beta.18` prerelease drops the
prerelease tag rather than advancing it - see `.claude/NOTES.md`'s 2026-09-28 entry). The code is identical to
`1.0.0-beta.18` plus the `@rapidmx/restapi` dependency bump below; it does **not** signal that this package has
reached release stability - the Outlook/MAPI-HTTP connectivity investigation referenced throughout this file's
earlier betas is still open. Once published, npm does not allow deleting a version, so this repository's
versioning continues forward from `1.0.0` with ordinary `patch`/`minor`/`major` releases rather than resuming a
beta prerelease line.

### Changed

- Bump the `@rapidmx/restapi` development dependency to `^0.25.1`, now that it's published.

## v1.0.0-beta.18

### Fixed

- **`/mapi/emsmdb` and `/mapi/nspi` no longer send the framework's own CSRF cookie.** beta.17's cookie-attribute fix wasn't enough on its own - a follow-up capture showed real Outlook desktop still never sending a `Cookie` header back, even with attributes now matching Exchange Online exactly. The one remaining difference: every response from this deployment also carried an unrelated `csrf=...` cookie, issued unconditionally by the framework's CSRF middleware (which only ever matters for cookie-based browser auth, never this route's Bearer/Basic auth) and mixed in alongside the real session cookies - something a real Exchange Online response never has. Cleared before this route's own cookies are set, entirely within this plugin (no framework or server change needed).

## v1.0.0-beta.17

### Fixed

- **Session cookies (`MapiContext`/`MapiSequence` on `/mapi/emsmdb`, `NspiContext` on `/mapi/nspi`) now carry `Path`/`Secure`/`SameSite=None`/`HttpOnly`, instead of being bare `name=value` pairs.** Real Outlook desktop never sent a `Cookie` header back to this deployment on any request, across every capture taken during this investigation - confirmed directly by comparing a live Fiddler capture of the same Outlook client's traffic against a working Exchange Online mailbox in the same session: Exchange Online's `Set-Cookie` headers are fully decorated this way, and Outlook reliably echoes them back (hundreds of real `Execute` calls observed); this deployment's bare cookies were never once echoed back in three separate full-session captures. This is very likely the actual cause of Outlook desktop never completing a mailbox logon against this deployment.

## v1.0.0-beta.16

### Fixed

- **`/mapi/nspi` (`Bind`/`Unbind`/`GetMatches`) now returns the mandatory `X-RequestId`/`X-ClientInfo` response headers.** `[MS-OXCMAPIHTTP]`'s common response format requires every response, of every request type, to echo back the caller's own `X-RequestId` unchanged - this route never set it at all, verified against both the spec text and Gromox's own reference implementation (which builds every MAPI/HTTP response, `Bind` included, from one shared header-setting function for exactly this reason). `/mapi/emsmdb` already did this correctly; this brings `/mapi/nspi` in line with it.

## v1.0.0-beta.15

### Added

- **Extend the temporary `MAPI_DEBUG` diagnostic logging to the raw wire bytes of `Connect` and `Bind`.** beta.14's capture showed both operations succeeding cleanly on every attempt, on a repeating cycle, with `Execute`/`GetMatches`/anything else never once being called - so the next thing to check is whether the response bytes themselves are actually spec-correct, not just "our own code thinks it succeeded." Logs `Connect`'s incoming/outgoing body hex (plus the `Set-Cookie` values sent) and every NSPI request type's incoming/outgoing body hex (via a temporary `res.send()` wrapper, so it covers `Bind` today and any other NSPI operation for free).

## v1.0.0-beta.14

### Added

- **Extend the temporary `MAPI_DEBUG` diagnostic logging to `/mapi/nspi`** (`BaseMapiNspiRoute`): the beta.13 capture showed Outlook desktop repeatedly `Connect`ing to the mailbox store but never once calling `Execute` - the Address Book (NSPI) provider, previously unlogged entirely, is the likely place the client is actually getting stuck. Logs dispatch entry, `Bind`/`GetMatches` success, and the real error on any throw.

## v1.0.0-beta.13

### Added

- **Temporary diagnostic logging** for the Outlook desktop "set of folders cannot be opened" investigation. `BaseMapiEmsmdbRoute` now logs (`warn` level, tag `MAPI_DEBUG`) each request's type, `Connect` success, `Execute`'s incoming/outgoing `RopsList` hex and every `Execute`-level failure path (session not found, invalid sequence, buffer too small); `RopDispatcher` now logs (`console.error`, tag `MAPI_DEBUG`) the real error whenever an individual ROP handler throws - previously swallowed entirely into a bare `MAPI_E_CALL_FAILED`/`MAPI_E_TOO_COMPLEX` with no server-side trace of why. Not a behavior change - purely observational - and expected to be removed again once the real root cause is confirmed from what this reveals.

## v1.0.0-beta.12

## v1.0.0-beta.11

### Changed

- The development dependency on `@rapidmx/restapi` is 0.23.0 (was 0.22.1); the plugin is otherwise unchanged. A minor bump (a new beta) because restapi's release is a minor.

## v1.0.0-beta.10

## v1.0.0-beta.9

## v1.0.0-beta.8

## v1.0.0-beta.7

### Changed

- Bump the `@rapidmx/restapi` development dependency to `^0.21.1` (peer range unchanged) and rebuild against it.

## v1.0.0-beta.6
