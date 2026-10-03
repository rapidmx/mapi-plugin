# Release Notes

## Unreleased

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
