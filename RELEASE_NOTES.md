# Release Notes

## Unreleased

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
