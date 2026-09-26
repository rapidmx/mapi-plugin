# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [1.0.0-beta.11] - 2026-09-26

### Changed
- Use @rapidmx/restapi 0.23.0 as the development dependency
- Note the dependency bump in the release notes

## [1.0.0-beta.10] - 2026-09-25

### Fixed
- Fixed code coverage

## [1.0.0-beta.9] - 2026-09-25

### Fixed
- Fixed repository URL (again)

## [1.0.0-beta.8] - 2026-09-25

### Changed
- Document that a downstream package's release bump level follows its upstream dependency's, minor for minor, patch for patch and major for major, in NOTES
- Upgraded rapidmx deps

### Fixed
- Fixed repository URL

## [1.0.0-beta.7] - 2026-09-25

### Changed
- Bump the @rapidmx/restapi development dependency to 0.21.1 and refresh the lockfile, leaving the peer range unchanged
- Document the bump in the release notes

## [1.0.0-beta.6] - 2026-09-24

### Added
- Added a RopDeleteMessagesHandler test proving the query budget stops the loop the same way it already stops RopDeleteFolder
- Added a route test proving Execute answers a generic 500 instead of leaking a raw error when building a failed ROP's own response throws something other than a DecodeError
- Added a route test exercising Execute with no auditLogClass configured
- Added a dedicated test file for TransportSend.sendOrThrow covering its missing-result, empty-accepted, non-empty-rejected and absent-field default branches
- Added tests for RopOpenMessageHandler's task-repo-absent branch and RopWriteStreamHandler's never-written write-stream branch
- Added a RopSubmitMessageHandler test proving the query budget stops resolving a recipient list partway through, mirroring RopDeleteMessagesHandler's own budget-exhaustion test
- Added tests proving a tight query budget now accounts for the audit write, not just the read/delete calls beside it, in both RopDeleteMessagesHandler and RopDeleteFolderHandler
- Added PropertyResolvers.writePropertyValueSafely, encoding into a scratch buffer first so a partial failed write can never leak stray bytes into the real response ahead of later columns/rows
- Added tests reproducing the PidTagSubject-requested-as-PtypGuid repro for both handlers and a direct unit-test block for writePropertyValueSafely

### Changed
- Charge RopDeleteMessagesHandler's per-message lookups and deletes against the Execute query budget, so a large client-supplied messageIdCount can't run unbounded sequential queries in one request
- Charge AddressList's per-recipient bare-display-name contact lookup against the Execute query budget, so a large To/Cc/Bcc list across many RopSubmitMessage calls can't run unmetered DB queries
- Charge the audit-log write inside RopDeleteMessagesHandler/RopDeleteFolderHandler's delete loops against the Execute query budget, centralized in auditMessageDelete() itself so it only charges when context.audit is actually configured
- Fall back to a type-appropriate default value instead of failing the whole RopQueryRows/RopGetPropertiesSpecific response when a column's client-requested PropertyType doesn't match what its propertyId actually resolves to
- Upgraded rapidrest and rapidmx deps

## [1.0.0-beta.5] - 2026-09-22

### Changed
- Bump the @rapidmx/restapi devDependency to ^0.17.0, since the folder count helpers this fix calls were added well after the ^0.10.0 this plugin was still pinned to in development
- Test every refresh site, including a real Mongo round trip proving a deleted message's folder ends with the right counts
- Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>

### Fixed
- Fixed Outlook's own folder pane showing stale unread and total counts after deleting mail, sending mail, or deleting a folder over MAPI, by refreshing and publishing restapi's derived folder counts after every write that changes a folder's messages - message delete, a folder deleted mid-way through its own item budget, and both the Outbox and Sent Items copies a submit creates

## [1.0.0-beta.4] - 2026-09-15

### Changed
- Update the README to the @rapidmx/mapi-plugin package name
- Co-Authored-By: Claude Opus 5 (1M context) <noreply@anthropic.com>
- Only use an EMSMDB session for Execute and Disconnect when it belongs to the signed-in user
- Bound NSPI GetMatches search patterns to the regex length limit and limit queries to the requested row count
- Co-Authored-By: Claude Opus 5 (1M context) <noreply@anthropic.com>
- Reject oversized or truncated ROP buffers with 400, cap handle tables at 255 and property tag arrays at 256, and limit ROPs per Execute
- Load MAPI sessions from Redis on every request with compare-and-set saves, cap sessions per user and session lifetime, and re-check mailbox ownership each Execute
- Keep MAPI hard deletes recoverable and audited since legal holds can't be checked from the plugin, walk folder deletes iteratively within the caller's mailbox with depth and size limits
- Page folder and contents table queries explicitly instead of the silent 100-row default, and cache parsed message bodies and FastTransfer streams outside the session
- Cap write streams, FastTransfer data and named properties, bump meeting sequence on meaningful changes, and only let organizers send invites
- Update only the caller's own copy on meeting responses and send an iTIP reply, and store the relayed Message-ID and raw message in Sent Items
- Keep minute precision in time zone decoding with a UTC fallback, reject CR/LF in recipient addresses, and release streams with their handles
- Co-Authored-By: Claude Opus 5 (1M context) <noreply@anthropic.com>
- Keep FastTransfer streams in a shared store and fail the ROP instead of paging a rebuilt, different stream
- Enforce MaxRopOut and the 16-bit RopSize limit, trimming ReadStream/GetBuffer/QueryRows output and answering RopBufferTooSmall
- Use random handle generations, cap session size and tracked MIDs, and store write-stream chunks outside the session
- Decline a single occurrence as a recurrence exception with a RECURRENCE-ID reply instead of deleting the series
- Isolate a failing ROP to its own error response and decode recurrence blobs with exceptions
- Lock each session during Execute and replay the stored response for a repeated X-RequestId
- Cap per-Execute rows and built bytes, decode native Outlook GlobalObjectIds, and return NOT_FOUND on no match
- Parse resolved display names in recipient lists, return X-ResponseCode 10 for a missing session, and enforce the per-user session cap atomically
- Fail meeting replies the transport rejects
- Co-Authored-By: Claude Opus 5 (1M context) <noreply@anthropic.com>
- Mark mail and meeting-response drafts submitted after any submit attempt and refuse further submits or edits, capping submits per Execute and recipients per message
- Charge folder walks, hierarchy fetches, contents windows and deletes against new per-Execute query and row budgets, and charge message bodies before parsing
- Delete FastTransfer and write-stream data from Redis on release and session end, cap stored bytes per session and user, and dedupe CopyProperties tags
- Look meetings up by bounded, exact-matched iCalendar UID, version-check updates of plain rows, and leave invites to restapi's MeetingSchedulingJob
- Answer ecBufferTooSmall when RopBufferTooSmall itself can't fit, and add declined override occurrences to the series exceptions
- Check session ownership before locking, renew the lock while running, and answer busy for a request id already in progress
- Stop ReadStream reporting a false end of stream, match recipient display names literally, evict the least recently used session, ignore unparseable dates and reminders, and pick the oldest well-known folder at logon
- Co-Authored-By: Claude Opus 5 (1M context) <noreply@anthropic.com>
- Updated @rapidrest/service-core to ^2.1.0 as both the dev dependency and the peer range
- Use ModelUtils.literal() for client-supplied contact display names and meeting iCalendar UIDs instead of eq() query strings
- Co-Authored-By: Claude Opus 5 (1M context) <noreply@anthropic.com>
- Charge write-stream chunks and owner byte quotas in constant time with running totals, a chunk cap per stream and rate-limited recounts, so repeated tiny writes can't stall Redis
- Reserve room for a RopBufferTooSmall response before each ROP so a response that doesn't fit is answered in place instead of discarding already-saved work
- Stamp inviteSequenceSent on appointment saves and request exactly one invite per submitted revision, so meetings that are only saved don't mail attendees
- Renew the in-progress marker only while it is still this request's, run renewals sequentially and stop them when the lock is lost
- Co-Authored-By: Claude Opus 5 (1M context) <noreply@anthropic.com>
- Upgraded restapi dep

## [1.0.0-beta.3] - 2026-09-14

### Added
- Added a test that each entry point exports only the mounted MAPI routes and that the manifest is present

### Changed
- Convert this library into a RapidMX server plugin: package.json carries a rapidmx.plugin manifest, and the ./mongo and ./sql entry points export MapiEmsmdbRoute and MapiNspiRoute mounted at /mapi/emsmdb and /mapi/nspi, so a server needs no wrapper classes
- Co-Authored-By: Claude Opus 5 (1M context) <noreply@anthropic.com>
- Upgraded deps
- Changing package name to @rapidmx/mapi-plugin

## [1.0.0-beta.2] - 2026-09-13

### Changed
- Switch NSPI GAL search from like() glob-wrapping to the new regex() operator for a genuine literal-substring match
- Change NspiGetMatchesHandler.findMatchingContacts to build a regex() query pattern via StringUtils.escapeRegExp(searchTerm) instead of wrapping the term in *...* glob wildcards for like(), closing the residual literal-*/?-acts-as-wildcard gap glob syntax can't escape
- Update NspiGetMatchesHandler.test.ts/BaseMapiNspiRoute.test.ts for the simplified signature and the regex() query assertions
- Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>

### Removed
- Removed the now-unneeded per-backend likePattern() hook from BaseMapiNspiRoute/MapiNspiRouteMongo/MapiNspiRouteSQL, since regex() compiles identically on both backends

### Changed
- Changed NSPI GAL search (GetMatches) from `like()` (glob syntax, requiring a `*...*`-wrapped pattern and still treating a literal `*`/`?` in the search term as a wildcard) to `regex()` (`@rapidrest/service-core` ^2.0's new operator), escaping the search term with `StringUtils.escapeRegExp` for a genuine literal-substring, case-insensitive match with no residual wildcard ambiguity - also removes the now-unneeded per-backend `likePattern()` hook from `BaseMapiNspiRoute`/`MapiNspiRouteMongo`/`MapiNspiRouteSQL`, since `regex()` compiles identically on both backends
- Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>

## [1.0.0-beta.1] - 2026-09-13

### Added
- Added read-only Categories support for messages (PidNameKeywords, PS_PUBLIC_STRINGS), resolving each Message.labelUids entry to its Label.name via a new labelRepo/labelClass RopContext field and PropertyResolvers.ts's existing per-call ResolutionCache
- Added end-to-end HTTP+DB integration tests for Categories in both the mongo and sql MapiEmsmdbRoute test suites

### Changed
- Confirm via a targeted test that FolderType.ARCHIVE needs no new mapi code
- Update README/CHANGELOG to document the new Categories support and the dependency bump
- Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>
- Update release notes

### Fixed
- Fixed NSPI GAL search regression from service-core 2.0's glob-syntax like() operator, add read-only Outlook Categories (PidNameKeywords) support, and bump @rapidmx/restapi to ^0.8.0/@rapidrest/service-core to 2.x
- Fixed NspiGetMatchesHandler/MapiNspiRouteMongo/MapiNspiRouteSQL's likePattern() wrapping search terms in raw *...* glob wildcards instead of regex-escaping them, since like()'s new glob grammar treats an unescaped term as an exact-match pattern with no wildcards

### Removed
- Removed @rapidrest/cli as a dep

### Added
- Added read-only Outlook Categories (PidNameKeywords, PS_PUBLIC_STRINGS) support for messages, resolving each Message.labelUids entry to its Label.name and returning them as PtypMultipleString via RopQueryRows/RopGetPropertiesSpecific
- Added labelRepo/labelClass as optional RopContext fields (mirroring contactRepo/taskRepo), populated by BaseMapiEmsmdbRoute
- Added end-to-end HTTP+DB integration tests for Categories in both the mongo and sql MapiEmsmdbRoute test suites
- Confirmed (via a targeted test against FolderType.ARCHIVE) that restapi's Archive folder type needs no new mapi code - generic folder/message browsing already handles any non-special-cased folder type identically

### Changed
- Bumped @rapidmx/restapi peer/dev dependency range to ^0.8.0 and @rapidrest/service-core to 2.x/^2.0.0
- Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>

### Fixed
- Fixed NSPI GAL search (GetMatches) returning zero matches against restapi ^0.8.0/@rapidrest/service-core 2.x, whose like() operator changed from raw substring matching to glob syntax (* / ?) that treats an unescaped search term as an exact-match pattern - likePattern() implementations now wrap the raw term in `*...*` instead of regex-escaping it
- Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>

## [1.0.0-beta.0] - 2026-09-09

### Added
- Added missing project files
- Added vite dev dependency
- Added read-only Contacts/Tasks folder browsing via new ContactTarget.ts/TaskTarget.ts, mirroring the existing CalendarEventTarget.ts pattern, wired through RopGetContentsTable/RopOpenMessage/RopGetPropertiesSpecific/RopQueryRows
- Added PidLidTaskStatus/PercentComplete/TaskDueDate/TaskComplete (PSETID_Task) named-property support for Task items via new TaskNamedProperties.ts
- Added PidTagGivenName/Surname/EmailAddress/BusinessTelephoneNumber/CompanyName/Title property support for Contact items
- Added PidTagDeferredSendTime ("do not deliver before") support to RopSubmitMessage, parking a deferred draft in Outbox with Message.scheduledSendTime for ScheduledSendJob to relay later
- Added PidTagReadReceiptRequested support to RopSubmitMessage, attaching a real Disposition-Notification-To header and recording Message.requestReceipt on the persisted Sent Items copy
- Added contactRepo/taskRepo/contactClass/taskClass as optional RopContext fields, populated by BaseMapiEmsmdbRoute, so existing RopContext test fixtures don't need updating
- Added end-to-end HTTP+DB integration tests for Contacts/Tasks folder browsing in both the mongo and sql MapiEmsmdbRoute test suites

### Changed
- Initial commit
- Update README's documented-gaps section to reflect the new Contacts/Tasks/deferred-send/read-receipt coverage and the deliberately-deferred Focused Inbox exposure
- Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>
- Change FolderTarget.assignOrGetFid/MessageTarget.assignOrGetMid/NamedPropertyRegistry.assignOrGetNamedPropertyId from an O(n) linear scan plus a Math.max(...spread) id-allocation to an O(1) reverse-index/counter pair
- Change assignOrGetNamedPropertyId to return 0x0000 once a session's entire named-property ID space is exhausted instead of assigning an out-of-range id that later crashed a 16-bit write
- Change RopQueryRowsHandler/RopGetPropertiesSpecificHandler to share one request-scoped ResolutionCache across every row/column resolved in a call, so per-row folder-list/mailbox fetches that only ever have one real answer per call happen at most once
- Change RopLogonHandler to resolve its four real-folder-type lookups concurrently instead of as four sequential round trips
- Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>

### Fixed
- Fixed @rapidmx/restapi peer dependency range (0.1.x) rejecting the ^0.3.0 releases this package is actually built and tested against
- Fixed RopSubmitMessageHandler's compose/send tests against a stale ScanPipeline mock missing the references/inReplyTo fields restapi ^0.3.0's scanAndRelay() now unconditionally reads to derive a conversation ID
- Fixed an unbounded PtypMultipleString/PtypMultipleString8 element count in RopSetProperties/NSPI GetMatches letting any authenticated caller hang or OOM the whole process with one crafted request
- Fixed BufferReader.readNullTerminatedUtf16LE/readNullTerminatedString8 silently returning truncated content past the buffer's end instead of throwing, closing the above DoS at its root
- Fixed BufferReader.readBytes() silently clamping an out-of-range or negative length instead of throwing, which let a malformed RopSize < 2 mis-parse decodeRopBuffer's handle table
- Fixed RopQueryRows/RopReadStream/RopFastTransferSourceGetBuffer throwing on any response exceeding 64KB, turning ordinary folder/message browsing and FastTransfer sync into a 500 - both now clamp to the response's real 16-bit size limit and let the client page for the rest
- Fixed RopSetColumns silently succeeding for a missing or wrong-type table handle instead of returning MAPI_E_INVALID_OBJECT
- Fixed RopLogon wholesale-resetting session.folderIds on a second logon within the same session, discarding every FID a client had already learned for a child folder
- Fixed RopSaveChangesMessage unconditionally blanking a calendar event's title (and, on at least one backend, its location/recurrence/reminder) on any update that didn't resend every property
- Fixed RopDeleteFolder checking only messageRepo/calendarEventRepo for emptiness, letting a non-empty Contacts/Tasks folder be deleted without DEL_MESSAGES and orphaning its Contact/Task rows
- Fixed meeting-response correlation being unreachable by a real Outlook client - this server's own outgoing invites never carried a PidLidGlobalObjectId, so decodeGlobalObjectId now recognizes the real "vCal-Uid"-wrapped VCALID form a client synthesizes from the plain iCalendar UID instead of assuming raw bytes
- Fixed a GAL search (NSPI GetMatches) TypeError on a Contact whose emails field is missing entirely
- Fixed changelog

### Removed
- Removed unused files

[Unreleased]: https://github.com/rapidmx/mapi-plugin/compare/v1.0.0-beta.11...HEAD
[1.0.0-beta.11]: https://github.com/rapidmx/mapi-plugin/compare/v1.0.0-beta.10...v1.0.0-beta.11
[1.0.0-beta.10]: https://github.com/rapidmx/mapi-plugin/compare/v1.0.0-beta.9...v1.0.0-beta.10
[1.0.0-beta.9]: https://github.com/rapidmx/mapi-plugin/compare/v1.0.0-beta.8...v1.0.0-beta.9
[1.0.0-beta.8]: https://github.com/rapidmx/mapi/compare/v1.0.0-beta.7...v1.0.0-beta.8
[1.0.0-beta.7]: https://github.com/RapidMX/mapi/compare/v1.0.0-beta.6...v1.0.0-beta.7
[1.0.0-beta.6]: https://github.com/RapidMX/mapi/compare/v1.0.0-beta.5...v1.0.0-beta.6
[1.0.0-beta.5]: https://github.com/RapidMX/mapi/compare/v1.0.0-beta.4...v1.0.0-beta.5
[1.0.0-beta.4]: https://github.com/RapidMX/mapi/compare/v1.0.0-beta.3...v1.0.0-beta.4
[1.0.0-beta.3]: https://github.com/RapidMX/mapi/compare/v1.0.0-beta.2...v1.0.0-beta.3
[1.0.0-beta.2]: https://github.com/RapidMX/mapi/compare/v1.0.0-beta.1...v1.0.0-beta.2
[1.0.0-beta.1]: https://github.com/RapidMX/mapi/compare/v1.0.0-beta.0...v1.0.0-beta.1
[1.0.0-beta.0]: https://github.com/RapidMX/mapi/releases/tag/v1.0.0-beta.0
