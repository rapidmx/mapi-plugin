# mapi — Design Decisions & Session Notes

This file exists so that Claude sessions working in this repo don't re-litigate settled
decisions or re-discover the same issues from scratch. It is local to this repo (not tied to
any one machine's global Claude memory), so it travels with the code.

**Maintenance rule:** when a standing decision changes, update the section below in place
(don't just append a contradiction lower down). When a new investigation/session produces a
decision, finding, or reverted approach worth remembering, add a dated entry under Session Log.
Keep entries terse — this is a reference, not a transcript.

## Standing design decisions & constraints

- **Vulnerability/review threat model: externally-exploitable only.** Only count issues reachable
  from a downstream, untrusted HTTP client hitting a service built on this package (anonymous or
  low-privilege caller). Do NOT flag developer-only footguns or purely theoretical races with no
  concrete external trigger path.
- **Commit discipline.** Don't `git commit` unless explicitly asked for *that specific piece of
  work*. An autonomous-execution/"commit as you go" approval given for one approved plan (e.g. via
  plan mode) is scoped to that plan only — it does not carry forward to later, separate requests in
  the same session, even ones that look similar in kind (a follow-up review-and-fix pass, a
  refactor, a new feature), and even after a full review-and-fix cycle with passing tests. Default
  to leaving changes staged/unstaged and saying so; only commit automatically within the exact
  scope of a plan that was explicitly approved as autonomous. If unsure whether new work falls
  inside that scope, treat it as outside and ask.
- **Commit message style: a flat list of one-line, verb-led items — no summary/title line, no
  `-`/`*` bullet markers.** This isn't just a style preference — it's dictated by how `release`
  (`@rapidrest/cli`) actually builds `CHANGELOG.md`. `collectChangelogBullets`/
  `classifyChangelogLine` (that repo's `src/lib/release.ts`) parse `git log --pretty=format:%B` and
  treat **every non-blank line of a commit's full message as its own changelog bullet** — there is
  no subject/body distinction. A conventional "short imperative subject + blank line + prose body"
  commit therefore leaks one changelog bullet per body sentence, and a `-`/`*`-prefixed line breaks
  `classifyChangelogLine`'s verb detection (it reads the line's first whitespace-delimited word as
  the verb; a leading `-` defeats that lookup and the dash leaks into the changelog text as
  `"- - Added foo"`). Correct format:
  - No separate summary/title line — if a commit needs an overview, that overview is itself just
    one more flat line, not a heading distinct from the rest.
  - No bullet-marker prefix of any kind — write bare lines.
  - Lead each line with an imperative verb where it fits: `Add`/`Fix`/`Remove` (and `-ing` forms)
    are recognized and become `Added`/`Fixed`/`Removed` entries; `Configuring`/`Converting`/
    `Refactoring`/`Updating`/etc. become `Changed`. Anything else still works, defaulting to
    `Changed` verbatim — see `CHANGELOG_VERB_REWRITES` in that repo's `src/lib/release.ts` for the
    full map.
  - A blank line before a trailing git trailer (`Co-Authored-By:`, `Signed-off-by:`, etc.) is fine
    — trailers matching `CHANGELOG_NOISE_PATTERNS` are dropped from the changelog — but nothing
    else should follow the item list.
  This mirrors JP's standing convention across his other repos; copy this exact rule verbatim into
  each sibling repo's own NOTES.md rather than paraphrasing it, since the paraphrase is what caused
  this to be gotten wrong in the first place (see `@rapidrest/cli`'s own NOTES.md, 2026-09-07 entry,
  for the full incident writeup and the `CHANGELOG_NOISE_PATTERNS` fix that accompanied it).

## Session Log

### 2026-09-06 — Repo split: `@rapidrest/mail` → four RapidMX packages

- **This repo is `@rapidmx/mapi`**, carved out of the former monolith `@rapidrest/mail`
  (`d:\github\rapidrest\mail`, still present there for reference/history) — was `src/mapi`/
  `test/mapi` there, moved to this repo's own root (not nested under a `mapi/` folder).
- Depends on [`@rapidmx/restapi`](https://github.com/RapidMX/restapi) (real published `^0.1.0`,
  not a local `portal:` link — see below) for the mailbox/folder/message/contact/calendar models,
  `resolveCallerMailboxUid`/`RecoverableRepoUtils`/`sendComposedMime` REST-layer helpers,
  `BlobStore`, and the scan pipeline.
- **Built from the start with two lessons already learned from the `activesync`/`autodiscover`
  splits (see their own NOTES.md for the full diagnosis of each) — do not regress either**:
  1. `vitest.config.ts`'s `ssr.noExternal` lists `@rapidmx/restapi` alongside
     `@rapidrest/service-core`/`@rapidrest/core`. Without it, Vite's SSR pipeline loads a second,
     natively-required copy of the framework packages for anything reached *through*
     `@rapidmx/restapi`, breaking `ModelUtils`'s static SQL-typeorm state (and, more generally, any
     `instanceof`-based framework behavior) between the two copies - manifesting as bare 500s from
     otherwise-correct code.
  2. `test/server-{mongo,sql}/models/index.ts` uses a **named** re-export of just the 5 model
     classes MAPI needs (`Folder`/`Message`/`Mailbox`/`Contact`/`CalendarEvent`), not a wildcard
     `export * from "@rapidmx/restapi/mongo"` - the wildcard form also pulls in every REST
     route/job class restapi bundles alongside its models, which the ClassLoader would then
     discover and try to initialize with no configured dependencies for them.
- Confirmed via a grep of `test/routes/{mongo,sql}/MapiEmsmdbRoute.test.ts` that (unlike
  `activesync`, which needed `FolderRoute`/`MessageRoute` REST mounts for its "renamed/deleted via
  REST" test scenarios) none of MAPI's own tests call out to a REST endpoint outside its own
  `/mapi/emsmdb`/`/mapi/nspi` base paths - no extra route mount files were needed here.
- For the original design rationale behind the ROP codec, the named-property registry, the
  calendar recurrence/timezone/GlobalObjectId codecs, and every other decision baked into this
  code, see the monolith's own `.claude/NOTES.md` (`d:\github\rapidrest\mail`) - that history
  wasn't duplicated here since it predates this repo's existence.

### 2026-09-08 — Synced against `@rapidmx/restapi` 0.2.0→0.3.x: fixed a stale peer-dependency range and a
broken test mock, added Contacts/Tasks browsing + deferred-send/read-receipt Submit support

Prompted by a "review restapi's latest activity and implement missing/broken MAPI protocol features" request.
Reviewed every restapi commit since this repo's own initial split (EAS-sync fields, TaskList, mail filters,
scheduled send, signatures/OOF, distribution lists, transport rules, audit log, domains, focused inbox,
bookings, plus-addressing, branding, read/delivery receipts) and filtered to what's actually MAPI/HTTP-protocol
relevant for a real Outlook client - **deliberately excluded** Branding/Bookings/TransportRule/AuditLog/Domains/
DNS-setup/DistributionList-admin/MailFilterRule-as-Outlook-Rules and EAS-only fields, since none of those are
things a MAPI client consumes (they're REST-admin/webmail/EAS-only surface) - implementing Outlook "Rules" ROPs
against `MailFilterRule` in particular was considered and set aside as a separate, much larger effort (real
`RopGetRulesTable`/`RopModifyRules` support, still listed as a documented gap) rather than folded into this pass.

- **Found two genuine regressions from the version drift** (peer dep pinned `0.1.x` in `package.json` while
  `devDependencies`/tests already ran against `^0.3.0` - a real installability bug for any downstream consumer
  on a modern restapi): bumped the peer range to `^0.3.0`. Separately, `test/rop/RopSubmitMessageHandler.test.ts`'s
  own `makeScanPipeline()` mock predated restapi's new `deriveConversationId()` call inside `scanAndRelay()`
  (added for `Message.conversationId`) - the mock's `run()` result had no `references`/`inReplyTo` fields, so
  `references[0]` threw on `undefined`. Fixed by adding both fields to the mock (production code was never
  actually broken - a real `ScanPipeline.run()` always populates them - only this repo's own test double was
  stale).
- **Added Contacts/Tasks folder browsing** - the biggest actual protocol gap found, not called out in the
  README's own "documented gaps" list before this (an oversight, not a deliberate exclusion): `RopGetContentsTable`
  already special-cased `FolderType.CALENDAR` to route to `calendarEventRepo` instead of `messageRepo`, but
  `CONTACTS`/`TASKS` folders fell through to `resolveFolderMessages()` and would always show empty, since
  `Contact`/`Task` are separate restapi entities, never `Message` rows. New `ContactTarget.ts`/`TaskTarget.ts`
  mirror `CalendarEventTarget.ts`'s exact shape (`resolveXInfo`/`resolveFolderXs`, `"contact:"`/`"task:"` target
  prefixes reusing `MessageTarget.assignOrGetMid`'s MID registry as-is). Wired through `RopGetContentsTableHandler`,
  `RopOpenMessageHandler` (subject/title resolution), and `PropertyResolvers.resolvePropertyValues` (new
  `contactValueFor`/`taskValueFor`). Task's own MAPI properties (`PidLidTaskStatus`/`PercentComplete`/
  `TaskDueDate`/`TaskComplete`, `PSETID_Task`) are named properties per real `[MS-OXOTASK]` - new
  `TaskNamedProperties.ts` mirrors `CalendarNamedProperties.ts`'s pattern; verified every LID/GUID against
  Microsoft's own Learn docs before implementing (`PidLidTaskStatus` = `0x8101`, `PercentComplete` = `0x8102`,
  `TaskDueDate` = `0x8105`, `TaskComplete` = `0x811C`, `PSETID_Task` = `00062003-0000-0000-c000-000000000046`) -
  Contact's own fields are all plain fixed `PidTag`s (`GivenName`/`Surname`/`EmailAddress`/
  `BusinessTelephoneNumber`/`CompanyName`/`Title`), no named-property lookup needed.
  **Deliberately read-only**: no `RopSetProperties`/`RopSaveChangesMessage` write-back for either kind in this
  pass (mirrors this package's own existing "browse only" precedent, e.g. no folder rename/move) - the REST API
  remains the way to create/edit Contacts/Tasks.
  **`contactRepo`/`taskRepo`/`contactClass`/`taskClass` were added to `RopContext` as *optional* fields**,
  not required ones - making them required would have forced every one of this codebase's ~15 existing
  `RopContext`-object-literal test fixtures (predating this feature) to grow four new fields each for no
  behavioral reason. `BaseMapiEmsmdbRoute` always populates them for real; every consumer degrades a
  `"contact:"`/`"task:"` target to empty rows/type-appropriate defaults when the repo is absent from a test
  double, the same "missing data, not a crash" stance already used for a vanished row.
  Added a real end-to-end HTTP+DB integration test in both `test/routes/mongo/MapiEmsmdbRoute.test.ts` and its
  `sql/` counterpart (Logon → OpenFolder(root) → GetHierarchyTable → SetColumns([DisplayName, FolderId]) →
  QueryRows to discover the Contacts/Tasks FID - neither is in `RopLogon`'s fixed 13-folder `FolderIds` list,
  unlike Inbox - then OpenFolder → GetContentsTable → SetColumns → QueryRows against it), plus full unit
  coverage for the new target/resolver files and the extended dispatch/handler branches.
- **Added `PidTagDeferredSendTime`/`PidTagReadReceiptRequested` support to `RopSubmitMessage`** - two documented
  gaps a real Outlook client's own Send dialog exercises routinely ("Do not deliver before" / "Request a read
  receipt"), now tracked by `RopSetPropertiesHandler` and honored in `RopSubmitMessageHandler`: a future deferred
  time parks the composed-but-unrelayed message in Outbox with `Message.scheduledSendTime` set (mirroring
  `BaseMessageRoute.send()`'s own identical branch, letting `ScheduledSendJob` relay it later via the same
  `scanAndRelay()` call), and a read-receipt request attaches a real `Disposition-Notification-To` header via
  `MailComposer`'s own `headers` option (no need to reach into restapi's internal `MimeHeaderUtils`) plus records
  `requestReceipt` on the persisted message. Deliberately **not** wired to this mailbox's own
  `alwaysRequestReceiptInternal`/`External` defaults - those apply to the REST/webmail compose path only,
  consistent with this handler's existing "no `RopModifyRecipients`, no address-book resolution" pragmatic scope.
- **Considered and explicitly deferred**: exposing `Message.inferenceClassification` (Focused Inbox) via
  `PidTagInferenceClassification` - restapi's own doc comment says this field's values were deliberately chosen
  to match what MAPI exposes, an open invitation to wire it up - but repeated web searches for this property's
  real fixed tag ID/wire-value encoding came back inconclusive (no canonical `[MS-OXPROPS]`/`[MS-OXCMSG]` page
  found, unlike every other property this session touched, all of which were verified against a real Microsoft
  Learn page first). Implementing a guessed tag/value mapping for a real protocol field risked silently
  mis-rendering a genuine Outlook client's Focused/Other view - worse than not exposing it at all - so this was
  left as a documented gap in the README rather than shipped on a guess. Revisit if a reliable source for the
  real encoding turns up.
- Full suite: 447/447 passing (up from 412 pre-session's own count after fixing the 5 pre-existing failures,
  net +35 new tests), 100% statement/function/line coverage, 98.82% branch (comfortably above this package's own
  95% floor). `yarn build`/`yarn tsc --noEmit`/`yarn lint` all clean. Not committed - left staged/unstaged per
  the standing commit-discipline rule.

### 2026-09-09 — Two-agent adversarial review + fix pass: DoS, correctness, and O(n²) session-registry bugs

Prompted by "review this codebase using two adversarial agents, look for correctness, bugs, vulnerabilities and
bottlenecks" followed by "fix everything." Ran two independent `general-purpose` agents in parallel (one
correctness/security-focused, one performance/independent-correctness-focused, neither seeing the other's
output), then personally verified every high/critical claim by reading the actual source before trusting it -
one claimed bug (recurrence/location silently cleared on `RopSaveChangesMessage` update) turned out to be
backend-dependent when traced through `RepoUtils.update()` (SQL's TypeORM skips `undefined` fields; Mongo's
behavior depends on a third-party BSON default I couldn't pin down statically) - only the `title` half of that
claim was unconditionally true, so that's what got reported as CONFIRMED, and the fix (below) closes both
regardless of the backend nuance by making it correct-by-construction instead of relying on undefined-skipping.

**Fixed, in order of severity:**
- **Critical DoS**: `PropertyValue.ts`'s `readCountedArray()` looped a client-controlled `uint32` element count
  with no cap, and `BufferCursor.ts`'s `readNullTerminatedUtf16LE`/`readNullTerminatedString8` never threw at
  end-of-buffer (returned `""` and kept going) despite the class's own doc comment claiming every read "bounds-
  checks and throws" - together, one `RopSetProperties` call (reachable before the handler even validates its
  target handle) with a `PtypMultipleString` count of `0xFFFFFFFF` could hang or OOM the whole process. Both
  string readers now throw when no terminator is found before the buffer ends; `readCountedArray` independently
  rejects any count exceeding the buffer's actual remaining bytes, since every element needs ≥1 byte.
- **`readBytes()` silently clamping instead of throwing** on a negative or oversized length - `Buffer.subarray`
  doesn't throw, so a malformed `RopSize < 2` rewound `decodeRopBuffer`'s cursor backwards and mis-parsed the
  handle table instead of failing loudly. Now bounds-checked explicitly.
- **Four separate uint16-response-field overflows** that turned ordinary (not just malicious) client usage into
  a 500: `RopBuffer.ts`'s `encodeRopBuffer` writes `RopSize` as a uint16 with no cap on `ropsList.length`
  (`RopQueryRows` against a few hundred messages already exceeds it), `RopReadStreamHandler`'s `DataSize` and
  `RopFastTransferSourceGetBufferHandler`'s `TransferBufferSize` both used a 32-bit requested/remaining size
  unclamped before a 16-bit write. Fixed `RopReadStream`/`FastTransferSourceGetBuffer` by clamping to `0xFFFF`
  and letting the client page for the rest (spec-legal - real Exchange does exactly this). **Not** fixed at the
  `encodeRopBuffer` level itself (would need a shared byte-budget threaded through every ROP in an `Execute`
  batch, a materially bigger change) - `RopQueryRows` was the one call site actually exercising this, and it's
  now addressed by the same pattern as the other two if it ever needs it; flagged as a narrower remaining gap
  in that handler's own doc comment rather than silently left unaddressed.
- **`RopSetColumns` silently succeeding for a missing/wrong-type handle** (`if (handle) {...}` with no `else` -
  every sibling ROP validates and returns `MAPI_E_INVALID_OBJECT`). Now matches the sibling pattern.
- **`RopLogon` wholesale-replacing `session.folderIds`** on every call, discarding child-folder FIDs a client
  had already learned via `RopGetHierarchyTable`/`RopQueryRows` on a re-logon (spec-legal within one session) -
  and risking a re-issued FID silently pointing at a *different* folder. Fixed by having `RopLogon` assign its
  13 special-folder FIDs through the exact same `FolderTarget.assignOrGetFid` reuse-or-mint logic child folders
  already use, instead of a bespoke always-reset-to-1 counter.
- **`RopSaveChangesMessage` blanking a calendar event's title** (unconditionally, any backend) whenever an
  update touched some other property without resending `PidTagSubject` - `decodeCalendarFieldsFromDraft`
  defaulted an absent `PidTagSubject` to `""` instead of leaving it `undefined` the way every other optional
  field already did, and the update path had no existing-value fallback for it (unlike `startDate`/`endDate`/
  `busyStatus`/`timezone`, which already did). Fixed `title` to stay `undefined` when absent and added the same
  `?? existing.xxx` fallback to `location`/`recurrenceRule`/`reminderMinutesBeforeStart` too, closing the
  backend-dependent uncertainty on those by making the fallback explicit rather than relying on `update()`'s
  undefined-handling.
- **`RopDeleteFolder` not knowing about Contacts/Tasks folder content** - emptiness was checked only against
  `messageRepo`/`calendarEventRepo`, so a non-empty Contacts/Tasks folder could be deleted without `DEL_MESSAGES`
  and orphan its `Contact`/`Task` rows (a real regression from the 2026-09-08 Contacts/Tasks feature work, not
  the documented "browse-only" limitation). Both repos now participate in the emptiness check and delete
  cascade, including the recursive subfolder path - guarded the same optional-repo way `RopGetContentsTableHandler`
  already does.
- **Meeting-response correlation was unreachable by a real Outlook client, not just a documented limitation**:
  `GlobalObjectId.ts`'s own doc comment claimed this server always embeds the `GlobalObjectId` it generates into
  the invites it sends - but `RopSubmitMessageHandler.submitAppointment()`'s ICS only ever emits a bare `UID:`
  line (invites go out as plain SMTP/iCalendar, not native MAPI Store objects), and `encodeGlobalObjectId` had
  zero call sites outside its own test file. A real client therefore always *synthesizes* its own
  `PidLidGlobalObjectId` from that bare UID using the standard `"vCal-Uid"`-wrapped `VCALID` form
  (`[MS-ASEMAIL]` §2.2.2.37 / the equivalent `[MS-OXCICAL]` algorithm - verified against Microsoft's own docs
  before implementing, the same bar applied to every other spec claim in this codebase), which the codec's
  bare-UTF8-bytes decode didn't recognize. Rewrote `encodeGlobalObjectId`/`decodeGlobalObjectId` to build/parse
  the real `VCALID` wire form (falling back to raw bytes for the `OutlookID` case a native Exchange server
  produces, which this server's own SMTP-based flow never does but decodes for completeness) - this is a real
  functional fix, not just documentation, since accept/decline responses from a genuine Outlook client would
  never have correlated back to the right `CalendarEvent` before this.
- **NSPI `GetMatches` `TypeError`** on a `Contact` whose `emails` field is missing entirely (optional-chained
  the array element, not the array itself) - low real-world likelihood (`@rapidmx/restapi`'s `Contact` always
  initializes `emails: []`) but cheap to close and inconsistent with `ContactTarget.ts`'s own fully-guarded
  version of the identical lookup.
- **O(n²) session registries**: `FolderTarget.assignOrGetFid`/`MessageTarget.assignOrGetMid` did a linear scan
  of a session-lifetime-cumulative map on every single table row, and `NamedPropertyRegistry
  .assignOrGetNamedPropertyId` did the same plus a `Math.max(...spread)` id-allocation that risked
  `RangeError: Maximum call stack size exceeded` on a session that registered enough named properties (past
  V8's argument-spread limit) - the code's own prior comments claimed these scans were "never a real cost,"
  which wasn't true once you account for the registry being cumulative across a whole session's worth of table
  paging, not reset per page. All three now use an O(1) reverse-index/counter pair (`session.folderTargetIds`/
  `nextFolderId`, `messageTargetIds`/`nextMessageId`, `namedPropertyIds`/`nextNamedPropertyId` - new
  `MapiSessionContext` fields, small and purely additive to session-cache size, unlike the alternative of
  caching whole resolved entities would have been). `assignOrGetNamedPropertyId` also now returns `0x0000`
  ("unmappable," a real spec-defined value already used for `Kind = 0xFF`) once the `0x8000`-`0xFFFF` ID space
  is exhausted, instead of assigning an out-of-range id that later crashed a 16-bit write.
- **Redundant identical per-row queries**: `RopQueryRowsHandler`/`RopGetPropertiesSpecificHandler` now share one
  request-scoped `ResolutionCache` (`PropertyResolvers.ts`) across every row/column resolved in one call, so a
  hierarchy table's `hasChildren` (which re-fetched the mailbox's *entire* folder list per row via
  `resolveFolderChildren`) and a calendar table's organizer/attendee mailbox lookup each happen at most once per
  call instead of once per row. **Deliberately not fixed**: the deeper per-row `repo.findOne(uid)` N+1 for
  distinct message/calendarEvent/contact/task rows - a real batched fix needs a backend-agnostic "fetch these
  uids" query this codebase's `RepoUtils<T>` abstraction doesn't expose (Mongo's `$in` and TypeORM's `In()`
  aren't interchangeable at a call site that doesn't know which backend it's running against), and caching
  whole resolved entities in session state instead would make the already-flagged session-serialization-size
  problem (large `session.handles`/`messageIds` re-`JSON.stringify`'d on every `Execute`) significantly worse -
  documented as a known limitation in `RopQueryRowsHandler`'s own doc comment rather than attempted blind.
  Also parallelized `RopLogonHandler`'s four sequential per-special-folder repo lookups via `Promise.all`.
- Full suite: 472/472 passing (up from 447, +25 net new tests targeting each fix specifically - not just
  "still passes" coverage). 100% statement/function/line, 99.14% branch. `yarn build`/`yarn tsc --noEmit`/
  `yarn lint` clean. Not committed - left staged/unstaged per the standing commit-discipline rule.

### 2026-09-13 — restapi 0.3.1 → 0.8.0 sync: a real `like()`-operator regression fix plus Categories

Prompted by "tons more changes have gone into restapi... do a full review of it, including the spec files...
and implement any missing MAPI features that restapi now supports." Bumped the `@rapidmx/restapi` peer/dev
dependency to `^0.8.0` and `@rapidrest/service-core` to `2.x`/`^2.0.0`, then read every file under
`../restapi/specs/*.md` and diffed restapi's own changelog/commit history against what this package already
covers, the same review discipline as the two prior sessions above.

**Found and fixed a real regression**: `@rapidrest/service-core` 2.0's `ModelUtils` `like()` operator changed
from raw substring matching to glob syntax (`*`/`?`, anchored `^...$` by default via `globToRegExpSource`/
`globToLike`). `NspiGetMatchesHandler.ts`'s `escapeForLikeQuery()` was written for the *old* semantics -
regex-escaping the user's search term so it behaved as a literal substring - which under the *new* semantics
just produces an exact-match glob pattern with no wildcards, so `RopGetMatches`/GAL "search as you type" broke
silently (0 results for the previously-passing `Filters by a ContentRestriction search term, matching only
Jane` Mongo test; SQL coincidentally still worked at the time, which would have hidden this backend-specific if
only one route's tests had been checked). First fix attempt: deleted `escapeForLikeQuery()` and had both
`MapiNspiRouteMongo.likePattern()`/`MapiNspiRouteSQL.likePattern()` wrap the *raw* term in `*...*` glob
wildcards instead. **Superseded within the same session**: JP pointed out `@rapidrest/service-core` 2.0 also
added a genuine `regex()` operator (validated server-side against ReDoS shapes via `isUnsafeRegexPattern`,
compiled case-insensitively and identically on both backends - Mongo `$regex`/`$options:"i"`, SQL `~*`/`REGEXP`
with a `better-sqlite3` `REGEXP` function auto-registered per-connection) - a strictly better fit for "literal
substring, case-insensitive" than glob-wrapping `like()`, since a `regex()` pattern built from `StringUtils
.escapeRegExp(term)` has no residual wildcard ambiguity for a literal `*`/`?` in the search term (glob syntax
has no escape for those two characters; `like()` would always have that gap). Switched to `regex()` and, since
it compiles identically on both backends (unlike the old glob-wrapping, which coincidentally also ended up
identical but didn't have to be), deleted the per-backend `likePattern()` hook entirely -
`BaseMapiNspiRoute`/`MapiNspiRouteMongo`/`MapiNspiRouteSQL` no longer need one. Guidance for future sessions:
prefer `regex()` over hand-rolled `like()` glob-wrapping specifically for "does this field contain this literal
substring" queries, but don't reach for `regex()` where a simpler operator (`eq`/`in`/plain `like()` for a
genuine prefix/suffix-style client-supplied glob) already does the job - it's the right tool for this one
narrow shape, not a blanket replacement for `like()`.

**Implemented**: read-only Outlook Categories (`PidNameKeywords`, `PS_PUBLIC_STRINGS` GUID
`00020329-0000-0000-c000-000000000046`, `Kind = "name"` not `"lid"` - the first named property in this
codebase keyed by name instead of LID, since every previous one used `PSETID_*`/lid pairs) backed by restapi's
pre-existing `Label`/`Message.labelUids` model (Gmail-style label-uid list, independent of folder placement).
Wired the same optional-`RopContext`-field pattern as `contactRepo`/`taskRepo` (`labelRepo`/`labelClass`,
`BaseMapiEmsmdbRoute` populates both) and extended `PropertyResolvers.ts`'s existing per-call `ResolutionCache`
so a table of 100 categorized messages fetches a mailbox's `Label` set once, not once per row - the same
n+1-avoidance pattern the 2026-09-09 session established for folder lists. Read-only by design: there's no
general "edit an existing persisted message's properties" ROP path in this codebase at all (matches the
existing Contacts/Tasks browse-only precedent) - assigning/removing labels goes through the REST API.

**Deliberately excluded** (confirmed against restapi's own specs, not just left unmentioned): End-to-End
Encryption (`specs/end-to-end_encryption.md` explicitly designs it so a real Outlook client needs *no* protocol
changes at all - native S/MIME - and KeyVault/Escrow/Discovery are REST/web-client-layer only) and the search
overhaul (`specs/search.md` states outright that parity with plaintext search "on Outlook... [is] Not
achievable; documented instead" once E2E is on - "a documented limitation, not a defect"). Also excluded every
compliance/admin-only feature (Branding, Bookings, TransportRule, Legal Hold, Matter, GDPR export/erasure) as
categorically not MAPI-protocol surface, consistent with this conversation's standing scope rule.

**Verified, not assumed**: confirmed via a targeted test change (`FolderType.ARCHIVE` in place of `USER` in an
existing folder-browsing test) that restapi's Archive folder type needs zero new mapi code, since generic
folder/message browsing already treats any non-special-cased `FolderType` identically.

Full suite: 486/486 passing (up from 472, +14 net new tests). 100% statement/function/line, 99.16% branch.
`yarn build` clean. Encountered and correctly identified as *not* a regression: repeated, non-reproducible
full-suite-only failures (different files, different garbage values each run) matching this repo's own
documented `mongodb-memory-server`/raw-socket-client flakiness pattern under load - every affected file passed
100% in isolation immediately after, and the specific failure was never the same twice. Not committed - left
staged/unstaged per the standing commit-discipline rule.

### 2026-09-14 — README package names

README still referenced the pre-rename package names (`@rapidmx/mapi`, `@rapidmx/autodiscover`, including the
npm badge and `import ... from "@rapidmx/mapi/mongo"`). Updated them to `@rapidmx/mapi-plugin` /
`@rapidmx/autodiscover-plugin`, matching `package.json`. Docs only, no code change, not committed.

### 2026-09-14 (2) — Round-2 review fixes (session ownership, NSPI GetMatches pattern/limit)

Each finding was confirmed in code first. Not committed; no version or peerDependency changes.

- **MEDIUM: `Execute`/`Disconnect` never checked who owned the session.** Both loaded the session named by the
  `MapiContext` cookie without comparing `session.userUid` to the authenticated JWT user. Anyone holding
  another user's cookie value could run ROPs against that user's mailbox, or end their session. A new private
  `loadOwnSession(req, user)` returns the session only when `session.userUid === user.uid`.
  - `Execute` treats a mismatch exactly like a missing session (`ERROR_SESSION_NOT_FOUND`).
  - `Disconnect` leaves another user's session alone and returns the same success body as for an unknown
    session.
  - The real-HTTP test is in `test/routes/mongo/MapiEmsmdbRoute.test.ts`: a second user's Execute gets
    session-not-found, their Disconnect doesn't end the owner's session, and the owner's Execute still succeeds.
- **LOW: `NspiGetMatches` pattern and limit.** `escapeRegExp(searchTerm)` was unbounded, so a long or
  metacharacter-heavy term produced a `regex()` operand over service-core's 100-character guard and failed the
  whole NSPI call. Results were also silently capped at the repo's 100-row default whatever `RowCount` asked
  for.
  - Copied activesync's `RegexPatternUtils.ts` (`boundedEscapedPattern`, `MAX_REGEX_PATTERN_LENGTH`) and its
    test into this repo, since there's no shared package. Keep the two copies in sync.
  - Every query now passes `limit = clamp(RowCount, 1, MAX_MATCH_ROWS = 1000)` in both the query (SQL reads
    that) and the options (Mongo reads that).
  - Known limitation: with a filter, each of the 4 per-field queries is limited separately before the in-memory
    merge and sort, so `TotalRecs` counts at most what was fetched, not every match.
  - Tests are in `test/RegexPatternUtils.test.ts` (copied) and `test/nspi/NspiGetMatchesHandler.test.ts`
    (truncation and limit clamping; existing expectations updated for `limit`).

### 2026-09-14 (3) — Round-3 review fixes (limits, session consistency, deletes, calendar, streams)

Each of the 15 findings was confirmed in code before fixing; all 15 were real. Not committed; no version or
peerDependency changes. 575 tests pass (was 495); coverage 100% statements/functions/lines, 99.12% branches;
`yarn lint` and `npx tsc --noEmit -p .` clean. (`tsconfig.test.json` still reports its pre-existing `RopContext`
fixture errors; it isn't part of the gate.)

- **Framing (1).** `decodeExecuteRequest()` (exported from `BaseMapiEmsmdbRoute.ts`) rejects `RopBufferSize` over
  32767 or past the body, and any decode error, with a 400. `decodeRopBuffer` rejects a handle table over 255 entries
  or with a partial entry. `encodeRopBuffer` writes the table as one buffer.
- **Tag arrays (2).** `readPropertyTagArray()` in `PropertyValue.ts` caps at `MAX_PROPERTY_TAG_COUNT = 256`. Past the
  cap it skips the bytes and returns `undefined`, so EMSMDB `RopSetColumns`/`RopGetPropertiesSpecific`/FastTransfer
  CopyTo/CopyProperties answer `MAPI_E_TOO_BIG` and later ROPs still parse. NSPI `readLargePropertyTagArray` throws.
  NSPI `MinimalIds` is skipped with one bounds-checked read instead of a loop.
- **Sessions (3, 8, 13).** `MapiSessionManager` no longer uses `RedisCache`: its per-process copy was served without
  checking Redis, so replicas diverged. It now injects the `cache` Redis client (`@Redis("cache", false)`) and uses
  `RedisMapiSessionStore` (every load is a Redis GET; save is a Lua compare-and-set on `version`), or
  `MemoryMapiSessionStore` without Redis (JSON strings, so loads are independent copies). `save()` returns
  `"saved" | "conflict" | "missing"`. On conflict, Execute answers `X-ResponseCode: 15` (Invalid Sequence); on
  missing, session-not-found. Side effects of that request's ROPs have already happened; only its handle changes are
  dropped. `create()` keeps a per-user index key and ends the oldest session past `MAX_SESSIONS_PER_USER = 20`.
  `load()` ends sessions older than `MAX_SESSION_LIFETIME_MS = 24h`. Execute also checks
  `resolveCallerMailboxUid(user) === session.mailboxUid` every time.
- **Deletes (4, 5, 6).** restapi's `assertNotOnLegalHold` is NOT exported from the package root (only used inside
  restapi), so `RopDeleteFolder` never purges: `DELETE_HARD_DELETE` is ignored and everything is soft-deleted, which a
  hold allows. Messages deleted by `RopDeleteFolder` and `RopDeleteMessages` are audited as `MESSAGE_DELETE` through
  `recordAuditLog` via an optional `RopContext.audit`. The routes set `auditLogClass` (`AuditLogEntryMongo`/`SQL`);
  the Mongo test server now exports `AuditLogEntryMongo` so the integration test can check the row. The subfolder walk
  is iterative, has a visited set, and is limited to `MAX_FOLDER_DEPTH = 32` and `MAX_FOLDERS_PER_DELETE = 1000`
  (else `MAPI_E_TOO_COMPLEX`). Every child and item query is filtered by `mailboxUid`. Items are paged, at most
  `MAX_ITEMS_PER_DELETE = 10000` per ROP; past that it reports `PartialCompletion` and keeps the remaining folders.
- **Paging (6).** `RepoPaging.ts` (`findPage`/`findAllCapped`/`findWindow`) always passes `limit`/`page`/`sort` in
  both query and options and ends sorts with `uid`. Contents tables no longer snapshot rows: `RopGetContentsTable`
  stores `contentsKind`, and `RopQueryRows` reads just the window (`ContentsTable.ts`), at most 500 rows per call.
  Trade-off: an item added or removed between QueryRows calls can shift later rows. Sorts: messages
  `receivedDate DESC`, events `startDate DESC`, contacts `displayName`, tasks `title`, folders `name`. Hierarchy
  tables and FastTransfer folder dumps still resolve lists up front, capped at 10000. `resolveFolderContacts`/
  `resolveFolderTasks` were removed (unused).
- **Streams and handle data (7, 8, 15).** `HandleDataCache.ts` is a process-local, byte-bounded LRU with TTL
  (256 MB, 15 min) for a read stream's decoded body (parsed once at `RopOpenStream`) and a FastTransfer stream. A miss
  (another replica, eviction) rebuilds from the handle, so it never returns wrong data. Keys include the handle's
  `generation`. `transferBufferBase64` is gone; FastTransfer handles keep `transferSourceType`/`transferColumns`/
  `transferExcludeIds`, and a stream over 32 MB fails with `MAPI_E_TOO_BIG`. `RopWriteStream` appends base64 by
  re-encoding only the last quantum (`appendBase64`) and caps at 4 MB. Named properties are capped at 4096 per
  session. `RopDispatcher` stops after 1024 ROPs. Every handle assignment goes through `assignHandle()` (fresh
  `generation`); `releaseHandle()` (used by `RopRelease` and by reassigning an index) drops cached data and releases
  write streams whose `writeTargetHandleIndex`/`writeTargetGeneration` match. `resolveDraftBody` also matches the
  generation.
- **Calendar (9, 10).** `RopSaveChangesMessage` bumps `sequence` when start, end, location, recurrence or the invited
  addresses change (same fields as restapi's `BaseCalendarEventRoute.update()`), and keeps a still-invited attendee's
  response. `submitAppointment` only sends when the event is in the caller's mailbox and the organizer is one of the
  caller's addresses. `submitMeetingResponse` now mirrors REST `respond()`: it finds the caller's own copy
  (`icalUid` + `mailboxUid`, preferring the exception whose `recurrenceId` date matches the GlobalObjectId instance
  bytes), updates it (decline soft-deletes it), and sends an iTIP REPLY built with restapi's `buildEventIcs` through
  `mailTransport.send`, swallowing send errors. It never touches the organizer's copy.
- **Submit (11, 14).** The Sent Items copy stores `scanAndRelay`'s returned `raw`, `messageId`, `conversationId` and
  `encrypted`. Address helpers moved to `AddressList.ts`: `isPlainEmailAddress` rejects whitespace (CR/LF),
  display-name forms and separators. Submit refuses the whole send if any recipient is invalid; attendee lists and
  invites drop invalid addresses.
- **Time zones (12).** `decodeTimeZoneStruct` keeps minutes: a whole-hour offset gives `Etc/GMT±N`, a known DST-less
  fractional offset gives a real zone (`Asia/Kolkata`, `Asia/Kathmandu`, ...), anything else gives a fixed offset like
  `"-03:30"` (Node's `Intl` accepts it). Offsets outside UTC-12..UTC+14 decode as `UTC`. `encodeTimeZoneStruct` falls
  back to UTC for a zone `Intl` rejects instead of throwing.
- Tests: new `test/ProtocolLimits.test.ts`, `test/MapiSessionManager.test.ts` (fake Redis client emulating the
  CAS script), `test/rop/HandleDataCache.test.ts`, plus Mongo integration tests for the 400, mailbox reassignment,
  conflict/missing save (spy on `MapiSessionManager.prototype.save`) and delete + audit row.

### 2026-09-14 (4) — Round-4 review fixes (FastTransfer, output size, sessions, meetings, recipients)

All 13 findings (at HEAD feac876) were confirmed in code before fixing; all were real. Not committed; no version or
peerDependency changes. 632 tests pass (was 575); coverage 100% statements/functions/lines, 98.92% branches;
`yarn lint` and `npx tsc --noEmit -p .` clean.

- **FastTransfer (1).** New `HandleDataStore` (`HandleDataCache.ts`): `RedisHandleDataStore` (base64 values in Redis,
  `mapi.handle.<key>`, local LRU in front) when the `cache` Redis client exists, else `MemoryHandleDataStore`.
  `MapiSessionManager.handleDataStore` picks one; the route passes it as `RopContext.handleData` (`handleDataStoreOf`
  falls back to the process store). Built streams are stored with `FAST_TRANSFER_TTL_SECONDS = 600`.
  `loadFastTransferBuffer` returns the stream, `"tooBig"` or `"lost"`: a miss rebuilds only while
  `transferPosition === 0`; mid-transfer it fails with `MAPI_E_CALL_FAILED` instead of paging a different stream. The
  32 MB cap applies to rebuilds, and `buildFastTransferStream` checks size after every item (`FastTransferTooBigError`).
- **Output size (2).** `decodeExecuteRequest` now returns `maxRopOut` (default 65535 when absent). `dispatchRops`
  takes `{ maxOutputBytes }` (route: `MaxRopOut - 2 - 4*handles`, never over `MAX_ROPS_LIST_BYTES = 65533`), sets
  `context.ropOutputRemaining` per ROP, and replaces a response that still doesn't fit with `RopBufferTooSmall`
  (`0xFF`, SizeNeeded, the remaining request bytes) and stops. `RopReadStream` clamps to the room; `GetBuffer` honors
  `MaximumBufferSize` for `0xBABE` (earlier it ignored it) and the room, answering `NoRoom` at zero room;
  `RopQueryRows` returns only rows that fit (cursor advances past just those), `ecBufferTooSmall` when none fit.
- **Generations (3).** `MapiObjectHandle.generation`/`writeTargetGeneration` are now `crypto.randomUUID()` strings
  (`nextHandleGeneration` removed), so a request whose save lost can't reuse a cache key.
- **Session size (4).** Session = two keys with a hash tag: `mapi.session.{id}` (JSON) and `mapi.session.{id}.version`;
  the CAS Lua compares the version key only (no `cjson`). `save()` returns `"tooBig"` past `MAX_SESSION_BYTES = 4 MB`
  (Execute: `X-ResponseCode 0`, body `ErrorCode MAPI_E_TOO_BIG`). Write streams store offset-keyed chunks in the
  `HandleDataStore` (`writeStreamKey`, Redis hash field per offset, so a retried write replaces its chunk); only
  `writeSize` stays in the session (`writeBufferBase64`/`appendBase64` removed). `readWriteStream` reassembles;
  Submit fails `MAPI_E_CALL_FAILED` when chunks are gone. MIDs are capped at `MAX_MESSAGE_IDS = 20000`, evicting the
  oldest via `session.firstMessageId`. Folder IDs are not capped (bounded by the mailbox's folders, 10000 cap).
- **Occurrence responses (5).** `submitMeetingResponse` returns a ReturnValue. With an instance date and no exception
  copy, the series is never deleted/updated for accept/tentative; decline appends the occurrence start (series local
  time on that date in the event tz, DST-aware) to the caller's `recurrenceRule.exceptions`. The REPLY carries
  `RECURRENCE-ID` (restapi's `buildEventIcs` emits it when `recurrenceId` is set, and drops `RRULE`). Exception copies
  match by `recurrenceId` date in the event's time zone (invalid tz -> UTC).
- **Per-ROP failures (6).** `dispatchRops` wraps its reader in a Proxy that turns reader `RangeError`s into
  `DecodeError` (new, `BufferCursor.ts`, extends `RangeError`); `readPropertyValue`'s unknown type and
  `readCountedArray` throw `DecodeError` too. Unknown RopId/`DecodeError` -> route saves the session, then 400. Any
  other handler error -> that ROP's partial output is dropped and a failure response written: RopId, the byte at
  `responseHandleIndexOffset` (default 2; 3 for ROPs echoing OutputHandleIndex), `MAPI_E_CALL_FAILED` (or
  `MAPI_E_TOO_COMPLEX` for the budget), plus `failureTailBytes` zeros (Read/WriteStream); `hasNoResponse` for
  RopRelease. `decodeAppointmentRecurrence` reads deleted instance dates into `exceptions` (minus same-day modified
  ones, plus StartTimeOffset) and stops before ExceptionInfo instead of throwing. `RopSetProperties` uses the new
  `BufferReader.seek` to skip to `PropertyValueSize`'s end and answers `MAPI_E_INVALID_PARAMETER` for mismatched or
  undecodable values; a size past the request still throws (400).
- **Lock/replay (7).** Execute takes `mapi.session.{id}.lock` (SET NX PX, `SESSION_LOCK_TTL_MS = 120s`, token-checked
  release) before loading; busy -> `X-ResponseCode 15` with no side effects. The last response is stored per session
  with its `X-RequestId` (`.last` key); a repeated id is answered from it without dispatching or saving.
  `MapiSequence` validation skipped: the test client (and many real retries) resend the Connect cookie, and lock +
  replay cover overlap and retries.
- **Work budget (8).** New `rop/ExecuteBudget.ts`: 20000 rows resolved (`resolvePropertyValues` charges one per call)
  and 64 MB built (FastTransfer streams, parsed bodies) per Execute; `bodies` dedupes body parses by message target.
- **GlobalObjectId (9).** Non-vCal ids decode to the uppercase hex of the whole blob with bytes 16-19 zeroed
  ([MS-OXCICAL]); lookup retries lowercase. `globalObjectIdInstanceDate` exported. Unknown class / missing or malformed
  id -> `MAPI_E_INVALID_PARAMETER`; no match / not an attendee -> `MAPI_E_NOT_FOUND`.
- **Recipients (10).** `AddressList.ts`: split on `;` only, each entry via nodemailer's `addressparser` (already a
  dependency; `nodemailer/lib/addressparser/index.js`), control chars stripped from names, CR/LF entries invalid.
  Bare names resolve against the caller's contacts (exact `displayName`, exactly one match, first valid email).
  Submit: invalid/no recipients or no sender -> `MAPI_E_INVALID_PARAMETER`, unresolved names -> `MAPI_E_NOT_FOUND`
  (was `0x80070005`, access denied). There is still no recipient table (no RopModifyRecipients) to fall back on.
  MIME To/Cc/Bcc carry display names; Sent Items recipients store `displayName`.
- **Response codes (11).** Context not found is `X-ResponseCode 10` ([MS-OXCMAPIHTTP] "Context Not Found"; no spec
  copy in the repo, value from the spec's response-code table); `0x80040111` only in the body `ErrorCode`.
- **Concurrent Connects (12).** Per-user index is a sorted set updated by one Lua script (prune dead, evict oldest to
  under the cap, ZADD). The script builds session key names from ids (not in KEYS) - fine on one Redis node, not
  cluster-safe. The memory store does the same synchronously.
- **REPLY send (13).** restapi's `sendOrThrow` is NOT exported from the package root, so `rop/TransportSend.ts` copies
  its rule (throw unless accepted > 0 and rejected == 0); keep in sync. A failed REPLY -> `MAPI_E_CALL_FAILED`, the
  response stays recorded.
- Tests: new `test/Round4Review.test.ts`, `test/fakeRedis.ts` (script-aware fake used by `MapiSessionManager.test.ts`
  and `HandleDataCache.test.ts`); Mongo integration tests for code 10, lock busy (spy `acquireLock`), X-RequestId
  replay, 400-after-save, MaxRopOut -> RopBufferTooSmall, `tooBig`.
- Lesson: bash heredocs containing some quote/backtick mixes fail in this environment (`unexpected EOF`); write
  scripts with the Write tool instead.

### 2026-09-14 (5) — Round-5 review fixes (submit replay, work budget, handle data cleanup, meeting lookups, locks)

All 13 findings (at HEAD 25e46b1) were checked against the code and were real; none skipped. The MID-map redesign was
out of scope and not touched. Not committed; no version or peerDependency changes. 680 tests pass (was 632); coverage
100% statements/functions/lines, 98.95% branches; `yarn lint` and `npx tsc --noEmit -p .` clean.

- **Submit replay (1).** `MapiObjectHandle.submitted`: set on a mail or meeting-response draft as soon as Submit gets
  past the handle check, whatever the outcome (the finding asked for "successful or not"). A submitted handle answers
  `MAPI_E_INVALID_OBJECT` to Submit, `RopSetProperties` and `RopWriteStream` (checked via `writeTargetHandleIndex`).
  The draft's write-stream chunks are deleted once read. `ExecuteBudget.chargeSubmit()` caps Submits at
  `MAX_SUBMITS_PER_EXECUTE = 16` (past it: `MAPI_E_TOO_COMPLEX`). `MAX_RECIPIENTS_PER_MESSAGE = 500` in
  `AddressList.ts`, same as activesync's `MAX_COMPOSE_RECIPIENTS` (restapi has no such cap): Submit counts parsed
  To/Cc/Bcc entries before any contact lookup (`MAPI_E_TOO_BIG`); `RopSaveChangesMessage` refuses a meeting with more
  attendees (`MAPI_E_TOO_BIG`, nothing saved). Appointment Submit is no longer a relay at all (see 6), so it isn't
  marked and can be repeated after an edit.
- **Work budget (2).** `ExecuteBudget` gained `MAX_QUERIES_PER_EXECUTE = 20000`, `MAX_ROWS_FETCHED_PER_EXECUTE =
  100000`, and `assertBytesLeft()`; the constructor takes a third `limits` object. `findPage`/`findAllCapped`/
  `findWindow` take an optional trailing `budget`: the query is charged before `repo.find`, the returned rows after.
  Threaded through `resolveFolderChildren`/`resolveFolderInfo` (hierarchy tables and `hasChildren`),
  `resolveFolderMessages`/`resolveFolderCalendarEvents` (FastTransfer), `resolveContentsWindow`, and every paged
  query in `RopDeleteFolder` (`collectSubtree`, `hasItems`, the child check, `deleteItems`); each item delete and folder
  delete charges a query first. `resolveContentsKind` charges its `findOne`. `loadStreamBody`/
  `resolveMessageBodyBytes` refuse at `bytesRemaining <= 0` before `findOne`, and charge the raw blob before
  `simpleParser` (then the decoded body, as before). Per-row `findOne`s in `resolvePropertyValues` stay under the row
  budget.
- **Handle data cleanup and quotas (3).** `HandleDataStore` changed: `set`/`putChunk` take optional `owners` and
  return `boolean` (`false` = quota refused, nothing stored); new `deleteOwnedBy(index)`; `writeStreamKey` moved to
  `HandleDataCache.ts` (re-exported from `RopWriteStreamHandler.ts`). `handleDataOwners(sessionUid, userUid)`:
  session index `session.<uid>` capped at `MAX_HANDLE_DATA_BYTES_PER_SESSION = 64 MB`, user index `user.<uid>` at
  `MAX_HANDLE_DATA_BYTES_PER_USER = 128 MB` (raw bytes). Redis: owner indexes are sorted sets
  `mapi.handle-owner.<index>` (member = key, score = raw bytes, TTL 24h); `SET_WITH_QUOTA_SCRIPT`/
  `PUT_CHUNK_WITH_QUOTA_SCRIPT` prune members whose key is gone, sum the rest (excluding the key being written), and
  store only under the cap. The value itself goes through EVAL (up to ~43 MB of base64 for a 32 MB stream). Chunks are
  now stored as unpadded `base64url` so the chunk script sizes an existing chunk exactly (`floor(len*3/4)`); Node's
  `"base64"` decode reads both forms. `DELETE_OWNED_SCRIPT` returns the deleted members so the local LRU drops them too.
  Like `ADD_TO_USER_INDEX_SCRIPT`, the quota/delete scripts build data keys from members, so they are single-node only
  (not cluster-safe). Memory store: same accounting in-process (`HandleDataCache.sizeOf` added, a non-touching peek).
  Cleanup: `releaseHandle` records released `fastTransfer`/`stream` handles' store keys in a module `WeakMap` beside
  the session (`takeReleasedHandleData`); the route deletes them after a `"saved"` save (best effort). `destroy()`
  (Disconnect, lifetime expiry) and eviction in `create()` call `deleteOwnedBy(session index)` (errors swallowed).
  FastTransfer open/rebuild pass owners; a refused open releases the new handle and answers `MAPI_E_TOO_BIG`, a
  refused rebuild reports `"tooBig"`. `RopWriteStream` refused by quota answers `MAPI_E_TOO_BIG`.
  `RopFastTransferSourceCopyProperties` dedupes tags by property id (first wins).
- **Meeting lookups (4, 11).** `rop/RestapiRules.ts` copies restapi's `boundIndexedValue` and `asEntity` (neither is
  exported from `@rapidmx/restapi` 0.9.0; keep in sync with `util/ConversationUtils.ts`/`util/EntityUtils.ts`) and
  adds `literalQueryValue(v) = "eq(v)"`: service-core's `op(x)` regex is greedy, so `eq(ne(x))` compares against
  `ne(x)`. It still coerces operands (`null`, numbers, `me` -> 403 without a user), so callers also filter rows in
  memory and treat a thrown query as no match. `findCallerCopies` queries `icalUid: eq(boundIndexedValue(uid))` and
  keeps `row.icalUid === key`; `AddressList.findContactsByDisplayName` queries `displayName: eq(name)` and keeps exact
  matches. Grepped the rest of `src` for queries built from client/sender strings: NSPI GetMatches already escapes its
  regex (round 2); every other query uses server-assigned uids/targets.
- **Versioned updates (5).** `asEntity(repo, row)` wraps the `existing` argument of every update of a restapi-owned
  row read with `find`/`findOne`: both updates in `MeetingMessageClassHandler`, the new series-exception update, and
  `RopSaveChangesMessageHandler`'s appointment update.
- **Invites (6).** Chose to stop sending from MAPI. restapi's `MeetingSchedulingJob` already sends a REQUEST for every
  organizer-owned event with `inviteSequenceSent !== sequence`, claims each revision with a versioned update, and uses
  `buildEventIcs` (RRULE + exceptions), so it gives the only correct recurring invite; stamping from MAPI would have
  kept the single-instance ICS. `submitAppointment` now only checks the organizer (success, or `MAPI_E_INVALID_OBJECT`
  as before); `buildMeetingRequestIcs` and helpers were removed. Trade-off: a meeting saved via MAPI is invited on the
  job's next run even if never submitted (same as a REST-created meeting).
- **RopBufferTooSmall (7).** When a response doesn't fit and neither does `RopBufferTooSmall` (3 bytes + rest of the
  request), `dispatchRops` throws `ExecuteBufferTooSmallError`; the route saves the session (earlier ROPs ran) and,
  unless the save failed, answers `X-ResponseCode 0` with body `ErrorCode` `ecBufferTooSmall` (`0x47D`) and no ROP
  buffer. A pre-run check was tried first and rejected: it failed Executes whose next response would have fit.
- **Declining an exception copy (8).** Decline via an override row soft-deletes it and appends its `recurrenceId` to
  the series' `recurrenceRule.exceptions` (versioned, skipped when already present or the series doesn't recur),
  mirroring `ScanQueueJob`'s occurrence CANCEL. Shared helper `addSeriesException`.
- **Lock (9).** Ownership (`loadOwnSession` + mailbox check) now runs before `acquireLock`; the session is loaded again
  under the lock (gone -> code 10). The lock is renewed every `SESSION_LOCK_RENEW_MS = 40s` (`RENEW_LOCK_SCRIPT`,
  token-checked PEXPIRE; `setInterval` unref'd, cleared in `finally`). With an `X-RequestId`, `markInProgress` stores
  `{requestId}` (no body) in `.last` with the lock's TTL, renewed with the lock; `storedResponse` returns
  `"inProgress"` for it and the route answers code 15. `storeResponse` replaces the marker; a request that ends without
  storing one calls `clearInProgress` (only deletes a still-in-progress marker for that id).
- **ReadStream (10).** At zero room with data left, the handler writes the full would-be response without advancing
  `streamPosition`, so the dispatcher turns it into `RopBufferTooSmall`; `DataSize 0` now only means end of stream.
- **LRU session cap (12).** `MapiSessionManager.touch(session)` (`TOUCH_USER_INDEX_SCRIPT`, `ZADD XX`) on every
  Execute, best effort, so the cap evicts the least recently used session.
- **Draft values and logon (13).** `validDate` (unparseable -> field left unset) for start/end, `validMinutes`
  (non-negative safe integer) for the reminder. `RopLogon`'s well-known folder query sorts `dateCreated ASC, uid ASC`
  in the query object with `limit 1`, like `findOrCreateWellKnownFolder`.
- Tests: new `test/Round5Review.test.ts`; `fakeRedis.ts` emulates the five new scripts; updates to
  `MeetingMessageClassHandler`, `RopSubmitMessageHandler` (appointment tests now assert nothing is sent), `RopDispatcher`,
  `Round4Review` tests; Mongo integration tests for `ecBufferTooSmall`, ownership-before-lock, in-progress marker,
  lock/marker renewal (spied `setInterval`), session vanishing under the lock, and released write-stream deletion.
- Lesson: the Bash tool also rejects some heredocs containing apostrophes inside a quoted `'PYEOF'` block; the Write tool
  plus a script file is reliable. Some test files flip to CRLF on disk after edits, so string-replace scripts should
  normalize `\r\n` first.

### 2026-09-14 (6) — `@rapidrest/service-core` 2.1.0 migration

Not committed; package version unchanged. `@rapidmx/restapi` and the other `@rapidmx` deps left at `^0.9.0`.

- **Deps.** service-core devDependency `^2.0.0` -> `^2.1.0`, peerDependency `2.x` -> `^2.1.0`; `yarn install`
  installed 2.1.0.
- **Suite before any code change: 680/680 passing** (57 files), coverage thresholds met. None of the 2.1.0 breaking
  changes (allowExistingACL on creates, stricter dates, 400 for `$or`/`$and`/`$` keys, plain-doc locking, duplicate
  key 400/409, recordACL truncate cap, Redis error handlers, auth fallthrough, per-IP anonymous `@RateLimit`, SQL
  `insert()`) broke a test. restapi 0.9.0's code paths this plugin reaches (`findOrCreateWellKnownFolder`, which
  creates the folder ACL at the new folder's own uid, Sent Items/Outbox `messageRepo.create`) did not hit
  `IDENTIFIER_EXISTS`; no restapi blocker for mapi.
- **Simplification.** `RestapiRules.literalQueryValue()` now returns `ModelUtils.literal(value)` instead of the
  hand-built `eq(${value})` string (used for the client-supplied contact display name in `AddressList` and the
  meeting `icalUid` in `MeetingMessageClassHandler`). Same result for every value that could match before; the only
  difference is that `me`, `null` and numeric-looking strings are no longer substituted/coerced (previously they
  matched nothing or threw and were treated as no match), so e.g. a contact literally named `me` now resolves. The
  in-memory exact re-check and the try/catch stay. Unit tests that asserted the `eq(...)` string now assert
  `ModelUtils.literal(...)` (`MeetingMessageClassHandler.test.ts`, `RopSubmitMessageHandler.test.ts`,
  `Round5Review.test.ts`).
- `asEntity` kept (doc comment updated): 2.1.0 locks plain documents too, so it is now defence in depth.
- Final: `yarn lint` and `npx tsc --noEmit -p .` clean; `yarn vitest run --coverage` 680/680, 100%
  statements/functions/lines, 98.95% branches.

### 2026-09-14 (7) — Round-6 review fixes (quota accounting, RopBufferTooSmall reserve, invites, markers)

All 4 findings were checked against the code and were real; none skipped. Not committed; no version or peerDependency
changes. 707 tests pass (was 680); coverage 100% statements/functions/lines, 98.97% branches; `yarn lint` and
`npx tsc --noEmit -p .` clean.

- **Quota accounting is O(1) (1, HIGH).** The chunk script walked every chunk (HKEYS + HSTRLEN) and every owner member
  (EXISTS) on each write, inside Redis's thread. Now (`HandleDataCache.ts`):
  - A write stream's hash keeps a `size` field (end of its furthest chunk). The stream is charged that size, so a retried
    write at the same offset still replaces its chunk and isn't charged twice. `MAX_CHUNKS_PER_STREAM = 8192` (HLEN
    check; a rewrite of an existing offset is still accepted). Chunks are still unpadded base64url, no longer needed for
    sizing.
  - Each owner index keeps a running total in `<index>.total` (always the sum of the zset's scores). `admits` checks it
    and only when it would refuse does a full recount (prunes members whose data key is gone), rate limited by
    `SET <index>.recount NX PX OWNER_RECOUNT_INTERVAL_MS (1000)`. An index without a total (pre-round-6 data, expiry)
    is recounted first. Totals only over-count.
  - `HandleDataStore.delete(key, owners?)`: with owners, `DELETE_WITH_OWNERS_SCRIPT` subtracts the key right away. The
    route's released-handle cleanup and Submit's chunk delete pass `handleDataOwners`. `deleteOwnedBy` also deletes the
    index's total and marker; the user index isn't touched (healed by the next recount).
  - Script ARGV now starts with prefix, member, index ttl, recount ms (`ownerScriptArguments`). The memory store mirrors
    it (per-stream size/offsets, per-index total + `recountAfter`; a stream counts as live while its first chunk is).
  - `test/fakeRedis.ts` emulates the new scripts, exposes `recounts`, and honors the recount marker with `Date.now()`.
- **RopBufferTooSmall reserve (2).** `dispatchRops` keeps room for a RopBufferTooSmall covering the next ROP and the rest
  of the request (`3 + remaining request bytes`). `context.ropOutputRemaining` is now a getter: room left minus the reserve
  from the reader's current position (so after a handler reads its request it excludes its own bytes). A response that
  doesn't leave the reserve becomes RopBufferTooSmall from that ROP, which always fits. `ExecuteBufferTooSmallError` is
  thrown only before any ROP runs (MaxRopOut can't hold a RopBufferTooSmall of the whole request), so ecBufferTooSmall
  never follows saved state and nothing is stored for its X-RequestId. This reinstates a pre-run check that round 5
  rejected: an Execute whose request is bigger than its MaxRopOut room now fails even if its responses would have fit
  (real clients send MaxRopOut ~0x10008 with requests <= 32767, so this doesn't arise in practice). ReadStream/QueryRows/
  GetBuffer size to the getter, so they are never replaced after moving a cursor.
- **Invites only for submitted revisions (3).** `RopSaveChangesMessage` creates meetings with `inviteSequenceSent: 0` and
  stamps an update with the new `sequence`, unless a submitted revision is still waiting for the job
  (`isInvitePending(existing)`, in `RestapiRules.ts`, mirrors the job's check), in which case the existing value (or
  null) is kept so the job still sends, now the edited meeting. Deliberate: dropping a submitted-but-unsent invite was
  judged worse than sending the edit. `submitAppointment` (organizer copy only): pending -> nothing; else versioned
  update `{uid, version, inviteSequenceSent: null}` (`asEntity`). The job only claims pending rows, so a version
  conflict is another edit (attendee reply, REST): re-read and retry, `MAX_INVITE_REQUEST_ATTEMPTS = 3`, then throw
  (MAPI_E_CALL_FAILED); a row gone on re-read -> `MAPI_E_NOT_FOUND`. Re-submitting a revision the job already sent
  sends it again. Integration tests run restapi's real `MeetingSchedulingJobMongo`/`SQL` (`objectFactory.newInstance`)
  against save-only, repeated submit, save-then-submit and edit-save-without-send; null writes work on both backends.
- **In-progress markers (4).** The renewal timer chains renewals (`renewing` promise, awaited in `finally` before
  `clearInProgress`), stops the interval when `renewLock` returns false (errors keep it going), and calls the new
  `MapiSessionManager.renewInProgress`, which extends `.last` only while it still equals this request's marker (reuses
  `renewLock`'s compare-and-PEXPIRE with the marker JSON as token). `clearInProgress` is now atomic via `releaseLock`
  (compare-and-delete). `markInProgress` stays unconditional (it replaces the previous request's stored response).
- Tests: new `test/Round6Review.test.ts`; updated `RopDispatcher.test.ts` (reserve-adjusted room), `Round5Review.test.ts`
  (deletes pass owners, new ARGV order), Mongo route tests (renewal sequencing, lost-lock stop, pre-run ecBufferTooSmall
  with nothing stored, ReadStream -> RopBufferTooSmall replayed from the store, meeting invite flow) and a SQL invite flow.
- Lesson: mongo route test's renewal test needs a real `setTimeout` pause after the fake ticks, since renewals are now
  sequential and the request can finish before the second one runs.

### 2026-09-22 — Round-7 review fix (RopDeleteMessages budget) + targeted coverage pass

One finding, confirmed real; not a re-litigation of anything above. Not committed as a version bump - `RELEASE_NOTES.md`
flipped back to `## Unreleased`. 725 tests pass (was 707 before this session's own two new test files plus additions to
five existing ones); coverage 100% statements/functions/lines, 99.34% branches (was 98.97%); `yarn lint` and
`npx tsc --noEmit -p .` clean.

- **MEDIUM: `RopDeleteMessagesHandler` had zero `ExecuteBudget` accounting**, the only bulk-mutation ROP round 5's
  work-budget pass missed - `RopDeleteFolderHandler.deleteItems()` and every paged query already charge
  `context.budget?.chargeQueries()` per DB round trip, but the per-`MessageId` loop here (`findOne` + `delete`, or a
  bare `delete` for a `calendarEvent:` target) never did, so a client could still walk ~4000 sequential DB round
  trips per `Execute` (a 32767-byte ROP buffer holds that many 8-byte `MessageId`s) against their own mailbox with no
  budget stopping it. Fixed by charging one query before each `findOne`/`delete` call, matching
  `resolveContentsKind`'s own single-row-lookup charge and `RopDeleteFolderHandler.deleteItems()`'s own per-delete
  charge exactly (a `WorkBudgetExceededError` thrown mid-loop propagates out of `handle()` uncaught, same as
  `RopDeleteFolderHandler`'s own direct-call budget test, and `dispatchRops` turns it into `MAPI_E_TOO_COMPLEX` for
  that ROP when reached through `Execute`). New test mirrors `Round5Review.test.ts`'s "the second delete never ran"
  case: a 3-query budget spends 2 on the first message (findOne+delete) and the third on the second message's findOne,
  leaving nothing for its delete, which throws `WorkBudgetExceededError`.
- **Coverage, all targeted gaps from an adversarial review pass, not a general sweep**:
  - `BaseMapiEmsmdbRoute.ts`'s `executeLocked()` rethrows whatever `dispatchRops` rejects with when it isn't a
    `DecodeError`/`ExecuteBufferTooSmallError` - untested because `dispatchRops` itself converts every ordinary
    handler-thrown error into a per-ROP failure response internally (see `RopDispatcher.test.ts`'s own handler-throws
    coverage) and never lets it escape. The one real way it *can* escape: `failureResponse()` (called from inside
    `dispatchRops`'s per-ROP catch, itself inside an outer try with only a `finally`, no enclosing catch) throws while
    building that failure response - e.g. a `RopHandler` whose `failureTailBytes` getter throws. New real-HTTP test in
    `test/routes/mongo/MapiEmsmdbRoute.test.ts` pokes a broken handler with exactly that shape into the route's own
    `ropHandlers` map (private is compile-time only) and confirms `Execute` answers a generic 500
    (`serializeError`'s non-`ApiError` fallback), not a raw leak or a crash.
  - The `this.auditLogClass ? ... : undefined` ternary's false side (no audit log class configured) has no real
    route to exercise it, since both concrete routes (`MapiEmsmdbRouteMongo`/`SQL`) always set one. New
    `test/server-mongo/routes/MapiEmsmdbRouteNoAudit.ts` mounts a second, otherwise-identical Mongo route at
    `/mongo/mapi/emsmdb-noaudit` with `auditLogClass` left unset, and a new test does a bare Connect+Execute against
    it (any Execute exercises the ternary, regardless of which ROP runs).
  - `TransportSend.ts`'s `sendOrThrow` had no dedicated test file at all despite being its own module (only exercised
    indirectly through `MeetingMessageClassHandler.test.ts`). New `test/rop/TransportSend.test.ts` covers all of it:
    no result, empty `accepted`, non-empty `rejected`, success, and the `?? []` fallback on each field when it's
    absent from the result object entirely (not just empty).
  - `RopOpenMessageHandler.ts`: the `context.taskRepo ? ... : ""` branch's false side (task-repo absent from the
    context) had no test, unlike the identical pattern one line up for `contactRepo` - added the mirrored test.
  - `RopWriteStreamHandler.ts`'s exported `readWriteStream()`: `stream.writeSize ?? 0` had no test for a handle
    nothing has ever been written to (`writeSize` genuinely `undefined`, not just `0`) - added a test asserting it
    reads back an empty buffer rather than needing a prior write.
  - Left alone (out of this pass's scope, not newly discovered): `RopDispatcher.ts` line 58's other
    `decodingReader` branch, `PropertyResolvers.ts` lines 395-397, and `BaseMapiEmsmdbRoute.ts` lines 90/301/362/407/520
    - pre-existing single-line branch gaps this review didn't call out specifically; the branch floor (95%) was
    already comfortably cleared before and after this session regardless.

#### Follow-up (same day) — a second bulk-mutation path missing budget accounting: `AddressList.ts`

A second-round review of the commit above found one more real instance of the same species of gap it fixed:
`src/rop/AddressList.ts`'s `findContactsByDisplayName`/`resolveRecipientList` had zero `ExecuteBudget` accounting -
unlike `FolderTarget.ts`/`MessageTarget.ts`/`CalendarEventTarget.ts`, `budget` wasn't even part of their context
type. `RopSubmitMessageHandler.ts` calls `resolveRecipientList` once each for To/Cc/Bcc (lines ~164-167); every
bare-display-name entry (no `@`, entirely client-controlled via a crafted draft's `PidTagDisplayTo`/`Cc`/`Bcc`) runs
a real `contactRepo.find()` through `findContactsByDisplayName`, gated only by the `MAX_RECIPIENTS_PER_MESSAGE = 500`
count cap - which bounds list length, not DB work. Combined with `MAX_SUBMITS_PER_EXECUTE = 16`, one `Execute` could
run up to 8,000 completely unmetered `contactRepo.find()` calls without ever touching `queriesRemaining`.

Fixed by widening both functions' context parameter from `Pick<RopContext, "mailboxUid" | "contactRepo">` to also
include `"budget"`, and charging `context.budget?.chargeQueries()` in `findContactsByDisplayName` immediately before
the real `contactRepo.find()` call (after the existing `!context.contactRepo` early-return, so a mailbox with no
contacts repo configured still charges nothing - consistent with `RopDeleteFolderHandler.itemRepos()`'s own
optional-repo skip). No call-site changes needed: `RopSubmitMessageHandler.ts`'s three call sites already pass the
full `RopContext`, which structurally satisfies the widened `Pick`.

New test in `test/rop/RopSubmitMessageHandler.test.ts` ("Round 5: once per draft" group, beside the existing
`MAX_RECIPIENTS_PER_MESSAGE`/submit-budget tests): a `PidTagDisplayTo` of two bare display names with a
`maxQueries: 1` budget throws `WorkBudgetExceededError` after the first name's lookup, leaving `contactFind` called
exactly once - the same shape as `RopDeleteMessagesHandler`'s own budget-exhaustion test from the commit above.

726 tests pass (was 725); coverage 100% statements/functions/lines, 99.34% branches (unchanged - the new charge
line is exercised by both the new test and every existing bare-display-name-resolution test). `yarn lint`,
`npx tsc --noEmit -p .` and `yarn build` all clean. Not a version bump - `RELEASE_NOTES.md` was already back at
`## Unreleased` from the commit above.

### 2026-09-23 — Round-3 review fixes (audit-write budget accounting, PropertyType-mismatch degradation)

Two findings, both confirmed real before fixing; two others (`Disconnect`'s missing session lock, a `.last`
Redis-key cleanup gap on session destroy) were reviewed and explicitly skipped as too low severity/self-inflicted-
only per the coordinator's own triage. Not a version bump - `RELEASE_NOTES.md` stayed at `## Unreleased`. 734 tests
pass (was 726); coverage 100% statements/functions/lines, 99.34% branches; `yarn lint`, `npx tsc --noEmit -p .` and
`yarn build` all clean.

- **MEDIUM: the audit write inside `RopDeleteMessagesHandler`/`RopDeleteFolderHandler`'s delete loops was the one
  real DB round trip in each loop iteration left uncharged** - the same species of gap as the two fixed in the
  commits above, just on `auditMessageDelete()`'s own `context.audit?.()` call (a real `create()` once
  `auditLogClass` is configured, which every real deployment does) rather than the `findOne`/`delete` calls beside
  it. `RopDeleteMessages` therefore cost 3 real DB round trips per message (`findOne`/`delete`/audit-`create`) but
  only charged 2; `RopDeleteFolder` cost 2 per message (`delete`/audit-`create`) but only charged 1 - up to a 100%
  undercount of real DB load relative to what `ExecuteBudget`'s own doc comment promises to cap.
  - Fixed centrally rather than at each of the two call sites: `auditMessageDelete()` (`RopDeleteMessagesHandler.ts`,
    shared by both handlers - see `RopDeleteFolderHandler.ts`'s own import) now charges
    `context.budget?.chargeQueries()` itself, immediately before its own `context.audit(...)` call, and returns
    early (no charge - there's no DB work to account for) when `context.audit` is absent. This was a deliberate
    change from "add a charge at each call site" (what the finding's own wording suggested): the call is only
    *sometimes* a real DB round trip (whether `context.audit` is configured), unlike the `delete()`/`findOne()`
    calls beside it which always are, so an unconditional charge at each call site would have over-charged (and
    broken) every existing no-audit-configured test that asserts an exact `queriesRemaining` count. Charging inside
    the function itself is also the only single point that both current call sites (and any future one) get right
    automatically, rather than needing the same `if (context.audit)` guard duplicated at each site.
  - New tests: `test/rop/RopDeleteMessagesHandler.test.ts` ("Charges the audit write too...") - with a 3-query
    budget and `audit` configured, the first message's `findOne`+`delete`+audit-`create` now spend all 3 charges
    (not 2), so the *second* message's `findOne` is what throws, one message earlier than the sibling test right
    above it with the same budget and no audit function. `test/Round5Review.test.ts` ("Charges RopDeleteFolder's
    audit write too...") mirrors the existing "Charges RopDeleteFolder's walk and each delete..." test: with audit
    configured the same scenario now spends 6 queries, not 5, and a 3-query budget lets `m2`'s `delete` run but
    throws on its audit charge specifically (`messageRepo.delete` called twice, `audit` still only once) - proving
    the audit charge, not just the delete, can be what a tight budget runs out on.
- **LOW: a client-supplied `PropertyType` that doesn't match what a `propertyId`'s resolver actually returns failed
  the *whole* `RopQueryRows`/`RopGetPropertiesSpecific` response with `MAPI_E_CALL_FAILED`**, discarding every row
  already built, instead of degrading just that one column - breaking both handlers' own documented "an unsupported
  property falls back to a type-appropriate default... never an error" contract, which was only ever actually true
  for a `propertyId` with no resolver data, not for a `propertyId` that resolves fine but to a value shape the
  *client's own requested* `propertyType` for that column doesn't accept. Root cause: every `xxxValueFor` function
  in `PropertyResolvers.ts` switches purely on `propertyId` and has no idea what `propertyType` the client actually
  asked for that column to be; the mismatch only surfaces later, in `PropertyValue.ts`'s `writePropertyValue()`
  (`encodeGuid()` throwing on a plain string, for example), by which point `RopQueryRowsHandler`'s per-row loop and
  `RopGetPropertiesSpecificHandler` have no try/catch around it, so the throw reaches `RopDispatcher` and fails the
  whole ROP. Repro (as given): `SetColumns([{propertyId: 0x0037 (PidTagSubject), propertyType: 0x0048 (PtypGuid)}])`
  then `QueryRows` against any mailbox with >= 1 message.
  - Checked first, per the finding's own suggestion: `FlaggedPropertyRow`'s per-column error-flag mechanism isn't
    implemented anywhere in this codebase (`RopQueryRowsHandler`'s own doc comment explicitly says so, and
    confirmed via grep) - building it from scratch (per-column flag byte + `PropertyErrorCode` encoding, spec
    compliance) was judged out of scope for a LOW-severity fix when the existing `defaultValueForType()` fallback
    (already used for a wholly-unmodeled `propertyId`) is right there and already type-correct per `PropertyType`.
  - Fix: new `PropertyResolvers.writePropertyValueSafely(writer, propertyType, value)` - encodes into a **scratch**
    `BufferWriter` first, and only copies its bytes into the real `writer` once the whole value encoded cleanly;
    falls back to encoding `defaultValueForType(propertyType)` (into its own fresh scratch writer) if the first
    attempt throws. The scratch-writer indirection matters: `PtypBinary`/every `PtypMultiple*` case writes more than
    one field per value (a count, then elements), and a wrong-*shaped* (not just wrong-type) value can throw
    partway through one of those - writing directly into the real `writer` would leave stray, wrongly-sized bytes
    ahead of every column/row that follows, corrupting the wire format worse than the original failure. Still
    throws (unchanged from today) when even the fallback fails, which only happens when `propertyType` itself is a
    raw value neither `switch` recognizes at all (`RopSetColumns` never validates `propertyType` against the known
    enum) - there's no byte encoding this codec knows how to produce for a wire type it doesn't implement, so that
    narrower, much rarer case is deliberately left as today's existing whole-ROP failure rather than invented on
    the spot.
  - `RopQueryRowsHandler.buildRow()` and `RopGetPropertiesSpecificHandler.handle()` both now call
    `writePropertyValueSafely` instead of `writePropertyValue` directly; both class doc comments updated to state
    the corrected, now-actually-true scope of their "never an error" claim.
  - New tests: mirrored repro tests in `test/rop/RopQueryRowsHandler.test.ts` and
    `test/rop/RopGetPropertiesSpecificHandler.test.ts` (PidTagSubject's string value requested as `PtypGuid`,
    asserting `ReturnValue` 0 and the row/property still comes back, decoding to `defaultValueForType`'s GUID) plus
    a direct `test/rop/PropertyResolvers.test.ts` unit-test block for `writePropertyValueSafely` itself: the
    as-is-success path, the fallback path, a sentinel-byte test proving no stray partial-write bytes leak ahead of
    the fallback's own bytes, and confirmation it still throws for a `propertyType` neither `switch` recognizes.
- Skipped, per explicit instruction (not re-investigated further here): a missing session lock on `Disconnect`
  (only lets a user race their own in-flight request; confirmed no cross-user impact) and a missing `.last`
  Redis-key cleanup on session destroy (pure TTL-bounded hygiene, self-heals within `SESSION_LOCK_TTL_MS`).

### 2026-09-25 - release bump levels follow upstream

When releasing packages that depend on each other (rapidmx: restapi / react-shared -> web-client -> meet-plugin, booking-plugin, autodiscover, mapi, activesync, server; rapidrest: core / service-core -> auth / auth-server / react / cli and the projects built on them), the bump level of a downstream release matches the level of the upstream release it picks up: an upstream **minor** is a downstream **minor**, an upstream patch a downstream patch, major to major. Where a downstream bump crosses several upstream releases, use the highest level among them, and never choose "patch" just because the downstream's own diff is only a `package.json` bump. Betas keep their prerelease line but follow the same idea - say which level was chosen.

Why: meet-plugin 0.4.2 and booking-plugin 0.5.2 were cut as patches after web-client 0.15.x -> 0.16.0 and react-shared 0.17.0 -> 0.18.0 (both minors), and autodiscover 1.1.1 after restapi 0.20.1 -> 0.21.0; the downstream versions then hid additive behaviour. JP accepted those releases as they were (2026-09-25) and asked for the rule going forward. Releases only happen when JP asks for them.
