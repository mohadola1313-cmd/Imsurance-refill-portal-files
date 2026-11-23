# Data Reliability and Snapshot Design

## Scope and safety boundary

- Repository: `mohadola1313-cmd/Imsurance-refill-portal-files`
- Work branch: `optimize-save-notes-status`
- Spreadsheet: approved test copy only (`1HTdkNkADKizQxycfDx-jgnWEsqaHVxJ1XkrkGlB3KxY`)
- Firestore target: project `insurance-refill-portal`, database `(default)`, branch `P0661`, region `me-central2`
- Google Sheets remains canonical. Firestore remains a read-optimized projection.
- No production/main branch, live deployment, or data mutation is authorized by this design.

## Current implementation observations

1. Explore currently lists the complete Firestore `purchaseLines` collection, then filters the rows in Apps Script for the selected due month/day.
2. The current Explore read builds both rows and follow-up data from one Firestore response, which avoids a second collection read but still downloads the complete collection on every uncached request.
3. Apps Script `CacheService` is currently used for short-lived sessions and rate limits. Those are appropriate ephemeral uses; it must not become the only copy of sales, follow-up state, or pending work.
4. Follow Up writes occur before Firestore sync in several user actions. A Firestore error can leave the canonical sheet updated while the Firestore read model remains stale.
5. No durable pending-sync queue or automatic retry worker exists in the reviewed implementation.
6. `setStatus`, `setModifiedDue`, and `setGroupDue` do not currently wrap the complete read-modify-write-sync operation in a script lock. This must be addressed before retry workers can safely share the same write paths.
7. `deletePeriod` currently calls the incremental sync with remaining rows, but does not explicitly provide deleted transaction/line identities. Deletion reconciliation needs a tombstone-aware path.

## Target architecture

### A. Canonical data and durable sync queue

- Keep Branch sheets as the canonical sales source.
- Keep the Customer Follow Up sheet as the canonical source for user-entered action, notes, status, and modified due date.
- Add a dedicated queue sheet in the approved test spreadsheet (for example, `_Firestore Sync Queue`). It is an operational journal, not a second business-data source.
- Queue entries identify affected branch/transaction keys (and explicit deleted line/transaction keys where needed); they should not duplicate entire purchase rows.
- Queue state should include a stable job ID, branch, transaction key/ID, operation type, created/updated timestamps, state, attempt count, next-attempt time, last error, and a revision/sequence.
- Coalesce duplicate pending jobs for the same transaction where safe. The worker must always rebuild the Firestore projection from the latest canonical Sheets/Follow Up state, not replay stale field values from an old event.
- Persist the queue intent before changing canonical rows while holding the same ScriptLock. This ordering prevents a successful sheet write from being left without a queue record if execution stops between the two operations. If the business write subsequently fails, the harmless queued job can reconcile the current canonical state.
- Ensure every relevant mutation path uses the same lock. The worker must acquire that lock too, so it cannot read halfway through a write.
- A scheduled worker processes a bounded number of pending jobs per run. It marks a job complete only after Firestore confirms success. On failure it retains the job, increments attempts, records a bounded error message, and schedules exponential backoff with a maximum delay.
- Recover jobs left in a processing state after an execution timeout by using a lease/processing timestamp.
- Use deterministic Firestore document IDs and transaction-level coalescing for idempotency. A retry must converge to the latest canonical state.
- Prevent stale overwrites by serializing mutations and worker execution with ScriptLock and rebuilding from current canonical data immediately before each commit. If later we permit out-of-band writers, add explicit per-transaction revision checks.
- Do not mark the index globally ready if unresolved initial-build or reconciliation work remains.

### B. Deletion/tombstone handling

- Deletion jobs must retain the identities of the removed purchase lines and affected transactions.
- For each affected transaction, compare current canonical lines with indexed lines; delete only stale line documents and rebuild the transaction aggregate from the remaining canonical lines.
- If a transaction has no remaining lines, delete its transaction document too.
- Never infer deleted identities from the post-delete remaining-row list alone.

### C. Snapshot and cache strategy

- Treat CacheService as a disposable performance layer only. It can expire or evict entries and must never be the authoritative copy.
- Do not put the complete branch snapshot into one cache key: the per-item cache size limit makes that unsafe as the dataset grows.
- Build versioned, bounded snapshot chunks (or month-scoped chunks) with a manifest containing schema version, data revision, chunk count, and build timestamp. Keep each cache value comfortably below the documented per-key size ceiling.
- Cache keys must include branch, scope (such as due month), schema version, and data revision. A successful Firestore projection update advances the durable revision; old cache keys then become unreachable without needing to delete every old entry.
- On a cache miss, missing chunk, malformed manifest, or version mismatch, discard the partial cache result and rebuild from the Firestore read model. Never return a partial snapshot as complete.
- If a sync job is pending for the requested scope, do not serve a potentially stale cached/Firestore snapshot as current. Temporarily fall back to canonical Sheets for that request (or apply a verified overlay); the first implementation should prefer the safe Sheets fallback.
- After a successful sync/reconciliation, advance the revision and allow a fresh snapshot to be built.
- Use a short TTL for warm cache entries. Cache expiry is expected and must affect speed only, never correctness.
- Measure cold-cache and warm-cache runtime, Firestore document reads, Sheets reads, and payload sizes before claiming an improvement.

### D. Snapshot data scope

- Explore needs the selected due month for calendar counts and the selected due day for its rows. The snapshot should therefore be scoped to the due month rather than serializing every branch row for every request.
- The selected day’s row details and month calendar counts must be derived from the same snapshot revision.
- Preserve current Explore semantics: completed transactions remain visible on the selected day; calendar counts include incomplete customers only; modified due date takes precedence over calculated due date.
- Member History is out of scope for the first implementation unless separately approved. Do not change its current read path in this phase.

## Implementation sequence

1. Add focused tests/diagnostics for key construction, queue coalescing, retry state transitions, cache chunk integrity, and deletion tombstones. Tests must be read-only or use in-memory fixtures; no temporary spreadsheet creation/deletion.
2. Add the durable queue sheet helpers and lock coverage for Follow Up mutation paths.
3. Add a bounded retry worker and explicit error/recovery behavior. Keep it test-branch-only.
4. Extend the sync path to handle deleted line/transaction documents using tombstones.
5. Add a versioned monthly snapshot cache with all-or-nothing chunk validation and safe fallback when pending jobs exist.
6. Validate behavior with controlled test fixtures and read-only diagnostics. Do not run initial index rebuilds, mutate Sheets/Firestore, or deploy without explicit authorization.
7. Record baseline and post-change runtime/read counts from an explicitly authorized test run before making performance claims.

## Acceptance criteria

- A successful canonical change is never left without a durable pending-sync intent.
- A failed Firestore commit leaves the queue job pending and retryable; it does not ask the user to repeat the business action.
- A retry converges to the latest canonical state without duplicate documents or stale overwrites.
- Deleted purchase lines and empty transactions are removed from Firestore correctly.
- A missing/expired/corrupt cache chunk causes a complete rebuild, never a partial response.
- Pending synchronization prevents stale cached data from being presented as current.
- Explore output remains semantically identical to the current behavior.
- Member History, production data, main branch, and deployments remain untouched in this phase.
- Performance improvement is claimed only after repeatable measurements.


## Implementation progress on the feature branch

The following source changes have been committed to `optimize-save-notes-status`:

- Added `firestore_reliability.gs` with a durable queue sheet named `_Firestore Sync Queue`, transaction-level job coalescing, retry attempts, exponential backoff, processing leases, and a manually installable five-minute worker.
- Updated P0661 upload, delete, status, modified-due, group-due, and follow-up mutation paths to enqueue affected transaction keys before canonical writes, then attempt immediate bounded processing while holding the existing/new ScriptLock.
- Retry reconciliation rebuilds each affected transaction from current Branch + Follow Up data, upserts current line documents, removes stale line documents, and deletes the transaction document when no canonical lines remain.
- Explore now requests a due-month snapshot. The new reader queries only that month by `effectiveDueDate`, caches bounded chunks with a manifest, validates all chunks before use, and uses canonical Sheets when pending sync or a concurrent revision change is detected.
- The initial index build advances the snapshot revision so older cache keys cannot be reused.
- Added a pure read-only helper self-test function: `firestoreReliabilitySelfTestP0661_()`.

### Not yet validated in a live Apps Script runtime

- The self-test function has been added but has not been executed in the Apps Script editor.
- The Firestore month query, cursor pagination, cache limits, retry trigger, and queue recovery have not yet been integration-tested against the test project's live runtime.
- The retry trigger has NOT been installed.
- No queue sheet has been created in Google Sheets by this code change.
- No Firestore or Google Sheets data has been written, and no deployment or merge has occurred.

### Required next validation

1. Add/synchronize the new `firestore_reliability.gs` file with the P0661 test Apps Script project alongside the updated `code.gs` and `firestore.gs`.
2. Run `firestoreReliabilitySelfTestP0661_()` first; it is designed to be read-only.
3. Review Apps Script execution logs and perform static review of the changed functions.
4. With explicit test-run authorization, run controlled scenarios on disposable/test records: successful sync, simulated Firestore failure and retry, repeat update/coalescing, deletion of a whole transaction, deletion of one line within a transaction, cold/warm monthly Explore, cache expiry, and concurrent mutation/read.
5. Only after those checks, manually install `firestoreInstallRetryTriggerP0661_()` and verify trigger execution.
6. Record actual cold/warm timings, Firestore document counts, and Sheets reads before making any performance claim.

Do not run `firestoreInitialSyncP0661()`, install the trigger, or deploy until the user explicitly authorizes the relevant test operation.


### Isolated verification completed

- JavaScript syntax parsing passed for `code.gs`, `firestore.gs`, and `firestore_reliability.gs`.
- The pure helper self-test passed all 7 assertions.
- An in-memory queue simulation verified pending-job coalescing, a simulated Firestore failure returning the job to pending with attempt count 1, and a later retry marking it complete.
- An in-memory snapshot simulation verified cold read from Firestore, warm read from cache, and full rebuild after a cache chunk is missing.
- An in-memory pending-sync simulation verified Explore selects the Sheets fallback without querying Firestore.
- These are isolated/mock checks only. They do not establish live Apps Script, Firestore REST query, trigger, authorization, or real performance behavior.
