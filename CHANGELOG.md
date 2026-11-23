# Change Log

## 2026-11-23 — Firestore sync reliability review

### Findings (current implementation on `optimize-save-notes-status`)
- Google Sheets Follow Up is written before the corresponding Firestore synchronization is attempted.
- Firestore commit failures are surfaced as errors by `firestoreCommit_()`; no sync-pending record or automatic retry/reconciliation mechanism was found in the reviewed code.
- Explore and Member History use Firestore when `FIRESTORE_INDEX_P0661_READY = 1`. A failed write can therefore leave the read model stale while the canonical Follow Up data already contains the newer change.
- Incremental synchronization of affected purchase lines/transactions is already present and should be preserved.

### Risk
A Follow Up update can succeed while its Firestore projection fails. The user’s change remains in the canonical Google Sheets data, but Explore/Member History may continue showing the older Firestore state until a successful synchronization occurs.

### Proposed improvement — Sync Queue / Pending Sync (not implemented)
1. Persist a durable pending-sync record after the canonical Follow Up write.
2. Give each operation a deterministic idempotency key so retries cannot create duplicate effects.
3. Track operation state, attempt count, timestamps, and the last error.
4. Attempt the incremental Firestore sync for affected transactions only.
5. Mark the queue item complete only after Firestore confirms success; retain failures for scheduled retry.
6. Protect against stale writes so an older retry cannot overwrite a newer user change.
7. Add a reconciliation path to detect and repair pending or inconsistent projections.

### Design constraints
- Google Sheets Branch remains the canonical source for sales data.
- Google Sheets Follow Up remains the canonical source for user-entered follow-up state.
- Firestore remains a read-optimized projection/index, not the primary source of truth.
- Preserve incremental updates; do not rescan or rewrite unaffected transactions.
- Do not claim performance improvement without runtime measurements.

### Status
**Documented for design; no application code, Sheets data, Firestore data, deployment, or main branch was changed as part of this review.**

## 2026-11-23 — Reliability and monthly Explore snapshot implementation

- Added `firestore_reliability.gs` for a durable P0661 outbox queue, transaction-level retry/reconciliation, processing leases, exponential backoff, and a manual trigger installer.
- Routed P0661 upload, delete, status, modified due date, group due date, and follow-up writes through the outbox before canonical writes; bounded immediate processing runs under ScriptLock.
- Added transaction reconciliation that removes stale purchase-line documents and deletes empty transaction documents.
- Replaced the P0661 Explore fast path with due-month Firestore queries and versioned, chunked CacheService snapshots. Pending sync or concurrent revision changes force a locked Sheets fallback.
- Added read-only helper self-tests. They have not yet been executed in Apps Script.
- No trigger installation, live Sheets/Firestore writes, initial index build, deployment, or merge was performed. Integration testing and performance measurements remain pending.
