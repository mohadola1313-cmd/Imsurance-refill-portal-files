# P0661 Reliability Layer — Test and Setup Guide

## Scope

All instructions below apply only to the approved test Apps Script project and test spreadsheet:

- Branch: P0661
- Spreadsheet ID: 1HTdkNkADKizQxycfDx-jgnWEsqaHVxJ1XkrkGlB3KxY
- Firestore project: insurance-refill-portal
- Feature branch: optimize-save-notes-status

Do not run these steps against production. Do not merge or deploy without separate approval.

## Source files to synchronize

Add the new firestore_reliability.gs file to the Apps Script test project alongside the updated code.gs and firestore.gs.

## Safe checks

1. Run firestoreReliabilitySelfTestP0661_() from the Apps Script editor.
2. Confirm the result reports passed: 7.
3. Review the execution log for syntax/runtime errors.

This self-test only exercises deterministic helpers. It does not read or write Sheets or Firestore.

## Controlled integration test sequence

Use disposable test transactions only, and obtain explicit authorization before running any test that writes to the spreadsheet or Firestore.

1. **Queue creation:** save one test status or modified due date. Confirm the _Firestore Sync Queue sheet receives one transaction job.
2. **Coalescing:** update the same transaction again before its job completes. Confirm it reuses the pending job rather than creating a duplicate.
3. **Normal sync:** confirm Firestore receives the latest line and transaction aggregate, then the job becomes complete.
4. **Retry:** simulate a Firestore failure in a controlled test environment. Confirm the canonical Sheets change remains saved, the queue row returns to pending, the attempt count increases, and Last Error is recorded. Restore Firestore access and confirm a later retry completes.
5. **Deletion:** delete a test transaction and confirm its purchase-line documents and transaction document are removed. Also test deleting one item line while other lines in that transaction remain.
6. **Snapshot cold/warm:** request a due month twice. Confirm the first request builds the monthly snapshot and the second can use the cache.
7. **Cache recovery:** remove/expire one cache chunk in a controlled test and confirm the reader rebuilds the complete snapshot instead of returning partial rows.
8. **Pending-sync safety:** while a queue job is pending, confirm Explore reads canonical Sheets rather than stale Firestore/cache data.
9. **Concurrency:** overlap an Explore request with a test update and confirm it returns a consistent pre-update or post-update view, never a mixed state.

## Trigger installation

The retry trigger is deliberately not installed automatically.

After the index is ready and the controlled tests pass, manually run:

firestoreInstallRetryTriggerP0661_()

It installs a five-minute trigger for firestoreRetryWorkerP0661_(). The installer refuses to run unless FIRESTORE_INDEX_P0661_READY equals 1.

To run one controlled retry batch manually, use firestoreRetryWorkerP0661_().

## Initial index rebuild

firestoreInitialSyncP0661() writes Firestore documents. It is not part of the read-only checks. It first disables the Firestore Explore read path; if the rebuild fails, Explore remains on the Sheets fallback. Run it only with explicit authorization and after verifying the test spreadsheet/project.

## Performance measurement

Record at least five cold and five warm Explore runs for the same month. Compare median duration, Firestore document reads, query pages, and cache hit/miss behavior. Also record the fallback path while sync is pending. Do not claim a speed improvement until measurements are repeatable.
