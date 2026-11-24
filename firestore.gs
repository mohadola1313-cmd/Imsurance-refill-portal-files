/**
 * Insurance Refill Portal — Firestore operational index
 * Target: TEST COPY ONLY (insurance speed test), branch P0661.
 * Google Sheets remains the canonical source in this phase.
 *
 * Add this file to the TEST Apps Script project alongside Code.gs.
 * The paired Code.gs must use TEST_CLOUD_ID below in CONFIG.CLOUD_ID.
 * No service-account key is used; requests use the Apps Script OAuth token.
 */

const FIRESTORE_INDEX_CONFIG = Object.freeze({
  PROJECT_ID: 'insurance-refill-portal',
  DATABASE_ID: '(default)',
  CLOUD_ID: '1HTdkNkADKizQxycfDx-jgnWEsqaHVxJ1XkrkGlB3KxY',
  BRANCH: 'P0661',
  ROOT: 'branches/P0661',
  TX_COLLECTION: 'branches/P0661/transactions',
  LINE_COLLECTION: 'branches/P0661/purchaseLines',
  SCHEMA_VERSION: 1,
  COMMIT_BATCH_SIZE: 400
});

/** One-time initial index build. Run manually from the Apps Script editor. */
function firestoreInitialSyncP0661() {
  firestoreAssertTestTarget_();
  // Disable Firestore reads during a rebuild; Explore will safely fall back to Sheets.
  PropertiesService.getScriptProperties().setProperty('FIRESTORE_INDEX_P0661_READY', '0');
  const branch = FIRESTORE_INDEX_CONFIG.BRANCH;
  const rows = readBranch_(branch); // reads canonical Sheets + current Follow Up state
  if (!rows.length) throw new Error('P0661 has no saved rows in the test spreadsheet.');

  const follow = followMap_(branch);
  const transactions = Object.create(null);
  const lines = [];
  const now = new Date().toISOString();

  rows.forEach(function (r) {
    const transactionKey = fkey_(r);
    const transactionId = firestoreStableId_('t', transactionKey);
    const purchaseLineId = firestoreLineIdForRow_(r,branch);
    const state = follow[flinekey_(r,branch)] || follow[transactionKey] || {};
    const status = s_(r[I.ST]).toLowerCase() === 'complete' ? 'complete' : 'incomplete';
    const due = s_(r[I.DUE]);
    const modifiedDueDate = s_(r[I.MOD]);
    const actionIndex = Number(state.a) || 0;
    const action = actionIndex > 0 && actionIndex <= CONFIG.ACTIONS.length ? CONFIG.ACTIONS[actionIndex - 1] : '';
    const notes = s_(state.n);
    const net = Number(r[I.NET]) || 0;
    const qty = Number(r[I.QTY]) || 0;

    if (!transactions[transactionKey]) {
      transactions[transactionKey] = {
        id: transactionId,
        branchId: branch,
        transactionId: transactionId,
        loyaltyPhone: s_(r[I.PHONE]),
        loyaltyName: s_(r[I.NAME]),
        trxDate: s_(r[I.DATE]),
        trxTime: s_(r[I.TIME]),
        receiptNumbers: [],
        membershipIds: [],
        itemCount: 0,
        totalQuantity: 0,
        totalPreNetValue: 0,
        completeCount: 0,
        incompleteCount: 0,
        status: 'incomplete',
        action: action,
        notes: notes,
        schemaVersion: FIRESTORE_INDEX_CONFIG.SCHEMA_VERSION,
        updatedAt: now
      };
    }
    const tx = transactions[transactionKey];
    const receipt = s_(r[I.RC]);
    const mid = s_(r[I.MID]);
    if (receipt && tx.receiptNumbers.indexOf(receipt) < 0) tx.receiptNumbers.push(receipt);
    if (mid && tx.membershipIds.indexOf(mid) < 0) tx.membershipIds.push(mid);
    tx.itemCount++;
    tx.totalQuantity += qty;
    tx.totalPreNetValue += net;
    if (status === 'complete') tx.completeCount++; else tx.incompleteCount++;
    // Parent-level action/notes are materialized onto each item if no item-level value exists,
    // matching the existing followMap_() inheritance behavior.
    if (!tx.action && action) tx.action = action;
    if (!tx.notes && notes) tx.notes = notes;

    lines.push({
      id: purchaseLineId,
      branchId: branch,
      transactionId: transactionId,
      purchaseLineId: purchaseLineId,
      receiptNumber: receipt,
      trxDate: s_(r[I.DATE]),
      trxTime: s_(r[I.TIME]),
      loyaltyPhone: s_(r[I.PHONE]),
      loyaltyName: s_(r[I.NAME]),
      itemCode: s_(r[I.CODE]),
      description: s_(r[I.DESC]),
      quantity: qty,
      preNetValue: net,
      netValueSum: Number(r[I.SUM]) || 0,
      averageQuantity: Number(r[I.AVG]) || 0,
      lastPurchaseDay: s_(r[I.LAST]),
      calculatedDueDate: due,
      modifiedDueDate: modifiedDueDate,
      effectiveDueDate: modifiedDueDate || due,
      membershipId: mid,
      status: status,
      action: action,
      notes: notes,
      statusUpdatedAt: s_(state.su),
      modifiedDate: s_(state.md),
      updatedAt: now,
      schemaVersion: FIRESTORE_INDEX_CONFIG.SCHEMA_VERSION
    });
  });

  const txDocs = Object.keys(transactions).map(function (k) {
    const tx = transactions[k];
    tx.status = tx.itemCount > 0 && tx.completeCount === tx.itemCount ? 'complete' : 'incomplete';
    tx.totalQuantity = firestoreRound_(tx.totalQuantity, 3);
    tx.totalPreNetValue = firestoreRound_(tx.totalPreNetValue, 2);
    return tx;
  });

  // Transaction documents contain aggregates only; all item detail is in purchaseLines.
  const writes = txDocs.map(function (x) { return firestoreWrite_(FIRESTORE_INDEX_CONFIG.TX_COLLECTION, x.id, x); })
    .concat(lines.map(function (x) { return firestoreWrite_(FIRESTORE_INDEX_CONFIG.LINE_COLLECTION, x.id, x); }));

  let committed = 0;
  for (let i = 0; i < writes.length; i += FIRESTORE_INDEX_CONFIG.COMMIT_BATCH_SIZE) {
    firestoreCommit_(writes.slice(i, i + FIRESTORE_INDEX_CONFIG.COMMIT_BATCH_SIZE));
    committed += Math.min(FIRESTORE_INDEX_CONFIG.COMMIT_BATCH_SIZE, writes.length - i);
  }

  const result = {
    mode: 'WRITE — INITIAL INDEX SYNC',
    target: 'TEST COPY ONLY',
    cloudFile: 'insurance speed test',
    branch: branch,
    transactionDocumentsWritten: txDocs.length,
    purchaseLineDocumentsWritten: lines.length,
    totalDocumentsWritten: committed,
    transactionStatus: {
      complete: txDocs.filter(function (x) { return x.status === 'complete'; }).length,
      incomplete: txDocs.filter(function (x) { return x.status !== 'complete'; }).length
    },
    note: 'Google Sheets remains canonical. Firestore documents are deterministic upserts.'
  };
  const indexProps = PropertiesService.getScriptProperties();
  indexProps.setProperty('FIRESTORE_SNAPSHOT_REV_P0661',
    String((Number(indexProps.getProperty('FIRESTORE_SNAPSHOT_REV_P0661'))||0)+1));
  indexProps.setProperty('FIRESTORE_INDEX_P0661_READY', '1');
  Logger.log(JSON.stringify(result));
  return result;
}

/** Stable branch-scoped document IDs. Branch is represented by the collection path. */
function firestoreStableId_(prefix, rawKey) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(rawKey), Utilities.Charset.UTF_8);
  return prefix + '_' + Utilities.base64EncodeWebSafe(bytes).replace(/=+$/g, '');
}

function firestoreAssertTestTarget_() {
  if (typeof CONFIG === 'undefined' || CONFIG.CLOUD_ID !== FIRESTORE_INDEX_CONFIG.CLOUD_ID) {
    throw new Error('STOPPED: CONFIG.CLOUD_ID must equal the approved test spreadsheet ID. No Firestore writes were made.');
  }
  if (FIRESTORE_INDEX_CONFIG.BRANCH !== 'P0661') {
    throw new Error('STOPPED: this implementation is restricted to branch P0661.');
  }
  const ss = cloud_();
  if (ss.getId() !== FIRESTORE_INDEX_CONFIG.CLOUD_ID) {
    throw new Error('STOPPED: opened spreadsheet does not match the approved test file.');
  }
  if (!ss.getSheetByName(FIRESTORE_INDEX_CONFIG.BRANCH)) {
    throw new Error('The P0661 sheet was not found in the approved test spreadsheet.');
  }
}

function firestoreWrite_(collectionPath, documentId, data) {
  const fields = {};
  Object.keys(data).forEach(function (key) { fields[key] = firestoreValue_(data[key]); });
  return { update: { name: firestoreDocumentName_(collectionPath, documentId), fields: fields } };
}

function firestoreValue_(value) {
  if (value === null || value === undefined) return { nullValue: null };
  if (value instanceof Date) return { timestampValue: value.toISOString() };
  if (typeof value === 'boolean') return { booleanValue: value };
  if (typeof value === 'number') {
    return Number.isInteger(value) ? { integerValue: String(value) } : { doubleValue: value };
  }
  if (Array.isArray(value)) return { arrayValue: { values: value.map(firestoreValue_) } };
  if (typeof value === 'object') {
    const fields = {};
    Object.keys(value).forEach(function (k) { fields[k] = firestoreValue_(value[k]); });
    return { mapValue: { fields: fields } };
  }
  return { stringValue: String(value) };
}

function firestoreDocumentName_(collectionPath, documentId) {
  return 'projects/' + FIRESTORE_INDEX_CONFIG.PROJECT_ID + '/databases/' +
    FIRESTORE_INDEX_CONFIG.DATABASE_ID + '/documents/' + collectionPath + '/' + documentId;
}

function firestoreCommit_(writes) {
  const url = 'https://firestore.googleapis.com/v1/projects/' + FIRESTORE_INDEX_CONFIG.PROJECT_ID +
    '/databases/' + FIRESTORE_INDEX_CONFIG.DATABASE_ID + '/documents:commit';
  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    payload: JSON.stringify({ writes: writes }),
    muteHttpExceptions: true
  });
  const code = response.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error('Firestore commit failed (' + code + '): ' + response.getContentText().slice(0, 1200));
  }
  return response.getContentText();
}

function firestoreRound_(n, places) {
  const p = Math.pow(10, places);
  return Math.round((Number(n) || 0) * p) / p;
}


/* ================= Operational reads and incremental synchronization ================= */

function firestoreListDocuments_(collectionPath) {
  const base = 'https://firestore.googleapis.com/v1/projects/' + FIRESTORE_INDEX_CONFIG.PROJECT_ID +
    '/databases/' + FIRESTORE_INDEX_CONFIG.DATABASE_ID + '/documents/' + collectionPath;
  let token = '', out = [];
  do {
    const url = base + '?pageSize=1000' + (token ? '&pageToken=' + encodeURIComponent(token) : '');
    const response = UrlFetchApp.fetch(url, {headers:{Authorization:'Bearer '+ScriptApp.getOAuthToken()},muteHttpExceptions:true});
    const code = response.getResponseCode();
    if (code < 200 || code >= 300) throw new Error('Firestore read failed ('+code+'): '+response.getContentText().slice(0,1000));
    const body = JSON.parse(response.getContentText() || '{}');
    (body.documents || []).forEach(function(doc){ out.push({id:doc.name.split('/').pop(), data:firestoreDecodeFields_(doc.fields || {})}); });
    token = body.nextPageToken || '';
  } while (token);
  return out;
}

// Query only documents whose transactionId is in the affected set.
// Firestore IN filters accept up to 30 values; chunks keep requests bounded.
function firestoreQueryByTransactionIds_(collectionPath, transactionIds) {
  const ids = (transactionIds || []).filter(String);
  if (!ids.length) return [];
  const parts = collectionPath.split('/');
  const collectionId = parts.pop();
  const parentPath = parts.join('/');
  const parent = 'projects/' + FIRESTORE_INDEX_CONFIG.PROJECT_ID + '/databases/' +
    FIRESTORE_INDEX_CONFIG.DATABASE_ID + '/documents/' + parentPath;
  const url = 'https://firestore.googleapis.com/v1/' + parent + ':runQuery';
  const out = [];
  for (let i = 0; i < ids.length; i += 30) {
    const chunk = ids.slice(i, i + 30);
    const payload = {
      structuredQuery: {
        from: [{collectionId: collectionId}],
        where: {
          fieldFilter: {
            field: {fieldPath: 'transactionId'},
            op: 'IN',
            value: {arrayValue: {values: chunk.map(function(id){return {stringValue:id};})}}
          }
        }
      }
    };
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: {Authorization: 'Bearer ' + ScriptApp.getOAuthToken()},
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    const status = response.getResponseCode();
    if (status < 200 || status >= 300) {
      throw new Error('Firestore targeted query failed (' + status + '): ' + response.getContentText().slice(0, 1200));
    }
    const results = JSON.parse(response.getContentText() || '[]');
    results.forEach(function(item) {
      const doc = item && item.document;
      if (doc) out.push({id: doc.name.split('/').pop(), data: firestoreDecodeFields_(doc.fields || {})});
    });
  }
  return out;
}

function firestoreDecodeFields_(fields) {
  const out = {};
  Object.keys(fields || {}).forEach(function(k){ out[k] = firestoreDecodeValue_(fields[k]); });
  return out;
}
function firestoreDecodeValue_(v) {
  if (!v) return null;
  if (Object.prototype.hasOwnProperty.call(v,'stringValue')) return v.stringValue;
  if (Object.prototype.hasOwnProperty.call(v,'integerValue')) return Number(v.integerValue);
  if (Object.prototype.hasOwnProperty.call(v,'doubleValue')) return Number(v.doubleValue);
  if (Object.prototype.hasOwnProperty.call(v,'booleanValue')) return !!v.booleanValue;
  if (Object.prototype.hasOwnProperty.call(v,'nullValue')) return null;
  if (v.arrayValue) return (v.arrayValue.values || []).map(firestoreDecodeValue_);
  if (v.mapValue) return firestoreDecodeFields_(v.mapValue.fields || {});
  if (v.timestampValue) return v.timestampValue;
  return null;
}

function firestoreLineIdForRow_(r, branch) {
  const key=flinekey_(r,branch);
  if(!key) throw new Error('Cannot build a Firestore purchase-line ID: Membership ID, full transaction time, branch, and item code are required.');
  return firestoreStableId_('l',key);
}
function firestoreTxIdForRow_(r) { return firestoreStableId_('t', fkey_(r)); }

// Builds canonical Firestore documents from the current Sheets state. No writes occur here.
function firestoreBuildCurrentProjection_(branch, rows, followOverride) {
  const follow = followOverride || followMap_(branch), txs = Object.create(null), lines = [], now = new Date().toISOString();
  rows.forEach(function(r){
    const txKey=fkey_(r), txId=firestoreTxIdForRow_(r), lineId=firestoreLineIdForRow_(r,branch);
    const state=follow[flinekey_(r,branch)] || follow[txKey] || {};
    const status=s_(state.status || r[I.ST]).toLowerCase()==='complete'?'complete':'incomplete';
    const actionIndex=Number(state.a)||0;
    const action=actionIndex>0&&actionIndex<=CONFIG.ACTIONS.length?CONFIG.ACTIONS[actionIndex-1]:'';
    const qty=Number(r[I.QTY])||0, net=Number(r[I.NET])||0;
    if(!txs[txKey]) txs[txKey]={id:txId,branchId:branch,transactionId:txId,loyaltyPhone:s_(r[I.PHONE]),loyaltyName:s_(r[I.NAME]),trxDate:s_(r[I.DATE]),trxTime:s_(r[I.TIME]),receiptNumbers:[],membershipIds:[],itemCount:0,totalQuantity:0,totalPreNetValue:0,completeCount:0,incompleteCount:0,status:'incomplete',action:action,notes:s_(state.n),schemaVersion:FIRESTORE_INDEX_CONFIG.SCHEMA_VERSION,updatedAt:now};
    const tx=txs[txKey], receipt=s_(r[I.RC]), mid=s_(r[I.MID]);
    if(receipt&&tx.receiptNumbers.indexOf(receipt)<0)tx.receiptNumbers.push(receipt);
    if(mid&&tx.membershipIds.indexOf(mid)<0)tx.membershipIds.push(mid);
    tx.itemCount++;tx.totalQuantity+=qty;tx.totalPreNetValue+=net;
    if(status==='complete')tx.completeCount++;else tx.incompleteCount++;
    if(!tx.action&&action)tx.action=action;if(!tx.notes&&state.n)tx.notes=s_(state.n);
    const due=s_(r[I.DUE]), mod=s_(r[I.MOD]);
    lines.push({id:lineId,branchId:branch,transactionId:txId,purchaseLineId:lineId,receiptNumber:receipt,trxDate:s_(r[I.DATE]),trxTime:s_(r[I.TIME]),loyaltyPhone:s_(r[I.PHONE]),loyaltyName:s_(r[I.NAME]),itemCode:s_(r[I.CODE]),description:s_(r[I.DESC]),quantity:qty,preNetValue:net,netValueSum:Number(r[I.SUM])||0,averageQuantity:Number(r[I.AVG])||0,lastPurchaseDay:s_(r[I.LAST]),calculatedDueDate:due,modifiedDueDate:mod,effectiveDueDate:mod||due,membershipId:mid,status:status,action:action,notes:s_(state.n),statusUpdatedAt:s_(state.su),modifiedDate:s_(state.md),updatedAt:now,schemaVersion:FIRESTORE_INDEX_CONFIG.SCHEMA_VERSION});
  });
  const transactions=Object.keys(txs).map(function(k){const t=txs[k];t.status=t.itemCount&&t.completeCount===t.itemCount?'complete':'incomplete';t.totalQuantity=firestoreRound_(t.totalQuantity,3);t.totalPreNetValue=firestoreRound_(t.totalPreNetValue,2);return t;});
  return {transactions:transactions,lines:lines};
}

function firestoreStableJson_(o) {
  const x={};Object.keys(o||{}).sort().forEach(function(k){if(k!=='updatedAt')x[k]=o[k];});return JSON.stringify(x);
}

// Hot-path sync for user actions: callers pass complete transaction rows from one Sheets snapshot.
// This avoids two Firestore read round-trips (purchaseLines + transactions) before the commit.
// It is safe because each affected transaction is rebuilt from all its canonical sheet item lines.
function firestoreSyncTransactionsP0661_(rows) {
  firestoreAssertTestTarget_();
  rows = rows || [];
  if (!rows.length) return {lineDocumentsWritten:0, transactionDocumentsWritten:0};
  const projection = firestoreBuildCurrentProjection_(FIRESTORE_INDEX_CONFIG.BRANCH, rows);
  const writes = [];
  projection.lines.forEach(function(line) {
    writes.push(firestoreWrite_(FIRESTORE_INDEX_CONFIG.LINE_COLLECTION, line.id, line));
  });
  projection.transactions.forEach(function(tx) {
    writes.push(firestoreWrite_(FIRESTORE_INDEX_CONFIG.TX_COLLECTION, tx.id, tx));
  });
  for (let i = 0; i < writes.length; i += FIRESTORE_INDEX_CONFIG.COMMIT_BATCH_SIZE) {
    firestoreCommit_(writes.slice(i, i + FIRESTORE_INDEX_CONFIG.COMMIT_BATCH_SIZE));
  }
  return {lineDocumentsWritten:projection.lines.length, transactionDocumentsWritten:projection.transactions.length};
}

// Upsert only changed documents. Used after sheet-side actions and before Explore reads.
function firestoreSyncCurrentP0661_(rows) {
  firestoreAssertTestTarget_();
  const projection=firestoreBuildCurrentProjection_(FIRESTORE_INDEX_CONFIG.BRANCH,rows||readBranch_(FIRESTORE_INDEX_CONFIG.BRANCH));
  const oldLines=firestoreListDocuments_(FIRESTORE_INDEX_CONFIG.LINE_COLLECTION), oldTx=firestoreListDocuments_(FIRESTORE_INDEX_CONFIG.TX_COLLECTION);
  const lineMap=Object.create(null), txMap=Object.create(null);
  oldLines.forEach(function(x){lineMap[x.id]=x.data;});oldTx.forEach(function(x){txMap[x.id]=x.data;});
  const writes=[];let lineChanges=0,txChanges=0;
  projection.lines.forEach(function(x){const old=lineMap[x.id];if(!old||firestoreStableJson_(old)!==firestoreStableJson_(x)){writes.push(firestoreWrite_(FIRESTORE_INDEX_CONFIG.LINE_COLLECTION,x.id,x));lineChanges++;}});
  projection.transactions.forEach(function(x){const old=txMap[x.id];if(!old||firestoreStableJson_(old)!==firestoreStableJson_(x)){writes.push(firestoreWrite_(FIRESTORE_INDEX_CONFIG.TX_COLLECTION,x.id,x));txChanges++;}});
  const desiredLines=Object.create(null), desiredTx=Object.create(null);
  projection.lines.forEach(function(x){desiredLines[x.id]=1;});projection.transactions.forEach(function(x){desiredTx[x.id]=1;});
  oldLines.forEach(function(x){if(!desiredLines[x.id])writes.push({delete:{name:firestoreDocumentName_(FIRESTORE_INDEX_CONFIG.LINE_COLLECTION,x.id)}});});
  oldTx.forEach(function(x){if(!desiredTx[x.id])writes.push({delete:{name:firestoreDocumentName_(FIRESTORE_INDEX_CONFIG.TX_COLLECTION,x.id)}});});
  for(let i=0;i<writes.length;i+=FIRESTORE_INDEX_CONFIG.COMMIT_BATCH_SIZE)firestoreCommit_(writes.slice(i,i+FIRESTORE_INDEX_CONFIG.COMMIT_BATCH_SIZE));
  return {lineDocumentsChanged:lineChanges,transactionDocumentsChanged:txChanges,staleDocumentsDeleted:writes.filter(function(w){return !!w.delete;}).length};
}

// Read-only guard for the hot Explore path. It intentionally avoids opening the
// spreadsheet; the configured Firestore project/branch and test spreadsheet ID are
// fixed in this test-only script snapshot. Write paths retain the stronger live-sheet guard.
function firestoreAssertIndexReadTarget_() {
  if (typeof CONFIG === 'undefined' || CONFIG.CLOUD_ID !== FIRESTORE_INDEX_CONFIG.CLOUD_ID) {
    throw new Error('STOPPED: Firestore index is not configured for the approved test spreadsheet.');
  }
  if (FIRESTORE_INDEX_CONFIG.PROJECT_ID !== 'insurance-refill-portal' ||
      FIRESTORE_INDEX_CONFIG.BRANCH !== 'P0661' ||
      FIRESTORE_INDEX_CONFIG.LINE_COLLECTION !== 'branches/P0661/purchaseLines') {
    throw new Error('STOPPED: Firestore read target is outside the approved P0661 test scope.');
  }
}

// Single-read Explore snapshot: derive both the portal row array and follow-up map
// from the same Firestore response instead of listing purchaseLines twice.
function firestoreReadExploreSnapshotLegacyP0661_() {
  firestoreAssertIndexReadTarget_();
  const docs = firestoreListDocuments_(FIRESTORE_INDEX_CONFIG.LINE_COLLECTION);
  const rows = [], follow = Object.create(null);
  docs.forEach(function(x) {
    const d = x.data, r = [];
    r[I.RC]=s_(d.receiptNumber); r[I.MON]=s_(d.trxDate).slice(0,7); r[I.DATE]=s_(d.trxDate);
    r[I.TIME]=s_(d.trxTime); r[I.NAME]=s_(d.loyaltyName); r[I.PHONE]=s_(d.loyaltyPhone);
    r[I.CODE]=s_(d.itemCode); r[I.DESC]=s_(d.description); r[I.QTY]=Number(d.quantity)||0;
    r[I.NET]=Number(d.preNetValue)||0; r[I.AVG]=Number(d.averageQuantity)||0; r[I.LAST]=s_(d.lastPurchaseDay);
    r[I.DUE]=s_(d.calculatedDueDate); r[I.SUM]=Number(d.netValueSum)||0; r[I.MID]=s_(d.membershipId);
    r[I.ST]=d.status==='complete'?'complete':''; r[I.MOD]=s_(d.modifiedDueDate);
    rows.push(r);
    const key=s_(d.loyaltyPhone)+'|'+s_(d.trxDate)+'|'+s_(d.trxTime)+'|'+s_(d.membershipId)+'|'+s_(d.itemCode);
    const actionIndex=CONFIG.ACTIONS.indexOf(s_(d.action))+1;
    follow[key]={a:actionIndex>0?actionIndex:0,n:s_(d.notes),u:'',status:s_(d.status)||'incomplete',su:s_(d.statusUpdatedAt),mod:s_(d.modifiedDueDate),md:s_(d.modifiedDate)};
  });
  return {rows:rows,follow:follow};
}

// Kept as compatibility wrappers for any other existing callers.
function firestoreReadBranchRowsP0661_() { return firestoreReadExploreSnapshotLegacyP0661_().rows; }

function firestoreSyncAffectedP0661_(rows) {
  firestoreAssertTestTarget_();
  rows = rows || [];
  if (!rows.length) return {lineDocumentsChanged:0, transactionDocumentsChanged:0};

  // Incremental write path: only the affected purchase lines are rebuilt from
  // the canonical Sheets state. We do NOT rescan the whole branch or compare
  // all 241 indexed documents. This keeps Follow Up writes isolated from the
  // fast Explore read path.
  const follow = followMap_(FIRESTORE_INDEX_CONFIG.BRANCH);
  const affected = Object.create(null);
  const txIds = Object.create(null);
  const now = new Date().toISOString();
  rows.forEach(function(r){
    const lineId = firestoreLineIdForRow_(r,FIRESTORE_INDEX_CONFIG.BRANCH);
    const txId = firestoreTxIdForRow_(r);
    const txKey = fkey_(r);
    const state = follow[flinekey_(r,FIRESTORE_INDEX_CONFIG.BRANCH)] || follow[txKey] || {};
    // Follow Up is the authoritative source for status after a status toggle.
    // The in-memory branch snapshot may predate the just-saved status change.
    const status = s_(state.status || r[I.ST]).toLowerCase()==='complete' ? 'complete' : 'incomplete';
    const actionIndex = Number(state.a)||0;
    const action = actionIndex>0 && actionIndex<=CONFIG.ACTIONS.length ? CONFIG.ACTIONS[actionIndex-1] : '';
    affected[lineId] = {
      id:lineId, branchId:FIRESTORE_INDEX_CONFIG.BRANCH, transactionId:txId, purchaseLineId:lineId,
      receiptNumber:s_(r[I.RC]), trxDate:s_(r[I.DATE]), trxTime:s_(r[I.TIME]),
      loyaltyPhone:s_(r[I.PHONE]), loyaltyName:s_(r[I.NAME]), itemCode:s_(r[I.CODE]), description:s_(r[I.DESC]),
      quantity:Number(r[I.QTY])||0, preNetValue:Number(r[I.NET])||0, netValueSum:Number(r[I.SUM])||0,
      averageQuantity:Number(r[I.AVG])||0, lastPurchaseDay:s_(r[I.LAST]), calculatedDueDate:s_(r[I.DUE]),
      modifiedDueDate:s_(r[I.MOD]), effectiveDueDate:s_(r[I.MOD])||s_(r[I.DUE]), membershipId:s_(r[I.MID]),
      status:status, action:action, notes:s_(state.n), statusUpdatedAt:s_(state.su),
      modifiedDate:s_(state.md), updatedAt:now, schemaVersion:FIRESTORE_INDEX_CONFIG.SCHEMA_VERSION
    };
    txIds[txId] = 1;
  });

  const wantedTxIds = Object.keys(txIds);
  const existingLines = firestoreQueryByTransactionIds_(FIRESTORE_INDEX_CONFIG.LINE_COLLECTION, wantedTxIds);
  const oldLineMap = Object.create(null), lineMap = Object.create(null);
  existingLines.forEach(function(x){ oldLineMap[x.id]=x.data; lineMap[x.id]=x.data; });
  Object.keys(affected).forEach(function(id){ lineMap[id]=affected[id]; });

  // Rebuild only transactions touched by the changed lines. Existing line docs
  // for those transactions are used as the remaining canonical index state.
  const txAgg = Object.create(null);
  Object.keys(lineMap).forEach(function(id){
    const line=lineMap[id], txId=s_(line.transactionId);
    if (!txIds[txId]) return;
    let tx=txAgg[txId];
    if(!tx) tx=txAgg[txId]={
      id:txId, branchId:FIRESTORE_INDEX_CONFIG.BRANCH, transactionId:txId,
      loyaltyPhone:s_(line.loyaltyPhone), loyaltyName:s_(line.loyaltyName), trxDate:s_(line.trxDate), trxTime:s_(line.trxTime),
      receiptNumbers:[], membershipIds:[], itemCount:0, totalQuantity:0, totalPreNetValue:0,
      completeCount:0, incompleteCount:0, status:'incomplete', action:'', notes:'', schemaVersion:FIRESTORE_INDEX_CONFIG.SCHEMA_VERSION, updatedAt:now
    };
    const receipt=s_(line.receiptNumber), mid=s_(line.membershipId);
    if(receipt && tx.receiptNumbers.indexOf(receipt)<0) tx.receiptNumbers.push(receipt);
    if(mid && tx.membershipIds.indexOf(mid)<0) tx.membershipIds.push(mid);
    tx.itemCount++; tx.totalQuantity += Number(line.quantity)||0; tx.totalPreNetValue += Number(line.preNetValue)||0;
    if(s_(line.status)==='complete') tx.completeCount++; else tx.incompleteCount++;
    // Keep transaction-level fields consistent with the full projection.
    if(!tx.action && s_(line.action)) tx.action=s_(line.action);
    if(!tx.notes && s_(line.notes)) tx.notes=s_(line.notes);
  });
  Object.keys(txAgg).forEach(function(id){
    const tx=txAgg[id];
    tx.status=tx.itemCount>0 && tx.completeCount===tx.itemCount ? 'complete':'incomplete';
    tx.totalQuantity=firestoreRound_(tx.totalQuantity,3);
    tx.totalPreNetValue=firestoreRound_(tx.totalPreNetValue,2);
  });

  const oldTx = firestoreQueryByTransactionIds_(FIRESTORE_INDEX_CONFIG.TX_COLLECTION, wantedTxIds);
  const oldTxMap=Object.create(null); oldTx.forEach(function(x){oldTxMap[x.id]=x.data;});
  const writes=[]; let lineChanges=0, txChanges=0;
  Object.keys(affected).forEach(function(id){
    const next=affected[id], old=oldLineMap[id];
    if(!old || firestoreStableJson_(old)!==firestoreStableJson_(next)){ writes.push(firestoreWrite_(FIRESTORE_INDEX_CONFIG.LINE_COLLECTION,id,next)); lineChanges++; }
  });
  Object.keys(txAgg).forEach(function(id){
    const next=txAgg[id], old=oldTxMap[id];
    if(!old || firestoreStableJson_(old)!==firestoreStableJson_(next)){ writes.push(firestoreWrite_(FIRESTORE_INDEX_CONFIG.TX_COLLECTION,id,next)); txChanges++; }
  });
  for(let i=0;i<writes.length;i+=FIRESTORE_INDEX_CONFIG.COMMIT_BATCH_SIZE) {
    firestoreCommit_(writes.slice(i,i+FIRESTORE_INDEX_CONFIG.COMMIT_BATCH_SIZE));
  }
  return {lineDocumentsChanged:lineChanges, transactionDocumentsChanged:txChanges};
}

// Fast Member History read for P0661. It uses the already-indexed purchase lines
// and preserves the existing invoice grouping: phone + transaction date + time.
function firestoreGetMemberHistoryP0661_(membershipId) {
  firestoreAssertIndexReadTarget_();
  const mid=String(membershipId||'').trim();
  if(!mid) throw new Error('Membership ID is required.');
  const docs=firestoreListDocuments_(FIRESTORE_INDEX_CONFIG.LINE_COLLECTION);
  const grouped=Object.create(null), order=[];
  docs.forEach(function(x){
    const d=x.data;
    if(s_(d.membershipId)!==mid) return;
    const k=s_(d.loyaltyPhone)+'|'+s_(d.trxDate)+'|'+s_(d.trxTime);
    if(!grouped[k]) {
      grouped[k]={membershipId:mid,phone:s_(d.loyaltyPhone),name:s_(d.loyaltyName),trxDate:s_(d.trxDate),trxTime:s_(d.trxTime),action:s_(d.action),notes:s_(d.notes),items:[]};
      order.push(k);
    } else {
      if(!grouped[k].action && s_(d.action)) grouped[k].action=s_(d.action);
      if(!grouped[k].notes && s_(d.notes)) grouped[k].notes=s_(d.notes);
    }
    grouped[k].items.push({code:s_(d.itemCode),description:s_(d.description),quantity:Number(d.quantity)||0,calculatedDueDate:s_(d.calculatedDueDate),modifiedDueDate:s_(d.modifiedDueDate)||'',dueDate:s_(d.effectiveDueDate)||s_(d.calculatedDueDate),status:s_(d.status)==='complete'?'Complete':'Incomplete'});
  });
  order.sort(function(a,c){ const aa=a.split('|'), cc=c.split('|'); return aa[1]===cc[1]?aa[2].localeCompare(cc[2]):aa[1].localeCompare(cc[1]); });
  return {membershipId:mid,invoices:order.map(function(k){return grouped[k];})};
}

function firestoreReadFollowMapP0661_() {
  firestoreAssertTestTarget_();
  const map=Object.create(null);
  firestoreListDocuments_(FIRESTORE_INDEX_CONFIG.LINE_COLLECTION).forEach(function(x){
    const d=x.data, key=s_(d.loyaltyPhone)+'|'+s_(d.trxDate)+'|'+s_(d.trxTime)+'|'+s_(d.membershipId)+'|'+s_(d.itemCode);
    const actionIndex=CONFIG.ACTIONS.indexOf(s_(d.action))+1;
    map[key]={a:actionIndex>0?actionIndex:0,n:s_(d.notes),u:'',status:s_(d.status)||'incomplete',su:s_(d.statusUpdatedAt),mod:s_(d.modifiedDueDate),md:s_(d.modifiedDate)};
  });
  return map;
}

