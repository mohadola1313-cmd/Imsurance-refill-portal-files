/**
 * Insurance Refill Portal — P0661 reliability layer.
 * TEST COPY ONLY. Google Sheets remains canonical; Firestore is a rebuildable read model.
 *
 * This file deliberately does not install triggers or touch live data at load time.
 * Run firestoreInstallRetryTriggerP0661_() manually only after reviewing the deployment.
 */

const FIRESTORE_RELIABILITY_P0661 = Object.freeze({
  QUEUE_SHEET: '_Firestore Sync Queue',
  QUEUE_HEADERS: ['Job ID','Branch','Transaction Key','Operation','State','Attempts',
    'Next Attempt','Created At','Updated At','Lease Until','Last Error'],
  MAX_JOBS_PER_RUN: 8,
  LEASE_MINUTES: 10,
  BASE_BACKOFF_MINUTES: 5,
  MAX_BACKOFF_MINUTES: 360,
  CACHE_TTL_SECONDS: 1800,
  CACHE_CHUNK_BYTES: 75000,
  SNAPSHOT_SCHEMA: 2
});

function firestoreReliabilityAssertP0661_() {
  firestoreAssertTestTarget_();
  if (FIRESTORE_INDEX_CONFIG.BRANCH !== 'P0661' ||
      FIRESTORE_INDEX_CONFIG.CLOUD_ID !== '1HTdkNkADKizQxycfDx-jgnWEsqaHVxJ1XkrkGlB3KxY') {
    throw new Error('STOPPED: reliability layer is restricted to the approved P0661 test copy.');
  }
}

function firestoreQueueSheetP0661_() {
  const ss = cloud_();
  if (ss.getId() !== FIRESTORE_INDEX_CONFIG.CLOUD_ID) {
    throw new Error('STOPPED: queue spreadsheet is not the approved test copy.');
  }
  let sh = ss.getSheetByName(FIRESTORE_RELIABILITY_P0661.QUEUE_SHEET);
  if (!sh) sh = ss.insertSheet(FIRESTORE_RELIABILITY_P0661.QUEUE_SHEET);
  const headers = FIRESTORE_RELIABILITY_P0661.QUEUE_HEADERS;
  if (sh.getLastRow() === 0) {
    sh.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  } else {
    const old = sh.getRange(1,1,1,Math.max(headers.length,sh.getLastColumn())).getValues()[0].map(String);
    headers.forEach(function(h,i) {
      if (old[i] !== h) sh.getRange(1,i+1).setValue(h);
    });
  }
  return sh;
}

function firestoreQueueEnqueueTransactionsP0661_(transactionKeys, operation) {
  firestoreReliabilityAssertP0661_();
  const keys = Object.create(null);
  (transactionKeys || []).forEach(function(k) { k = s_(k); if (k) keys[k] = true; });
  const list = Object.keys(keys);
  if (!list.length) return 0;

  const sh = firestoreQueueSheetP0661_();
  const last = sh.getLastRow();
  const vals = last > 1 ? sh.getRange(2,1,last-1,FIRESTORE_RELIABILITY_P0661.QUEUE_HEADERS.length).getValues() : [];
  const pending = Object.create(null), changed = Object.create(null);
  vals.forEach(function(r,i) {
    if (s_(r[1]) === 'P0661' && (s_(r[4]) === 'pending' || s_(r[4]) === 'processing')) {
      pending[s_(r[2])] = i;
    }
  });

  const now = new Date().toISOString();
  const add = [];
  list.forEach(function(key) {
    const ix = pending[key];
    if (ix !== undefined) {
      vals[ix][3] = String(operation || 'upsert');
      vals[ix][5] = 0;
      vals[ix][6] = now;
      vals[ix][8] = now;
      vals[ix][9] = '';
      vals[ix][10] = '';
      changed[ix] = true;
      // A pending job always rebuilds from the latest canonical Sheets state.
      return;
    }
    add.push([Utilities.getUuid(),'P0661',key,String(operation || 'upsert'),'pending',0,
      now,now,now,'','']);
  });
  Object.keys(changed).map(Number).sort(function(a,b){return a-b;}).forEach(function(ix) {
    firestoreQueueSetRowP0661_(sh,ix+2,vals[ix]);
  });
  if (add.length) sh.getRange(sh.getLastRow()+1,1,add.length,FIRESTORE_RELIABILITY_P0661.QUEUE_HEADERS.length).setValues(add);
  return list.length;
}

function firestoreQueueHasPendingP0661_() {
  const ss = cloud_();
  if (ss.getId() !== FIRESTORE_INDEX_CONFIG.CLOUD_ID) throw new Error('STOPPED: unexpected spreadsheet.');
  const sh = ss.getSheetByName(FIRESTORE_RELIABILITY_P0661.QUEUE_SHEET);
  if (!sh || sh.getLastRow() < 2) return false;
  const vals = sh.getRange(2,2,sh.getLastRow()-1,4).getValues();
  return vals.some(function(r) {
    return s_(r[0]) === 'P0661' && (s_(r[3]) === 'pending' || s_(r[3]) === 'processing');
  });
}

function firestoreQueueSetRowP0661_(sh, rowNumber, row) {
  sh.getRange(rowNumber,1,1,FIRESTORE_RELIABILITY_P0661.QUEUE_HEADERS.length).setValues([row]);
}

function firestoreReconcileTransactionP0661_(transactionKey, branchRows) {
  firestoreReliabilityAssertP0661_();
  const branch = FIRESTORE_INDEX_CONFIG.BRANCH;
  const txRows = (branchRows || readBranch_(branch)).filter(function(r) { return fkey_(r) === transactionKey; });
  const txId = firestoreStableId_('t', transactionKey);
  const existing = firestoreQueryByTransactionIds_(FIRESTORE_INDEX_CONFIG.LINE_COLLECTION,[txId]);
  const existingMap = Object.create(null);
  existing.forEach(function(x) { existingMap[x.id] = x.data; });

  const projection = txRows.length ? firestoreBuildCurrentProjection_(branch,txRows) : {transactions:[],lines:[]};
  const desired = Object.create(null), writes = [];
  projection.lines.forEach(function(line) {
    desired[line.id] = true;
    const old = existingMap[line.id];
    if (!old || firestoreStableJson_(old) !== firestoreStableJson_(line)) {
      writes.push(firestoreWrite_(FIRESTORE_INDEX_CONFIG.LINE_COLLECTION,line.id,line));
    }
  });
  existing.forEach(function(x) {
    if (!desired[x.id]) writes.push({delete:{name:firestoreDocumentName_(FIRESTORE_INDEX_CONFIG.LINE_COLLECTION,x.id)}});
  });

  if (projection.transactions.length) {
    const tx = projection.transactions[0];
    writes.push(firestoreWrite_(FIRESTORE_INDEX_CONFIG.TX_COLLECTION,tx.id,tx));
  } else {
    writes.push({delete:{name:firestoreDocumentName_(FIRESTORE_INDEX_CONFIG.TX_COLLECTION,txId)}});
  }

  for (let i=0;i<writes.length;i+=FIRESTORE_INDEX_CONFIG.COMMIT_BATCH_SIZE) {
    firestoreCommit_(writes.slice(i,i+FIRESTORE_INDEX_CONFIG.COMMIT_BATCH_SIZE));
  }
  return {transactionId:txId,lines:projection.lines.length,writes:writes.length};
}

function firestoreProcessPendingLockedP0661_() {
  firestoreReliabilityAssertP0661_();
  if (PropertiesService.getScriptProperties().getProperty('FIRESTORE_INDEX_P0661_READY') !== '1') {
    return {processed:0,reason:'index-not-ready'};
  }
  const sh = firestoreQueueSheetP0661_();
  const last = sh.getLastRow();
  if (last < 2) return {processed:0};
  const vals = sh.getRange(2,1,last-1,FIRESTORE_RELIABILITY_P0661.QUEUE_HEADERS.length).getValues();
  const now = Date.now(), nowIso = new Date(now).toISOString();
  const max = FIRESTORE_RELIABILITY_P0661.MAX_JOBS_PER_RUN;
  let processed = 0, changed = false;
  const eligibleIndexes=[];
  for(let i=0;i<vals.length;i++) {
    const r=vals[i],state=s_(r[4]),next=Date.parse(s_(r[6]))||0,lease=Date.parse(s_(r[9]))||0;
    if(s_(r[1])==='P0661' && (state==='pending'&&next<=now || state==='processing'&&lease<=now)) {
      eligibleIndexes.push(i);
      if(eligibleIndexes.length>=max) break;
    }
  }
  if(!eligibleIndexes.length) return {processed:0};
  const branchRows = readBranch_(FIRESTORE_INDEX_CONFIG.BRANCH);

  for (let j=0;j<eligibleIndexes.length;j++) {
    const i=eligibleIndexes[j],r=vals[i];
    r[4] = 'processing';
    r[8] = nowIso;
    r[9] = new Date(now + FIRESTORE_RELIABILITY_P0661.LEASE_MINUTES*60000).toISOString();
    firestoreQueueSetRowP0661_(sh,i+2,r);
    changed = true;
    try {
      firestoreReconcileTransactionP0661_(s_(r[2]),branchRows);
      r[4] = 'complete';
      r[8] = new Date().toISOString();
      r[9] = '';
      r[10] = '';
      PropertiesService.getScriptProperties().setProperty('FIRESTORE_SNAPSHOT_REV_P0661',
        String((Number(PropertiesService.getScriptProperties().getProperty('FIRESTORE_SNAPSHOT_REV_P0661'))||0)+1));
    } catch (err) {
      const attempts = (Number(r[5])||0)+1;
      const delay = Math.min(FIRESTORE_RELIABILITY_P0661.MAX_BACKOFF_MINUTES,
        FIRESTORE_RELIABILITY_P0661.BASE_BACKOFF_MINUTES*Math.pow(2,Math.min(attempts-1,10)));
      r[4] = 'pending';
      r[5] = attempts;
      r[6] = new Date(Date.now()+delay*60000).toISOString();
      r[8] = new Date().toISOString();
      r[9] = '';
      r[10] = String(err && err.message || err).slice(0,1000);
    }
    firestoreQueueSetRowP0661_(sh,i+2,r);
    processed++;
  }
  return {processed:processed,changed:changed};
}

/** Time-driven worker. Safe to run manually for a controlled test after index readiness. */
function firestoreRetryWorkerP0661_() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    return firestoreProcessPendingLockedP0661_();
  } finally {
    lock.releaseLock();
  }
}

/** Installs one five-minute retry trigger. This function is intentionally manual. */
function firestoreInstallRetryTriggerP0661_() {
  firestoreReliabilityAssertP0661_();
  if (PropertiesService.getScriptProperties().getProperty('FIRESTORE_INDEX_P0661_READY') !== '1') {
    throw new Error('Build and verify the P0661 Firestore index before installing the retry trigger.');
  }
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'firestoreRetryWorkerP0661_') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('firestoreRetryWorkerP0661_').timeBased().everyMinutes(5).create();
  return 'Firestore retry trigger installed for the P0661 test project.';
}

function firestoreChangedTransactionKeysP0661_(beforeRows,afterRows) {
  const before = Object.create(null), after = Object.create(null), changed = Object.create(null);
  (beforeRows||[]).forEach(function(r) { before[flinekey_(r)] = r; });
  (afterRows||[]).forEach(function(r) { after[flinekey_(r)] = r; });
  Object.keys(before).forEach(function(k) {
    if (!after[k] || JSON.stringify(before[k]) !== JSON.stringify(after[k])) changed[fkey_(before[k])] = true;
  });
  Object.keys(after).forEach(function(k) {
    if (!before[k] || JSON.stringify(before[k]) !== JSON.stringify(after[k])) changed[fkey_(after[k])] = true;
  });
  return Object.keys(changed);
}

function firestoreAutomaticAffectedTransactionKeysP0661_(rows) {
  const latest = Object.create(null), affected = Object.create(null);
  (rows||[]).forEach(function(r) {
    const mid=s_(r[I.MID]), code=s_(r[I.CODE]), stamp=s_(r[I.DATE])+'T'+s_(r[I.TIME]);
    if (!mid || !code || !s_(r[I.DATE]) || !s_(r[I.TIME])) return;
    const k=mid+'|'+code;
    if (!latest[k] || stamp>latest[k]) latest[k]=stamp;
  });
  (rows||[]).forEach(function(r) {
    const mid=s_(r[I.MID]), code=s_(r[I.CODE]), stamp=s_(r[I.DATE])+'T'+s_(r[I.TIME]);
    if (mid && code && latest[mid+'|'+code] && stamp<latest[mid+'|'+code] &&
        s_(r[I.ST]).toLowerCase() !== 'complete') affected[fkey_(r)]=true;
  });
  return Object.keys(affected);
}

/* ================= Monthly Explore snapshot cache ================= */

function firestoreExploreMonthBoundsP0661_(month) {
  if (!/^\d{4}-\d{2}$/.test(String(month||''))) return null;
  const p=String(month).split('-'), y=Number(p[0]), m=Number(p[1]);
  if (m<1 || m>12) return null;
  const start=String(month)+'-01';
  const next=new Date(Date.UTC(y,m,1)).toISOString().slice(0,10);
  return {start:start,end:next};
}

function firestoreQueryExploreMonthP0661_(month) {
  firestoreAssertIndexReadTarget_();
  const bounds=firestoreExploreMonthBoundsP0661_(month);
  if (!bounds) throw new Error('A valid YYYY-MM month is required for the Explore snapshot.');
  const parent='projects/'+FIRESTORE_INDEX_CONFIG.PROJECT_ID+'/databases/'+FIRESTORE_INDEX_CONFIG.DATABASE_ID+
    '/documents/branches/P0661';
  const url='https://firestore.googleapis.com/v1/'+parent+':runQuery';
  const out=[];
  let cursor=null;
  while (true) {
    const query={
      from:[{collectionId:'purchaseLines'}],
      where:{compositeFilter:{op:'AND',filters:[
        {fieldFilter:{field:{fieldPath:'effectiveDueDate'},op:'GREATER_THAN_OR_EQUAL',value:{stringValue:bounds.start}}},
        {fieldFilter:{field:{fieldPath:'effectiveDueDate'},op:'LESS_THAN',value:{stringValue:bounds.end}}}
      ]}},
      orderBy:[{field:{fieldPath:'effectiveDueDate'},direction:'ASCENDING'},{field:{fieldPath:'__name__'},direction:'ASCENDING'}],
      limit:1000
    };
    if (cursor) query.startAt={values:[{stringValue:cursor.due},{referenceValue:cursor.name}],before:false};
    const response=UrlFetchApp.fetch(url,{
      method:'post',contentType:'application/json',
      headers:{Authorization:'Bearer '+ScriptApp.getOAuthToken()},
      payload:JSON.stringify({structuredQuery:query}),muteHttpExceptions:true
    });
    const code=response.getResponseCode();
    if(code<200||code>=300) throw new Error('Firestore monthly Explore query failed ('+code+'): '+response.getContentText().slice(0,1000));
    const results=JSON.parse(response.getContentText()||'[]');
    let page=0,last=null;
    results.forEach(function(item){
      if(!item||!item.document)return;
      const doc=item.document, data=firestoreDecodeFields_(doc.fields||{});
      out.push({id:doc.name.split('/').pop(),data:data});
      last={due:s_(data.effectiveDueDate),name:doc.name};
      page++;
    });
    if(page<1000 || !last) break;
    cursor=last;
  }
  return out;
}

function firestoreExploreRowsFromDocsP0661_(docs) {
  const rows=[],follow=Object.create(null);
  (docs||[]).forEach(function(x){
    const d=x.data,r=[];
    r[I.RC]=s_(d.receiptNumber);r[I.MON]=s_(d.trxDate).slice(0,7);r[I.DATE]=s_(d.trxDate);
    r[I.TIME]=s_(d.trxTime);r[I.NAME]=s_(d.loyaltyName);r[I.PHONE]=s_(d.loyaltyPhone);
    r[I.CODE]=s_(d.itemCode);r[I.DESC]=s_(d.description);r[I.QTY]=Number(d.quantity)||0;
    r[I.NET]=Number(d.preNetValue)||0;r[I.AVG]=Number(d.averageQuantity)||0;r[I.LAST]=s_(d.lastPurchaseDay);
    r[I.DUE]=s_(d.calculatedDueDate);r[I.SUM]=Number(d.netValueSum)||0;r[I.MID]=s_(d.membershipId);
    r[I.ST]=d.status==='complete'?'complete':'';r[I.MOD]=s_(d.modifiedDueDate);rows.push(r);
    const key=s_(d.loyaltyPhone)+'|'+s_(d.trxDate)+'|'+s_(d.trxTime)+'|'+s_(d.membershipId)+'|'+s_(d.itemCode);
    const actionIndex=CONFIG.ACTIONS.indexOf(s_(d.action))+1;
    follow[key]={a:actionIndex>0?actionIndex:0,n:s_(d.notes),u:'',status:s_(d.status)||'incomplete',
      su:s_(d.statusUpdatedAt),mod:s_(d.modifiedDueDate),md:s_(d.modifiedDate)};
  });
  return {rows:rows,follow:follow};
}

function firestoreSnapshotCacheKeyP0661_(month,revision,part) {
  return 'ifr:p0661:explore:v'+FIRESTORE_RELIABILITY_P0661.SNAPSHOT_SCHEMA+':r'+revision+':'+month+':'+part;
}

function firestoreReadCanonicalExploreFallbackP0661_(source) {
  const lock=LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const branch=FIRESTORE_INDEX_CONFIG.BRANCH;
    return {rows:readBranch_(branch),follow:followMap_(branch),source:source||'sheets-fallback'};
  } finally {
    lock.releaseLock();
  }
}

function firestoreExploreSnapshotStillCurrentP0661_(revision) {
  const current=Number(PropertiesService.getScriptProperties().getProperty('FIRESTORE_SNAPSHOT_REV_P0661'))||0;
  return current===revision && !firestoreQueueHasPendingP0661_();
}

function firestoreReadExploreSnapshotP0661_(month) {
  firestoreAssertIndexReadTarget_();
  const bounds=firestoreExploreMonthBoundsP0661_(month);
  if(!bounds) return null;
  // A pending write makes the Firestore projection potentially stale. Read canonical Sheets instead.
  if(firestoreQueueHasPendingP0661_()) return firestoreReadCanonicalExploreFallbackP0661_('sheets-pending-sync');
  const props=PropertiesService.getScriptProperties();
  const revision=Number(props.getProperty('FIRESTORE_SNAPSHOT_REV_P0661'))||0;
  const cache=CacheService.getScriptCache();
  const manifestKey=firestoreSnapshotCacheKeyP0661_(month,revision,'manifest');
  const rawManifest=cache.get(manifestKey);
  if(rawManifest) {
    try {
      const manifest=JSON.parse(rawManifest),docs=[];
      if(manifest.schema===FIRESTORE_RELIABILITY_P0661.SNAPSHOT_SCHEMA &&
         manifest.revision===revision && manifest.month===month && manifest.count>0) {
        let valid=true;
        for(let i=0;i<manifest.count;i++) {
          const raw=cache.get(firestoreSnapshotCacheKeyP0661_(month,revision,'part'+i));
          if(!raw){valid=false;break;}
          const chunk=JSON.parse(raw);
          if(!Array.isArray(chunk)){valid=false;break;}
          Array.prototype.push.apply(docs,chunk);
        }
        if(valid && docs.length===Number(manifest.documents)) {
          const snapshot=firestoreExploreRowsFromDocsP0661_(docs);
          snapshot.source='cache';
          if(firestoreExploreSnapshotStillCurrentP0661_(revision)) return snapshot;
          return firestoreReadCanonicalExploreFallbackP0661_('sheets-concurrent-update');
        }
      }
    } catch(e) {
      // Treat malformed or incomplete cache entries as a cache miss.
    }
  }

  const docs=firestoreQueryExploreMonthP0661_(month);
  const snapshot=firestoreExploreRowsFromDocsP0661_(docs);
  snapshot.source='firestore';
  if(!firestoreExploreSnapshotStillCurrentP0661_(revision)) {
    return firestoreReadCanonicalExploreFallbackP0661_('sheets-concurrent-update');
  }
  const chunks=[];let chunk=[],bytes=2;
  docs.forEach(function(doc){
    const item=JSON.stringify(doc),size=Utilities.newBlob(item).getBytes().length+1;
    if(size>FIRESTORE_RELIABILITY_P0661.CACHE_CHUNK_BYTES){chunks.length=0;chunk=null;return;}
    if(chunk && bytes+size>FIRESTORE_RELIABILITY_P0661.CACHE_CHUNK_BYTES){chunks.push(chunk);chunk=[];bytes=2;}
    if(chunk){chunk.push(doc);bytes+=size;}
  });
  if(chunk && chunk.length) chunks.push(chunk);
  if(docs.length===0) chunks.push([]);
  if(chunks.length && chunks.every(function(c){return c!==null;})) {
    chunks.forEach(function(c,i){cache.put(firestoreSnapshotCacheKeyP0661_(month,revision,'part'+i),JSON.stringify(c),FIRESTORE_RELIABILITY_P0661.CACHE_TTL_SECONDS);});
    cache.put(manifestKey,JSON.stringify({schema:FIRESTORE_RELIABILITY_P0661.SNAPSHOT_SCHEMA,revision:revision,month:month,count:chunks.length,documents:docs.length,builtAt:new Date().toISOString()}),FIRESTORE_RELIABILITY_P0661.CACHE_TTL_SECONDS);
  }
  return snapshot;
}


/** Pure, read-only checks for deterministic helpers. Does not access Sheets or Firestore. */
function firestoreReliabilitySelfTestP0661_() {
  const assert = function(ok,message) { if(!ok) throw new Error('Self-test failed: '+message); };
  const bounds=firestoreExploreMonthBoundsP0661_('2024-02');
  assert(bounds && bounds.start==='2024-02-01' && bounds.end==='2024-03-01','leap-year month bounds');
  assert(firestoreExploreMonthBoundsP0661_('2024-13')===null,'invalid month rejected');
  assert(firestoreSnapshotCacheKeyP0661_('2026-11',3,'manifest') !==
    firestoreSnapshotCacheKeyP0661_('2026-11',4,'manifest'),'revision changes cache key');
  assert(firestoreSnapshotCacheKeyP0661_('2026-11',3,'manifest') !==
    firestoreSnapshotCacheKeyP0661_('2026-12',3,'manifest'),'month changes cache key');

  const row=function(phone,date,time,mid,code,status) {
    const r=[];r[I.PHONE]=phone;r[I.DATE]=date;r[I.TIME]=time;r[I.MID]=mid;r[I.CODE]=code;r[I.ST]=status||'incomplete';return r;
  };
  const before=[row('0550000000','2026-11-01','10:00:00','M1','A1')];
  const after=[row('0550000000','2026-11-01','10:00:00','M1','A1','complete')];
  const changed=firestoreChangedTransactionKeysP0661_(before,after);
  assert(changed.length===1 && changed[0]==='0550000000|2026-11-01|10:00:00','changed line maps to one transaction');
  const autoRows=[
    row('0550000000','2026-10-01','09:00:00','M1','A1','incomplete'),
    row('0550000000','2026-11-01','09:00:00','M1','A1','incomplete'),
    row('0550000000','2026-09-01','09:00:00','M1','B2','incomplete')
  ];
  const auto=firestoreAutomaticAffectedTransactionKeysP0661_(autoRows);
  assert(auto.length===1 && auto[0]==='0550000000|2026-10-01|09:00:00','only older incomplete matching item is queued');
  const alreadyDone=firestoreAutomaticAffectedTransactionKeysP0661_([
    row('0550000000','2026-10-01','09:00:00','M1','A1','complete'),
    row('0550000000','2026-11-01','09:00:00','M1','A1','incomplete')
  ]);
  assert(alreadyDone.length===0,'already-complete history is not re-queued');
  return {passed:7,readOnly:true};
}
