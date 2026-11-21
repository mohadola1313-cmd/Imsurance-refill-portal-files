/**
 * Insurance Refill Portal - server code (Google Apps Script), version 7.
 * Project files: Code.gs (this file), Index.html, appsscript.json
 *
 * A purchase = one loyalty phone at one Trx Date + Trx Time, with its own average qt, net value sum
 * and due date. A purchase line (phone + item code + date + time) is saved once; a repeat replaces it.
 *
 * Due date of a purchase:
 *  - milk rule (when any line of the purchase has an item code listed in the milk items file):
 *      total milk quantity of the purchase >= 10 -> 28 days, otherwise quantity x 2.8 days (rounded);
 *      every line of that purchase gets this due date.
 *  - otherwise average qt: <= 1.5 -> 28 days, > 1.5 and <= 2.5 -> 56 days, > 2.5 -> 84 days.
 * A line can also carry a "modified due date" (set by the user); it is shown instead of the calculated one.
 *
 * Customer follow up: the action (one of five choices) and the notes recorded for a customer are saved in the
 * "customer follow up" sheet of the cloud file, one row per purchase process (branch + loyalty phone + due date +
 * trx date + trx time). That sheet is separate from the branch sheets, so deleting a period never removes it.
 *
 * Rows arriving from the browser carry 11 values: indexes 0-9 as saved, and index 10 = Membership ID.
 */

const CONFIG = {
  CLOUD_ID: '1HTdkNkADKizQxycfDx-jgnWEsqaHVxJ1XkrkGlB3KxY', // TEST COPY ONLY: insurance speed test
  MASTER_ID: '1fWFb82D-zvRzj3GOVBOMw_KSVFIletVH',           // master file; item codes in column A ("Code")
  MASTER_SHEET: '',                                          // '' = first sheet
  MILK_ID: '1F25KgSGsaP9R3sIll9DzwcwXG-GkzkQM4L75Uka1umc',   // milk items file; item codes in column A ("Item code")
  MILK_SHEET: '',                                            // '' = first sheet
  USERS_SHEET: 'Users',                                      // Branch | PasswordHash | Role | Approval | Contact | Requested
  FOLLOW_SHEET: 'customer follow up',                        // actions and notes recorded on the follow up page
  ACTIONS: ['Customer ineligible for refill', 'Not contacted', 'Contacted with no answer',
            'Contacted \u2013 customer approved refill', 'Contacted \u2013 customer rejected refill'],
  SESSION_SECONDS: 21600,
  TZ: 'Asia/Riyadh',
  OUT: ['Receipt Number', 'Trx Month', 'Trx Date', 'Trx Time', 'Loyalty Name', 'Loyalty Phone', 'Code',
        'Description', 'Quantity', 'Pre-Net Value', 'average qt', 'Last Purchase Day', 'Due Date',
        'net value sum', 'Membership ID', 'status', 'modified due date']
};

// column positions in a saved row
const I = { RC: 0, MON: 1, DATE: 2, TIME: 3, NAME: 4, PHONE: 5, CODE: 6, DESC: 7, QTY: 8, NET: 9,
            AVG: 10, LAST: 11, DUE: 12, SUM: 13, MID: 14, ST: 15, MOD: 16 };
const IN_MID = 10;
const CORE_COLS = 16; // columns present since version 3; "modified due date" is the 17th

function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};
  const page = String(params.page || '').toLowerCase();
  const template = HtmlService.createTemplateFromFile('Index');

  // Pass the reset token through the server-rendered page as well as the URL.
  // Apps Script may move the web app into its userCodeAppPanel URL and lose
  // the original query string in the browser, so the client must not rely on it.
  const rawResetToken = page === 'reset' ? String(params.token || '') : '';
  // Reset tokens are opaque identifiers only. Never pass arbitrary URL text into the HTML template.
  template.resetToken = /^[A-Za-z0-9-]+$/.test(rawResetToken) ? rawResetToken : '';

  return template.evaluate()
    .setTitle('Insurance Refill Portal')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ================= setup (run once from the editor) ================= */

function setupUsersSheet() {
  const ss = cloud_();
  const sh = ss.getSheetByName(CONFIG.USERS_SHEET) || ss.insertSheet(CONFIG.USERS_SHEET);
  sh.getRange(1, 1, 1, 6).setValues([['Branch', 'PasswordHash', 'Role', 'Approval', 'Contact', 'Requested']]).setFontWeight('bold');
  sh.getRange('E:F').setNumberFormat('@');
  sh.setFrozenRows(1);
  salt_();
  return 'Users sheet ready';
}

// Secure initial setup. Add the five BOOTSTRAP_* values temporarily under
// Project Settings > Script properties, then run setupUsersSecure() once.
// Passwords are hashed before being written to Users; bootstrap values are deleted on success.
function setupUsersSecure() {
  setupUsersSheet();
  const props = PropertiesService.getScriptProperties();
  const adminId = String(props.getProperty('BOOTSTRAP_ADMIN_ID') || 'admin').trim();
  const adminPw = String(props.getProperty('BOOTSTRAP_ADMIN_PASSWORD') || '');
  const email = String(props.getProperty('BOOTSTRAP_ADMIN_EMAIL') || '').trim().toLowerCase();
  const branchId = String(props.getProperty('BOOTSTRAP_BRANCH_ID') || 'P0806').trim();
  const branchPw = String(props.getProperty('BOOTSTRAP_BRANCH_PASSWORD') || '');
  if (!adminId || !branchId || adminPw.length < 6 || branchPw.length < 6 ||
      !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    throw new Error('Set BOOTSTRAP_ADMIN_ID, BOOTSTRAP_ADMIN_PASSWORD, BOOTSTRAP_ADMIN_EMAIL, BOOTSTRAP_BRANCH_ID, and BOOTSTRAP_BRANCH_PASSWORD in Script properties. Passwords must be at least 6 characters.');
  }
  if (adminId.toLowerCase() === branchId.toLowerCase()) throw new Error('Admin ID and branch ID must be different.');
  secureUpsertUser_(adminId, adminPw, 'admin', email);
  secureUpsertUser_(branchId, branchPw, 'user', '');
  ['BOOTSTRAP_ADMIN_ID','BOOTSTRAP_ADMIN_PASSWORD','BOOTSTRAP_ADMIN_EMAIL',
   'BOOTSTRAP_BRANCH_ID','BOOTSTRAP_BRANCH_PASSWORD'].forEach(function(k) { props.deleteProperty(k); });
  return 'Admin and initial branch credentials saved as password hashes. Temporary bootstrap properties were cleared.';
}

function secureUpsertUser_(id, password, role, contact) {
  const sh = users_(), rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim().toLowerCase() === id.toLowerCase()) {
      const oldContact = String(rows[i][4] || '');
      sh.getRange(i + 1, 2).setNumberFormat('@').setValue(hashPw_(id, password));
      sh.getRange(i + 1, 3, 1, 3).setValues([[role, 'approved', contact || oldContact]]);
      return;
    }
  }
  sh.appendRow([id, hashPw_(id, password), role, 'approved', contact || '', new Date().toISOString()]);
}

// Update the recovery email using a temporary Script Property, then clear it.
function setAdminRecoveryEmail() {
  const props = PropertiesService.getScriptProperties();
  const email = String(props.getProperty('ADMIN_RECOVERY_EMAIL') || '').trim().toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    throw new Error('Set ADMIN_RECOVERY_EMAIL in Script properties, then run setAdminRecoveryEmail() again.');
  }
  const sh = users_(), data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][2]).trim().toLowerCase() === 'admin') {
      sh.getRange(i + 1, 5).setNumberFormat('@').setValue(email);
      props.deleteProperty('ADMIN_RECOVERY_EMAIL');
      return 'Admin recovery email saved and temporary property cleared.';
    }
  }
  throw new Error('No Admin account was found. Run setupUsersSecure() first.');
}

function addOrUpdateUser(branch, password, role) {
  role = String(role || 'user').toLowerCase();
  const sh = users_();
  const data = sh.getDataRange().getValues();
  const row = [branch, hashPw_(branch, password), role, 'approved', '', new Date().toISOString()];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === String(branch).trim().toLowerCase()) {
      sh.getRange(i + 1, 1, 1, 6).setValues([row]);
      return 'Updated ' + branch;
    }
  }
  sh.appendRow(row);
  return 'Added ' + branch;
}

// Run from the editor to check access to the cloud file, the Users sheet, the master file and the milk file.
// Optional: checkSetup('100335,106295') also tests whether those item codes are in each list.
function checkSetup(codes) {
  const out = [];
  out.push('Cloud file: ' + cloud_().getName());
  out.push('Users sheet: ' + (cloud_().getSheetByName(CONFIG.USERS_SHEET) ? 'found' : 'missing - run setupUsersSheet()'));
  out.push('Customer follow up sheet: ' + (cloud_().getSheetByName(CONFIG.FOLLOW_SHEET) ? 'found' : 'not created yet (created automatically on the first save)'));
  const test = String(codes || '').split(',').map(function (c) { return c.trim(); }).filter(String);
  [['Master file', masterCodes_], ['Milk items file', milkCodes_]].forEach(function (p) {
    try {
      const m = p[1]();
      out.push(p[0] + ' item codes found: ' + Object.keys(m).length);
      test.forEach(function (c) { out.push('  Code ' + c + ': ' + (m[c] ? 'found' : 'NOT found')); });
    } catch (e) { out.push(p[0] + ' problem: ' + e.message); }
  });
  Logger.log(out.join('\n'));
  return out.join('\n');
}

/* ================= registration, login, sessions ================= */

function register(branch, password, contact) {
  branch = String(branch || '').trim().replace(/\s+/g, ' ');
  password = String(password || '');
  contact = String(contact || '').trim().replace(/^[=+\-@]+/, '').slice(0, 80);
  if (!/^[A-Za-z0-9\u0600-\u06FF][A-Za-z0-9\u0600-\u06FF _.\-()]{1,48}$/.test(branch)) {
    throw new Error('Use 2 to 50 letters, numbers, spaces, or . _ - ( ) in the branch name.');
  }
  if (password.length < 6) throw new Error('The password must have at least 6 characters.');
  const cache = CacheService.getScriptCache();
  const n = Number(cache.get('reg:count') || 0);
  if (n >= 30) throw new Error('Too many registration requests. Try again later.');
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const reserved = [CONFIG.USERS_SHEET, CONFIG.FOLLOW_SHEET].map(function (n) { return n.toLowerCase(); });
    if (reserved.indexOf(branch.toLowerCase()) >= 0) throw new Error('This branch name is not available.');
    const sh = users_();
    const rows = sh.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim().toLowerCase() === branch.toLowerCase()) {
        throw new Error(String(rows[i][3] || 'approved').toLowerCase() === 'pending'
          ? 'A registration for this branch is already waiting for admin approval.'
          : 'This branch is already registered. Please log in.');
      }
    }
    if (cloud_().getSheetByName(branch)) throw new Error('This branch name is not available.');
    sh.appendRow([branch, hashPw_(branch, password), 'user', 'pending', contact, new Date().toISOString()]);
    cache.put('reg:count', String(n + 1), 3600);
    return 'Registration sent. You can log in after the admin approves it.';
  } finally {
    lock.releaseLock();
  }
}

function login(id, pw, role) {
  id = String(id || '').trim();
  pw = String(pw || '');
  role = String(role || 'user').toLowerCase();
  const cache = CacheService.getScriptCache();
  const lockKey = 'fail:' + id.toLowerCase();
  if (Number(cache.get(lockKey) || 0) >= 5) throw new Error('Too many attempts. Try again in 15 minutes.');
  const rows = users_().getDataRange().getValues();
  const hash = hashPw_(id, pw);
  let found = null;
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (String(r[0]).trim().toLowerCase() === id.toLowerCase() && String(r[1]) === hash &&
        String(r[2]).trim().toLowerCase() === role) {
      found = { id: String(r[0]).trim(), role: role, approval: String(r[3] || 'approved').trim().toLowerCase() };
      break;
    }
  }
  if (!found) {
    cache.put(lockKey, String(Number(cache.get(lockKey) || 0) + 1), 900);
    throw new Error('Wrong ID, password, or login type.');
  }
  cache.remove(lockKey);
  if (found.approval !== 'approved') throw new Error('Your registration is waiting for admin approval.');
  const token = Utilities.getUuid();
  cache.put('s:' + token, JSON.stringify({ id: found.id, role: found.role }), CONFIG.SESSION_SECONDS);
  return { token: token, id: found.id, role: found.role };
}

function logout(token) {
  CacheService.getScriptCache().remove('s:' + token);
  return true;
}

/* ================= Admin password recovery ================= */

const RESET_SECONDS = 900; // 15 minutes
const RESET_MAX_REQUESTS = 3;

// Public endpoint: it only issues a reset link when the supplied ID belongs to an Admin
// and the supplied email matches that Admin's stored recovery email. The response is intentionally generic.
function requestAdminPasswordReset(adminId, email) {
  adminId = String(adminId || '').trim();
  email = String(email || '').trim().toLowerCase();
  const generic = 'If the Admin ID and recovery email match, a password reset link has been sent to that email.';
  if (!adminId || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return generic;

  const cache = CacheService.getScriptCache();
  const rateKey = 'resetreq:' + adminId.toLowerCase();
  const count = Number(cache.get(rateKey) || 0);
  if (count >= RESET_MAX_REQUESTS) return generic;

  const rows = users_().getDataRange().getValues();
  let adminRow = -1, storedEmail = '';
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim().toLowerCase() === adminId.toLowerCase() &&
        String(rows[i][2]).trim().toLowerCase() === 'admin' &&
        String(rows[i][3] || 'approved').trim().toLowerCase() === 'approved') {
      adminRow = i;
      storedEmail = String(rows[i][4] || '').trim().toLowerCase();
      break;
    }
  }
  if (adminRow < 0 || !storedEmail || storedEmail !== email) return generic;

  const token = Utilities.getUuid() + Utilities.getUuid().replace(/-/g, '');
  cache.put('reset:' + token, JSON.stringify({ id: String(rows[adminRow][0]).trim(), email: storedEmail }), RESET_SECONDS);
  cache.put(rateKey, String(count + 1), 3600);

  const url = ScriptApp.getService().getUrl();
  if (!url) throw new Error('The web app URL is not available. Deploy the project as a web app first.');
  const resetUrl = url + '?page=reset&token=' + encodeURIComponent(token);
  const subject = 'Insurance Refill Portal - Admin password reset';
  const body = 'A password reset was requested for your Insurance Refill Portal Admin account.\n\n' +
    'Use this link within 15 minutes to create a new password/PIN:\n' + resetUrl + '\n\n' +
    'If you did not request this, ignore this email. The link expires automatically.';
  MailApp.sendEmail({to: storedEmail, subject: subject, body: body});
  return generic;
}

// Called by the reset page. The token is single-use and expires after 15 minutes.
function resetAdminPassword(token, newPassword) {
  token = String(token || '').trim();
  if (!/^[A-Za-z0-9-]+$/.test(token)) throw new Error('This reset link is invalid or has expired. Request a new link.');
  newPassword = String(newPassword || '');
  if (newPassword.length < 6) throw new Error('The new password/PIN must have at least 6 characters.');
  const cache = CacheService.getScriptCache();
  const key = 'reset:' + token;
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const raw = cache.get(key);
    if (!raw) throw new Error('This reset link is invalid or has expired. Request a new link.');
    const info = JSON.parse(raw);
    const sh = users_();
    const rows = sh.getDataRange().getValues();
    let rowNo = -1;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim().toLowerCase() === String(info.id).trim().toLowerCase() &&
          String(rows[i][2]).trim().toLowerCase() === 'admin' &&
          String(rows[i][4] || '').trim().toLowerCase() === String(info.email).trim().toLowerCase()) {
        rowNo = i + 1;
        break;
      }
    }
    if (rowNo < 0) throw new Error('The Admin account could not be verified. Request a new link.');
    sh.getRange(rowNo, 2).setNumberFormat('@').setValue(hashPw_(info.id, newPassword));
    cache.remove(key);
    return 'Password reset successfully. You can now log in with the new password/PIN.';
  } finally {
    lock.releaseLock();
  }
}

function getHome(token, branch) {
  const s = sess_(token);
  const out = { role: s.role, id: s.id, branches: [], pending: [], branch: '', hasData: false, rows: 0,
                firstInvoice: '', lastInvoice: '', problem: '' };
  if (s.role === 'admin') {
    out.branches = branchList_();
    out.pending = pending_();
    if (!branch) return out;
  }
  const b = branchOf_(token, branch);
  out.branch = b;
  let rows = [];
  try { rows = readBranch_(b); } catch (e) { out.problem = e.message; return out; }
  out.rows = rows.length;
  out.hasData = rows.length > 0;
  out.firstInvoice = rows.reduce(function (m, r) { return !m || r[I.DATE] < m ? r[I.DATE] : m; }, '');
  out.lastInvoice = lastDate_(rows);
  return out;
}

function reviewRegistration(token, branch, action) {
  if (sess_(token).role !== 'admin') throw new Error('Only an admin can review registrations.');
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sh = users_();
    const rows = sh.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === String(branch).trim() && String(rows[i][3]).toLowerCase() === 'pending') {
        if (action === 'approve') sh.getRange(i + 1, 4).setValue('approved');
        else sh.deleteRow(i + 1);
        return true;
      }
    }
    throw new Error('This registration was not found. Refresh the page.');
  } finally {
    lock.releaseLock();
  }
}

/* ================= upload, clean, calculate, save ================= */

function saveUpload(token, branch, files, keepMilk) {
  const b = branchOf_(token, branch);
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const saved = readBranch_(b);
    const previousLast = lastDate_(saved);
    const done = doneMap_(saved), mods = modMap_(saved);

    // 1. append the uploaded files; the same purchase line (phone + item + date + time) is kept once
    let incoming = [], rawCount = 0;
    (files || []).forEach(function (f) {
      const rows = cleanIn_(f.rows);
      rawCount += rows.length;
      incoming = mergeLines_(incoming, rows);
    });
    if (!incoming.length) throw new Error('No valid sales rows were found in the uploaded file.');

    // 2. dates older than the saved data: not saved
    const newMax = incoming.reduce(function (m, r) { return r[I.DATE] > m ? r[I.DATE] : m; }, '');
    if (previousLast && newMax <= previousLast) {
      throw new Error('These dates are older than the data already saved (last invoice date saved previously: ' +
        previousLast + '). The new file cannot be saved.');
    }

    // 3. clean with the master file. Saved lines stay when their item is a master or milk item.
    //    New lines stay when their item is a master item, or a milk item while "save milk items" is checked.
    const master = masterCodes_(), milk = milkCodes_();
    const gone = {};
    let removed = 0, milkKept = 0;
    const drop = function (r) {
      const g = gone[r[I.CODE]] || (gone[r[I.CODE]] = { c: r[I.CODE], d: r[I.DESC], n: 0 });
      g.n++;
      removed++;
    };
    const base = saved.map(base_).filter(function (r) {
      const ok = master[r[I.CODE]] || milk[r[I.CODE]];
      if (!ok) drop(r);
      return ok;
    });
    const fresh = incoming.filter(function (r) {
      if (master[r[I.CODE]]) return true;
      if (keepMilk && milk[r[I.CODE]]) { milkKept++; return true; }
      drop(r);
      return false;
    });

    // 4. merge: a line already saved is replaced, any other line is added
    const merged = mergeLines_(base, fresh);
    const duplicates = (rawCount - incoming.length) + (base.length + fresh.length - merged.length);
    const removedCodes = Object.keys(gone).map(function (k) { return gone[k]; })
      .sort(function (a, c) { return c.n - a.n; }).slice(0, 100)
      .map(function (g) { return [g.c, g.d, g.n]; });

    // 5. per purchase: average qt, net value sum, due date (milk rule or average qt rule)
    const finalRows = calc_(merged, done, mods, milk);
    writeBranch_(b, finalRows);
    applyAutomaticRefillCompletion_(b, finalRows);
    if (b === 'P0661' && typeof firestoreSyncAffectedP0661_ === 'function') firestoreSyncAffectedP0661_(finalRows);

    return { branch: b, previousLast: previousLast, newLast: lastDate_(finalRows), rowsSaved: finalRows.length,
             removed: removed, removedCodes: removedCodes, duplicates: duplicates, milkKept: milkKept };
  } finally {
    lock.releaseLock();
  }
}

function deletePeriod(token, branch, from, to, dryRun) {
  const b = branchOf_(token, branch);
  if (!isDate_(from) || !isDate_(to) || from > to) throw new Error('Choose a valid period: From must not be after To.');
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const rows = readBranch_(b);
    const inside = function (r) { return r[I.DATE] >= from && r[I.DATE] <= to; };
    const gone = rows.filter(inside);
    const purchases = {};
    gone.forEach(function (r) { purchases[pkey_(r)] = 1; });
    const nPurchases = Object.keys(purchases).length;
    if (dryRun) return { count: gone.length, purchases: nPurchases, remaining: rows.length - gone.length };
    if (!gone.length) throw new Error('No saved rows fall in this period.');
    const keep = rows.filter(function (r) { return !inside(r); }).map(base_);
    const finalRows = calc_(keep, doneMap_(rows), modMap_(rows), milkCodes_());
    writeBranch_(b, finalRows);
    if (b === 'P0661' && typeof firestoreSyncAffectedP0661_ === 'function') firestoreSyncAffectedP0661_(finalRows);
    return { deleted: gone.length, purchases: nPurchases, remaining: finalRows.length, lastInvoice: lastDate_(finalRows) };
  } finally {
    lock.releaseLock();
  }
}

/* ================= explore, status, modified due date, export ================= */

function explore(token, branch, date, min, max, ym) {
  const b = branchOf_(token, branch);
  let exploreRows, follow;

  // FAST PATH: for the approved P0661 test branch, Firestore is the read model.
  // Do not call readBranch_(), migrateBranchState_(), applyAutomaticRefillCompletion_(),
  // syncFollowStatuses_(), or followMap_() during an ordinary Explore request.
  // Those functions touch Sheets and were the main source of repeated latency.
  const firestoreReady = b === 'P0661' &&
    typeof firestoreReadExploreSnapshotP0661_ === 'function' &&
    PropertiesService.getScriptProperties().getProperty('FIRESTORE_INDEX_P0661_READY') === '1';

  if (firestoreReady) {
    const snapshot = firestoreReadExploreSnapshotP0661_();
    exploreRows = snapshot.rows;
    follow = snapshot.follow;
  } else {
    // Compatibility path for other branches, or P0661 before its first index build.
    const rows = readBranch_(b);
    applyAutomaticRefillCompletion_(b, rows);
    const refreshedRows = readBranch_(b);
    syncFollowStatuses_(b, refreshedRows);
    exploreRows = refreshedRows;
    follow = followMap_(b);
  }

  const lo = (min === '' || min == null) ? -Infinity : Number(min);
  const hi = (max === '' || max == null) ? Infinity : Number(max);
  const mon = String(ym || date || '').slice(0, 7);
  const recs = date ? follow : {};
  const perDay = {}, list = [];
  exploreRows.forEach(function (r) {
    const due = eff_(r);
    if (!due || r[I.SUM] < lo || r[I.SUM] > hi) return;
    if (due.slice(0, 7) === mon && r[I.ST] !== 'complete') (perDay[due] = perDay[due] || {})[r[I.PHONE]] = 1;
    // Explore shows every transaction due on the selected day, including completed ones.
    // Calendar badges remain based on incomplete customers only.
    if (date && due === date) list.push(r);
  });
  const counts = {};
  Object.keys(perDay).forEach(function (k) { counts[k] = Object.keys(perDay[k]).length; });
  list.sort(function (a, c) {
    return a[I.NAME].localeCompare(c[I.NAME]) || a[I.PHONE].localeCompare(c[I.PHONE]) ||
           a[I.CODE].localeCompare(c[I.CODE]) || a[I.TIME].localeCompare(c[I.TIME]);
  });
  return { counts: counts, rows: list.map(function (r) { return detail_(r, recs); }) };
}

// code '*' marks every purchase line of the customer due on that date (conclusion page);
// otherwise only the line with that item code, purchase date and time (details page).
// `due` is the due date as shown (the modified one when there is one).
function upsertFollowState_(b, phone, due, tdate, ttime, code, membershipId, patch) {
  const sh=followSheet_(), last=sh.getLastRow();
  const vals=last>1?sh.getRange(2,1,last-1,FOLLOW_COLS).getValues():[];
  let ix=-1;
  for(let i=0;i<vals.length;i++) if(s_(vals[i][0])===b && s_(vals[i][1])===phone && s_(vals[i][3])===tdate && s_(vals[i][4])===ttime && s_(vals[i][15])===String(code||'') && s_(vals[i][6])===String(membershipId||'')){ix=i;break;}
  if(ix<0) for(let i=0;i<vals.length;i++) if(s_(vals[i][0])===b && s_(vals[i][1])===phone && s_(vals[i][3])===tdate && s_(vals[i][4])===ttime && s_(vals[i][15])===String(code||'') && !s_(vals[i][6])){ix=i;break;}
  const now=Utilities.formatDate(new Date(),CONFIG.TZ,"yyyy-MM-dd'T'HH:mm:ss");
  if(ix<0){ vals.push([b,phone,due,tdate,ttime,'',String(membershipId||''),'','','','', 'incomplete','', '', '',String(code||'')]); ix=vals.length-1; }
  else if(membershipId && !s_(vals[ix][6])) vals[ix][6]=String(membershipId);
  vals[ix][2]=due;
  Object.keys(patch||{}).forEach(function(k){ const col={status:11,su:12,mod:13,md:14}[k]; if(col!==undefined) vals[ix][col]=patch[k]; });
  if(patch && ('status' in patch)) vals[ix][12]=now;
  if(patch && ('mod' in patch)) vals[ix][14]=now;
  sh.getRange(2,1,vals.length,FOLLOW_COLS).setNumberFormat('@').setValues(vals);
}

function setStatus(token, branch, phone, due, code, tdate, ttime, complete) {
  const b=branchOf_(token,branch), rows=readBranch_(b);
  const targets=rows.filter(function(r){return r[I.PHONE]===phone && eff_(r)===due && (code==='*'||(r[I.CODE]===code&&r[I.DATE]===tdate&&r[I.TIME]===ttime));});
  if(!targets.length) throw new Error('This row was not found. Explore again and retry.');
  targets.forEach(function(r){ upsertFollowState_(b,phone,due,r[I.DATE],r[I.TIME],r[I.CODE],r[I.MID],{status:complete?'complete':'incomplete'}); });
  // Preserve the established automatic replacement rule at write time, rather than
  // depending on the next Explore visit to repair an older purchase line.
  const auto = applyAutomaticRefillCompletion_(b, rows);
  if (b === 'P0661' && typeof firestoreSyncAffectedP0661_ === 'function') {
    const affected = targets.concat(auto && auto.rows ? auto.rows : []);
    firestoreSyncAffectedP0661_(affected);
  }
  return targets.length;
}

function setModifiedDue(token, branch, phone, code, tdate, ttime, newDate) {
  const b=branchOf_(token,branch); newDate=String(newDate||'');
  if(newDate&&!isDate_(newDate)) throw new Error('Choose a valid date.');
  const rows=readBranch_(b), r=rows.find(function(x){return x[I.PHONE]===phone&&x[I.CODE]===code&&x[I.DATE]===tdate&&x[I.TIME]===ttime;});
  if(!r) throw new Error('This line was not found. Explore again and retry.');
  const value=newDate===r[I.DUE]?'':newDate, effective=value||r[I.DUE];
  upsertFollowState_(b,phone,effective,tdate,ttime,code,r[I.MID],{mod:value});
  if (b === 'P0661' && typeof firestoreSyncAffectedP0661_ === 'function') firestoreSyncAffectedP0661_([r]);
  return effective;
}

function setGroupDue(token, branch, phone, due, newDate) {
  const b=branchOf_(token,branch); newDate=String(newDate||'');
  if(newDate&&!isDate_(newDate)) throw new Error('Choose a valid date.');
  const rows=readBranch_(b).filter(function(r){return r[I.PHONE]===phone&&eff_(r)===due;});
  if(!rows.length) throw new Error('This customer is no longer due on that date. Explore again.');
  rows.forEach(function(r){const value=newDate===r[I.DUE]?'':newDate, effective=value||r[I.DUE]; upsertFollowState_(b,phone,effective,r[I.DATE],r[I.TIME],r[I.CODE],r[I.MID],{mod:value});});
  if (b === 'P0661' && typeof firestoreSyncAffectedP0661_ === 'function') firestoreSyncAffectedP0661_(rows);
  return rows.length;
}

// view: 'summary' (conclusion page, one row per customer and due date) or 'details' (one row per item)
function exportDue(token, branch, from, to, status, min, max, view, fmt, selection) {
  const b = branchOf_(token, branch);
  if (!isDate_(from) || !isDate_(to) || from > to) throw new Error('Choose a valid period: From must not be after To.');
  const rows = readBranch_(b);
  syncFollowStatuses_(b, rows);
  const follow = followMap_(b);
  const lo = (min === '' || min == null) ? -Infinity : Number(min);
  const hi = (max === '' || max == null) ? Infinity : Number(max);
  const list = rows.filter(function (r) {
    const due = eff_(r);
    if (!due || due < from || due > to || r[I.SUM] < lo || r[I.SUM] > hi) return false;
    const isDone = r[I.ST] === 'complete';
    if (!(status === 'all' || (status === 'complete' ? isDone : !isDone))) return false;
    if (selection && selection !== 'all') {
      const f = follow[fkey_(r)];
      if (!f || f.a !== Number(selection)) return false;
    }
    return true;
  });
  if (!list.length) throw new Error('No customers match this selection, so there is nothing to export.');
  list.sort(function (a, c) {
    return eff_(a).localeCompare(eff_(c)) || a[I.NAME].localeCompare(c[I.NAME]) ||
           a[I.PHONE].localeCompare(c[I.PHONE]) || a[I.CODE].localeCompare(c[I.CODE]);
  });

  const label = function (done) { return done ? 'Complete' : 'Incomplete'; };
  let heads, data, numCol;
  if (view === 'details') {
    heads = ['Loyalty Name', 'Loyalty Phone', 'Item Code', 'Item Description', 'Quantity', 'Membership ID',
             'Last Purchase Day', 'Due Date', 'Status'];
    data = list.map(function (r) {
      return [r[I.NAME], r[I.PHONE], r[I.CODE], r[I.DESC], r[I.QTY], r[I.MID], r[I.LAST], eff_(r), label(r[I.ST] === 'complete')];
    });
    numCol = 5;
  } else {
    heads = ['Loyalty Name', 'Loyalty Phone', 'Membership ID', 'Last Purchase Day', 'Due Date', 'Sum Pre-Net Value',
             'Actions', 'Notes', 'Status'];
    data = groupRows_(list, followMap_(b)).map(function (g) {
      return [g.name, g.phone, g.ids.join(', '), g.last, g.due, Math.round(g.net * 100) / 100,
              g.act ? CONFIG.ACTIONS[g.act - 1] : '', g.note, label(g.done)];
    });
    numCol = 6;
  }

  const tmp = SpreadsheetApp.create('export-' + Date.now());
  const id = tmp.getId();
  try {
    const sh = tmp.getSheets()[0];
    for (let c = 1; c <= heads.length; c++) {
      if (c !== numCol) sh.getRange(2, c, data.length, 1).setNumberFormat('@');
    }
    sh.getRange(1, 1, 1, heads.length).setValues([heads]);
    sh.getRange(2, 1, data.length, heads.length).setValues(data);
    sh.getRange(1, 1, 1, heads.length).setFontWeight('bold').setBackground('#0d5c4d').setFontColor('#ffffff');
    sh.getRange(1, 1, data.length + 1, heads.length)
      .setBorder(true, true, true, true, true, true, '#9fb3ab', SpreadsheetApp.BorderStyle.SOLID);
    sh.autoResizeColumns(1, heads.length);
    SpreadsheetApp.flush();
    const pdf = fmt === 'pdf';
    const url = 'https://docs.google.com/spreadsheets/d/' + id + '/export?format=' + (pdf ?
      'pdf&portrait=false&size=A4&fitw=true&gridlines=false&sheetnames=false&printtitle=false' : 'xlsx');
    const res = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) throw new Error('Export failed (' + res.getResponseCode() + ').');
    return {
      name: String(b).replace(/[^\w\-]+/g, '_') + '_due_' + from + '_to_' + to + '_' + status + '_' +
            (view === 'details' ? 'details' : 'conclusion') + (pdf ? '.pdf' : '.xlsx'),
      mime: pdf ? 'application/pdf' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      b64: Utilities.base64Encode(res.getBlob().getBytes())
    };
  } finally {
    try { DriveApp.getFileById(id).setTrashed(true); } catch (e) {}
  }
}

/* ================= monthly analytics ================= */

function monthlyAnalytics(token, branch, ym) {
  const b = branchOf_(token, branch);
  if (!/^\d{4}-\d{2}$/.test(String(ym || ''))) throw new Error('Choose a valid month.');
  const rows = readBranch_(b);
  syncFollowStatuses_(b, rows);
  const follow = followMap_(b);
  const days = new Date(Number(ym.slice(0,4)), Number(ym.slice(5,7)), 0).getDate();
  const byDay = {};
  for (let d = 1; d <= days; d++) byDay[ym + '-' + ('0' + d).slice(-2)] = {planned:0, complete:0, incomplete:0, contacted:0, notContacted:0, action:[]};
  const groups = {};
  rows.forEach(function(r){
    const due = eff_(r);
    if (!due || due.slice(0,7) !== ym || !r[I.PHONE]) return;
    const k = r[I.PHONE] + '|' + due;
    let g = groups[k];
    if (!g) g = groups[k] = {phone:r[I.PHONE], name:r[I.NAME], due:due, net:0, complete:true, action:0};
    g.net += Number(r[I.NET]) || 0;
    if (r[I.ST] !== 'complete') g.complete = false;
    const f = follow[fkey_(r)];
    if (f && f.a) g.action = f.a;
  });
  Object.keys(groups).forEach(function(k){
    const g = groups[k], d = byDay[g.due];
    if (!d) return;
    d.planned++;
    if (g.complete) d.complete++; else d.incomplete++;
    if (g.action) d.contacted++;
    if (g.action === 2) d.notContacted++;
    d.action[g.action || 0] = (d.action[g.action || 0] || 0) + 1;
  });
  let plannedCum=0, completeCum=0;
  const progress = Object.keys(byDay).map(function(day){
    plannedCum += byDay[day].planned;
    completeCum += byDay[day].complete;
    return {day:day, planned:byDay[day].planned, complete:byDay[day].complete, incomplete:byDay[day].incomplete,
            plannedCum:plannedCum, completeCum:completeCum, contacted:byDay[day].contacted, notContacted:byDay[day].notContacted};
  });
  const groupKeys = Object.keys(groups);
  const total = groupKeys.length;
  const complete = groupKeys.filter(k=>groups[k].complete).length;
  const totalValue = groupKeys.reduce(function(sum, k){ return sum + (Number(groups[k].net) || 0); }, 0);
  // Approved means the explicit Customer approved refill action (action 4).
  const approved = groupKeys.filter(k=>groups[k].action === 4).length;
  const approvedValue = groupKeys.reduce(function(sum, k){ return sum + (groups[k].action === 4 ? (Number(groups[k].net) || 0) : 0); }, 0);
  const actions = {};
  groupKeys.forEach(function(k){ actions[groups[k].action || 0] = (actions[groups[k].action || 0] || 0) + 1; });
  return {month:ym, total:total, complete:complete, incomplete:total-complete,
          approved:approved, totalValue:Math.round(totalValue * 100) / 100,
          approvedValue:Math.round(approvedValue * 100) / 100,
          contacted:(actions[3]||0)+(actions[4]||0)+(actions[5]||0), notContacted:actions[2]||0,
          actions:actions, progress:progress, labels:CONFIG.ACTIONS};
}

/* ================= customer follow up ================= */

function purchaseStatus_(rows, phone, due, tdate, ttime) {
  let found = false, allDone = true;
  rows.forEach(function (r) {
    // Invoice completion is evaluated across every item in the same transaction, independent of due-date edits.
    if (r[I.PHONE] !== phone || r[I.DATE] !== tdate || r[I.TIME] !== ttime) return;
    found = true;
    if (r[I.ST] !== 'complete') allDone = false;
  });
  return found && allDone ? 'complete' : 'incomplete';
}

// Backfill the permanent follow-up status for existing rows that predate the Status column.
function syncFollowStatuses_(b, rows) {
  const sh = cloud_().getSheetByName(CONFIG.FOLLOW_SHEET);
  if (!sh || sh.getLastRow() < 2) return;
  const vals = sh.getRange(2, 1, sh.getLastRow() - 1, FOLLOW_COLS).getValues();
  const now = Utilities.formatDate(new Date(), CONFIG.TZ, "yyyy-MM-dd'T'HH:mm:ss");
  let changed = false;
  vals.forEach(function (r) {
    if (s_(r[0]) !== b || s_(r[11])) return;
    const status = purchaseStatus_(rows, s_(r[1]), s_(r[2]), s_(r[3]), s_(r[4]));
    r[11] = status; r[12] = now; changed = true;
  });
  if (changed) sh.getRange(2, 1, vals.length, FOLLOW_COLS).setValues(vals);
  return changed;
}

// Saves the action (1-5, or 0 for none) and the notes of a customer for a due date.
// One row is kept per purchase process of that customer due on that date.
function saveFollow(token, branch, phone, due, actionIdx, notes) {
  const b = branchOf_(token, branch);
  const idx = Number(actionIdx) || 0;
  if (idx < 0 || idx > CONFIG.ACTIONS.length) throw new Error('Choose one of the five actions.');
  notes = String(notes == null ? '' : notes).replace(/\r/g, '').slice(0, 1000);
  if (!idx && !notes.trim()) throw new Error('Choose an action or write a note.');
  if (!isDate_(due)) throw new Error('Choose a valid due date.');
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sourceRows = readBranch_(b);
    const purch = {};
    sourceRows.forEach(function (r) {
      if (r[I.PHONE] !== phone || eff_(r) !== due) return;
      const k = r[I.DATE] + '|' + r[I.TIME];
      const p = purch[k] || (purch[k] = { date: r[I.DATE], time: r[I.TIME], name: r[I.NAME], ids: [] });
      if (r[I.MID] && p.ids.indexOf(r[I.MID]) < 0) p.ids.push(r[I.MID]);
    });
    const keys = Object.keys(purch);
    if (!keys.length) throw new Error('This customer is no longer due on that date. Explore again.');
    const sh = followSheet_();
    const last = sh.getLastRow();
    const vals = last > 1 ? sh.getRange(2, 1, last - 1, FOLLOW_COLS).getValues() : [];
    const at = {}, duplicateTransactionRows = [];
    vals.forEach(function (r, i) {
      if (s_(r[0]) !== b || s_(r[15])) return;
      const k = s_(r[1]) + '|' + s_(r[3]) + '|' + s_(r[4]);
      if (at[k]) duplicateTransactionRows.push(i + 2);
      else at[k] = i + 2;
    });
    const stamp = Utilities.formatDate(new Date(), CONFIG.TZ, "yyyy-MM-dd'T'HH:mm:ss");
    const by = sess_(token).id;
    const text = idx ? CONFIG.ACTIONS[idx - 1] : '';
    keys.forEach(function (k) {
      const p = purch[k];
      const status = purchaseStatus_(readBranch_(b), phone, due, p.date, p.time);
      const key = phone + '|' + p.date + '|' + p.time;
      const oldRow = at[key] ? vals[at[key]-2] : [];
      const row = [b, phone, due, p.date, p.time, p.name, p.ids.join(', '), text, notes, stamp, by, (oldRow[11] || status), (oldRow[12] || stamp), (oldRow[13] || ''), (oldRow[14] || ''), ''];
      const at1 = at[key] || (sh.getLastRow() + 1);
      sh.getRange(at1, 1, 1, FOLLOW_COLS).setNumberFormat('@').setValues([row]);
    });
    duplicateTransactionRows.sort(function(a,c){return c-a;}).forEach(function(rowNo){sh.deleteRow(rowNo);});
    if (b === 'P0661' && typeof firestoreSyncAffectedP0661_ === 'function') {
      // The original branch snapshot already contains every line affected by
      // this customer + due date. Reuse it; do not read the whole branch again.
      const affectedRows = sourceRows.filter(function(r){ return r[I.PHONE] === phone && eff_(r) === due; });
      firestoreSyncAffectedP0661_(affectedRows);
    }
    return keys.length;
  } finally {
    lock.releaseLock();
  }
}

const FOLLOW_COLS = 16;

// If a later purchase exists for the same Membership ID + Item Code, every older
// purchase line for that pair is permanently marked Complete. Other invoice items are untouched.
function applyAutomaticRefillCompletion_(b, rows) {
  rows = rows || [];
  const latest = {};
  rows.forEach(function(r){
    const mid=s_(r[I.MID]), code=s_(r[I.CODE]), date=s_(r[I.DATE]), time=s_(r[I.TIME]);
    if(!mid || !code || !date || !time) return;
    const key=mid+'|'+code, stamp=date+'T'+time;
    if(!latest[key] || stamp>latest[key]) latest[key]=stamp;
  });
  const sh=followSheet_(), last=sh.getLastRow();
  const vals=last>1?sh.getRange(2,1,last-1,FOLLOW_COLS).getValues():[];
  const index={};
  vals.forEach(function(r,i){
    if(s_(r[0])===b && s_(r[15])) index[s_(r[1])+'|'+s_(r[3])+'|'+s_(r[4])+'|'+s_(r[6])+'|'+s_(r[15])]=i;
  });
  const now=Utilities.formatDate(new Date(),CONFIG.TZ,"yyyy-MM-dd'T'HH:mm:ss");
  let changed=false;
  const changedRows=[];
  rows.forEach(function(r){
    const mid=s_(r[I.MID]), code=s_(r[I.CODE]), date=s_(r[I.DATE]), time=s_(r[I.TIME]);
    if(!mid || !code || !date || !time) return;
    const stamp=date+'T'+time, newest=latest[mid+'|'+code];
    if(!newest || stamp>=newest) return;
    const key=s_(r[I.PHONE])+'|'+date+'|'+time+'|'+mid+'|'+code;
    const ix=index[key];
    if(ix!==undefined){
      if(s_(vals[ix][11])!=='complete'){ vals[ix][11]='complete'; vals[ix][12]=now; changed=true; changedRows.push(r); }
      if(s_(vals[ix][2])!==eff_(r)){ vals[ix][2]=eff_(r); changed=true; }
    } else {
      const newRow=[b,s_(r[I.PHONE]),eff_(r),date,time,s_(r[I.NAME]),mid,'','','','', 'complete',now,'','',code];
      index[key]=vals.length; vals.push(newRow); changed=true;
    }
  });
  if(changed) sh.getRange(2,1,vals.length,FOLLOW_COLS).setNumberFormat('@').setValues(vals);
  return {changed:changed, rows:changedRows};
}

// follow up records are keyed by branch + phone + Trx Date + Trx Time; item-state rows add Item Code
function followMap_(b) {
  const sh=cloud_().getSheetByName(CONFIG.FOLLOW_SHEET), m={};
  if(!sh||sh.getLastRow()<2) return m;
  const width=Math.max(13,sh.getLastColumn());
  const rows=sh.getRange(2,1,sh.getLastRow()-1,Math.min(width,FOLLOW_COLS)).getValues();
  const base={};
  rows.forEach(function(r){
    if(s_(r[0])!==b || s_(r[15])) return;
    const k=s_(r[1])+'|'+s_(r[3])+'|'+s_(r[4]);
    base[k]={a:CONFIG.ACTIONS.indexOf(s_(r[7]))+1,n:s_(r[8]),u:s_(r[9]),status:s_(r[11])||'incomplete',su:s_(r[12]),mod:s_(r[13]),md:s_(r[14])};
  });
  Object.keys(base).forEach(function(k){m[k]=base[k];});
  rows.forEach(function(r){
    if(s_(r[0])!==b || !s_(r[15])) return;
    const k=s_(r[1])+'|'+s_(r[3])+'|'+s_(r[4]), parent=base[k]||{};
    const entry={a:CONFIG.ACTIONS.indexOf(s_(r[7]))+1,n:s_(r[8]),u:s_(r[9]),status:s_(r[11])||'incomplete',su:s_(r[12]),mod:s_(r[13]),md:s_(r[14])};
    const lineEntry=Object.assign({},parent,entry,{a:entry.a||parent.a||0,n:entry.n||parent.n||'',u:entry.u||parent.u||''});
    m[k+'|'+s_(r[6])+'|'+s_(r[15])]=lineEntry;
    if(!s_(r[6])) m[k+'|'+s_(r[15])]=lineEntry;
  });
  return m;
}

function followSheet_() {
  const ss=cloud_(), sh=ss.getSheetByName(CONFIG.FOLLOW_SHEET)||ss.insertSheet(CONFIG.FOLLOW_SHEET);
  const H=['Branch','Loyalty Phone','Due Date','Trx Date','Trx Time','Loyalty Name','Membership ID','Action','Notes','Updated','Updated By','Status','Status Updated','Modified Due Date','Modified Date','Item Code'];
  if(sh.getLastRow()===0){sh.getRange(1,1,1,H.length).setValues([H]).setFontWeight('bold');}
  else {
    const old=sh.getRange(1,1,1,Math.max(13,sh.getLastColumn())).getValues()[0].map(String);
    H.forEach(function(h,i){if(old[i]!==h) sh.getRange(1,i+1).setValue(h);});
  }
  sh.getRange(1,1,Math.max(1,sh.getLastRow()),H.length).setNumberFormat('@'); sh.setFrozenRows(1); return sh;
}

// transaction identity ignores Due Date: changing a due date must not create another transaction record
function fkey_(r) { return r[I.PHONE] + '|' + r[I.DATE] + '|' + r[I.TIME]; }
function flinekey_(r) { return fkey_(r) + '|' + r[I.MID] + '|' + r[I.CODE]; }

/* ================= search by Membership ID or mobile number ================= */

function checkQuery_(query) {
  const q = String(query == null ? '' : query).trim();
  if (q.length < 4) throw new Error('Type at least 4 characters of the Membership ID or the mobile number.');
  return q;
}

// Finds the customers (loyalty phones) that match, whatever their due dates are, then returns ALL saved lines
// of those customers. A Membership ID matches exactly (or by its digits once 6 or more are typed); a mobile
// number matches by its ending, so 0541611096 finds 966541611096.
function searchRows_(rows, q) {
  const ph = q.replace(/\D/g, '').replace(/^0+/, '');
  const phones = {};
  rows.forEach(function (r) {
    if (!r[I.PHONE]) return;
    const mid = r[I.MID];
    const byId = mid && (mid === q || (q.length >= 6 && mid.indexOf(q) >= 0));
    const byPhone = ph.length >= 6 && r[I.PHONE].slice(-ph.length) === ph;
    if (byId || byPhone) phones[r[I.PHONE]] = 1;
  });
  const n = Object.keys(phones).length;
  if (n > 50) throw new Error('More than 50 customers match. Type more of the number.');
  return { list: rows.filter(function (r) { return phones[r[I.PHONE]] && eff_(r); }), members: n };
}

// detail rows (same layout as explore) of the matching customers for all due dates
function searchMember(token, branch, query) {
  const b = branchOf_(token, branch);
  const found = searchRows_(readBranch_(b), checkQuery_(query));
  const recs = followMap_(b);
  found.list.sort(function (a, c) {
    return a[I.NAME].localeCompare(c[I.NAME]) || a[I.PHONE].localeCompare(c[I.PHONE]) ||
           eff_(a).localeCompare(eff_(c)) || a[I.CODE].localeCompare(c[I.CODE]) || a[I.TIME].localeCompare(c[I.TIME]);
  });
  return { members: found.members, rows: found.list.map(function (r) { return detail_(r, recs); }) };
}

// Full purchase history for one exact Membership ID, including completed lines.
// Invoice groups are keyed by phone + transaction date + transaction time; due-date edits never split a transaction.
function getMemberHistory(token, branch, membershipId) {
  const b = branchOf_(token, branch);
  const mid = String(membershipId == null ? '' : membershipId).trim();
  if (!mid) throw new Error('Membership ID is required.');
  if (b === 'P0661' && typeof firestoreGetMemberHistoryP0661_ === 'function' &&
      PropertiesService.getScriptProperties().getProperty('FIRESTORE_INDEX_P0661_READY') === '1') {
    return firestoreGetMemberHistoryP0661_(mid);
  }
  let rows = readBranch_(b);
  applyAutomaticRefillCompletion_(b, rows);
  rows = readBranch_(b);
  const recs = followMap_(b), grouped = {}, order = [];
  rows.filter(function(r){ return s_(r[I.MID]) === mid; }).forEach(function(r){
    const k = r[I.PHONE] + '|' + r[I.DATE] + '|' + r[I.TIME];
    if (!grouped[k]) {
      const f = recs[r[I.PHONE] + '|' + r[I.DATE] + '|' + r[I.TIME]] || {};
      grouped[k] = { membershipId: mid, phone: r[I.PHONE], name: r[I.NAME], trxDate: r[I.DATE], trxTime: r[I.TIME], action: f.a ? CONFIG.ACTIONS[f.a-1] : '', notes: f.n || '', items: [] };
      order.push(k);
    }
    grouped[k].items.push({ code:r[I.CODE], description:r[I.DESC], quantity:r[I.QTY], calculatedDueDate:r[I.DUE], modifiedDueDate:r[I.MOD] || '', dueDate:eff_(r), status:r[I.ST] === 'complete' ? 'Complete' : 'Incomplete' });
  });
  order.sort(function(a,c){ return a.split('|')[1]===c.split('|')[1] ? a.split('|')[2].localeCompare(c.split('|')[2]) : a.split('|')[1].localeCompare(c.split('|')[1]); });
  return { membershipId:mid, invoices:order.map(function(k){ return grouped[k]; }) };
}

// one saved line as a detail row: name, phone, item code, item description, quantity, Membership ID,
// purchase date, due date shown, complete?, time, pre-net value, modified?, calculated due date,
// follow up action, notes, saved at, permanently mirrored status, status updated at
function detail_(r, recs) {
  const f = recs[flinekey_(r)] || recs[fkey_(r)+'|'+r[I.CODE]] || recs[fkey_(r)] || { a: 0, n: '', u: '', status: r[I.ST] === 'complete' ? 'complete' : 'incomplete', su: '' };
  return [r[I.NAME], r[I.PHONE], r[I.CODE], r[I.DESC], r[I.QTY], r[I.MID], r[I.LAST], eff_(r),
          f.status === 'complete', r[I.TIME], r[I.NET], !!f.mod, r[I.DUE], f.a, f.n, f.u, f.status, f.su];
}

/* ================= data helpers ================= */

function key_(v) { return String(v).replace(/^0+/, '') || '0'; }

// who a purchase belongs to: the loyalty phone (the receipt number only when a row has no phone)
function who_(r) { return r[I.PHONE] || ('R' + key_(r[I.RC])); }

// one purchase = one customer at one date and time
function pkey_(r) { return who_(r) + '|' + r[I.DATE] + '|' + r[I.TIME]; }

// one purchase line = one customer, one item, at one date and time
function lkey_(r) { return who_(r) + '|' + r[I.CODE] + '|' + r[I.DATE] + '|' + r[I.TIME]; }

// the due date shown: the user's modified date when there is one, otherwise the calculated one
function eff_(r) { return r[I.MOD] || r[I.DUE]; }

// the same purchase line appears once; when it appears again the newer row replaces the older one
function mergeLines_(base, add) {
  const m = new Map();
  base.concat(add).forEach(function (r) { m.set(lkey_(r), r); });
  return Array.from(m.values());
}

function cleanIn_(rows) {
  const out = [];
  (rows || []).forEach(function (r) {
    const rc = String(r[I.RC] == null ? '' : r[I.RC]).trim();
    const dt = String(r[I.DATE] || '');
    if (!rc || !isDate_(dt)) return;
    const tm = /^\d{2}:\d{2}:\d{2}$/.test(String(r[I.TIME])) ? String(r[I.TIME]) : '';
    const q = Number(r[I.QTY]), n = Number(r[I.NET]);
    out.push([rc, String(r[I.MON] || ''), dt, tm, String(r[I.NAME] || '').trim(), String(r[I.PHONE] || '').replace(/\D/g, ''),
              String(r[I.CODE] == null ? '' : r[I.CODE]).trim(), String(r[I.DESC] || '').trim(),
              isNaN(q) ? 0 : q, isNaN(n) ? 0 : n, String(r[IN_MID] == null ? '' : r[IN_MID]).trim()]);
  });
  return out;
}

// saved row -> the layout used while merging (the first 10 fields plus the Membership ID)
function base_(r) { return r.slice(0, 10).concat([r[I.MID]]); }

// a completed line stays completed as long as the same purchase line exists
function doneMap_(rows) {
  const m = {};
  rows.forEach(function (r) { if (r[I.ST] === 'complete') m[lkey_(r)] = true; });
  return m;
}

// a modified due date stays with its purchase line
function modMap_(rows) {
  const m = {};
  rows.forEach(function (r) { if (r[I.MOD]) m[lkey_(r)] = r[I.MOD]; });
  return m;
}

// milk rule: 10 or more -> 28 days; below 10 -> quantity x 2.8 days (rounded to the nearest day)
function milkDays_(qty) {
  return qty >= 10 ? 28 : Math.max(0, Math.round(qty * 2.8));
}

function calc_(rows, done, mods, milk) {
  done = done || {}; mods = mods || {}; milk = milk || {};
  const g = {};
  rows.forEach(function (r) {
    const o = g[pkey_(r)] || (g[pkey_(r)] = { s: 0, t: 0, n: 0, mq: 0, milk: false });
    o.s += r[I.QTY];
    o.t += r[I.NET];
    o.n++;
    if (milk[r[I.CODE]]) { o.milk = true; o.mq += r[I.QTY]; }
  });
  return rows.map(function (r) {
    const o = g[pkey_(r)];
    const avg = Math.round(o.s / o.n * 1e6) / 1e6;
    const own = r[I.PHONE] ? r[I.DATE] : '';           // Last Purchase Day = the purchase's own date
    let due = '';
    if (own) due = addDays_(own, o.milk ? milkDays_(o.mq) : (avg <= 1.5 ? 28 : (avg <= 2.5 ? 56 : 84)));
    const k = lkey_(r);
    return r.slice(0, 10).concat([Math.round(avg * 10000) / 10000, own, due, Math.round(o.t * 100) / 100,
      r[IN_MID] || '', done[k] ? 'complete' : '', mods[k] || '']);
  });
}

function groupRows_(list, recs) {
  const g = {}, order = [];
  list.forEach(function (r) {
    const due = eff_(r);
    const k = r[I.PHONE] + '|' + due;
    let o = g[k];
    if (!o) { o = g[k] = { name: r[I.NAME], phone: r[I.PHONE], ids: [], last: r[I.LAST], due: due, net: 0, done: true, act: 0, note: '', upd: '' }; order.push(k); }
    if (r[I.MID] && o.ids.indexOf(r[I.MID]) < 0) o.ids.push(r[I.MID]);
    if (r[I.LAST] > o.last) o.last = r[I.LAST];
    o.net += r[I.NET];
    if (r[I.ST] !== 'complete') o.done = false;
    const f = recs && recs[fkey_(r)];
    if (f && f.u > o.upd) { o.act = f.a; o.note = f.n; o.upd = f.u; }
  });
  return order.map(function (k) { return g[k]; });
}

/* ================= item code lists (master file, milk items file) ================= */

function masterCodes_() { return codesFrom_(CONFIG.MASTER_ID, CONFIG.MASTER_SHEET, 'master file', 'MASTER_ID'); }
function milkCodes_() { return codesFrom_(CONFIG.MILK_ID, CONFIG.MILK_SHEET, 'milk items file', 'MILK_ID'); }

// reads the item codes from column A; the first row is skipped when it is the header ("Code" / "Item code")
function codesFrom_(fileId, sheetName, what, configKey) {
  let ss = null, tmpId = null;
  try { ss = SpreadsheetApp.openById(fileId); } catch (e) { ss = null; }
  if (!ss) {
    // the file may be an Excel file stored in Drive: convert a temporary copy
    try {
      const copy = Drive.Files.copy({ name: 'tmp-codes-' + Date.now(), mimeType: MimeType.GOOGLE_SHEETS }, fileId);
      tmpId = copy.id;
      ss = SpreadsheetApp.openById(tmpId);
    } catch (e) {
      throw new Error('Cannot read the ' + what + '. Save it as a Google Sheet (File > Save as Google Sheets) ' +
        'and put the new ID in CONFIG.' + configKey + '. Details: ' + e.message);
    }
  }
  try {
    const sh = sheetName ? ss.getSheetByName(sheetName) : ss.getSheets()[0];
    const v = sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 1).getValues();
    const set = {};
    v.forEach(function (r, i) {
      const c = String(r[0]).trim();
      if (c && !(i === 0 && /^(item\s*)?code$/i.test(c))) set[c] = true;
    });
    return set;
  } finally {
    if (tmpId) { try { DriveApp.getFileById(tmpId).setTrashed(true); } catch (e) {} }
  }
}

/* ================= sheet access and misc helpers ================= */

function cloud_() { return SpreadsheetApp.openById(CONFIG.CLOUD_ID); }

function users_() {
  const sh = cloud_().getSheetByName(CONFIG.USERS_SHEET);
  if (!sh) throw new Error('The Users sheet is missing. Run setupUsersSheet() once from the script editor.');
  return sh;
}

function sess_(token) {
  const v = CacheService.getScriptCache().get('s:' + token);
  if (!v) throw new Error('Session expired. Please log in again.');
  return JSON.parse(v);
}

function branchOf_(token, branch) {
  const s = sess_(token);
  if (s.role !== 'admin') return s.id;
  branch = String(branch || '').trim();
  if (!branch || branchList_().indexOf(branch) < 0) throw new Error('Choose a branch first.');
  return branch;
}

function branchList_() {
  return users_().getDataRange().getValues().slice(1).filter(function (r) {
    return String(r[2]).trim().toLowerCase() === 'user' && String(r[3] || 'approved').trim().toLowerCase() === 'approved';
  }).map(function (r) { return String(r[0]).trim(); }).filter(String).sort();
}

function pending_() {
  return users_().getDataRange().getValues().slice(1).filter(function (r) {
    return String(r[3]).trim().toLowerCase() === 'pending';
  }).map(function (r) {
    return { branch: String(r[0]).trim(), contact: s_(r[4]), requested: s_(r[5]).slice(0, 10) };
  });
}

// Branch sheets now contain sales data only. User-controlled status and due-date overrides
// are stored in the permanent Follow Up sheet and joined back into the in-memory row model.
function migrateBranchState_(b) {
  const sh = cloud_().getSheetByName(b);
  if (!sh || sh.getLastRow() < 2 || sh.getLastColumn() <= 15) return;
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const hasLegacyState = headers.some(function(h){ return /^(status|modified due date)$/i.test(h.trim()); });
  if (!hasLegacyState) return;
  const raw = sh.getRange(2, 1, sh.getLastRow()-1, Math.min(sh.getLastColumn(), 17)).getValues();
  const fsh = followSheet_();
  const last = fsh.getLastRow();
  const vals = last > 1 ? fsh.getRange(2,1,last-1,FOLLOW_COLS).getValues() : [];
  const index = {};
  vals.forEach(function(r,i){ if(s_(r[0])===b) index[s_(r[1])+'|'+s_(r[2])+'|'+s_(r[3])+'|'+s_(r[4])+'|'+s_(r[15])] = i; });
  const now = Utilities.formatDate(new Date(), CONFIG.TZ, "yyyy-MM-dd'T'HH:mm:ss");
  const backup = [];
  raw.forEach(function(r,i){
    const phone=s_(r[5]), date=s_(r[2]), time=s_(r[3]), due=s_(r[16]) || s_(r[12]);
    const status=s_(r[15]), mod=s_(r[16]) && s_(r[16])!==s_(r[12]) ? s_(r[16]) : '';
    // Legacy layout: column 16 is Status and column 17 is Modified Due Date.
    const legacyStatus=s_(r[15]);
    const legacyMod=s_(r[16]);
    if (!phone || !date || !time || (!legacyStatus && !legacyMod)) return;
    const effectiveDue=legacyMod || s_(r[12]);
    const code=s_(r[6]);
    const key=phone+'|'+effectiveDue+'|'+date+'|'+time+'|'+code;
    let ix=index[key];
    if(ix==null){ ix=vals.length; index[key]=ix; vals.push([b,phone,effectiveDue,date,time,s_(r[4]),s_(r[14]),'','','','',legacyStatus==='complete'?'complete':'incomplete',now,legacyMod,now,code]); }
    else {
      if(legacyStatus) vals[ix][11]=legacyStatus==='complete'?'complete':'incomplete';
      if(legacyMod) vals[ix][13]=legacyMod;
      vals[ix][12]=now; vals[ix][14]=now;
    }
    backup.push([b,i+2,phone,s_(r[12]),date,time,legacyStatus,legacyMod,now,s_(r[6])]);
  });
  if(vals.length) fsh.getRange(2,1,vals.length,FOLLOW_COLS).setNumberFormat('@').setValues(vals);
  // Verify the migrated records are present before archiving and removing legacy columns.
  const verify = fsh.getLastRow()>1 ? fsh.getRange(2,1,fsh.getLastRow()-1,FOLLOW_COLS).getValues() : [];
  const verified = backup.every(function(x){ return verify.some(function(r){ return s_(r[0])===b && s_(r[1])===x[2] && s_(r[3])===x[4] && s_(r[4])===x[5] && s_(r[15])===x[9] && (!x[6] || s_(r[11])===(x[6]==='complete'?'complete':'incomplete')) && (s_(r[13])===x[7] || !x[7]); }); });
  if(!verified) throw new Error('Legacy status migration verification failed for '+b+'. Original Branch columns were retained.');
  if(backup.length){
    const ss=cloud_(), name='Branch State Backup';
    const bk=ss.getSheetByName(name)||ss.insertSheet(name);
    if(bk.getLastRow()===0) bk.appendRow(['Branch','Source Row','Loyalty Phone','Original Due Date','Trx Date','Trx Time','Status','Modified Due Date','Archived At','Item Code']);
    bk.getRange(bk.getLastRow()+1,1,backup.length,10).setNumberFormat('@').setValues(backup);
  }
  // Remove the legacy Status and Modified Due Date columns only after copy + verification.
  if(sh.getLastColumn()>=17) sh.deleteColumns(16,2); else if(sh.getLastColumn()>=16) sh.deleteColumn(16);
}

function migrateAllBranchState() {
  const branches=branchList_(), results=[];
  branches.forEach(function(b){ migrateBranchState_(b); results.push(b); });
  return 'Migration checked '+results.length+' approved branch sheets. Legacy state columns were archived and removed only where verification succeeded.';
}

function readBranch_(b) {
  migrateBranchState_(b);
  const sh = cloud_().getSheetByName(b);
  if (!sh || sh.getLastRow() < 2) return [];
  const head = sh.getRange(1, 1, 1, 15).getValues()[0].map(String);
  if (head.join('|') !== CONFIG.OUT.slice(0, 15).join('|')) {
    throw new Error('The sheet "' + b + '" in the cloud file uses an older column layout. Rename or delete that sheet, then upload the files again.');
  }
  const recs=followMap_(b);
  return sh.getRange(2,1,sh.getLastRow()-1,15).getValues().map(function(r){
    const row=[s_(r[0]),s_(r[1]),s_(r[2]),s_(r[3]),s_(r[4]),s_(r[5]),s_(r[6]),s_(r[7]),Number(r[8])||0,Number(r[9])||0,Number(r[10])||0,s_(r[11]),s_(r[12]),Number(r[13])||0,s_(r[14]),'', ''];
    const f=recs[flinekey_(row)] || recs[fkey_(row)+'|'+row[I.CODE]] || recs[fkey_(row)];
    if(f){ row[I.ST]=f.status==='complete'?'complete':''; row[I.MOD]=f.mod||''; }
    return row;
  });
}

function writeBranch_(b, rows) {
  const ss=cloud_(), sh=ss.getSheetByName(b)||ss.insertSheet(b), H=CONFIG.OUT.slice(0,15);
  sh.clear();
  if(sh.getMaxRows()<rows.length+1) sh.insertRowsAfter(sh.getMaxRows(),rows.length+1-sh.getMaxRows());
  sh.getRange(1,1,1,H.length).setValues([H]).setFontWeight('bold'); sh.setFrozenRows(1);
  if(!rows.length) return;
  [1,2,3,4,5,6,7,8,12,13,15].forEach(function(c){sh.getRange(2,c,rows.length,1).setNumberFormat('@');});
  const values=rows.map(function(r){return r.slice(0,15);}), CH=20000;
  for(let i=0;i<values.length;i+=CH){const part=values.slice(i,i+CH);sh.getRange(2+i,1,part.length,15).setValues(part);}
}

function s_(v) {
  return v instanceof Date ? Utilities.formatDate(v, CONFIG.TZ, 'yyyy-MM-dd') : String(v == null ? '' : v);
}

function isDate_(s) { return /^\d{4}-\d{2}-\d{2}$/.test(String(s)); }

function lastDate_(rows) {
  return rows.reduce(function (m, r) { return r[I.DATE] > m ? r[I.DATE] : m; }, '');
}

function addDays_(s, n) {
  const p = s.split('-');
  return Utilities.formatDate(new Date(Date.UTC(+p[0], +p[1] - 1, +p[2] + n)), 'UTC', 'yyyy-MM-dd');
}

function salt_() {
  const props = PropertiesService.getScriptProperties();
  let s = props.getProperty('AUTH_SALT');
  if (!s) { s = Utilities.getUuid(); props.setProperty('AUTH_SALT', s); }
  return s;
}

function hashPw_(id, pw) {
  const raw = salt_() + '|' + String(id).trim().toLowerCase() + '|' + pw;
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw, Utilities.Charset.UTF_8)
    .map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join('');
}
