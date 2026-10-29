/**
 * ระบบเสนออนุมัติการอนุเคราะห์/ค้างชำระ (API)
 * กลุ่มงานสังคมสงเคราะห์ โรงพยาบาลสมเด็จพระสังฆราช (อมฺพรมหาเถร)
 *
 * Apps Script ตัวนี้เป็น API ให้หน้าเว็บบน GitHub Pages เรียกใช้
 * ข้อมูลเก็บใน Google Sheet ตาม SHEET_ID ด้านล่าง
 *
 * ติดตั้ง: ดู README.md
 *   1) วางโค้ดนี้ใน Apps Script แล้วรันฟังก์ชัน setup() หนึ่งครั้ง
 *   2) ดู PIN ของผู้ใช้ admin ใน Execution log
 *   3) Deploy > New deployment > Web app
 *      Execute as: Me / Who has access: Anyone
 */

const SHEET_ID = '1Qslzy0FfgsKFOMhYrWlz6i4khBq_3nQC11HT_V0UeI0';
const TZ = 'Asia/Bangkok';
const LIMITS = [3000, 20000, 50000, Infinity];           // วงเงินสูงสุดของแต่ละระดับ
const TIER_NAMES = ['นักสังคมสงเคราะห์', 'หัวหน้านักสังคมสงเคราะห์',
  'หัวหน้ากลุ่มภารกิจด้านบริการทุติยภูมิและตติยภูมิ/รองผู้อำนวยการฝ่ายการแพทย์', 'ผู้อำนวยการ'];
const SESSION_SECONDS = 6 * 60 * 60;                       // อายุการเข้าสู่ระบบ 6 ชั่วโมง
const MAX_FAILS = 5, LOCK_SECONDS = 15 * 60;               // ใส่ PIN ผิด 5 ครั้ง ล็อก 15 นาที
const HASH_ROUNDS = 300;

const SHEETS = {
  Cases: ['id', 'submittedAt', 'updatedAt', 'status', 'tier', 'requestedTier', 'hn', 'an', 'ptype', 'assessedOn', 'fy',
    'level', 'result', 'cost', 'covered', 'other', 'owe', 'payNow', 'inst', 'nInst', 'members', 'income', 'pc', 'fra', 'd870',
    'amount', 'approvedAmount', 'submittedBy', 'decidedBy', 'decidedAt', 'swNote', 'data'],
  History: ['at', 'caseId', 'hn', 'by', 'action', 'tier', 'amount', 'note'],
  Users: ['username', 'name', 'tiers', 'admin', 'active', 'pinHash', 'salt', 'mustChange', 'updatedAt']
};
const TEXT_COLS = { Cases: ['id', 'submittedAt', 'updatedAt', 'hn', 'an', 'assessedOn', 'submittedBy', 'decidedBy', 'decidedAt', 'swNote', 'data'],
  History: ['at', 'caseId', 'hn', 'by', 'note'], Users: ['username', 'tiers', 'pinHash', 'salt', 'updatedAt'] };

/* ======================= setup ======================= */
function setup() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  Object.keys(SHEETS).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    const cols = SHEETS[name];
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold').setBackground('#e8eef5');
    sh.setFrozenRows(1);
    (TEXT_COLS[name] || []).forEach(c => sh.getRange(2, cols.indexOf(c) + 1, sh.getMaxRows() - 1, 1).setNumberFormat('@'));
  });
  const users = sheet_('Users');
  if (users.getLastRow() < 2) {
    const pin = String(Math.floor(100000 + Math.random() * 900000));
    const salt = Utilities.getUuid();
    users.appendRow(['admin', 'ผู้ดูแลระบบ', '', true, true, hash_(pin, salt), salt, true, now_()]);
    Logger.log('สร้างผู้ใช้ admin แล้ว PIN เริ่มต้น = ' + pin + ' (ระบบจะให้เปลี่ยนเมื่อเข้าสู่ระบบครั้งแรก)');
  } else {
    Logger.log('มีผู้ใช้อยู่แล้ว ไม่ได้สร้าง admin ใหม่');
  }
  bump_();
  Logger.log('ตั้งค่าเสร็จ');
}

/** ใช้ฉุกเฉินเมื่อลืม PIN ของ admin: แก้ค่า pin ด้านล่างแล้วรันใน editor */
function resetAdminPin() {
  const pin = '123456';
  setPin_('admin', pin, true);
  Logger.log('ตั้ง PIN ของ admin เป็น ' + pin + ' แล้ว (ต้องเปลี่ยนเมื่อเข้าสู่ระบบ)');
}

/* ======================= HTTP ======================= */
function doGet() {
  return out_({ ok: true, service: 'sw-assist-api', time: now_() });
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (x) { return out_({ ok: false, error: 'bad_request' }); }
  try {
    const fn = ACTIONS[req.action];
    if (!fn) throw err_('unknown_action');
    const user = req.action === 'login' ? null : auth_(req.token);
    return out_(Object.assign({ ok: true }, fn(req, user) || {}));
  } catch (x) {
    return out_({ ok: false, error: x.code || 'server_error', message: x.code ? x.message : String(x) });
  }
}

const ACTIONS = {
  login(req) {
    const username = String(req.username || '').trim().toLowerCase();
    const cache = CacheService.getScriptCache();
    const failKey = 'fail:' + username;
    const fails = +(cache.get(failKey) || 0);
    if (fails >= MAX_FAILS) throw err_('locked', 'ใส่ PIN ผิดหลายครั้ง ลองใหม่ใน 15 นาที');
    const u = findUser_(username);
    if (!u || !truthy_(u.active) || hash_(String(req.pin || ''), u.salt) !== u.pinHash) {
      cache.put(failKey, String(fails + 1), LOCK_SECONDS);
      throw err_('bad_login', 'ชื่อผู้ใช้หรือ PIN ไม่ถูกต้อง');
    }
    cache.remove(failKey);
    const token = Utilities.getUuid() + Utilities.getUuid().slice(0, 8);
    cache.put('tok:' + token, u.username, SESSION_SECONDS);
    return { token: token, me: pub_(u, true), users: userList_() };
  },

  me(req, user) { return { me: pub_(user, true), users: userList_() }; },

  logout(req) { CacheService.getScriptCache().remove('tok:' + req.token); return {}; },

  changePin(req, user) {
    if (hash_(String(req.oldPin || ''), user.salt) !== user.pinHash) throw err_('bad_login', 'PIN เดิมไม่ถูกต้อง');
    checkPin_(req.newPin);
    setPin_(user.username, String(req.newPin), false);
    return {};
  },

  list(req) {
    const ver = getVer_();
    if (req.ver && String(req.ver) === ver) return { unchanged: true, ver: ver };
    return { ver: ver, cases: readCases_(), users: userList_() };
  },

  submit(req, user) {
    const c = req.case || {};
    const level = +c.level;
    const amount = Math.round(+c.amount);
    if (!(level >= 1 && level <= 3)) throw err_('invalid', 'ผลการประเมินต้องเป็นอนุเคราะห์หรือค้างชำระ');
    if (!(amount > 0)) throw err_('invalid', 'ยอดที่เสนอต้องมากกว่า 0');
    if (!String(c.hn || '').trim()) throw err_('invalid', 'ต้องระบุ HN');
    const tier = tierOf_(amount);
    return withLock_(() => {
      const now = now_();
      let row, old = null;
      if (req.editingId) {
        row = findCaseRow_(req.editingId);
        if (!row) throw err_('not_found', 'ไม่พบรายการเดิม');
        old = row.obj;
        if (old.status !== 'returned') throw err_('conflict', 'รายการนี้ไม่ได้อยู่ในสถานะส่งกลับแก้ไข');
        if (old.submittedBy !== user.username && !hasTier_(user, 0) && !truthy_(user.admin)) throw err_('forbidden', 'ไม่มีสิทธิ์แก้ไขรายการนี้');
      }
      const id = old ? old.id : Utilities.getUuid().replace(/-/g, '').slice(0, 20);
      const history = old ? (old.history || []) : [];
      const h = { at: now, by: user.username, action: old ? 'resubmit' : 'submit', tier: tier, amount: amount };
      history.push(h);
      const data = Object.assign({}, c, {
        amount: amount, tier: tier, requestedTier: tier, status: 'pending', decision: null,
        submittedBy: old ? old.submittedBy : user.username, submittedAt: old ? old.submittedAt : now, updatedAt: now,
        history: history
      });
      writeCase_(id, data, row ? row.index : null);
      logHistory_(id, data.hn, h);
      bump_();
      return { id: id, tier: tier };
    });
  },

  decide(req, user) {
    const action = req.decision;
    if (['approve', 'escalate', 'return', 'reject'].indexOf(action) < 0) throw err_('invalid', 'ผลการพิจารณาไม่ถูกต้อง');
    const note = String(req.note || '').slice(0, 1000);
    return withLock_(() => {
      const row = findCaseRow_(req.id);
      if (!row) throw err_('not_found', 'ไม่พบรายการ');
      const c = row.obj;
      if (c.status !== 'pending') throw err_('conflict', 'รายการนี้มีผู้พิจารณาไปแล้ว');
      if (req.tier != null && +req.tier !== +c.tier) throw err_('conflict', 'รายการนี้ถูกเสนอระดับอื่นไปแล้ว');
      const mx = maxTier_(user);
      const own = c.submittedBy === user.username;
      const selfEscalate = action === 'escalate' && own && c.tier >= 1 && hasTier_(user, c.tier);
      if (!selfEscalate) {
        if (mx < c.tier) throw err_('forbidden', 'ยอดนี้เกินวงเงินของคุณ');
        if (c.tier >= 1 && own) throw err_('forbidden', 'ผู้เสนอพิจารณารายของตนเองไม่ได้ในระดับนี้ ให้เสนอระดับถัดไป');
      }
      if ((action === 'reject' || action === 'return') && !note) throw err_('invalid', 'ระบุเหตุผลในช่องความเห็น');
      if (action === 'escalate' && c.tier >= 3) throw err_('invalid', 'เป็นระดับสูงสุดแล้ว');
      const now = now_();
      const h = { at: now, by: user.username, action: action, tier: c.tier };
      if (action === 'approve') {
        const amt = Math.round(+req.amount);
        if (!(amt >= 0)) throw err_('invalid', 'ระบุยอดที่อนุมัติ');
        if (amt > LIMITS[mx]) throw err_('forbidden', 'ยอดเกินวงเงินของคุณ ให้เสนอระดับถัดไป');
        h.amount = amt;
      }
      if (note) h.note = note;
      c.history = (c.history || []).concat([h]);
      c.updatedAt = now;
      if (action === 'approve') { c.status = 'approved'; c.decision = h; }
      else if (action === 'reject') { c.status = 'rejected'; c.decision = h; }
      else if (action === 'return') { c.status = 'returned'; }
      else if (action === 'escalate') { c.tier = Math.min(3, c.tier + 1); }
      writeCase_(c.id, c, row.index);
      logHistory_(c.id, c.hn, h);
      bump_();
      return { status: c.status, tier: c.tier };
    });
  },

  /* ---------- admin ---------- */
  adminUsers(req, user) {
    requireAdmin_(user);
    return { users: readUsers_().map(u => pub_(u, true)) };
  },

  saveUser(req, user) {
    requireAdmin_(user);
    const u = req.user || {};
    const username = String(u.username || '').trim().toLowerCase();
    if (!/^[a-z0-9._-]{3,30}$/.test(username)) throw err_('invalid', 'ชื่อผู้ใช้ใช้ a-z 0-9 . _ - ยาว 3–30 ตัว');
    const tiers = (u.tiers || []).map(Number).filter(t => t >= 0 && t <= 3);
    return withLock_(() => {
      const sh = sheet_('Users'); const cols = SHEETS.Users;
      const found = findUserRow_(username);
      if (found) {
        if (username === user.username && u.admin === false) throw err_('invalid', 'นำสิทธิ์ผู้ดูแลของตนเองออกไม่ได้');
        const r = found.index;
        sh.getRange(r, cols.indexOf('name') + 1).setValue(String(u.name || '').slice(0, 100));
        sh.getRange(r, cols.indexOf('tiers') + 1).setValue(tiers.join(','));
        sh.getRange(r, cols.indexOf('admin') + 1).setValue(!!u.admin);
        sh.getRange(r, cols.indexOf('active') + 1).setValue(u.active !== false);
        sh.getRange(r, cols.indexOf('updatedAt') + 1).setValue(now_());
      } else {
        checkPin_(u.pin);
        const salt = Utilities.getUuid();
        sh.appendRow([username, String(u.name || '').slice(0, 100), tiers.join(','), !!u.admin, u.active !== false,
          hash_(String(u.pin), salt), salt, true, now_()]);
      }
      bump_();
      return { users: readUsers_().map(x => pub_(x, true)) };
    });
  },

  resetPin(req, user) {
    requireAdmin_(user);
    checkPin_(req.pin);
    setPin_(String(req.username || '').toLowerCase(), String(req.pin), true);
    return {};
  }
};

/* ======================= helpers ======================= */
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function err_(code, message) { const e = new Error(message || code); e.code = code; return e; }
function now_() { return new Date().toISOString(); }
function truthy_(v) { return v === true || String(v).toUpperCase() === 'TRUE'; }
function sheet_(name) {
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(name);
  if (!sh) throw err_('not_setup', 'ยังไม่ได้รัน setup() ใน Apps Script');
  return sh;
}
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw err_('busy', 'ระบบกำลังบันทึกข้อมูลอื่น ลองอีกครั้ง');
  try { return fn(); } finally { lock.releaseLock(); }
}
function getVer_() { return PropertiesService.getScriptProperties().getProperty('ver') || '0'; }
function bump_() { PropertiesService.getScriptProperties().setProperty('ver', String(Date.now())); }
function tierOf_(a) { for (let i = 0; i < LIMITS.length; i++) if (a <= LIMITS[i]) return i; return 3; }

function hash_(pin, salt) {
  let h = salt + '|' + pin;
  for (let i = 0; i < HASH_ROUNDS; i++) {
    h = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h + salt, Utilities.Charset.UTF_8)
      .map(b => ('0' + (b & 255).toString(16)).slice(-2)).join('');
  }
  return h;
}
function checkPin_(pin) { if (!/^\d{6}$/.test(String(pin || ''))) throw err_('invalid', 'PIN ต้องเป็นตัวเลข 6 หลัก'); }
function setPin_(username, pin, mustChange) {
  withLock_(() => {
    const found = findUserRow_(username);
    if (!found) throw err_('not_found', 'ไม่พบผู้ใช้');
    const sh = sheet_('Users'); const cols = SHEETS.Users; const salt = Utilities.getUuid();
    sh.getRange(found.index, cols.indexOf('pinHash') + 1, 1, 3).setValues([[hash_(pin, salt), salt, !!mustChange]]);
    sh.getRange(found.index, cols.indexOf('updatedAt') + 1).setValue(now_());
  });
}

/* users */
function readUsers_() {
  const sh = sheet_('Users'); const cols = SHEETS.Users;
  const v = sh.getDataRange().getValues().slice(1);
  return v.filter(r => r[0]).map(r => { const o = {}; cols.forEach((c, i) => o[c] = r[i]); o.username = String(o.username).toLowerCase(); return o; });
}
function findUser_(username) { return readUsers_().filter(u => u.username === username)[0] || null; }
function findUserRow_(username) {
  const v = sheet_('Users').getRange('A:A').getValues();
  for (let i = 1; i < v.length; i++) if (String(v[i][0]).toLowerCase() === username) return { index: i + 1 };
  return null;
}
function tiersOf_(u) { return String(u.tiers || '').split(',').filter(x => x !== '').map(Number); }
function hasTier_(u, t) { return tiersOf_(u).indexOf(t) >= 0; }
function maxTier_(u) { const t = tiersOf_(u); return t.length ? Math.max.apply(null, t) : -1; }
function pub_(u, full) {
  const o = { username: u.username, name: u.name || u.username, tiers: tiersOf_(u) };
  if (full) { o.admin = truthy_(u.admin); o.active = truthy_(u.active); o.mustChange = truthy_(u.mustChange); }
  return o;
}
function userList_() { return readUsers_().map(u => pub_(u, false)); }
function requireAdmin_(u) { if (!truthy_(u.admin)) throw err_('forbidden', 'เฉพาะผู้ดูแลระบบ'); }
function auth_(token) {
  if (!token) throw err_('auth', 'กรุณาเข้าสู่ระบบ');
  const cache = CacheService.getScriptCache();
  const username = cache.get('tok:' + token);
  if (!username) throw err_('auth', 'หมดเวลาการใช้งาน กรุณาเข้าสู่ระบบใหม่');
  const u = findUser_(username);
  if (!u || !truthy_(u.active)) throw err_('auth', 'บัญชีถูกปิดการใช้งาน');
  cache.put('tok:' + token, username, SESSION_SECONDS);   // ต่ออายุเมื่อมีการใช้งาน
  return u;
}

/* cases */
function rowToCase_(r) {
  const cols = SHEETS.Cases; const o = {};
  cols.forEach((c, i) => { let v = r[i]; if (v instanceof Date) v = Utilities.formatDate(v, TZ, "yyyy-MM-dd'T'HH:mm:ssXXX"); o[c] = v; });
  let data = {};
  try { data = JSON.parse(o.data || '{}'); } catch (x) {}
  delete o.data;
  const merged = Object.assign({}, data, o);
  ['tier', 'requestedTier', 'level'].forEach(k => merged[k] = +merged[k]);
  return merged;
}
function readCases_() {
  const sh = sheet_('Cases');
  const last = sh.getLastRow();
  if (last < 2) return [];
  const start = Math.max(2, last - 1499);                   // ส่ง 1,500 รายการล่าสุด
  const v = sh.getRange(start, 1, last - start + 1, SHEETS.Cases.length).getValues();
  return v.filter(r => r[0]).map(rowToCase_).reverse();
}
function findCaseRow_(id) {
  const sh = sheet_('Cases');
  const f = sh.getRange('A:A').createTextFinder(String(id)).matchEntireCell(true).findNext();
  if (!f) return null;
  const r = f.getRow();
  return { index: r, obj: rowToCase_(sh.getRange(r, 1, 1, SHEETS.Cases.length).getValues()[0]) };
}
function writeCase_(id, c, rowIndex) {
  const d = c.decision || {};
  const rec = {
    id: id, submittedAt: c.submittedAt, updatedAt: c.updatedAt, status: c.status, tier: c.tier, requestedTier: c.requestedTier,
    hn: c.hn, an: c.an, ptype: c.ptype, assessedOn: c.assessedOn, fy: c.fy, level: c.level, result: c.result,
    cost: c.cost, covered: c.covered, other: c.other, owe: c.owe, payNow: c.payNow, inst: c.inst, nInst: c.nInst,
    members: c.members, income: c.income, pc: c.pc, fra: c.fra, d870: c.d870, amount: c.amount,
    approvedAmount: c.status === 'approved' ? d.amount : '', submittedBy: c.submittedBy,
    decidedBy: d.by || '', decidedAt: d.at || '', swNote: c.swNote
  };
  const data = Object.assign({}, c);
  Object.keys(rec).forEach(k => { if (k !== 'id') delete data[k]; });
  ['status', 'tier', 'submittedBy', 'submittedAt', 'updatedAt'].forEach(k => data[k] = c[k]);
  data.decision = c.decision || null;
  rec.data = JSON.stringify(data);
  const row = SHEETS.Cases.map(k => rec[k] == null ? '' : rec[k]);
  const sh = sheet_('Cases');
  if (rowIndex) sh.getRange(rowIndex, 1, 1, row.length).setValues([row]);
  else sh.appendRow(row);
}
function logHistory_(caseId, hn, h) {
  sheet_('History').appendRow([h.at, caseId, hn || '', h.by, h.action, h.tier, h.amount == null ? '' : h.amount, h.note || '']);
}
