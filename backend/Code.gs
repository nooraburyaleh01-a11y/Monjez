/** MONJEZ Pro — account-linked activation and Google Sheets storage.
 * Preserves existing YAWM_DATA / YAWM_LICENSES rows and legacy license data.
 * Activation keys are account credentials; a linked account can sign in on another device.
 */
const DATA_SHEET = 'YAWM_DATA';
const LICENSE_SHEET = 'YAWM_LICENSES';
const ACCOUNT_SHEET = 'MONJEZ_ACCOUNTS';
const STATE_PREFIX = 'ACCOUNT_STATE_';
const LEGACY_STATE_PREFIX = 'APP_STATE_JSON_';

function doGet() {
  ensureDataSheet_(); ensureLicenseSheet_(); ensureAccountSheet_();
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('منجز | MONJEZ')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, viewport-fit=cover');
}

/** JSON API bridge used by the Vercel-hosted frontend. Keep the web app deployed as 'Anyone'. */
function doPost(e) {
  try {
    const request = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const method = String(request.method || '');
    const args = Array.isArray(request.args) ? request.args : [];
    const allowed = ['activateMonjez','completeMonjezProfile','resumeMonjezSession','getMonjezData','saveMonjezData'];
    if (allowed.indexOf(method) < 0) throw new Error('طلب API غير مسموح.');
    const result = dispatchMonjezApi_(method, args);
    return ContentService.createTextOutput(JSON.stringify({ ok: true, result }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err && err.message || err || 'حدث خطأ غير متوقع.') }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
function dispatchMonjezApi_(method, args) {
  switch (method) {
    case 'activateMonjez': return activateMonjez.apply(null, args);
    case 'completeMonjezProfile': return completeMonjezProfile.apply(null, args);
    case 'resumeMonjezSession': return resumeMonjezSession.apply(null, args);
    case 'getMonjezData': return getMonjezData.apply(null, args);
    case 'saveMonjezData': return saveMonjezData.apply(null, args);
    default: throw new Error('طلب API غير مسموح.');
  }
}

/** First use: validate key, bind it to an account, and return either profile setup or account login. */
function activateMonjez(key, installId) {
  key = normalizeKey_(key); installId = validateInstallId_(installId);
  if (!key) throw new Error('أدخل مفتاح التفعيل.');
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const sh = ensureLicenseSheet_();
    const last = sh.getLastRow();
    const rows = last > 1 ? sh.getRange(2, 1, last - 1, 9).getValues() : [];
    const idx = rows.findIndex(r => normalizeKey_(r[0]) === key);
    if (idx < 0) throw new Error('مفتاح التفعيل غير صحيح.');
    const rowNum = idx + 2, row = rows[idx];
    if (String(row[1]) !== 'ACTIVE') throw new Error('مفتاح التفعيل غير متاح.');
    let licenseId = String(row[3] || '');
    if (!licenseId) { licenseId = 'lic_' + Utilities.getUuid(); sh.getRange(rowNum, 4).setValue(licenseId); }
    if (!row[5]) sh.getRange(rowNum, 6).setValue(new Date());
    let accountId = String(row[7] || '');
    if (!accountId) { accountId = 'acc_' + Utilities.getUuid(); sh.getRange(rowNum, 8).setValue(accountId); }
    const account = findAccount_(accountId);
    let profile = account ? account.profile : null;
    // One-time migration path for existing installations that already saved a profile and cloud data.
    if (!profile) {
      const legacy = readStateByKey_(LEGACY_STATE_PREFIX + licenseId);
      if (legacy && legacy.profile && String(legacy.profile.name || '').trim()) {
        profile = normalizeProfile_(legacy.profile);
        upsertAccount_(accountId, profile);
        writeAccountState_(accountId, legacy);
      }
    }
    if (profile) return { ok: true, needsProfile: false, accountId, licenseId, profile, message: 'أهلًا بعودتك إلى منجز.' };
    return { ok: true, needsProfile: true, accountId, licenseId, message: 'هذا مفتاح جديد. أنشئ ملفك الشخصي للمتابعة.' };
  } finally { lock.releaseLock(); }
}

/** Save the profile for a newly activated account. Country is currently Jordan only. */
function completeMonjezProfile(key, accountId, profile, installId) {
  key = normalizeKey_(key); accountId = String(accountId || ''); installId = validateInstallId_(installId);
  if (!key || !/^acc_[\w-]{10,100}$/.test(accountId)) throw new Error('جلسة التفعيل غير صالحة. أعد إدخال المفتاح.');
  const name = String((profile || {}).name || '').trim();
  const birthDate = String((profile || {}).birthDate || '').trim();
  const country = String((profile || {}).country || 'Jordan');
  if (name.length < 2) throw new Error('اكتب الاسم الكامل أو الاسم الذي تريد استخدامه.');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(birthDate)) throw new Error('اختر تاريخ الميلاد.');
  if (country !== 'Jordan') throw new Error('الدولة المتاحة حاليًا هي الأردن.');
  const dob = new Date(birthDate + 'T12:00:00');
  if (isNaN(dob.getTime()) || dob > new Date()) throw new Error('تاريخ الميلاد غير صالح.');
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const sh = ensureLicenseSheet_(), rows = sh.getLastRow() > 1 ? sh.getRange(2,1,sh.getLastRow()-1,9).getValues() : [];
    const idx = rows.findIndex(r => normalizeKey_(r[0]) === key);
    if (idx < 0 || String(rows[idx][1]) !== 'ACTIVE') throw new Error('مفتاح التفعيل غير صحيح أو غير متاح.');
    const rowNum = idx + 2, row = rows[idx];
    if (String(row[7] || '') !== accountId) throw new Error('هذا المفتاح مرتبط بحساب آخر.');
    const normalized = { name, birthDate, country: 'Jordan', timezone: 'Asia/Amman', lang: 'ar', dark: false,
      waterGoal: 8, calorieGoal: 2000, qada: 0, onboardingComplete: true, tutorialDone: false };
    upsertAccount_(accountId, normalized);
    // Migrate legacy cloud state once, then attach the new profile without erasing user data.
    const accountState = readStateByKey_(STATE_PREFIX + accountId);
    if (!accountState) {
      const legacy = readStateByKey_(LEGACY_STATE_PREFIX + String(row[3] || ''));
      if (legacy) { legacy.profile = Object.assign({}, legacy.profile || {}, normalized); writeAccountState_(accountId, legacy); }
    }
    sh.getRange(rowNum, 9).setValue(JSON.stringify({ country:'Jordan', timezone:'Asia/Amman' }));
    return { ok:true, accountId, licenseId:String(row[3]), profile:normalized, message:'تم إنشاء حسابك بنجاح.' };
  } finally { lock.releaseLock(); }
}

/** Resume a previously activated account on this device. */
function resumeMonjezSession(licenseId, accountId, installId) {
  const row = getValidLicense_(licenseId, accountId, installId);
  const account = findAccount_(accountId);
  if (!account || !account.profile) throw new Error('أكمل معلومات حسابك للمتابعة.');
  return { ok:true, accountId, licenseId, profile:account.profile };
}

function validateMonjezActivation(licenseId, installId) {
  const row = getValidLicense_(licenseId, '', installId);
  return {ok:true, licenseId:licenseId, accountId:String(row[7] || '')};
}
function getMonjezData(accountId, licenseId, installId) {
  getValidLicense_(licenseId, accountId, installId);
  let state = readStateByKey_(STATE_PREFIX + accountId);
  if (state) return state;
  // Backward compatibility with previous license-keyed cloud state.
  const legacy = readStateByKey_(LEGACY_STATE_PREFIX + licenseId);
  if (legacy) { writeAccountState_(accountId, legacy); return legacy; }
  return null;
}
function saveMonjezData(jsonString, accountId, licenseId, installId) {
  getValidLicense_(licenseId, accountId, installId);
  if (typeof jsonString !== 'string' || !jsonString.trim()) throw new Error('لا توجد بيانات صالحة للحفظ.');
  let state; try { state = JSON.parse(jsonString); } catch (e) { throw new Error('صيغة البيانات غير صالحة.'); }
  if (!state || typeof state !== 'object' || !Array.isArray(state.tasks) || !Array.isArray(state.habits)) throw new Error('بيانات التطبيق غير مكتملة.');
  writeAccountState_(accountId, state);
  return {ok:true, savedAt:new Date().toISOString()};
}

function getValidLicense_(licenseId, accountId, installId) {
  if (!licenseId) throw new Error('أدخل مفتاح التفعيل.');
  validateInstallId_(installId);
  const sh = ensureLicenseSheet_(), last = sh.getLastRow();
  if (last < 2) throw new Error('لم يتم العثور على مفتاح التفعيل.');
  const rows = sh.getRange(2,1,last-1,9).getValues();
  const row = rows.find(r => String(r[3]) === String(licenseId));
  if (!row || String(row[1]) !== 'ACTIVE') throw new Error('مفتاح التفعيل غير متاح.');
  if (accountId && String(row[7] || '') !== String(accountId)) throw new Error('مفتاح التفعيل لا يطابق هذا الحساب.');
  return row;
}
function normalizeProfile_(p) {
  return { name:String(p.name || '').trim(), birthDate:String(p.birthDate || ''), country:'Jordan', timezone:'Asia/Amman', lang:'ar', dark:false,
    waterGoal:Number(p.waterGoal || 8), calorieGoal:Number(p.calorieGoal || 2000), qada:Number(p.qada || 0), onboardingComplete:!!p.onboardingComplete, tutorialDone:!!p.tutorialDone };
}
function findAccount_(accountId) {
  const sh = ensureAccountSheet_(), last = sh.getLastRow(); if (last < 2) return null;
  const rows = sh.getRange(2,1,last-1,5).getValues();
  const row = rows.find(r => String(r[0]) === String(accountId)); if (!row) return null;
  let profile = null; try { profile = row[1] ? JSON.parse(String(row[1])) : null; } catch (e) {}
  return {accountId:String(row[0]), profile, createdAt:row[2], updatedAt:row[3]};
}
function upsertAccount_(accountId, profile) {
  const sh = ensureAccountSheet_(), last = sh.getLastRow(); let rowNum = 0;
  if (last > 1) { const ids = sh.getRange(2,1,last-1,1).getValues().flat(); const i = ids.findIndex(v => String(v) === String(accountId)); if (i >= 0) rowNum = i + 2; }
  const now = new Date();
  if (rowNum) { const created = sh.getRange(rowNum,3).getValue() || now; sh.getRange(rowNum,1,1,5).setValues([[accountId,JSON.stringify(profile),created,now,'ACTIVE']]); }
  else sh.appendRow([accountId,JSON.stringify(profile),now,now,'ACTIVE']);
}
function readStateByKey_(key) {
  const sh = ensureDataSheet_(), last = sh.getLastRow(); if (last < 2) return null;
  const values = sh.getRange(2,1,last-1,2).getValues();
  for (let i=0;i<values.length;i++) if (String(values[i][0]) === String(key) && values[i][1]) {
    try { return JSON.parse(String(values[i][1])); } catch(e) { throw new Error('تعذر قراءة البيانات المحفوظة.'); }
  }
  return null;
}
function writeAccountState_(accountId, state) {
  const sh = ensureDataSheet_(), key = STATE_PREFIX + accountId, json = JSON.stringify(state);
  if (json.length > 45000) throw new Error('حجم بيانات الحساب كبير جدًا للحفظ في خلية واحدة.');
  const last = sh.getLastRow(); let target = 0;
  if (last > 1) { const keys = sh.getRange(2,1,last-1,1).getValues().flat(); const i = keys.findIndex(v => String(v) === key); if (i >= 0) target = i + 2; }
  if (!target) target = Math.max(2, sh.getLastRow() + 1);
  sh.getRange(target,1,1,3).setValues([[key,json,new Date()]]);
  sh.getRange(target,3).setNumberFormat('yyyy-mm-dd hh:mm:ss');
}
function validateInstallId_(id) { id=String(id||''); if(!/^[a-zA-Z0-9_-]{16,100}$/.test(id)) throw new Error('معرّف تثبيت غير صالح.'); return id; }
function normalizeKey_(key) { return String(key||'').trim().toUpperCase().replace(/\s+/g,''); }
function ensureDataSheet_() {
  const ss=SpreadsheetApp.getActiveSpreadsheet(); if(!ss) throw new Error('اربط المشروع بملف Google Sheets.');
  let sh=ss.getSheetByName(DATA_SHEET); if(!sh) sh=ss.insertSheet(DATA_SHEET);
  if(!sh.getLastRow()){sh.getRange(1,1,1,3).setValues([['KEY','JSON_DATA','UPDATED_AT']]);sh.setFrozenRows(1);} return sh;
}
function ensureLicenseSheet_() {
  const ss=SpreadsheetApp.getActiveSpreadsheet(); if(!ss) throw new Error('اربط المشروع بملف Google Sheets.');
  let sh=ss.getSheetByName(LICENSE_SHEET); if(!sh) sh=ss.insertSheet(LICENSE_SHEET);
  if(!sh.getLastRow()){sh.getRange(1,1,1,9).setValues([['ACTIVATION_KEY','STATUS','LEGACY_INSTALL_ID','LICENSE_ID','CREATED_AT','ACTIVATED_AT','NOTE','ACCOUNT_ID','ACCOUNT_META_JSON']]);sh.setFrozenRows(1);}
  else {
    const current = sh.getRange(1,1,1,Math.max(7,sh.getLastColumn())).getValues()[0];
    if (!current[7]) sh.getRange(1,8).setValue('ACCOUNT_ID');
    if (!current[8]) sh.getRange(1,9).setValue('ACCOUNT_META_JSON');
  }
  return sh;
}
function ensureAccountSheet_() {
  const ss=SpreadsheetApp.getActiveSpreadsheet(); if(!ss) throw new Error('اربط المشروع بملف Google Sheets.');
  let sh=ss.getSheetByName(ACCOUNT_SHEET); if(!sh) sh=ss.insertSheet(ACCOUNT_SHEET);
  if(!sh.getLastRow()){sh.getRange(1,1,1,5).setValues([['ACCOUNT_ID','PROFILE_JSON','CREATED_AT','UPDATED_AT','STATUS']]);sh.setFrozenRows(1);sh.getRange(1,1,1,5).setFontWeight('bold');}
  return sh;
}
/** Run manually in Apps Script to issue 10 new permanent activation keys. */
function issueMonjezLicenseKeys() {
  const count=10, sh=ensureLicenseSheet_(), now=new Date(), rows=[];
  for(let i=0;i<count;i++){const raw=Utilities.getUuid().replace(/-/g,'').toUpperCase().slice(0,16);rows.push(['MNJ-'+raw.slice(0,4)+'-'+raw.slice(4,8)+'-'+raw.slice(8,12)+'-'+raw.slice(12,16),'ACTIVE','','',now,'','MONJEZ','','']);}
  sh.getRange(sh.getLastRow()+1,1,rows.length,9).setValues(rows); return 'تم إصدار '+count+' مفاتيح تفعيل في '+LICENSE_SHEET+'.';
}
function setupMonjez(){ensureDataSheet_();ensureLicenseSheet_();ensureAccountSheet_();return 'تم تجهيز منجز. لم يتم حذف أي بيانات موجودة.';}
