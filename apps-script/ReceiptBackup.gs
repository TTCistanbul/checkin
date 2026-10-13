/**
 * 報帳收據 → Google Drive 自動備份（Google Apps Script）
 *
 * 每天自動執行一次：
 *   1. 把「已確認」的收據檔案存到 Drive：「TTC 報帳收據備份 / YYYY-MM /」
 *      檔名跟打卡系統下載的一樣：日期_廠商_金額幣別_員工編號_代碼.jpg/pdf
 *   2. 每個月份資料夾裡維護一份「清冊_YYYY-MM.csv」（該月全部收據，含未確認的）
 *   3. 確認後又被修改（檔名變了）的收據，會存新版、舊版移到垃圾桶
 *   4. 打卡系統裡刪掉的收據，Drive 上的備份不會被刪——這就是備份的用意
 *
 * 安裝（只要做一次）：
 *   1. 到 https://script.google.com 用辦公室的 Google 帳號「新增專案」，
 *      把這整份程式碼貼上取代原本內容，存檔。
 *   2. 左邊齒輪「專案設定」→ 最下面「指令碼屬性」→ 新增屬性：
 *        名稱 SERVICE_ACCOUNT_JSON，值＝Firebase 服務帳號金鑰 JSON 的「整份內容」
 *   3. 回到編輯器，上方函式選單選 setup → 執行 → 依指示授權。
 *      setup 會建立備份資料夾、設定每天凌晨自動執行，並立刻先備份一次。
 *   之後想馬上備份，執行 backupNow 即可。
 */

const PROJECT_ID = 'office-checkin-855a6';
const ROOT_FOLDER_NAME = 'TTC 報帳收據備份';
const MAX_FILES_PER_RUN = 60;            // Apps Script 單次最多跑 6 分鐘，分批處理
const FS = `https://firestore.googleapis.com/v1/projects/${PROJECT_ID}/databases/(default)/documents`;

const CAT = { transport: '交通', meals: '餐費', office: '文具辦公', postage: '郵電快遞', entertain: '交際招待', misc: '雜支' };
const STATUS = { submitted: '待確認', verified: '已確認' };

/* ---------- 安裝與執行 ---------- */
function setup() {
  getRootFolder_();
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'backupNow')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('backupNow').timeBased().everyDays(1).atHour(3).create();
  console.log('已設定每天凌晨 3–4 點自動備份。現在先執行一次…');
  backupNow();
}

function backupNow() {
  const token = getAccessToken_();
  const receipts = listCollection_('receipts', token);
  const employees = {};
  listCollection_('employees', token).forEach(e => { employees[e.id] = e; });
  const nameOf = uid => (employees[uid] && employees[uid].name) || uid;
  const empNoOf = uid => (employees[uid] && (employees[uid].empNo || employees[uid].id)) || uid;

  const todo = planBackup_(receipts, empNoOf);
  const touched = {};
  let saved = 0;
  for (const r of todo.slice(0, MAX_FILES_PER_RUN)) {
    const month = monthOf_(r);
    const folder = getMonthFolder_(month);
    const name = fileName_(r, empNoOf);
    const blob = loadFile_(r, name, token);
    const file = folder.createFile(blob);
    if (r.driveFileId) {
      try { DriveApp.getFileById(r.driveFileId).setTrashed(true); } catch (e) { /* 舊檔可能已被手動刪除 */ }
    }
    patchDoc_('receipts/' + r.id, {
      driveFileId: { stringValue: file.getId() },
      backupName: { stringValue: name },
      backedUpAt: { timestampValue: new Date().toISOString() },
    }, token);
    // 同步更新記憶體裡的這筆，下面寫清冊時「已備份」欄才會是最新狀態
    r.driveFileId = file.getId(); r.backupName = name; r.backedUpAt = new Date().toISOString();
    touched[month] = true;
    saved++;
  }
  // 清冊：這次有動到的月份 + 本月，都重寫一次（反映最新的金額、狀態）
  touched[Utilities.formatDate(new Date(), 'Europe/Istanbul', 'yyyy-MM')] = true;
  Object.keys(touched).forEach(m => writeLedger_(m, receipts, nameOf, empNoOf));
  console.log(`備份完成：新存 ${saved} 個檔案，尚待下次處理 ${Math.max(0, todo.length - saved)} 個；更新清冊 ${Object.keys(touched).join(', ')}`);
}

/* ---------- 規則（哪些收據要備份、怎麼命名） ---------- */
function planBackup_(receipts, empNoOf) {
  return receipts.filter(r =>
    r.status === 'verified' && (!r.backedUpAt || r.backupName !== fileName_(r, empNoOf)));
}

function monthOf_(r) {
  return /^\d{4}-\d{2}/.test(r.date || '') ? r.date.slice(0, 7) : 'no-date';
}

function slug_(s) {
  return String(s || '').trim().replace(/[\\/:*?"<>|#%&{}$!'@+`=\s.,;]+/g, '-')
    .replace(/-+/g, '-').replace(/^-|-$/g, '').slice(0, 40);
}

// 跟打卡系統網頁的 rcFileName() 完全相同，下載的檔名和 Drive 上的才會一致
function fileName_(r, empNoOf) {
  const d = (r.date || '').replace(/-/g, '') || 'nodate';
  const v = slug_(r.vendor) || 'vendor';
  const a = (r.amount !== null && r.amount !== undefined && r.amount !== '')
    ? Number(r.amount).toFixed(2) + (r.currency || '') : 'noamount';
  const ext = r.mime === 'application/pdf' ? 'pdf' : 'jpg';
  return `${d}_${v}_${a}_${empNoOf(r.uid)}_${r.id.slice(0, 4)}.${ext}`;
}

function csvCell_(v) {
  const s = String(v === null || v === undefined ? '' : v);
  return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
}

function ledgerCsv_(month, receipts, nameOf, empNoOf) {
  const head = ['日期 Date', '經手人 By', '員工編號 EmpNo', '廠商 Vendor', '金額 Amount', '幣別 Currency',
    '類別 Category', '用途 Purpose', '狀態 Status', '稅號 Tax ID', '單據號碼 Invoice No', '檔名 File', '已備份 Backed up'];
  const rows = receipts.filter(r => monthOf_(r) === month)
    .sort((a, b) => (a.date || '').localeCompare(b.date || ''))
    .map(r => [r.date, nameOf(r.uid), empNoOf(r.uid), r.vendor || '',
      (r.amount === null || r.amount === undefined || r.amount === '') ? '' : Number(r.amount).toFixed(2),
      r.currency || '', CAT[r.category] || '', r.note || '', STATUS[r.status] || r.status || '',
      r.tax_id || '', r.invoice_no || '', fileName_(r, empNoOf),
      (r.backedUpAt && r.backupName === fileName_(r, empNoOf)) ? '✓' : '']);
  return '﻿' + [head].concat(rows).map(row => row.map(csvCell_).join(',')).join('\n');
}

/* ---------- Google Drive ---------- */
function getRootFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('ROOT_FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) { /* 被刪了就重建 */ } }
  const it = DriveApp.getFoldersByName(ROOT_FOLDER_NAME);
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder(ROOT_FOLDER_NAME);
  props.setProperty('ROOT_FOLDER_ID', folder.getId());
  return folder;
}

function getMonthFolder_(month) {
  const root = getRootFolder_();
  const it = root.getFoldersByName(month);
  return it.hasNext() ? it.next() : root.createFolder(month);
}

function writeLedger_(month, receipts, nameOf, empNoOf) {
  if (!receipts.some(r => monthOf_(r) === month)) return;
  const folder = getMonthFolder_(month);
  const name = `清冊_${month}.csv`;
  const csv = ledgerCsv_(month, receipts, nameOf, empNoOf);
  const it = folder.getFilesByName(name);
  if (it.hasNext()) it.next().setContent(csv);
  else folder.createFile(name, csv, MimeType.CSV);
}

/* ---------- Firestore（用服務帳號登入） ---------- */
function getAccessToken_() {
  const raw = PropertiesService.getScriptProperties().getProperty('SERVICE_ACCOUNT_JSON');
  if (!raw) throw new Error('找不到指令碼屬性 SERVICE_ACCOUNT_JSON，請先在「專案設定」貼上 Firebase 服務帳號金鑰');
  const sa = JSON.parse(raw);
  const now = Math.floor(Date.now() / 1000);
  const enc = o => Utilities.base64EncodeWebSafe(JSON.stringify(o)).replace(/=+$/, '');
  const input = enc({ alg: 'RS256', typ: 'JWT' }) + '.' + enc({
    iss: sa.client_email, scope: 'https://www.googleapis.com/auth/datastore',
    aud: sa.token_uri || 'https://oauth2.googleapis.com/token', iat: now, exp: now + 3600,
  });
  const sig = Utilities.base64EncodeWebSafe(
    Utilities.computeRsaSha256Signature(input, sa.private_key)).replace(/=+$/, '');
  const res = UrlFetchApp.fetch(sa.token_uri || 'https://oauth2.googleapis.com/token', {
    method: 'post', muteHttpExceptions: true,
    payload: { grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer', assertion: input + '.' + sig },
  });
  if (res.getResponseCode() !== 200) throw new Error('服務帳號登入失敗：' + res.getContentText().slice(0, 300));
  return JSON.parse(res.getContentText()).access_token;
}

function fsFetch_(path, token, opts) {
  const res = UrlFetchApp.fetch(FS + '/' + path, Object.assign({
    headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true,
  }, opts || {}));
  const code = res.getResponseCode();
  if (code === 404) return null;
  if (code !== 200) throw new Error(`Firestore ${path} → HTTP ${code}: ${res.getContentText().slice(0, 300)}`);
  return JSON.parse(res.getContentText());
}

function listCollection_(name, token) {
  const out = [];
  let pageToken = '';
  do {
    const data = fsFetch_(`${name}?pageSize=300${pageToken ? '&pageToken=' + encodeURIComponent(pageToken) : ''}`, token);
    ((data && data.documents) || []).forEach(d => {
      const obj = decodeFields_(d.fields || {});
      obj.id = d.name.split('/').pop();
      out.push(obj);
    });
    pageToken = data && data.nextPageToken;
  } while (pageToken);
  return out;
}

function loadFile_(r, name, token) {
  let data = '';
  for (let i = 0; i < (r.chunks || 1); i++) {
    const d = fsFetch_(`receiptFiles/${r.id}_${i}`, token);
    if (!d) throw new Error(`收據 ${r.id} 缺少檔案分段 ${i}`);
    data += decodeFields_(d.fields || {}).data || '';
  }
  const m = data.match(/^data:([^;]+);base64,([\s\S]*)$/);
  if (!m) throw new Error(`收據 ${r.id} 的檔案格式不正確`);
  return Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], name);
}

function patchDoc_(path, fields, token) {
  const mask = Object.keys(fields).map(k => 'updateMask.fieldPaths=' + encodeURIComponent(k)).join('&');
  fsFetch_(`${path}?${mask}`, token, {
    method: 'patch', contentType: 'application/json', payload: JSON.stringify({ fields }),
  });
}

function decodeValue_(v) {
  if ('stringValue' in v) return v.stringValue;
  if ('integerValue' in v) return Number(v.integerValue);
  if ('doubleValue' in v) return v.doubleValue;
  if ('booleanValue' in v) return v.booleanValue;
  if ('nullValue' in v) return null;
  if ('timestampValue' in v) return v.timestampValue;
  if ('arrayValue' in v) return (v.arrayValue.values || []).map(decodeValue_);
  if ('mapValue' in v) return decodeFields_(v.mapValue.fields || {});
  return null;
}

function decodeFields_(fields) {
  const o = {};
  Object.keys(fields).forEach(k => { o[k] = decodeValue_(fields[k]); });
  return o;
}
