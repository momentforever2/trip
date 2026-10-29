/**
 * 家族旅行 app 的後端（Google Apps Script）
 * 資料存在這份試算表的「data」工作表：每一列是一筆資料（col / id / json / updatedAt）。
 *
 * 安裝：擴充功能 → Apps Script → 貼上本檔 → 部署 → 新增部署作業 → 網頁應用程式
 *       執行身分：我　／　誰可以存取：任何人
 * 收據辨識：專案設定 → 指令碼屬性 → 新增 GEMINI_API_KEY（Google AI Studio 申請的金鑰）
 *          可選：GEMINI_MODEL（預設 gemini-3.8-flash）
 */

var SHEET = 'data';
var COLS = ['trip', 'days', 'pool', 'topups', 'splits', 'spots'];
var PHOTO_FOLDER = '家族旅行 收據照片';

function sheet_() {
  var ss = SpreadsheetApp.getActive();
  var s = ss.getSheetByName(SHEET);
  if (!s) {
    s = ss.insertSheet(SHEET);
    s.getRange(1, 1, 1, 4).setValues([['col', 'id', 'json', 'updatedAt']]);
    s.getRange('C:C').setNumberFormat('@');
  }
  return s;
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function props_() { return PropertiesService.getScriptProperties(); }
function aiKey_() { return props_().getProperty('GEMINI_API_KEY'); }
function aiModel_() { return props_().getProperty('GEMINI_MODEL') || 'gemini-3.8-flash'; }

/* ---------- GET：讀取全部資料 ---------- */
function doGet(e) {
  try {
    return out_({ ok: true, data: readAll_(), ai: !!aiKey_(), ts: Date.now() });
  } catch (err) {
    return out_({ ok: false, error: String(err && err.message || err) });
  }
}

function readAll_() {
  var s = sheet_();
  var last = s.getLastRow();
  var data = {};
  COLS.forEach(function (c) { data[c] = []; });
  if (last < 2) return data;
  var v = s.getRange(2, 1, last - 1, 4).getValues();
  for (var i = 0; i < v.length; i++) {
    var col = String(v[i][0]), id = String(v[i][1]);
    if (!col || !id || !data[col]) continue;
    try {
      var doc = JSON.parse(v[i][2]);
      doc.id = id;
      data[col].push(doc);
    } catch (x) { /* 略過壞掉的列 */ }
  }
  return data;
}

/* ---------- POST：寫入、刪除、AI ---------- */
function doPost(e) {
  var body;
  try { body = JSON.parse(e.postData.contents); }
  catch (x) { return out_({ ok: false, error: 'bad_request' }); }
  try {
    switch (body.action) {
      case 'save': return out_(save_(body.col, body.id, body.data));
      case 'delete': return out_(remove_(body.col, body.id));
      case 'receipt': return out_(receipt_(body));
      case 'ai': return out_({ ok: true, result: gemini_([{ text: String(body.prompt || '').slice(0, 20000) }]) });
      default: return out_({ ok: false, error: 'unknown_action' });
    }
  } catch (err) {
    return out_({ ok: false, error: String(err && err.message || err) });
  }
}

function check_(col, id) {
  if (COLS.indexOf(col) < 0) throw new Error('bad_collection');
  if (!/^[A-Za-z0-9_-]{1,40}$/.test(String(id))) throw new Error('bad_id');
}

function findRow_(s, col, id) {
  var last = s.getLastRow();
  if (last < 2) return 0;
  var v = s.getRange(2, 1, last - 1, 2).getValues();
  for (var i = 0; i < v.length; i++) {
    if (String(v[i][0]) === col && String(v[i][1]) === String(id)) return i + 2;
  }
  return 0;
}

function save_(col, id, data) {
  check_(col, id);
  if (!data || typeof data !== 'object') throw new Error('bad_data');
  delete data.id;
  var json = JSON.stringify(data);
  if (json.length > 45000) throw new Error('too_large');
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var s = sheet_();
    var row = findRow_(s, col, id);
    var now = Date.now();
    var vals = [[col, "'" + id, json, now]];
    if (row) s.getRange(row, 1, 1, 4).setValues(vals);
    else s.getRange(s.getLastRow() + 1, 1, 1, 4).setValues(vals);
    return { ok: true, u: now };
  } finally {
    lock.releaseLock();
  }
}

function remove_(col, id) {
  check_(col, id);
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var s = sheet_();
    var row = findRow_(s, col, id);
    if (row) s.deleteRow(row);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- Gemini ---------- */
function gemini_(parts) {
  var key = aiKey_();
  if (!key) throw new Error('no_key');
  var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + aiModel_() + ':generateContent';
  var level = props_().getProperty('GEMINI_THINKING') || 'minimal'; // 讓模型少想一點，回得比較快
  var call = function (withThinking) {
    var gen = { responseMimeType: 'application/json' };
    if (withThinking && level !== 'off') gen.thinkingConfig = { thinkingLevel: level };
    return UrlFetchApp.fetch(url, {
      method: 'post', contentType: 'application/json', headers: { 'x-goog-api-key': key }, muteHttpExceptions: true,
      payload: JSON.stringify({ contents: [{ role: 'user', parts: parts }], generationConfig: gen })
    });
  };
  var res = call(true);
  var code = res.getResponseCode();
  if (code === 400 && /think/i.test(res.getContentText())) { res = call(false); code = res.getResponseCode(); }
  if (code === 429) throw new Error('ai_busy');
  if (code !== 200) throw new Error('ai_error ' + code + ': ' + res.getContentText().slice(0, 300));
  var j = JSON.parse(res.getContentText());
  var text = '';
  var c = j.candidates && j.candidates[0];
  if (c && c.content && c.content.parts) c.content.parts.forEach(function (p) { if (p.text && !p.thought) text += p.text; });
  if (!text) throw new Error('ai_empty');
  return parseJson_(text);
}

function parseJson_(t) {
  try { return JSON.parse(t); } catch (x) {}
  var m = t.match(/```(?:json)?\s*([\s\S]*?)```/);
  if (m) { try { return JSON.parse(m[1]); } catch (x) {} }
  var a = t.search(/[\[{]/), b = Math.max(t.lastIndexOf('}'), t.lastIndexOf(']'));
  if (a >= 0 && b > a) return JSON.parse(t.slice(a, b + 1));
  throw new Error('ai_bad_json');
}

var RECEIPT_PROMPT = [
  '你是家族日本旅遊的記帳助理。附上的照片是一張收據、發票或消費明細，可能是日文、英文或中文。',
  '請讀出內容，並把店名與品名翻譯成台灣慣用的繁體中文。只回傳一個 JSON 物件，格式如下：',
  '{"store":"收據上的原文店名","storeZh":"店名中文（例如：一蘭拉麵 新宿店）","date":"YYYY-MM-DD，看不出來就給空字串",',
  '"currency":"JPY 或 TWD","total":含稅總金額數字,"category":"餐飲|交通|門票|購物|藥妝|住宿|其他 擇一",',
  '"items":[{"orig":"原文品名","zh":"中文品名","price":該品項小計數字}],',
  '"note":"值得記下的資訊，例如免稅、稅額、付款方式、點數折抵，沒有就空字串"}',
  '規則：金額只寫數字，不含符號與逗號；items 最多 40 項，看不清楚的品項可略過；不要編造照片上沒有的內容。'
].join('\n');

function receipt_(b) {
  if (!b.image) throw new Error('no_image');
  var mime = b.mime || 'image/jpeg';
  var result = gemini_([{ text: RECEIPT_PROMPT }, { inline_data: { mime_type: mime, data: b.image } }]);
  var photoUrl = '';
  try {
    var folders = DriveApp.getFoldersByName(PHOTO_FOLDER);
    var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(PHOTO_FOLDER);
    var name = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd_HHmmss') + '.jpg';
    var file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(b.image), mime, name));
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (x) {}
    photoUrl = file.getUrl();
  } catch (x) { /* 照片沒存成功不影響辨識 */ }
  return { ok: true, result: result, photoUrl: photoUrl };
}

/* ---------- 在編輯器裡手動執行，用來確認設定 ---------- */
function testSetup() {
  Logger.log('資料筆數：' + JSON.stringify(Object.keys(readAll_()).map(function (k) { return k + '=' + readAll_()[k].length; })));
  if (!aiKey_()) { Logger.log('尚未設定 GEMINI_API_KEY，收據辨識不能用。'); return; }
  var t0 = Date.now();
  var r = gemini_([{ text: '只回傳 JSON：{"ok":true,"msg":"你好"}' }]);
  Logger.log('Gemini 回應：' + JSON.stringify(r) + '（花了 ' + Math.round((Date.now() - t0) / 1000) + ' 秒）');
  DriveApp.getRootFolder(); // 觸發雲端硬碟授權
}
