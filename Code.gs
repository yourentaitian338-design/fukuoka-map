// ===== 福岡おすすめスポットMAP：保存係（Googleスプレッドシート用） =====
// このプログラムは、アプリから届いた「一覧ちょうだい」「スポット追加」「コメント追加」を
// 受け取って、スプレッドシートに読み書きします。合言葉が合わない人には何も返しません。

const SHEET_SPOTS = 'スポット';
const SHEET_COMMENTS = 'コメント';
const SHEET_CONFIG = '設定';
const PLACEHOLDER = 'ここに合言葉を書く';

// 最初に1回だけ実行する：必要なシート（表）を3つ作る
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  makeSheet(ss, SHEET_SPOTS, ['id', '名前', 'ジャンル', '予算', '緯度', '経度', 'Googleマップ', 'おすすめした人', 'ひとこと', '登録日', '削除']);
  makeSheet(ss, SHEET_COMMENTS, ['スポットid', 'ニックネーム', 'コメント', '日付', 'id', '削除', '写真']);
  const config = makeSheet(ss, SHEET_CONFIG, ['合言葉', PLACEHOLDER]);
  config.getRange('A2').setValue('↑ B1のマスに合言葉を書いてください。変えたいときもここを書き換えるだけです。');
}

function makeSheet(ss, name, header) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(header);
    sheet.setFrozenRows(name === SHEET_CONFIG ? 0 : 1);
  }
  return sheet;
}

// ブラウザでURLを開いたときの動作確認用
function doGet() {
  return ContentService.createTextOutput('福岡おすすめスポットMAPの保存係は動いています。');
}

// アプリからの通信はすべてここに届く
function doPost(e) {
  let out;
  try {
    out = handle(JSON.parse(e.postData.contents));
  } catch (err) {
    out = { ok: false, error: String(err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function handle(req) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pass = String(ss.getSheetByName(SHEET_CONFIG).getRange('B1').getValue()).trim();
  if (!pass || pass === PLACEHOLDER) throw new Error('setup');
  if (String(req.pass || '').trim() !== pass) throw new Error('pass');

  ensureColumns(ss);
  let newId = '';
  if (req.action === 'addSpot') newId = addSpot(ss, req);
  else if (req.action === 'addComment') addComment(ss, req);
  else if (req.action === 'deleteSpot') hide(ss, SHEET_SPOTS, 1, 11, req);
  else if (req.action === 'deleteComment') hide(ss, SHEET_COMMENTS, 5, 6, req);
  else if (req.action === 'photo') return getPhoto(ss, req);   // 写真1枚だけ返す（一覧は返さない）
  else if (req.action !== 'list') throw new Error('unknown action');
  return { ok: true, id: newId, spots: listSpots(ss) };
}

function addSpot(ss, req) {
  const name = clean(req.name, 40), by = clean(req.by, 20), hitokoto = clean(req.hitokoto, 100);
  const genre = clean(req.genre, 20), budget = clean(req.budget, 20);
  if (!name || !by || !hitokoto || !genre || !budget) throw new Error('入力が足りません');
  const lat = Number(req.lat), lng = Number(req.lng);
  const hasPos = req.lat !== null && req.lat !== '' && isFinite(lat) && isFinite(lng) && Math.abs(lat) <= 90 && Math.abs(lng) <= 180;
  const gmap = /^https:\/\//.test(String(req.gmap || '')) ? String(req.gmap).slice(0, 300) : '';
  if (!hasPos && !gmap) throw new Error('場所がありません');
  const id = Utilities.getUuid();
  withLock(() => ss.getSheetByName(SHEET_SPOTS).appendRow(
    [id, name, genre, budget, hasPos ? lat : '', hasPos ? lng : '', gmap, by, hitokoto, today()]));
  return id;
}

function addComment(ss, req) {
  const spotId = clean(req.spotId, 60), by = clean(req.by, 20), text = clean(req.text, 100);
  if (!spotId || !by || (!text && !req.photo)) throw new Error('入力が足りません');
  const id = Utilities.getUuid();
  const photoId = req.photo ? savePhoto(req.photo, id) : '';
  withLock(() => ss.getSheetByName(SHEET_COMMENTS).appendRow([spotId, by, text, today(), id, '', photoId]));
}

// ----- 写真 -----
// 写真はあなたのGoogleドライブの「福岡スポットMAP写真」フォルダに保存する（共有はしない）。
// アプリには、合言葉が合う人にだけ、保存係を通して渡す。
const PHOTO_FOLDER = '福岡スポットMAP写真';
const PHOTO_MAX_BYTES = 1500000; // 1枚1.5MBまで（アプリ側で小さくしてから送るので普通は200KBほど）

function photoFolder() {
  const props = PropertiesService.getScriptProperties();
  const saved = props.getProperty('photoFolderId');
  if (saved) { try { return DriveApp.getFolderById(saved); } catch (e) { /* 消されていたら作り直す */ } }
  const folder = DriveApp.createFolder(PHOTO_FOLDER);
  props.setProperty('photoFolderId', folder.getId());
  return folder;
}
function savePhoto(base64, commentId) {
  const bytes = Utilities.base64Decode(String(base64));
  if (bytes.length > PHOTO_MAX_BYTES) throw new Error('写真が大きすぎます');
  if ((bytes[0] & 0xff) !== 0xff || (bytes[1] & 0xff) !== 0xd8) throw new Error('写真の形式が違います'); // JPEGだけ受け付ける
  return photoFolder().createFile(Utilities.newBlob(bytes, 'image/jpeg', today() + '_' + commentId + '.jpg')).getId();
}
// 表に載っている（削除されていない）コメントの写真だけ返す。ドライブの他のファイルは決して返さない
function getPhoto(ss, req) {
  const want = String(req.photoId || '');
  const found = want && rows(ss, SHEET_COMMENTS).some(r => str(r[6] || '') === want && !str(r[5] || ''));
  if (!found) throw new Error('写真が見つかりません');
  const blob = DriveApp.getFileById(want).getBlob();
  return { ok: true, data: Utilities.base64Encode(blob.getBytes()), type: 'image/jpeg' };
}

// 削除は「非表示」にするだけ。「削除」の列に日付と名前が入る。
// 元に戻したいときは、スプレッドシートでその行の「削除」のマスを空にする。
function hide(ss, sheetName, idCol, delCol, req) {
  const id = String(req.id || '').trim();
  const by = clean(req.by, 20) || '（名前なし）';
  if (!id) throw new Error('入力が足りません');
  withLock(() => {
    const sheet = ss.getSheetByName(sheetName);
    const ids = sheet.getRange(1, idCol, sheet.getLastRow(), 1).getValues();
    for (let i = 1; i < ids.length; i++) {
      if (str(ids[i][0]) === id) { sheet.getRange(i + 1, delCol).setValue(today() + ' ' + by + 'が削除'); return; }
    }
  });
}

// 削除機能を足す前に作った表にも、必要な列（見出しとコメントのid）を足す
function ensureColumns(ss) {
  const spots = ss.getSheetByName(SHEET_SPOTS), comments = ss.getSheetByName(SHEET_COMMENTS);
  if (spots.getRange(1, 11).getValue() !== '' && comments.getRange(1, 5).getValue() !== '' && comments.getRange(1, 7).getValue() !== '') {
    const last = comments.getLastRow();
    if (last < 2 || comments.getRange(2, 5, last - 1, 1).getValues().every(r => r[0] !== '')) return;
  }
  withLock(() => {
    if (spots.getRange(1, 11).getValue() === '') spots.getRange(1, 11).setValue('削除');
    if (comments.getRange(1, 5).getValue() === '') comments.getRange(1, 5, 1, 2).setValues([['id', '削除']]);
    if (comments.getRange(1, 7).getValue() === '') comments.getRange(1, 7).setValue('写真');
    const last = comments.getLastRow();
    if (last < 2) return;
    const range = comments.getRange(2, 5, last - 1, 1);
    range.setValues(range.getValues().map(r => [r[0] === '' ? Utilities.getUuid() : r[0]]));
  });
}

function listSpots(ss) {
  const comments = {};
  rows(ss, SHEET_COMMENTS).filter(r => !str(r[5] || '')).forEach(r => {
    const id = str(r[0]);
    (comments[id] = comments[id] || []).push({ id: str(r[4] || ''), by: str(r[1]), text: str(r[2]), date: str(r[3]), photo: str(r[6] || '') });
  });
  return rows(ss, SHEET_SPOTS).filter(r => str(r[0]) && !str(r[10] || '')).map(r => ({
    id: str(r[0]), name: str(r[1]), genre: str(r[2]), budget: str(r[3]),
    lat: r[4] === '' ? null : Number(r[4]), lng: r[5] === '' ? null : Number(r[5]),
    gmap: str(r[6]), by: str(r[7]), hitokoto: str(r[8]), date: str(r[9]),
    comments: comments[str(r[0])] || [],
  })).reverse(); // 新しい順
}

// ----- 小さな道具 -----
function rows(ss, name) {
  const v = ss.getSheetByName(name).getDataRange().getValues();
  return v.slice(1); // 1行目は見出しなので外す
}
function str(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM-dd');
  return String(v).trim();
}
// 文字数を制限し、表計算の「式」として動いてしまう文字で始まる入力を無害にする
function clean(v, max) {
  const s = String(v == null ? '' : v).trim().slice(0, max);
  return /^[=+\-@]/.test(s) ? ' ' + s : s;
}
function today() {
  return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
}
// 2人が同時に書き込んでも表が壊れないように、順番待ちをさせる
function withLock(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try { fn(); } finally { lock.releaseLock(); }
}
