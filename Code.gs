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
  makeSheet(ss, SHEET_SPOTS, ['id', '名前', 'ジャンル', '予算', '緯度', '経度', 'Googleマップ', 'おすすめした人', 'ひとこと', '登録日']);
  makeSheet(ss, SHEET_COMMENTS, ['スポットid', 'ニックネーム', 'コメント', '日付']);
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

  let newId = '';
  if (req.action === 'addSpot') newId = addSpot(ss, req);
  else if (req.action === 'addComment') addComment(ss, req);
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
  if (!spotId || !by || !text) throw new Error('入力が足りません');
  withLock(() => ss.getSheetByName(SHEET_COMMENTS).appendRow([spotId, by, text, today()]));
}

function listSpots(ss) {
  const comments = {};
  rows(ss, SHEET_COMMENTS).forEach(r => {
    const id = str(r[0]);
    (comments[id] = comments[id] || []).push({ by: str(r[1]), text: str(r[2]), date: str(r[3]) });
  });
  return rows(ss, SHEET_SPOTS).filter(r => str(r[0])).map(r => ({
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
