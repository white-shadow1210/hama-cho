/**
 * 販売・仕入れ管理 — スプレッドシート側のプログラム（Google Apps Script）
 *
 * 使い方（くわしくは README.md）
 *  1. 新しいGoogleスプレッドシートを作る
 *  2. 拡張機能 → Apps Script を開き、このファイルの中身をすべて貼り付ける
 *  3. 下の PASSCODE を家族で決めた合言葉に変えて保存
 *  4. 上のメニューで「setup」を選んで ▶実行（シートが自動でできます）
 *  5. デプロイ → 新しいデプロイ → 種類「ウェブアプリ」
 *       次のユーザーとして実行：自分 ／ アクセスできるユーザー：全員
 *  6. 表示された「ウェブアプリのURL」をアプリの設定に入れる
 */

// ▼ 家族で決めた合言葉（setup を実行すると保存されます。変えたら setup をもう一度実行）
const PASSCODE = 'ここを合言葉に変える';

const TZ = 'Asia/Tokyo';

// 列の並び（左から）。 t=文字 n=数値 b=はい/いいえ
const SCHEMA = {
  items: { sheet: '品目', cols: [
    ['id','ID','t'],['name','品目名','t'],['cat','種類','t'],['unit','単位','t'],['price','販売単価','n'],
    ['active','扱い中','b'],['lastCost','前回仕入単価','n'],['lastCostDate','前回仕入日','t'],['created','登録日','t'],
    ['packKg','1個の重さ(kg)','n'],['yieldPct','歩留まり目安(%)','n'] ]},
  customers: { sheet: '顧客', cols: [
    ['id','ID','t'],['name','お名前','t'],['phone','連絡先','t'],['area','地域','t'],['since','取引開始日','t'] ]},
  purchases: { sheet: '仕入れ', cols: [
    ['id','ID','t'],['date','日付','t'],['dow','曜日','t'],['itemId','品目ID','t'],['itemName','品目','t'],
    ['qty','数量','n'],['unit','単位','t'],['unitCost','仕入単価','n'],['total','仕入金額','n'],['note','メモ','t'],['ts','登録時刻','n'] ]},
  processes: { sheet: '加工', cols: [
    ['id','ID','t'],['batchId','加工ID','t'],['date','日付','t'],['fromItemId','原料ID','t'],['fromItemName','原料','t'],
    ['fromQty','使った量','n'],['fromUnit','原料の単位','t'],['toItemId','商品ID','t'],['toItemName','できあがり','t'],
    ['toQty','できた数','n'],['toUnit','単位','t'],['note','メモ','t'],['ts','登録時刻','n'] ]},
  sales: { sheet: '販売', cols: [
    ['id','ID','t'],['orderId','注文ID','t'],['date','日付','t'],['dow','曜日','t'],['customerId','顧客ID','t'],
    ['customerName','お客さま','t'],['area','地域','t'],['itemId','品目ID','t'],['itemName','品目','t'],['qty','数量','n'],
    ['unit','単位','t'],['listPrice','いつもの単価','n'],['unitPrice','販売単価','n'],['total','販売金額','n'],
    ['discount','値引き額','n'],['mode','売り方','t'],['pay','支払い','t'],['paidDate','入金日','t'],['note','メモ','t'],['ts','登録時刻','n'] ]}
};

/* ---------- 最初に1回だけ実行 ---------- */
function setup() {
  if (!PASSCODE || PASSCODE === 'ここを合言葉に変える') throw new Error('先に PASSCODE を合言葉に変えてください');
  PropertiesService.getScriptProperties().setProperty('PASSCODE', PASSCODE);
  const ss = SpreadsheetApp.getActive();
  Object.keys(SCHEMA).forEach(function (t) {
    const def = SCHEMA[t];
    let sh = ss.getSheetByName(def.sheet);
    if (!sh) sh = ss.insertSheet(def.sheet);
    const head = def.cols.map(function (c) { return c[1]; });
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#1d3a5f').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    def.cols.forEach(function (c, i) {
      const col = sh.getRange(2, i + 1, sh.getMaxRows() - 1, 1);
      if (c[2] === 't') col.setNumberFormat('@');           // 電話番号の先頭0などが消えないように
      if (c[2] === 'n') col.setNumberFormat('#,##0.##');
    });
    sh.hideColumns(1); // ID列は隠す
  });
  const first = ss.getSheetByName('シート1');
  if (first && first.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(first);
  return 'セットアップ完了';
}

/* ---------- 受け口 ---------- */
function doGet(e) {
  return respond(function () {
    const p = e.parameter || {};
    checkKey(p.key);
    if (p.action === 'ping') return { ok: true };
    if (p.action === 'load') return load(p.month, p.today);
    throw new Error('不明な操作です');
  });
}

function doPost(e) {
  return respond(function () {
    const body = JSON.parse(e.postData.contents);
    checkKey(body.key);
    const lock = LockService.getScriptLock();
    lock.waitLock(25000);
    try {
      const ops = body.ops || [];
      ops.forEach(applyOp);
      SpreadsheetApp.flush();
      return { ok: true, applied: ops.length };
    } finally {
      lock.releaseLock();
    }
  });
}

function respond(fn) {
  let out;
  try { out = fn(); } catch (err) { out = { ok: false, error: String(err && err.message || err) }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function checkKey(k) {
  const pass = PropertiesService.getScriptProperties().getProperty('PASSCODE');
  if (!pass) throw new Error('スプレッドシート側で setup がまだ実行されていません');
  if (String(k || '') !== pass) throw new Error('合言葉がちがいます');
}

/* ---------- 読み込み ---------- */
function load(month, today) {
  if (!/^\d{4}-\d{2}$/.test(month || '')) throw new Error('月の指定がおかしいです');
  const prev = prevMonth(month);
  const t = /^\d{4}-\d{2}-\d{2}$/.test(today || '') ? today : Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  const since = addDays(t, -120); // 「いつもの品目」用に直近120日
  const inRange = function (r) { const m = String(r.date || '').slice(0, 7); return m === month || m === prev; };
  return {
    ok: true,
    items: readAll('items'),
    customers: readAll('customers'),
    purchases: readAll('purchases').filter(inRange),
    processes: readAll('processes').filter(inRange),
    sales: readAll('sales').filter(function (r) { return inRange(r) || String(r.date) >= since || r.pay === '掛け'; })
  };
}

function readAll(table) {
  const def = SCHEMA[table];
  const sh = SpreadsheetApp.getActive().getSheetByName(def.sheet);
  if (!sh || sh.getLastRow() < 2) return [];
  const vals = sh.getRange(2, 1, sh.getLastRow() - 1, def.cols.length).getValues();
  const out = [];
  vals.forEach(function (row) {
    if (!row[0]) return;
    const o = {};
    def.cols.forEach(function (c, i) {
      let v = row[i];
      if (v instanceof Date) v = Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
      if (c[2] === 'n') v = (v === '' || v === null) ? '' : Number(v);
      if (c[2] === 'b') v = !(v === false || v === 'FALSE' || v === 'false');
      if (c[2] === 't') v = (v === null || v === undefined) ? '' : String(v);
      o[c[0]] = v;
    });
    out.push(o);
  });
  return out;
}

/* ---------- 書き込み ---------- */
function applyOp(op) {
  const def = SCHEMA[op.table];
  if (!def) throw new Error('不明な表: ' + op.table);
  const sh = SpreadsheetApp.getActive().getSheetByName(def.sheet);
  if (!sh) throw new Error('シートがありません: ' + def.sheet + '（setup を実行してください）');
  const id = String(op.id || '');
  if (!id) throw new Error('IDがありません');
  const row = findRow(sh, id);
  if (op.op === 'delete') { if (row) sh.deleteRow(row); return; }
  if (op.op !== 'upsert') throw new Error('不明な操作: ' + op.op);
  const d = op.data || {};
  const values = def.cols.map(function (c) {
    const v = c[0] === 'id' ? id : d[c[0]];
    if (c[2] === 'n') return (v === '' || v === null || v === undefined || isNaN(Number(v))) ? '' : Number(v);
    if (c[2] === 'b') return v === false || v === 'FALSE' ? false : true;
    return v === null || v === undefined ? '' : String(v);
  });
  if (row) sh.getRange(row, 1, 1, values.length).setValues([values]);
  else sh.appendRow(values);
}

function findRow(sh, id) {
  if (sh.getLastRow() < 2) return 0;
  const hit = sh.getRange(2, 1, sh.getLastRow() - 1, 1).createTextFinder(id).matchEntireCell(true).findNext();
  return hit ? hit.getRow() : 0;
}

/* ---------- 日付 ---------- */
function prevMonth(m) { const y = +m.slice(0, 4), mm = +m.slice(5, 7); const d = new Date(y, mm - 2, 1); return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2); }
function addDays(s, n) { const p = s.split('-').map(Number); const d = new Date(p[0], p[1] - 1, p[2] + n); return Utilities.formatDate(d, TZ, 'yyyy-MM-dd'); }
