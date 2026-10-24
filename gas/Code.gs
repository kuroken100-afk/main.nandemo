/**
 * せどりノート: Gmailのメルカリ通知を自動で読み取り、スプレッドシートに保存して、
 * index.html（せどりノートアプリ）へ配信する Apps Script です。
 *
 * ■ セットアップ手順
 * 1. Googleスプレッドシートを新規作成し、「拡張機能」→「Apps Script」を開く
 * 2. このファイルの中身をすべて貼り付けて保存
 * 3. 「プロジェクトの設定」→「スクリプト プロパティ」に追加
 *      プロパティ: TOKEN   値: 推測されにくい長い文字列（例: パスワードマネージャーで生成）
 * 4. 関数 syncMail を一度手動で実行し、Gmailの読み取り権限を許可
 * 5. 「トリガー」→「トリガーを追加」
 *      実行する関数: syncMail / イベントのソース: 時間主導型 / 30分おき
 * 6. 「デプロイ」→「新しいデプロイ」→ 種類: ウェブアプリ
 *      次のユーザーとして実行: 自分 / アクセスできるユーザー: 全員
 *      発行された URL の末尾に ?token=（手順3の TOKEN の値）を付けたものを
 *      せどりノートの「自動取得」欄に貼り付けて「今すぐ同期」
 */

// 取り込み対象のメール検索条件（メルカリの「取引完了」「購入」通知）
const SEARCH_QUERY = 'from:no-reply@mercari.jp newer_than:60d (subject:取引が完了しました OR subject:購入しました)';
const SHEET_NAME = 'sales';
const HEADER = ['externalId', 'name', 'soldPrice', 'soldDate', 'status'];

// 時間トリガーから呼ばれる: 新しいメール分をシートに追加する（商品IDが既にあれば追加しない）
function syncMail() {
  const sheet = getSheet_();
  const known = new Set(sheet.getDataRange().getValues().slice(1).map(r => String(r[0])));
  // 検索結果は新しい順。最初に見つかったものを残すので、新しい日付が優先される
  for (const thread of GmailApp.search(SEARCH_QUERY, 0, 100)) {
    for (const msg of thread.getMessages()) {
      const date = Utilities.formatDate(msg.getDate(), 'Asia/Tokyo', 'yyyy-MM-dd');
      for (const it of parseMercariMail_(msg.getPlainBody())) {
        if (known.has(it.externalId)) continue;
        sheet.appendRow([it.externalId, it.name, it.soldPrice, date, 'sold']);
        known.add(it.externalId);
      }
    }
  }
}

// せどりノート（index.html）から呼ばれる: 売却済み一覧をJSONで返す
function doGet(e) {
  const token = PropertiesService.getScriptProperties().getProperty('TOKEN');
  if (!token || !e.parameter || e.parameter.token !== token) {
    return json_({ error: 'unauthorized' });
  }
  const items = getSheet_().getDataRange().getValues().slice(1)
    .filter(r => r[0])
    .map(r => ({
      externalId: String(r[0]),
      name: String(r[1]),
      soldPrice: Number(r[2]),
      soldDate: r[3] instanceof Date ? Utilities.formatDate(r[3], 'Asia/Tokyo', 'yyyy-MM-dd') : String(r[3]),
      status: String(r[4]),
    }));
  return json_({ items: items });
}

// メルカリ通知本文から 商品ID・商品名・商品価格 を抜き出す（index.html の parseMercariMail と同じ規則）
function parseMercariMail_(text) {
  const re = /商品ID\s*[:：]\s*(m\d+)[\s\S]*?商品名\s*[:：]\s*([\s\S]*?)\s*商品価格\s*[:：]\s*([\d,]+)\s*円/g;
  const out = [];
  let m;
  while ((m = re.exec(text))) {
    out.push({ externalId: m[1], name: m[2].replace(/\s+/g, ' ').trim(), soldPrice: Number(m[3].replace(/,/g, '')) });
  }
  return out;
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADER);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
