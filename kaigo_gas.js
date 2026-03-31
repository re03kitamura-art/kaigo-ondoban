// ================================================================
// 介護記録システム - Google Apps Script
// ================================================================
// 【セットアップ手順】
// 1. script.google.com で「新しいプロジェクト」を作成
// 2. このコードを貼り付け
// 3. FOLDER_ID を自分のGoogleドライブのフォルダIDに変更
// 4. 「デプロイ」→「新しいデプロイ」→種類「ウェブアプリ」
//    ・実行ユーザー：自分
//    ・アクセス：全員（Googleアカウント不要）
// 5. 発行されたURLをWebアプリの GAS_URL に貼り付ける
// ================================================================

// ▼▼▼ ここを自分のフォルダIDに変更 ▼▼▼
// Googleドライブでフォルダを開いたときのURLの末尾の文字列
// 例: https://drive.google.com/drive/folders/1ABC123XYZ
//                                             ↑これがID
const FOLDER_ID = "ここにフォルダIDを貼り付ける";
// ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

// ================================================================
// シート定義（列ヘッダー）
// ================================================================
const SHEET_HEADERS = {
  "食事・水分": ["記録日時", "利用者名", "部屋", "食事区分", "摂取内容", "水分量", "特記事項"],
  "排泄記録":   ["記録日時", "利用者名", "部屋", "回数集計", "時刻別ログ", "浣腸・下剤", "特記事項"],
  "看護記録":   ["記録日時", "利用者名", "部屋", "バイタル",  "観察所見",   "特記事項"],
};

// ================================================================
// HTTPエントリーポイント
// ================================================================

function doPost(e) {
  try {
    const data   = JSON.parse(e.postData.contents);
    const action = data.action;
    let result;

    if (action === "addRecord") {
      result = addRecord(data);
    } else if (action === "ping") {
      result = { ok: true, message: "GAS接続OK" };
    } else {
      result = { ok: false, error: "不明なアクション: " + action };
    }

    return jsonResponse(result);
  } catch (err) {
    return jsonResponse({ ok: false, error: err.toString() });
  }
}

// GETリクエスト（接続テスト）
function doGet(e) {
  return jsonResponse({
    ok: true,
    message: "介護記録GAS起動中",
    time: new Date().toLocaleString("ja-JP")
  });
}

// ================================================================
// メイン処理：記録をスプレッドシートに追記
// ================================================================
function addRecord(data) {
  const { residentId, residentName, sheetName, row } = data;

  // 利用者ブックを取得（なければ新規作成）
  const ss = getOrCreateBook(residentId, residentName);

  // 対象シートを取得（なければ新規作成）
  const sheet = getOrCreateSheet(ss, sheetName);

  // データ追記
  sheet.appendRow(row);

  // 書式整形
  formatNewRow(sheet, sheetName, sheet.getLastRow());

  return { ok: true, residentId, sheetName, totalRows: sheet.getLastRow() - 1 };
}

// ================================================================
// スプレッドシート管理
// ================================================================

function getOrCreateBook(residentId, residentName) {
  const folder   = DriveApp.getFolderById(FOLDER_ID);
  const fileName = residentName + "_介護記録";

  // 既存ファイルを検索
  const files = folder.getFilesByName(fileName);
  if (files.hasNext()) {
    return SpreadsheetApp.open(files.next());
  }

  // 新規作成
  const ss = SpreadsheetApp.create(fileName);

  // デフォルトシートを「食事・水分」に
  const sheet1 = ss.getSheets()[0];
  sheet1.setName("食事・水分");
  setHeader(sheet1, "食事・水分");

  // 排泄記録シート
  const sheet2 = ss.insertSheet("排泄記録");
  setHeader(sheet2, "排泄記録");

  // 看護記録シート
  const sheet3 = ss.insertSheet("看護記録");
  setHeader(sheet3, "看護記録");

  // 指定フォルダに移動
  const file = DriveApp.getFileById(ss.getId());
  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);

  return ss;
}

function getOrCreateSheet(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    setHeader(sheet, sheetName);
  }
  return sheet;
}

// ================================================================
// 書式設定
// ================================================================

function setHeader(sheet, sheetName) {
  const headers = SHEET_HEADERS[sheetName];
  if (!headers) return;

  const range = sheet.getRange(1, 1, 1, headers.length);
  range.setValues([headers]);
  range
    .setBackground("#2A6045")
    .setFontColor("#FFFFFF")
    .setFontWeight("bold")
    .setFontSize(10)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(1, 28);
  sheet.setFrozenRows(1);

  // 列幅設定
  const widths = {
    "食事・水分": [130, 80, 55, 55, 200, 70, 150],
    "排泄記録":   [130, 80, 55, 100, 250, 80, 150],
    "看護記録":   [130, 80, 55, 200, 120, 150],
  };
  const w = widths[sheetName] || [];
  w.forEach((width, i) => sheet.setColumnWidth(i + 1, width));
}

function formatNewRow(sheet, sheetName, rowNum) {
  const headers = SHEET_HEADERS[sheetName];
  if (!headers || rowNum < 2) return;

  const range = sheet.getRange(rowNum, 1, 1, headers.length);

  // 行の背景色（交互）
  range.setBackground(rowNum % 2 === 0 ? "#EBF4EE" : "#FFFFFF");
  range.setFontSize(10).setVerticalAlignment("middle");
  sheet.setRowHeight(rowNum, 22);

  // 日時列の書式
  sheet.getRange(rowNum, 1).setNumberFormat("yyyy/mm/dd hh:mm");
}

// ================================================================
// ユーティリティ
// ================================================================

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ================================================================
// テスト関数（スクリプトエディタから手動実行して確認できます）
// ================================================================

function testPing() {
  Logger.log("フォルダアクセステスト...");
  const folder = DriveApp.getFolderById(FOLDER_ID);
  Logger.log("フォルダ名: " + folder.getName());
  Logger.log("OK - GASは正常に動作しています");
}

function testFood() {
  const result = addRecord({
    residentId:   "r01",
    residentName: "山田 花子",
    sheetName:    "食事・水分",
    row: ["2025/3/26 08:00", "山田 花子", "101号室", "朝食", "主食:全量　副食:全量", "800ml", ""]
  });
  Logger.log(JSON.stringify(result));
}

function testToilet() {
  const result = addRecord({
    residentId:   "r01",
    residentName: "山田 花子",
    sheetName:    "排泄記録",
    row: ["2025/3/26 10:00", "山田 花子", "101号室", "排尿2回 排便1回", "08:00 排尿、09:00 排便（普通便）、10:00 排尿", "", ""]
  });
  Logger.log(JSON.stringify(result));
}

function testNurse() {
  const result = addRecord({
    residentId:   "r01",
    residentName: "山田 花子",
    sheetName:    "看護記録",
    row: ["2025/3/26 14:00", "山田 花子", "101号室", "体温36.5℃　血圧128/78　SpO2 98%　脈拍72", "良眠", ""]
  });
  Logger.log(JSON.stringify(result));
}
