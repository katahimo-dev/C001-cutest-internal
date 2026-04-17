/**
 * 構成設定（Config）
 * スタッフ情報の同期設定
 */
const STAFF_SYNC_CONFIG = {
  // ■ 参照元（Staff）の設定
  SOURCE_FILE_NAME: 'Staff',  // 検索するファイル名
  SOURCE_SHEET_NAME: 'Staff', // シート名
  SOURCE_DATA_START_ROW: 2,            // データ開始行
  
  // 読み込む列のインデックス（A列=0, B列=1, ... K列=10）
  SRC_IDX_ID: 0,       // A列: ID
  SRC_IDX_NAME: 1,     // B列: 氏名
  SRC_IDX_EMAIL: 4,    // E列: mail address
  SRC_IDX_HIRE: 6,     // G列: 入社日
  SRC_IDX_RESIGN: 7,   // H列: 退職日
  SRC_IDX_ADMIN: 10,   // K列: 管理者
  SRC_MAX_COL: 11,     // 取得する最大列数（A～K列まで読み込むため 11）

  // ■ 転記先（デモ管理用シート）の設定
  DEST_SHEET_NAME: '表紙', // 転記先のシート名
  DEST_START_ROW: 9,       // 書き込み開始行
  DEST_START_COL: 1,       // 書き込み開始列（A列=1）
  DEST_COL_WIDTH: 6,       // 書き込む列の幅（A～F列の6列分）
  DEST_LINK_START_COL: 7   // 出勤簿リンク開始列（G列）
};

// /**
//  * メニューを追加する関数
//  */
// function onOpen() {
//   SpreadsheetApp.getUi()
//     .createMenu('★スタッフ情報更新')
//     .addItem('最新情報を取得する', 'syncStaffInfo')
//     .addToUi();
// }

/**
 * スタッフ情報を同期するメイン関数
 * 管理者, ID, 氏名, メールアドレス, 入社日, 退職日 を取得して反映します
 */
function syncStaffInfo() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // 1. 転記先（このスプレッドシート）のシートを取得
    const destSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const destSheet = destSpreadsheet.getSheetByName(STAFF_SYNC_CONFIG.DEST_SHEET_NAME);
    
    if (!destSheet) {
      throw new Error(`転記先シート「${STAFF_SYNC_CONFIG.DEST_SHEET_NAME}」が見つかりません。`);
    }

    // 2. 参照元（Staff_20251224）ファイルを検索して開く
    const files = DriveApp.getFilesByName(STAFF_SYNC_CONFIG.SOURCE_FILE_NAME);
    if (!files.hasNext()) {
      throw new Error(`参照元ファイル「${STAFF_SYNC_CONFIG.SOURCE_FILE_NAME}」が見つかりません。`);
    }
    const sourceFile = files.next();
    const sourceSpreadsheet = SpreadsheetApp.open(sourceFile);
    
    // シート取得（名前で見つからない場合は1番目のシート）
    let sourceSheet = sourceSpreadsheet.getSheetByName(STAFF_SYNC_CONFIG.SOURCE_SHEET_NAME);
    if (!sourceSheet) {
      sourceSheet = sourceSpreadsheet.getSheets()[0];
    }

    // 3. 参照元データの読み込み（A列～K列まで一括取得）
    const lastRow = sourceSheet.getLastRow();
    if (lastRow < STAFF_SYNC_CONFIG.SOURCE_DATA_START_ROW) {
      ui.alert("参照元にデータがありません。");
      return;
    }
    
    const numRows = lastRow - STAFF_SYNC_CONFIG.SOURCE_DATA_START_ROW + 1;
    const today = toDateAtNoon_(new Date());
    
    // A列(1)からK列(11)までの範囲を取得
    const sourceRangeValues = sourceSheet.getRange(
      STAFF_SYNC_CONFIG.SOURCE_DATA_START_ROW, 
      1, 
      numRows, 
      STAFF_SYNC_CONFIG.SRC_MAX_COL
    ).getValues();

    // 必要なデータ（管理者, ID, 氏名, メールアドレス, 入社日, 退職日）だけを抽出・整形
    const outputValues = sourceRangeValues.map(row => {
      const rawResignDate = row[STAFF_SYNC_CONFIG.SRC_IDX_RESIGN];
      const isRetired = isRetired_(rawResignDate, today);

      return [
        isRetired ? '' : row[STAFF_SYNC_CONFIG.SRC_IDX_ADMIN], // 管理者
        row[STAFF_SYNC_CONFIG.SRC_IDX_ID],     // ID
        row[STAFF_SYNC_CONFIG.SRC_IDX_NAME],   // 氏名
        row[STAFF_SYNC_CONFIG.SRC_IDX_EMAIL],  // メールアドレス
        row[STAFF_SYNC_CONFIG.SRC_IDX_HIRE],   // 入社日
        rawResignDate                          // 退職日
      ];
    });

    const retiredEmails = collectRetiredEmails_(sourceRangeValues, today);

    // 4. 転記先への書き込み（A列～F列）
    const destLastRow = destSheet.getLastRow();
    
    // 既存データをクリア（A9以降の A～F列 をクリア）
    if (destLastRow >= STAFF_SYNC_CONFIG.DEST_START_ROW) {
      destSheet.getRange(
        STAFF_SYNC_CONFIG.DEST_START_ROW, 
        STAFF_SYNC_CONFIG.DEST_START_COL, 
        destLastRow - STAFF_SYNC_CONFIG.DEST_START_ROW + 1, 
        STAFF_SYNC_CONFIG.DEST_COL_WIDTH
      ).clearContent();
    }

    // 新しいデータを書き込み
    destSheet.getRange(
      STAFF_SYNC_CONFIG.DEST_START_ROW, 
      STAFF_SYNC_CONFIG.DEST_START_COL, 
      outputValues.length, 
      STAFF_SYNC_CONFIG.DEST_COL_WIDTH
    ).setValues(outputValues);

    revokeEditorsFromAttendanceFiles_(destSheet, retiredEmails);

    // 完了通知
    destSpreadsheet.toast('管理者, ID, 氏名, メールアドレス, 入社日, 退職日を更新しました', '更新完了');

  } catch (e) {
    console.error(e);
    ui.alert('エラーが発生しました: ' + e.message);
  }
}

function collectRetiredEmails_(sourceRangeValues, today) {
  const retiredEmails = [];
  const seen = {};

  for (let i = 0; i < sourceRangeValues.length; i++) {
    const row = sourceRangeValues[i];
    const email = row[STAFF_SYNC_CONFIG.SRC_IDX_EMAIL];
    const resignDate = row[STAFF_SYNC_CONFIG.SRC_IDX_RESIGN];

    if (!isRetired_(resignDate, today)) continue;
    if (!email || String(email).trim() === '') continue;

    const normalizedEmail = String(email).trim();
    if (seen[normalizedEmail]) continue;

    seen[normalizedEmail] = true;
    retiredEmails.push(normalizedEmail);
  }

  return retiredEmails;
}

function revokeEditorsFromAttendanceFiles_(destSheet, retiredEmails) {
  if (!retiredEmails || retiredEmails.length === 0) return;

  const lastRow = destSheet.getLastRow();
  const lastCol = destSheet.getLastColumn();
  if (lastRow < STAFF_SYNC_CONFIG.DEST_START_ROW) return;
  if (lastCol < STAFF_SYNC_CONFIG.DEST_LINK_START_COL) return;

  const numRows = lastRow - STAFF_SYNC_CONFIG.DEST_START_ROW + 1;
  const numCols = lastCol - STAFF_SYNC_CONFIG.DEST_LINK_START_COL + 1;
  const linkRange = destSheet.getRange(
    STAFF_SYNC_CONFIG.DEST_START_ROW,
    STAFF_SYNC_CONFIG.DEST_LINK_START_COL,
    numRows,
    numCols
  );
  const linkValues = linkRange.getValues();
  const linkFormulas = linkRange.getFormulas();
  const fileUrls = collectAttendanceFileUrls_(linkValues, linkFormulas);

  for (let i = 0; i < fileUrls.length; i++) {
    const fileUrl = fileUrls[i];

    try {
      const file = DriveApp.getFileById(extractFileIdFromUrl_(fileUrl));
      removeEditorsSafely_(file, retiredEmails);
    } catch (e) {
      console.warn('権限削除スキップ: ' + fileUrl + ' / ' + e.message);
    }
  }
}

function collectAttendanceFileUrls_(linkValues, linkFormulas) {
  const urls = [];
  const seen = {};

  for (let rowIndex = 0; rowIndex < linkValues.length; rowIndex++) {
    const valueRow = linkValues[rowIndex];
    const formulaRow = linkFormulas[rowIndex];

    for (let colIndex = 0; colIndex < valueRow.length; colIndex++) {
      const formula = formulaRow[colIndex];
      const value = valueRow[colIndex];
      const url = extractUrlFromCell_(formula, value);

      if (!url || seen[url]) continue;

      seen[url] = true;
      urls.push(url);
    }
  }

  return urls;
}

function extractUrlFromCell_(formula, value) {
  if (formula && formula.indexOf('http') !== -1) {
    const match = formula.match(/"(https?:\/\/[^\"]+)"/);
    if (match) return match[1];
  }

  if (value && String(value).indexOf('http') !== -1) {
    return String(value);
  }

  return '';
}

function extractFileIdFromUrl_(url) {
  const match = String(url).match(/[-\w]{25,}/);
  if (!match) {
    throw new Error('ファイルIDを抽出できませんでした');
  }

  return match[0];
}

function removeEditorsSafely_(file, emails) {
  for (let i = 0; i < emails.length; i++) {
    const email = emails[i];

    try {
      file.removeEditor(email);
    } catch (e) {
      console.warn('編集権限削除失敗: ' + email + ' / ' + e.message);
    }
  }
}

function isRetired_(rawResignDate, today) {
  const resignDate = parseDateValue_(rawResignDate);
  if (!resignDate) return false;

  return today.getTime() > resignDate.getTime();
}

function parseDateValue_(rawValue) {
  if (!rawValue) return null;

  if (rawValue instanceof Date) {
    return toDateAtNoon_(rawValue);
  }

  if (typeof rawValue === 'string' && rawValue.trim() !== '') {
    const normalized = rawValue.replace(/\./g, '/').replace(/-/g, '/');
    const parsed = new Date(normalized);
    if (!isNaN(parsed.getTime())) {
      return toDateAtNoon_(parsed);
    }
  }

  return null;
}

function toDateAtNoon_(date) {
  const normalized = new Date(date);
  normalized.setHours(12, 0, 0, 0);
  return normalized;
}