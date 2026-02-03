/**
 * 構成設定（Config）
 * スタッフ情報の同期設定
 */
const STAFF_SYNC_CONFIG = {
  // ■ 参照元（Staff）の設定
  SOURCE_FILE_NAME: 'Staff',  // 検索するファイル名
  SOURCE_SHEET_NAME: 'Staff', // シート名
  SOURCE_DATA_START_ROW: 2,            // データ開始行
  
  // 読み込む列のインデックス（A列=0, B列=1, ... H列=7, I列=8）
  SRC_IDX_ID: 0,      // A列: ID
  SRC_IDX_NAME: 1,    // B列: 氏名
  SRC_IDX_HIRE: 6,    // G列: 入社日
  SRC_IDX_RESIGN: 7,  // H列: 退職日
  SRC_MAX_COL: 9,     // 取得する最大列数（A～I列まで読み込むため 9）

  // ■ 転記先（デモ管理用シート）の設定
  DEST_SHEET_NAME: '表紙', // 転記先のシート名
  DEST_START_ROW: 9,       // 書き込み開始行
  DEST_START_COL: 2,       // 書き込み開始列（B列=2）
  DEST_COL_WIDTH: 4        // 書き込む列の幅（B,C,D,E の4列分）
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
 * ID, 氏名, 入社日, 退職日 を取得して反映します
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

    // 3. 参照元データの読み込み（A列～I列まで一括取得）
    const lastRow = sourceSheet.getLastRow();
    if (lastRow < STAFF_SYNC_CONFIG.SOURCE_DATA_START_ROW) {
      ui.alert("参照元にデータがありません。");
      return;
    }
    
    const numRows = lastRow - STAFF_SYNC_CONFIG.SOURCE_DATA_START_ROW + 1;
    
    // A列(1)からI列(9)までの範囲を取得
    const sourceRangeValues = sourceSheet.getRange(
      STAFF_SYNC_CONFIG.SOURCE_DATA_START_ROW, 
      1, 
      numRows, 
      STAFF_SYNC_CONFIG.SRC_MAX_COL
    ).getValues();

    // 必要なデータ（ID, 氏名, 入社日, 退職日）だけを抽出・整形
    const outputValues = sourceRangeValues.map(row => {
      return [
        row[STAFF_SYNC_CONFIG.SRC_IDX_ID],     // ID
        row[STAFF_SYNC_CONFIG.SRC_IDX_NAME],   // 氏名
        row[STAFF_SYNC_CONFIG.SRC_IDX_HIRE],   // 入社日 (H列) -> D列へ
        row[STAFF_SYNC_CONFIG.SRC_IDX_RESIGN]  // 退職日 (I列) -> E列へ
      ];
    });

    // 4. 転記先への書き込み（B列～E列）
    const destLastRow = destSheet.getLastRow();
    
    // 既存データをクリア（B9以降の B,C,D,E列 をクリア）
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

    // 完了通知
    destSpreadsheet.toast('ID, 氏名, 入社日, 退職日を更新しました', '更新完了');

  } catch (e) {
    console.error(e);
    ui.alert('エラーが発生しました: ' + e.message);
  }
}