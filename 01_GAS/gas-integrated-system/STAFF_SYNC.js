/**
 * ファイル名: StaffSync
 * 機能: 外部Staffファイルから最新情報を取得し、デモ管理用シートを更新する
 */

const STAFF_SYNC_CONFIG = {
  // ■ 参照元（Staffファイル）の設定
  SOURCE_FILE_NAME: 'Staff',           // 検索するファイル名
  SOURCE_SHEET_NAME: 'Staff_20251224', // シート名
  SOURCE_DATA_START_ROW: 2,            // データ開始行
  
  // 読み込む列のインデックス（A列=0 スタート）
  SRC_IDX_ID: 0,      // A列: ID
  SRC_IDX_NAME: 1,    // B列: 氏名
  SRC_IDX_MAIL: 4,    // E列: メールアドレス
  SRC_IDX_COL_G: 6,   // G列: 入社日（想定）
  SRC_IDX_COL_H: 7,   // H列: 退職日（想定）
  
  SRC_MAX_COL: 8,     // 取得する最大列数（A～H列まで読み込むため 8列目まで必要）

  // ■ 転記先（デモ管理用シート）の設定
  DEST_SHEET_NAME: 'Staff', // 転記先のシート名
  DEST_START_ROW: 9,       // 書き込み開始行
  DEST_START_COL: 2,       // 書き込み開始列（B列=2）
  DEST_COL_WIDTH: 5        // 書き込む列の幅（B,C,D,E,F の5列分）
};

/**
 * メニューを追加する関数
 */

// スプレッドシートを開いたときにメニューを追加する（統合版）
function onOpen() {
  var ui = SpreadsheetApp.getUi();

  // 1つ目のメニュー：スタッフ情報更新
  ui.createMenu('✳️スタッフ情報更新')
    .addItem('最新情報を取得する', 'syncStaffInfo')
    .addToUi();

  // 2つ目のメニュー：アプリ管理（URL送信）
  ui.createMenu('📱アプリURL送付')
    .addItem('☑️ チェックした人にURLを送信', 'sendAppUrlToSelected')
    .addToUi();
}



/**
 * スタッフ情報を同期するメイン関数
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

    // 2. 参照元ファイルを検索して開く
    const files = DriveApp.getFilesByName(STAFF_SYNC_CONFIG.SOURCE_FILE_NAME);
    if (!files.hasNext()) {
      throw new Error(`参照元ファイル「${STAFF_SYNC_CONFIG.SOURCE_FILE_NAME}」が見つかりません。`);
    }
    const sourceFile = files.next();
    const sourceSpreadsheet = SpreadsheetApp.open(sourceFile);
    
    // シート取得
    let sourceSheet = sourceSpreadsheet.getSheetByName(STAFF_SYNC_CONFIG.SOURCE_SHEET_NAME);
    if (!sourceSheet) {
      sourceSheet = sourceSpreadsheet.getSheets()[0];
    }

    // 3. 参照元データの読み込み
    const lastRow = sourceSheet.getLastRow();
    if (lastRow < STAFF_SYNC_CONFIG.SOURCE_DATA_START_ROW) {
      ui.alert("参照元にデータがありません。");
      return;
    }
    
    const numRows = lastRow - STAFF_SYNC_CONFIG.SOURCE_DATA_START_ROW + 1;
    
    // A列(1)からH列(8)までの範囲を取得
    const sourceRangeValues = sourceSheet.getRange(
      STAFF_SYNC_CONFIG.SOURCE_DATA_START_ROW, 
      1, 
      numRows, 
      STAFF_SYNC_CONFIG.SRC_MAX_COL
    ).getValues();

    // ★マッピング処理：指定の列配置にデータを並べ替えます
    const outputValues = sourceRangeValues.map(row => {
      return [
        row[STAFF_SYNC_CONFIG.SRC_IDX_ID],    // B列へ: ID (A列)
        row[STAFF_SYNC_CONFIG.SRC_IDX_NAME],  // C列へ: 氏名 (B列)
        row[STAFF_SYNC_CONFIG.SRC_IDX_MAIL],  // D列へ: メール (E列)
        row[STAFF_SYNC_CONFIG.SRC_IDX_COL_G], // E列へ: 入社日 (G列)
        row[STAFF_SYNC_CONFIG.SRC_IDX_COL_H]  // F列へ: 退職日 (H列)
      ];
    });

    // 4. 転記先への書き込み
    const destLastRow = destSheet.getLastRow();
    
    // 既存データをクリア（B列～F列の範囲）
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

    destSpreadsheet.toast('スタッフ情報を更新しました', '更新完了');

  } catch (e) {
    console.error(e);
    ui.alert('エラーが発生しました: ' + e.message);
  }
}
