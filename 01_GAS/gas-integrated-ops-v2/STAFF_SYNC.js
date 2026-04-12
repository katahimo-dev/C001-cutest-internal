/**
 * 【統合版】 STAFF_SYNC.js
 * 
 * 機能: 外部 Staff ファイルから最新情報を取得し、
 * gas-integrated-ops-v2 の Staff シートを更新する
 * 
 * ★注意: doget.js の onOpen() と重複しているメニュー定義は、
 * 本ファイル内で呼ぶことで統合する
 */

const STAFF_SYNC_CONFIG = {
  // 参照元（外部マスター Staffファイル）の設定
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

  // 転記先（gas-integrated-ops-v2 の Staff シート）の設定
  DEST_SHEET_NAME: 'Staff', // 転記先のシート名
  DEST_START_ROW: 2,        // 書き込み開始行
  DEST_START_COL: 2,        // 書き込み開始列（B列=2）
  DEST_COL_WIDTH: 5         // 書き込む列の幅（B,C,D,E,F の5列分）
};

/**
 * スプレッドシートを開いたときにメニューを追加する（統合版）
 * 
 * doget.js の onOpen() と合わせることで、UI メニューを整理
 */
function initializeMenus() {
  const ui = SpreadsheetApp.getUi();

  // ✳️ スタッフ情報更新
  ui.createMenu('✳️スタッフ情報更新')
    .addItem('最新情報を取得する', 'syncStaffInfo')
    .addToUi();

  // 📱 アプリ管理（URL送信）
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
    const destSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const destSheet = destSpreadsheet.getSheetByName(STAFF_SYNC_CONFIG.DEST_SHEET_NAME);
    
    if (!destSheet) {
      throw new Error(`転記先シート「${STAFF_SYNC_CONFIG.DEST_SHEET_NAME}」が見つかりません`);
    }

    // 参照元ファイルを検索する
    const files = DriveApp.getFilesByName(STAFF_SYNC_CONFIG.SOURCE_FILE_NAME);
    if (!files.hasNext()) {
      throw new Error(`参照元ファイル「${STAFF_SYNC_CONFIG.SOURCE_FILE_NAME}」が見つかりません`);
    }
    
    const sourceFile = files.next();
    const sourceSpreadsheet = SpreadsheetApp.open(sourceFile);
    
    // シートを取得（指定のシート名、またはデフォルトシート）
    let sourceSheet = sourceSpreadsheet.getSheetByName(STAFF_SYNC_CONFIG.SOURCE_SHEET_NAME);
    if (!sourceSheet) {
      sourceSheet = sourceSpreadsheet.getSheets()[0];
    }

    // 参照元データを読み込む
    const lastRow = sourceSheet.getLastRow();
    if (lastRow < STAFF_SYNC_CONFIG.SOURCE_DATA_START_ROW) {
      ui.alert("ℹ️ 参照元にデータがありません");
      return;
    }
    
    const numRows = lastRow - STAFF_SYNC_CONFIG.SOURCE_DATA_START_ROW + 1;
    
    const sourceRangeValues = sourceSheet.getRange(
      STAFF_SYNC_CONFIG.SOURCE_DATA_START_ROW, 
      1, 
      numRows, 
      STAFF_SYNC_CONFIG.SRC_MAX_COL
    ).getValues();

    // データを指定の列配置に並べ替え
    const outputValues = sourceRangeValues.map(row => {
      return [
        row[STAFF_SYNC_CONFIG.SRC_IDX_ID],     // B列へ: ID
        row[STAFF_SYNC_CONFIG.SRC_IDX_NAME],   // C列へ: 氏名
        row[STAFF_SYNC_CONFIG.SRC_IDX_MAIL],   // D列へ: メール
        row[STAFF_SYNC_CONFIG.SRC_IDX_COL_G],  // E列へ: 入社日
        row[STAFF_SYNC_CONFIG.SRC_IDX_COL_H]   // F列へ: 退職日
      ];
    });

    // 転記先の既存データをクリア
    const destLastRow = destSheet.getLastRow();
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

    ui.alert('✅ 完了', 'スタッフ情報を更新しました', ui.ButtonSet.OK);

  } catch (e) {
    console.error('syncStaffInfo error:', e.message);
    ui.alert('❌ エラー', e.message, ui.ButtonSet.OK);
  }
}
