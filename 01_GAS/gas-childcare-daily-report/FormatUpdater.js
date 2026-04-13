/**
 * 管理シートの構造に合わせ、37行目（文字含む）もそっくりコピーして一括更新する
 */
function updateAllStaffAttendanceSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const listSheet = ss.getSheetByName("表紙"); 
  const masterSheet = ss.getSheetByName("シートマスター"); 
  
  // 【設定：行番号の指定】
  const HEADER_ROW = 8;
  const START_ROW = 9;
  const TARGET_COLUMN = 7; // G列(2025年度)
  
  const MONTHS = ["4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月", "1月", "2月", "3月"];
  
  // 1. 年度ラベルを取得
  const yearLabel = listSheet.getRange(HEADER_ROW, TARGET_COLUMN).getValue().toString().trim();
  if (!yearLabel) {
    Browser.msgBox("エラー：年度が取得できません。");
    return;
  }
  
  // 2. マスターから「4行目(日次)」、「35行目(合計)」の計算式情報を取得
  // ※37行目は文字なのでここではなく、下の関数内で copyTo を使って丸ごとコピーします
  const lastCol = masterSheet.getLastColumn();
  const row4Formulas = masterSheet.getRange(4, 1, 1, lastCol).getFormulasR1C1()[0];
  const row35Formulas = masterSheet.getRange(35, 1, 1, lastCol).getFormulasR1C1()[0];
  
  // 3. 管理シートをスキャン
  const lastRow = listSheet.getLastRow();
  if (lastRow < START_ROW) return;
  const listData = listSheet.getRange(START_ROW, 1, lastRow - START_ROW + 1, TARGET_COLUMN).getValues();
  
  let processedCount = 0;

  listData.forEach((row) => {
    const staffName = row[2];
    const cellStatus = row[TARGET_COLUMN - 1];

    if (!staffName || !cellStatus || cellStatus === "-" || cellStatus.toString().includes("退職済")) {
      return;
    }

    const targetFileName = staffName + "_出勤簿_" + yearLabel;
    const files = DriveApp.getFilesByName(targetFileName);
    
    if (files.hasNext()) {
      const targetSS = SpreadsheetApp.open(files.next());
      console.log("更新中: " + targetFileName);
      
      const tempMaster = masterSheet.copyTo(targetSS);
      
      MONTHS.forEach(month => {
        const monthSheet = targetSS.getSheetByName(month);
        if (monthSheet) {
          // 数式データを渡して更新
          updateSingleSheet(monthSheet, tempMaster, row4Formulas, row35Formulas);
        }
      });
      
      targetSS.deleteSheet(tempMaster);
      processedCount++;
    }
  });

  Browser.msgBox(processedCount + " 名分の出勤簿を更新しました。");
}

/**
 * 内部関数：35行目の合計式を区別して適用し、37行目をそっくりコピー
 */
function updateSingleSheet(targetSheet, tempMaster, row4Formulas, row35Formulas) {
  const lastCol = tempMaster.getLastColumn();
  
  // 1. ヘッダー（1-3行目）をコピー
  tempMaster.getRange(1, 1, 3, lastCol).copyTo(targetSheet.getRange(1, 1, 3, lastCol));
  
  // 2. データ行（4行目〜34行目）に日次数式を適用
  for (let col = 0; col < row4Formulas.length; col++) {
    if (row4Formulas[col] !== "") {
      // 4行目から34行目（31日間分）に適用
      const dataRange = targetSheet.getRange(4, col + 1, 31, 1);
      dataRange.setFormulaR1C1(row4Formulas[col]);
    }
    
    // 3. 合計行（35行目）に専用の数式を適用
    if (row35Formulas[col] !== "") {
      const totalCell = targetSheet.getRange(35, col + 1);
      totalCell.setFormulaR1C1(row35Formulas[col]);
    }
  }

  // ★追加：4. 37行目（文字や書式を含む）をそっくりコピー
  tempMaster.getRange(37, 1, 1, lastCol).copyTo(targetSheet.getRange(37, 1, 1, lastCol));
}
