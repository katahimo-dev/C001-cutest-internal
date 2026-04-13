/**
 * メンテナンス用：X列・AA列 プルダウンメニュー更新 ＆ 不要な色ルール削除スクリプト
 * 機能: 全ファイルのX列・AA列の選択肢を更新し、過去に設定してしまった条件付き書式をクリアする
 * Ver.Batch_ResetColor_Fixed (getBackgroundエラー修正版)
 */

const CONFIG_FIX2 = {
  // 管理シートの設定
  STAFF_START_ROW: 2,       // データ開始行
  COL_LINK_2025: 6,         // 6 (F列)
  COL_NAME: 3,              // C列 (氏名)
  
  // 更新対象の設定
  TARGET_SHEET_NAMES: ["4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月", "1月", "2月", "3月"],
  
  TARGET_START_ROW: 4,      // 4行目から開始
  TARGET_ROWS: 31,          // 31行分
  
  // 対象の列番号 (A=1, X=24, AA=27)
  COL_X: 24,
  COL_AA: 27,
  
  // 新しいプルダウンの選択肢
  NEW_OPTIONS: [
    "カスタマーサポート業務",
    "広報業務",
    "新聞チェック",
    "月例会（オンライン参加）",
    "Cutest Shere（週次mtg）",
    "チームmtg",
    "リーダーmtg",
    "その他事務作業"
  ]
};

// function onOpen() {
//   SpreadsheetApp.getUi()
//     .createMenu('🔧 メンテナンス')
//     .addItem('【注意】X・AA列のプルダウンを一括更新(色リセット)', 'updateDropdownAllFiles')
//     .addToUi();
// }

function updateDropdownAllFiles() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'プルダウン一括更新＆色リセットの実行',
    '管理シートのG列(7列目)にあるリンク先を開き、以下の処理を行います。\n\n1. X列・AA列のプルダウン選択肢を更新\n2. X列・AA列に残っている不要な色付けルール(条件付き書式)を削除\n\n実行してもよろしいですか？',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const listSheet = ss.getSheets()[0]; // 一番左のシート（名簿）
  const lastRow = listSheet.getLastRow();
  
  if (lastRow < CONFIG_FIX2.STAFF_START_ROW) {
    ui.alert("データ行が見つかりません。");
    return;
  }

  // --- データの入力規則（プルダウン）を作成 ---
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(CONFIG_FIX2.NEW_OPTIONS) // リストを指定
    .setAllowInvalid(true) // リスト外の入力を許可する
    .build();

  // --- データ一括取得 ---
  const numRows = lastRow - CONFIG_FIX2.STAFF_START_ROW + 1;
  const range = listSheet.getRange(CONFIG_FIX2.STAFF_START_ROW, 1, numRows, 10); 
  const values = range.getValues();
  const formulas = range.getFormulas();
  
  let successCount = 0;
  let errorCount = 0;
  let skipCount = 0;
  let log = "";

  ss.toast("更新処理を開始しました...", "処理中", -1);

  for (let i = 0; i < values.length; i++) {
    const rowValues = values[i];
    const rowFormulas = formulas[i];
    
    const staffName = rowValues[CONFIG_FIX2.COL_NAME - 1] || "不明なスタッフ";
    const linkVal = rowValues[CONFIG_FIX2.COL_LINK_2025 - 1];
    const linkFormula = rowFormulas[CONFIG_FIX2.COL_LINK_2025 - 1];

    if (!rowValues[CONFIG_FIX2.COL_NAME - 1]) continue;

    let fileUrl = null;

    // URL抽出
    if (linkFormula && linkFormula.includes("http")) {
      const match = linkFormula.match(/"(https?:\/\/[^"]+)"/);
      if (match) fileUrl = match[1];
    }
    if (!fileUrl && linkVal && String(linkVal).includes("http")) {
      fileUrl = linkVal;
    }

    if (!fileUrl) {
      console.warn(`[SKIP] ${staffName}: G列に有効なリンクがありません`);
      skipCount++;
      continue;
    }

    try {
      const targetSS = SpreadsheetApp.openByUrl(fileUrl);
      console.log(`更新中: ${staffName}`);

      CONFIG_FIX2.TARGET_SHEET_NAMES.forEach(sheetName => {
        const sheet = targetSS.getSheetByName(sheetName);
        if (sheet) {
          // --- 1. プルダウン適用 ---
          sheet.getRange(CONFIG_FIX2.TARGET_START_ROW, CONFIG_FIX2.COL_X, CONFIG_FIX2.TARGET_ROWS, 1)
               .setDataValidation(rule);
          
          sheet.getRange(CONFIG_FIX2.TARGET_START_ROW, CONFIG_FIX2.COL_AA, CONFIG_FIX2.TARGET_ROWS, 1)
               .setDataValidation(rule);
               
          // --- 2. 不要な条件付き書式をクリア（修正箇所） ---
          const rules = sheet.getConditionalFormatRules();
          const newRules = [];
          
          // 削除したい色のリスト（前回コードで設定した色）
          const removeColors = ["#d9ead3", "#cfe2f3", "#fff2cc", "#ead1dc", "#f4cccc", "#fce5cd", "#d0e0e3", "#eeeeee"];
          
          for (let r = 0; r < rules.length; r++) {
            const currentRule = rules[r];
            const booleanCondition = currentRule.getBooleanCondition(); // ★ここを修正

            // グラデーションルールなどは booleanCondition が null になるので、それは削除せず残す
            if (!booleanCondition) {
              newRules.push(currentRule);
              continue;
            }

            // 背景色を取得
            const bg = booleanCondition.getBackground();

            // 削除リストに含まれていない色なら、残す（nullの場合は色設定がないので残す）
            if (bg == null || !removeColors.includes(bg.toLowerCase())) {
              newRules.push(currentRule);
            }
          }
          
          sheet.setConditionalFormatRules(newRules);
        }
      });

      successCount++;

    } catch (e) {
      errorCount++;
      log += `[NG] ${staffName}: ${e.message}\n`;
    }
  }

  ss.toast("すべての処理が完了しました。", "完了", 5);
  
  ui.alert(
    '処理完了',
    `成功: ${successCount}件\nエラー: ${errorCount}件\nスキップ: ${skipCount}件\n\n${errorCount > 0 ? "エラー詳細:\n" + log : ""}`,
    ui.ButtonSet.OK
  );
}
