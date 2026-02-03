/**
 * メンテナンス用：AM列（訪問等回数）算定式一括更新スクリプト
 * 機能: 管理シートにある全スタッフのファイルを開き、全月シートのAM列に新しい数式を適用する
 * 
 * このスクリプトは、既に運用している出勤簿の一部を書き換えるものです。数式を変えたい場合は、config_fixの内容を書き換えてください。
 * 列の追加等は、自動転記GASの内容も書き換えないと、挙動が変わってしまうことがあります。
 */

const CONFIG_FIX = {
  // 管理シートの設定（自動転記GASと合わせる）
  STAFF_START_ROW: 2,       // データ開始行
  COL_LINK_2025: 6,         // G列 (2025年度リンク)
  
  // 更新対象の設定
  TARGET_SHEET_NAMES: ["4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月", "1月", "2月", "3月"],
  TARGET_COL_AM: 39,        // AM列は39番目
  TARGET_START_ROW: 4,      // 4行目から開始
  TARGET_ROWS: 31,          // 31行分（1ヶ月分）更新する
  
  // 新しい数式（4行目用）
  // ※GASで一括入力すると、5行目以降は自動的にAH5, AH6...とズレて入力されます
  NEW_FORMULA: '=IF(ISNUMBER(AH4), 3, IF(ISNUMBER(AG4), 2, IF(OR(ISNUMBER(AI4), ISNUMBER(AJ4)), 1, 0)))'
};

// function onOpen() {
//   SpreadsheetApp.getUi()
//     .createMenu('🔧 メンテナンス')
//     .addItem('【注意】AM列の数式を一括更新', 'updateFormulaAllFiles')
//     .addToUi();
// }

function updateFormulaAllFiles() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '数式一括更新の実行',
    '管理シートにある「全スタッフ」のファイルを開き、\n全月のシートのAM列(訪問回数)の数式を強制的に上書きします。\n\n実行してもよろしいですか？',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const listSheet = ss.getSheets()[0]; // 一番左のシート（名簿）
  const lastRow = listSheet.getLastRow();
  
  // 名簿データの取得
  const staffData = listSheet.getRange(CONFIG_FIX.STAFF_START_ROW, 1, lastRow - CONFIG_FIX.STAFF_START_ROW + 1, 10).getValues();
  
  let successCount = 0;
  let errorCount = 0;
  let log = "";

  // トースト通知で開始を知らせる
  ss.toast("更新処理を開始しました。完了までそのままお待ちください...", "処理中", -1);

  // スタッフごとのループ
  for (let i = 0; i < staffData.length; i++) {
    const row = staffData[i];
    const linkUrl = row[CONFIG_FIX.COL_LINK_2025 - 1]; // G列

    // URL抽出（HYPERLINK関数対策は簡易的に文字列チェック）
    let fileUrl = linkUrl;
    // もしセルの中身が数式なら、URLを抽出する処理を入れることも可能ですが、
    // ここでは値としてURLが入っている、またはHYPERLINKの表示テキストではなく値を取得できている前提で進めます。
    // ※もし文字列としてURLが取れない場合は、getFormulasで取る必要がありますが、前回のコードと同様の処理であれば以下で通ります。
    
    // URLの簡単なチェック
    if (!fileUrl || String(fileUrl).indexOf("http") === -1) {
      // 数式からURLを取り出す必要がある場合（念のため）
      const formula = listSheet.getRange(CONFIG_FIX.STAFF_START_ROW + i, CONFIG_FIX.COL_LINK_2025).getFormula();
      if (formula && formula.includes("http")) {
        const match = formula.match(/"(https?:\/\/[^"]+)"/);
        if (match) fileUrl = match[1];
      }
    }

    if (!fileUrl || String(fileUrl).indexOf("http") === -1) continue; // URLがない行はスキップ

    try {
      const targetSS = SpreadsheetApp.openByUrl(fileUrl);
      const staffName = targetSS.getName(); // ファイル名取得（ログ用）
      
      // 月ごとのループ
      CONFIG_FIX.TARGET_SHEET_NAMES.forEach(sheetName => {
        const sheet = targetSS.getSheetByName(sheetName);
        if (sheet) {
          // 数式を一括セット (AM4:AM34)
          // setFormulaに4行目の式を渡すと、範囲全体に対して自動的に行番号をインクリメントして適用してくれます
          sheet.getRange(CONFIG_FIX.TARGET_START_ROW, CONFIG_FIX.TARGET_COL_AM, CONFIG_FIX.TARGET_ROWS, 1)
               .setFormula(CONFIG_FIX.NEW_FORMULA);

          // ★追加：AO4セルの値を消去（AO列は41番目）
          sheet.getRange("AO4").clearContent();

        }
      });

      successCount++;
      console.log(`[OK] ${staffName}`);

    } catch (e) {
      errorCount++;
      log += `[NG] 行${CONFIG_FIX.STAFF_START_ROW + i}: ${e.message}\n`;
      console.error(`[NG] 行${CONFIG_FIX.STAFF_START_ROW + i}: ${e.message}`);
    }
  }

  ss.toast("処理が完了しました。", "完了", 5);
  
  ui.alert(
    '処理完了',
    `成功: ${successCount}件\nエラー: ${errorCount}件\n\n${errorCount > 0 ? "エラー詳細:\n" + log : ""}`,
    ui.ButtonSet.OK
  );
}