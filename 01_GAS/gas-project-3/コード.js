/**
 * 手動修正マーク（背景色変更）
 * 条件：ファイル名に「出勤簿」が含まれ、かつシート名が「1月」〜「12月」の場合
 */
// データが始まる行番号（例：ヘッダーが1行目なら、データは2行目から）
const START_ROW_FOR_DATA = 2; 

function onEdit(e) {
  // オブジェクトが存在しない場合（スクリプトエディタから直接実行した場合など）は終了
  if (!e) return;

  const range = e.range;
  const sheet = range.getSheet();
  const sheetName = sheet.getName();
  
  // 1. ファイル名チェック
  // ★修正: 「〜で終わる(endsWith)」から「〜を含む(includes)」に変更しました
  const fileName = SpreadsheetApp.getActiveSpreadsheet().getName();
  
  if (!fileName.includes("出勤簿")) return;

  // 2. シート名チェック
  // 「1月」〜「12月」のいずれかであるかチェック（正規表現）
  const monthPattern = /^(1[0-2]|[1-9])月$/;
  if (!monthPattern.test(sheetName)) return;

  // 3. 行番号チェック
  // データ開始行より上の行（ヘッダーなど）は無視
  if (range.getRow() < START_ROW_FOR_DATA) return;

  // 4. 背景色を薄赤に変更
  // 編集されたセル範囲に対して色を設定します
  range.setBackground("#fce4e4"); 
}

/**
 * ファイルを開いたときに、今月のシートを自動で開くスクリプト
 */
function onOpen() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 今日の日付から「X月」を作成
  const today = new Date();
  const month = today.getMonth() + 1; // 0始まりなので+1
  const sheetName = month + "月";
  
  // その月のシートが存在すればアクティブにする
  const targetSheet = ss.getSheetByName(sheetName);
  if (targetSheet) {
    targetSheet.activate();
  }
}