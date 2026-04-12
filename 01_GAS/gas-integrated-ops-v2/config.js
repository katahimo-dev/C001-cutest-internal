/**
 * 【統合版】共通設定ファイル
 * gas-integrated-ops-v2 用の集約設定
 * 各プロジェクトの CONFIG を統一管理して、重複コード削減 and 保守性向上
 */

// ==========================================
// 【基本設定】
// ==========================================
const CONFIG_UNIFIED = {
  // ■ 年度管理
  YEAR: '2025',                // 対象年度（会計年度: 4月～3月）
  YEAR_START_MONTH: 4,         // 会計年度開始月
  
  // ■ スプレッドシート設定
  STAFF_SHEET_NAME: 'Staff',   // スタッフマスタ シート名
  STAFF_SHEET_ID: '1exqD69qZqACm9KOUPpa0fVWRYD2qEZfce7I6TOs_VDk', // 既設スタッフマスター
  SUMMARY_SHEET_ID: '(TBD)', // ⚠️ NEW: 月間集約シート（作成後にID変更）
  ROUTE_FOLDER_ID: '1wLjR6iZ447tbUa3ff59bejoM5aXh8clC', // ルート集計ファイルを保存するフォルダ
  TEMPLATE_FILE_NAME: '出勤簿テンプレート',
  
  // ■ Google Calendar (カレンダー同期用)
  CALENDAR_ID: 'primary',  // 全社カレンダー（通常 primary）
  
  // ■ LineWorks 設定（オプション）
  USE_LINEWORKS: false,    // true にすると LINE WORKS 通知が有効
  LW: {
    BOT_ID: PropertiesService.getScriptProperties().getProperty('LW_BOT_ID'),
    // その他認証情報は STAFF_SYNC.js で管理
  },
  
  // ■ UI設定
  APP_TITLE: 'スマホde勤務管理',    // アプリタイトル
  APP_ICON: '📱',
};

const RUNTIME_CONFIG = {
  SUMMARY_SHEET_ID_PROP: 'SUMMARY_SHEET_ID',
  SUMMARY_SHEET_FILE_NAME: '月間集約_勤務管理'
};

// ==========================================
// 【スプレッドシート列定義】
// ==========================================
// 月間集約シートの列構造
// A列: 日付
// B列: スタッフ名
// C-E列: ♯1訪問先（訪問先名, 開始時刻, 終了時刻）
// F-H列: ♯2訪問先
// I-K列: ♯3訪問先
// L-N列: 作業内容・時間
// O列: 天候
// P列: 事務フラグ
// Q列: 修正者名（管理者が最終確定時に記入）
// R列: 修正日時

const SUMMARY_SHEET_COLUMNS = {
  DATE: 'A',
  STAFF_NAME: 'B',
  
  // 訪問先 #1
  VISIT_1_NAME: 'C',
  VISIT_1_START: 'D',
  VISIT_1_END: 'E',
  
  // 訪問先 #2
  VISIT_2_NAME: 'F',
  VISIT_2_START: 'G',
  VISIT_2_END: 'H',
  
  // 訪問先 #3
  VISIT_3_NAME: 'I',
  VISIT_3_START: 'J',
  VISIT_3_END: 'K',
  
  // 作業
  WORK_TYPE: 'L',
  WORK_START: 'M',
  WORK_END: 'N',
  
  // その他
  WEATHER: 'O',
  IS_ADMIN_WORK: 'P',     // 事務作業フラグ
  CONFIRMED_BY: 'Q',      // 管理者名
  CONFIRMED_AT: 'R',      // 確定日時
};

// ==========================================
// 【UI用: 入力フォーム項目】
// ==========================================
// gas-integrated-system の INPUT_COLUMNS から、月間集約対応に修正
const INPUT_COLUMNS_SUMMARY = {
  [SUMMARY_SHEET_COLUMNS.VISIT_1_NAME]: 
    {type: 'text', label: '訪問先①', readonly: true},
  [SUMMARY_SHEET_COLUMNS.VISIT_1_START]: 
    {type: 'time', label: '開始①'},
  [SUMMARY_SHEET_COLUMNS.VISIT_1_END]: 
    {type: 'time', label: '終了①'},
    
  [SUMMARY_SHEET_COLUMNS.VISIT_2_NAME]: 
    {type: 'text', label: '訪問先②', readonly: true},
  [SUMMARY_SHEET_COLUMNS.VISIT_2_START]: 
    {type: 'time', label: '開始②'},
  [SUMMARY_SHEET_COLUMNS.VISIT_2_END]: 
    {type: 'time', label: '終了②'},
    
  [SUMMARY_SHEET_COLUMNS.VISIT_3_NAME]: 
    {type: 'text', label: '訪問先③', readonly: true},
  [SUMMARY_SHEET_COLUMNS.VISIT_3_START]: 
    {type: 'time', label: '開始③'},
  [SUMMARY_SHEET_COLUMNS.VISIT_3_END]: 
    {type: 'time', label: '終了③'},
    
  [SUMMARY_SHEET_COLUMNS.WORK_TYPE]: 
    {type: 'select', label: '作業内容'},
  [SUMMARY_SHEET_COLUMNS.WORK_START]: 
    {type: 'time', label: '作業開始'},
  [SUMMARY_SHEET_COLUMNS.WORK_END]: 
    {type: 'time', label: '作業終了'},
    
  [SUMMARY_SHEET_COLUMNS.WEATHER]: 
    {type: 'select', label: '天候'},
  [SUMMARY_SHEET_COLUMNS.IS_ADMIN_WORK]: 
    {type: 'select', label: '事務作業'},
};

// ==========================================
// 【ユーティリティ】
// ==========================================

/**
 * 月間集約スプレッドシートIDを取得。
 * 未設定時は自動作成し、Script Properties に保存する。
 * @returns {string}
 */
function getSummarySpreadsheetId() {
  const props = PropertiesService.getScriptProperties();
  const runtimeId = props.getProperty(RUNTIME_CONFIG.SUMMARY_SHEET_ID_PROP);
  if (runtimeId && runtimeId !== '') {
    return runtimeId;
  }

  if (CONFIG_UNIFIED.SUMMARY_SHEET_ID && CONFIG_UNIFIED.SUMMARY_SHEET_ID !== '(TBD)') {
    props.setProperty(RUNTIME_CONFIG.SUMMARY_SHEET_ID_PROP, CONFIG_UNIFIED.SUMMARY_SHEET_ID);
    return CONFIG_UNIFIED.SUMMARY_SHEET_ID;
  }

  const ss = SpreadsheetApp.create(RUNTIME_CONFIG.SUMMARY_SHEET_FILE_NAME);
  const newId = ss.getId();
  props.setProperty(RUNTIME_CONFIG.SUMMARY_SHEET_ID_PROP, newId);

  // 可能なら指定フォルダへ移動（失敗しても処理は継続）
  try {
    const file = DriveApp.getFileById(newId);
    const folder = DriveApp.getFolderById(CONFIG_UNIFIED.ROUTE_FOLDER_ID);
    folder.addFile(file);
    DriveApp.getRootFolder().removeFile(file);
  } catch (e) {
    console.warn('Summary spreadsheet move skipped:', e.message);
  }

  return newId;
}

/**
 * 月間集約スプレッドシートを取得する。
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet}
 */
function getOrCreateSummarySpreadsheet() {
  const id = getSummarySpreadsheetId();
  return SpreadsheetApp.openById(id);
}

/**
 * 月シートを取得。存在しない場合はヘッダー付きで作成する。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string} sheetName
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getOrCreateMonthSheet(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) return sheet;

  sheet = ss.insertSheet(sheetName);
  sheet.getRange('A1:R1').setValues([[
    '日付', 'スタッフ名',
    '訪問先1', '開始1', '終了1',
    '訪問先2', '開始2', '終了2',
    '訪問先3', '開始3', '終了3',
    '作業内容', '作業開始', '作業終了',
    '天候', '事務フラグ', '確定者', '確定日時'
  ]]);
  return sheet;
}

/**
 * 会計年度を計算（月から）
 * @param {number} month - 1-12
 * @param {number} year - 西暦
 * @returns {number} 会計年度（e.g., 2025 = FY2025年度）
 */
function calcFiscalYear(month, year) {
  const startMonth = CONFIG_UNIFIED.YEAR_START_MONTH;
  return (month < startMonth) ? year - 1 : year;
}

/**
 * 月シート名を生成（YYYYMM形式）
 * @param {Date} date
 * @returns {string} '202504' など
 */
function generateMonthSheetName(date) {
  const y = date.getFullYear();
  const m = ('0' + (date.getMonth() + 1)).slice(-2);
  return `${y}${m}`;
}

/**
 * 列文字を列番号に変換（A->1, B->2, ..., Z->26, AA->27など）
 * @param {string} colChar - 列文字 (e.g., 'A', 'AA')
 * @returns {number} 列番号
 */
function columnToNumber(colChar) {
  let result = 0;
  for (let i = 0; i < colChar.length; i++) {
    result = result * 26 + (colChar.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
  }
  return result;
}

/**
 * 列番号を列文字に変換（1->A, 26->Z, 27->AA など）
 * @param {number} colNum
 * @returns {string}
 */
function numberToColumn(colNum) {
  let colChar = '';
  while (colNum > 0) {
    colNum--;
    colChar = String.fromCharCode(65 + (colNum % 26)) + colChar;
    colNum = Math.floor(colNum / 26);
  }
  return colChar;
}

/**
 * 日付を比較（タイムゾーン考慮）
 * @param {Date|string} d1
 * @param {Date|string} d2
 * @returns {boolean}
 */
function isSameDate(d1, d2) {
  const tz = 'Asia/Tokyo';
  try {
    const s1 = Utilities.formatDate(new Date(d1), tz, 'yyyyMMdd');
    const s2 = Utilities.formatDate(new Date(d2), tz, 'yyyyMMdd');
    return s1 === s2;
  } catch (e) {
    return false;
  }
}

/**
 * 値を比較（日付・時刻対応）
 * @param {*} oldVal
 * @param {*} newVal
 * @returns {boolean}
 */
function isSameValue(oldVal, newVal) {
  const sOld = String(oldVal);
  const sNew = String(newVal);
  if (sOld === sNew) return true;
  
  if (oldVal instanceof Date) {
    const formattedOld = Utilities.formatDate(oldVal, 'Asia/Tokyo', 'HH:mm');
    if (formattedOld === sNew) return true;
  }
  return false;
}
