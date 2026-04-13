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
  CUSTOMER_FOLDER_ID: '1wLjR6iZ447tbUa3ff59bejoM5aXh8clC', // 顧客CSV保存フォルダ
  STAFF_SHEET_NAME: 'Staff',   // スタッフマスタ シート名
  STAFF_SHEET_ID: '1exqD69qZqACm9KOUPpa0fVWRYD2qEZfce7I6TOs_VDk', // 既設スタッフマスター
  SUMMARY_SHEET_ID: '(TBD)', // ⚠️ NEW: 月間集約シート（作成後にID変更）
  ROUTE_FOLDER_ID: '1K2x_z3Tr0qLgpSRVpLn1f8SaBdW25G7q', // 月間集約ファイルの保存先フォルダ
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
// S-AG列: 勤怠計算用の補助列（予約URL・移動/出退勤経路情報）

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

  // 勤怠計算用の補助列
  VISIT_1_RESERVA_URL: 'S',
  VISIT_1_ATTENDANCE_URL: 'T',
  VISIT_1_ATTENDANCE_MIN: 'U',
  VISIT_1_ATTENDANCE_KM: 'V',
  VISIT_2_RESERVA_URL: 'W',
  VISIT_2_MOVE_URL: 'X',
  VISIT_2_MOVE_MIN: 'Y',
  VISIT_2_MOVE_KM: 'Z',
  VISIT_3_RESERVA_URL: 'AA',
  VISIT_3_MOVE_URL: 'AB',
  VISIT_3_MOVE_MIN: 'AC',
  VISIT_3_MOVE_KM: 'AD',
  LEAVING_URL: 'AE',
  LEAVING_MIN: 'AF',
  LEAVING_KM: 'AG',
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
    ensureFileInTargetFolder_(runtimeId);
    return runtimeId;
  }

  if (CONFIG_UNIFIED.SUMMARY_SHEET_ID && CONFIG_UNIFIED.SUMMARY_SHEET_ID !== '(TBD)') {
    props.setProperty(RUNTIME_CONFIG.SUMMARY_SHEET_ID_PROP, CONFIG_UNIFIED.SUMMARY_SHEET_ID);
    ensureFileInTargetFolder_(CONFIG_UNIFIED.SUMMARY_SHEET_ID);
    return CONFIG_UNIFIED.SUMMARY_SHEET_ID;
  }

  const ss = SpreadsheetApp.create(RUNTIME_CONFIG.SUMMARY_SHEET_FILE_NAME);
  const newId = ss.getId();
  props.setProperty(RUNTIME_CONFIG.SUMMARY_SHEET_ID_PROP, newId);

  ensureFileInTargetFolder_(newId);

  return newId;
}

/**
 * 対象ファイルを設定フォルダへ配置する。
 * すでに配置済みなら何もしない。
 * @param {string} fileId
 */
function ensureFileInTargetFolder_(fileId) {
  try {
    const file = DriveApp.getFileById(fileId);
    const targetFolder = DriveApp.getFolderById(CONFIG_UNIFIED.ROUTE_FOLDER_ID);

    // 既存親フォルダにターゲットが含まれているか判定
    let alreadyInTarget = false;
    const parents = file.getParents();
    while (parents.hasNext()) {
      if (parents.next().getId() === targetFolder.getId()) {
        alreadyInTarget = true;
        break;
      }
    }

    if (!alreadyInTarget) {
      targetFolder.addFile(file);
      try {
        DriveApp.getRootFolder().removeFile(file);
      } catch (eRoot) {
        console.warn('Root folder detach skipped:', eRoot.message);
      }
    }
  } catch (e) {
    console.warn('ensureFileInTargetFolder_ skipped:', e.message);
  }
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
 * 現在利用中の月間集約スプレッドシートURLをログ表示する。
 */
function showSummarySpreadsheetInfo() {
  const id = getSummarySpreadsheetId();
  const url = 'https://docs.google.com/spreadsheets/d/' + id;
  console.log('Summary Spreadsheet ID: ' + id);
  console.log('Summary Spreadsheet URL: ' + url);
}

/**
 * 月シートを取得。存在しない場合はヘッダー付きで作成する。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string} sheetName
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getOrCreateMonthSheet(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  ensureMonthSheetStructure_(sheet);
  return sheet;
}

/**
 * 月シートのヘッダーと列数を現行定義へ揃える。
 * 既存データは保持し、足りない列だけ右側へ追加する。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function ensureMonthSheetStructure_(sheet) {
  const headers = getMonthSheetHeaderValues_();
  const requiredColumns = headers.length;
  const currentMaxColumns = sheet.getMaxColumns();

  if (currentMaxColumns < requiredColumns) {
    sheet.insertColumnsAfter(currentMaxColumns, requiredColumns - currentMaxColumns);
  }

  sheet.getRange(1, 1, 1, requiredColumns).setValues([headers]);
}

/**
 * 月間集約シートのヘッダー定義を返す。
 * @returns {string[]}
 */
function getMonthSheetHeaderValues_() {
  return [
    '日付', 'スタッフ名',
    '訪問先1', '開始1', '終了1',
    '訪問先2', '開始2', '終了2',
    '訪問先3', '開始3', '終了3',
    '作業内容', '作業開始', '作業終了',
    '天候', '事務フラグ', '確定者', '確定日時',
    '予約詳細URL1', '出勤経路URL1', '出勤時間1', '出勤距離1',
    '予約詳細URL2', '移動経路URL2', '移動時間2', '移動距離2',
    '予約詳細URL3', '移動経路URL3', '移動時間3', '移動距離3',
    '退勤経路URL', '退勤時間', '退勤距離'
  ];
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

/**
 * スタッフマップを取得（同期処理用）
 * Staff シートから、メール → スタッフ情報のマップを作成
 * 退職済み(H列に日付あり)は除外
 * 
 * @returns {Object} { normalizedName -> { name, email, id, lwId, isAdmin } }
 */
function getAuthorizedStaffMapForSync() {
  try {
    const staffSheetId = CONFIG_UNIFIED.STAFF_SHEET_ID;
    if (!staffSheetId || staffSheetId === '(TBD)') {
      throw new Error('STAFF_SHEET_ID が未設定です');
    }

    const ss = SpreadsheetApp.openById(staffSheetId);
    const sheet = ss.getSheetByName(CONFIG_UNIFIED.STAFF_SHEET_NAME);
    if (!sheet) {
      throw new Error(`シート「${CONFIG_UNIFIED.STAFF_SHEET_NAME}」が見つかりません`);
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return {};

    // Staff マスター（A〜L）
    // A:ID, B:氏名, C:住所, D:TEL, E:mail address, F:緯度経度
    // G:入社日, H:退職日, I:ユーザーID, J:PW, K:管理者, L:LW_ID
    const data = sheet.getRange(2, 1, lastRow - 1, 12).getValues();
    const map = {};

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const id = row[0];          // A列
      const name = row[1];        // B列
      const address = row[2] || ''; // C列
      const email = row[4];       // E列
      const latLng = row[5] || ''; // F列
      const retiredDate = row[7]; // H列
      const isAdmin = row[10];    // K列
      const lwId = row[11];       // L列

      let lat = '';
      let lng = '';
      if (typeof latLng === 'string' && latLng.includes(',')) {
        const parts = latLng.split(',');
        lat = parseFloat(String(parts[0]).trim()) || '';
        lng = parseFloat(String(parts[1]).trim()) || '';
      }

      // アクティブなスタッフのみ（退職日が未設定）
      if (name && name !== '' && email && email !== '' && (!retiredDate || retiredDate === '')) {
        const normalizedName = String(name).trim();
        map[normalizedName] = {
          id: id || '',
          name: normalizedName,
          address: address,
          lat: lat,
          lng: lng,
          email: String(email).trim(),
          lwId: lwId || '',
          isAdmin: isAdmin || ''
        };
      }
    }

    return map;
  } catch (e) {
    console.error('[getAuthorizedStaffMapForSync] error:', e.message);
    throw e;
  }
}
