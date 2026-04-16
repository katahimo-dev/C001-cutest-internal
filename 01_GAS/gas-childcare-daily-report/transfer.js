/**
 * 勤怠管理自動転記システム（年度別リンク・HYPERLINK関数対応版）
 */

const CONFIG_TRANS = {
  FOLDER_ID_INFO: '1FA2aSBddgBakETEbzJhJIx1vWG06P46J', // ルート情報フォルダ
  ROUTE_FILE_NAME: 'ルート集計',
  ATTENDANCE_FILE_NAME: '勤怠集計',
  
  STAFF_START_ROW: 9,
  COL_NAME: 3,
  BASE_YEAR: 2025,
  BASE_COL_LINK: 6,
  DATA_START_ROW: 4
};

// メニュー作成
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🟩スタッフ情報更新')
    .addItem(' 最新情報を取得する', 'syncStaffInfo')
    .addToUi();

  ui.createMenu('🟦 出勤簿入力')
    .addItem('出勤簿ファイルへ転記（日付指定）', 'runAttendanceTransfer')
    .addItem('出勤簿ファイルへ転記（期間指定）', 'runAttendanceTransferRangeInteractive')
    .addItem('【手動】今日の分を今すぐ転記', 'autoRunDailyTransfer')
    .addToUi();
  
  ui.createMenu('🟧 出勤簿作成（年度）')
    .addItem('出勤簿一括作成', 'promptAndCreateFiles')
    .addToUi();
  
  // ui.createMenu('🔧 メンテナンス')
    
  //   .addItem('【注意】AM列の数式を一括更新', 'updateFormulaAllFiles')
  //   .addToUi();

}

/**
 * 【手動用】日付を指定して実行する関数
 */
function runAttendanceTransfer() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('データ転記', '転記する日付を「YYYY-MM-DD」形式で入力してください\n(例: 2026-01-08)', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  const targetDateStr = response.getResponseText().trim();

  try {
    const report = processTransfer(targetDateStr);
    ui.alert('処理結果', report, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('エラーが発生しました', e.message, ui.ButtonSet.OK);
  }
}

/**
 * 【トリガー専用】実行日の「当日」日付で自動実行する関数
 * ※夜に実行して、当日分を転記する運用
 */
function autoRunDailyTransfer() {
  const today = new Date();
  const targetDate = new Date(today);
  
  const year = targetDate.getFullYear();
  const month = ("0" + (targetDate.getMonth() + 1)).slice(-2);
  const day = ("0" + targetDate.getDate()).slice(-2);
  const targetDateStr = `${year}-${month}-${day}`;
  
  console.log("自動実行(当日分)を開始します: " + targetDateStr);
  
  try {
    const report = processTransfer(targetDateStr);
    console.log("処理結果:\n" + report); 
  } catch (e) {
    console.error("エラーが発生しました: " + e.message);
  }
}

function calcDurationMin(startTime, endTime) {
  if (!startTime || !endTime) return null;

  const toMin = (value) => {
    const normalized = String(value).substring(0, 5);
    const parts = normalized.split(':');
    if (parts.length < 2) return null;
    const h = Number(parts[0]);
    const m = Number(parts[1]);
    if (isNaN(h) || isNaN(m)) return null;
    return h * 60 + m;
  };

  const s = toMin(startTime);
  const e = toMin(endTime);
  if (s === null || e === null) return null;

  if (e >= s) return e - s;
  return (24 * 60 - s) + e;
}

function isOfficeWorkEntry(entry) {
  if (entry.eventType === 'OFFICE WORK') return true;

  if (entry.eventType === 'CUSTOMER APPOINTMENT') {
    const duration = calcDurationMin(entry.startTime, entry.endTime);
    return duration === 15;
  }

  return false;
}



/**
 * 転記処理のコアロジック（修正版）
 * ★修正点: Q列への書き込みを停止
 */
function processTransfer(targetDateStr) {
  // --- 0. 日付情報の解析と対象シートの特定 ---
  const normalizedDateStr = targetDateStr.replace(/\//g, '-');
  const dateObj = new Date(normalizedDateStr);
  
  if (isNaN(dateObj.getTime())) {
    throw new Error("日付形式が正しくありません: " + targetDateStr);
  }

  const year = dateObj.getFullYear();
  const month = dateObj.getMonth() + 1;
  const day = dateObj.getDate();
  
  const monthStr = month < 10 ? '0' + month : '' + month;
  const routeSheetName = year + monthStr;

  // --- 1. 勤怠集計データの読み込み ---
  const infoFolder = DriveApp.getFolderById(CONFIG_TRANS.FOLDER_ID_INFO);
  const routeFile = findFileByName(infoFolder, CONFIG_TRANS.ATTENDANCE_FILE_NAME);
  if (!routeFile) throw new Error('「勤怠集計」ファイルが見つかりません。');
  
  const routeSS = SpreadsheetApp.open(routeFile);
  const routeSheet = routeSS.getSheetByName(routeSheetName);
  if (!routeSheet) {
    throw new Error('勤怠集計ファイルに、対象月のシート「' + routeSheetName + '」が見つかりません。');
  }

  const routeDisplayValues = routeSheet.getDataRange().getDisplayValues();
  const dailyData = {};
  let foundRowCount = 0;

  const targetHyphen = normalizedDateStr;
  const targetSlash = normalizedDateStr.replace(/-/g, '/');

  for (let i = 1; i < routeDisplayValues.length; i++) {
    let cellStr = routeDisplayValues[i][0];
    if (!cellStr) continue;

    if (cellStr.indexOf(targetHyphen) !== -1 || cellStr.indexOf(targetSlash) !== -1) {
      let staffName = routeDisplayValues[i][1] ? routeDisplayValues[i][1].trim() : "";
      if (!staffName && cellStr.length > 10) {
        staffName = cellStr.replace(targetHyphen, "").replace(targetSlash, "").trim();
      }

      if (!staffName) continue;
      if (!dailyData[staffName]) dailyData[staffName] = [];
      
      dailyData[staffName].push({
        eventType: routeDisplayValues[i][2],
        customer: routeDisplayValues[i][3],
        startTime: routeDisplayValues[i][4],
        endTime: routeDisplayValues[i][5],
        travelTime: routeDisplayValues[i][8],
        travelDist: routeDisplayValues[i][9],
        commuteDist: routeDisplayValues[i][12],
        returnDist: routeDisplayValues[i][15]
      });
      foundRowCount++;
    }
  }

  if (foundRowCount === 0) return targetDateStr + " のデータは勤怠集計シート「" + routeSheetName + "」に見つかりませんでした。";

  // --- 2. 会計年度の計算 ---
  const fiscalYear = (month <= 3) ? year - 1 : year;
  
  // --- 3. スタッフ一覧表からリンクを取得して転記 ---
  const managerSS = SpreadsheetApp.getActiveSpreadsheet();
  const listSheet = managerSS.getSheets()[0];
  
  if (fiscalYear < CONFIG_TRANS.BASE_YEAR) throw new Error(CONFIG_TRANS.BASE_YEAR + '年度以前のデータには対応していません。');
  const targetLinkCol = CONFIG_TRANS.BASE_COL_LINK + (fiscalYear - CONFIG_TRANS.BASE_YEAR);

  const lastRow = listSheet.getLastRow();
  const listRange = listSheet.getRange(CONFIG_TRANS.STAFF_START_ROW, 1, lastRow - CONFIG_TRANS.STAFF_START_ROW + 1, targetLinkCol);
  
  const staffListValues = listRange.getValues();
  const staffListFormulas = listRange.getFormulas();

  let successStaffs = [];
  let errorStaffs = [];

  for (const staff in dailyData) {
    const rowIndex = staffListValues.findIndex(r => r[CONFIG_TRANS.COL_NAME - 1] === staff);
    
    if (rowIndex === -1) {
      errorStaffs.push(staff + "(名簿なし)");
      continue;
    }

    let fileUrl = staffListValues[rowIndex][targetLinkCol - 1];
    const cellFormula = staffListFormulas[rowIndex][targetLinkCol - 1];

    if (cellFormula && cellFormula.includes("http")) {
      const match = cellFormula.match(/"(https?:\/\/[^"]+)"/);
      if (match) fileUrl = match[1]; 
    }

    if (!fileUrl || fileUrl === "-" || String(fileUrl).indexOf("http") === -1) {
      errorStaffs.push(staff + "(" + fiscalYear + "年度ファイルなし)");
      continue;
    }

    try {
      const ss = SpreadsheetApp.openByUrl(fileUrl);
      const sheetName = month + "月";
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        errorStaffs.push(staff + "(シート" + sheetName + "なし)");
        continue;
      }

      const targetRow = CONFIG_TRANS.DATA_START_ROW + day - 1;
      const allEntries = dailyData[staff];
      const officeWorks = allEntries.filter(isOfficeWorkEntry);
      const visits = allEntries.filter(v => !isOfficeWorkEntry(v));

      // 再転記時の残存データ防止のため、転記対象列を先にクリア
      sheet.getRange(targetRow, 3, 1, 3).clearContent();   // C:E
      sheet.getRange(targetRow, 8).clearContent();          // H
      sheet.getRange(targetRow, 12, 1, 3).clearContent();   // L:N
      sheet.getRange(targetRow, 21, 1, 3).clearContent();   // U:W
      sheet.getRange(targetRow, 24, 1, 6).clearContent();   // X:AC
      sheet.getRange(targetRow, 33, 1, 4).clearContent();   // AG:AJ

      // --- 書き込み処理（修正箇所） ---
      visits.forEach((v, idx) => {
        if (idx === 0) {
          // 1件目
          sheet.getRange(targetRow, 3).setValue(v.customer);   // C
          sheet.getRange(targetRow, 4).setValue(v.startTime);  // D
          sheet.getRange(targetRow, 5).setValue(v.endTime);    // E
          sheet.getRange(targetRow, 35).setValue(v.commuteDist); // AI
        } else if (idx === 1) {
          // 2件目
          sheet.getRange(targetRow, 12).setValue(v.customer);  // L
          sheet.getRange(targetRow, 13).setValue(v.startTime); // M
          sheet.getRange(targetRow, 14).setValue(v.endTime);   // N
          // J列, O列は指定されていないため書き込みなし
          if (v.travelTime) sheet.getRange(targetRow, 8).setValue(v.travelTime); // H (ここはそのまま)
          if (v.travelDist) sheet.getRange(targetRow, 33).setValue(v.travelDist); // AG
        } else if (idx === 2) {
          // 3件目
          sheet.getRange(targetRow, 21).setValue(v.customer);  // U
          sheet.getRange(targetRow, 22).setValue(v.startTime); // V
          sheet.getRange(targetRow, 23).setValue(v.endTime);   // W
          
          
          if (v.travelTime) sheet.getRange(targetRow, 17).setValue(v.travelTime); 
          
          if (v.travelDist) sheet.getRange(targetRow, 34).setValue(v.travelDist); // AH
          // S列は指定されていないため書き込みなし
        }
      });

      if (visits.length > 0) {
        const lastVisit = visits[visits.length - 1];
        sheet.getRange(targetRow, 36).setValue(lastVisit.returnDist); // AJ
      }

      if (officeWorks[0]) {
        sheet.getRange(targetRow, 24).setValue(officeWorks[0].customer);  // X
        sheet.getRange(targetRow, 25).setValue(officeWorks[0].startTime); // Y
        sheet.getRange(targetRow, 26).setValue(officeWorks[0].endTime);   // Z
      }

      if (officeWorks[1]) {
        sheet.getRange(targetRow, 27).setValue(officeWorks[1].customer);  // AA
        sheet.getRange(targetRow, 28).setValue(officeWorks[1].startTime); // AB
        sheet.getRange(targetRow, 29).setValue(officeWorks[1].endTime);   // AC
      }

      // ★追加点: 転記したシートをアクティブ（開いた状態）にする
      sheet.activate();

      successStaffs.push(staff);

    } catch (e) {
      errorStaffs.push(staff + "(エラー: " + e.message + ")");
    }
  }

  let msg = "【完了】対象日: " + targetDateStr + "\n";
  msg += "対象シート: " + routeSheetName + "\n";
  msg += "成功: " + successStaffs.length + "件 (" + successStaffs.join(', ') + ")\n";
  if (errorStaffs.length > 0) {
    msg += "未転記: " + errorStaffs.join(', ');
  }
  return msg;
}

// ---------------------------------------------
// 期間指定一括転記: ヘルパーとUIトリガー
// ---------------------------------------------
function formatYMD(d) {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}

function runTransferForRange(startDateStr, endDateStr) {
  const s = new Date(String(startDateStr).replace(/\//g, '-'));
  const e = new Date(String(endDateStr).replace(/\//g, '-'));
  if (isNaN(s.getTime()) || isNaN(e.getTime()) || s > e) {
    throw new Error('開始／終了日が不正です');
  }
  const results = [];
  const cur = new Date(s.getTime());
  while (cur <= e) {
    const dateStr = formatYMD(cur);
    try {
      processTransfer(dateStr);
      results.push(dateStr + ' → OK');
    } catch (err) {
      results.push(dateStr + ' → エラー: ' + err.message);
    }
    cur.setDate(cur.getDate() + 1);
  }
  return results.join('\n');
}

function runAttendanceTransferRangeInteractive() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('期間転記', '開始日と終了日を「YYYY-MM-DD,YYYY-MM-DD」の形式で入力してください\n例: 2026-01-01,2026-01-07', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const parts = res.getResponseText().split(',');
  if (parts.length !== 2) { ui.alert('入力形式が不正です'); return; }
  try {
    const report = runTransferForRange(parts[0].trim(), parts[1].trim());
    ui.alert('処理結果', report, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('エラー', e.message, ui.ButtonSet.OK);
  }
}

/**
 * 補助関数（ルート集計検索用）
 */
function findFileByName(folder, name) {
  const files = folder.getFiles();
  const searchName = name.replace(/\s+/g, "");
  while (files.hasNext()) {
    let file = files.next();
    let fileName = file.getName().replace(/\s+/g, "");
    if (fileName.indexOf(searchName) !== -1) return file;
  }
  return null;
}
