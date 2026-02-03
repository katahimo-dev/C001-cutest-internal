// ==========================================
// 【最終完成版】スマホde出勤簿修正アプリ
// ==========================================

var YEAR = '2025'; // 年度
var STAFF_SHEET_NAME = 'Staff'; // タブ名

// ==========================================
// ★設定エリア：列の定義
// ==========================================
var INPUT_COLUMNS = {
  // ■1件目
  'C': {type: 'text', label: '#1訪問先等', readonly: true}, // 保護列（書き込みなし）
  'D': {type: 'time', label: '#1始業時刻'},
  'E': {type: 'time', label: '#1終業時刻'},
  'I': {type: 'text', label: '天候'},
  
  // ■2件目
  'L': {type: 'text', label: '#2訪問先等', readonly: true}, // 保護列（書き込みなし）
  'M': {type: 'time', label: '#2始業時刻'},
  'N': {type: 'time', label: '#2終業時刻'},
  'R': {type: 'text', label: '天候'},
  
  // ■3件目
  'U': {type: 'text', label: '#3訪問先等', readonly: true}, // 保護列（書き込みなし）
  'V': {type: 'time', label: '#3始業時刻'},
  'W': {type: 'time', label: '#3終業時刻'}, 
  
  // ■作業記録
  'X': {type: 'text', label: '作業１'},
  'Y': {type: 'time', label: '作業１開始'},
  'Z': {type: 'time', label: '作業１終了'},
  
  'AA': {type: 'text', label: '作業２'},
  'AB': {type: 'time', label: '作業２開始'},
  'AC': {type: 'time', label: '作業２終了'},
  
  // ■その他
  // ★変更：ここを 'text' から 'select' に変えました
  'AN': {type: 'select', label: '買物代行'}, 
  'AO': {type: 'text', label: '備考'},
};

// ==========================================
// 1. 基本設定・メニュー
// ==========================================

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('スマホde出勤簿修正')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// function onOpen() {
//   var ui = SpreadsheetApp.getUi();
//   ui.createMenu('📱アプリURL送付')
//     .addItem('☑️ チェックした人にURLを送信', 'sendAppUrlToSelected')
//     .addToUi();
// }

// ==========================================
// 2. スタッフ特定・権限確認
// ==========================================

function getAuthorizedStaffMap() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(STAFF_SHEET_NAME);
  
  if (!sheet) {
    throw new Error('このファイル内に「' + STAFF_SHEET_NAME + '」シートが見つかりません。');
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return {}; 

  var data = sheet.getRange(2, 2, lastRow - 1, 5).getValues(); 
  var map = {};

  for (var i = 0; i < data.length; i++) {
    var name = data[i][1];       // C列
    var email = data[i][2];      // D列
    var retiredDate = data[i][4]; // F列

    if (email && email !== "" && (!retiredDate || retiredDate === "")) {
      map[String(email).trim()] = name;
    }
  }
  return map;
}

// ==========================================
// 3. データ取得処理
// ==========================================

function getInitialData() {
  try {
    var email = Session.getActiveUser().getEmail();
    var staffMap = getAuthorizedStaffMap();
    var staffName = staffMap[email];
    
    if (!staffName) {
      throw new Error('あなたのメールアドレス (' + email + ') はスタッフ登録されていません。');
    }

    var spreadsheet = findSpreadsheetGlobal(staffName, YEAR);
    var sheet = spreadsheet.getSheets()[0]; 

    // シートから選択肢を読み取る関数
    var getOptions = function(colChar) {
      var range = sheet.getRange(colChar + "5"); 
      var rule = range.getDataValidation();
      if (rule != null) {
        var criteria = rule.getCriteriaType();
        var args = rule.getCriteriaValues();
        if (criteria === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
          return args[0]; 
        } else if (criteria === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
          var sourceRange = args[0];
          var values = sourceRange.getValues();
          return values.flat().filter(function(v){ return v !== ""; });
        }
      }
      return []; 
    };

    return {
      success: true,
      userName: staffName,
      optionsI: getOptions('I'),
      optionsR: getOptions('R'),
      optionsX: getOptions('X'),
      optionsAA: getOptions('AA'),
      // ★追加：買物代行の選択肢（ここで数字を自由に設定してください）
      // 必要に応じて ['100', '200', '300'] や ['あり', 'なし'] などに変えてください
      optionsAN: ['1', '2', '3', '4', '5']
    };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function getDataByDate(dateString) {
  try {
    var email = Session.getActiveUser().getEmail();
    var staffMap = getAuthorizedStaffMap();
    var staffName = staffMap[email];

    if (!staffName) throw new Error('スタッフ登録が確認できません。');

    var spreadsheet = findSpreadsheetGlobal(staffName, YEAR);
    
    var sheets = spreadsheet.getSheets();
    var targetSheetName = null;
    var targetRow = -1;
    
    var searchDate = new Date(dateString.replace(/-/g, '/')); 
    
    for (var s = 0; s < sheets.length; s++) {
      var sheet = sheets[s];
      var lastRow = sheet.getLastRow();
      if (lastRow < 1) continue; 

      var dateValues = sheet.getRange(1, 1, lastRow, 1).getValues();
      for (var i = 0; i < dateValues.length; i++) {
        var val = dateValues[i][0];
        if (!val) continue;
        if (isSameDate(val, searchDate)) {
          targetSheetName = sheet.getName();
          targetRow = i + 1;
          break; 
        }
      }
      if (targetRow !== -1) break; 
    }

    if (targetRow === -1) {
      throw new Error('指定された日付 (' + dateString + ') の行が見つかりませんでした。');
    }

    PropertiesService.getUserProperties().setProperty('LAST_SHEET_NAME', targetSheetName);

    var targetSheet = spreadsheet.getSheetByName(targetSheetName);
    var rowData = {};
    for (var colChar in INPUT_COLUMNS) {
      var colNum = columnToNumber(colChar);
      var val = targetSheet.getRange(targetRow, colNum).getValue();
      if (val instanceof Date) {
        val = Utilities.formatDate(val, 'Asia/Tokyo', 'HH:mm');
      }
      rowData[colChar] = val;
    }

    return { success: true, rowData: rowData, rowNumber: targetRow };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ==========================================
// 4. データ保存処理
// ==========================================

function updateData(formObject) {
  try {
    var targetDateStr = formObject.targetDateStr; 
    if (!targetDateStr) throw new Error('日付情報が不足しています。');

    var targetDate = new Date(targetDateStr.replace(/-/g, '/'));
    var today = new Date();
    targetDate.setHours(0,0,0,0); today.setHours(0,0,0,0);
    var limitDate = new Date(today); limitDate.setDate(today.getDate() - 7);

    if (targetDate < limitDate) {
      throw new Error('修正期限切れです。\n7日前（' + Utilities.formatDate(limitDate, 'Asia/Tokyo', 'MM/dd') + '）より過去の記録は変更できません。');
    }

    var rowNumber = parseInt(formObject.rowNumber);
    if (!rowNumber) throw new Error('行番号が不明です。');

    var email = Session.getActiveUser().getEmail();
    var staffMap = getAuthorizedStaffMap();
    var staffName = staffMap[email];

    if (!staffName) throw new Error('スタッフ登録が確認できません。');

    var spreadsheet = findSpreadsheetGlobal(staffName, YEAR);
    
    var targetSheetName = PropertiesService.getUserProperties().getProperty('LAST_SHEET_NAME');
    var sheet = targetSheetName ? spreadsheet.getSheetByName(targetSheetName) : spreadsheet.getSheets()[0];

    // 書き込み制御
    for (var colChar in INPUT_COLUMNS) {
      var config = INPUT_COLUMNS[colChar];
      
      // readonlyならスキップ
      if (config.readonly === true) {
        continue;
      }

      var inputKey = 'col_' + colChar;
      var newValue = formObject[inputKey]; 
      var colNum = columnToNumber(colChar);

      if (newValue !== undefined) { 
        var cell = sheet.getRange(rowNumber, colNum);
        var oldValue = cell.getValue(); 
        
        if (!isSameValue(oldValue, newValue)) {
          cell.setValue(newValue);      
          cell.setBackground("#fce4e4"); 
        }
      }
    }
    return { success: true, message: '修正完了しました。お疲れ様でした！' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ==========================================
// 5. ファイル検索・メール・権限関数
// ==========================================

function findSpreadsheetGlobal(staffName, year) {
  var searchKey = staffName + '_出勤簿_' + year + '年度';
  
  var query = "title = '" + searchKey + "' and mimeType = 'application/vnd.google-apps.spreadsheet' and trashed = false";
  
  try {
    var files = DriveApp.searchFiles(query);
    if (files.hasNext()) {
      return SpreadsheetApp.open(files.next());
    }
  } catch (e) {
    console.log("検索クエリエラー(title一致): " + e.message);
  }
  
  var searchKeyNoSpace = searchKey.replace(/\s+/g, '').replace(/　/g, '');
  var allFiles = DriveApp.searchFiles("title contains '" + staffName + "' and title contains '出勤簿' and trashed = false");
  
  while (allFiles.hasNext()) {
    var file = allFiles.next();
    var fileName = file.getName();
    var normalized = fileName.replace(/\s+/g, '').replace(/　/g, '');
    if (normalized === searchKeyNoSpace) {
      return SpreadsheetApp.open(file);
    }
  }

  throw new Error('あなたの出勤簿ファイルが見つかりません。\n検索したファイル名: ' + searchKey);
}

function sendAppUrlToSelected() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Staff'); 
  
  var appUrl = sheet.getRange('H8').getValue();
  
  if (!appUrl || !String(appUrl).match(/^https:\/\/script\.google\.com/)) {
    Browser.msgBox('エラー', 'H8セルに正しいアプリのURL（https://...）が入力されていません。', Browser.Buttons.OK);
    return;
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  
  var data = sheet.getRange(2, 3, lastRow - 1, 5).getValues();
  var sentCount = 0;
  var errorLog = ""; 

  for (var i = 0; i < data.length; i++) {
    var rowNum = i + 2;
    var name = data[i][0];       // C列
    var email = data[i][1];      // D列
    var retiredDate = data[i][3]; // F列
    var isChecked = data[i][4];  // G列

    if (isChecked === true && email && (!retiredDate || retiredDate === "")) {
      try {
        var htmlBody = 
          name + ' さん<br><br>' +
          'お疲れ様です。<br>' +
          'スマホから出勤簿を修正できるアプリのURLをお送りします。<br><br>' +
          '以下のリンクをタップしてアクセスしてください：<br><br>' +
          '<b><a href="' + appUrl + '" style="font-size: 16px;">👉 アプリURL</a></b><br><br>' +
          'このアプリは、' + name + ' さんだけが、出勤簿を修正できます。';

        MailApp.sendEmail({
          to: email,
          subject: '【業務連絡】勤怠修正アプリのURLをお知らせします',
          htmlBody: htmlBody
        });
        
        sheet.getRange(rowNum, 7).setValue(false); 
        sentCount++;
        
      } catch (e) {
        errorLog += "・" + rowNum + "行目 (" + name + "): " + e.message + "\n";
      }
    }
  }
  var ui = SpreadsheetApp.getUi();
  if (sentCount > 0) {
    ui.alert('完了', sentCount + ' 名にメールを送信しました。', ui.ButtonSet.OK);
  } else {
    if (errorLog !== "") {
      ui.alert('送信エラー', '送信時にエラーが発生しました。\n\n' + errorLog, ui.ButtonSet.OK);
    } else {
      ui.alert('確認', '送信対象がいませんでした。', ui.ButtonSet.OK);
    }
  }
}

function forceAuth() {
  MailApp.getRemainingDailyQuota();
  DriveApp.getRootFolder();
  SpreadsheetApp.create("Dummy");
  console.log("すべての権限認証に成功しました！");
}

function isSameValue(oldVal, newVal) {
  var sOld = String(oldVal);
  var sNew = String(newVal);
  if (sOld === sNew) return true;
  if (oldVal instanceof Date) {
    var formattedOld = Utilities.formatDate(oldVal, 'Asia/Tokyo', 'HH:mm');
    if (formattedOld === sNew) return true;
  }
  return false;
}

function isSameDate(d1, d2) {
  var tz = 'Asia/Tokyo'; 
  try {
    var s1 = Utilities.formatDate(new Date(d1), tz, 'yyyyMMdd');
    var s2 = Utilities.formatDate(new Date(d2), tz, 'yyyyMMdd');
    return s1 === s2;
  } catch (e) { return false; }
}

function columnToNumber(column) {
  var result = 0;
  for (var i = 0; i < column.length; i++) {
    result = result * 26 + (column.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
  }
  return result;
}