/**
 * スタッフ別出勤簿一括生成システム（タイムアウト回避・重複スキップ版）
 */

const CONFIG_GEN = {
  STAFF_START_ROW: 9,      
  COL_ID: 2,               
  COL_NAME: 3,             
  COL_RESIGN: 5,           
  BASE_YEAR: 2025,         
  BASE_COL_LINK: 6,        
  TEMPLATE_DATA_ROW: 4,    
  TEMPLATE_FOLDER_NAME: "出勤簿",      
  TEMPLATE_FILE_NAME: "出勤簿テンプレート",
  // ★追加: タイムアウト対策（ミリ秒）。5分（300,000ms）経過したら安全に停止
  TIME_LIMIT_MS: 300000 
};

function promptAndCreateFiles() {
  const ui = SpreadsheetApp.getUi();
  const startTime = new Date().getTime(); // 開始時刻を記録
  
  // --- ステップ1: 年度の入力 ---
  const responseYear = ui.prompt('ステップ 1/2: 年度の入力', '作成したい【会計年度】を入力してください（例: 2025）', ui.ButtonSet.OK_CANCEL);
  if (responseYear.getSelectedButton() !== ui.Button.OK) return;
  
  const fiscalYear = parseInt(responseYear.getResponseText().trim());
  if (isNaN(fiscalYear) || fiscalYear < CONFIG_GEN.BASE_YEAR) {
    ui.alert('エラー', '有効な年度を入力してください。', ui.ButtonSet.OK);
    return;
  }

  // --- ステップ2: 対象者の選択 ---
  const responseId = ui.prompt('ステップ 2/2: 対象者の選択', 
    '特定のスタッフのみ作成する場合: 【スタッフID】を入力してください。\n' +
    '全員分を一括作成する場合: 【空白】のままOKを押してください。', 
    ui.ButtonSet.OK_CANCEL);
  
  if (responseId.getSelectedButton() !== ui.Button.OK) return;
  const targetId = responseId.getResponseText().trim(); 

  // --- 実行 ---
  let msg = targetId === "" ? "【全員分】" : "スタッフID: " + targetId + " のみ";
  
  // 処理開始のトースト表示（ユーザーへのフィードバック）
  SpreadsheetApp.getActiveSpreadsheet().toast('処理を開始します...', 'ステータス', 5);

  try {
    const result = createStaffAttendanceFiles(fiscalYear, targetId, startTime);
    
    if (result.status === 'TIMEOUT') {
      ui.alert('タイムアウト回避', 
        '処理時間が長くなったため、一旦停止しました。\n' +
        '現在の作成数: ' + result.count + '件\n\n' +
        '★続きを作成するには、もう一度メニューから実行してください。\n（作成済みのファイルは自動的にスキップされます）', 
        ui.ButtonSet.OK);
    } else if (result.count > 0) {
      ui.alert('完了', fiscalYear + '年度の出勤簿作成が完了しました。\n対象: ' + msg + '\n今回作成数: ' + result.count + '件\nスキップ数: ' + result.skipped + '件', ui.ButtonSet.OK);
    } else {
      ui.alert('完了', '新規作成対象はいませんでした。\n（作成済み、または対象者なし）\nスキップ数: ' + result.skipped + '件', ui.ButtonSet.OK);
    }
  } catch (e) {
    ui.alert('エラーが発生しました', e.message, ui.ButtonSet.OK);
  }
}

function createStaffAttendanceFiles(fiscalYear, targetId, startTime) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const staffSheet = ss.getSheets()[0];
  
  const currentFolder = DriveApp.getFileById(ss.getId()).getParents().next();
  const templateFiles = currentFolder.getFilesByName(CONFIG_GEN.TEMPLATE_FILE_NAME);
  if (!templateFiles.hasNext()) {
    throw new Error("テンプレート「" + CONFIG_GEN.TEMPLATE_FILE_NAME + "」が見つかりません。");
  }
  const templateFile = templateFiles.next(); 
  
  // テンプレートの保護情報を取得
  const originalTemplateSS = SpreadsheetApp.open(templateFile);
  const originalTemplateSheet = originalTemplateSS.getSheets()[0]; 

  const targetFolderName = fiscalYear + "出勤簿";
  let targetFolder;
  const folders = currentFolder.getFoldersByName(targetFolderName);
  if (folders.hasNext()) {
    targetFolder = folders.next();
  } else {
    targetFolder = currentFolder.createFolder(targetFolderName);
  }

  const targetLinkCol = CONFIG_GEN.BASE_COL_LINK + (fiscalYear - CONFIG_GEN.BASE_YEAR);
  const headerRow = CONFIG_GEN.STAFF_START_ROW - 1; 
  if (headerRow > 0) {
    staffSheet.getRange(headerRow, targetLinkCol).setValue(fiscalYear + "年度");
  }

  const lastRow = staffSheet.getLastRow();
  if (lastRow < CONFIG_GEN.STAFF_START_ROW) return { count: 0, skipped: 0, status: 'COMPLETE' };
  
  const maxCol = Math.max(CONFIG_GEN.COL_RESIGN, targetLinkCol); 
  const staffValues = staffSheet.getRange(CONFIG_GEN.STAFF_START_ROW, 1, lastRow - CONFIG_GEN.STAFF_START_ROW + 1, maxCol).getValues();

  let processCount = 0;
  let skipCount = 0;
  let status = 'COMPLETE';

  // --- ループ処理 ---
  // forEachではなくforループにして中断できるようにする
  for (let i = 0; i < staffValues.length; i++) {
    // ★タイムアウトチェック
    if (new Date().getTime() - startTime > CONFIG_GEN.TIME_LIMIT_MS) {
      status = 'TIMEOUT';
      break; // ループを抜ける
    }

    const row = staffValues[i];
    const rowIndex = CONFIG_GEN.STAFF_START_ROW + i;
    const staffId = row[CONFIG_GEN.COL_ID - 1];   
    const staffName = row[CONFIG_GEN.COL_NAME - 1]; 
    const rawResignDate = row[CONFIG_GEN.COL_RESIGN - 1];

    if (!staffName) continue;
    if (targetId !== "" && String(staffId) !== String(targetId)) continue;

    // --- ①対策：重複チェック ---
    const newFileName = staffName + "_出勤簿_" + fiscalYear + "年度";
    if (targetFolder.getFilesByName(newFileName).hasNext()) {
      // 既にファイルがある場合はスキップし、リンクだけ念の為更新（または何もしない）
      const existingFile = targetFolder.getFilesByName(newFileName).next();
      const formula = '=HYPERLINK("' + existingFile.getUrl() + '", "' + staffId + '出勤簿")';
      staffSheet.getRange(rowIndex, targetLinkCol).setFormula(formula);
      
      skipCount++;
      continue; // 次の人の処理へ
    }

    // --- 退職判定 ---
    // (ここは元のロジック通りですが、少し整理しても良いでしょう)
    let isResigned = false;
    const fyStartDate = new Date(fiscalYear, 3, 1, 12, 0, 0);
    if (rawResignDate) {
      let rDate = null;
      if (rawResignDate instanceof Date) {
        rDate = new Date(rawResignDate);
      } else if (typeof rawResignDate === 'string' && rawResignDate.trim() !== "") {
        let dateStr = rawResignDate.replace(/\./g, '/').replace(/-/g, '/');
        let parsedTime = Date.parse(dateStr);
        if (!isNaN(parsedTime)) rDate = new Date(parsedTime);
      }
      if (rDate) {
        rDate.setHours(12, 0, 0, 0); 
        if (rDate < fyStartDate) isResigned = true;
      }
    }

    if (isResigned) {
      staffSheet.getRange(rowIndex, targetLinkCol).setValue("- (退職済)");
      continue; 
    }

    // --- ファイル作成 ---
    const newFile = templateFile.makeCopy(newFileName, targetFolder);
    const newSS = SpreadsheetApp.open(newFile);
    const baseSheet = newSS.getSheets()[0];
    
    const months = [4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3];
    months.forEach(month => {
      let currentYear = (month >= 4) ? fiscalYear : fiscalYear + 1;
      let newSheet = baseSheet.copyTo(newSS);
      newSheet.setName(month + "月");
      setupAttendanceSheet(newSheet, fiscalYear, currentYear, month, staffId, staffName, originalTemplateSheet);
    });

    newSS.deleteSheet(baseSheet);

    const linkText = staffId + "出勤簿"; 
    const formula = '=HYPERLINK("' + newSS.getUrl() + '", "' + linkText + '")';
    staffSheet.getRange(rowIndex, targetLinkCol).setFormula(formula);
    
    processCount++;
  }

  return { count: processCount, skipped: skipCount, status: status };
}

function setupAttendanceSheet(sheet, fiscalYear, currentYear, month, staffId, staffName, templateSheet) {
  sheet.getRange("A2").setValue(currentYear); 
  sheet.getRange("B2").setValue(month);       
  sheet.getRange("C2").setValue(staffId);
  sheet.getRange("D2").setValue(staffName);

  const lastDay = new Date(currentYear, month, 0).getDate();
  const weekDays = ["日", "月", "火", "水", "木", "金", "土"];
  let dateValues = [];
  
  for (let d = 1; d <= lastDay; d++) {
    let dateObj = new Date(currentYear, month - 1, d, 12, 0, 0);
    dateValues.push([dateObj, weekDays[dateObj.getDay()]]);
  }
  
  const targetRange = sheet.getRange(CONFIG_GEN.TEMPLATE_DATA_ROW, 1, dateValues.length, 2);
  sheet.getRange(CONFIG_GEN.TEMPLATE_DATA_ROW, 1, 31, 2).clearContent();
  targetRange.setValues(dateValues);
  sheet.getRange(CONFIG_GEN.TEMPLATE_DATA_ROW, 1, dateValues.length, 1).setNumberFormat("d");

  if (templateSheet) {
    copyProtectionSettings(templateSheet, sheet);
  }
}

/**
 * テンプレートシートの保護設定を確実にコピーする
 */
function copyProtectionSettings(sourceSheet, targetSheet) {
  const protections = sourceSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);

  protections.forEach(p => {
    // 1. 新しい保護を作成（この時点でデフォルトで「自分のみ」編集可になる）
    const newProtection = targetSheet.protect();
    newProtection.setDescription(p.getDescription()); 
    newProtection.setWarningOnly(p.isWarningOnly()); 

    // 2. テンプレートの「編集許可範囲（UnprotectedRanges）」を取得
    const unprotectedRanges = p.getUnprotectedRanges();
    const newUnprotectedRanges = [];
    
    unprotectedRanges.forEach(range => {
      newUnprotectedRanges.push(targetSheet.getRange(range.getA1Notation()));
    });
    
    // 3. 編集許可範囲をセット
    if (newUnprotectedRanges.length > 0) {
      newProtection.setUnprotectedRanges(newUnprotectedRanges);
    }
    
    // 4. 重要: 共有ドライブ対策
    // 警告のみの設定でない場合（＝完全に編集不可にしたい場合）、明示的にエディタを削除しようとする
    if (!p.isWarningOnly()) {
      // オーナー（実行者）のみが編集できるようにする設定
      // ※共有ドライブではremoveEditorsがエラーになることがあるためtry-catchする
      // しかし、protect()した時点で作成者がオーナーになるため、
      // 共有ドライブ権限者以外に対してはこれでロックがかかる。
      try {
        const me = Session.getEffectiveUser();
        newProtection.addEditor(me);
        newProtection.removeEditors(newProtection.getEditors()); // 自分以外を削除
        if (newProtection.getEditors().length > 1) {
           // それでも消えない（共有ドライブの管理者など）場合は、警告モードにはせずそのままにする
           // GASの仕様上、共有ドライブの管理者を排除できない場合がある
        }
      } catch (e) {
        // 共有ドライブ環境でEditor削除に失敗しても処理を止めない
        console.warn("保護設定: エディタの削除に失敗しましたが保護は適用されました。: " + e.message);
      }
    }
  });
}