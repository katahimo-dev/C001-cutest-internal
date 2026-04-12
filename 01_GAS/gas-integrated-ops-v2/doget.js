/**
 * 【統合版】 gas-integrated-ops-v2 メインロジック
 * 
 * 機能:
 * 1. Webアプリエントリ (doGet)
 * 2. スタッフ認証・権限管理
 * 3. 月間集約シートからのデータ取得・修正
 * 4. カレンダー→月間集約シート自動転記（calendar_sync.js と連携）
 * 5. 個票生成・転記（gas-childcare-daily-report の流用）
 */

// ==========================================
// 1. Webアプリエントリ・メニュー
// ==========================================

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle(CONFIG_UNIFIED.APP_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * スプレッドシート上部のメニューを作成
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('📅カレンダー同期')
    .addItem('本日の予定を取得・反映', 'syncCalendarToday')
    .addItem('指定日の予定を取得・反映', 'syncCalendarSpecificDate')
    .addItem('月次一括同期（YYYY-MM）', 'syncCalendarMonthlyBulk')
    .addToUi();
  
  ui.createMenu('✅管理者用')
    .addItem('個票を生成・更新', 'generateIndividualSheets')
    .addItem('本月の確定', 'confirmMonthData')
    .addToUi();
}

// ==========================================
// 2. スタッフ認証・権限確認
// ==========================================

/**
 * スタッフメールと名前のマッピングを取得
 * （Staff シートから、退職済みを除外）
 */
function getAuthorizedStaffMap() {
  try {
    let ss = null;

    // Webアプリ（standalone）では ActiveSpreadsheet が null になり得るためID優先で開く
    if (CONFIG_UNIFIED.STAFF_SHEET_ID && CONFIG_UNIFIED.STAFF_SHEET_ID !== '(TBD)') {
      ss = SpreadsheetApp.openById(CONFIG_UNIFIED.STAFF_SHEET_ID);
    } else {
      ss = SpreadsheetApp.getActiveSpreadsheet();
    }

    if (!ss) {
      throw new Error('スタッフ参照スプレッドシートを取得できません。CONFIG_UNIFIED.STAFF_SHEET_ID を設定してください。');
    }

    const sheet = ss.getSheetByName(CONFIG_UNIFIED.STAFF_SHEET_NAME);
    
    if (!sheet) {
      throw new Error(CONFIG_UNIFIED.STAFF_SHEET_NAME + ' シートが見つかりません');
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return {};

    // Staff マスター（A〜L）
    // A:ID, B:氏名, C:住所, D:TEL, E:mail address, F:緯度経度,
    // G:入社日, H:退職日, I:ユーザーID, J:PW, K:管理者, L:LW_ID
    const data = sheet.getRange(2, 1, lastRow - 1, 12).getValues();
    const map = {};

    for (let i = 0; i < data.length; i++) {
      const name = data[i][1];       // B列: 氏名
      const email = data[i][4];      // E列: mail address
      const retiredDate = data[i][7]; // H列: 退職日

      if (email && email !== "" && (!retiredDate || retiredDate === "")) {
        map[String(email).trim().toLowerCase()] = name;
      }
    }
    return map;
  } catch (e) {
    console.error('getAuthorizedStaffMap error:', e.message);
    throw e;
  }
}

/**
 * 月間集約シートを取得（sheet.getSheetByName でも、ID指定でもOK）
 */
function getSummarySheet(sheetName = null) {
  try {
    const ss = getOrCreateSummarySpreadsheet();
    
    if (!sheetName) {
      // 本月のシート名を自動生成（YYYYMM形式）
      const today = new Date();
      sheetName = generateMonthSheetName(today);
    }
    
    const sheet = getOrCreateMonthSheet(ss, sheetName);
    return sheet;
  } catch (e) {
    console.error('getSummarySheet error:', e.message);
    throw e;
  }
}

// ==========================================
// 3. 初期データ取得（Web UI用）
// ==========================================

/**
 * Webアプリ起動時に UI に初期情報を提供
 * スタッフ名・選択肢などを返す
 */
function getInitialData() {
  try {
    const email = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
    const staffMap = getAuthorizedStaffMap();
    const staffName = staffMap[email];
    
    if (!staffName) {
      throw new Error(`\n❌ あなたのメールアドレス (${email}) はスタッフ登録されていません。`);
    }

    // 月間集約シートから、該当スタッフの行を全て取得して選択肢も拾ってくる
    const summarySheet = getSummarySheet();
    
    // ★選択肢取得は、月間集約シートの特定行からデータ入力規則を読む
    // （将来的に config.js から選択肢マスターを参照するように改良可能）
    const getOptions = (colChar) => {
      try {
        const range = summarySheet.getRange(colChar + "2"); // 2行目のデータ入力規則を参照
        const rule = range.getDataValidation();
        if (rule) {
          const criteria = rule.getCriteriaType();
          const args = rule.getCriteriaValues();
          if (criteria === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
            return args[0];
          } else if (criteria === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
            const sourceRange = args[0];
            const values = sourceRange.getValues();
            return values.flat().filter(v => v !== "");
          }
        }
      } catch (e) {
        console.warn(`getOptions(${colChar}) failed:`, e.message);
      }
      return [];
    };

    return {
      success: true,
      userName: staffName,
      optionsWorkType: getOptions(SUMMARY_SHEET_COLUMNS.WORK_TYPE),
      optionsWeather: getOptions(SUMMARY_SHEET_COLUMNS.WEATHER),
      optionsAdminWork: getOptions(SUMMARY_SHEET_COLUMNS.IS_ADMIN_WORK),
      currentMonth: generateMonthSheetName(new Date()),
    };
  } catch (e) {
    console.error('getInitialData error:', e.message);
    return { success: false, error: e.toString() };
  }
}

// ==========================================
// 4. 日付別データ取得
// ==========================================

/**
 * Webアプリで日付を選択した時に、該当行のデータを取得
 * @param {string} dateString - 'YYYY-MM-DD' 形式
 */
function getDataByDate(dateString) {
  try {
    const email = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
    const staffMap = getAuthorizedStaffMap();
    const staffName = staffMap[email];

    if (!staffName) {
      throw new Error('スタッフ登録が確認できません');
    }

    const searchDate = new Date(dateString.replace(/-/g, '/'));
    const sheetMonth = generateMonthSheetName(searchDate);
    const summarySheet = getSummarySheet(sheetMonth);
    
    const lastRow = summarySheet.getLastRow();
    if (lastRow < 2) {
      throw new Error('データがありません');
    }

    // 日付・スタッフ名で該当行を探す
    const dateValues = summarySheet.getRange('A2:B' + lastRow).getValues();
    let targetRow = -1;

    for (let i = 0; i < dateValues.length; i++) {
      const date = dateValues[i][0];
      const name = dateValues[i][1];
      
      if (isSameDate(date, searchDate) && name === staffName) {
        targetRow = i + 2; // +1 は header, +1 は 0-indexed
        break;
      }
    }

    if (targetRow === -1) {
      throw new Error(`${staffName} さんの ${dateString} のデータが見つかりません`);
    }

    // 該当行のすべてのデータを取得
    const rowData = {};
    for (const colChar in SUMMARY_SHEET_COLUMNS) {
      const col = SUMMARY_SHEET_COLUMNS[colChar];
      const colNum = columnToNumber(col);
      let val = summarySheet.getRange(targetRow, colNum).getValue();
      
      // 時刻をフォーマット
      if (val instanceof Date) {
        val = Utilities.formatDate(val, 'Asia/Tokyo', 'HH:mm');
      }
      rowData[col] = val;
    }

    // 後続の updateData() で使用するため、ユーザープロパティに row 情報を保存
    PropertiesService.getUserProperties().setProperty('LAST_SUMMARY_SHEET_ROW', targetRow);
    PropertiesService.getUserProperties().setProperty('LAST_SUMMARY_SHEET_MONTH', generateMonthSheetName(searchDate));

    return { 
      success: true, 
      rowData: rowData, 
      rowNumber: targetRow,
      staffName: staffName 
    };
  } catch (e) {
    console.error('getDataByDate error:', e.message);
    return { success: false, error: e.toString() };
  }
}

// ==========================================
// 5. データ修正・保存
// ==========================================

/**
 * Web UI からフォーム送信されたデータを月間集約シートに書き込む
 * @param {Object} formObject - {col_C: "値", col_D: "時刻", ...}
 */
function updateData(formObject) {
  try {
    const email = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
    const staffMap = getAuthorizedStaffMap();
    const staffName = staffMap[email];

    if (!staffName) {
      throw new Error('スタッフ登録が確認できません');
    }

    // 前回の getDataByDate() で保存したrow情報を復元
    const rowNumber = parseInt(PropertiesService.getUserProperties().getProperty('LAST_SUMMARY_SHEET_ROW'));
    const sheetMonth = PropertiesService.getUserProperties().getProperty('LAST_SUMMARY_SHEET_MONTH');
    
    if (!rowNumber || !sheetMonth) {
      throw new Error('セッション情報が失効しています。ブラウザをリロードしてください。');
    }

    // 修正期限チェック（7日前まで）
    const targetDate = new Date();
    targetDate.setHours(0, 0, 0, 0);
    const limitDate = new Date(targetDate);
    limitDate.setDate(targetDate.getDate() - 7);

    const monthYear = sheetMonth.substring(0, 4) + '-' + sheetMonth.substring(4);
    const dateStr = formObject.targetDateStr || ''; // UIから日付が送られてくる前提
    const dataDate = new Date(dateStr.replace(/-/g, '/'));
    dataDate.setHours(0, 0, 0, 0);

    if (dataDate < limitDate) {
      throw new Error(`修正期限切れです（7日以内のみ修正可能）。`);
    }

    const summarySheet = getSummarySheet(sheetMonth);

    // INPUT_COLUMNS_SUMMARY 定義に従って、該当列に書き込む
    for (const colChar in SUMMARY_SHEET_COLUMNS) {
      const col = SUMMARY_SHEET_COLUMNS[colChar];
      const config = INPUT_COLUMNS_SUMMARY[col];
      
      // readonly フラグがあっ場合はスキップ
      if (config && config.readonly) {
        continue;
      }

      const inputKey = 'col_' + col;
      const newValue = formObject[inputKey];
      const colNum = columnToNumber(col);

      if (newValue !== undefined && newValue !== '') {
        const cell = summarySheet.getRange(rowNumber, colNum);
        const oldValue = cell.getValue();
        
        if (!isSameValue(oldValue, newValue)) {
          cell.setValue(newValue);
          cell.setBackground('#fce4e4'); // 修正セルを目立つ色で可視化
        }
      }
    }

    return { 
      success: true, 
      message: '修正完了しました。お疲れ様でした！' 
    };
  } catch (e) {
    console.error('updateData error:', e.message);
    return { success: false, error: e.toString() };
  }
}

// ==========================================
// 6. 管理者用：カレンダー同期トリガー
// ==========================================

/**
 * 本日のカレンダー予定を月間集約シートに同期
 */
function syncCalendarToday() {
  const today = new Date();
  syncCalendarSpecificDate_impl(today);
}

/**
 * 指定日のカレンダー予定を同期（管理者メニュー用）
 */
function syncCalendarSpecificDate() {
  const ui = getUiSafe_();
  if (!ui) {
    throw new Error('UIコンテキストではありません。syncCalendarSpecificDateByArg_("YYYY-MM-DD") を使ってください。');
  }

  const response = ui.prompt('同期する日付を入力してください (YYYY-MM-DD):');
  if (response.getSelectedButton() === ui.Button.OK) {
    syncCalendarSpecificDateByArg_(response.getResponseText());
  }
}

/**
 * UI非依存で指定日同期するための補助関数
 * @param {string} dateStr - YYYY-MM-DD
 */
function syncCalendarSpecificDateByArg_(dateStr) {
  const date = new Date(String(dateStr).replace(/-/g, '/'));
  if (isNaN(date.getTime())) {
    throw new Error('無効な日付です。YYYY-MM-DD 形式で指定してください。');
  }
  syncCalendarSpecificDate_impl(date);
}

/**
 * 指定月を一括同期（YYYY-MM）
 * 例: 2026-03 -> 2026-03-01 から 2026-03-31 まで毎日同期
 */
function syncCalendarMonthlyBulk() {
  const ui = getUiSafe_();
  if (!ui) {
    const ymFromProp = PropertiesService.getScriptProperties().getProperty('BULK_SYNC_YYYYMM');
    if (!ymFromProp) {
      throw new Error('UIコンテキストではありません。Script Properties に BULK_SYNC_YYYYMM=YYYY-MM を設定するか、syncCalendarMonthlyBulkByArg_("YYYY-MM") を実行してください。');
    }
    syncCalendarMonthlyBulkByArg_(ymFromProp);
    return;
  }

  const response = ui.prompt('同期する年月を入力してください (YYYY-MM):');
  if (response.getSelectedButton() !== ui.Button.OK) return;
  syncCalendarMonthlyBulkByArg_(response.getResponseText());
}

/**
 * UI非依存で指定月を一括同期
 * @param {string} ym - YYYY-MM
 */
function syncCalendarMonthlyBulkByArg_(ym) {
  const ymText = String(ym || '').trim();
  const m = ymText.match(/^(\d{4})-(\d{2})$/);
  if (!m) {
    throw new Error('無効な形式です。YYYY-MM で入力してください。');
  }

  const year = parseInt(m[1], 10);
  const month = parseInt(m[2], 10);
  if (month < 1 || month > 12) {
    throw new Error('月は 01〜12 で入力してください。');
  }

  const startDate = new Date(year, month - 1, 1);
  const endDate = new Date(year, month, 0);
  syncCalendarDateRange_impl(startDate, endDate);
}

/**
 * 日付範囲を1日ずつ同期
 * @param {Date} startDate
 * @param {Date} endDate
 */
function syncCalendarDateRange_impl(startDate, endDate) {
  try {
    const ui = getUiSafe_();
    const s = Utilities.formatDate(startDate, 'Asia/Tokyo', 'yyyy-MM-dd');
    const e = Utilities.formatDate(endDate, 'Asia/Tokyo', 'yyyy-MM-dd');

    const diffDays = Math.floor((endDate.getTime() - startDate.getTime()) / (24 * 60 * 60 * 1000)) + 1;
    if (diffDays <= 0) {
      notifyInfo_('❌ 開始日と終了日の指定が不正です。', '入力エラー');
      return;
    }
    if (diffDays > 62) {
      notifyInfo_('❌ 同期範囲が広すぎます。最大62日までにしてください。', '入力エラー');
      return;
    }

    if (ui) {
      const confirm = ui.alert(
        '⚠️ 確認',
        `${s} 〜 ${e}（${diffDays}日分）を一括同期します。\n\n実行してよろしいですか？`,
        ui.ButtonSet.YES_NO
      );
      if (confirm !== ui.Button.YES) return;
    }

    let successDays = 0;
    let totalRows = 0;
    const errors = [];

    const cursor = new Date(startDate);
    while (cursor.getTime() <= endDate.getTime()) {
      const current = new Date(cursor);
      const dayStr = Utilities.formatDate(current, 'Asia/Tokyo', 'yyyy-MM-dd');
      try {
        const result = fetchCalendarAndUpdateSheet(current);
        successDays++;
        totalRows += Number(result.rowsAdded || 0);
        console.log(`[monthly-sync] ${dayStr}: ${result.message}`);
      } catch (eDay) {
        errors.push(`${dayStr}: ${eDay.message}`);
        console.error(`[monthly-sync] ${dayStr} failed:`, eDay.message);
      }
      cursor.setDate(cursor.getDate() + 1);
    }

    let message = `完了: ${s} 〜 ${e}\n成功日数: ${successDays}/${diffDays}\n更新行数: ${totalRows}`;
    if (errors.length > 0) {
      const preview = errors.slice(0, 5).join('\n');
      message += `\n\nエラー(${errors.length}件):\n${preview}`;
      if (errors.length > 5) {
        message += '\n...';
      }
    }
    if (ui) {
      ui.alert(errors.length > 0 ? '⚠️ 一部エラーで完了' : '✅ 一括同期完了', message, ui.ButtonSet.OK);
    } else {
      console.log(`[bulk-sync] ${errors.length > 0 ? '一部エラー' : '完了'}: ${message}`);
    }
  } catch (e) {
    console.error('syncCalendarDateRange_impl error:', e.message);
    notifyInfo_('❌ エラー: ' + e.message, '同期エラー');
  }
}

/**
 * calendar_sync.js を呼び出す（該当日のカレンダーを月間集約シートに反映）
 */
function syncCalendarSpecificDate_impl(targetDate) {
  try {
    console.log('syncCalendarSpecificDate_impl:', Utilities.formatDate(targetDate, 'JST', 'yyyy-MM-dd'));

    const result = fetchCalendarAndUpdateSheet(targetDate);
    notifyInfo_('✅ ' + result.message, '同期完了');
  } catch (e) {
    console.error('syncCalendarSpecificDate_impl error:', e.message);
    notifyInfo_('❌ エラー: ' + e.message, '同期エラー');
  }
}

/**
 * UI取得を安全に行う（UIがない実行コンテキストでは null を返す）
 * @returns {GoogleAppsScript.Base.Ui|null}
 */
function getUiSafe_() {
  try {
    return SpreadsheetApp.getUi();
  } catch (e) {
    return null;
  }
}

/**
 * UIがあれば alert、なければログへ出力
 * @param {string} message
 * @param {string} title
 */
function notifyInfo_(message, title) {
  const ui = getUiSafe_();
  if (ui) {
    ui.alert(title || '通知', message, ui.ButtonSet.OK);
  } else {
    console.log(`[${title || '通知'}] ${message}`);
  }
}

// ==========================================
// 7. 管理者用：個票生成・転記
// ==========================================

/**
 * 個別スタッフの出勤簿ファイルに対して、
 * 月間集約シートのデータを転記 or 初回作成
 */
function generateIndividualSheets() {
  try {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      '⚠️ 確認',
      '本月の月間集約シートから、各スタッフの出勤簿を生成・更新します。\n\n実行してよろしいですか？',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) return;

    // attendance.js / transfer.js の ロジックを流用
    // 個票生成処理の実装は Phase 4 で詳細化
    console.log('generateIndividualSheets: NOT YET IMPLEMENTED');
    
    ui.alert('実装予定', '個票生成処理は Phase 4 で実装予定です', ui.ButtonSet.OK);
  } catch (e) {
    console.error('generateIndividualSheets error:', e.message);
    SpreadsheetApp.getUi().alert('❌ エラー: ' + e.message);
  }
}

/**
 * 本月のデータを「確定」(管理者による最終承認)
 */
function confirmMonthData() {
  try {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      '⚠️ 確認',
      '本月のデータを確定します。\n確定後は修正できなくなります。\n\n実行してよろしいですか？',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) return;

    const today = new Date();
    const sheetMonth = generateMonthSheetName(today);
    const summarySheet = getSummarySheet(sheetMonth);

    // Q列(CONFIRMED_BY) に管理者名、R列(CONFIRMED_AT) に日時を記入
    const confirmEmail = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
    const adminName = getAuthorizedStaffMap()[confirmEmail] || confirmEmail;
    const now = new Date();
    const confirmedAtStr = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');

    const lastRow = summarySheet.getLastRow();
    const colQ = columnToNumber(SUMMARY_SHEET_COLUMNS.CONFIRMED_BY);
    const colR = columnToNumber(SUMMARY_SHEET_COLUMNS.CONFIRMED_AT);

    // 全行に管理者名・確定日時を記入（既に確定済みの行は上書きしない）
    for (let row = 2; row <= lastRow; row++) {
      const qCell = summarySheet.getRange(row, colQ);
      if (!qCell.getValue() || qCell.getValue() === '') {
        qCell.setValue(adminName);
        summarySheet.getRange(row, colR).setValue(confirmedAtStr);
      }
    }

    ui.alert('✅ 完了', `本月 (${sheetMonth}) のデータが確定しました`, ui.ButtonSet.OK);
  } catch (e) {
    console.error('confirmMonthData error:', e.message);
    SpreadsheetApp.getUi().alert('❌ エラー: ' + e.message);
  }
}

// ==========================================
// 8. その他ユーティリティ (config.js から重複回避)
// ==========================================

// 列文字→列番号, 列番号→列文字 は config.js から引き込むか、ここでインライン
// isSameDate, isSameValue も config.js から引き込む

/**
 * メール送信（管理者通知など）
 */
function sendAppUrlToSelected() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG_UNIFIED.STAFF_SHEET_NAME);
    
    // Store で WebApp URL を取得（H8 セルに保存されている想定）
    const appUrl = sheet.getRange('H8').getValue();
    
    if (!appUrl || !String(appUrl).match(/^https:\/\/script\.google\.com/)) {
      SpreadsheetApp.getUi().alert('❌ エラー', 'H8セルに正しいアプリのURL（https://...）が入力されていません。');
      return;
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
    
    const data = sheet.getRange(2, 3, lastRow - 1, 5).getValues();
    let sentCount = 0;

    for (let i = 0; i < data.length; i++) {
      const rowNum = i + 2;
      const name = data[i][0];
      const email = data[i][1];
      const retiredDate = data[i][3];
      const isChecked = data[i][4];

      if (isChecked === true && email && (!retiredDate || retiredDate === "")) {
        try {
          const htmlBody = 
            `${name} さん<br><br>` +
            'お疲れ様です。<br>' +
            '勤務管理アプリのURLをお送りします。<br><br>' +
            `<a href="${appUrl}" style="font-size: 16px;">👉 アプリURL</a><br><br>`;

          MailApp.sendEmail({
            to: email,
            subject: '【業務連絡】勤務管理アプリのURLをお知らせします',
            htmlBody: htmlBody
          });
          
          sheet.getRange(rowNum, 7).setValue(false);
          sentCount++;
        } catch (e) {
          console.warn(`Mail send failed for ${name}:`, e.message);
        }
      }
    }

    const ui = SpreadsheetApp.getUi();
    if (sentCount > 0) {
      ui.alert('✅ 完了', `${sentCount} 名にメールを送信しました`, ui.ButtonSet.OK);
    } else {
      ui.alert('ℹ️ 確認', '送信対象がいませんでした', ui.ButtonSet.OK);
    }
  } catch (e) {
    console.error('sendAppUrlToSelected error:', e.message);
    SpreadsheetApp.getUi().alert('❌ エラー: ' + e.message);
  }
}

/**
 * 権限を強制設定（初期セットアップ用）
 */
function forceAuth() {
  try {
    MailApp.getRemainingDailyQuota();
    DriveApp.getRootFolder();
    SpreadsheetApp.create("Dummy");
    console.log("✅ すべての権限認証に成功しました！");
  } catch (e) {
    console.error('forceAuth error:', e.message);
  }
}
