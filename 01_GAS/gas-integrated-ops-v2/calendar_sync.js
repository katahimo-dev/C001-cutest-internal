/**
 * 【統合版】 calendar_sync.js
 * 
 * 機能: Google Calendar から予定を読み込み、月間集約シートに転記
 * 
 * 参考: gas-root-serach の main.js からロジックを抽出・改良
 */

/**
 * ★フェーズ2実装: カレンダーから月間集約シートへデータ取得・転記
 * 
 * 処理フロー:
 * 1. targetDate のカレンダー予定をすべて取得
 * 2. タイトルのタグ（[予約確定]等）で種別判定
 * 3. スタッフ名を抽出（招待ゲストから照合）
 * 4. 月間集約シートの該当（日付×スタッフ）行に書き込む
 * 
 * ※ルート計算（Google Maps API）は本PJでは未対応
 */

/**
 * カレンダーから指定日のイベントを取得し、
 * 月間集約シートに転記する
 * 
 * @param {Date} targetDate
 */
function fetchCalendarAndUpdateSheet(targetDate) {
  try {
    const dateStr = Utilities.formatDate(targetDate, 'Asia/Tokyo', 'yyyy-MM-dd');
    const sheetMonth = generateMonthSheetName(targetDate);
    
    console.log(`[calendar_sync] 開始: ${dateStr} のカレンダー同期`);
    
    // 1. カレンダーイベントを取得
    const events = fetchCalendarEvents(targetDate);
    console.log(`[calendar_sync] イベント数: ${events.length}`);
    
    if (events.length === 0) {
      console.log(`[calendar_sync] ${dateStr} には予定がありません`);
      return {
        success: true,
        message: `${dateStr} には予定がありません`
      };
    }

    // 2. スタッフマスターを取得
    const staffList = getStaffDataFromSpreadsheet(CONFIG_UNIFIED.STAFF_SHEET_ID);
    const staffMap = {};
    staffList.forEach(s => {
      staffMap[s.name.replace(/\s+/g, '')] = s;
    });
    
    // 3. イベントを "スタッフ単位" にグループ化
    const scheduleByStaff = groupEventsByStaff(events, staffMap, targetDate);
    console.log(`[calendar_sync] グループ化後: ${Object.keys(scheduleByStaff).length} 名のスタッフ`);
    
    // 4. 月間集約シートに書き込む
    const summarySheet = getSummarySheet(sheetMonth);
    const rowsAdded = updateSummarySheetWithSchedule(summarySheet, dateStr, scheduleByStaff);
    
    console.log(`[calendar_sync] ${rowsAdded} 行を更新しました`);
    return {
      success: true,
      message: `${dateStr} の予定を ${rowsAdded} 件 転記しました`,
      rowsAdded: rowsAdded
    };
  } catch (e) {
    console.error('[calendar_sync] error:', e.message);
    throw e;
  }
}

/**
 * ===============================================
 * 1. カレンダーイベント取得
 * ===============================================
 */

/**
 * 指定日のカレンダーイベントをすべてのカレンダーから取得
 * 
 * @param {Date} targetDate
 * @returns {Array} イベント配列
 */
function fetchCalendarEvents(targetDate) {
  const calendars = CalendarApp.getAllCalendars();
  const events = [];
  const uniqueEventsMap = new Map(); // 重複排除用
  
  const startTime = new Date(targetDate);
  startTime.setHours(0, 0, 0, 0);
  
  const endTime = new Date(targetDate);
  endTime.setHours(23, 59, 59, 999);

  console.log(`[fetchCalendarEvents] ${calendars.length} 個のカレンダーをスキャン`);

  for (let i = 0; i < calendars.length; i++) {
    const cal = calendars[i];
    try {
      const dayEvents = cal.getEvents(startTime, endTime);
      
      for (let j = 0; j < dayEvents.length; j++) {
        const evt = dayEvents[j];
        const eventId = evt.getId();
        
        // 重複排除（同じイベントが複数のカレンダーに登録されている場合）
        if (!uniqueEventsMap.has(eventId)) {
          uniqueEventsMap.set(eventId, evt);
        }
      }
    } catch (e) {
      console.warn(`[fetchCalendarEvents] カレンダー「${cal.getName()}」の読込エラー:`, e.message);
    }
  }

  // Map から配列に変換
  uniqueEventsMap.forEach(evt => {
    events.push({
      id: evt.getId(),
      title: evt.getTitle(),
      description: evt.getDescription(),
      startTime: evt.getStartTime(),
      endTime: evt.getEndTime(),
      guests: evt.getGuestList(),
      location: evt.getLocation() || ''
    });
  });

  return events;
}

/**
 * ===============================================
 * 2. イベント分類・スタッフ特定
 * ===============================================
 */

/**
 * イベント配列をスタッフ単位でグループ化
 * 
 * @param {Array} events - イベント配列
 * @param {Object} staffMap - スタッフマップ { normalizedName -> staffInfo }
 * @param {Date} targetDate - 処理対象日
 * @returns {Object} { staffName -> [schedule1, schedule2, ...] }
 */
function groupEventsByStaff(events, staffMap, targetDate) {
  const grouped = {};

  for (let i = 0; i < events.length; i++) {
    const evt = events[i];
    const classified = classifyEvent(evt);
    
    if (!classified) {
      console.log(`[groupEventsByStaff] スキップ: ${evt.title}`);
      continue; // 不明なタイプはスキップ
    }

    // スタッフ名を特定
    let staffNames = [];
    
    if (classified.type === 'VISIT' || classified.type === 'VISIT_NEW') {
      // [予約確定] / [新規]: 本文の「施設：〇〇」から施設ID/スタッフを抽出（カスタムロジック）
      // 簡易版: 施設マスタのロジックはスキップ、代わりに説明文から「スタッフ」の記述を探す
      staffNames = extractStaffNamesFromDescription(evt.description || '');
      
      if (staffNames.length === 0) {
        // スタッフが未指定の場合、招待ゲストから推定
        staffNames = extractStaffFromGuests(evt.guests, staffMap);
      }
    } else if (classified.type === 'EVENT') {
      // [イベント]: 招待ゲストから特定
      staffNames = extractStaffFromGuests(evt.guests, staffMap);
    } else if (classified.type === 'ADMIN') {
      // [事務]: 説明文またはゲストから（重要度: 低）
      staffNames = extractStaffFromGuests(evt.guests, staffMap);
    }

    // 各スタッフの grouped に登録
    for (let j = 0; j < staffNames.length; j++) {
      const staffName = staffNames[j];
      if (!grouped[staffName]) {
        grouped[staffName] = [];
      }
      
      grouped[staffName].push({
        type: classified.type,
        title: evt.title,
        startTime: evt.startTime,
        endTime: evt.endTime,
        location: evt.location,
        description: evt.description || ''
      });
    }
  }

  return grouped;
}

/**
 * イベントタイトルからタグを抽出し、種別を判定
 * 
 * @param {Object} evt - イベントオブジェクト
 * @returns {Object|null} { type: 'VISIT'|'VISIT_NEW'|'EVENT'|'ADMIN'|null, tag: string }
 */
function classifyEvent(evt) {
  const title = evt.title || '';
  
  // オンライン判定（説明にオンラインが含まれている場合はスキップ）
  if ((evt.description || '').includes('オンライン')) {
    return null;
  }

  if (title.includes('[予約確定]')) {
    return { type: 'VISIT', tag: '[予約確定]' };
  } else if (title.includes('[新規]')) {
    return { type: 'VISIT_NEW', tag: '[新規]' };
  } else if (title.includes('[イベント]')) {
    return { type: 'EVENT', tag: '[イベント]' };
  } else if (title.includes('[事務]')) {
    return { type: 'ADMIN', tag: '[事務]' };
  }

  return null; // タグがない、または不明
}

/**
 * イベント説明からスタッフ名を抽出
 * 簡易実装: 「スタッフ：○○」という記述を探す
 * 
 * @param {string} description
 * @returns {Array} スタッフ名配列
 */
function extractStaffNamesFromDescription(description) {
  const staffNames = [];
  const patterns = [
    /スタッフ[：:]\s*([^\n(（]*)/g,
    /担当[：:]\s*([^\n(（]*)/g,
    /訪問者[：:]\s*([^\n(（]*)/g
  ];

  for (let i = 0; i < patterns.length; i++) {
    let match;
    while ((match = patterns[i].exec(description)) !== null) {
      const name = match[1].trim();
      if (name && staffNames.indexOf(name) === -1) {
        staffNames.push(name);
      }
    }
  }

  return staffNames;
}

/**
 * イベント招待ゲストからスタッフ名を抽出
 * 
 * @param {Array} guests - ゲストリスト
 * @param {Object} staffMap - スタッフマップ
 * @returns {Array} スタッフ名配列
 */
function extractStaffFromGuests(guests, staffMap) {
  const staffNames = [];

  for (let i = 0; i < guests.length; i++) {
    const guest = guests[i];
    const email = guest.getEmail() || '';
    
    // Email がスタッフマップに存在するか確認
    for (const normalizedName in staffMap) {
      const staffInfo = staffMap[normalizedName];
      if (staffInfo.email && email.toLowerCase() === staffInfo.email.toLowerCase()) {
        if (staffNames.indexOf(staffInfo.name) === -1) {
          staffNames.push(staffInfo.name);
        }
      }
    }
  }

  return staffNames;
}

/**
 * ===============================================
 * 3. 月間集約シートへの書き込み
 * ===============================================
 */

/**
 * 月間集約シートにスケジュール情報を書き込む
 * 
 * @param {Sheet} summarySheet - 月間集約シート
 * @param {string} dateStr - 'YYYY-MM-DD'
 * @param {Object} scheduleByStaff - { staffName -> [schedule, ...] }
 * @returns {number} 追加/更新した行数
 */
function updateSummarySheetWithSchedule(summarySheet, dateStr, scheduleByStaff) {
  let rowsAdded = 0;

  // 月間集約シートの既存行を確認し、該当（日付×スタッフ）の行を探す or 新規作成
  for (const staffName in scheduleByStaff) {
    const schedules = scheduleByStaff[staffName];
    
    if (schedules.length === 0) continue;

    // 該当行を検索（日付×スタッフ名で一意に決定）
    let targetRow = findOrCreateSummaryRow(summarySheet, dateStr, staffName);
    
    // スケジュール情報を書き込む
    writeScheduleToRow(summarySheet, targetRow, dateStr, staffName, schedules);
    rowsAdded++;
  }

  return rowsAdded;
}

/**
 * 月間集約シートで（日付×スタッフ）の行を検索し、
 * なければ新規行を追加
 * 
 * @param {Sheet} sheet
 * @param {string} dateStr - 'YYYY-MM-DD'
 * @param {string} staffName
 * @returns {number} 行番号
 */
function findOrCreateSummaryRow(sheet, dateStr, staffName) {
  const lastRow = sheet.getLastRow();
  const col_A = columnToNumber(SUMMARY_SHEET_COLUMNS.DATE);
  const col_B = columnToNumber(SUMMARY_SHEET_COLUMNS.STAFF_NAME);

  // 既存行を検索
  for (let row = 2; row <= lastRow; row++) {
    const dateCell = sheet.getRange(row, col_A).getValue();
    const nameCell = sheet.getRange(row, col_B).getValue();
    
    // 日付とスタッフ名が一致する行を発見
    if (isSameDate(dateCell, new Date(dateStr.replace(/-/g, '/'))) && 
        nameCell === staffName) {
      return row;
    }
  }

  // 該当行がない場合、新規行を追加
  const newRow = lastRow + 1;
  const col_A_num = columnToNumber(SUMMARY_SHEET_COLUMNS.DATE);
  const col_B_num = columnToNumber(SUMMARY_SHEET_COLUMNS.STAFF_NAME);
  
  sheet.getRange(newRow, col_A_num).setValue(new Date(dateStr.replace(/-/g, '/')));
  sheet.getRange(newRow, col_B_num).setValue(staffName);
  
  return newRow;
}

/**
 * スケジュール情報を月間集約シートの行に書き込む
 * 
 * @param {Sheet} sheet
 * @param {number} rowNum
 * @param {string} dateStr
 * @param {string} staffName
 * @param {Array} schedules - [{ type, title, startTime, endTime, location, description }, ...]
 */
function writeScheduleToRow(sheet, rowNum, dateStr, staffName, schedules) {
  // 訪問スケジュール（種別別に分類）
  const visits = schedules.filter(s => s.type === 'VISIT' || s.type === 'VISIT_NEW');
  const adminItems = schedules.filter(s => s.type === 'ADMIN');

  // 訪問先1-3を書き込む
  let visitIndex = 0;
  for (let i = 0; i < Math.min(visits.length, 3); i++) {
    const visit = visits[i];
    const colPrefix = i === 0 ? 'VISIT_1_' : i === 1 ? 'VISIT_2_' : 'VISIT_3_';
    
    // 訪問先名
    const colName = SUMMARY_SHEET_COLUMNS[colPrefix + 'NAME'];
    sheet.getRange(rowNum, columnToNumber(colName)).setValue(visit.location || visit.title);
    
    // 開始時刻
    const colStart = SUMMARY_SHEET_COLUMNS[colPrefix + 'START'];
    sheet.getRange(rowNum, columnToNumber(colStart)).setValue(visit.startTime);
    
    // 終了時刻
    const colEnd = SUMMARY_SHEET_COLUMNS[colPrefix + 'END'];
    sheet.getRange(rowNum, columnToNumber(colEnd)).setValue(visit.endTime);
  }

  // 事務作業フラグ
  if (adminItems.length > 0) {
    const colP = SUMMARY_SHEET_COLUMNS.IS_ADMIN_WORK;
    sheet.getRange(rowNum, columnToNumber(colP)).setValue('Yes');
  }

  console.log(`[writeScheduleToRow] ${staffName} - ${dateStr} を更新`);
}

/**
 * ===============================================
 * 4. ユーティリティ関数
 * ===============================================
 */

/**
 * スタッフマスター（外部ファイル）を読み込む
 * gas-root-serach の getStaffDataFromSpreadsheet() を参考
 * 
 * @param {string} spreadsheetId
 * @returns {Array}  [{ id, name, email, lwId }, ...]
 */
function getStaffDataFromSpreadsheet(spreadsheetId) {
  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheet = ss.getSheets()[0]; // デフォルトシート
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    // Staff マスター（A〜L）
    // A:ID, B:氏名, E:mail address, H:退職日, K:管理者, L:LW_ID
    const data = sheet.getRange(2, 1, lastRow - 1, 12).getValues();
    const staffList = [];

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const id = row[0];
      const name = row[1];
      const email = row[4] || ''; // E列
      const retiredDate = row[7]; // H列
      const isAdmin = row[10] || ''; // K列
      const lwId = row[11] || ''; // L列

      if (name && name !== '' && (!retiredDate || retiredDate === '')) {
        staffList.push({
          id: id || '',
          name: name,
          email: email || '',
          lwId: lwId,
          isAdmin: isAdmin
        });
      }
    }

    return staffList;
  } catch (e) {
    console.error('[getStaffDataFromSpreadsheet] error:', e.message);
    throw e;
  }
}

/**
 * 月間集約シートを取得
 * doget.js と共通のヘルパー
 */
function getSummarySheet(sheetName) {
  try {
    const ss = getOrCreateSummarySpreadsheet();
    const sheet = getOrCreateMonthSheet(ss, sheetName);
    return sheet;
  } catch (e) {
    console.error('[getSummarySheet] error:', e.message);
    throw e;
  }
}
