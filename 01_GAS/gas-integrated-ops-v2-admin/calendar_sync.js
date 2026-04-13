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
 * 4. 顧客CSV・スタッフ住所をもとに移動/出退勤経路を計算
 * 5. 月間集約シートの該当（日付×スタッフ）行に書き込む
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

    // 2. スタッフマスターを取得（同期用タイプ）
    const staffMap = getAuthorizedStaffMapForSync();

    // 3. 顧客CSVを取得（ルート集計で使用）
    let customerList = [];
    try {
      customerList = getCustomerDataForRouting_();
    } catch (customerError) {
      console.warn('[calendar_sync] 顧客CSV読込失敗。顧客名のみで継続します:', customerError.message);
    }
    
    // 4. イベントを "スタッフ単位" にグループ化
    const scheduleByStaff = groupEventsByStaff(events, staffMap, targetDate, customerList);
    applyRouteInfoToSchedules_(scheduleByStaff, staffMap, targetDate);
    console.log(`[calendar_sync] グループ化後: ${Object.keys(scheduleByStaff).length} 名のスタッフ`);
    
    // 5. 月間集約シートに書き込む
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
 * 各イベントに対して、所有者（カレンダー名）を保持する
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

  // ===== gas-root-serach に合わせる =====
  // 各カレンダーをループして、誰のカレンダーかを保持する
  for (let i = 0; i < calendars.length; i++) {
    const cal = calendars[i];
    const calendarName = cal.getName(); // カレンダー名（スタッフ名）
    
    try {
      const dayEvents = cal.getEvents(startTime, endTime);
      
      for (let j = 0; j < dayEvents.length; j++) {
        const evt = dayEvents[j];
        const eventId = evt.getId();
        
        // 重複排除（同じイベントが複数のカレンダーに登録されている場合）
        if (!uniqueEventsMap.has(eventId)) {
          uniqueEventsMap.set(eventId, {
            event: evt,
            ownerName: calendarName  // ← カレンダーの所有者を記録
          });
        }
      }
    } catch (e) {
      console.warn(`[fetchCalendarEvents] カレンダー「${cal.getName()}」の読込エラー:`, e.message);
    }
  }

  // Map から配列に変換
  uniqueEventsMap.forEach(data => {
    const evt = data.event;
    const ownerName = data.ownerName;
    
    events.push({
      id: evt.getId(),
      title: evt.getTitle(),
      description: evt.getDescription() || '',
      startTime: evt.getStartTime(),
      endTime: evt.getEndTime(),
      guests: evt.getGuestList(),
      location: evt.getLocation() || '',
      ownerName: ownerName  // ← スタッフ候補として保持
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
 * @param {Array} customerList - 顧客一覧
 * @returns {Object} { staffName -> [schedule1, schedule2, ...] }
 */
function groupEventsByStaff(events, staffMap, targetDate, customerList) {
  const grouped = {};
  const normalize = (str) => str ? str.replace(/\s+/g, "") : "";

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
      // [予約確定] / [新規]: 説明文またはタイトルからスタッフ名を抽出
      staffNames = extractStaffNamesFromDescription(evt.description || '', evt.title, evt, staffMap);
      
      if (staffNames.length === 0) {
        // スタッフ抽出失敗時は、ゲストから推定を試みる
        staffNames = extractStaffFromGuests(evt.guests, staffMap);
      }
    } else if (classified.type === 'EVENT') {
      // [イベント]: 説明文またはゲストから特定
      staffNames = extractStaffNamesFromDescription(evt.description || '', evt.title, evt, staffMap);
      
      if (staffNames.length === 0) {
        staffNames = extractStaffFromGuests(evt.guests, staffMap);
      }
    } else if (classified.type === 'ADMIN') {
      // [事務]: 説明文またはゲストから特定
      staffNames = extractStaffNamesFromDescription(evt.description || '', evt.title, evt, staffMap);
      
      if (staffNames.length === 0) {
        staffNames = extractStaffFromGuests(evt.guests, staffMap);
      }
    }

    // ===== gas-root-serach に合わせた照合 =====
    // staffNames は抽出した「名前」の配列
    // staffMap で、normalize して検索
    for (let j = 0; j < staffNames.length; j++) {
      const staffNameCandidate = staffNames[j];
      
      // staffMap から該当スタッフを検索（normalize して比較）
      let matchedStaff = null;
      for (const normalizedKey in staffMap) {
        const staffInfo = staffMap[normalizedKey];
        if (normalize(staffInfo.name) === normalize(staffNameCandidate)) {
          matchedStaff = staffInfo;
          break;
        }
      }

      if (matchedStaff) {
        const staffName = matchedStaff.name;
        if (!grouped[staffName]) {
          grouped[staffName] = [];
        }

        grouped[staffName].push(buildScheduleFromEvent_(evt, classified, matchedStaff, customerList));
      } else {
        // デバッグ: マッチしなかったスタッフ名をログ
        console.log(`[groupEventsByStaff] マッチなし: ${staffNameCandidate} (イベント: ${evt.title})`);
      }
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
  const title = normalizeEventTitle_(evt.title || '');
  
  // オンライン判定（説明にオンラインが含まれている場合はスキップ）
  if ((evt.description || '').includes('オンライン')) {
    return null;
  }

  const durationMin = (evt.endTime && evt.startTime)
    ? (evt.endTime.getTime() - evt.startTime.getTime()) / (1000 * 60)
    : 0;
  if (durationMin === 15) {
    return { type: 'ADMIN', tag: '[15分イベント]' };
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
 * gas-root-serach に合わせた実装
 * 1. 説明文から「施設：〇〇[」パターンを検索
 * 2. 失敗時は、カレンダー所有者（ownerName）を使用
 * 3. [独立] などの括弧内テキストを除去
 * 
 * @param {string} description
 * @param {string} title - イベントタイトル
 * @param {Object} evt - イベント元オブジェクト（ownerName を持つ）
 * @param {Object} staffMap - スタッフマップ
 * @returns {Array} スタッフ名配列
 */
function extractStaffNamesFromDescription(description, title, evt, staffMap) {
  const staffNames = [];
  const normalize = (str) => str ? str.replace(/\s+/g, "") : "";

  // ===== パターン1: 説明文から「施設：〇〇[」を探す =====
  if (description) {
    const staffMatch = description.match(/施設：(.*?)\[/);
    if (staffMatch) {
      let staffName = staffMatch[1].trim();
      if (staffName && staffNames.indexOf(staffName) === -1) {
        staffNames.push(staffName);
        return staffNames;
      }
    }
  }

  // ===== パターン2: タイトルから全角スペース区切りで抽出 =====
  if (title) {
    // [タグ] を除去
    let titleWithoutTag = title.replace(/\[.*?\]$/, '').trim();
    
    // 最後の「　」（全角スペース）で分割してスタッフ名候補を抽出
    const parts = titleWithoutTag.split('　');
    if (parts.length >= 2) {
      const possibleStaffName = parts[parts.length - 1].trim();
      if (possibleStaffName && staffNames.indexOf(possibleStaffName) === -1) {
        staffNames.push(possibleStaffName);
        return staffNames;
      }
    }
  }

  // ===== パターン3: フォールバック - カレンダー所有者 =====
  if (evt.ownerName && staffNames.indexOf(evt.ownerName) === -1) {
    staffNames.push(evt.ownerName);
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
  clearDerivedScheduleCells_(sheet, rowNum);

  // 訪問/イベントスケジュール（種別別に分類）
  const visits = schedules.filter(s => s.type !== 'ADMIN');
  const adminItems = schedules.filter(s => s.type === 'ADMIN');

  // 訪問先1-3と勤怠計算用補助列を書き込む
  for (let i = 0; i < Math.min(visits.length, 3); i++) {
    const visit = visits[i];
    const colPrefix = i === 0 ? 'VISIT_1_' : i === 1 ? 'VISIT_2_' : 'VISIT_3_';
    
    // 訪問先名
    const colName = SUMMARY_SHEET_COLUMNS[colPrefix + 'NAME'];
    sheet.getRange(rowNum, columnToNumber(colName)).setValue(getScheduleDisplayName_(visit));
    
    // 開始時刻
    const colStart = SUMMARY_SHEET_COLUMNS[colPrefix + 'START'];
    sheet.getRange(rowNum, columnToNumber(colStart)).setValue(visit.startTime);
    
    // 終了時刻
    const colEnd = SUMMARY_SHEET_COLUMNS[colPrefix + 'END'];
    sheet.getRange(rowNum, columnToNumber(colEnd)).setValue(visit.endTime);

    // 予約詳細URL
    const reservaCol = SUMMARY_SHEET_COLUMNS[colPrefix + 'RESERVA_URL'];
    if (reservaCol && visit.reservaUrl) {
      sheet.getRange(rowNum, columnToNumber(reservaCol)).setValue(visit.reservaUrl);
    }

    // 出勤/移動経路情報
    if (i === 0 && visit.attendanceRoute) {
      if (visit.attendanceRoute.url) {
        sheet.getRange(rowNum, columnToNumber(SUMMARY_SHEET_COLUMNS.VISIT_1_ATTENDANCE_URL)).setValue(visit.attendanceRoute.url);
      }
      if (visit.attendanceRoute.min !== '') {
        sheet.getRange(rowNum, columnToNumber(SUMMARY_SHEET_COLUMNS.VISIT_1_ATTENDANCE_MIN)).setValue(visit.attendanceRoute.min);
      }
      if (visit.attendanceRoute.km !== '') {
        sheet.getRange(rowNum, columnToNumber(SUMMARY_SHEET_COLUMNS.VISIT_1_ATTENDANCE_KM)).setValue(visit.attendanceRoute.km);
      }
    }

    if (i === 1 || i === 2) {
      const routeInfo = visit.moveRoute || { url: '', min: '', km: '' };
      const suffix = i === 1 ? 'VISIT_2_MOVE_' : 'VISIT_3_MOVE_';
      if (routeInfo.url) {
        sheet.getRange(rowNum, columnToNumber(SUMMARY_SHEET_COLUMNS[suffix + 'URL'])).setValue(routeInfo.url);
      }
      if (routeInfo.min !== '') {
        sheet.getRange(rowNum, columnToNumber(SUMMARY_SHEET_COLUMNS[suffix + 'MIN'])).setValue(routeInfo.min);
      }
      if (routeInfo.km !== '') {
        sheet.getRange(rowNum, columnToNumber(SUMMARY_SHEET_COLUMNS[suffix + 'KM'])).setValue(routeInfo.km);
      }
    }
  }

  // 退勤経路
  const lastVisit = visits.length > 0 ? visits[Math.min(visits.length, 3) - 1] : null;
  if (lastVisit && lastVisit.leavingRoute) {
    if (lastVisit.leavingRoute.url) {
      sheet.getRange(rowNum, columnToNumber(SUMMARY_SHEET_COLUMNS.LEAVING_URL)).setValue(lastVisit.leavingRoute.url);
    }
    if (lastVisit.leavingRoute.min !== '') {
      sheet.getRange(rowNum, columnToNumber(SUMMARY_SHEET_COLUMNS.LEAVING_MIN)).setValue(lastVisit.leavingRoute.min);
    }
    if (lastVisit.leavingRoute.km !== '') {
      sheet.getRange(rowNum, columnToNumber(SUMMARY_SHEET_COLUMNS.LEAVING_KM)).setValue(lastVisit.leavingRoute.km);
    }
  }

  // 事務作業フラグと作業欄
  if (adminItems.length > 0) {
    const colP = SUMMARY_SHEET_COLUMNS.IS_ADMIN_WORK;
    sheet.getRange(rowNum, columnToNumber(colP)).setValue('Yes');

    const workTitles = adminItems.map(item => stripEventTags_(item.title)).filter(Boolean);
    if (workTitles.length > 0) {
      sheet.getRange(rowNum, columnToNumber(SUMMARY_SHEET_COLUMNS.WORK_TYPE)).setValue(workTitles.join(' / '));
    }

    const sortedAdmin = adminItems.slice().sort((a, b) => a.startTime - b.startTime);
    sheet.getRange(rowNum, columnToNumber(SUMMARY_SHEET_COLUMNS.WORK_START)).setValue(sortedAdmin[0].startTime);
    sheet.getRange(rowNum, columnToNumber(SUMMARY_SHEET_COLUMNS.WORK_END)).setValue(sortedAdmin[sortedAdmin.length - 1].endTime);
  }

  console.log(`[writeScheduleToRow] ${staffName} - ${dateStr} を更新`);
}

/**
 * 同期対象の派生セルをクリアする。
 * A/B/O/Q/R は保持する。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} rowNum
 */
function clearDerivedScheduleCells_(sheet, rowNum) {
  sheet.getRange(rowNum, columnToNumber('C'), 1, columnToNumber('N') - columnToNumber('C') + 1).clearContent();
  sheet.getRange(rowNum, columnToNumber('P')).clearContent();
  sheet.getRange(rowNum, columnToNumber('S'), 1, columnToNumber('AG') - columnToNumber('S') + 1).clearContent();
}

/**
 * 顧客CSVを読み込む。実行中はメモリキャッシュを使う。
 * @returns {Array}
 */
function getCustomerDataForRouting_() {
  if (typeof globalThis.__CUSTOMER_DATA_CACHE__ !== 'undefined' && globalThis.__CUSTOMER_DATA_CACHE__) {
    return globalThis.__CUSTOMER_DATA_CACHE__;
  }

  const folder = DriveApp.getFolderById(CONFIG_UNIFIED.CUSTOMER_FOLDER_ID);
  const customerData = getCustomerDataFromCsvForRouting_(folder);
  globalThis.__CUSTOMER_DATA_CACHE__ = customerData;
  return customerData;
}

/**
 * イベントを月間集約用スケジュールへ変換する。
 * @param {Object} evt
 * @param {Object} classified
 * @param {Object} staffInfo
 * @param {Array} customerList
 * @returns {Object}
 */
function buildScheduleFromEvent_(evt, classified, staffInfo, customerList) {
  const subjectName = stripEventTags_(evt.title || '');
  const reservaUrl = extractReservationUrl_(evt.description || '');

  const schedule = {
    type: classified.type,
    title: evt.title,
    subjectName: subjectName,
    startTime: evt.startTime,
    endTime: evt.endTime,
    location: evt.location,
    description: evt.description || '',
    reservaUrl: reservaUrl,
    staffInfo: staffInfo,
    attendanceRoute: { url: '', min: '', km: '' },
    moveRoute: { url: '', min: '', km: '' },
    leavingRoute: { url: '', min: '', km: '' },
    customerInfo: {
      id: '',
      name: subjectName,
      address: '',
      lat: '',
      lng: '',
      address2: '',
      address2_start: '',
      address2_end: ''
    }
  };

  if (classified.type === 'VISIT') {
    const matchedCustomer = findCustomerBySubject_(subjectName, customerList);
    if (matchedCustomer) {
      schedule.customerInfo = matchedCustomer;
    }
  } else if (classified.type === 'VISIT_NEW' || classified.type === 'EVENT') {
    const geo = evt.location ? getLatLngFromAddressForRouting_(evt.location) : { lat: '', lng: '' };
    schedule.customerInfo = {
      id: '',
      name: subjectName,
      address: evt.location || '',
      lat: geo.lat,
      lng: geo.lng,
      address2: '',
      address2_start: '',
      address2_end: ''
    };
  }

  return schedule;
}

/**
 * スタッフごとのスケジュールへ経路情報を付与する。
 * @param {Object} scheduleByStaff
 * @param {Object} staffMap
 * @param {Date} targetDate
 */
function applyRouteInfoToSchedules_(scheduleByStaff, staffMap, targetDate) {
  for (const staffName in scheduleByStaff) {
    const schedules = scheduleByStaff[staffName].slice().sort((a, b) => a.startTime - b.startTime);
    scheduleByStaff[staffName] = schedules;

    const validLocSchedules = schedules.filter(hasRouteLocation_);
    const staffInfo = staffMap[staffName];
    if (!staffInfo) {
      continue;
    }

    for (let i = 0; i < validLocSchedules.length; i++) {
      const current = validLocSchedules[i];
      if (i === 0) {
        current.attendanceRoute = getRouteDetailsForRouting_(staffInfo, current.customerInfo, targetDate);
      } else {
        current.moveRoute = getRouteDetailsForRouting_(validLocSchedules[i - 1].customerInfo, current.customerInfo, targetDate);
      }

      if (i === validLocSchedules.length - 1) {
        current.leavingRoute = getRouteDetailsForRouting_(current.customerInfo, staffInfo, targetDate);
      }
    }
  }
}

/**
 * ルート計算対象かどうかを返す。
 * @param {Object} schedule
 * @returns {boolean}
 */
function hasRouteLocation_(schedule) {
  return !!(
    schedule &&
    schedule.customerInfo &&
    ((schedule.customerInfo.lat && schedule.customerInfo.lng) ||
      (schedule.customerInfo.address && String(schedule.customerInfo.address).trim() !== ''))
  );
}

/**
 * 顧客名をイベントタイトルから検索する。
 * @param {string} subjectName
 * @param {Array} customerList
 * @returns {Object|null}
 */
function findCustomerBySubject_(subjectName, customerList) {
  const normalizedSubject = normalizeForMatch_(subjectName);
  for (let i = 0; i < customerList.length; i++) {
    const customer = customerList[i];
    if (normalizeForMatch_(customer.name) === normalizedSubject) {
      return customer;
    }
  }
  return null;
}

/**
 * 表示用の訪問先名を返す。
 * @param {Object} schedule
 * @returns {string}
 */
function getScheduleDisplayName_(schedule) {
  if (schedule.customerInfo && schedule.customerInfo.name) {
    return schedule.customerInfo.name;
  }
  if (schedule.subjectName) {
    return schedule.subjectName;
  }
  return schedule.location || schedule.title || '';
}

/**
 * タイトルの表記ゆれを吸収する。
 * @param {string} title
 * @returns {string}
 */
function normalizeEventTitle_(title) {
  return String(title || '')
    .replace(/［/g, '[')
    .replace(/］/g, ']');
}

/**
 * イベントタイトルから同期用タグを除去する。
 * @param {string} title
 * @returns {string}
 */
function stripEventTags_(title) {
  return normalizeEventTitle_(title)
    .replace(/^【[^】]*】/g, '')
    .replace(/\[予約確定\]/g, '')
    .replace(/\[新規\]/g, '')
    .replace(/\[イベント\]/g, '')
    .replace(/\[事務\]/g, '')
    .replace(/　+/g, ' ')
    .trim();
}

/**
 * 説明文から予約詳細URLを抽出する。
 * @param {string} description
 * @returns {string}
 */
function extractReservationUrl_(description) {
  const match = String(description || '').match(/https?:\/\/[\w/:%#$&?()~.=+\-]+/);
  return match ? match[0] : '';
}

/**
 * 名前照合用に空白を除去する。
 * @param {string} value
 * @returns {string}
 */
function normalizeForMatch_(value) {
  return String(value || '').replace(/\s+/g, '');
}

/**
 * 住所から緯度経度を取得する。
 * @param {string} address
 * @returns {{lat: *, lng: *}}
 */
function getLatLngFromAddressForRouting_(address) {
  try {
    const response = Maps.newGeocoder().geocode(address);
    if (response.status === 'OK' && response.results.length > 0) {
      const result = response.results[0].geometry.location;
      return { lat: result.lat, lng: result.lng };
    }
  } catch (e) {
    console.warn(`[getLatLngFromAddressForRouting_] ジオコーディング失敗: ${address}`, e.message);
  }
  return { lat: '', lng: '' };
}

/**
 * 出発地/目的地の座標を解決する。
 * @param {Object} loc
 * @param {Date} targetDate
 * @returns {{lat: *, lng: *}}
 */
function resolveLocationForRouting_(loc, targetDate) {
  let finalAddress = loc.address || '';
  let finalLat = loc.lat || '';
  let finalLng = loc.lng || '';

  if (loc.address2 && loc.address2_start && loc.address2_end) {
    const startDate = new Date(loc.address2_start);
    const endDate = new Date(loc.address2_end);
    const checkDate = new Date(targetDate.getTime());
    checkDate.setHours(0, 0, 0, 0);
    startDate.setHours(0, 0, 0, 0);
    endDate.setHours(0, 0, 0, 0);

    if (checkDate >= startDate && checkDate <= endDate) {
      finalAddress = loc.address2;
      finalLat = '';
      finalLng = '';
    }
  }

  if ((!finalLat || !finalLng) && finalAddress) {
    const geo = getLatLngFromAddressForRouting_(finalAddress);
    finalLat = geo.lat;
    finalLng = geo.lng;
  }

  return { lat: finalLat, lng: finalLng };
}

/**
 * 2地点間の経路URL/時間/距離を取得する。
 * @param {Object} from
 * @param {Object} to
 * @param {Date} targetDate
 * @returns {{url: string, min: *, km: *}}
 */
function getRouteDetailsForRouting_(from, to, targetDate) {
  const origin = resolveLocationForRouting_(from, targetDate);
  const destination = resolveLocationForRouting_(to, targetDate);

  if (!origin.lat || !origin.lng || !destination.lat || !destination.lng) {
    return { url: '', min: '', km: '' };
  }

  try {
    const directions = Maps.newDirectionFinder()
      .setOrigin(origin.lat, origin.lng)
      .setDestination(destination.lat, destination.lng)
      .setMode(Maps.DirectionFinder.Mode.DRIVING)
      .getDirections();

    if (directions.routes && directions.routes.length > 0) {
      const leg = directions.routes[0].legs[0];
      return {
        url: `https://www.google.com/maps/dir/?api=1&origin=${origin.lat},${origin.lng}&destination=${destination.lat},${destination.lng}&travelmode=driving`,
        min: Math.round(leg.duration.value / 60),
        km: (leg.distance.value / 1000).toFixed(2)
      };
    }
  } catch (e) {
    console.warn('[getRouteDetailsForRouting_] ルート計算失敗:', e.message);
  }

  return { url: '', min: '', km: '' };
}

/**
 * 指定フォルダ内の最新CSVを取得して解析する。
 * @param {GoogleAppsScript.Drive.Folder} folder
 * @returns {Array}
 */
function getCustomerDataFromCsvForRouting_(folder) {
  const files = folder.getFiles();
  const targetFiles = [];

  while (files.hasNext()) {
    const file = files.next();
    const name = file.getName();
    if (name.indexOf('Kokyaku_') === 0 && /\.csv$/i.test(name)) {
      targetFiles.push(file);
    }
  }

  if (targetFiles.length === 0) {
    throw new Error('対象のCSVファイルが見つかりません。');
  }

  targetFiles.sort((a, b) => b.getLastUpdated() - a.getLastUpdated());
  const latestFile = targetFiles[0];
  const csvContent = latestFile.getBlob().getDataAsString('UTF-16');
  const values = Utilities.parseCsv(csvContent, '\t');

  if (values.length < 2) {
    return [];
  }

  const header = values.shift();
  const idx = getColumnIndicesForRouting_(header);

  return values.map(row => {
    const lastName = String(row[idx.lastName] || '');
    const firstName = String(row[idx.firstName] || '');
    let lat = row[idx.lat];
    let lng = row[idx.lng];
    const combinedLatLng = String(row[idx.lat_lng] || '');

    if (combinedLatLng && combinedLatLng.includes(',')) {
      const parts = combinedLatLng.split(',');
      lat = parseFloat(parts[0].trim());
      lng = parseFloat(parts[1].trim());
    }

    return {
      id: row[idx.customerId],
      name: `${lastName} ${firstName}`.trim(),
      address: row[idx.address] || '',
      parkingarea: row[idx.parkingArea] || '',
      lat: lat || '',
      lng: lng || '',
      address2: row[idx.address2] || '',
      address2_start: row[idx.address2_start] || '',
      address2_end: row[idx.address2_end] || ''
    };
  }).filter(item => item.id);
}

/**
 * 顧客CSVの列インデックスを返す。
 * @param {Array} header
 * @returns {Object}
 */
function getColumnIndicesForRouting_(header) {
  const find = names => header.findIndex(h => names.some(name => String(h || '').includes(name)));
  return {
    customerId: find(['顧客ID']),
    lastName: find(['姓', '名字']),
    firstName: find(['名', '名前']),
    address: find(['住所']),
    lat_lng: find(['緯度・経度', '緯度/経度', '緯度,経度']),
    lat: find(['緯度', 'lat']),
    lng: find(['経度', 'lng']),
    parkingArea: find(['駐車場']),
    address2: find(['住所2', '住所２']),
    address2_start: find(['住所2[適用開始日YYYY/MM/DD]']),
    address2_end: find(['住所2[適用終了日YYYY/MM/DD]'])
  };
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

/**
 * ===============================================
 * デバッグ関数: カレンダーイベント一覧表示
 * ===============================================
 */

/**
 * 指定日のカレンダーイベントをすべて列挙（デバッグ用）
 * タグの有無、スタッフマッチング結果を表示
 * 
 * @param {string} dateStr - 'YYYY-MM-DD' 形式
 */
function debugShowCalendarEventsForDate(dateStr) {
  try {
    const date = new Date(dateStr.replace(/-/g, '/'));
    const events = fetchCalendarEvents(date);
    
    console.log(`\n========== デバッグ: ${dateStr} のカレンダーイベント一覧 ==========`);
    console.log(`取得イベント数: ${events.length}`);
    
    if (events.length === 0) {
      console.log('イベントなし');
      return;
    }

    const staffMap = getAuthorizedStaffMapForSync();
    
    for (let i = 0; i < events.length; i++) {
      const evt = events[i];
      const classified = classifyEvent(evt);
      
      console.log(`\n【イベント ${i + 1}】`);
      console.log('  タイトル: ' + evt.title);
      console.log('  時間: ' + Utilities.formatDate(evt.startTime, 'Asia/Tokyo', 'HH:mm') + 
                  ' ～ ' + Utilities.formatDate(evt.endTime, 'Asia/Tokyo', 'HH:mm'));
      console.log('  場所: ' + evt.location);
      console.log('  説明: ' + (evt.description || 'なし').substring(0, 100));
      console.log('  カレンダー所有者: ' + evt.ownerName);
      
      if (classified) {
        console.log('  ✅ タグあり: ' + classified.tag + ' (種別: ' + classified.type + ')');
      } else {
        console.log('  ❌ タグなし（フィルタされます）');
      }
      
      // ===== スタッフマッチング結果（改良版） =====
      const staffNamesFromDesc = extractStaffNamesFromDescription(
        evt.description || '', 
        evt.title, 
        evt, 
        staffMap
      );
      
      console.log('  【スタッフ抽出結果】');
      if (staffNamesFromDesc.length > 0) {
        console.log('    説明文/タイトル: ' + staffNamesFromDesc.join(', '));
      } else {
        console.log('    説明文/タイトル: なし');
      }
      
      // ゲスト情報
      if (evt.guests && evt.guests.length > 0) {
        console.log('  ゲスト:');
        for (let j = 0; j < evt.guests.length; j++) {
          console.log('    - ' + evt.guests[j].getEmail());
        }
        
        // ゲストからのスタッフマッチング結果
        const staffNamesFromGuests = extractStaffFromGuests(evt.guests, staffMap);
        if (staffNamesFromGuests.length > 0) {
          console.log('    ゲスト: ' + staffNamesFromGuests.join(', '));
        } else {
          console.log('    ゲスト: マッチなし');
        }
      } else {
        console.log('  ゲスト: なし');
      }
    }
    
    console.log(`\n========== デバッグ終了 ==========\n`);
    
  } catch (e) {
    console.error('[debugShowCalendarEventsForDate] error:', e.message);
  }
}

/**
 * 指定期間のカレンダーイベント統計（デバッグ用）
 * 
 * @param {string} startDateStr - 'YYYY-MM-DD'
 * @param {string} endDateStr - 'YYYY-MM-DD'
 */
function debugShowCalendarEventStats(startDateStr, endDateStr) {
  try {
    const startDate = new Date(startDateStr.replace(/-/g, '/'));
    const endDate = new Date(endDateStr.replace(/-/g, '/'));
    
    console.log(`\n========== デバッグ統計: ${startDateStr} ～ ${endDateStr} ==========`);
    
    let totalEvents = 0;
    let taggedEvents = 0;
    let untaggedEvents = 0;
    const titlePatterns = {};
    
    const staffMap = getAuthorizedStaffMapForSync();
    
    for (let date = new Date(startDate); date <= endDate; date.setDate(date.getDate() + 1)) {
      const events = fetchCalendarEvents(new Date(date));
      
      for (let i = 0; i < events.length; i++) {
        const evt = events[i];
        const classified = classifyEvent(evt);
        totalEvents++;
        
        if (classified) {
          taggedEvents++;
        } else {
          untaggedEvents++;
        }
        
        // タイトルパターン集計
        const titleWithoutTag = evt.title.replace(/\[.*?\]/g, '').trim();
        titlePatterns[titleWithoutTag] = (titlePatterns[titleWithoutTag] || 0) + 1;
      }
    }
    
    console.log(`総イベント数: ${totalEvents}`);
    console.log(`タグあり: ${taggedEvents}`);
    console.log(`タグなし（スキップ）: ${untaggedEvents}`);
    console.log(`スキップ率: ${(untaggedEvents / totalEvents * 100).toFixed(1)}%`);
    
    console.log('\nタイトルパターン (トップ10):');
    const sorted = Object.entries(titlePatterns)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 10);
    
    for (let i = 0; i < sorted.length; i++) {
      console.log(`  ${sorted[i][0]}: ${sorted[i][1]}件`);
    }
    
    console.log(`\n========== デバッグ統計終了 ==========\n`);
    
  } catch (e) {
    console.error('[debugShowCalendarEventStats] error:', e.message);
  }
}
