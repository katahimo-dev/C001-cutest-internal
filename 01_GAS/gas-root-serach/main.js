// ==========================================
// 設定・定数（変更があれば修正してください）
// ==========================================
const CONFIG = {
  CUSTOMER_FOLDER_ID: "1wLjR6iZ447tbUa3ff59bejoM5aXh8clC",
  STAFF_SS_ID: "1exqD69qZqACm9KOUPpa0fVWRYD2qEZfce7I6TOs_VDk",
  ROUTE_SUMMARY_FILE_NAME: "ルート集計", // 保存用ファイル名
  ATTENDANCE_FILE_NAME: "勤怠集計", // 勤怠転記用ファイル名
  ATTENDANCE_FOLDER_ID: "1FA2aSBddgBakETEbzJhJIx1vWG06P46J", // 勤怠集計保存先（gas-childcare-daily-report と共通）
  REPORT_APP_URL: "https://script.google.com/macros/s/AKfycbwmef1CgKJNhqjJXr-ZqkRWx4_-OJ8pjzdzBdRC8iOZ8_v5TXrIEWv01xFlCKGaurs/exec", // 日報アプリURL
  // LINE WORKS設定
  LW: {
    CLIENT_ID: PropertiesService.getScriptProperties().getProperty('LW_CLIENT_ID'),
    CLIENT_SECRET: PropertiesService.getScriptProperties().getProperty('LW_CLIENT_SECRET'),
    SERVICE_ACCOUNT: PropertiesService.getScriptProperties().getProperty('LW_SERVICE_ACCOUNT'),
    PRIVATE_KEY: PropertiesService.getScriptProperties().getProperty('LW_PRIVATE_KEY'),
    BOT_ID: PropertiesService.getScriptProperties().getProperty('LW_BOT_ID')
  }
};

/**
 * 【指定日実行用】特定の日付で実行したい時は以下の手順で実行する
 * 1. main("")の中の日付を指定日に書き換える
 * 2. 「デバッグ」右のトグルから「runSpecifiedDate」を選択
 * 3. 「実行」をクリック
 */
function runSpecifiedDate() {
  main("2026/4/24"); // ここを書き換えるだけでOK
}

/**
 * メイン実行関数
 * @param {string} date - "2025/12/25" 形式の文字列（指定日実行用）
 */
function main(date) {
  let targetDate;

  if (date && typeof date === 'string') {
    targetDate = new Date(date);
    console.log(`--- 手動実行モード：指定日 ${date} を処理します ---`);
  } else {
    targetDate = new Date();
    targetDate.setDate(targetDate.getDate() + 1);
    console.log("--- 自動実行モード：翌日の予定を処理します ---");
  }

  if (isNaN(targetDate.getTime())) {
    console.error("エラー: 無効な日付が指定されました。");
    return;
  }

  console.log("実行対象日:", Utilities.formatDate(targetDate, "JST", "yyyy-MM-dd"));

  const dailyData = buildDailyRouteRows(targetDate);
  if (!dailyData.hasEvents) {
    console.log("予定がありません。終了します。");
    return;
  }

  outputToSheetAppend(dailyData.outputRows, dailyData.sheetName);
  sendDailyScheduleToLineWorks(targetDate);
}

/**
 * 勤怠集計用データを保存する（通知なし）
 */
function saveAttendanceData(targetDate, preloadedCustomerData, preloadedStaffData) {
  if (!(targetDate instanceof Date) || isNaN(targetDate.getTime())) {
    throw new Error("saveAttendanceData: 無効な日付です。");
  }

  const dailyData = buildDailyRouteRows(targetDate, preloadedCustomerData, preloadedStaffData);
  console.log(`勤怠集計作成: ${dailyData.dateStr} / ${dailyData.sheetName}`);

  if (!dailyData.hasEvents) {
    outputToAttendanceFile([], dailyData.sheetName, dailyData.dateStr);
    console.log("予定がありません。勤怠集計の当日データ削除のみ実施しました。");
    return;
  }

  outputToAttendanceFile(dailyData.outputRows, dailyData.sheetName, dailyData.dateStr);
}

/**
 * 1日分のルート計算用データを作成する
 */
function buildDailyRouteRows(targetDate, preloadedCustomerData, preloadedStaffData) {
  const y = targetDate.getFullYear();
  const m = ('0' + (targetDate.getMonth() + 1)).slice(-2);
  const d = ('0' + targetDate.getDate()).slice(-2);
  const dateStr = `${y}-${m}-${d}`;
  const sheetName = `${y}${m}`;

  console.log(`処理対象日: ${dateStr} / シート名: ${sheetName}`);

  const folder = DriveApp.getFolderById(CONFIG.CUSTOMER_FOLDER_ID);
  const customerData = preloadedCustomerData || getCustomerDataFromCsv(folder);
  const staffData = preloadedStaffData || getStaffDataFromSpreadsheet(CONFIG.STAFF_SS_ID);
  const calendarEvents = getCalendarEvents(targetDate, customerData);

  if (calendarEvents.length === 0) {
    return { dateStr, sheetName, hasEvents: false, outputRows: [] };
  }

  const groupedEvents = groupEventsByStaff(calendarEvents, staffData);
  const outputRows = calculateDetailedRoutes(groupedEvents, dateStr, targetDate);

  return { dateStr, sheetName, hasEvents: true, outputRows };
}

function autoRunSaveAttendance() {
  const targetDate = new Date();
  targetDate.setHours(0, 0, 0, 0);
  saveAttendanceData(targetDate);
}

function debugSaveAttendanceMarch2026() {
  const start = new Date("2026/03/01");
  const end = new Date("2026/03/31");
  saveAttendanceDataInRange(start, end);
}

function debugSaveAttendanceApril2026() {
  saveAttendanceDataInRange(new Date("2026/04/01"), new Date("2026/04/18"));
}

function saveAttendanceDataInRange(startDate, endDate) {
  const folder = DriveApp.getFolderById(CONFIG.CUSTOMER_FOLDER_ID);
  const customerData = getCustomerDataFromCsv(folder);
  const staffData = getStaffDataFromSpreadsheet(CONFIG.STAFF_SS_ID);
  console.log(`データ読み込み完了。期間: ${Utilities.formatDate(startDate, 'JST', 'yyyy-MM-dd')} ～ ${Utilities.formatDate(endDate, 'JST', 'yyyy-MM-dd')}`);

  const cursor = new Date(startDate.getTime());
  while (cursor <= endDate) {
    const current = new Date(cursor.getTime());
    saveAttendanceData(current, customerData, staffData);
    cursor.setDate(cursor.getDate() + 1);
  }
}

// ==========================================
// 1. カレンダー情報取得
// ==========================================
function getCalendarEvents(date, customerList) {
  const calendars = CalendarApp.getAllCalendars();
  const uniqueEventsMap = new Map();

  const startTime = new Date(date);
  startTime.setHours(0, 0, 0, 0);
  const endTime = new Date(startTime);
  endTime.setHours(23, 59, 59, 999);

  calendars.forEach(calendar => {
    const calendarName = calendar.getName();
    const events = calendar.getEvents(startTime, endTime);

    events.forEach(event => {
      if (event.getMyStatus() === CalendarApp.GuestStatus.NO) {
        return;
      }
      const eventId = event.getId();
      if (!uniqueEventsMap.has(eventId)) {
        uniqueEventsMap.set(eventId, {
          event: event,
          ownerName: calendarName
        });
      }
    });
  });

  const allEventData = Array.from(uniqueEventsMap.values());
  const processedEvents = [];

  allEventData.forEach(data => {
    const event = data.event;
    const ownerName = data.ownerName;

    const subject = event.getTitle();
    const description = event.getDescription() || "";
    const location = event.getLocation();

    const isConfirmed = subject.includes("[予約確定]");
    const isNewCustomer = subject.includes("[新規]");
    const isSpecialEvent = subject.includes("[イベント]");
    const isOfficeWork = subject.includes("[事務]");

    const cleanedSubjectName = subject
          .replace("[予約確定]", "")
          .replace("[新規]", "")
          .replace("[イベント]", "")
          .replace("[事務]", "")
          .trim();

    if (isConfirmed) {
      const staffMatch = description.match(/施設：(.*?)\[/);
      const staffNameFromDesc = staffMatch ? staffMatch[1].trim() : ownerName;

      const isOnline = description.includes("オンライン");

      let customer = customerList.find(c => {
        const targetName = String(c.name || "").replace(/\s+/g, "");
        const searchName = cleanedSubjectName.replace(/\s+/g, "");
        return targetName === searchName;
      });

      if (!customer) {
        customer = {
          name: cleanedSubjectName,
          address: "",
          lat: "",
          lng: "",
          parkingarea: ""
        };
      }

      if (isOnline) {
        customer.address = "";
        customer.lat = "";
        customer.lng = "";
      }

      const urlMatch = description.match(/https?:\/\/[\w/:%#\$&\?\(\)~\.=\+\-]+/);

      processedEvents.push({
        startTime: event.getStartTime(),
        endTime: event.getEndTime(),
        eventType: "CUSTOMER APPOINTMENT",
        customerInfo: customer,
        staffNameRaw: staffNameFromDesc,
        reservaUrl: urlMatch ? urlMatch[0] : "",
        isMeeting: false
      });
    }
    else if (isNewCustomer) {
      let geo = location ? getLatLngFromAddress(location) : { lat: "", lng: "" };

      processedEvents.push({
        startTime: event.getStartTime(),
        endTime: event.getEndTime(),
        eventType: "CUSTOMER APPOINTMENT",
        customerInfo: {
          name: cleanedSubjectName,
          address: location,
          lat: geo.lat,
          lng: geo.lng,
          parkingarea: ""
        },
        staffNameRaw: ownerName,
        reservaUrl: "",
        isMeeting: false
      });
    }
    else if (isSpecialEvent) {
      let geo = location ? getLatLngFromAddress(location) : { lat: "", lng: "" };
      const guestNames = event.getGuestList().map(guest => guest.getName() || guest.getEmail());

      processedEvents.push({
        startTime: event.getStartTime(),
        endTime: event.getEndTime(),
        eventType: "EVENT",
        customerInfo: {
          name: cleanedSubjectName,
          address: location,
          lat: geo.lat,
          lng: geo.lng,
          parkingarea: ""
        },
        guestNames: guestNames,
        ownerName: ownerName,
        reservaUrl: "",
        isMeeting: true
      });
    }
    else if (isOfficeWork) {
      let geo = location ? getLatLngFromAddress(location) : { lat: "", lng: "" };

      processedEvents.push({
        startTime: event.getStartTime(),
        endTime: event.getEndTime(),
        eventType: "OFFICE WORK",
        customerInfo: {
          name: cleanedSubjectName,
          address: location || "",
          lat: geo.lat,
          lng: geo.lng,
          parkingarea: ""
        },
        staffNameRaw: ownerName,
        reservaUrl: "",
        isMeeting: false
      });
    }
  });

  return mergeOverlappingOfficeWork(processedEvents);
}

function mergeOverlappingOfficeWork(events) {
  const officeWorks = events.filter(e => e.eventType === "OFFICE WORK");
  const nonOfficeWorks = events.filter(e => e.eventType !== "OFFICE WORK");

  if (officeWorks.length === 0) {
    return events;
  }

  const staffGroups = {};
  officeWorks.forEach(event => {
    const staffName = event.staffNameRaw || "";
    if (!staffGroups[staffName]) {
      staffGroups[staffName] = [];
    }
    staffGroups[staffName].push(event);
  });

  const mergedEvents = [];
  Object.keys(staffGroups).forEach(staffName => {
    const staffEvents = staffGroups[staffName];
    staffEvents.sort((a, b) => a.startTime - b.startTime);

    const processed = [];
    let currentMergeGroup = null;

    staffEvents.forEach(event => {
      if (!currentMergeGroup) {
        currentMergeGroup = {
          names: [event.customerInfo.name],
          startTime: event.startTime,
          endTime: event.endTime,
          baseEvent: event
        };
      } else if (event.startTime < currentMergeGroup.endTime) {
        currentMergeGroup.names.push(event.customerInfo.name);
        currentMergeGroup.endTime = new Date(Math.max(currentMergeGroup.endTime.getTime(), event.endTime.getTime()));
      } else {
        processed.push(currentMergeGroup);
        currentMergeGroup = {
          names: [event.customerInfo.name],
          startTime: event.startTime,
          endTime: event.endTime,
          baseEvent: event
        };
      }
    });

    if (currentMergeGroup) {
      processed.push(currentMergeGroup);
    }

    processed.forEach(group => {
      mergedEvents.push({
        startTime: group.startTime,
        endTime: group.endTime,
        eventType: "OFFICE WORK",
        customerInfo: {
          name: group.names.join(","),
          address: group.baseEvent.customerInfo.address,
          lat: group.baseEvent.customerInfo.lat,
          lng: group.baseEvent.customerInfo.lng,
          parkingarea: group.baseEvent.customerInfo.parkingarea
        },
        staffNameRaw: staffName,
        reservaUrl: "",
        isMeeting: false
      });
    });
  });

  return [...nonOfficeWorks, ...mergedEvents];
}

function getLatLngFromAddress(address) {
  try {
    const response = Maps.newGeocoder().geocode(address);
    if (response.status === 'OK' && response.results.length > 0) {
      const result = response.results[0].geometry.location;
      return { lat: result.lat, lng: result.lng };
    }
  } catch (e) {
    console.warn(`ジオコーディング失敗: ${address}`, e);
  }
  return { lat: "", lng: "" };
}

// ==========================================
// 2. スタッフごとのグループ化（招待ゲスト対応）
// ==========================================
function groupEventsByStaff(events, staffList) {
  const staffGroups = {};
  const normalize = (str) => str ? str.replace(/\s+/g, "") : "";

  events.forEach(event => {
    if (event.isMeeting) {
      if (event.guestNames && event.guestNames.length > 0) {
        event.guestNames.forEach(guestName => {
          const staff = staffList.find(s => normalize(s.name) === normalize(guestName));
          if (staff) {
            addEventToGroup(staffGroups, staff, event);
          }
        });
      } else if (event.ownerName) {
        const staff = staffList.find(s => normalize(s.name) === normalize(event.ownerName));
        if (staff) {
          addEventToGroup(staffGroups, staff, event);
        }
      }
    } else {
      const staff = staffList.find(s => normalize(s.name) === normalize(event.staffNameRaw));
      if (staff) {
        addEventToGroup(staffGroups, staff, event);
      }
    }
  });

  Object.keys(staffGroups).forEach(key => {
    staffGroups[key].appointments.sort((a, b) => a.startTime - b.startTime);
  });
  return staffGroups;
}

function addEventToGroup(groups, staff, event) {
  if (!groups[staff.id]) {
    groups[staff.id] = { staffInfo: staff, appointments: [] };
  }
  groups[staff.id].appointments.push(event);
}

// ==========================================
// 3. ルート計算
// ==========================================
function calculateDetailedRoutes(groupedEvents, dateStr, targetDate) {
  const allRows = [];
  Object.values(groupedEvents).forEach(group => {
    const staff = group.staffInfo;
    const apps = group.appointments;

    const validLocApps = apps.filter(app =>
      app.customerInfo && (
        (app.customerInfo.lat && app.customerInfo.lng) ||
        (app.customerInfo.address && app.customerInfo.address.trim() !== "")
      )
    );

    for (let i = 0; i < apps.length; i++) {
      const currentApp = apps[i];
      const hasLocation = currentApp.customerInfo && (
        (currentApp.customerInfo.lat && currentApp.customerInfo.lng) ||
        (currentApp.customerInfo.address && currentApp.customerInfo.address.trim() !== "")
      );

      let moveInfo = { url: "", min: "", km: "" };
      let attendanceInfo = { url: "", min: "", km: "" };
      let leavingInfo = { url: "", min: "", km: "" };

      if (hasLocation) {
        if (currentApp === validLocApps[0]) {
          attendanceInfo = getRouteDetails(staff, currentApp.customerInfo, targetDate);
        }

        const myIndexInValid = validLocApps.indexOf(currentApp);
        if (myIndexInValid > 0) {
          const prevValidApp = validLocApps[myIndexInValid - 1];
          moveInfo = getRouteDetails(prevValidApp.customerInfo, currentApp.customerInfo, targetDate);
        }

        if (currentApp === validLocApps[validLocApps.length - 1]) {
          leavingInfo = getRouteDetails(currentApp.customerInfo, staff, targetDate);
        }
      } else {
        let moveFromLocation = null;
        for (let j = i - 1; j >= 0; j--) {
          const prevApp = apps[j];
          const prevHasLocation = prevApp.customerInfo && (
            (prevApp.customerInfo.lat && prevApp.customerInfo.lng) ||
            (prevApp.customerInfo.address && prevApp.customerInfo.address.trim() !== "")
          );
          if (prevHasLocation) {
            moveFromLocation = prevApp.customerInfo;
            break;
          }
        }

        if (moveFromLocation && currentApp.customerInfo && currentApp.customerInfo.address) {
          moveInfo = getRouteDetails(moveFromLocation, currentApp.customerInfo, targetDate);
        }
      }

      // 顧客ID を末尾に追加 (RESERVA 顧客ID、CRM と同じ形式)
      allRows.push([
        dateStr,
        staff.name,
        currentApp.eventType,
        currentApp.customerInfo.name,
        Utilities.formatDate(currentApp.startTime, "JST", "HH:mm"),
        Utilities.formatDate(currentApp.endTime, "JST", "HH:mm"),
        currentApp.reservaUrl,
        moveInfo.url, moveInfo.min, moveInfo.km,
        attendanceInfo.url, attendanceInfo.min, attendanceInfo.km,
        leavingInfo.url, leavingInfo.min, leavingInfo.km,
        currentApp.customerInfo.id || ''   // ← 追加: RESERVA 顧客ID
      ]);
    }
  });
  return allRows;
}


function getRouteDetails(from, to, targetDate) {
  const origin = resolveLocation(from, targetDate);
  const destination = resolveLocation(to, targetDate);

  if (!origin.lat || !origin.lng || !destination.lat || !destination.lng) {
    return { url: "", min: "", km: "" };
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
        km: (leg.distance.value / 1000).toFixed(2),
        min: Math.round(leg.duration.value / 60)
      };
    }
  } catch (e) {
    console.warn("ルート計算失敗", e);
  }
  return { url: "", min: "", km: "" };
}

function resolveLocation(loc, targetDate) {
  let finalAddress = loc.address;
  let finalLat = loc.lat;
  let finalLng = loc.lng;

  if (loc.address2 && loc.address2_start && loc.address2_end) {
    const startDate = new Date(loc.address2_start);
    const endDate = new Date(loc.address2_end);

    const checkDate = new Date(targetDate.getTime());
    checkDate.setHours(0, 0, 0, 0);
    startDate.setHours(0, 0, 0, 0);
    endDate.setHours(0, 0, 0, 0);

    if (checkDate >= startDate && checkDate <= endDate) {
      finalAddress = loc.address2;
      finalLat = "";
      finalLng = "";
    }
  }

  if (!finalLat || !finalLng) {
    if (finalAddress) {
      const geo = getLatLngFromAddress(finalAddress);
      finalLat = geo.lat;
      finalLng = geo.lng;
    }
  }

  return { lat: finalLat, lng: finalLng };
}


function getCustomerDataFromCsv(folder) {
  const files = folder.getFiles();
  const targetFiles = [];

  while (files.hasNext()) {
    const file = files.next();
    const name = file.getName();
    if (name.startsWith("Kokyaku_") && name.endsWith(".csv")) {
      targetFiles.push(file);
    }
  }

  if (targetFiles.length === 0) {
    throw new Error("対象のCSVファイルが見つかりません。");
  }

  targetFiles.sort((a, b) => b.getLastUpdated() - a.getLastUpdated());
  const latestFile = targetFiles[0];
  console.log(`読み込み対象ファイル: ${latestFile.getName()} (更新日: ${latestFile.getLastUpdated()})`);

  const csvContent = latestFile.getBlob().getDataAsString("UTF-16");
  const values = Utilities.parseCsv(csvContent, '\t');

  if (values.length < 2) return [];

  const header = values.shift();
  const idx = getColumnIndices(header);

  return values.map(row => {
    const lastName = String(row[idx.lastName] || "");
    const firstName = String(row[idx.firstName] || "");

    let lat = row[idx.lat];
    let lng = row[idx.lng];

    const combinedLatLng = String(row[idx.lat_lng] || "");
    if (combinedLatLng && combinedLatLng.includes(",")) {
      const parts = combinedLatLng.split(",");
      lat = parseFloat(parts[0].trim());
      lng = parseFloat(parts[1].trim());
    }

    return {
      id: row[idx.customerId],
      name: `${lastName} ${firstName}`.trim(),
      address: row[idx.address],
      parkingarea: row[idx.parkingArea],
      lat: lat,
      lng: lng,
      address2: row[idx.address2],
      address2_start: row[idx.address2_start],
      address2_end: row[idx.address2_end],
    };
  }).filter(item => item.id);
}


function getStaffDataFromSpreadsheet(spreadsheetId) {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheets()[0];

  const data = sheet.getDataRange().getValues();

  if (data.length < 2) throw new Error("スプレッドシートにデータが見つかりません。");

  const header = data.shift();
  const idx = getColumnIndices(header);

  return data.map(row => {
    const lastName = String(row[idx.lastName] || "");
    const firstName = String(row[idx.firstName] || "");

    const id = extractIdFromName(lastName) || extractIdFromName(firstName) || row[0];

    let lat = row[idx.lat];
    let lng = row[idx.lng];

    const combinedLatLng = String(row[idx.lat_lng] || "");
    if (combinedLatLng && combinedLatLng.includes(",")) {
      const parts = combinedLatLng.split(",");
      lat = parseFloat(parts[0].trim());
      lng = parseFloat(parts[1].trim());
    }

    return {
      id: id,
      name: `${lastName} ${firstName}`.trim() || row[idx.name] || row[1],
      address: row[idx.address],
      lat: lat,
      lng: lng,
      lwId: row[idx.lwId]
    };
  });
}


function getColumnIndices(header) {
  const find = (names) => header.findIndex(h => names.some(name => h.includes(name)));
  return {
    customerId: find(["顧客ID"]),
    lastName: find(["姓", "名字"]),
    firstName: find(["名", "名前"]),
    address: find(["住所"]),
    lat_lng: find(['緯度・経度', '緯度/経度', '緯度,経度']),
    lat: find(["緯度", "lat"]),
    lng: find(["経度", "lng"]),
    parkingArea: find(["駐車場"]),
    lwId: find(["LW_ID"]),
    name: find(["氏名", "スタッフ名"]),
    address2: find(['住所2', '住所2']),
    address2_start: find(['住所2[適用開始日YYYY/MM/DD]']),
    address2_end: find(['住所2[適用終了日YYYY/MM/DD]'])
  };
}

function extractIdFromName(str) {
  if (!str) return null;
  const match = str.match(/\[(.*?)\]/);
  return match ? match[1].trim() : null;
}

function outputToSheetAppend(rows, sheetName) {
  const ss = getOrCreateSpreadsheetInSameFolder(CONFIG.ROUTE_SUMMARY_FILE_NAME);
  let sheet = ss.getSheetByName(sheetName);
  // ヘッダーに「顧客ID」を末尾追加 (2026-04-23 CRM連携対応)
  const header = [
    "日付", "スタッフ名", "種別", "顧客名", "開始時間", "終了時間", "予約詳細URL",
    "移動経路URL", "移動時間（分）", "移動距離（km）",
    "出勤経路URL", "出勤経路時間（分）", "出勤経路距離（km）",
    "退勤経路URL", "退勤経路時間（分）", "退勤経路距離（km）",
    "顧客ID"
  ];
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(header);
  }
  if (rows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, header.length).setValues(rows);
  }
}

function outputToAttendanceFile(rows, sheetName, dateStr) {
  const ss = getOrCreateSpreadsheetInFolder(CONFIG.ATTENDANCE_FOLDER_ID, CONFIG.ATTENDANCE_FILE_NAME);
  let sheet = ss.getSheetByName(sheetName);
  // ヘッダーに「顧客ID」を末尾追加
  const header = [
    "日付", "スタッフ名", "種別", "顧客名", "開始時間", "終了時間", "予約詳細URL",
    "移動経路URL", "移動時間（分）", "移動距離（km）",
    "出勤経路URL", "出勤経路時間（分）", "出勤経路距離（km）",
    "退勤経路URL", "退勤経路時間（分）", "退勤経路距離（km）",
    "顧客ID"
  ];

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(header);
  }

  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const dateValues = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
    const dateSlash = dateStr.replace(/-/g, '/');

    for (let i = dateValues.length - 1; i >= 0; i--) {
      const rowDate = (dateValues[i][0] || "").toString();
      if (rowDate.indexOf(dateStr) !== -1 || rowDate.indexOf(dateSlash) !== -1) {
        sheet.deleteRow(i + 2);
      }
    }
  }

  if (rows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, header.length).setValues(rows);
  }
}

function getOrCreateSpreadsheetInSameFolder(fileName) {
  const files = DriveApp.getFilesByName(fileName);
  if (files.hasNext()) return SpreadsheetApp.open(files.next());
  return SpreadsheetApp.create(fileName);
}

function getOrCreateSpreadsheetInFolder(folderId, fileName) {
  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFilesByName(fileName);
  if (files.hasNext()) {
    return SpreadsheetApp.open(files.next());
  }

  const ss = SpreadsheetApp.create(fileName);
  const file = DriveApp.getFileById(ss.getId());
  folder.addFile(file);

  try {
    DriveApp.getRootFolder().removeFile(file);
  } catch (e) {
    console.warn("作成ファイルのマイドライブ除去をスキップしました: " + e.message);
  }

  return ss;
}


function formatTimeValue(val) {
  if (val instanceof Date) {
    return Utilities.formatDate(val, "JST", "HH:mm");
  }
  if (typeof val === 'string' && val.includes(':')) {
    return val.substring(0, 5);
  }
  return val;
}

function debugDumpAllCalendarsForDate(dateStr) {
  const targetDate = new Date(dateStr);
  if (isNaN(targetDate.getTime())) {
    console.error("無効な日付です:", dateStr);
    return;
  }
  const startTime = new Date(targetDate);
  startTime.setHours(0,0,0,0);
  const endTime = new Date(startTime);
  endTime.setHours(23,59,59,999);

  const calendars = CalendarApp.getAllCalendars();
  console.log(`--- カレンダーダンプ ${dateStr} 開始 ---`);
  calendars.forEach(cal => {
    const cname = cal.getName();
    const events = cal.getEvents(startTime, endTime);
    console.log(`Calendar: ${cname} / id: ${cal.getId()} / events: ${events.length}`);
    events.forEach(ev => {
      const sid = ev.getId();
      const title = ev.getTitle();
      const desc = ev.getDescription() || "";
      const loc = ev.getLocation() || "";
      const startS = Utilities.formatDate(ev.getStartTime(), "JST", "yyyy-MM-dd HH:mm:ss");
      const endS = Utilities.formatDate(ev.getEndTime(), "JST", "yyyy-MM-dd HH:mm:ss");
      const guests = ev.getGuestList().map(g => g.getName() || g.getEmail()).join(", ");
      console.log(`  --- Event ---`);
      console.log(`  id: ${sid}`);
      console.log(`  title: ${title}`);
      console.log(`  start: ${startS}`);
      console.log(`  end: ${endS}`);
      console.log(`  location: ${loc}`);
      console.log(`  description: ${desc}`);
      console.log(`  guests: ${guests}`);
      console.log(`  myStatus: ${ev.getMyStatus()}`);
    });
  });
  console.log(`--- カレンダーダンプ ${dateStr} 終了 ---`);
}

function runDebug_20260313() {
  debugDumpAllCalendarsForDate("2026/03/13");
}

function calcDurationMinForAttendance_(startTime, endTime) {
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

function isOfficeWorkAppointment_(appointment) {
  if (appointment.eventType === 'OFFICE WORK') return true;

  if (appointment.eventType === 'CUSTOMER APPOINTMENT') {
    return calcDurationMinForAttendance_(appointment.startTime, appointment.endTime) === 15;
  }

  return false;
}

function buildTimesheetRowDataFromAppointments_(appointments) {
  const rowData = {
    C: '', D: '', E: '',
    H: '',
    L: '', M: '', N: '',
    Q: '',
    U: '', V: '', W: '',
    X: '', Y: '', Z: '',
    AA: '', AB: '', AC: '',
    AG: '', AH: '', AI: '', AJ: ''
  };

  const officeWorks = appointments.filter(isOfficeWorkAppointment_);
  const visits = appointments.filter(app => !isOfficeWorkAppointment_(app));

  if (visits[0]) {
    rowData.C = visits[0].customerName || '';
    rowData.D = visits[0].startTime || '';
    rowData.E = visits[0].endTime || '';
    rowData.AI = visits[0].attendanceKm || '';
  }

  if (visits[1]) {
    rowData.L = visits[1].customerName || '';
    rowData.M = visits[1].startTime || '';
    rowData.N = visits[1].endTime || '';
    rowData.H = visits[1].moveMin || '';
    rowData.AG = visits[1].moveKm || '';
  }

  if (visits[2]) {
    rowData.U = visits[2].customerName || '';
    rowData.V = visits[2].startTime || '';
    rowData.W = visits[2].endTime || '';
    rowData.Q = visits[2].moveMin || '';
    rowData.AH = visits[2].moveKm || '';
  }

  if (visits.length > 0) {
    rowData.AJ = visits[visits.length - 1].leavingKm || '';
  }

  if (officeWorks[0]) {
    rowData.X = officeWorks[0].customerName || '';
    rowData.Y = officeWorks[0].startTime || '';
    rowData.Z = officeWorks[0].endTime || '';
  }

  if (officeWorks[1]) {
    rowData.AA = officeWorks[1].customerName || '';
    rowData.AB = officeWorks[1].startTime || '';
    rowData.AC = officeWorks[1].endTime || '';
  }

  return rowData;
}

function refreshAttendanceForStaffOnDate(staffName, dateString) {
  const normalizedStaff = String(staffName || "").trim();
  if (!normalizedStaff) throw new Error("staffName が指定されていません。");

  const targetDate = new Date(String(dateString || "").replace(/-/g, "/"));
  if (isNaN(targetDate.getTime())) {
    throw new Error("dateString が不正です。YYYY-MM-DD 形式で指定してください。");
  }
  targetDate.setHours(0, 0, 0, 0);

  const y = targetDate.getFullYear();
  const m = ("0" + (targetDate.getMonth() + 1)).slice(-2);
  const d = ("0" + targetDate.getDate()).slice(-2);
  const dateStr = `${y}-${m}-${d}`;
  const sheetName = `${y}${m}`;

  const folder = DriveApp.getFolderById(CONFIG.CUSTOMER_FOLDER_ID);
  const customerData = getCustomerDataFromCsv(folder);
  const staffData = getStaffDataFromSpreadsheet(CONFIG.STAFF_SS_ID);
  const calendarEvents = getCalendarEvents(targetDate, customerData);
  const groupedEvents = groupEventsByStaff(calendarEvents, staffData);

  const normalize = (str) => String(str || "").replace(/\s+/g, "");
  const staff = staffData.find(s => normalize(s.name) === normalize(normalizedStaff));

  let outputRows = [];
  if (staff && groupedEvents[staff.id]) {
    const singleGroup = {};
    singleGroup[staff.id] = groupedEvents[staff.id];
    outputRows = calculateDetailedRoutes(singleGroup, dateStr, targetDate);
  }

  const ss = getOrCreateSpreadsheetInFolder(CONFIG.ATTENDANCE_FOLDER_ID, CONFIG.ATTENDANCE_FILE_NAME);
  let sheet = ss.getSheetByName(sheetName);
  const header = [
    "日付", "スタッフ名", "種別", "顧客名", "開始時間", "終了時間", "予約詳細URL",
    "移動経路URL", "移動時間（分）", "移動距離（km）",
    "出勤経路URL", "出勤経路時間（分）", "出勤経路距離（km）",
    "退勤経路URL", "退勤経路時間（分）", "退勤経路距離（km）",
    "顧客ID"
  ];
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(header);
  } else {
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const dateStaffVals = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
      const targetHyphen = dateStr;
      const targetSlash = dateStr.replace(/-/g, "/");
      for (let i = dateStaffVals.length - 1; i >= 0; i--) {
        const rowDate = String(dateStaffVals[i][0] || "").trim();
        const rowStaff = normalize(String(dateStaffVals[i][1] || ""));
        if (rowStaff === normalize(normalizedStaff) &&
            (rowDate.indexOf(targetHyphen) !== -1 || rowDate.indexOf(targetSlash) !== -1)) {
          sheet.deleteRow(i + 2);
        }
      }
    }
  }

  if (outputRows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, outputRows.length, header.length).setValues(outputRows);
  }

  const appointments = outputRows.map(row => ({
    eventType:     row[2],
    customerName:  row[3],
    startTime:     String(row[4] || ""),
    endTime:       String(row[5] || ""),
    reservaUrl:    row[6],
    moveUrl:       row[7],
    moveMin:       row[8],
    moveKm:        row[9],
    attendanceUrl: row[10],
    attendanceMin: row[11],
    attendanceKm:  row[12],
    leavingUrl:    row[13],
    leavingMin:    row[14],
    leavingKm:     row[15],
    customerId:    row[16]
  }));

  const rowData = buildTimesheetRowDataFromAppointments_(appointments);

  return {
    success: true,
    date: dateStr,
    staffName: normalizedStaff,
    appointments: appointments,
    rowData: rowData
  };
}

function getScheduleForStaffOnDate(staffName, dateString) {
  const normalizedStaffName = String(staffName || "").trim();
  if (!normalizedStaffName) {
    throw new Error("staffName が指定されていません。");
  }

  const targetDate = new Date(String(dateString || "").replace(/-/g, "/"));
  if (isNaN(targetDate.getTime())) {
    throw new Error("dateString が不正です。YYYY-MM-DD 形式で指定してください。");
  }
  targetDate.setHours(0, 0, 0, 0);

  const folder = DriveApp.getFolderById(CONFIG.CUSTOMER_FOLDER_ID);
  const customerData = getCustomerDataFromCsv(folder);
  const staffData = getStaffDataFromSpreadsheet(CONFIG.STAFF_SS_ID);
  const calendarEvents = getCalendarEvents(targetDate, customerData);
  const groupedEvents = groupEventsByStaff(calendarEvents, staffData);

  const normalize = (str) => String(str || "").replace(/\s+/g, "");
  const staff = staffData.find(s => normalize(s.name) === normalize(normalizedStaffName));
  if (!staff) {
    return {
      success: true,
      date: Utilities.formatDate(targetDate, "Asia/Tokyo", "yyyy-MM-dd"),
      staffName: normalizedStaffName,
      appointments: []
    };
  }

  const group = groupedEvents[staff.id];
  const appointments = group ? group.appointments : [];

  return {
    success: true,
    date: Utilities.formatDate(targetDate, "Asia/Tokyo", "yyyy-MM-dd"),
    staffName: staff.name,
    appointments: appointments.map(app => ({
      title: app.customerInfo ? app.customerInfo.name : "",
      eventType: app.eventType,
      start: Utilities.formatDate(app.startTime, "Asia/Tokyo", "HH:mm"),
      end: Utilities.formatDate(app.endTime, "Asia/Tokyo", "HH:mm")
    }))
  };
}

function getAttendancePrefillForStaffOnDate(staffName, dateString) {
  const schedule = getScheduleForStaffOnDate(staffName, dateString);
  const apps = (schedule.appointments || []).slice(0, 3);
  const rowData = {
    C: "", D: "", E: "",
    L: "", M: "", N: "",
    U: "", V: "", W: ""
  };

  if (apps[0]) {
    rowData.C = apps[0].title || "";
    rowData.D = apps[0].start || "";
    rowData.E = apps[0].end || "";
  }
  if (apps[1]) {
    rowData.L = apps[1].title || "";
    rowData.M = apps[1].start || "";
    rowData.N = apps[1].end || "";
  }
  if (apps[2]) {
    rowData.U = apps[2].title || "";
    rowData.V = apps[2].start || "";
    rowData.W = apps[2].end || "";
  }

  return {
    success: true,
    date: schedule.date,
    staffName: schedule.staffName,
    appointments: apps,
    rowData: rowData
  };
}
