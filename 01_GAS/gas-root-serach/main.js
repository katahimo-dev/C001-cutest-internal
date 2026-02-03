// ==========================================
// 設定・定数（変更があれば修正してください）
// ==========================================
const CONFIG = {
  CUSTOMER_FOLDER_ID: "1wLjR6iZ447tbUa3ff59bejoM5aXh8clC", 
  STAFF_SS_ID: "1exqD69qZqACm9KOUPpa0fVWRYD2qEZfce7I6TOs_VDk",
  ROUTE_SUMMARY_FILE_NAME: "ルート集計", // 保存用ファイル名
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
  main("2026/1/30"); // ここを書き換えるだけでOK
}

/**
 * メイン実行関数
 * @param {string} date - "2025/12/25" 形式の文字列（指定日実行用）
 */
function main(date) {
  let targetDate;

  if (date && typeof date === 'string') {
    // 1. 引数がある場合（指定日実行用）
    targetDate = new Date(date);
    console.log(`--- 手動実行モード：指定日 ${date} を処理します ---`);
  } else {
    // 2. 引数がない場合（自動実行・通常のボタン実行）
    targetDate = new Date();
    targetDate.setDate(targetDate.getDate() + 1); // 翌日
    console.log("--- 自動実行モード：翌日の予定を処理します ---");
  }

  // 念のため、日付が正しく生成されたか確認
  if (isNaN(targetDate.getTime())) {
    console.error("エラー: 無効な日付が指定されました。");
    return;
  }

  console.log("実行対象日:", Utilities.formatDate(targetDate, "JST", "yyyy-MM-dd"));

  const y = targetDate.getFullYear();
  const m = ('0' + (targetDate.getMonth() + 1)).slice(-2);
  const d = ('0' + targetDate.getDate()).slice(-2);
  const dateStr = `${y}-${m}-${d}`; 
  const sheetName = `${y}${m}`;    

  console.log(`処理対象日: ${dateStr} / シート名: ${sheetName}`);

  const folder = DriveApp.getFolderById(CONFIG.CUSTOMER_FOLDER_ID);
  const customerData = getCustomerDataFromCsv(folder);
  const staffData = getStaffDataFromSpreadsheet(CONFIG.STAFF_SS_ID);  

  // 1. 予定取得（通常予約と[イベント]の両方）
  const calendarEvents = getCalendarEvents(targetDate, customerData);
  
  if (calendarEvents.length === 0) {
    console.log("予定がありません。終了します。");
    return;
  }

  // 2. グループ化（招待ゲストとの照合ロジックを含む）
  const groupedEvents = groupEventsByStaff(calendarEvents, staffData);

  // 3. ルート計算
  const outputRows = calculateDetailedRoutes(groupedEvents, dateStr, targetDate);

  // 4. ルート集計へ出力
  outputToSheetAppend(outputRows, sheetName);

  // 5. LW送信
  sendDailyScheduleToLineWorks(targetDate);
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

  // 1. 各カレンダーからイベントを収集し、誰のカレンダーかを保持
  calendars.forEach(calendar => {
    const calendarName = calendar.getName(); // カレンダー名（スタッフ名）
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
    const ownerName = data.ownerName; // カレンダーの持ち主（スタッフ名）
    
    const subject = event.getTitle();
    const description = event.getDescription() || "";
    const location = event.getLocation();
    
    const isConfirmed = subject.includes("[予約確定]");
    const isNewCustomer = subject.includes("[新規]");
    const isSpecialEvent = subject.includes("[イベント]");
    const isOfficeWork = subject.includes("[事務]");

    // タイトルから[予約確定]等のタグを取り除く
    const cleanedSubjectName = subject
          .replace("[予約確定]", "")
          .replace("[新規]", "")
          .replace("[イベント]", "")
          .replace("[事務]", "")
          .trim();

// --- パターンA: 通常の顧客予約（[予約確定] が必須） ---
    if (isConfirmed) {
      // 1. 本文からスタッフ名を取得（施設：〇〇[独立] の形式）
      const staffMatch = description.match(/施設：(.*?)\[/);
      const staffNameFromDesc = staffMatch ? staffMatch[1].trim() : ownerName;

      // 2. メニューに「オンライン」が含まれるかチェック
      const isOnline = description.includes("オンライン");

      // 3. 顧客名簿との照合
      let customer = customerList.find(c => {
        const targetName = String(c.name || "").replace(/\s+/g, "");
        const searchName = cleanedSubjectName.replace(/\s+/g, "");
        return targetName === searchName;
      });

      // 4. 名簿にない場合は名前だけ保持し、場所を空にする
      if (!customer) {
        customer = {
          name: cleanedSubjectName,
          address: "",
          lat: "",
          lng: "",
          parkingarea: ""
        };
      }

      // オンラインの場合は名簿にあっても場所情報を消去（ルート計算除外用）
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
        staffNameRaw: staffNameFromDesc, // 本文から抽出した名前を優先
        reservaUrl: urlMatch ? urlMatch[0] : "",
        isMeeting: false
      });
    }
    // --- パターンB: [新規] ---
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
        staffNameRaw: ownerName, // カレンダーの持ち主名を採用
        reservaUrl: "",
        isMeeting: false
      });
    }
    // --- パターンC: 会議等の [イベント] ---
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
        reservaUrl: "",
        isMeeting: true
      });
    }
    // --- パターンD: [事務] ---  
    else if (isOfficeWork) {
      // 事務作業は移動計算に含めないため、住所や座標は空のまま
      processedEvents.push({
        startTime: event.getStartTime(),
        endTime: event.getEndTime(),
        eventType: "OFFICE WORK", // 種別を区別
        customerInfo: {
          name: cleanedSubjectName,
          address: "",
          lat: "",
          lng: "",
          parkingarea: ""
        },
        staffNameRaw: ownerName, 
        reservaUrl: "",
        isMeeting: false 
      });
    }
  });


  return processedEvents;
}

// 住所から緯度経度を取得するヘルパー
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
      // [イベント] の場合：招待ゲスト全員とスタッフ名簿を照合
      event.guestNames.forEach(guestName => {
        const staff = staffList.find(s => normalize(s.name) === normalize(guestName));
        if (staff) {
          addEventToGroup(staffGroups, staff, event);
        }
      });
    } else {
      // 通常予約の場合
      const staff = staffList.find(s => normalize(s.name) === normalize(event.staffNameRaw));
      if (staff) {
        addEventToGroup(staffGroups, staff, event);
      }
    }
  });

  // 時間順にソート
  Object.keys(staffGroups).forEach(key => {
    staffGroups[key].appointments.sort((a, b) => a.startTime - b.startTime);
  });
  return staffGroups;
}

// グループ追加用ヘルパー
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

    // 「場所がある予定」だけを抽出したリストを作る
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
        // --- 1. 出勤 (自宅 -> 最初の場所) ---
        // 有効な予定リストの中で、自分が「最初」なら出勤計算
        if (currentApp === validLocApps[0]) {
          attendanceInfo = getRouteDetails(staff, currentApp.customerInfo, targetDate);
        }

        // --- 2. 移動 (前の場所 -> 現在の場所) ---
        // 有効な予定リストの中で、自分より前に場所がある予定があれば計算
        const myIndexInValid = validLocApps.indexOf(currentApp);
        if (myIndexInValid > 0) {
          const prevValidApp = validLocApps[myIndexInValid - 1];
          moveInfo = getRouteDetails(prevValidApp.customerInfo, currentApp.customerInfo, targetDate);
        }

        // --- 3. 退勤 (最後の場所 -> 自宅) ---
        // 有効な予定リストの中で、自分が「最後」なら退勤計算
        if (currentApp === validLocApps[validLocApps.length - 1]) {
          leavingInfo = getRouteDetails(currentApp.customerInfo, staff, targetDate);
        }
      }

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
        // currentApp.customerInfo.parkingarea
      ]);
    }
  });
  return allRows;
}


/**
 * ルート詳細を取得する
 * @param {Object} from 出発地情報
 * @param {Object} to 目的地情報
 * @param {Date} targetDate 判定対象日
 */
function getRouteDetails(from, to, targetDate) {
  // 1. 各地点の最終的な住所と座標を決定する
  const origin = resolveLocation(from, targetDate);
  const destination = resolveLocation(to, targetDate);

  // どちらかの座標が取得できない場合は計算不可
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
        // 元のコードのURL形式を維持（緯度経度ベース）
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

/**
 * 期間限定住所の判定と座標の補完を行う内部補助関数
 */
function resolveLocation(loc, targetDate) {
  let finalAddress = loc.address;
  let finalLat = loc.lat;
  let finalLng = loc.lng;

  // --- A. 期間限定住所 (address2) の判定 ---
  if (loc.address2 && loc.address2_start && loc.address2_end) {
    const startDate = new Date(loc.address2_start);
    const endDate = new Date(loc.address2_end);
    
    // 時間をリセットして日付のみで比較
    const checkDate = new Date(targetDate.getTime());
    checkDate.setHours(0, 0, 0, 0);
    startDate.setHours(0, 0, 0, 0);
    endDate.setHours(0, 0, 0, 0);

    if (checkDate >= startDate && checkDate <= endDate) {
      finalAddress = loc.address2;
      // address2に切り替わった場合は、元の緯度経度は使えないためリセット
      finalLat = "";
      finalLng = "";
    }
  }

  // --- B. 座標の補完 (緯度経度が空の場合) ---
  if (!finalLat || !finalLng) {
    if (finalAddress) {
      const geo = getLatLngFromAddress(finalAddress); // 既存のヘルパー関数を呼び出し
      finalLat = geo.lat;
      finalLng = geo.lng;
    }
  }

  return { lat: finalLat, lng: finalLng };
}


/**
 * 指定フォルダ内の最新CSV（Kokyaku_...）を取得して解析する
 */
function getCustomerDataFromCsv(folder) {
  // 1. 「Kokyaku_」で始まるファイルをすべて取得
  const files = folder.getFiles();
  const targetFiles = [];

  while (files.hasNext()) {
    const file = files.next();
    const name = file.getName();
    // ファイル名が Kokyaku_ で始まり .csv で終わるものを抽出
    if (name.startsWith("Kokyaku_") && name.endsWith(".csv")) {
      targetFiles.push(file);
    }
  }

  if (targetFiles.length === 0) {
    throw new Error("対象のCSVファイルが見つかりません。");
  }

  // 2. 最終更新日時（Last Updated）でソートし、最新の1件を取得
  targetFiles.sort((a, b) => b.getLastUpdated() - a.getLastUpdated());
  const latestFile = targetFiles[0];
  console.log(`読み込み対象ファイル: ${latestFile.getName()} (更新日: ${latestFile.getLastUpdated()})`);

  // 3. CSVデータを解析（文字コードを考慮してBlobから取得）
  const csvContent = latestFile.getBlob().getDataAsString("UTF-16"); 
  const values = Utilities.parseCsv(csvContent, '\t');

  if (values.length < 2) return []; // データがない場合

  // 4. ヘッダーとインデックスの取得
  const header = values.shift();
  const idx = getColumnIndices(header);

  // 5. データの成形
  return values.map(row => {
    const lastName = String(row[idx.lastName] || "");
    const firstName = String(row[idx.firstName] || "");

    // const id = extractIdFromName(lastName) || extractIdFromName(firstName);
    // --- 緯度・経度の分割処理 ---
    let lat = row[idx.lat];
    let lng = row[idx.lng];
    
    // 「緯度・経度」カラムに値がある場合、カンマで分割して優先的に使用
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
  }).filter(item => item.id); // IDが存在するものだけを返す
}


/**
 * スプレッドシートからスタッフデータを取得する
 * @param {string} spreadsheetId スプレッドシートのID
 */
function getStaffDataFromSpreadsheet(spreadsheetId) {
  // 1. IDでスプレッドシートを開く
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheets()[0]; // 一番左のシートを取得
  
  // 2. データをすべて取得（二次元配列）
  const data = sheet.getDataRange().getValues();
  
  if (data.length < 2) throw new Error("スプレッドシートにデータが見つかりません。");

  // 3. ヘッダーを取得し、列番号のインデックスを特定
  const header = data.shift();
  const idx = getColumnIndices(header);

  // 4. 各行をオブジェクトに変換
  return data.map(row => {
    const lastName = String(row[idx.lastName] || "");
    const firstName = String(row[idx.firstName] || "");
    
    // 名前からIDを抽出、または1列目(row[0])をデフォルトIDとする
    const id = extractIdFromName(lastName) || extractIdFromName(firstName) || row[0];
    
    // --- 緯度・経度の分割処理 ---
    let lat = row[idx.lat];
    let lng = row[idx.lng];
    
    // 「緯度・経度」カラムに値がある場合、カンマで分割して優先的に使用
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
    address2: find(['住所2', '住所２']),
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
  const header = [
    "日付", "スタッフ名", "顧客ID", "顧客名", "開始時間", "終了時間", "予約詳細URL",
    "移動経路URL", "移動時間（分）", "移動距離（km）",
    "出勤経路URL", "出勤経路時間（分）", "出勤経路距離（km）",
    "退勤経路URL", "退勤経路時間（分）", "退勤経路距離（km）", 
    // "駐車場"
  ];
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(header);
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


/**
 * Dateオブジェクトまたは時刻データを "HH:mm" 形式の文字列に変換する
 */
function formatTimeValue(val) {
  if (val instanceof Date) {
    return Utilities.formatDate(val, "JST", "HH:mm");
  }
  // すでに文字列で "10:00:00" などとなっている場合の簡易処理
  if (typeof val === 'string' && val.includes(':')) {
    return val.substring(0, 5);
  }
  return val;
}