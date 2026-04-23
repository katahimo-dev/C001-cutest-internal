// ==========================================
// CRM 連携設定 (2026-04-23 追加)
// ==========================================
const CRM_CONFIG = {
  CRM_SHEET_ID: '1GdrbxkUpFvL5mrLjjOAJYyn3AQxXJDl6NYvNKhL1HQk',
  INCOMPLETE_TAB: '要情報更新顧客',
  CUSTOMER_ID_COL: 16  // 0-indexed、ルート集計末尾に追加した顧客ID列 (Q列)
};

/**
 * CRM の「要情報更新顧客」タブを読み込んで、顧客ID → 理由 のマップを返す
 */
function loadIncompleteCustomerMap_() {
  try {
    const ss = SpreadsheetApp.openById(CRM_CONFIG.CRM_SHEET_ID);
    const sheet = ss.getSheetByName(CRM_CONFIG.INCOMPLETE_TAB);
    if (!sheet) {
      console.log('要情報更新顧客タブが見つかりません (CRM側で未生成?)');
      return {};
    }
    const data = sheet.getDataRange().getValues();
    const map = {};
    for (let i = 1; i < data.length; i++) {
      const id = String(data[i][0] || '').trim();
      const reasons = String(data[i][2] || '').trim();
      if (id && reasons) map[id] = reasons;
    }
    console.log('要情報更新顧客 ' + Object.keys(map).length + '件 ロード完了');
    return map;
  } catch (e) {
    console.warn('要情報更新顧客リスト読み込み失敗: ' + e.message);
    return {};
  }
}

/**
 * LINE WORKSへ翌日の予定を送信するメイン関数
 */
function sendDailyScheduleToLineWorks(targetDate) {

  const y = targetDate.getFullYear();
  const m = ('0' + (targetDate.getMonth() + 1)).slice(-2);
  const d = ('0' + targetDate.getDate()).slice(-2);
  const dateStr = `${y}-${m}-${d}`;
  const sheetName = `${y}${m}`;

  // 2. スタッフ名簿（CSV）から LINE WORKS ID (メールID) を取得
  const staffList = getStaffDataFromSpreadsheet(CONFIG.STAFF_SS_ID);

  // 3. ルート集計シートからデータを取得
  const ss = getOrCreateSpreadsheetInSameFolder(CONFIG.ROUTE_SUMMARY_FILE_NAME);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return console.log("対象シートがありません");

  const data = sheet.getDataRange().getValues();
  data.shift(); // ヘッダー除去

  // 4. 指定日のデータのみ抽出、スタッフごとにグループ化
  const scheduleByStaff = {};
  data.forEach(row => {
    const rowDate = (row[0] instanceof Date) ? Utilities.formatDate(row[0], "JST", "yyyy-MM-dd") : row[0];
    if (rowDate === dateStr) {
      const staffName = row[1];
      if (!scheduleByStaff[staffName]) scheduleByStaff[staffName] = [];
      scheduleByStaff[staffName].push(row);
    }
  });

  if (Object.keys(scheduleByStaff).length === 0) {
    console.log(`${dateStr} の予定データがスプレッドシート内に見つかりませんでした。`);
    return;
  }

  // 5. アクセストークン取得
  const accessToken = getLineWorksAccessToken();

  // 5-b. 要情報更新顧客マップ事前ロード (2026-04-23 追加)
  const incompleteMap = loadIncompleteCustomerMap_();

  // 6. メッセージ作成と送信
  for (const staffName in scheduleByStaff) {
    const staffInfo = staffList.find(s => s.name.replace(/\s+/g, "") === staffName.replace(/\s+/g, ""));
    if (!staffInfo || !staffInfo.lwId) {
      console.warn(`スタッフ ${staffName} のID（LINE WORKS ID）が見つかりません。`);
      continue;
    }

    const appointments = scheduleByStaff[staffName];
    let message = `【明日の予定: ${dateStr}】\n\n`;
    const reminderNames = []; // 要情報更新顧客の名前リスト (2026-04-23 追加)

    appointments.forEach((row, index) => {
      const customerName = row[3];
      const customerId = String(row[CRM_CONFIG.CUSTOMER_ID_COL] || '').trim(); // 2026-04-23 追加
      const startTime = formatTimeValue(row[4]);
      const endTime = formatTimeValue(row[5]);
      const timeRange = `${startTime}~${endTime}`;
      const reservaUrl = row[6];
      const moveUrl = row[7];      // 案件間移動
      const attendanceUrl = row[10]; // 出勤経路

      const num = index + 1;
      message += `予約${num}: ${customerName}\n`;
      message += `時刻${num}: ${timeRange}\n`;
      if (reservaUrl) message += `予約詳細${num}: ${reservaUrl}\n`;

      // 最初は出勤経路、それ以外は案件間移動を表示
      if (index === 0) {
        message += `ルート（スタッフ宅→顧客）${num}: ${attendanceUrl}\n`;
      } else {
        message += `ルート（顧客→顧客）${num}: ${moveUrl}\n`;
      }
      message += `\n`;

      // 要情報更新顧客の記録 (2026-04-23 追加)
      if (customerId && incompleteMap[customerId]) {
        reminderNames.push(customerName);
      }
    });

    // 最後の退勤経路
    const lastApp = appointments[appointments.length - 1];
    const leavingUrl = lastApp[13];
    message += `ルート（顧客→スタッフ宅）: ${leavingUrl}\n`;

    // 日報アプリのURL
    message += `日報アプリ: ${CONFIG.REPORT_APP_URL}`;

    // ★要情報更新顧客の追記 (2026-04-23 追加)
    if (reminderNames.length > 0) {
      message += `\n\n━━━ 訪問時お願い ━━━\n`;
      reminderNames.forEach(name => {
        message += `★${name} 様へ、ご本人の生年月日情報のご登録（家族情報欄に記載）を促してください\n`;
      });
    }

    // 送信実行
    console.log(`送信メッセージ: ${message}`);
    sendLineWorksMessage(accessToken, staffInfo.lwId, message);
  }
}

/**
 * LINE WORKS メッセージ送信実行
 */
function sendLineWorksMessage(token, userId, text) {
  const url = `https://www.worksapis.com/v1.0/bots/${CONFIG.LW.BOT_ID}/users/${userId}/messages`;
  console.log({url})
  const payload = {
    "content": {
      "type": "text",
      "text": text
    }
  };

  const options = {
    method: "post",
    headers: {"Authorization": "Bearer " + token},
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };


  const res = UrlFetchApp.fetch(url, options);
  console.log(`送信結果 (${userId}): ${res.getContentText()}`);
}

/**
 * LINE WORKS アクセストークン取得 (JWT認証)
 */
function getLineWorksAccessToken() {
  const rawKey = CONFIG.LW.PRIVATE_KEY;
  if (!rawKey) throw new Error("PrivateKeyがスクリプトプロパティに設定されていません。");

  const formattedKey = formatPrivateKey(rawKey);

  const url = "https://auth.worksmobile.com/oauth2/v2.0/token";
  const header = { alg: "RS256", typ: "JWT" };
  const now = Math.floor(Date.now() / 1000);

  const payload = {
    iss: CONFIG.LW.CLIENT_ID,
    sub: CONFIG.LW.SERVICE_ACCOUNT,
    iat: now,
    exp: now + 3600
  };

  const jwt = createJwt(header, payload, formattedKey);

  const options = {
    method: "post",
    payload: {
      assertion: jwt,
      grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
      client_id: CONFIG.LW.CLIENT_ID,
      client_secret: CONFIG.LW.CLIENT_SECRET,
      scope: "bot"
    },
    muteHttpExceptions: true
  };

  const res = UrlFetchApp.fetch(url, options);
  const json = JSON.parse(res.getContentText());

  if (!json.access_token) {
    throw new Error("トークン取得失敗: " + res.getContentText());
  }

  return json.access_token;
}

// JWT生成用ユーティリティ
function createJwt(header, payload, privateKey) {
  const encode = (obj) => Utilities.base64EncodeWebSafe(JSON.stringify(obj)).replace(/=+$/, "");
  const stringToSign = encode(header) + "." + encode(payload);
  const signature = Utilities.computeRsaSha256Signature(stringToSign, privateKey);
  return stringToSign + "." + Utilities.base64EncodeWebSafe(signature).replace(/=+$/, "");
}


/**
 * Helper to ensure Private Key is in correct PEM format for GAS
 */
function formatPrivateKey(key) {
    if (!key) return "";

    let body = key;
    body = body.replace(/\\n/g, '');
    body = body.replace(/-----BEGIN[^-]+-----/g, '');
    body = body.replace(/-----END[^-]+-----/g, '');
    body = body.replace(/\s+/g, '');

    const chunks = body.match(/.{1,64}/g);
    if (!chunks) {
        console.error("Failed to parse Private Key body.");
        return key;
    }
    const chunkedBody = chunks.join('\n');

    return '-----BEGIN PRIVATE KEY-----\n' + chunkedBody + '\n-----END PRIVATE KEY-----';
}

/**
 * 秘密鍵を正しくプロパティに保存するための関数
 */
function savePrivateKeyToProperty() {
  const myKey = `-----BEGIN PRIVATE KEY-----
（ここに入力）
-----END PRIVATE KEY-----`;

  PropertiesService.getScriptProperties().setProperty('LW_PRIVATE_KEY', myKey);
  console.log("秘密鍵を保存しました。");
}
