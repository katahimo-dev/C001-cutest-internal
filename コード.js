// ==========================================
// Web App Logic
// ==========================================

const SPREADSHEET_ID = '1Y0ciyR4LHwCPwIkcuqwhQnD9k_-WBHn0nnwNRzFaacM';
const TARGET_GID = 1224762512;
const PROMPT_SHEET_NAME = 'ＡＩプロンプト';
const REPORT_SHEET_NAME = '日報';
const STAFF_SS_ID = "1exqD69qZqACm9KOUPpa0fVWRYD2qEZfce7I6TOs_VDk";
const STAFF_GID = 2002628493;

// New Constants for Receipt Images
const RECEIPT_FOLDER_ID = '1fruKZdH4gigbUB54VOrvYDa31cotVSdD';
const IMAGE_LOG_SS_ID = '18tkd37ck6fmhPMc_Ep1YEs4FXjB2HqXgMRfm0oJlkWI';

const PROMPT_KEYS = {
    GENERATE_WITH_WARNINGS: 'GenerateWithWarnings',
    GENERATE_ACCIDENT: 'GenerateAccident',
    PLACEHOLDER_DAILY: 'PlaceholderDaily',
    PLACEHOLDER_ACCIDENT: 'PlaceholderAccident',
    HINT_ACCIDENT: 'HintAccident',
    PLACEHOLDER_HIYARI: 'PlaceholderHiyari'
};

const DEFAULT_PROMPTS = {
    [PROMPT_KEYS.GENERATE_WITH_WARNINGS]: `
あなたは保育士の業務を支援するAIアシスタントです。
以下の「保育日報のメモ（口語）」をもとに、日報を作成してください。

# 必須情報チェック
以下の3点が入力テキストに含まれているか確認してください。
1. **訪問当日のサポート内容** (具体的に何をしたか)
2. **お客様情報** (家庭内の状況や家族との会話から見えた生活状況など)
3. **振り返り** (自分のサポートに対しての内省・今回どうだったか)

# 指示
- 不足している必須情報があれば、その項目名を "warnings" 配列にリストアップしてください（例: ["お客様情報", "振り返り"]）。
- 不足情報の有無に関わらず、入力された情報を元に可能な範囲でレポートを作成してください。

# 入力テキスト
{anonymizedText}
時間情報: {timeInfo}

# 出力フォーマット (JSON)
{
  "warnings": ["不足項目名1", "不足項目名2"], // なければ空配列 []
  "internal": "社内向けレポート内容（事実・客観的）。読みやすさのため、適宜改行コード(\\n)を含めてください。",
  "customer": "保護者向けレポート内容（親しみやすく）。読みやすさのため、適宜改行コード(\\n)を含めてください。"
}
`,
    [PROMPT_KEYS.GENERATE_ACCIDENT]: `
あなたは保育園の事故報告書作成を支援するAIです。
入力された状況説明（メモ）から、以下の項目に整理・分解してJSON形式で出力してください。

# 入力テキスト
{anonymizedText}
時間情報: {timeInfo}

# 出力項目とルール
- occurrenceTime: 発生日時（令和〇年〇月〇日...の形式が望ましいが、入力から推測できる範囲で。不明なら「要確認」としてください）
- location: 発生場所（施設名＋部屋名、屋外ならエリアなど）
- accidentContent: 事故内容（端的な見出し。例：転倒による額切創）
- situation: 発生状況（5W1H、時系列。推測は避け事実のみ）
- immediateResponse: 発生時の対応（誰が、何分後に、何をしたか。タイムライン形式など）
- parentCorrespondence: 保護者への対応（連絡手段、時刻、反応、受診予定など）
- diagnosisTreatment: 診断名および処置状況/必要診察日数（未受診なら「診療前」と明記）
- prevention: 事故防止に向けた今後の対応（原因分析、一次対策、恒久対策）

# 出力フォーマット (JSON)
{
  "occurrenceTime": "...",
  "location": "...",
  "accidentContent": "...",
  "situation": "...",
  "immediateResponse": "...",
  "parentCorrespondence": "...",
  "diagnosisTreatment": "...",
  "prevention": "..."
}
`
    ,
    [PROMPT_KEYS.PLACEHOLDER_DAILY]: `①訪問当日のサポート内容
   　（実際に実施した保育・家事・対応内容など）
② お客様情報
   　（家庭内の状況、保護者や子どもの様子、会話から見えた生活状況・要望・健康面など）
③　振り返り
   　（支援中の状況→対応→結果、気づき、改善点、次回への申し送りなど）`,
    [PROMPT_KEYS.PLACEHOLDER_ACCIDENT]: `①事実を時系列で、客観的に
感情的な表現や推測は避け、見聞きした事実のみを時系列に並べます。

②「5W1H＋初動対応」を意識
いつ・どこで・誰が・何をしていて・何が起こり・どう対処したかを必ず押さえます。

③ヒヤリハットも記録
ヒヤリハットも重大事故と同じ視点で記録し、要因分析と再発防止策を残すことで重大事故を防げます`,
    [PROMPT_KEYS.HINT_ACCIDENT]: `事故報告書 記載項目と記載要領

発生日時
「令和〇年〇月〇日（曜）午後〇時〇分頃」の形で、分単位まで記載。発見時刻と発生時刻が異なる場合は両方書く。

発生場所
施設名＋部屋名／屋外の場合はエリアまで具体的に
（例：〇〇公園すべり台下）。

事故内容
端的な見出し語で
「転倒による額切創」「アレルギー症状（じんましん）」など
原因＋結果をセットで。

発生状況
①環境 ②子どもの行動 ③職員配置 ④事故発生の瞬間
の順に、1文1事実で記録。観察できない部分は書かない。

発生時の対応
①誰が ②何分後に ③何をしたのかをタイムライン形式で。
「14:05 冷水で5分間冷却 → 14:10 止血確認 → 14:12 保護者へ電話」など。

保護者への対応
連絡手段・時刻・先方の反応・今後の受診予定を簡潔に。
「14:12 母・携帯へ連絡、15:00 来園し受診同意」

診断名および処置状況／必要診察日数
受診後に医師の診断名を正式に転記。
未受診の段階では「診察前」と明記し暫定措置を書く。

事故防止に向けた今後の対応
①原因分析（環境・人・手順の観点で）
→②一次対策（急ぎの安全策）
→③恒久対策（マニュアル改訂・研修など）
を箇条書きで。`,
    [PROMPT_KEYS.PLACEHOLDER_HIYARI]: `■ヒヤリハットを記入するときの追加留意点
①「もし○○していたら重大事故」まで想定して原因を書く
例：「高さ60 cmの踏み台から足を滑らせたが、すぐ横に職員がいて転落を回避」
②再発防止策を必ず具体化（配置変更、備品購入、声かけ方法など）
 
■よくあるNG集
 
NG例	修正方法
主観的表現 「急に暴れ出した」	行動を具体的に「立ち上がって走り出した」
「たぶん眠かった」	憶測を削除 or 根拠を追記「午睡前で目をこすっていたため眠気があった可能性」
時刻抜け・曖昧な順序	タイムラインで整理し、時計を確認して都度メモ。
再発防止策が抽象的 「注意する」	「○月○日までに踏み台に滑り止めテープを貼付、写真を共有」など行動・期限・担当を明示。`
};

function doGet() {
    return HtmlService.createTemplateFromFile('index').evaluate()
        .setTitle('保育日報アプリ')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Helper to get the Spreadsheet object
 */
function getSpreadsheet() {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
}

/**
 * Finds the sheet by GID
 */
function getSheetByGid(ss, gid) {
    const sheets = ss.getSheets();
    for (let i = 0; i < sheets.length; i++) {
        if (sheets[i].getSheetId() === gid) {
            return sheets[i];
        }
    }
    return null;
}

/**
 * Fetches Customer data from the specified Spreadsheet/Sheet.
 * Columns: B (Surname/Name), S (Address).
 * Filter: Extract City from Address.
 */
/**
 * Fetches Customer data from the specified Spreadsheet/Sheet.
 * Returns full data for detailed view.
 */
function getData() {
    const ss = getSpreadsheet();
    let cSheet = ss.getSheetByName(CUSTOMER_DB_NAME);
    let fSheet = ss.getSheetByName(FAMILY_DB_NAME);

    // Fallback if sheets don't exist yet (return empty)
    if (!cSheet || !fSheet) return { cities: [], customers: [] };

    const cData = cSheet.getDataRange().getValues();
    const fData = fSheet.getDataRange().getValues();

    if (cData.length === 0) return { cities: [], customers: [] };

    // Headers are row 0. Data starts row 1.
    const cHeaders = cData[0];
    const cRows = cData.slice(1);
    const fRows = fData.slice(1);

    // Identify Columns by Header Name
    const mapIndices = {};
    cHeaders.forEach((h, i) => mapIndices[h] = i);

    // Helper indices
    const idxId = mapIndices['顧客ID'] !== undefined ? mapIndices['顧客ID'] : 0;
    const idxSei = mapIndices['姓'];
    const idxMei = mapIndices['名'];
    const idxName = mapIndices['氏名'];
    const idxAddr = mapIndices['住所'];
    // For Map
    const idxLat = mapIndices['緯度'];
    const idxLng = mapIndices['経度'];

    // Helper to safe string
    const toStr = (val) => (val === null || val === undefined) ? "" : String(val);
    const fmtDate = (val) => {
        if (val instanceof Date) {
            return Utilities.formatDate(val, "Asia/Tokyo", "yyyy/MM/dd");
        }
        return toStr(val);
    };

    // Map Families by CustomerID
    const familyMap = {};
    fRows.forEach(row => {
        const cId = toStr(row[0]);
        if (!cId) return;
        if (!familyMap[cId]) familyMap[cId] = [];
        familyMap[cId].push({
            name: toStr(row[1]),
            dob: fmtDate(row[2]),
            job: toStr(row[3]),
            allergy: toStr(row[4]),
            info: toStr(row[5])
        });
    });

    const customers = cRows.map(row => {
        const id = toStr(row[idxId]);
        if (!id) return null; // Skip empty IDs

        let name = "";
        // Construct Name
        if (idxSei !== undefined && idxMei !== undefined) {
            const s = toStr(row[idxSei]);
            const m = toStr(row[idxMei]);
            name = s + " " + m;
        } else if (idxName !== undefined) {
            name = toStr(row[idxName]);
        } else {
            name = "Unknown";
        }
        name = name.trim();

        const address = idxAddr !== undefined ? toStr(row[idxAddr]) : "";
        const lat = idxLat !== undefined ? row[idxLat] : "";
        const lng = idxLng !== undefined ? row[idxLng] : "";

        const fmtDateTime = (val) => {
            if (val instanceof Date) {
                return Utilities.formatDate(val, "Asia/Tokyo", "yyyy/MM/dd HH:mm");
            }
            return toStr(val);
        };


        // Build Full Data Object (excluding sensitive)
        // Use array to guarantee order
        const orderedDetails = [];
        cHeaders.forEach((h, i) => {
            // Exclude fields
            if (h === 'パスワード' || h === '顧客ID') return;

            let val = row[i];
            // Smart formatting based on type and header
            if (val instanceof Date) {
                if (h.includes('生年月日') || h.includes('誕生日')) {
                    val = Utilities.formatDate(val, "Asia/Tokyo", "yyyy/MM/dd");
                } else {
                    // Assume timestamp for other dates like '登録日時', '最終更新日時'
                    val = Utilities.formatDate(val, "Asia/Tokyo", "yyyy/MM/dd HH:mm");
                }
            } else {
                val = toStr(val);
            }

            orderedDetails.push({ key: h, value: val });
        });

        // Extract City
        let city = '';
        if (address) {
            const cityMatch = address.match(/(?:東京都|北海道|(?:京都|大阪)府|.{2,3}県)([^市区町村]+[市区町村])/);
            if (cityMatch) {
                city = cityMatch[1];
            } else {
                const simpleMatch = address.match(/^([^0-9]+?[市区町村])/);
                if (simpleMatch) city = simpleMatch[1];
            }
        }

        return {
            id: id,
            name: name,
            address: address,
            city: city,
            lat: lat,
            lng: lng,
            family: familyMap[id] || [],
            details: orderedDetails
        };
    }).filter(c => c !== null);

    // Unique cities for filter
    const cities = [...new Set(customers.map(c => c.city).filter(c => c))].sort();

    return { cities, customers, version: checkDataVersion() };
}

/**
 * Gets prompt from sheet or returns default (and saves it).
 */
function getPrompt(key) {
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName(PROMPT_SHEET_NAME);

    if (!sheet) {
        sheet = ss.insertSheet(PROMPT_SHEET_NAME);
        sheet.appendRow(['Key', 'Prompt Template']);
        Object.keys(DEFAULT_PROMPTS).forEach(k => {
            sheet.appendRow([k, DEFAULT_PROMPTS[k]]);
        });
        return DEFAULT_PROMPTS[key];
    }

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
        if (data[i][0] === key) {
            return data[i][1];
        }
    }

    const defaultVal = DEFAULT_PROMPTS[key];
    if (defaultVal) {
        sheet.appendRow([key, defaultVal]);
        return defaultVal;
    }

    return "";
}

/**
 * Generates report and returns warnings if info is missing.
 */
// Refactored helper to support multimodal or text-only
function callGemini(apiKey, contentParts, generationConfig, modelName) {
    const model = modelName || 'gemini-2.5-flash';
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;

    const payload = {
        contents: [{ parts: contentParts }],
        generationConfig: generationConfig || { responseMimeType: "application/json" }
    };

    try {
        const response = UrlFetchApp.fetch(url, {
            method: 'post',
            contentType: 'application/json',
            payload: JSON.stringify(payload),
            muteHttpExceptions: true
        });

        if (response.getResponseCode() !== 200) {
            return { error: "API Error: " + response.getContentText() };
        }

        const json = JSON.parse(response.getContentText());
        if (!json.candidates || json.candidates.length === 0) {
            return { error: "No candidates returned" };
        }

        const text = json.candidates[0].content.parts[0].text;
        // Clean markdown JSON if present
        const cleanText = text.replace(/```json/g, '').replace(/```/g, '').trim();
        return JSON.parse(cleanText);

    } catch (e) {
        return { error: "System Error: " + e.message };
    }
}

function generateReportWithWarnings(inputData) {
    const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!apiKey) return { warnings: ["API Key Missing"], internal: "Error: API Key not set", customer: "" };

    let promptTemplate = getPrompt(PROMPT_KEYS.GENERATE_WITH_WARNINGS);
    if (!promptTemplate) promptTemplate = DEFAULT_PROMPTS[PROMPT_KEYS.GENERATE_WITH_WARNINGS];

    let text = "";
    let timeInfo = "時間指定なし";

    if (typeof inputData === 'string') {
        text = inputData;
    } else {
        text = inputData.text;
        if (inputData.start && inputData.end) {
            timeInfo = `${inputData.start}〜${inputData.end}`;
        }
    }

    let prompt = promptTemplate.replace('{anonymizedText}', text);
    prompt = prompt.replace('{timeInfo}', timeInfo);

    // Schema for Daily Report
    const dailyReportSchema = {
        type: "OBJECT",
        properties: {
            warnings: { type: "ARRAY", items: { type: "STRING" } },
            internal: { type: "STRING" },
            customer: { type: "STRING" }
        },
        required: ["warnings", "internal", "customer"]
    };

    // Call with schema
    const result = callGemini(apiKey, [{ text: prompt }], {
        responseMimeType: "application/json",
        responseSchema: dailyReportSchema
    });

    if (result.error) {
        return { warnings: ["API Error"], internal: result.error, customer: "" };
    }
    return result;
}

/**
 * Extracts amount from receipt image
 */
function extractAmountFromImage(base64Image) {
    const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!apiKey) return { amount: "" };

    const prompt = `
    Analyze the image of this receipt.
    Identify the Total Amount (Total, 合計, 支払い金額).
    Return the result in JSON format: {"amount": number}
    If the amount cannot be read, return {"amount": 0}.
    Do NOT include currency symbols in the number value.
    `;

    // Removing header "data:image/jpeg;base64,"
    const rawBase64 = base64Image.split(',')[1];

    // OCR uses gemini-2.5-flash-lite as requested
    const result = callGemini(apiKey, [
        { text: prompt },
        { inline_data: { mime_type: "image/jpeg", data: rawBase64 } }
    ], null, 'gemini-2.5-flash-lite');

    if (result.error) {
        console.error(result.error);
        return { amount: 0 };
    }
    return result;
}

// --- Firestore Integration ---
function getFirestore() {
    /*
    const props = PropertiesService.getScriptProperties();
    const email = props.getProperty('FIREBASE_CLIENT_EMAIL');
    const key = props.getProperty('FIREBASE_PRIVATE_KEY');
    const projectId = props.getProperty('FIREBASE_PROJECT_ID');

    if (!email || !key || !projectId) {
        console.warn("Firestore Config Missing: Email=" + !!email + ", Key=" + !!key + ", ProjectId=" + !!projectId);
        return null;
    }

    return FirestoreApp.getFirestore(email, key.replace(/\\n/g, '\n'), projectId);
    */
    return null;
}

/**
 * Saves the final report to 'Reports' sheet.
 */
/**
 * Saves the final report to 'Reports' sheet.
 */
/**
 * Saves the final report to 'Reports' sheet.
 */
function saveReport(reportData) {
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName(REPORT_SHEET_NAME);

    if (!sheet) {
        sheet = ss.insertSheet(REPORT_SHEET_NAME);
        sheet.appendRow(['Timestamp', 'StartTime', 'EndTime', 'User', 'CustomerId', 'CustomerName', 'InputText', 'InternalReport', 'CustomerReport', 'RiskRating', 'EsRating']);
    }

    let timestampJST;
    if (reportData.reportDate) {
        // Use the selected date and start time
        const timePart = reportData.start || "00:00";
        // Ensure date uses slashes for better compatibility if needed, though most environments handle standard formats
        const datePart = reportData.reportDate.replace(/-/g, '/');
        const d = new Date(`${datePart} ${timePart}`);
        if (!isNaN(d.getTime())) {
            timestampJST = Utilities.formatDate(d, "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");
        } else {
            timestampJST = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");
        }
    } else {
        timestampJST = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");
    }

    sheet.appendRow([
        timestampJST,
        reportData.start || "",
        reportData.end || "",
        reportData.staffName || "", // Use selected staff name only
        reportData.customerId || "",
        reportData.customerName || "",
        reportData.inputText,
        reportData.internalText,
        reportData.customerText,
        reportData.riskRating || "",
        reportData.esRating || ""
    ]);

    // --- Firestore Write ---
    // --- Firestore Write ---
    /*
    const firestore = getFirestore();
    if (firestore) {
        try {
            const docData = {
                timestamp: timestampJST,
                start: reportData.start || "",
                end: reportData.end || "",
                staffName: reportData.staffName || "",
                customerId: reportData.customerId || "",
                customerName: reportData.customerName || "",
                inputText: reportData.inputText,
                internalText: reportData.internalText,
                customerText: reportData.customerText,
                riskRating: reportData.riskRating || 0,
                esRating: reportData.esRating || 0,
                reportDate: reportData.reportDate || "",
                type: 'daily',
                createdAt: new Date()
            };
            firestore.createDocument("reports", docData);
        } catch (e) {
            console.error("Firestore Write Error: " + e.message);
        }
    }
    */

    // Handle Image Uploads with Amount
    let imageStatus = "No Images";
    if (reportData.images && reportData.images.length > 0) {
        try {
            const count = processReceiptImages(reportData.images, reportData.staffName, reportData.customerId);
            imageStatus = `Uploaded ${count} images`;
        } catch (e) {
            console.error("Image Upload Failed: " + e.message);
            // We return the error details so the frontend can warn the user
            // even if the report text was saved.
            return { success: true, message: "日報は保存されましたが、画像保存に失敗しました: " + e.message };
        }
    }

    // --- LineWorks Notification ---
    try {
        const star = (n) => "★".repeat(Number(n) || 0) + "☆".repeat(5 - (Number(n) || 0));
        let ratingsInfo = "";
        if (reportData.riskRating || reportData.esRating) {
            ratingsInfo = "\n\n【評価指標】";
            if (reportData.riskRating) ratingsInfo += `\nリスク: ${star(reportData.riskRating)} (${reportData.riskRating})`;
            if (reportData.esRating) ratingsInfo += `\n満足度: ${star(reportData.esRating)} (${reportData.esRating})`;
        }

        // Requested format: Staff Name + Internal Report + Ratings
        const lwText = `【日報提出】\n担当: ${reportData.staffName}\n\n${reportData.internalText}${ratingsInfo}`;
        sendToLineWorks(lwText);
    } catch (e) {
        console.error("LineWorks Notification Failed: " + e.message);
    }
    // ------------------------------

    return { success: true, message: "保存しました (" + imageStatus + ")" };
}

/**
 * Saves receipt images to Drive and logs to Spreadsheet
 * Updated to handle Amount and CustomerID
 */
function processReceiptImages(imagesData, staffId, customerId) {
    if (!imagesData || imagesData.length === 0) return 0;

    const folder = DriveApp.getFolderById(RECEIPT_FOLDER_ID);
    const ss = SpreadsheetApp.openById(IMAGE_LOG_SS_ID);
    const sheet = ss.getSheets()[0];
    // Ensure header matches new requirements
    // Old: ['日時', 'ユーザーID', 'Googleドライブ写真ファイルへのリンク']
    // New: ['日時', 'ユーザーID', '顧客ID', '金額', 'Googleドライブ写真ファイルへのリンク']
    // We won't delete old data, just append new columns effectively from now on if row 1 is empty
    if (sheet.getLastRow() === 0) {
        sheet.appendRow(['日時', 'ユーザーID', '顧客ID', '金額', 'Googleドライブ写真ファイルへのリンク']);
    }

    const now = new Date();
    const timestamp = Utilities.formatDate(now, "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");
    const filePrefix = Utilities.formatDate(now, "Asia/Tokyo", "yyyyMMdd_HHmmss");

    let successCount = 0;

    // Support legacy array of strings or new array of objects
    imagesData.forEach((item, index) => {
        let base64Data = "";
        let amount = "";

        if (typeof item === 'string') {
            base64Data = item;
        } else {
            base64Data = item.data;
            amount = item.amount;
        }

        const split = base64Data.split(',');
        if (split.length < 2) return; // Skip invalid

        const contentType = split[0].split(':')[1].split(';')[0];
        const bytes = Utilities.base64Decode(split[1]);

        const fileName = `${filePrefix}_${index + 1}_${staffId || 'unknown'}.jpg`;
        const blob = Utilities.newBlob(bytes, contentType, fileName);
        const file = folder.createFile(blob);
        const fileUrl = file.getUrl();

        // Append to Spreadsheet
        sheet.appendRow([timestamp, staffId || '', customerId || '', amount || '', fileUrl]);
        successCount++;
    });
    return successCount;
}

/**
 * Generates accident report decomposition.
 */
function generateAccidentReport(inputData) {
    const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!apiKey) return { error: "API Key Missing" };

    let promptTemplate = getPrompt(PROMPT_KEYS.GENERATE_ACCIDENT);
    if (!promptTemplate) promptTemplate = DEFAULT_PROMPTS[PROMPT_KEYS.GENERATE_ACCIDENT];

    let text = "";
    let timeInfo = "時間指定なし";

    if (typeof inputData === 'string') {
        text = inputData;
    } else {
        text = inputData.text;
        if (inputData.start && inputData.end) {
            timeInfo = `${inputData.start}〜${inputData.end}`;
        } else if (inputData.start) {
            timeInfo = inputData.start; // Just occurrence time
        }
    }

    let prompt = promptTemplate.replace('{anonymizedText}', text);
    prompt = prompt.replace('{timeInfo}', timeInfo);

    // Schema for Accident Report
    const accidentSchema = {
        type: "OBJECT",
        properties: {
            occurrenceTime: { type: "STRING" },
            location: { type: "STRING" },
            accidentContent: { type: "STRING" },
            situation: { type: "STRING" },
            immediateResponse: { type: "STRING" },
            parentCorrespondence: { type: "STRING" },
            diagnosisTreatment: { type: "STRING" },
            prevention: { type: "STRING" }
        },
        required: ["occurrenceTime", "location", "accidentContent", "situation", "immediateResponse", "parentCorrespondence", "diagnosisTreatment", "prevention"]
    };

    const result = callGemini(apiKey, [{ text: prompt }], {
        responseMimeType: "application/json",
        responseSchema: accidentSchema
    });

    if (result.error) {
        return { error: result.error };
    }
    return result;
}

const ACCIDENT_SHEET_NAME = '事故報告';
const PROMPT_ACCIDENT_KEY = 'GenerateAccident';

/**
 * Saves the accident report to 'AccidentReports' sheet.
 */
function saveAccidentReport(reportData) {
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName(ACCIDENT_SHEET_NAME);

    if (!sheet) {
        sheet = ss.insertSheet(ACCIDENT_SHEET_NAME);
        sheet.appendRow([
            'Timestamp', 'Reporter', 'CustomerId', 'CustomerName',
            'OccurrenceTime', 'Location', 'AccidentContent',
            'Situation', 'ImmediateResponse', 'ParentCorrespondence',
            'DiagnosisTreatment', 'Prevention', 'OriginalInput', 'ReportType'
        ]);
    }

    const timestampJST = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");

    sheet.appendRow([
        timestampJST,
        reportData.staffName || "",
        reportData.customerId || "",
        reportData.customerName || "",
        reportData.targetName || "",
        reportData.targetDob || "",
        reportData.occurrenceTime,
        reportData.location,
        reportData.accidentContent,
        reportData.situation,
        reportData.immediateResponse,
        reportData.parentCorrespondence,
        reportData.diagnosisTreatment,
        reportData.prevention,
        reportData.inputText,
        reportData.reportType || "事故報告" // Default to Accident if missing
    ]);

    // --- Firestore Write (Accident) ---
    // --- Firestore Write (Accident) ---
    /*
    const firestore = getFirestore();
    if (firestore) {
        try {
            const docData = {
                timestamp: timestampJST,
                staffName: reportData.staffName || "",
                customerId: reportData.customerId || "",
                customerName: reportData.customerName || "",
                targetName: reportData.targetName || "",
                targetDob: reportData.targetDob || "",
                occurrenceTime: reportData.occurrenceTime,
                location: reportData.location,
                accidentContent: reportData.accidentContent,
                situation: reportData.situation,
                immediateResponse: reportData.immediateResponse,
                parentCorrespondence: reportData.parentCorrespondence,
                diagnosisTreatment: reportData.diagnosisTreatment,
                prevention: reportData.prevention,
                inputText: reportData.inputText,
                reportType: reportData.reportType || "事故報告",
                type: 'accident',
                createdAt: new Date()
            };
            firestore.createDocument("reports", docData);
        } catch (e) {
            console.error("Firestore Write Error (Accident): " + e.message);
        }
    }
    */

    // --- LineWorks Notification ---
    try {
        const typeLabel = reportData.reportType || "事故報告";
        const lwText = `【${typeLabel}】
担当: ${reportData.staffName}
対象: ${reportData.targetName}
生年月日: ${reportData.targetDob}
発生日時: ${reportData.occurrenceTime}
発生場所: ${reportData.location}
事故内容: ${reportData.accidentContent}
発生状況: ${reportData.situation}
発生時の対応: ${reportData.immediateResponse}
保護者への対応: ${reportData.parentCorrespondence}
診断名および処置状況: ${reportData.diagnosisTreatment}
今後の対応: ${reportData.prevention}`;
        sendToLineWorks(lwText);
    } catch (e) {
        console.error("LineWorks Accident Notification Failed: " + e.message);
    }
    // ------------------------------

    return "Success";
}





function verifyLogin(userId, password) {
    try {
        const ss = SpreadsheetApp.openById(STAFF_SS_ID);
        let sheet = null;
        const sheets = ss.getSheets();
        for (let i = 0; i < sheets.length; i++) {
            if (sheets[i].getSheetId() === STAFF_GID) {
                sheet = sheets[i];
                break;
            }
        }

        if (!sheet) return { success: false, message: "Staff sheet not found" };

        const data = sheet.getDataRange().getValues().slice(1); // Skip header

        // New Layout:
        // Col 1: Name
        // Col 9: UserID
        // Col 10: Password
        // Col 11: Admin (1 = true)
        const user = data.find(row => String(row[9]) === userId && String(row[10]) === password);

        if (user) {
            const isAdmin = (user[11] == 1 || user[11] === '1');
            return { success: true, name: user[1], isAdmin: isAdmin };
        } else {
            return { success: false, message: "Invalid ID or password" };
        }
    } catch (e) {
        return { success: false, message: e.message };
    }
}

const CUSTOMER_DB_NAME = '顧客DB_New';
const FAMILY_DB_NAME = '家族DB_New';

// Google Drive Folder ID for Auto CSV Import
const AUTO_CSV_FOLDER_ID = '1wLjR6iZ447tbUa3ff59bejoM5aXh8clC';

/**
 * Checks for the latest CSV in the dedicated folder and imports it if newer.
 * Can be run manually or by a time-based trigger.
 */
function checkAndImportLatestCsv() {
    try {
        const folder = DriveApp.getFolderById(AUTO_CSV_FOLDER_ID);
        // getFilesByName exact match might not work with prefix. Use getFiles() and filter.
        let latestFile = null;
        let latestTime = 0;

        const iterator = folder.getFiles();
        while (iterator.hasNext()) {
            const file = iterator.next();
            const name = file.getName();

            // Match pattern: Kokyaku_YYYYMMDDHHmm_1.csv
            // We just need to parse the timestamp part purely string based or regex
            const match = name.match(/^Kokyaku_(\d{12})_\d+\.csv$/);
            if (match) {
                const timeStr = match[1];
                // Simple numerical comparison is enough for YYYYMMDDHHmm format
                const timeVal = parseInt(timeStr, 10);

                if (timeVal > latestTime) {
                    latestTime = timeVal;
                    latestFile = file;
                }
            }
        }

        if (!latestFile) {
            console.log("No matching CSV files found.");
            return "No files found";
        }

        // Check if we already imported this version
        const props = PropertiesService.getScriptProperties();
        const currentVersion = parseInt(props.getProperty('LATEST_CSV_VERSION') || '0', 10);

        if (latestTime <= currentVersion) {
            console.log("Already up to date. Latest: " + latestTime);
            return "Already up to date";
        }

        console.log("Newer CSV found: " + latestFile.getName());

        // Read content
        // Try UTF-16LE first as it appeared in error logs, then Shift_JIS, then UTF-8
        const blob = latestFile.getBlob();
        let csvContent = "";
        let detected = false;
        const encodings = ['UTF-16LE', 'Shift_JIS', 'UTF-8'];

        for (const enc of encodings) {
            try {
                const text = blob.getDataAsString(enc);
                // Check for key header to validate encoding
                if (text.includes('顧客ID') || text.includes('Customer')) {
                    csvContent = text;
                    detected = true;
                    console.log("Detected Encoding: " + enc);
                    break;
                }
            } catch (e) { }
        }

        if (!detected) {
            console.log("Encoding detection failed, using default UTF-8");
            csvContent = blob.getDataAsString('UTF-8');
        }

        // Use shared helper - pass raw string!
        const result = updateDatabaseFromLinesV2(csvContent);

        if (result === "Upload Successful") {
            props.setProperty('LATEST_CSV_VERSION', latestTime.toString());
            console.log("Import Success");
            return "Imported: " + latestFile.getName();
        } else {
            console.error("Import Failed: " + result);
            return "Failed: " + result;
        }

    } catch (e) {
        console.error("Auto Import Error: " + e.message);
        return "Error: " + e.message;
    }
}
/**
 * Manually forces the import of the latest CSV, ignoring version checks.
 * Run this function from the GAS Editor to update data immediately.
 */
function forceImportCsv() {
    try {
        // Reset Version
        PropertiesService.getScriptProperties().deleteProperty('LATEST_CSV_VERSION');
        console.log("Version reset. Starting import...");

        // Run standard check (which will now see '0' as current version)
        const result = checkAndImportLatestCsv();
        console.log("Force Import Result: " + result);
        return result;
    } catch (e) {
        console.error("Force Import Error: " + e.message);
        return "Error: " + e.message;
    }
}



/**
 * Shared helper to update database from Parsed CSV Lines (Array of Arrays)
 */
function updateDatabaseFromLines(rawString) {
    if (!rawString) return "Empty Data";

    // Detect Delimiter: Try Tab then Comma
    const firstLine = rawString.substring(0, rawString.indexOf('\n'));
    let delimiter = ',';
    // If tab count > visible comma count, assume tab
    if ((firstLine.match(/\t/g) || []).length > (firstLine.match(/,/g) || []).length) {
        delimiter = '\t';
    }

    const csvData = Utilities.parseCsv(rawString, delimiter);
    if (!csvData || csvData.length < 2) return "Invalid CSV Data";

    const ss = getSpreadsheet();
    let cSheet = ss.getSheetByName(CUSTOMER_DB_NAME);
    let fSheet = ss.getSheetByName(FAMILY_DB_NAME);

    // Ensure sheets exist
    if (!cSheet) cSheet = ss.insertSheet(CUSTOMER_DB_NAME);
    if (!fSheet) fSheet = ss.insertSheet(FAMILY_DB_NAME);

    const cHeader = ['顧客ID', '氏名', '住所 (建物名・部屋番号)'];
    // Check/Init Headers
    if (cSheet.getLastRow() === 0) cSheet.appendRow(cHeader);

    const fHeader = ['顧客ID', '家族氏名', '生年月日', '職業', 'アレルギー情報', '備考'];
    if (fSheet.getLastRow() === 0) fSheet.appendRow(fHeader);

    // Existing Data Map for Updates (only Customer DB)
    const cCurrent = cSheet.getDataRange().getValues();
    const cMap = new Map();
    // Start from 1 (skip header)
    for (let i = 1; i < cCurrent.length; i++) {
        cMap.set(String(cCurrent[i][0]), i + 1); // ID -> RowNum
    }

    const updates = [];
    const newCRows = [];
    // Always append new Family data logic
    const fRows = [];

    // Assuming first row is header
    const headerRow = csvData[0];
    const dataRows = csvData.slice(1);

    // Detect Column Indices 
    const idxId = headerRow.findIndex(h => h.includes('顧客ID'));
    // Handle separate Name columns if necessary
    const idxLastName = headerRow.findIndex(h => h === '姓');
    const idxFirstName = headerRow.findIndex(h => h === '名');
    const idxName = headerRow.findIndex(h => h.includes('氏名') || h.includes('Name')); // Fallback

    const idxAddr = headerRow.findIndex(h => h.includes('住所') || h.includes('地番'));

    // Family column might be "世帯全員の情報" or similar
    const idxFamily = headerRow.findIndex(h => h.includes('世帯') || h.includes('家族') || h.includes('Family'));

    if (idxId === -1) return "Invalid CSV Format (Missing ID)";

    for (const row of dataRows) {
        const id = String(row[idxId]);
        if (!id) continue;

        let name = "";
        if (idxLastName !== -1 && idxFirstName !== -1) {
            name = (row[idxLastName] || "") + " " + (row[idxFirstName] || "");
        } else if (idxName !== -1) {
            name = row[idxName];
        } else {
            // Fallback if no specific name header found, try col 1+2
            if (row.length > 2) name = row[1] + " " + row[2];
        }
        name = name.trim();

        // Address
        let address = "";
        if (idxAddr !== -1 && idxAddr < row.length) {
            address = row[idxAddr];
        } else {
            // Fallback for address?
        }

        const familyRaw = (idxFamily !== -1 && idxFamily < row.length) ? row[idxFamily] : "";

        // Customer Upsert
        if (cMap.has(id)) {
            updates.push({ row: cMap.get(id), values: [id, name, address] });
        } else {
            newCRows.push([id, name, address]);
        }

        // Family Parsing
        if (familyRaw) {
            const blocks = familyRaw.split(/[\r\n]+/).filter(line => line.trim() !== "");
            blocks.forEach(line => {
                const parts = line.trim().split(/[\s　]+/);

                // Heuristic: Find DOB index
                let dobIndex = -1;
                for (let j = 0; j < parts.length; j++) {
                    if (parts[j].match(/\d{4}\/\d{1,2}\/\d{1,2}/)) {
                        dobIndex = j;
                        break;
                    }
                }

                if (dobIndex !== -1) {
                    const fName = parts.slice(0, dobIndex).join(" ");
                    const fDob = parts[dobIndex];
                    const fJob = (dobIndex + 1 < parts.length) ? parts[dobIndex + 1] : "";
                    const fAllergy = (dobIndex + 2 < parts.length) ? parts[dobIndex + 2] : "";
                    const fInfo = (dobIndex + 3 < parts.length) ? parts.slice(dobIndex + 3).join(" ") : "";
                    fRows.push([id, fName, fDob, fJob, fAllergy, fInfo]);
                } else {
                    // Fallback
                    if (parts.length > 0) {
                        const fName = parts[0] + (parts.length > 1 ? " " + parts[1] : "");
                        const fRest = parts.slice(2).join(" ");
                        fRows.push([id, fName, "", "", "", fRest]);
                    }
                }
            });
        }
    }

    // Batch Update Customers
    updates.forEach(u => {
        cSheet.getRange(u.row, 1, 1, 3).setValues([u.values]);
    });
    if (newCRows.length > 0) {
        cSheet.getRange(cSheet.getLastRow() + 1, 1, newCRows.length, 3).setValues(newCRows);
    }

    // Rewrite Families (Filtering old + Adding new)
    if (fRows.length > 0) {
        // Only clear if we have new data to put in? 
        // Or if this is a partial update? Previous app seemed to be "Bulk Upload" style.
        // I will follow the "Bulk Replace" pattern for Family DB to avoid duplicates if that was the logic.
        // Warning: This wipes data not in the CSV.

        // SAFEGUARD: If this is an automatic import, we expect the CSV to be full dump.
        const lastRow = fSheet.getLastRow();
        if (lastRow > 1) {
            fSheet.getRange(2, 1, lastRow - 1, fHeader.length).clearContent();
        }
        fSheet.getRange(2, 1, fRows.length, fHeader.length).setValues(fRows);
    }

    // Update Data Version
    PropertiesService.getScriptProperties().setProperty('DATA_VERSION', String(Date.now()));

    return "Upload Successful";
}

function checkDataVersion() {
    return PropertiesService.getScriptProperties().getProperty('DATA_VERSION') || '0';
}

function testDriveAuth() {
    // Run this function in the GAS Editor to authorize DriveApp permissions (Write Access)
    const folder = DriveApp.getFolderById(RECEIPT_FOLDER_ID);
    console.log("Folder Found: " + folder.getName());

    // Create a dummy file to force Writer Authorization
    const file = folder.createFile("AuthTest.txt", "Write Permission Test");
    console.log("Write Successful: " + file.getName());

    // Clean up
    file.setTrashed(true);
    console.log("Cleanup Successful");
}

function getStaffList() {
    try {
        const ss = SpreadsheetApp.openById(STAFF_SS_ID);
        const sheet = ss.getSheets().find(s => s.getSheetId() == STAFF_GID);
        if (!sheet) return [];

        const lastRow = sheet.getLastRow();
        if (lastRow < 2) return [];

        // Name is in Col A (index 0 for getValues, but getRange is 1-based, so Column 1)
        const range = sheet.getRange(2, 1, lastRow - 1, 1);
        const values = range.getValues();
        return values.map(row => row[0]).filter(name => name);

    } catch (e) {
        console.error("Staff List Error: " + e.message);
        return [];
    }
}

/**
 * V2: Shared helper to update database from Parsed CSV Lines (Array of Arrays)
 * Supports full CSV column import.
 */
function updateDatabaseFromLinesV2(rawString) {
    if (!rawString) return "Empty Data";

    // Detect Delimiter: Try Tab then Comma
    const firstLine = rawString.substring(0, rawString.indexOf('\n'));
    let delimiter = ',';
    // If tab count > visible comma count, assume tab
    if ((firstLine.match(/\t/g) || []).length > (firstLine.match(/,/g) || []).length) {
        delimiter = '\t';
    }

    const csvData = Utilities.parseCsv(rawString, delimiter);
    if (!csvData || csvData.length < 2) return "Invalid CSV Data";

    const ss = getSpreadsheet();
    let cSheet = ss.getSheetByName(CUSTOMER_DB_NAME);
    let fSheet = ss.getSheetByName(FAMILY_DB_NAME);

    // Ensure sheets exist
    if (!cSheet) cSheet = ss.insertSheet(CUSTOMER_DB_NAME);
    if (!fSheet) fSheet = ss.insertSheet(FAMILY_DB_NAME);

    // Full Replace for Customer DB based on Requirement "CSV has all data"
    // We will clear the sheet and rewrite it to match the CSV structure exactly.
    cSheet.clearContents();

    const headerRow = csvData[0];
    const dataRows = csvData.slice(1);

    // Write Headers
    if (headerRow) cSheet.appendRow(headerRow);

    // Write All Data
    if (dataRows.length > 0) {
        cSheet.getRange(2, 1, dataRows.length, headerRow.length).setValues(dataRows);
    }

    // Family DB Handling (Extract from CSV)
    // We still need to parse Family info to populate FAMILY_DB_NAME
    // Family DB Handling (Extract from CSV)
    // We still need to parse Family info to populate FAMILY_DB_NAME
    const idxFamily = headerRow.findIndex(h => h.includes('世帯') || h.includes('家族') || h.includes('Family'));
    const idxId = headerRow.findIndex(h => h.includes('顧客ID'));

    const fHeader = ['顧客ID', '家族氏名', '生年月日', '職業', 'アレルギー情報', '備考'];
    // Reset Family DB as well
    fSheet.clearContents();
    fSheet.appendRow(fHeader);

    const fRows = [];

    if (idxId !== -1 && idxFamily !== -1) {
        for (const row of dataRows) {
            const id = String(row[idxId]);
            if (!id) continue;

            const familyRaw = (idxFamily < row.length) ? row[idxFamily] : "";
            if (familyRaw) {
                const parsedFamilies = parseFamilyInfo(familyRaw);
                parsedFamilies.forEach(f => {
                    // f: { name, dob, info }
                    // Try to extract Allergy from info
                    let allergy = "";
                    let otherInfo = f.info;

                    // Simple heuristic to extract allergy
                    // If info contains "アレルギー", try to extract closest segment
                    // But for now, just putting everything in info is safer, or column alignment
                    // The user prompt had specific "アレルギー:..." or "アレルギーなし"

                    // We will put everything else in "備考" for now, as splitting Job/Allergy is ambiguous without named fields
                    fRows.push([id, f.name, f.dob, "", "", otherInfo]);
                });
            }
        }
    }

    if (fRows.length > 0) {
        fSheet.getRange(2, 1, fRows.length, fHeader.length).setValues(fRows);
    }

    // Update Data Version
    PropertiesService.getScriptProperties().setProperty('DATA_VERSION', String(Date.now()));

    return "Upload Successful";
}

/**
 * Parses raw family text block into structured objects.
 * Handles various formats: YYYY.MM.DD, YYYYMMDD, Japanese Era, etc.
 * Supports multi-line blocks.
 */
function parseFamilyInfo(rawText) {
    const results = [];
    if (!rawText) return results;

    // Normalize newlines
    const lines = rawText.split(/[\r\n]+/);

    // Regex for Dates
    // Supports:
    // 1. YYYY.MM.DD, YYYY/MM/DD, YYYY-MM-DD
    // 2. YYYY年M月D日
    // 3. 和暦 (明治|大正|昭和|平成|令和|M|T|S|H|R)N(.|年)M(.|月)D(日)
    // 4. Compact 8 digit: 19860921
    // Case insensitive flag 'i' is used.
    const reDate = /((?:19|20)\d{2}[\.\/\-]\d{1,2}[\.\/\-]\d{1,2}|(?:明治|大正|昭和|平成|令和|[MTSHR])\.?\s*[0-9元]+[\.\-年]\s*[0-9]+[\.\-月]\s*[0-9]+(?:日|生)?|(?:19|20)\d{6})/gi;

    let current = null;

    const flush = () => {
        if (current && (current.name || current.dob)) {
            // Join info array
            current.info = current.infoList.join(" ");
            delete current.infoList;
            results.push(current);
        }
        current = { name: "", dob: "", infoList: [] };
    };

    // Initialize first person
    current = { name: "", dob: "", infoList: [] };

    lines.forEach(line => {
        let clean = line.trim();
        if (!clean) return;

        // Check for Date
        reDate.lastIndex = 0;
        const match = reDate.exec(clean);

        if (match) {
            const dateStr = match[0];
            const idx = match.index;
            const pre = clean.substring(0, idx).trim();
            const post = clean.substring(idx + dateStr.length).trim();

            // Detect if this is a "Property Line" (e.g. "生年月日: 1990...")
            // or an "Inline Name Line" (e.g. "Name 1990...")
            const isProperty = /生年月日|誕生日|DOB|Date/.test(pre) || pre.endsWith(":") || pre.endsWith("：");

            if (isProperty) {
                // Determine format:
                // If current person already has DOB, this must be a new person (or error, but assume new)
                // UNLESS valid Name was just set previously (e.g. L1: Name, L2: DOB)
                if (current.dob && !current.justStarted) {
                    flush();
                }
                current.dob = normalizeDateStr(dateStr);
                // 'pre' is likely label, ignore. 'post' might be info.
                if (post) current.infoList.push(post);
                current.justStarted = false; // logic handled
            } else {
                // Likely "Name Date Info" format
                // If we already have a partial person with Name only, and this line has NO name (Pre is empty), then it's that person's DOB.
                if (current.name && !current.dob && pre === "") {
                    current.dob = normalizeDateStr(dateStr);
                    if (post) current.infoList.push(post);
                    current.justStarted = false;
                } else {
                    // Otherwise, it's a fully self-contained line OR a new person line
                    // Flush previous if exists
                    if (current.name || current.dob) flush();

                    current.name = pre; // Name is before date
                    current.dob = normalizeDateStr(dateStr);
                    if (post) current.infoList.push(post);
                    current.justStarted = true; // Mark as fresh to capture subsequent info lines
                }
            }
        } else {
            // No Date Found on this line
            // Is it Name or Info?

            // Heuristics for Info
            const infoKeywords = ["職業", "勤務", "園", "学校", "社", "アレルギー", "疾患", "病", "薬", "申請", "検討", "利用", "金額", "備考", "共有", "男児", "女児", "時", "分"];
            const isInfoKey = infoKeywords.some(k => clean.includes(k));
            const isLong = clean.length > 20;

            // Heuristics for Name (Start of new block)
            // If current person is "Complete" (has DOB) and line looks like a Name -> Flush and Start New
            // If current person is Empty, this is Name.

            const isLikelyName = !isInfoKey && !isLong;

            if (current.dob) {
                // Current person has DOB. 
                // If this line looks like a Name, start new person.
                // UNLESS it's just a short remark? e.g. "主婦"
                if (isLikelyName && !["主婦", "夫", "妻", "パート", "学生", "無職", "会社員", "自営業"].includes(clean)) {
                    flush();
                    current.name = clean;
                    current.justStarted = true;
                } else {
                    current.infoList.push(clean);
                }
            } else {
                // No DOB yet.
                if (!current.name) {
                    // Empty person. Assume Name.
                    if (isLikelyName) {
                        current.name = clean;
                        current.justStarted = true;
                    } else {
                        // Weird to have info before name, but maybe partial?
                        // Or general note. Attach to current (which is empty) or previous?
                        // Just stash it.
                        current.infoList.push(clean);
                    }
                } else {
                    // Has Name, No DOB. 
                    // This line is likely intermediate info or multi-line name?
                    // "夫 \n 鎌弥" -> handled by regex? No.
                    // If Name="夫", and this line="鎌弥", merge?
                    // If Line 1 was very short (<3 chars) and "Role-like", maybe append?
                    if (current.name.length < 5 && isLikelyName) {
                        current.name += " " + clean;
                    } else {
                        current.infoList.push(clean);
                    }
                }
            }
        }
    });

    flush(); // Final flush to capture last person

    return results;
}

/**
 * Normalizes date string to YYYY/MM/DD format.
 */
function normalizeDateStr(dateStr) {
    if (!dateStr) return "";

    // Remove '生' suffix if matched
    let str = dateStr.replace(/生$/, '').trim();

    // 1. Compact 8 digit: 19860921
    if (/^\d{8}$/.test(str)) {
        return str.substring(0, 4) + '/' + str.substring(4, 6) + '/' + str.substring(6, 8);
    }

    // 2. Japanese Era (Kanji or Alpha, Dot or Kanji Separator)
    // Matches: S59.5.9, 昭和59年5月9日, R5.12.21
    const eraMatch = str.match(/^([明治大正昭和平成令和MTSHR])\.?\s*([0-9元]+)[\.\-年]\s*([0-9]+)[\.\-月]\s*([0-9]+)(?:日)?$/i);
    if (eraMatch) {
        let era = eraMatch[1];
        let year = (eraMatch[2] === '元') ? 1 : parseInt(eraMatch[2], 10);
        const month = parseInt(eraMatch[3], 10);
        const day = parseInt(eraMatch[4], 10);

        // Normalize alpha to Kanji
        if (/m/i.test(era)) era = '明治';
        else if (/t/i.test(era)) era = '大正';
        else if (/s/i.test(era)) era = '昭和';
        else if (/h/i.test(era)) era = '平成';
        else if (/r/i.test(era)) era = '令和';

        if (era === '明治') year += 1867;
        else if (era === '大正') year += 1911;
        else if (era === '昭和') year += 1925;
        else if (era === '平成') year += 1988;
        else if (era === '令和') year += 2018;

        return `${year}/${month}/${day}`;
    }

    // 3. Standard YYYY.MM.DD or YYYY-MM-DD
    // Replace '年' '月' '.' '-' with '/'
    let norm = str
        .replace(/[年月\.\-]/g, '/')
        .replace(/日/g, '');

    return norm;
}


// --- Assessment Definitions ---
const ASSESSMENT_DEFINITIONS = {
    risk: {
        title: "産後うつ・児童虐待総合評価",
        levels: [
            { score: 5, label: "安心・良好", desc: "全く懸念がない状態。\n保護者の表情も明るく、お子様も衛生・情緒ともに安定している。\n部屋も安全に保たれている。" },
            { score: 4, label: "通常", desc: "一般的な家庭の状態。\n多少の疲れや散らかりはあるが、保育に支障はなく、親子の関わりも標準的。" },
            { score: 3, label: "要観察", desc: "「少し気になる」レベル。\n保護者がひどく疲れている、部屋が不衛生になりつつある、子供の情緒が少し不安定など。\n※次回の担当者に引き継ぎたい内容がある。" },
            { score: 2, label: "注意", desc: "明らかに異変を感じる状態。\n保護者の反応が鈍い（無視・無表情）、子供の体や服が著しく汚れている、怒鳴り声が多いなど。\n※管理者への報告を強く推奨。" },
            { score: 1, label: "危険・緊急", desc: "緊急の介入が必要な状態。\n明らかな虐待の痕跡（あざ・傷）、育児放棄（ネグレクト）、保護者の心身耗弱が激しく子供の安全が守れない。\n※直ちに管理者に電話連絡が必要。" }
        ]
    },
    es: {
        title: "従業員満足度(ES)",
        levels: [
            { score: 5, label: "最高", desc: "ぜひまた担当したい（優先希望）。\n顧客の態度が非常に良く、感謝されており、環境も快適。\n精神的にも報酬以上のやりがいを感じる。" },
            { score: 4, label: "良", desc: "問題なく担当できる。\n常識的な対応をしていただき、業務遂行にストレスがない。\n標準的な「良いお客様」。" },
            { score: 3, label: "可", desc: "担当しても良い（許容範囲）。\n多少のやりにくさ（細かい指示や部屋の環境など）はあるが、仕事として割り切れる範囲。" },
            { score: 2, label: "難あり", desc: "できれば担当したくない（回避希望）。\n高圧的な態度、契約外の要求が多い、部屋が極端に不衛生などで、精神的・体力的に消耗が激しい。" },
            { score: 1, label: "NG", desc: "二度と担当できない（ブラック）。\nハラスメント（暴言・セクハラ）、身の危険を感じる、著しい契約違反など。\n※担当を外れることを希望するレベル。" }
        ]
    }
};

/**
 * Retrieves UI configuration (placeholders, hints) from the prompt sheet.
 */
function getUiConfig() {
    return {
        dailyPlaceholder: getPrompt(PROMPT_KEYS.PLACEHOLDER_DAILY),
        accidentPlaceholder: getPrompt(PROMPT_KEYS.PLACEHOLDER_ACCIDENT),
        accidentHint: getPrompt(PROMPT_KEYS.HINT_ACCIDENT),
        hiyariPlaceholder: getPrompt(PROMPT_KEYS.PLACEHOLDER_HIYARI),
        assessments: ASSESSMENT_DEFINITIONS
    };
}

/**
 * Retrieves past reports for a specific customer.
 * Returns both Daily and Accident reports, sorted by date (newest first).
 */
function getCustomerReports(customerId, startAfterTime) {
    const limit = 5;

    // --- Firestore Read (Skipped) ---

    // --- Sheet Fallback ---

    // NOTE: This approach reads all data then filters. 
    // For much larger datasets, we would need to read only the bottom N rows.
    // However, given user request for "Load More" and speed, sending small chunks is the main win here.

    const ss = getSpreadsheet();
    const results = [];
    const fmt = (d) => {
        if (!d) return "";
        if (d instanceof Date) return Utilities.formatDate(d, "Asia/Tokyo", "yyyy/MM/dd HH:mm");
        return String(d);
    };

    // Helper to scan sheet backwards
    const scanSheet = (sheetName, mapFn) => {
        const sheet = ss.getSheetByName(sheetName);
        if (!sheet) return;
        const lastRow = sheet.getLastRow();
        if (lastRow < 2) return;

        // Optimization Idea: If we knew exactly where the time limit was, we could range read.
        // For now, read all is safest for correctness with mixed report types.
        const values = sheet.getDataRange().getValues();

        // Loop backwards
        for (let i = values.length - 1; i >= 1; i--) {
            const row = values[i];
            const res = mapFn(row);
            if (res) results.push(res);
        }
    };

    // Daily Reports
    scanSheet(REPORT_SHEET_NAME, (row) => {
        if (String(row[4]) !== String(customerId)) return null;
        return {
            type: 'daily',
            timestamp: fmt(row[0]),
            staff: row[3],
            original: row[6],
            internal: row[7],
            customer: row[8],
            risk: row[9],
            es: row[10]
        };
    });

    // Accident Reports
    scanSheet(ACCIDENT_SHEET_NAME || '事故報告', (row) => {
        if (String(row[2]) !== String(customerId)) return null;

        const parts = [];
        if (row[6]) parts.push(`【発生時間】${row[6]}`);
        if (row[7]) parts.push(`【場所】${row[7]}`);
        if (row[9]) parts.push(`【状況】\n${row[9]}`);
        if (row[8]) parts.push(`【事故内容】\n${row[8]}`);
        if (row[10]) parts.push(`【応急処置】\n${row[10]}`);
        if (row[12]) parts.push(`【受診・治療】\n${row[12]}`);
        if (row[13]) parts.push(`【再発防止策】\n${row[13]}`);

        const internalText = parts.join('\n\n');

        return {
            type: 'accident',
            timestamp: fmt(row[0]),
            staff: row[1],
            original: row[14],
            internal: internalText,
            customer: row[11],
            isAccident: true,
            subtype: row[15] || '事故報告'
        };
    });

    // Sort by date descending
    results.sort((a, b) => {
        const da = new Date(a.timestamp);
        const db = new Date(b.timestamp);
        return db - da; // Descending
    });

    // Pagination Logic
    let startIndex = 0;
    if (startAfterTime) {
        // Find the index of the item that matches startAfterTime
        // And start from the next one.
        // Since timestamps might be non-unique, strictly finding the object is hard without distinct ID.
        // We will assume checking timestamp < startAfterTime is enough for "next page".

        const startTime = new Date(startAfterTime).getTime();

        // Find first item that is strictly older than startAfterTime
        // (Since sorted descending, we look for timestamp < startTime)
        // However, if we have duplicate timestamps, this skips all of them.
        // A simple array scan findIndex is better if we assume results are consistent.
        // But since sheet might have been modified, filtering by time is more robust.

        const filtered = results.filter(r => new Date(r.timestamp).getTime() < startTime);
        return filtered.slice(0, limit);
    }

    return results.slice(0, limit);
}

/**
 * One-time migration function to copy data from Sheet to Firestore.
 * Run this manually from GAS Editor.
 */
function migrateDataToFirestore() {
    const firestore = getFirestore();
    if (!firestore) {
        console.error("Firestore not configured. Please checks Script Properties.");
        return "Firestore Not Configured";
    }

    const ss = getSpreadsheet();
    let migratedCount = 0;

    // 1. Daily Reports
    const dailySheet = ss.getSheetByName(REPORT_SHEET_NAME);
    if (dailySheet && dailySheet.getLastRow() > 1) {
        const rows = dailySheet.getDataRange().getValues().slice(1);
        rows.forEach(row => {
            try {
                if (!row[4]) return; // Skip if no Customer ID

                const timeStr = (row[0] instanceof Date) ? Utilities.formatDate(row[0], "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss") : String(row[0]);
                // Create deterministic ID to prevent duplicates on re-run
                // ID: D_CustomerID_Timestamp(NumbersOnly)
                const docId = `D_${String(row[4])}_${timeStr.replace(/[^0-9]/g, '')}`;

                const docData = {
                    timestamp: timeStr,
                    start: String(row[1]),
                    end: String(row[2]),
                    staffName: String(row[3]),
                    customerId: String(row[4]),
                    customerName: String(row[5]),
                    inputText: String(row[6]),
                    internalText: String(row[7]),
                    customerText: String(row[8]),
                    riskRating: Number(row[9]) || 0,
                    esRating: Number(row[10]) || 0,
                    type: 'daily',
                    migratedAt: new Date()
                };

                // Use ID in path to ensure Upsert behavior (overwrite if exists with this ID)
                firestore.createDocument("reports/" + docId, docData);
                migratedCount++;
            } catch (e) {
                console.warn("Skipped Daily Row: " + e.message);
            }
        });
    }

    // 2. Accident Reports
    const accSheet = ss.getSheetByName(ACCIDENT_SHEET_NAME || '事故報告');
    if (accSheet && accSheet.getLastRow() > 1) {
        const rows = accSheet.getDataRange().getValues().slice(1);
        rows.forEach(row => {
            try {
                if (!row[2]) return; // CustomerID

                const timeStr = (row[0] instanceof Date) ? Utilities.formatDate(row[0], "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss") : String(row[0]);
                // ID: A_CustomerID_Timestamp(NumbersOnly)
                const docId = `A_${String(row[2])}_${timeStr.replace(/[^0-9]/g, '')}`;

                const docData = {
                    timestamp: timeStr,
                    staffName: String(row[1]),
                    customerId: String(row[2]),
                    customerName: String(row[3]),
                    targetName: String(row[4]),
                    targetDob: (row[5] instanceof Date) ? Utilities.formatDate(row[5], "Asia/Tokyo", "yyyy/MM/dd") : String(row[5]),
                    occurrenceTime: String(row[6]),
                    location: String(row[7]),
                    accidentContent: String(row[8]),
                    situation: String(row[9]),
                    immediateResponse: String(row[10]),
                    parentCorrespondence: String(row[11]),
                    diagnosisTreatment: String(row[12]),
                    prevention: String(row[13]),
                    inputText: String(row[14] || ""),
                    reportType: String(row[15] || "事故報告"),
                    type: 'accident',
                    migratedAt: new Date()
                };

                firestore.createDocument("reports/" + docId, docData);
                migratedCount++;
            } catch (e) {
                console.warn("Skipped Accident Row: " + e.message);
            }
        });
    }

    console.log(`Migration Completed. Total documents processed: ${migratedCount}`);
    return `Migration Completed. ${migratedCount} records processed.`;
}
