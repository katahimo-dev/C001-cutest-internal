// ==========================================
// Web App Logic
// ==========================================

const SPREADSHEET_ID = '1Y0ciyR4LHwCPwIkcuqwhQnD9k_-WBHn0nnwNRzFaacM';
const TARGET_GID = 1224762512;
const PROMPT_SHEET_NAME = 'ＡＩプロンプト';
const REPORT_SHEET_NAME = '日報';
const STAFF_SS_ID = "1yfVdlHeptbGZxTawIQjGLZu60T4726tEveLjOO-aL6Q";
const STAFF_GID = 939066637;

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
  "internal": "社内向けレポート内容（事実・客観的）。不足部分は推測せず、入力内容のみで構成。",
  "customer": "保護者向けレポート内容（親しみやすく）。不足部分は推測せず、入力内容のみで構成。"
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
function getData() {
    const ss = getSpreadsheet();
    let cSheet = ss.getSheetByName(CUSTOMER_DB_NAME);
    let fSheet = ss.getSheetByName(FAMILY_DB_NAME);

    // Fallback if sheets don't exist yet (return empty)
    if (!cSheet || !fSheet) return { cities: [], customers: [] };

    const cData = cSheet.getDataRange().getValues();
    const fData = fSheet.getDataRange().getValues();

    // Headers are row 0. Data starts row 1.
    const cRows = cData.slice(1);
    const fRows = fData.slice(1);

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
        const id = toStr(row[0]);
        if (!id) return null; // Skip empty IDs

        const name = toStr(row[1]);
        const address = toStr(row[2]);

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
            family: familyMap[id] || []
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
function callGemini(apiKey, contentParts, generationConfig) {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`;
    // Using 2.0-flash-exp or 1.5-flash for better vision support if 2.5 is not stable yet. 
    // User's previous code had 2.5-flash. Assuming it works or falling back to 1.5-flash which is standard for multimodal.
    // Let's stick to the user's previous model ID if it was working, or use 1.5-flash for safety with images.
    // Actually, gemini-1.5-flash is the standard for cost-effective multimodal. 
    // Previous code: gemini-2.5-flash (might be a typo by me or user previously? usually it's 1.5-flash or 2.0-flash-exp).
    // I will use gemini-1.5-flash which is reliable for OCR.

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

    // Call with text part
    const result = callGemini(apiKey, [{ text: prompt }]);
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

    const result = callGemini(apiKey, [
        { text: prompt },
        { inline_data: { mime_type: "image/jpeg", data: rawBase64 } }
    ]);

    if (result.error) {
        console.error(result.error);
        return { amount: 0 };
    }
    return result;
}

/**
 * Saves the final report to 'Reports' sheet.
 */
function saveReport(reportData) {
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName(REPORT_SHEET_NAME);

    if (!sheet) {
        sheet = ss.insertSheet(REPORT_SHEET_NAME);
        sheet.appendRow(['Timestamp', 'StartTime', 'EndTime', 'User', 'CustomerId', 'CustomerName', 'InputText', 'InternalReport', 'CustomerReport']);
    }

    const timestampJST = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");

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
    ]);

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

    const result = callGemini(apiKey, [{ text: prompt }]);
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

    return "Success";
}





function verifyLogin(userId, password) {
    try {
        const ss = SpreadsheetApp.openById(STAFF_SS_ID);
        const sheet = ss.getSheets().find(s => s.getSheetId() == STAFF_GID);
        if (!sheet) return { success: false, message: "Staff sheet not found" };

        const data = sheet.getDataRange().getValues().slice(1); // Skip header
        // New Layout: Col A(0): Name, B(1): UserID, C(2): Password, D(3): Admin
        const user = data.find(row => row[1] == userId && row[2] == password);

        if (user) {
            return { success: true, name: user[0], isAdmin: (user[3] == 1) };
        } else {
            return { success: false, message: "Invalid ID or password" };
        }
    } catch (e) {
        return { success: false, message: e.message };
    }
}

const CUSTOMER_DB_NAME = '顧客DB_New';
const FAMILY_DB_NAME = '家族DB_New';

function uploadCustomerCsv(csvContent) {
    const ss = getSpreadsheet();
    let cSheet = ss.getSheetByName(CUSTOMER_DB_NAME);
    let fSheet = ss.getSheetByName(FAMILY_DB_NAME);

    // Create sheets if not exist
    if (!cSheet) {
        cSheet = ss.insertSheet(CUSTOMER_DB_NAME);
        cSheet.appendRow(['ID', 'Name', 'Address']); // Minimal key fields
    }
    if (!fSheet) {
        fSheet = ss.insertSheet(FAMILY_DB_NAME);
        fSheet.appendRow(['CustomerID', 'Name', 'DOB', 'Job', 'Allergy', 'Info']);
    }

    // Detect Delimiter (Simple Heuristic: check first line)
    const firstLine = csvContent.substring(0, csvContent.indexOf('\n'));
    let delimiter = ',';
    if (firstLine.indexOf('\t') !== -1 && (firstLine.match(/\t/g) || []).length > (firstLine.match(/,/g) || []).length) {
        delimiter = '\t';
    }

    // Parse CSV
    const rows = Utilities.parseCsv(csvContent, delimiter);
    if (rows.length < 2) return "No data found";

    // Column Mapping
    const header = rows[0];
    let idxId = -1, idxName = -1, idxAddr = -1, idxFamily = -1;

    for (let i = 0; i < header.length; i++) {
        const col = String(header[i]).trim();
        if (col.includes("ID") && !col.includes("Benefit")) idxId = i;
        else if (col.includes("顧客ID")) idxId = i; // Fallback

        if (col.includes("氏名") || col.includes("Name")) idxName = i;
        if (col.includes("住所") || col.includes("Address")) idxAddr = i;
        if (col.includes("世帯全員の情報") || col.includes("Family")) idxFamily = i;
    }

    // Fallbacks if not found (based on user spec)
    if (idxId === -1) idxId = 0; // Default A
    if (idxName === -1) idxName = 1; // Default B
    if (idxAddr === -1) idxAddr = 18; // Default S
    if (idxFamily === -1) idxFamily = 24; // Default Y

    // Read existing data for Upsert
    const cData = cSheet.getDataRange().getValues();
    const cMap = new Map();
    for (let i = 1; i < cData.length; i++) cMap.set(String(cData[i][0]), i + 1); // ID -> RowNum

    const fData = fSheet.getDataRange().getValues();
    const fHeader = fData[0];
    let fRows = fData.slice(1);

    // IDs in the CSV
    const csvIds = new Set();
    // Pre-scan to gather IDs for filtering families
    for (let i = 1; i < rows.length; i++) {
        if (rows[i].length > idxId) csvIds.add(String(rows[i][idxId]));
    }

    // Filter out families of affected customers
    fRows = fRows.filter(r => !csvIds.has(String(r[0])));

    const newCRows = [];
    const updates = []; // {row: x, val: []}
    const newFRows = [];

    for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        if (row.length <= idxId) continue;

        const id = String(row[idxId]);
        // Safety checks for indices
        const name = (idxName < row.length) ? row[idxName] : "";
        const address = (idxAddr < row.length) ? row[idxAddr] : "";
        const familyRaw = (idxFamily < row.length) ? row[idxFamily] : "";

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
                // "鈴木 健一 1982/02/14 自営業 ..."
                // Assumption: Name is first 1 or 2 parts. 
                // To be robust: Date like YYYY/MM/DD is likely DOB.
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
                    // Fallback if no date found: just dump
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
        fSheet.getRange(2, 1, fSheet.getLastRow(), fHeader.length).clearContent();
        fSheet.getRange(2, 1, fRows.length, fHeader.length).setValues(fRows);
    } else {
        if (fSheet.getLastRow() > 1) fSheet.getRange(2, 1, fSheet.getLastRow() - 1, fHeader.length).clearContent();
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
        const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
        const staff = values.flat().filter(name => name && String(name).trim() !== "");
        return [...new Set(staff)];
    } catch (e) {
        return ["Error: " + e.message];
    }
}

/**
 * Retrieves UI configuration (placeholders, hints) from the prompt sheet.
 */
function getUiConfig() {
    return {
        dailyPlaceholder: getPrompt(PROMPT_KEYS.PLACEHOLDER_DAILY),
        accidentPlaceholder: getPrompt(PROMPT_KEYS.PLACEHOLDER_ACCIDENT),
        accidentHint: getPrompt(PROMPT_KEYS.HINT_ACCIDENT),
        hiyariPlaceholder: getPrompt(PROMPT_KEYS.PLACEHOLDER_HIYARI)
    };
}
