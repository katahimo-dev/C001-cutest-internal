// ==========================================
// Web App Logic
// ==========================================

const SPREADSHEET_ID = '1Y0ciyR4LHwCPwIkcuqwhQnD9k_-WBHn0nnwNRzFaacM';
const TARGET_GID = 1224762512;
const PROMPT_SHEET_NAME = 'ＡＩプロンプト';
const REPORT_SHEET_NAME = '日報';
const STAFF_SS_ID = "1yfVdlHeptbGZxTawIQjGLZu60T4726tEveLjOO-aL6Q";
const STAFF_GID = 939066637;
const PROMPT_KEYS = {
    GENERATE_WITH_WARNINGS: 'GenerateWithWarnings'
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
`
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
    const sheet = getSheetByGid(ss, TARGET_GID);

    // If sheet doesn't exist, return empty
    if (!sheet) return { cities: [], customers: [] };

    const data = sheet.getDataRange().getValues();
    // Assuming row 1 is header, data starts from row 2
    const rows = data.slice(1);

    const customers = rows.map((row, index) => {
        // B column is index 1 (Surname)
        // C column is index 2 (Given Name)
        // S column is index 18 (Address)
        const customerId = row[0];
        const lastName = row[1];
        const firstName = row[2];
        const address = row[18];

        if (!lastName) return null;

        const name = `${lastName} ${firstName}`;

        // Extract City: "〇〇市"
        let city = "その他";
        if (address) {
            // Match anything ending in 市.
            // Often addresses are "State City ...". We want the "City".
            // Simplest heuristic: grab the first occurrence of "...市"
            // Or look for 県...市?
            // User example: "〇〇市" suggests we just want the City part.
            // Let's crudely regex for `(\S+市)` or just look for '市' index.
            // Better: `(東京都|北海道|...)(...市|...区|...郡)`?
            // Simple approach: Match /(.+?市)/
            const cityMatch = address.match(/(.+?市)/);
            if (cityMatch) {
                // If address starts with Prefecture, keep it? Or just the city?
                // Usually "XX市" implies just the city name?
                // "住所から〇〇市を抜き出して" -> Extract "City".
                // We will use the match.
                city = cityMatch[1];

                // Cleanup: sometimes it includes prefecture. 
                // If it's too long, maybe trim? But for filter "Yokohama-shi" is fine.
                // We'll leave it as the extracted chunk ending in 市.
            }
        }

        return {
            id: customerId, // Use the ID from Column A
            name: name,
            address: address,
            city: city
        };
    }).filter(c => c !== null);

    // Unique cities for filter
    const cities = [...new Set(customers.map(c => c.city))].sort();

    return { cities, customers };
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
function generateReportWithWarnings(inputData) {
    const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!apiKey) return { warnings: ["API Key Missing"], internal: "Error: API Key not set in Script Properties", customer: "Error: API Key not set" };

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

    return callGemini(apiKey, prompt);
}

// ... existing code ...
// ... existing code ...
function callGemini(apiKey, promptText) {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`;
    const payload = {
        contents: [{ parts: [{ text: promptText }] }],
        generationConfig: { responseMimeType: "application/json" }
    };
    // ... existing code ...
    // ... existing code ...

    try {
        const response = UrlFetchApp.fetch(url, {
            method: 'post',
            contentType: 'application/json',
            payload: JSON.stringify(payload),
            muteHttpExceptions: true
        });

        if (response.getResponseCode() !== 200) {
            // Fallback checking?
            return { warnings: ["API Error"], internal: "Error: " + response.getContentText(), customer: "" };
        }

        const json = JSON.parse(response.getContentText());
        const candidates = json.candidates;
        if (!candidates || candidates.length === 0) {
            return { warnings: ["No Content"], internal: "Error: No response candidates", customer: "" };
        }

        const text = candidates[0].content.parts[0].text;
        // Strip markdown code block if present
        const jsonText = text.replace(/```json/g, '').replace(/```/g, '').trim();
        return JSON.parse(jsonText);
    } catch (e) {
        return { warnings: ["System Error"], internal: "Error: " + e.message, customer: "" };
    }
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

    return "Success";
}



function getStaffList() {
    try {
        const ss = SpreadsheetApp.openById(STAFF_SS_ID);
        const sheet = ss.getSheets().find(s => s.getSheetId() == STAFF_GID);
        if (!sheet) return [];

        const lastRow = sheet.getLastRow();
        if (lastRow < 2) return [];

        const values = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
        const staff = values.flat().filter(name => name && String(name).trim() !== "");
        return [...new Set(staff)];
    } catch (e) {
        return ["Error: " + e.message];
    }
}
