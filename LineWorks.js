/**
 * LineWorks API 2.0 Integration
 */

// Script Properties Keys
const LW_PROPS = {
    CLIENT_ID: 'LW_CLIENT_ID',
    CLIENT_SECRET: 'LW_CLIENT_SECRET',
    SERVICE_ACCOUNT: 'LW_SERVICE_ACCOUNT',
    PRIVATE_KEY: 'LW_PRIVATE_KEY', // PEM format
    BOT_ID: 'LW_BOT_ID',
    DOMAIN_ID: 'LW_DOMAIN_ID',
    TARGET_ID: 'LW_TARGET_ID' // New property for Channel or User ID
};

const DEFAULT_TARGET_ID = 'cu.45082@cutestjapan';

/**
 * Gets the configured Target ID (Channel or User), or default.
 */
function getLwTargetId() {
    const props = PropertiesService.getScriptProperties();
    let id = props.getProperty(LW_PROPS.TARGET_ID);
    // Backward compatibility: check old channel ID if new one missing
    if (!id) id = props.getProperty('LW_CHANNEL_ID');

    return id || DEFAULT_TARGET_ID;
}

/**
 * Saves the Global Target ID. 
 * Only admins should call this (check on frontend or separate verification).
 */
function saveLwTargetId(newId) {
    if (!newId) return "ID cannot be empty";
    PropertiesService.getScriptProperties().setProperty(LW_PROPS.TARGET_ID, newId.trim());
    return { success: true, message: "通知先IDを保存しました" };
}

/**
 * Sends a text message to LineWorks.
 * Currently configured to send to a specific channel/user defined in properties or arguments.
 */
function sendToLineWorks(text) {
    try {
        const token = getLwAccessToken();
        if (!token) {
            console.error("LineWorks: Failed to get Access Token");
            return;
        }

        const props = PropertiesService.getScriptProperties();
        const botId = props.getProperty(LW_PROPS.BOT_ID);

        // Get configured Target ID
        let targetId = getLwTargetId();

        // Auto-detect type: If it has '@', assume User ID. Otherwise assume Channel ID.
        let targetType = 'channels';
        if (targetId.includes('@')) {
            targetType = 'users';
        }

        if (!targetId) {
            console.error("LineWorks: No Target ID configured.");
            return;
        }

        if (!botId) {
            console.error("LineWorks: Bot ID missing.");
            return;
        }

        // URL structure differs: 
        // /bots/{botId}/channels/{channelId}/messages
        // /bots/{botId}/users/{userId}/messages
        const url = `https://www.worksapis.com/v1.0/bots/${botId}/${targetType}/${targetId}/messages`;

        const payload = {
            content: {
                type: "text",
                text: text
            }
        };

        const options = {
            method: 'post',
            headers: {
                'Authorization': 'Bearer ' + token,
                'Content-Type': 'application/json'
            },
            payload: JSON.stringify(payload),
            muteHttpExceptions: true
        };

        const res = UrlFetchApp.fetch(url, options);
        if (res.getResponseCode() !== 201) {
            console.error("LineWorks Send Failed (" + targetType + "): " + res.getContentText());
        } else {
            console.log("LineWorks Message Sent to " + targetType + ":" + targetId);
        }

    } catch (e) {
        console.error("LineWorks Error: " + e.message);
    }
}

/**
 * Gets a new Access Token using JWT (OAuth 2.0 Service Account)
 */
function getLwAccessToken() {
    const props = PropertiesService.getScriptProperties();
    const clientId = props.getProperty(LW_PROPS.CLIENT_ID)?.trim();
    const clientSecret = props.getProperty(LW_PROPS.CLIENT_SECRET)?.trim();
    const serviceAccount = props.getProperty(LW_PROPS.SERVICE_ACCOUNT)?.trim();

    // 1. Try Script Property
    let privateKey = props.getProperty(LW_PROPS.PRIVATE_KEY);

    // 2. If Property is missing or short, try Drive File
    if (!privateKey || privateKey.length < 50) {
        // Updated standard filename
        const keyFileName = 'lineworks_private.key';
        console.log(`Checking for ${keyFileName} in Drive...`);
        privateKey = getPrivateKeyFromDrive(keyFileName);
    }

    if (!clientId || !clientSecret || !serviceAccount || !privateKey) {
        console.error("LineWorks Credentials Missing or Empty");
        console.log(`Debug Info: ClientID=${clientId ? 'OK' : 'Missing'}, Secret=${clientSecret ? 'OK' : 'Missing'}, SA=${serviceAccount ? 'OK' : 'Missing'}, Key=${privateKey ? 'OK' : 'Missing'}`);
        return null;
    }

    // Robust Private Key Formatting
    privateKey = formatPrivateKey(privateKey);

    // Create JWT
    const jwt = createLwJwt(clientId, serviceAccount, privateKey);

    // Request Access Token
    const url = 'https://auth.worksmobile.com/oauth2/v2.0/token';
    const payload = {
        grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
        assertion: jwt,
        client_id: clientId,
        client_secret: clientSecret,
        scope: 'bot' // Minimal scope needed for bot messages
    };

    const options = {
        method: 'post',
        payload: payload,
        muteHttpExceptions: true
    };

    const res = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(res.getContentText());

    if (json.access_token) {
        return json.access_token;
    } else {
        console.error("LineWorks Token Error: " + res.getContentText());
        return null;
    }
}

/**
 * Helper to ensure Private Key is in correct PEM format for GAS
 * Re-constructs the key from scratch to fix common pasting errors (single line, missing newlines, etc.)
 */
function formatPrivateKey(key) {
    if (!key) return "";

    // 1. Remove all headers, footers (using regex to catch typos like 4 dashes), and whitespace
    let body = key;
    // Remove literal "\n" characters if they exist (common in JSON strings)
    body = body.replace(/\\n/g, '');
    // Remove headers/footers loosely
    body = body.replace(/-----BEGIN[^-]+-----/g, '');
    body = body.replace(/-----END[^-]+-----/g, '');
    // Remove all remaining whitespace
    body = body.replace(/\s+/g, '');

    // 2. Split into 64-character chunks (Standard PEM requirement)
    const chunks = body.match(/.{1,64}/g);
    if (!chunks) {
        console.error("Failed to parse Private Key body.");
        return key;
    }
    const chunkedBody = chunks.join('\n');

    // 3. Re-assemble with correct headers and newlines
    return '-----BEGIN PRIVATE KEY-----\n' + chunkedBody + '\n-----END PRIVATE KEY-----';
}

/**
 * Searches for a file in Google Drive and reads its content.
 */
function getPrivateKeyFromDrive(fileName) {
    try {
        const files = DriveApp.getFilesByName(fileName);
        if (files.hasNext()) {
            const file = files.next();
            console.log("Found Private Key File: " + file.getName());
            return file.getBlob().getDataAsString();
        }
    } catch (e) {
        console.warn("Drive Search Error: " + e.message);
    }
    return null;
}

/**
 * Creates a signed JWT for LineWorks
 */
function createLwJwt(clientId, serviceAccount, privateKey) {
    const header = {
        alg: 'RS256',
        typ: 'JWT'
    };

    const now = Math.floor(Date.now() / 1000);
    const claim = {
        iss: clientId,
        sub: serviceAccount,
        iat: now,
        exp: now + 3600 // 1 hour
    };

    console.log("Debug: JWT Claim Payload -> " + JSON.stringify(claim));

    const toSign = Utilities.base64EncodeWebSafe(JSON.stringify(header)) + '.' +
        Utilities.base64EncodeWebSafe(JSON.stringify(claim));

    // Check if key is empty or corrupted
    if (!privateKey || privateKey.length < 50) {
        console.error("Debug: Private Key seems too short or empty inside createLwJwt");
    }

    const signature = Utilities.computeRsaSha256Signature(toSign, privateKey);

    return toSign + '.' + Utilities.base64EncodeWebSafe(signature);
}

/**
 * DEBUG HELPER: Run this function to see a list of channels the Bot is in.
 * Use this to find your 'channelId'.
 */
function listBotChannels() {
    const token = getLwAccessToken();
    if (!token) return;

    const props = PropertiesService.getScriptProperties();
    const botId = props.getProperty(LW_PROPS.BOT_ID);

    console.log("Access Token acquired successfully.");
    console.log("Please verify `LW_CHANNEL_ID` is set in Script Properties.");
    console.log("If you do not know the Channel ID, you can get it by adding the Bot to a room, then checking the audit logs or using a callback server (which is complex for GAS).");
    console.log("Easier method: If you just created the Bot, try to message a User ID directly first to test.");
}

/**
 * DEBUG: Run this function manually to test the LineWorks connection.
 * It will log each step so you can see where it fails.
 */
function testLineWorksConnection() {
    console.log("=== Starting LineWorks Connection Test ===");

    const props = PropertiesService.getScriptProperties();
    const keys = [
        'LW_CLIENT_ID',
        'LW_CLIENT_SECRET',
        'LW_SERVICE_ACCOUNT',
        'LW_BOT_ID'
    ];

    // 1. Check Properties
    let missing = [];
    keys.forEach(k => {
        if (!props.getProperty(k)) missing.push(k);
    });

    if (missing.length > 0) {
        console.error("❌ Missing Properties: " + missing.join(', '));
        return;
    }
    console.log("✅ Basic properties present.");

    // 1.5 Check Private Key (Property OR File)
    let privKey = props.getProperty('LW_PRIVATE_KEY');
    const keyFileName = 'lineworks_private.key';

    if (!privKey) {
        console.log(`ℹ️ No LW_PRIVATE_KEY property. Checking Drive for ${keyFileName}...`);
        const driveKey = getPrivateKeyFromDrive(keyFileName);
        if (driveKey) {
            console.log("✅ Found Private Key in Drive.");
            privKey = driveKey;
        } else {
            console.error(`❌ Private Key missing in both Property and Drive (${keyFileName}).`);
            console.error("💡 Please upload your Private Key file to Google Drive with the name: " + keyFileName);
            return;
        }
    } else {
        console.log("✅ Found Private Key in Property.");
    }

    // 2. Check Private Key Format
    if (!privKey.includes('BEGIN PRIVATE KEY')) {
        console.warn("⚠️ Private Key might need formatting. System will attempt to fix.");
    } else {
        console.log("✅ Private Key format check passed (basic).");
    }

    // 3. Try to get Access Token
    console.log("🔄 Attempting to get Access Token...");
    console.log(`Debug: Using ClientID: ${props.getProperty(LW_PROPS.CLIENT_ID)?.trim()}`);
    console.log(`Debug: Using ServiceAccount: ${props.getProperty(LW_PROPS.SERVICE_ACCOUNT)?.trim()}`);
    // Note: Do not log the full private key for security, but maybe the first few chars
    console.log(`Debug: Private Key starts with: ${privKey.substring(0, 30).replace(/\n/g, ' ')}...`);

    const token = getLwAccessToken();

    if (!token) {
        console.error("❌ Failed to get Access Token. Check Client ID, Secret, Service Account, and Private Key.");
        return;
    }
    console.log("✅ Access Token acquired successfully!");

    // 4. Send Test Message
    console.log("🔄 Attempting to send test message...");

    try {
        const botId = props.getProperty('LW_BOT_ID');
        let targetId = props.getProperty('LW_CHANNEL_ID');
        let targetType = 'channels';

        if (!targetId) {
            console.log("ℹ️ No Channel ID set. Using Debug User ID.");
            targetId = 'cu.45082@cutestjapan';
            targetType = 'users';
        }

        console.log(`Sending to ${targetType}: ${targetId}`);

        const url = `https://www.worksapis.com/v1.0/bots/${botId}/${targetType}/${targetId}/messages`;
        const payload = {
            content: {
                type: "text",
                text: "LineWorks Connection Test Successful! ✅\n" + new Date().toString()
            }
        };

        const options = {
            method: 'post',
            headers: {
                'Authorization': 'Bearer ' + token,
                'Content-Type': 'application/json'
            },
            payload: JSON.stringify(payload),
            muteHttpExceptions: true
        };

        const res = UrlFetchApp.fetch(url, options);
        const code = res.getResponseCode();
        const content = res.getContentText();

        if (code === 201) {
            console.log("✅ Message sent successfully!");
        } else {
            console.error(`❌ Send Failed. Code: ${code}`);
            console.error(`Response: ${content}`);

            if (code === 400 || code === 404) {
                console.error("Hint: Check if Bot ID is correct and if the Bot is published/added to the user/channel.");
            }
            if (code === 403) {
                console.error("Hint: Permission denied. Check Bot scope or domain settings.");
            }
        }

    } catch (e) {
        console.error("❌ Exeption during send: " + e.message);
    }

    console.log("=== Test Complete ===");
}
