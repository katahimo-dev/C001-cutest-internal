// Test function to debug address parsing
function testAddressExtraction() {
    const testAddresses = [
        "仙台市泉区朝日",
        "仙台市若林区五橋",
        "宮城県仙台市青葉区中央",
        "仙台市泉区あすと長町",
        "仙台市若林区相川知美"
    ];

    testAddresses.forEach(addr => {
        Logger.log("Original: " + addr);

        // Test the regex
        let pref = "";
        const prefRegex = /^(東京都|北海道|(?:京都|大阪)府|.{2,3}県)\s*/;
        const prefMatch = addr.match(prefRegex);
        let remainder = addr;

        if (prefMatch) {
            pref = prefMatch[0];
            remainder = addr.substring(pref.length);
            Logger.log("  Prefecture: " + pref);
            Logger.log("  Remainder: " + remainder);
        }

        // Try to match City + Ward
        const cityWardMatch = remainder.match(/^([^0-9\s]+市[^0-9\s]+区)/);

        if (cityWardMatch) {
            Logger.log("  City+Ward: " + cityWardMatch[1]);
            Logger.log("  Full: " + pref + cityWardMatch[1]);
        } else {
            Logger.log("  NO MATCH!");
        }

        Logger.log("---");
    });
}
