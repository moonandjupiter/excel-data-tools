(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded.
    Office.onReady(function (info) {
        if (info.host === Office.HostType.Excel) {
            // Assign event handlers and other initialization logic.
            document.getElementById("tab-inserter").onclick = () => switchTab('inserter');
            document.getElementById("tab-cleanser").onclick = () => switchTab('cleanser');
            document.getElementById("generate-rows").onclick = generateAndInsertRows;
            document.getElementById("cleanse-data").onclick = cleanseAndPasteData;

            // Set the initial tab
            switchTab('inserter');
        }
    });

    function switchTab(tabName) {
        document.getElementById('inserter-content').style.display = 'none';
        document.getElementById('cleanser-content').style.display = 'none';
        document.getElementById('tab-inserter').classList.remove('active');
        document.getElementById('tab-cleanser').classList.remove('active');

        document.getElementById(tabName + '-content').style.display = 'block';
        document.getElementById('tab-' + tabName).classList.add('active');
    }

    async function generateAndInsertRows() {
        const rowCountInput = document.getElementById('rowCount');
        const count = parseInt(rowCountInput.value, 10);
        const status = document.getElementById('rowStatus');

        if (isNaN(count) || count < 1) {
            status.textContent = 'Please enter a valid number.';
            setTimeout(() => { status.textContent = ''; }, 2000);
            return;
        }

        try {
            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                const selection = context.workbook.getSelectedRange();
                selection.load("rowIndex, rowCount");
                await context.sync();

                const insertAtRow = selection.rowIndex + selection.rowCount;
                for (let i = 0; i < count; i++) {
                    sheet.getRangeByIndexes(insertAtRow, 0, 1, 1).insert(Excel.InsertShiftDirection.down);
                }
                await context.sync();
                status.textContent = `Successfully inserted ${count} row(s).`;
            });
        } catch (error) {
            console.error(error);
            status.textContent = 'Error inserting rows.';
        }
        setTimeout(() => { status.textContent = ''; }, 3000);
    }

    async function cleanseAndPasteData() {
        const rawData = document.getElementById('rawData').value;
        const cleansedData = parseData(rawData);
        const status = document.getElementById('cleanseStatus');

        if (cleansedData.length === 0) {
            status.textContent = 'No data to paste.';
            setTimeout(() => { status.textContent = ''; }, 3000);
            return;
        }

        try {
            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                const selection = context.workbook.getSelectedRange();
                selection.load("rowIndex, columnIndex");
                await context.sync();

                const dataToInsert = cleansedData.map(item => [item]);
                const targetRange = sheet.getRangeByIndexes(selection.rowIndex, selection.columnIndex, dataToInsert.length, 1);
                targetRange.values = dataToInsert;

                await context.sync();
                status.textContent = `Successfully pasted ${cleansedData.length} items.`;
            });
        } catch (error) {
            console.error(error);
            status.textContent = 'Error pasting data.';
        }
        setTimeout(() => { status.textContent = ''; }, 3000);
    }

    function parseData(rawData) {
        const lines = rawData.split('\n').filter(line => line.trim() !== '');
        let cleansedLines = [];

        for (const line of lines) {
            // Handle ranges like "17P-07 to 09" or "16P-2301-5 to 8"
            const rangeMatch = line.match(/(.*?)(\d+)\s*(?:to|-)\s*(\d+)$/i);
            if (rangeMatch) {
                const prefix = rangeMatch[1].trim();
                const startStr = rangeMatch[2];
                const endStr = rangeMatch[3];

                const startNum = parseInt(startStr, 10);
                const endNum = parseInt(endStr, 10);
                
                // The padding length is determined by the starting number's format.
                const padLength = startStr.length;

                if (!isNaN(startNum) && !isNaN(endNum) && endNum >= startNum) {
                    for (let i = startNum; i <= endNum; i++) {
                        const paddedNum = String(i).padStart(padLength, '0');
                        cleansedLines.push(`${prefix}${paddedNum}`);
                    }
                    continue; // Done with this line, move to the next.
                }
            }

            // Handle complex formats with delimiters (commas, slashes, etc.)
            const snKeywordMatch = line.match(/(S#s|S#|SN:|Ser\. No\.|SN)/i);
            if (snKeywordMatch) {
                const snIndex = snKeywordMatch.index;
                const description = line.substring(0, snIndex).trim();
                const keyword = snKeywordMatch[0];
                const serialsString = line.substring(snIndex + keyword.length).trim();
                const serialParts = serialsString.split(/[,/&;]|\s+/).filter(p => p.trim() !== '');
                
                let lastFullSerial = "";
                let lastPrefix = "";

                for (const part of serialParts) {
                    let currentSerial = part.replace(/\.+/g, '').trim(); // Remove '..' or '...'
                    if (!currentSerial) continue;
                    
                    // A part is considered "full" if it contains non-numeric characters or is long.
                    // This handles switching prefixes like "0141HM2122/.../0142HM2145"
                    if (isNaN(currentSerial) || currentSerial.length > 6 || !lastPrefix) {
                        lastFullSerial = currentSerial;
                        // Find the last number part to determine the prefix
                        const match = lastFullSerial.match(/^(.*?)(\d+)$/);
                        if (match) {
                           lastPrefix = match[1];
                        } else {
                           lastPrefix = lastFullSerial; // It might be all non-numeric
                        }
                    } else {
                        // This is a partial number, append it to the last prefix.
                        currentSerial = lastPrefix + currentSerial;
                    }
                    cleansedLines.push(`${description} ${keyword} ${currentSerial}`);
                }
                continue; // Done with this line
            }
            
            // If no specific format is matched, add the line as is.
            if(line.trim()){
                 cleansedLines.push(line.trim());
            }
        }

        return cleansedLines;
    }
})();

