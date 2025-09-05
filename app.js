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

    function tryParseRange(line) {
        const results = [];
        const rangeMatch = line.match(/(.*?)(\d+)\s*(?:to|-)\s*(\d+)$/i);
        if (rangeMatch) {
            const prefix = rangeMatch[1].trim();
            const startStr = rangeMatch[2];
            const endStr = rangeMatch[3];
            const startNum = parseInt(startStr, 10);
            const endNum = parseInt(endStr, 10);
            const padLength = startStr.length;

            if (!isNaN(startNum) && !isNaN(endNum) && endNum >= startNum) {
                for (let i = startNum; i <= endNum; i++) {
                    const paddedNum = String(i).padStart(padLength, '0');
                    results.push(`${prefix}${paddedNum}`);
                }
            }
        }
        return results;
    }

    function tryParseSerialNumbers(line) {
        const results = [];
        const snKeywordMatch = line.match(/(S#s|S#|SN:|Ser\. No\.|SN)/i);
        if (snKeywordMatch) {
            const snIndex = snKeywordMatch.index;
            const description = line.substring(0, snIndex).trim();
            const keyword = snKeywordMatch[0];
            const serialsString = line.substring(snIndex + keyword.length).trim();
            const serialParts = serialsString.split(/[,/&;]|\s+/).filter(p => p.trim() !== '');
            
            let lastPrefix = "";

            for (const part of serialParts) {
                let currentSerial = part.replace(/\.+/g, '').trim();
                if (!currentSerial) continue;

                if (isNaN(currentSerial) || currentSerial.length > 6 || !lastPrefix) {
                    const match = currentSerial.match(/^(.*?)(\d+)$/);
                    if (match) {
                        lastPrefix = match[1];
                    } else {
                        lastPrefix = currentSerial;
                    }
                } else {
                    currentSerial = lastPrefix + currentSerial;
                }
                results.push(`${description} ${keyword} ${currentSerial}`);
            }
        }
        return results;
    }

    function parseData(rawData) {
        const lines = rawData.split('\n').filter(line => line.trim() !== '');
        let allCleansedLines = [];

        for (const line of lines) {
            let processedLines = tryParseRange(line);

            if (processedLines.length === 0) {
                processedLines = tryParseSerialNumbers(line);
            }

            if (processedLines.length === 0 && line.trim()) {
                allCleansedLines.push(line.trim());
            } else {
                allCleansedLines.push(...processedLines);
            }
        }

        return allCleansedLines;
    }
})();

