(function () {
    "use strict";

    // Wait for the DOM to be loaded before initializing anything.
    document.addEventListener('DOMContentLoaded', initialize);

    function initialize() {
        // --- UI Initialization ---
        // Assign event handlers for the UI elements.
        document.getElementById("tab-inserter").onclick = () => switchTab('inserter');
        document.getElementById("tab-cleanser").onclick = () => switchTab('cleanser');
        document.getElementById("generate-rows").onclick = generateAndInsertRows;
        document.getElementById("cleanse-data").onclick = cleanseAndPasteData;

        // Restore the "focus to select all" functionality.
        document.getElementById('rowCount').onfocus = function() { this.select(); };
        document.getElementById('rawData').onfocus = function() { this.select(); };

        // Set the initial tab view.
        switchTab('inserter');

        // --- Office-Specific Initialization ---
        Office.onReady(function (info) {
            // Office host is ready.
        });
    }

    function switchTab(tabName) {
        document.getElementById('inserter-content').style.display = 'none';
        document.getElementById('cleanser-content').style.display = 'none';
        document.getElementById('tab-inserter').classList.remove('active-tab');
        document.getElementById('tab-cleanser').classList.remove('active-tab');
        document.getElementById(tabName + '-content').style.display = 'block';
        document.getElementById('tab-' + tabName).classList.add('active-tab');
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

        if (typeof Excel === 'undefined') {
            status.textContent = 'This feature only works inside Excel.';
            return;
        }

        try {
            await Excel.run(async (context) => {
                const range = context.workbook.getSelectedRange();
                // Get the last row of the current selection.
                const lastRow = range.getLastRow();
                lastRow.load("rowIndex");
                await context.sync();

                // Define a range representing the 'count' of entire rows below the selection.
                const rowsToInsert = lastRow.worksheet.getRangeByIndexes(lastRow.rowIndex + 1, 0, count, 0);
                // Insert the new rows.
                rowsToInsert.insert(Excel.InsertShiftDirection.down);
                await context.sync();
                status.textContent = `Successfully inserted ${count} row(s).`;
            });
        } catch (error) {
            console.error(error);
            status.textContent = 'Error: Could not insert rows.';
        }
        setTimeout(() => { status.textContent = ''; }, 3000);
    }

    async function cleanseAndPasteData() {
        const rawData = document.getElementById('rawData').value;
        const cleansedData = parseData(rawData);
        const status = document.getElementById('cleanseStatus');

        if (cleansedData.length === 0) {
            status.textContent = rawData.trim() ? 'Could not parse data.' : 'No data to paste.';
            setTimeout(() => { status.textContent = ''; }, 3000);
            return;
        }

        if (typeof Excel === 'undefined') {
            status.textContent = 'This feature only works inside Excel.';
            console.log("Cleansed Data (Browser Mode):", cleansedData);
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
            status.textContent = 'Error: Could not paste data.';
        }
        setTimeout(() => { status.textContent = ''; }, 3000);
    }
    
    function parseData(rawData) {
        const allCleansedLines = [];
        const lines = rawData.split('\n').filter(line => line.trim() !== '');

        lines.forEach(line => {
            let results = [];
            let processed = false;

            // Pattern 1: Simple Ranges (e.g., "17P-07 to 09")
            const rangeMatch = line.match(/^(.*?)(\d+)\s*(?:to|-)\s*(\d+)$/i);
            if (rangeMatch) {
                const prefix = rangeMatch[1].trim();
                const startStr = rangeMatch[2];
                const endStr = rangeMatch[3];
                const startNum = parseInt(startStr, 10);
                const endNum = parseInt(endStr, 10);
                const padLength = startStr.length;

                if (!isNaN(startNum) && !isNaN(endNum) && endNum >= startNum) {
                    for (let i = startNum; i <= endNum; i++) {
                        results.push(`${prefix}${String(i).padStart(padLength, '0')}`);
                    }
                    if (results.length > 0) processed = true;
                }
            }

            // Pattern 2: Serial Number Lists with implied prefixes
            if (!processed) {
                 if (line.includes(',') || line.includes('&') || line.match(/(S#s|S#|SN:|Ser\. No\.|SN)/i)) {
                    let description = '';
                    let serialsString = line;
                    const snKeywordMatch = line.match(/(.*?)(S#s|S#|SN:|Ser\. No\.|SN)\s*(.*)/i);

                    if (snKeywordMatch) {
                        description = `${snKeywordMatch[1]}${snKeywordMatch[2]} `.trim();
                        serialsString = snKeywordMatch[3];
                    }

                    const parts = serialsString.split(/[,/&]|\s+/).filter(p => p && p.trim() !== '');
                    let lastPrefix = '';

                    parts.forEach(part => {
                        let currentSerial = part.trim();
                        if (!currentSerial) return;

                        if (/[a-zA-Z-]/.test(currentSerial)) {
                            results.push(`${description}${currentSerial}`);
                            const match = currentSerial.match(/^(.*?)(\d+)$/);
                            if (match) {
                                lastPrefix = match[1];
                            }
                        } else if (lastPrefix) {
                            results.push(`${description}${lastPrefix}${currentSerial}`);
                        } else {
                            results.push(`${description}${currentSerial}`);
                        }
                    });
                    
                    if (results.length > 0) processed = true;
                }
            }
            
            if (processed) {
                allCleansedLines.push(...results);
            } else {
                allCleansedLines.push(line.trim());
            }
        });

        return allCleansedLines;
    }

})();

