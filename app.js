(function () {
    "use strict";

    // Wait for the DOM to be loaded before initializing anything.
    document.addEventListener('DOMContentLoaded', initialize);

    function initialize() {
        // --- UI Initialization ---
        // Assign event handlers for the UI elements that should work everywhere, including a browser.
        document.getElementById("tab-inserter").onclick = () => switchTab('inserter');
        document.getElementById("tab-cleanser").onclick = () => switchTab('cleanser');
        document.getElementById("generate-rows").onclick = generateAndInsertRows;
        document.getElementById("cleanse-data").onclick = cleanseAndPasteData;

        // Set the initial tab view.
        switchTab('inserter');

        // --- Office-Specific Initialization ---
        // Office.onReady is used to ensure the Office host is ready before we try to use its APIs.
        Office.onReady(function (info) {
            // You can add any host-specific logic here. For now, it's just confirming readiness.
        });
    }

    function switchTab(tabName) {
        // Hide both content panes
        document.getElementById('inserter-content').style.display = 'none';
        document.getElementById('cleanser-content').style.display = 'none';
        
        // Remove the active style from both tab buttons
        document.getElementById('tab-inserter').classList.remove('active-tab');
        document.getElementById('tab-cleanser').classList.remove('active-tab');

        // Show the selected content pane
        document.getElementById(tabName + '-content').style.display = 'block';
        
        // Add the active style to the selected tab button
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

        // Check if we are running inside Excel before using the Excel API.
        if (!window.Excel) {
            status.textContent = 'This feature only works inside Excel.';
            console.error("Excel API is not available in this context.");
            setTimeout(() => { status.textContent = ''; }, 3000);
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
            status.textContent = rawData.trim() ? 'Could not parse data.' : 'No data to paste.';
            setTimeout(() => { status.textContent = ''; }, 3000);
            return;
        }

        // Check if we are running inside Excel before using the Excel API.
        if (!window.Excel) {
            status.textContent = 'This feature only works inside Excel.';
            console.error("Excel API is not available in this context.");
            // For browser testing, you can log the output to see if parsing works.
            console.log("Cleansed Data (Browser Mode):", cleansedData);
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

            // Pattern 2: Serial Number Lists
            if (!processed) {
                const snKeywordMatch = line.match(/(S#s|S#|SN:|Ser\. No\.|SN)/i);
                if (snKeywordMatch) {
                    const snIndex = snKeywordMatch.index;
                    const description = line.substring(0, snIndex).trim();
                    const keyword = snKeywordMatch[0];
                    const serialsString = line.substring(snIndex + keyword.length).trim();
                    const serialParts = serialsString.split(/[,/&;]|\s+/).filter(p => p && p.trim() !== '' && !p.startsWith('.'));
                    
                    if (serialParts.length > 0) {
                        let lastPrefix = "";
                        let fullDescription = `${description} ${keyword}`;

                        serialParts.forEach(part => {
                            let currentSerial = part.trim();
                            
                            if (/[a-zA-Z-]/.test(currentSerial) || currentSerial.length > 6 || !lastPrefix) {
                                 const match = currentSerial.match(/^(.*?)(\d+)$/);
                                 if (match) {
                                     lastPrefix = match[1];
                                 } else {
                                     lastPrefix = ''; 
                                 }
                                 results.push(`${fullDescription} ${currentSerial}`);
                            } else {
                                results.push(`${fullDescription} ${lastPrefix}${currentSerial}`);
                            }
                        });
                        if (results.length > 0) processed = true;
                    }
                }
            }
            
            // Final decision: push the results or fall back to the original line
            if (processed) {
                allCleansedLines.push(...results);
            } else {
                allCleansedLines.push(line.trim());
            }
        });

        return allCleansedLines;
    }
})();

