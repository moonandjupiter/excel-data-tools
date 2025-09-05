(function () {
    "use strict";

    // Wait for the DOM to be loaded before initializing anything.
    document.addEventListener('DOMContentLoaded', initialize);

    function initialize() {
        // --- UI Initialization ---
        document.getElementById("tab-inserter").onclick = () => switchTab('inserter');
        document.getElementById("tab-cleanser").onclick = () => switchTab('cleanser');
        document.getElementById("tab-gemini").onclick = () => switchTab('gemini');
        
        document.getElementById("generate-rows").onclick = generateAndInsertRows;
        document.getElementById("cleanse-data").onclick = cleanseAndPasteData;
        document.getElementById("process-gemini").onclick = processWithGemini;
        document.getElementById("copy-gemini-result").onclick = copyGeminiResult;

        document.getElementById('rowCount').onfocus = function() { this.select(); };
        document.getElementById('rawData').onfocus = function() { this.select(); };
        document.getElementById('gemini-input').onfocus = function() { this.select(); };

        switchTab('inserter');

        // --- Office-Specific Initialization ---
        Office.onReady(function (info) {
            // Office host is ready.
        });
    }

    function switchTab(tabName) {
        // Hide all content panels
        document.getElementById('inserter-content').style.display = 'none';
        document.getElementById('cleanser-content').style.display = 'none';
        document.getElementById('gemini-content').style.display = 'none';
        
        // Deactivate all tab buttons
        document.getElementById('tab-inserter').classList.remove('active-tab');
        document.getElementById('tab-cleanser').classList.remove('active-tab');
        document.getElementById('tab-gemini').classList.remove('active-tab');
        
        // Show the selected content panel and activate the corresponding tab
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
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                const range = context.workbook.getSelectedRange();
                
                const lastRowOfSelection = range.getEntireRow().getLastRow();
                lastRowOfSelection.load("rowIndex");
                await context.sync();

                const insertStartRow = lastRowOfSelection.rowIndex + 2;
                const insertEndRow = insertStartRow + count - 1;
                const rangeAddress = `${insertStartRow}:${insertEndRow}`;
                
                sheet.getRange(rangeAddress).insert(Excel.InsertShiftDirection.down);
                
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
                selection.load("rowIndex", "columnIndex");
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
    
    function copyGeminiResult() {
        const output = document.getElementById('gemini-output');
        const status = document.getElementById('gemini-status');
        if (output.value) {
            output.select();
            document.execCommand('copy');
            status.textContent = 'Result copied to clipboard!';
            setTimeout(() => { status.textContent = ''; }, 2000);
        }
    }

    async function processWithGemini() {
        const input = document.getElementById('gemini-input').value.trim();
        const status = document.getElementById('gemini-status');
        const button = document.getElementById('process-gemini');
        const outputContainer = document.getElementById('gemini-output-container');
        const outputTextarea = document.getElementById('gemini-output');

        outputContainer.style.display = 'none';
        outputTextarea.value = '';

        if (!input) {
            status.textContent = 'Please enter some data to process.';
            setTimeout(() => { status.textContent = ''; }, 3000);
            return;
        }

        if (typeof Excel === 'undefined') {
            status.textContent = 'This feature only works inside Excel.';
            return;
        }

        button.disabled = true;
        status.textContent = 'Processing with Gemini...';

        try {
            const prompt = `
                You are a highly efficient data cleansing and expansion tool. Your task is to take a compressed or ranged string of data and expand it into a list where each line represents a single, complete item. The output should be plain text, with one item per line, and nothing else.
                Follow these rules strictly:
                1. Maintain the full prefix (the part of the string before the serial number, model, or item identifier).
                2. Expand ranges. For example, "10702P to 10704P" becomes three separate lines.
                3. Expand abbreviated serial numbers. For example, if the input is "HP DL380p Gen8 S#SGH438WACA,...WAC8", the second entry should be completed to "HP DL380p Gen8 S#SGH438WAC8".
                4. Handle complex formats where the description comes after the serial numbers, like "SN: 200108F & 200107F, Brand: Arrow...". Each output line must contain the full description.
                Input Data to process:
                ${input}
                Final Output:
            `;

            const payload = {
                contents: [{ parts: [{ text: prompt }] }],
                generationConfig: { temperature: 0 },
            };
            const apiKey = ""; // Canvas will provide this.
            const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash-preview-05-20:generateContent?key=${apiKey}`;

            const response = await fetch(apiUrl, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            });

            if (!response.ok) throw new Error(`API call failed: ${response.status}`);

            const result = await response.json();
            const generatedText = result?.candidates?.[0]?.content?.parts?.[0]?.text;

            if (!generatedText) throw new Error("No text generated by the model.");
            
            outputTextarea.value = generatedText.trim();
            outputContainer.style.display = 'block';

            const cleansedData = generatedText.trim().split('\n');

            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                const selection = context.workbook.getSelectedRange();
                selection.load("rowIndex", "columnIndex");
                await context.sync();

                const dataToInsert = cleansedData.map(item => [item]);
                const targetRange = sheet.getRangeByIndexes(selection.rowIndex, selection.columnIndex, dataToInsert.length, 1);
                targetRange.values = dataToInsert;
                await context.sync();
                status.textContent = `Successfully pasted ${cleansedData.length} items.`;
            });

        } catch (error) {
            console.error(error);
            status.textContent = 'Error: Could not process data.';
            outputContainer.style.display = 'none';
        } finally {
            button.disabled = false;
            if (status.textContent.startsWith('Error')) {
                 setTimeout(() => { status.textContent = ''; }, 4000);
            }
        }
    }

    function parseData(rawData) {
        const allCleansedLines = [];
        const lines = rawData.split('\n').filter(line => line.trim() !== '');

        lines.forEach(line => {
            let results = [];
            let processed = false;

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

            if (!processed) {
                const snKeywordMatch = line.match(/(S#s|S#|SN:|Ser\. No\.|SN)/i);
                if (snKeywordMatch) {
                    const keyword = snKeywordMatch[0];
                    const lineParts = line.split(keyword);
                    const descBefore = lineParts[0];
                    const everythingAfter = lineParts.slice(1).join(keyword).trim();

                    let serialsString = everythingAfter;
                    let descAfter = '';

                    const splitMatch = everythingAfter.match(/,(?=\s*(?:Brand|Model|Type|w\/))/i);
                    if (splitMatch) {
                        serialsString = everythingAfter.substring(0, splitMatch.index);
                        descAfter = everythingAfter.substring(splitMatch.index);
                    }

                    const parts = serialsString.split(/[,/&]|\s+/).filter(p => p && p.trim() !== '');
                    let lastPrefix = '';

                    parts.forEach(part => {
                        let currentSerial = part.trim();
                        if (!currentSerial) return;

                        if (/[a-zA-Z-]/.test(currentSerial)) {
                            results.push(`${descBefore}${keyword} ${currentSerial}${descAfter}`.trim());
                            const match = currentSerial.match(/^(.*[a-zA-Z-])(\d+)$/);
                            if (match) {
                                lastPrefix = match[1];
                            }
                        } else if (lastPrefix) {
                            results.push(`${descBefore}${keyword} ${lastPrefix}${currentSerial}${descAfter}`.trim());
                        } else {
                             results.push(`${descBefore}${keyword} ${currentSerial}${descAfter}`.trim());
                        }
                    });
                    
                    if (results.length > 0) processed = true;
                }
            }
            
            if (processed) {
                allCleansedLines.push(...results);
            } else {
                if (line.includes(',')) {
                    const parts = line.split(/[,&]/).map(p => p.trim()).filter(Boolean);
                    let lastPrefix = '';
                    parts.forEach(part => {
                        if (/[a-zA-Z-]/.test(part)) {
                            allCleansedLines.push(part);
                            const match = part.match(/^(.*[a-zA-Z-])(\d+)$/);
                            if (match) lastPrefix = match[1];
                        } else if (lastPrefix) {
                            allCleansedLines.push(lastPrefix + part);
                        } else {
                            allCleansedLines.push(part);
                        }
                    });
                } else {
                    allCleansedLines.push(line.trim());
                }
            }
        });

        return allCleansedLines;
    }

})();

