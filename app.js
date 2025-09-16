(function () {
    "use strict";

    // This listener ensures all UI elements are ready before we try to attach events to them.
    document.addEventListener('DOMContentLoaded', initialize);

    function initialize() {
        // --- UI Initialization ---
        document.getElementById("generate-rows").onclick = generateAndInsertRows;
        document.getElementById("get-cleanse-paste").onclick = getCleanseAndPaste;
        document.getElementById("get-from-selection").onclick = getFromSelection;
        document.getElementById("cleanse-data").onclick = cleanseAndPasteData;

        document.getElementById('rowCount').onfocus = function() { this.select(); };
        document.getElementById('rawData').onfocus = function() { this.select(); };

        // The Office.onReady call is only to confirm the host, not for UI setup.
        Office.onReady((info) => {
            if (info.host === Office.HostType.Excel) {
                console.log("Add-in is ready and running in Excel.");
            }
        });
    }

    async function generateAndInsertRows() {
        const rowCountInput = document.getElementById('rowCount');
        const countToInsert = parseInt(rowCountInput.value, 10);
        const status = document.getElementById('rowStatus');

        if (isNaN(countToInsert) || countToInsert < 1) {
            status.textContent = 'Please enter a valid number.';
            setTimeout(() => { status.textContent = ''; }, 2000);
            return;
        }

        try {
            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                const range = context.workbook.getSelectedRange();
                
                range.load(["rowIndex", "rowCount"]);
                await context.sync();

                const startRow = range.rowIndex;
                const numSelectedRows = range.rowCount;

                for (let i = numSelectedRows - 1; i >= 0; i--) {
                    const currentRowIndex = startRow + i;
                    const insertAtRowNumber = currentRowIndex + 2;
                    const insertRangeAddress = `${insertAtRowNumber}:${insertAtRowNumber + countToInsert - 1}`;
                    sheet.getRange(insertRangeAddress).insert(Excel.InsertShiftDirection.down);
                }
                
                await context.sync();
                const totalInserted = numSelectedRows * countToInsert;
                status.textContent = `Successfully inserted ${totalInserted} row(s).`;
            });
        } catch (error) {
            console.error(error);
            status.textContent = 'Error: Could not insert rows.';
        }
        setTimeout(() => { status.textContent = ''; }, 3000);
    }

    async function getCleanseAndPaste() {
        const rawDataTextarea = document.getElementById('rawData');
        const status = document.getElementById('cleanseStatus');

        try {
            await Excel.run(async (context) => {
                const selection = context.workbook.getSelectedRange();
                selection.load("values");
                await context.sync();
                const selectionText = selection.values.map(row => row.join("\t")).join("\n");
                
                let cleansedData = parseData(selectionText);
                
                // Update the textarea with the formatted, numbered list for display
                const displayData = formatForDisplay(cleansedData);
                rawDataTextarea.value = displayData;

                // The data to be pasted remains the original un-numbered version
                cleansedData = cleansedData.filter(line => (line && !/^\s*$/.test(line)) || line === '');

                if (cleansedData.length === 0) {
                    status.textContent = "No data to paste after cleansing.";
                    return; 
                }

                const dataToInsert = cleansedData.map(item => [item]);
                const targetRange = selection.getCell(0, 0).getResizedRange(dataToInsert.length - 1, 0);
                targetRange.values = dataToInsert;

                selection.getCell(0, 0).select();
                await context.sync();
                status.textContent = `Pasted ${cleansedData.length} cleansed items.`;
            });
        } catch (error) {
            console.error(error);
            status.textContent = "Error: Could not process selection.";
        }
        setTimeout(() => { status.textContent = ''; }, 3000);
    }

    async function getFromSelection() {
        const rawDataTextarea = document.getElementById('rawData');
        const status = document.getElementById('cleanseStatus');

        try {
            await Excel.run(async (context) => {
                const range = context.workbook.getSelectedRange();
                range.load("values");
                await context.sync();

                const selectionText = range.values.map(row => row.join("\t")).join("\n");
                rawDataTextarea.value = selectionText;
                status.textContent = "Data loaded from selection.";
            });
        } catch (error) {
            console.error(error);
            status.textContent = "Error: Could not get data.";
        }
        setTimeout(() => { status.textContent = ''; }, 3000);
    }

    async function cleanseAndPasteData() {
        const rawDataTextarea = document.getElementById('rawData');
        const rawData = rawDataTextarea.value;
        let cleansedData = parseData(rawData);
        const status = document.getElementById('cleanseStatus');
        
        // Update textarea with the formatted version for visual feedback
        const displayData = formatForDisplay(cleansedData);
        rawDataTextarea.value = displayData;

        // Filter the actual data to be pasted
        cleansedData = cleansedData.filter(line => (line && !/^\s*$/.test(line)) || line === '');

        if (cleansedData.length === 0) {
            status.textContent = rawData.trim() ? 'Could not parse data.' : 'No data to paste.';
            setTimeout(() => { status.textContent = ''; }, 3000);
            return;
        }

        try {
            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                const selection = context.workbook.getSelectedRange();
                
                const dataToInsert = cleansedData.map(item => [item]);
                
                const targetRange = selection.getCell(0, 0).getResizedRange(dataToInsert.length - 1, 0);
                targetRange.values = dataToInsert;

                selection.getCell(0, 0).select();
                await context.sync();
                status.textContent = `Successfully pasted ${cleansedData.length} items.`;
            });
        } catch (error) {
            console.error(error);
            status.textContent = 'Error: Could not paste data.';
        }
        setTimeout(() => { status.textContent = ''; }, 3000);
    }
    
    function formatForDisplay(cleansedData) {
        let displayLines = [];
        let counter = 1;
        cleansedData.forEach(line => {
            if (line === '') {
                displayLines.push('');
                counter = 1; // Reset counter for the next block
            } else {
                displayLines.push(`${counter}. ${line}`);
                counter++;
            }
        });
        return displayLines.join('\n');
    }

    function parseData(rawData) {
        const allCleansedLines = [];
        const blocks = rawData.split(/\n\s*\n/);

        blocks.forEach((block, blockIndex) => {
            const linesInBlock = block.split('\n').filter(line => !/^\s*$/.test(line));
            const uniqueLines = [...new Set(linesInBlock)];

            uniqueLines.forEach(line => {
                let results = [];
                let processed = false;

                if (line.includes(';')) {
                    const segments = line.split(';').map(s => s.trim());
                    let prefix = '';
                    let processedInThisBlock = false;
                    
                    const prefixMatch = segments[0].match(/^(.*?)\d/);
                    if (prefixMatch) {
                        prefix = prefixMatch[1].trim();
                    }

                    segments.forEach(segment => {
                        const normalizedSegment = segment.replace(/\s+to\s+/i, '-').replace(/\s/g, '');
                        const rangeMatch = normalizedSegment.match(/(\d+)-(\d+)/);
                        
                        if (rangeMatch) {
                            const startStr = rangeMatch[1];
                            const endStr = rangeMatch[2];
                            const startNum = parseInt(startStr, 10);
                            const endNum = parseInt(endStr, 10);
                            const padLength = startStr.length;
                            
                            if (!isNaN(startNum) && !isNaN(endNum) && endNum >= startNum) {
                                for (let i = startNum; i <= endNum; i++) {
                                    results.push(`${prefix}${String(i).padStart(padLength, '0')}`);
                                }
                                processedInThisBlock = true;
                            }
                        } else if (/^\d+$/.test(normalizedSegment)) {
                            const numStr = normalizedSegment;
                            const firstNumInLine = line.match(/\d+/);
                            const padLength = firstNumInLine ? firstNumInLine[0].length : numStr.length;
                            results.push(`${prefix}${String(numStr).padStart(padLength, '0')}`);
                            processedInThisBlock = true;
                        }
                    });

                    if (processedInThisBlock) {
                        processed = true;
                    }
                }

                if (!processed) {
                    const rangeMatch = line.match(/^(.*?)(\d+)\s*(?:to|-)\s*(\d+)(\D*)$/i);
                    if (rangeMatch) {
                        const prefix = rangeMatch[1].trim();
                        const startStr = rangeMatch[2];
                        const endStr = rangeMatch[3];
                        const suffix = rangeMatch[4].trim();
                        const startNum = parseInt(startStr, 10);
                        const endNum = parseInt(endStr, 10);
                        const padLength = startStr.length;

                        if (!isNaN(startNum) && !isNaN(endNum) && endNum >= startNum) {
                            for (let i = startNum; i <= endNum; i++) {
                                results.push(`${prefix}${String(i).padStart(padLength, '0')}${suffix}`);
                            }
                            if (results.length > 0) processed = true;
                        }
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
                                if (match) lastPrefix = match[1];
                            } else if (lastPrefix) {
                                results.push(`${descBefore}${keyword} ${lastPrefix}${currentSerial}${descAfter}`.trim());
                            } else {
                                results.push(`${descBefore}${keyword} ${currentSerial}${descAfter}`.trim());
                            }
                        });
                        if (results.length > 0) processed = true;
                    }
                }
                
                if (!processed && (line.includes(',') || line.includes('&'))) {
                    const parts = line.split(/[,&]/).map(p => p.trim()).filter(Boolean);
                    if (parts.length > 1) {
                        let firstPart = parts[0];
                        let lastPrefix = '';
                        
                        const match = firstPart.match(/^(.*\D)(\d+)$/);
                        if (match) {
                            lastPrefix = match[1];
                            results.push(firstPart);
                            for (let i = 1; i < parts.length; i++) {
                                results.push(lastPrefix + parts[i]);
                            }
                        } else {
                            results.push(...parts);
                        }
                        if(results.length > 0) processed = true;
                    }
                }

                if (processed) {
                    allCleansedLines.push(...results);
                } else {
                    allCleansedLines.push(line.trim());
                }
            });

            if (blockIndex < blocks.length - 1) {
                allCleansedLines.push('');
            }
        });

        return allCleansedLines;
    }

})();

