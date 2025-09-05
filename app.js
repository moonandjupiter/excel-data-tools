/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// Initialize the Office Add-in
Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.addEventListener("DOMContentLoaded", function() {
            // Set the default tab on load
            switchTab('rowInserter');
        });
    }
});

/**
 * Switches the visible tab in the add-in.
 * @param {string} tabName The ID of the tab to make active.
 */
function switchTab(tabName) {
    const tabs = ['rowInserter', 'dataCleanser'];
    tabs.forEach(tab => {
        const content = document.getElementById(`content-${tab}`);
        const tabButton = document.getElementById(`tab-${tab}`);
        if (tab === tabName) {
            content.classList.remove('hidden');
            tabButton.classList.add('tab-active');
            tabButton.classList.remove('border-transparent', 'text-gray-600', 'hover:text-gray-800', 'hover:border-gray-400');
        } else {
            content.classList.add('hidden');
            tabButton.classList.remove('tab-active');
            tabButton.classList.add('border-transparent', 'text-gray-600', 'hover:text-gray-800', 'hover:border-gray-400');
        }
    });
}

/**
 * Inserts a specified number of empty rows into the current Excel worksheet.
 */
async function insertRows() {
    const rowCountInput = document.getElementById('rowCount');
    const status = document.getElementById('rowStatus');
    const count = parseInt(rowCountInput.value, 10);

    if (isNaN(count) || count < 1) {
        status.textContent = 'Please enter a valid number.';
        setTimeout(() => { status.textContent = ''; }, 3000);
        return;
    }

    try {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.load("address, rowCount");
            await context.sync();

            // To insert rows *below* the selection, we start from the row after the selection ends.
            const insertRange = range.worksheet.getRangeByIndexes(range.rowIndex + range.rowCount, 0, count, 1);
            insertRange.insert(Excel.InsertShiftDirection.down);
            await context.sync();

            status.textContent = `${count} row${count > 1 ? 's' : ''} inserted.`;
            setTimeout(() => { status.textContent = ''; }, 3000);
        });
    } catch (error) {
        console.error(error);
        status.textContent = 'Error inserting rows.';
        setTimeout(() => { status.textContent = ''; }, 3000);
    }
}

/**
 * Takes raw text, cleanses it into a list of serial numbers, and pastes the result into Excel.
 */
async function cleanseAndPasteData() {
    const rawData = document.getElementById('rawData').value;
    const status = document.getElementById('status');
    const cleansedData = cleanseData(rawData);

    if (cleansedData.length === 0) {
        status.textContent = 'No data to paste.';
        setTimeout(() => { status.textContent = ''; }, 3000);
        return;
    }

    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const range = sheet.getUsedRange();
            const lastCell = range.getLastCell();
            lastCell.load('rowIndex');
            await context.sync();
            
            // Determine the starting cell for pasting. If the sheet is empty, start at A1. Otherwise, start on the next row.
            const startRow = range.address ? lastCell.rowIndex + 1 : 0;
            const targetRange = sheet.getRangeByIndexes(startRow, 0, cleansedData.length, 1);
            
            // Format data for Excel's 2D array requirement
            const formattedData = cleansedData.map(item => [item]);

            targetRange.values = formattedData;
            targetRange.select();
            await context.sync();

            status.textContent = `${cleansedData.length} item(s) pasted.`;
            setTimeout(() => { status.textContent = ''; }, 3000);
        });
    } catch (error) {
        console.error(error);
        status.textContent = 'Error pasting data.';
        setTimeout(() => { status.textContent = ''; }, 3000);
    }
}

/**
 * The core data parsing and cleansing logic.
 * Handles ranges, various delimiters, and changing serial number prefixes.
 * @param {string} rawText The raw input string from the textarea.
 * @returns {string[]} An array of cleansed, individual serial number lines.
 */
function cleanseData(rawText) {
    const finalOutput = [];
    const lines = rawText.split('\n').filter(line => line.trim() !== '');

    for (const line of lines) {
        const snIndex = line.toUpperCase().lastIndexOf('SN');
        let description = "";
        let serialsStr = line;

        if (snIndex !== -1) {
            description = line.substring(0, snIndex);
            serialsStr = line.substring(snIndex + 2).trim();
        }

        // Tokenize the string by replacing all delimiters with spaces and then splitting.
        // This handles commas, slashes, ampersands, and spaces gracefully.
        let items = serialsStr.replace(/,/g, ' ').replace(/\//g, ' ').replace(/&/g, ' ').split(/\s+/).filter(Boolean);

        let expandedSerials = [];
        let lastPrefix = "";

        for (let i = 0; i < items.length; i++) {
            let currentItem = items[i];

            // Handle ranges like "449 to 482"
            if (i > 0 && items[i - 1].toLowerCase() === 'to') {
                let startItem = expandedSerials.pop(); // The item before "to" is the start of our range
                let endItem = currentItem;

                let startPrefixMatch = startItem.match(/^(\D*)(\d+)$/);
                if (startPrefixMatch) {
                    let prefix = startPrefixMatch[1];
                    let startNum = parseInt(startPrefixMatch[2], 10);
                    let endNum = parseInt(endItem, 10);

                    if (!isNaN(startNum) && !isNaN(endNum)) {
                        expandedSerials.push(startItem); // Add the start item back
                        lastPrefix = prefix;
                        // Add all numbers from start+1 to end
                        for (let j = startNum + 1; j <= endNum; j++) {
                            expandedSerials.push(prefix + j);
                        }
                    }
                }
                 continue; // We've processed the range end, so skip to the next item
            }
            
            if (currentItem.toLowerCase() === 'to') {
                continue; // Skip the "to" keyword itself
            }

            // Handle individual items (not part of a range)
            if (/^\d+$/.test(currentItem)) {
                // Item is purely numeric, so prepend the last known prefix
                if (lastPrefix) {
                    expandedSerials.push(lastPrefix + currentItem);
                }
            } else {
                // Item contains non-digits, so it's a full serial number
                expandedSerials.push(currentItem);
                // Update the last known prefix for any subsequent numeric-only items
                const potentialPrefix = currentItem.replace(/\d+$/, '');
                if (potentialPrefix !== currentItem) { // Ensure a prefix was actually found
                    lastPrefix = potentialPrefix;
                }
            }
        }
        
        // Format the final output lines with the original description
        expandedSerials.forEach(sn => {
             if (description) {
                finalOutput.push(`${description.trim()} SN: ${sn}`);
             } else {
                finalOutput.push(sn);
             }
        });
    }

    return finalOutput;
}

