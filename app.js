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
        // ... (This function remains unchanged from previous versions)
    }

    async function cleanseAndPasteData() {
        // ... (This function remains unchanged from previous versions)
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
        const apiKey = document.getElementById('gemini-api-key').value.trim();
        const input = document.getElementById('gemini-input').value.trim();
        const status = document.getElementById('gemini-status');
        const button = document.getElementById('process-gemini');
        const outputContainer = document.getElementById('gemini-output-container');
        const outputTextarea = document.getElementById('gemini-output');

        outputContainer.style.display = 'none';
        outputTextarea.value = '';

        if (!apiKey) {
            status.textContent = 'Please enter your Gemini API key.';
            setTimeout(() => { status.textContent = ''; }, 3000);
            return;
        }

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
        // ... (This function remains unchanged from the previous version)
    }

})();

