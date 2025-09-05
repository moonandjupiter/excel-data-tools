(function () {
    "use strict";

    const BACKEND_URL = "http://localhost:8000/process-data";

    document.addEventListener('DOMContentLoaded', initialize);

    function initialize() {
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

        // This ensures the first tab is visible on load.
        switchTab('inserter');

        Office.onReady();
    }

    function switchTab(tabName) {
        document.getElementById('inserter-content').classList.add('hidden');
        document.getElementById('cleanser-content').classList.add('hidden');
        document.getElementById('gemini-content').classList.add('hidden');
        
        document.getElementById('tab-inserter').classList.remove('active-tab');
        document.getElementById('tab-cleanser').classList.remove('active-tab');
        document.getElementById('tab-gemini').classList.remove('active-tab');
        
        document.getElementById(tabName + '-content').classList.remove('hidden');
        document.getElementById('tab-' + tabName).classList.add('active-tab');
    }

    async function generateAndInsertRows() {
        // This function's logic remains the same
    }

    async function cleanseAndPasteData() {
       // This function's logic remains the same
    }
    
    function copyGeminiResult() {
        // This function's logic remains the same
    }

    async function processWithGemini() {
        // This function's logic remains the same
    }

    function parseData(rawData) {
        // This function's logic remains the same
    }

})();

