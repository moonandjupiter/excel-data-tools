Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        console.log("Add-in is ready for Excel!");
    }
});

const tabs = ['rowInserter', 'dataCleanser'];

function switchTab(selectedTab) {
    tabs.forEach(tab => {
        const tabButton = document.getElementById(`tab-${tab}`);
        const tabContent = document.getElementById(`content-${tab}`);
        if (tab === selectedTab) {
            tabButton.classList.add('tab-active');
            tabButton.classList.remove('border-transparent', 'text-gray-500', 'hover:text-gray-700', 'hover:border-gray-300');
            tabContent.classList.remove('hidden');
        } else {
            tabButton.classList.remove('tab-active');
            tabButton.classList.add('border-transparent', 'text-gray-500', 'hover:text-gray-700', 'hover:border-gray-300');
            tabContent.classList.add('hidden');
        }
    });
}

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
            const entireRow = range.getEntireRow();
            
            for (let i = 0; i < count; i++) {
                 entireRow.insert(Excel.InsertShiftDirection.down);
            }

            await context.sync();
            status.textContent = `${count} row(s) inserted successfully!`;
            setTimeout(() => { status.textContent = ''; }, 3000);
        });
    } catch (error) {
        console.error(error);
        status.textContent = 'Error inserting rows.';
        setTimeout(() => { status.textContent = ''; }, 3000);
    }
}


async function cleanseAndPasteData() {
    const status = document.getElementById('status');
    const cleansedData = getCleansedData();

    if (!cleansedData || cleansedData.length === 0) {
        status.textContent = 'Nothing to paste.';
        setTimeout(() => { status.textContent = ''; }, 3000);
        return;
    }

    try {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            const resizedRange = range.getResizedRange(cleansedData.length - 1, 0);
            resizedRange.values = cleansedData;
            resizedRange.select();
            
            await context.sync();
            status.textContent = 'Data pasted successfully!';
            setTimeout(() => { status.textContent = ''; }, 3000);
        });
    } catch (error) {
        console.error(error);
        status.textContent = 'Error pasting data.';
        setTimeout(() => { status.textContent = ''; }, 3000);
    }
}


function getCleansedData() {
    const rawData = document.getElementById('rawData').value;
    const lines = rawData.split('\n');
    let result = [];

    lines.forEach(line => {
        if (!line.trim()) return;

        if (line.toLowerCase().includes('sn:')) {
            const parts = line.split(/sn:/i);
            const description = parts[0] + 'SN: ';
            const serials = parts[1].split(/[,&\s]+/).filter(s => s.trim() !== '');
            let lastFullSerial = '';

            serials.forEach(serial => {
                 let currentSerial = serial.trim();
                 if (lastFullSerial && !isNaN(currentSerial) && lastFullSerial.match(/[^0-9]/)) {
                    const prefix = lastFullSerial.slice(0, lastFullSerial.search(/[0-9]+$/));
                    currentSerial = prefix + currentSerial;
                 } else {
                    const nextSerialIndex = serials.indexOf(serial) + 1;
                    if (nextSerialIndex < serials.length && isNaN(serials[nextSerialIndex].trim()) === false && currentSerial.match(/[a-zA-Z]/)) {
                       const potentialPrefix = currentSerial.slice(0, currentSerial.search(/[0-9]+$/));
                       if (potentialPrefix) lastFullSerial = currentSerial;
                    }
                 }
                result.push([description + currentSerial]);
                if(currentSerial.match(/[a-zA-Z]/)){
                   lastFullSerial = currentSerial;
                }
            });
        } else {
            const prefixMatch = line.match(/^([a-zA-Z0-9]+-)/);
            const prefix = prefixMatch ? prefixMatch[0] : '';
            const cleanedLine = line.replace(prefix, '').replace(/&/g, ',');
            const parts = cleanedLine.split(/[,&\s]+/).filter(p => p.trim() !== '');
            
            let i = 0;
            while(i < parts.length){
                let current = parts[i];
                 if (current.toLowerCase() === 'to' && i > 0 && i < parts.length - 1) {
                    let start = parseInt(parts[i - 1]);
                    let end = parseInt(parts[i + 1]);
                    if (!isNaN(start) && !isNaN(end)) {
                        for (let j = start + 1; j <= end; j++) {
                             result.push([prefix + j]);
                        }
                    }
                    i += 2; // Skip 'to' and the end number
                 } else {
                    let num = parseInt(current);
                    if(!isNaN(num)){
                        result.push([prefix + num]);
                    }
                    i++;
                 }
            }
        }
    });
    return result;
}
