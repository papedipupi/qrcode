document.addEventListener('DOMContentLoaded', () => {
    let originalData = [];

    document.getElementById('process').addEventListener('click', () => {
        const fileInput = document.getElementById('csv-file');
        if (fileInput.files.length === 0) {
            alert('Please select a CSV file to upload.');
            return;
        }

        const file = fileInput.files[0];
        logMessage(`Selected file: ${file.name}`);

        Papa.parse(file, {
            header: true,
            complete: function (results) {
                const data = results.data;
                originalData = data;
                processData(data);
                document.getElementById('download').style.display = 'block';
                document.getElementById('filter-dates').style.display = 'block';
                setupDatePickers(data);
            },
            error: function (error) {
                logMessage(`Error parsing CSV file: ${error.message}`, true);
            }
        });
    });

    document.getElementById('filter-dates').addEventListener('click', () => {
        const startDate = document.getElementById('start-date')._flatpickr.selectedDates[0];
        const endDate = document.getElementById('end-date')._flatpickr.selectedDates[0];

        if (!startDate || !endDate) {
            alert('Please select both start and end dates.');
            return;
        }

        const filteredData = originalData.filter(row => {
            const rowDate = new Date(row['Date/time']);
            return rowDate >= startDate && rowDate <= endDate;
        });

        processData(filteredData);
    });

    document.getElementById('download').addEventListener('click', () => {
        const workbook = XLSX.utils.book_new();
        const sheetData = [];
        let currentRow = 0;

        // Extract and append table data to sheetData
        currentRow = appendTableToSheetData(sheetData, currentRow, 'Scans Over Time', 'Date', 'Total Scans', 'Unique Scans', document.querySelector('h2:nth-of-type(1) + table'));
        currentRow = appendTableToSheetData(sheetData, currentRow, 'Scans by Operating Systems', 'OS', 'Scans', '%', document.querySelector('h2:nth-of-type(2) + table'));
        currentRow = appendTableToSheetData(sheetData, currentRow, 'Scans by Top 5 Cities', 'City', 'Scans', '%', document.querySelector('h2:nth-of-type(3) + table'));
        currentRow = appendTableToSheetData(sheetData, currentRow, 'Scans by Top 5 Countries', 'Country', 'Scans', '%', document.querySelector('h2:nth-of-type(4) + table'));

        const worksheet = XLSX.utils.aoa_to_sheet(sheetData);
        applyStyles(worksheet, sheetData);

        XLSX.utils.book_append_sheet(workbook, worksheet, 'QR Code Scans Data');
        XLSX.writeFile(workbook, 'qr_code_scans_data.xlsx');
        logMessage(`Excel file 'qr_code_scans_data.xlsx' has been downloaded.`);

        // Log Kawaii cat ASCII art
        logMessage(`
          /\\_/\\  
         ( o.o ) 
          > ^ < 
        `);
    });
});

function appendTableToSheetData(sheetData, currentRow, title, col1, col2, col3, table) {
    sheetData[currentRow] = [title];
    currentRow++;
    sheetData[currentRow] = [col1, col2, col3]; // Add headers
    currentRow++;

    table.querySelectorAll('tr').forEach((row, index) => {
        if (index === 0) return;  // Skip the table header row
        const rowData = [];
        row.querySelectorAll('td, th').forEach(cell => {
            rowData.push(cell.innerText);
        });
        sheetData[currentRow] = rowData;
        currentRow++;
    });
    sheetData[currentRow] = []; // Blank row to separate sections
    currentRow++;
    return currentRow;
}

function applyStyles(worksheet, sheetData) {
    const range = XLSX.utils.decode_range(worksheet['!ref']);

    for (let rowNum = range.s.r; rowNum <= range.e.r; rowNum++) {
        for (let colNum = range.s.c; colNum <= range.e.c; colNum++) {
            const cellAddress = { c: colNum, r: rowNum };
            const cellRef = XLSX.utils.encode_cell(cellAddress);

            if (!worksheet[cellRef]) continue;

            worksheet[cellRef].s = {
                border: {
                    top: { style: "thin", color: { auto: 1 } },
                    bottom: { style: "thin", color: { auto: 1 } },
                    left: { style: "thin", color: { auto: 1 } },
                    right: { style: "thin", color: { auto: 1 } }
                },
                alignment: {
                    vertical: "center",
                    horizontal: "center"
                }
            };

            // Bold headers and titles
            if (rowNum === 0 || sheetData[rowNum][0] === 'Scans Over Time' ||
                sheetData[rowNum][0] === 'Scans by Operating Systems' ||
                sheetData[rowNum][0] === 'Scans by Top 5 Cities' ||
                sheetData[rowNum][0] === 'Scans by Top 5 Countries' ||
                sheetData[rowNum][0] === 'Date' ||
                sheetData[rowNum][0] === 'OS' ||
                sheetData[rowNum][0] === 'City' ||
                sheetData[rowNum][0] === 'Country') {
                worksheet[cellRef].s.font = { bold: true };
            }
        }
    }
}

function processData(data) {
    const tablesHTML = generateTables(data);
    document.getElementById('tables-container').innerHTML = tablesHTML;
}

function generateTables(data) {
    const dailyScansTable = generateDailyScansTable(data);
    const osScansTable = generateOSScansTable(data);
    const cityScansTable = generateCityScansTable(data);
    const countryScansTable = generateCountryScansTable(data);
    return `
        <h2>Scans Over Time</h2>${dailyScansTable}
        <h2>Scans by Operating Systems</h2>${osScansTable}
        <h2>Scans by Top 5 Cities</h2>${cityScansTable}
        <h2>Scans by Top 5 Countries</h2>${countryScansTable}
    `;
}

function generateDailyScansTable(data) {
    let dateCounts = {};
    let uniqueCounts = {};

    data.forEach(row => {
        const dateString = row['Date/time'];
        const date = new Date(dateString);
        if (isNaN(date)) {
            logMessage(`Invalid date encountered: ${dateString}`, true);
            logMessage(`Row data: ${JSON.stringify(row)}`);
            return; // Skip invalid dates
        }
        logMessage(`Processing date: ${dateString}`);
        logMessage(`Row data: ${JSON.stringify(row)}`);
        const dateKey = date.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: '2-digit' });

        dateCounts[dateKey] = (dateCounts[dateKey] || 0) + 1;
        if (row['Unique Visitor'] === '1') {
            uniqueCounts[dateKey] = (uniqueCounts[dateKey] || 0) + 1;
        }
    });

    const dateKeys = Object.keys(dateCounts).sort((a, b) => new Date(a) - new Date(b));
    let startDate = new Date(dateKeys[0]);
    let endDate = new Date(dateKeys[dateKeys.length - 1]);
    startDate.setHours(0, 0, 0, 0); // Ensure time is at start of the day
    endDate.setHours(23, 59, 59, 999); // Ensure time is at the end of the day
    const daysDifference = (endDate - startDate) / (1000 * 60 * 60 * 24) + 1;

    if (daysDifference > 35) {
        const weekCounts = {};
        const weekUniqueCounts = {};

        let startOfWeek = new Date(startDate);
        let endOfWeek = new Date(startOfWeek);
        endOfWeek.setDate(startOfWeek.getDate() + (6 - startOfWeek.getDay()));

        const firstWeekKey = `${startOfWeek.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: '2-digit' })} - ${endOfWeek.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: '2-digit' })}`;
        let totalScans = 0;
        let totalUnique = 0;
        for (let d = new Date(startDate); d <= endOfWeek; d.setDate(d.getDate() + 1)) {
            const key = d.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: '2-digit' });
            totalScans += (dateCounts[key] || 0);
            totalUnique += (uniqueCounts[key] || 0);
        }
        weekCounts[firstWeekKey] = totalScans;
        weekUniqueCounts[firstWeekKey] = totalUnique;

        startDate = new Date(endOfWeek);
        startDate.setDate(startDate.getDate() + 1);
        while (startDate <= endDate) {
            startOfWeek = new Date(startDate);
            endOfWeek = new Date(startOfWeek);
            endOfWeek.setDate(startOfWeek.getDate() + 6);

            let weekKey;
            if (endOfWeek <= endDate) {
                weekKey = `${startOfWeek.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: '2-digit' })} - ${endOfWeek.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: '2-digit' })}`;
            } else {
                weekKey = `${startOfWeek.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: '2-digit' })} - ${endDate.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: '2-digit' })}`;
            }

            let totalScans = 0;
            let totalUnique = 0;
            for (let d = new Date(startOfWeek); d <= (endOfWeek <= endDate ? endOfWeek : endDate); d.setDate(d.getDate() + 1)) {
                const key = d.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: '2-digit' });
                totalScans += (dateCounts[key] || 0);
                totalUnique += (uniqueCounts[key] || 0);
            }
            weekCounts[weekKey] = totalScans;
            weekUniqueCounts[weekKey] = totalUnique;

            startDate = new Date(endOfWeek);
            startDate.setDate(startDate.getDate() + 1);
        }

        const totalScansOverall = Object.values(weekCounts).reduce((a, b) => a + b, 0);
        const totalUniqueScansOverall = Object.values(weekUniqueCounts).reduce((a, b) => a + b, 0);

        let table = '<table><thead><tr><th>Date</th><th>Total Scans</th><th>Unique Scans</th></tr></thead><tbody>';
        Object.keys(weekCounts).forEach(week => {
            table += `<tr><td>${week}</td><td>${weekCounts[week]}</td><td>${weekUniqueCounts[week] || 0}</td></tr>`;
        });
        table += `<tr class="total"><td>Total</td><td>${totalScansOverall}</td><td>${totalUniqueScansOverall}</td></tr>`;
        table += '</tbody></table>';
        logMessage(`Generated weekly scans table with total scans: ${totalScansOverall} and total unique scans: ${totalUniqueScansOverall}`);
        return table;
    } else {
        const totalScans = Object.values(dateCounts).reduce((a, b) => a + b, 0);
        const totalUniqueScans = Object.values(uniqueCounts).reduce((a, b) => a + b, 0);

        let table = '<table><thead><tr><th>Date</th><th>Total Scans</th><th>Unique Scans</th></tr></thead><tbody>';
        dateKeys.forEach(date => {
            table += `<tr><td>${date}</td><td>${dateCounts[date]}</td><td>${uniqueCounts[date] || 0}</td></tr>`;
        });
        table += `<tr class="total"><td>Total</td><td>${totalScans}</td><td>${totalUniqueScans}</td></tr>`;
        table += '</tbody></table>';
        logMessage(`Generated daily scans table with total scans: ${totalScans} and total unique scans: ${totalUniqueScans}`);
        return table;
    }
}

function generateOSScansTable(data) {
    const osCounts = {};

    data.forEach(row => {
        const os = row['Operating System'] || 'Unknown';  // Handle missing OS
        osCounts[os] = (osCounts[os] || 0) + 1;
    });

    const totalScans = data.length;
    let table = '<table><thead><tr><th>OS</th><th>Scans</th><th>%</th></thead><tbody>';
    Object.keys(osCounts).forEach(os => {
        const percentage = ((osCounts[os] / totalScans) * 100).toFixed(2);
        table += `<tr><td>${os}</td><td>${osCounts[os]}</td><td>${percentage}%</td></tr>`;
    });
    table += '</tbody></table>';
    logMessage(`Generated OS scans table.`);
    return table;
}

function generateCityScansTable(data) {
    const cityCounts = {};

    data.forEach(row => {
        const city = row['City'] || 'Unknown';  // Handle missing city
        cityCounts[city] = (cityCounts[city] || 0) + 1;
    });

    const totalScans = data.length;
    let table = '<table><thead><tr><th>City</th><th>Scans</th><th>%</th></thead><tbody>';
    Object.keys(cityCounts)
        .sort((a, b) => cityCounts[b] - cityCounts[a])
        .slice(0, 5)
        .forEach(city => {
            const percentage = ((cityCounts[city] / totalScans) * 100).toFixed(2);
            table += `<tr><td>${city}</td><td>${cityCounts[city]}</td><td>${percentage}%</td></tr>`;
        });
    table += '</tbody></table>';
    logMessage(`Generated city scans table.`);
    return table;
}

function generateCountryScansTable(data) {
    const countryCounts = {};

    data.forEach(row => {
        const country = row['Country Name'] || 'Unknown';  // Handle missing country
        countryCounts[country] = (countryCounts[country] || 0) + 1;
    });

    const totalScans = data.length;
    let table = '<table><thead><tr><th>Country</th><th>Scans</th><th>%</th></thead><tbody>';
    Object.keys(countryCounts)
        .sort((a, b) => countryCounts[b] - countryCounts[a])
        .slice(0, 5)
        .forEach(country => {
            const percentage = ((countryCounts[country] / totalScans) * 100).toFixed(2);
            table += `<tr><td>${country}</td><td>${countryCounts[country]}</td><td>${percentage}%</td></tr>`;
        });
    table += '</tbody></table>';
    logMessage(`Generated country scans table.`);
    return table;
}

function convertUKtoET(ukTimeString, row) {
    const ukDate = new Date(ukTimeString);
    if (isNaN(ukDate)) {
        logMessage(`Invalid UK time string encountered: ${ukTimeString}`, true);
        logMessage(`Row content presence: ${Object.values(row).some(value => value.trim() !== '')}`);
        logMessage(`Row data: ${JSON.stringify(row)}`);
        return ukTimeString;  // Return original if invalid date
    }
    const etOffset = -5 * 60; // ET is UTC-5
    const etDate = new Date(ukDate.getTime() + etOffset * 60 * 1000);
    return etDate.toISOString().slice(0, 19).replace('T', ' '); // Format as 'YYYY-MM-DD HH:MM:SS'
}

function logMessage(message, isError = false) {
    const logContainer = document.getElementById('log-container');
    const logEntry = document.createElement('p');
    logEntry.textContent = message;
    logEntry.style.color = isError ? 'red' : 'black';
    logContainer.appendChild(logEntry);
    console.log(message);
}

function setupDatePickers(data) {
    const dateStrings = data.map(row => row['Date/time']).filter(date => !isNaN(new Date(date)));
    const dates = dateStrings.map(dateString => new Date(dateString));
    const minDate = new Date(Math.min(...dates));
    const maxDate = new Date(Math.max(...dates));

    flatpickr("#start-date", {
        defaultDate: minDate,
        minDate: minDate,
        maxDate: maxDate,
        dateFormat: "Y-m-d"
    });

    flatpickr("#end-date", {
        defaultDate: maxDate,
        minDate: minDate,
        maxDate: maxDate,
        dateFormat: "Y-m-d"
    });
}

logMessage("Script loaded successfully.");
