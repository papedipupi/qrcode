document.addEventListener('DOMContentLoaded', () => {
    let originalData = [];
    let rawData = [];  // To store the raw data

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
                rawData = results.data;  // Store the raw data
                const data = results.data.map(row => ({
                    ...row,
                    'Date/time': convertUKtoET(row['Date/time'], row)
                }));
                originalData = data;
                processData(data);
                document.getElementById('download').style.display = 'block';
                document.getElementById('filter-dates').style.display = 'block';
                setupDatePickers(data);
                generateProcessedDataTable(data);
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

        processData(filteredData, startDate, endDate);
    });

    document.getElementById('download').addEventListener('click', () => {
        const workbook = XLSX.utils.book_new();

        // First Sheet: 4 Tables
        const sheetData = [];
        let currentRow = 0;

        // Extract and append table data to sheetData
        currentRow = appendTableToSheetData(sheetData, currentRow, 'Scans Over Time', 'Date', 'Total Scans', 'Unique Scans', document.querySelector('h2:nth-of-type(1) + table'));
        currentRow = appendTableToSheetData(sheetData, currentRow, 'Scans by Operating Systems', 'OS', 'Scans', '%', document.querySelector('h2:nth-of-type(2) + table'));
        currentRow = appendTableToSheetData(sheetData, currentRow, 'Scans by Top 5 Cities', 'City', 'Scans', '%', document.querySelector('h2:nth-of-type(3) + table'));
        currentRow = appendTableToSheetData(sheetData, currentRow, 'Scans by Top 5 Countries', 'Country', 'Scans', '%', document.querySelector('h2:nth-of-type(4) + table'));

        const tablesWorksheet = XLSX.utils.aoa_to_sheet(sheetData);
        applyStyles(tablesWorksheet, sheetData);
        XLSX.utils.book_append_sheet(workbook, tablesWorksheet, 'QR Code Scans Data');

        // Second Sheet: Processed Data
        const processedDataSheetData = [['Date/time', 'Country Name', 'Country ISO', 'City', 'Device', 'Operating System', 'Unique Visitor']];
        originalData.forEach(row => {
            processedDataSheetData.push([
                row['Date/time'],
                row['Country Name'],
                row['Country ISO'],
                row['City'],
                row['Device'],
                row['Operating System'],
                row['Unique Visitor']
            ]);
        });
        const processedDataSheet = XLSX.utils.aoa_to_sheet(processedDataSheetData);
        XLSX.utils.book_append_sheet(workbook, processedDataSheet, 'Processed Data');

        // Third Sheet: Raw Data
        const rawDataSheetData = [['Date/time', 'Country Name', 'Country ISO', 'City', 'Device', 'Operating System', 'Unique Visitor']];
        rawData.forEach(row => {
            rawDataSheetData.push([
                row['Date/time'],
                row['Country Name'],
                row['Country ISO'],
                row['City'],
                row['Device'],
                row['Operating System'],
                row['Unique Visitor']
            ]);
        });
        const rawDataSheet = XLSX.utils.aoa_to_sheet(rawDataSheetData);
        XLSX.utils.book_append_sheet(workbook, rawDataSheet, 'Raw Data');

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

function processData(data, startDate = null, endDate = null) {
    const tablesHTML = generateTables(data, startDate, endDate);
    document.getElementById('tables-container').innerHTML = tablesHTML;
}

function generateTables(data, startDate, endDate) {
    const dailyScansTable = generateDailyScansTable(data, startDate, endDate);
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

function generateDailyScansTable(data, startDate, endDate) {
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
        const dateKey = date.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: '2-digit' });

        dateCounts[dateKey] = (dateCounts[dateKey] || 0) + 1;
        if (row['Unique Visitor'] === '1') {
            uniqueCounts[dateKey] = (uniqueCounts[dateKey] || 0) + 1;
        }
    });

    if (!startDate || !endDate) {
        const dateKeys = Object.keys(dateCounts).sort((a, b) => new Date(a) - new Date(b));
        startDate = new Date(dateKeys[0]);
        endDate = new Date(dateKeys[dateKeys.length - 1]);
    }
    const daysDifference = (endDate - startDate) / (1000 * 60 * 60 * 24) + 1;

    if (daysDifference > 35) {
        const weekCounts = {};
        const weekUniqueCounts = {};

        let startOfWeek = new Date(startDate);
        let endOfWeek = new Date(startOfWeek);
        endOfWeek.setDate(startOfWeek.getDate() + 6); // Add 7 days (start date to end date)

        while (startOfWeek <= endDate) {
            let weekKey;
            if (endOfWeek <= endDate) {
                weekKey = `${startOfWeek.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: '2-digit' })} - ${endOfWeek.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: '2-digit' })}`;
            } else if (startOfWeek === endOfWeek) {
                weekKey = `${startOfWeek.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: '2-digit' })}`;
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

            startOfWeek = new Date(endOfWeek);
            startOfWeek.setDate(startOfWeek.getDate() + 1);
            endOfWeek = new Date(startOfWeek);
            endOfWeek.setDate(startOfWeek.getDate() + 6);
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
        const dailyCounts = appendMissingDays(dateCounts, uniqueCounts, startDate, endDate);

        const totalScans = Object.values(dailyCounts).reduce((a, b) => a + b.total, 0);
        const totalUniqueScans = Object.values(dailyCounts).reduce((a, b) => a + b.unique, 0);

        let table = '<table><thead><tr><th>Date</th><th>Total Scans</th><th>Unique Scans</th></tr></thead><tbody>';
        Object.keys(dailyCounts).forEach(date => {
            table += `<tr><td>${date}</td><td>${dailyCounts[date].total}</td><td>${dailyCounts[date].unique}</td></tr>`;
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
    const hoursOffset = -5; // Adjusting time by deducting 5 hours for standard time difference
    const etDate = new Date(ukDate.getTime() + hoursOffset * 60 * 60 * 1000);
    const etTimeString = etDate.toISOString().slice(0, 19).replace('T', ' '); // Format as 'YYYY-MM-DD HH:MM:SS'
    
    // Log the conversion to the conversion log container
    logConversionMessage(`UK Time: ${ukTimeString} => ET Time: ${etTimeString}`);
    
    return etTimeString;
}

function convertETtoUK(etTimeString) {
    const etDate = new Date(etTimeString);
    if (isNaN(etDate)) {
        logMessage(`Invalid ET time string encountered: ${etTimeString}`, true);
        return etTimeString;  // Return original if invalid date
    }
    const hoursOffset = 5; // Adjusting time by adding 5 hours for standard time difference
    const ukDate = new Date(etDate.getTime() + hoursOffset * 60 * 60 * 1000);
    return ukDate.toISOString().slice(0, 19).replace('T', ' ');
}

function logMessage(message, isError = false) {
    const logContainer = document.getElementById('log-container');
    const logEntry = document.createElement('p');
    logEntry.textContent = message;
    logEntry.style.color = isError ? 'red' : 'black';
    logContainer.appendChild(logEntry);
    console.log(message);
}

function logConversionMessage(message) {
    const conversionLogContainer = document.getElementById('conversion-log-container');
    const logEntry = document.createElement('p');
    logEntry.textContent = message;
    logEntry.style.color = 'blue'; // Different color for conversion log
    conversionLogContainer.appendChild(logEntry);
    console.log(message);
}

function setupDatePickers(data) {
    const dateStrings = data.map(row => row['Date/time']).filter(date => !isNaN(new Date(date)));
    const dates = dateStrings.map(dateString => new Date(dateString));
    const minDate = new Date(Math.min(...dates));
    const maxDate = new Date(Math.max(...dates));

    const earliestPossibleDate = new Date(minDate);
    earliestPossibleDate.setDate(earliestPossibleDate.getDate() - 34); // Allow selection of up to 34 days before the data starts

    const latestPossibleDate = new Date(maxDate);
    latestPossibleDate.setDate(latestPossibleDate.getDate() + 34); // Allow selection of up to 34 days after the data ends

    flatpickr("#start-date", {
        defaultDate: minDate,
        minDate: earliestPossibleDate,
        maxDate: latestPossibleDate,
        dateFormat: "Y-m-d"
    });

    flatpickr("#end-date", {
        defaultDate: maxDate,
        minDate: earliestPossibleDate,
        maxDate: latestPossibleDate,
        dateFormat: "Y-m-d"
    });
}

function appendMissingDays(dateCounts, uniqueCounts, startDate, endDate) {
    const dailyCounts = {};
    for (let d = new Date(startDate); d <= endDate; d.setDate(d.getDate() + 1)) {
        const key = d.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: '2-digit' });
        dailyCounts[key] = {
            total: dateCounts[key] || 0,
            unique: uniqueCounts[key] || 0
        };
    }
    return dailyCounts;
}

function generateProcessedDataTable(data) {
    let table = '<h2>Processed Data</h2><table><thead><tr>';
    const headers = ['Date/time', 'Country Name', 'Country ISO', 'City', 'Device', 'Operating System', 'Unique Visitor'];
    headers.forEach(header => {
        table += `<th>${header}</th>`;
    });
    table += '</tr></thead><tbody>';

    data.forEach(row => {
        table += '<tr>';
        headers.forEach(header => {
            table += `<td>${row[header]}</td>`;
        });
        table += '</tr>';
    });

    table += '</tbody></table>';
    document.getElementById('processed-data-table').innerHTML = table;
    logMessage(`Generated processed data table with ${data.length} rows.`);
}

logMessage("Script loaded successfully.");
