// Global variables for storing data and selections
let sheetData = []; // To hold the fetched data from the sheet
let selectedRows = []; // To track selected rows
let selectedColumns = []; // To track selected columns

// Function to apply operations based on user selections
document.getElementById('apply-operation').onclick = function() {
    const primaryColumn = document.getElementById('primary-column').value;
    const rangeColumns = document.getElementById('operation-columns').value.split(',');
    const rowRangeFrom = parseInt(document.getElementById('row-range-from').value);
    const rowRangeTo = parseInt(document.getElementById('row-range-to').value);
    const operationType = document.getElementById('operation-type').value;
    const operation = document.getElementById('operation').value;
    const contentType = document.getElementById('content-type').value;

    // Reset selected rows and columns
    selectedRows = [];
    selectedColumns = [];

    // Highlighting logic for selected ranges
    highlightRanges(rangeColumns, rowRangeFrom, rowRangeTo);
    
    // Perform operation based on selected filters
    const filteredData = filterData(sheetData, primaryColumn, rangeColumns, rowRangeFrom, rowRangeTo, operation, contentType);
    displayData(filteredData);
};

// Function to highlight selected ranges in the displayed data
function highlightRanges(rangeColumns, rowRangeFrom, rowRangeTo) {
    const sheetContent = document.getElementById('sheet-content');
    const rows = sheetContent.querySelectorAll('tr');
    
    rows.forEach((row, rowIndex) => {
        if (rowIndex >= rowRangeFrom - 1 && rowIndex <= rowRangeTo - 1) {
            rangeColumns.forEach(col => {
                const cell = row.querySelector(`td:nth-child(${col.trim().charCodeAt(0) - 65 + 1})`);
                if (cell) {
                    cell.style.backgroundColor = 'lightyellow';
                    selectedRows.push(rowIndex);
                    selectedColumns.push(col.trim());
                }
            });
        }
    });
}

// Function to filter data based on user-defined criteria
function filterData(data, primaryColumn, rangeColumns, rowRangeFrom, rowRangeTo, operation, contentType) {
    return data.filter((row, rowIndex) => {
        if (rowIndex < rowRangeFrom - 1 || rowIndex > rowRangeTo - 1) return false;

        const primaryValue = row[primaryColumn];
        const isNullCheck = (operation === 'null') ? isNull(row) : !isNull(row);
        const matchesContentType = checkContentType(row, contentType);

        return isNullCheck && matchesContentType;
    });
}

// Function to check if a row meets the null/not-null criteria
function isNull(row) {
    return Object.values(row).every(value => value === '' || value === null);
}

// Function to check if the row's content matches the selected content type
function checkContentType(row, contentType) {
    return Object.values(row).every(value => {
        switch (contentType) {
            case 'word':
                return /^[a-zA-Z]+$/.test(value);
            case 'number':
                return /^\d+$/.test(value);
            case 'link':
                return /^(https?|ftp):\/\/[^\s/$.?#].[^\s]*$/.test(value);
            case 'all':
            default:
                return true; // No restriction
        }
    });
}

// Function to display filtered data in the content area
function displayData(data) {
    const sheetContent = document.getElementById('sheet-content');
    sheetContent.innerHTML = ''; // Clear previous content

    const table = document.createElement('table');
    const headerRow = document.createElement('tr');
    Object.keys(data[0]).forEach(header => {
        const th = document.createElement('th');
        th.innerText = header;
        headerRow.appendChild(th);
    });
    table.appendChild(headerRow);

    data.forEach(row => {
        const rowElement = document.createElement('tr');
        Object.values(row).forEach(cell => {
            const td = document.createElement('td');
            td.innerText = cell;
            rowElement.appendChild(td);
        });
        table.appendChild(rowElement);
    });

    sheetContent.appendChild(table);
}

// Function to initiate the download of the selected data
function downloadFile() {
    const filename = document.getElementById('filename').value || 'download';
    const fileFormat = document.getElementById('file-format').value;

    switch (fileFormat) {
        case 'xlsx':
            downloadAsExcel(filename);
            break;
        case 'csv':
            downloadAsCSV(filename);
            break;
        case 'pdf':
            downloadAsPDF(filename);
            break;
        case 'img':
            downloadAsImage(filename);
            break;
        case 'txt':
            downloadAsText(filename);
            break;
        default:
            console.error('Unsupported format');
    }

    closeDownloadModal();
}

// Function to download data as an Excel file
function downloadAsExcel(filename) {
    const worksheet = XLSX.utils.json_to_sheet(sheetData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
    XLSX.writeFile(workbook, `${filename}.xlsx`);
}

// Function to download data as a CSV file
function downloadAsCSV(filename) {
    const csvContent = sheetData.map(e => Object.values(e).join(",")).join("\n");
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.setAttribute('download', `${filename}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

// Placeholder function for downloading as a PDF
function downloadAsPDF(filename) {
    console.log("Downloading as PDF:", filename + ".pdf");
    // Implement PDF download logic here
}

// Placeholder function for downloading as an image
function downloadAsImage(filename) {
    console.log("Downloading as Image:", filename + ".png");
    // Implement image download logic here
}

// Placeholder function for downloading as a text file
function downloadAsText(filename) {
    const textContent = sheetData.map(e => JSON.stringify(e)).join("\n");
    const blob = new Blob([textContent], { type: 'text/plain;charset=utf-8;' });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.setAttribute('download', `${filename}.txt`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

// Event listeners for the buttons
document.getElementById('download-button').onclick = openDownloadModal;
document.getElementById('confirm-download').onclick = downloadFile;
document.getElementById('close-modal').onclick = closeDownloadModal;

// Functions to handle modal open and close
function openDownloadModal() {
    document.getElementById('download-modal').style.display = 'flex';
}

function closeDownloadModal() {
    document.getElementById('download-modal').style.display = 'none';
}

// Example function to fetch and process data from a Google Sheet or other source
async function fetchData(sheetUrl) {
    // Fetch data from Google Sheets or other data sources here
    // For now, we'll just simulate with dummy data
    sheetData = [
        { A: 'Name', B: 'Age', C: 'Link' },
        { A: 'Alice', B: '30', C: 'http://example.com' },
        { A: 'Bob', B: '25', C: '' },
        { A: '', B: '35', C: 'http://example.org' },
        { A: 'Charlie', B: '', C: 'http://example.net' }
    ];
    displayData(sheetData);
}

// Call fetchData with the Google Sheet URL or a similar source when the page loads
// fetchData('your-google-sheet-url');
