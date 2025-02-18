// Global variables to store data
let excelData = null;
let workbook = null;

// DOM Elements
const fileInput = document.getElementById('fileInput');
const filterSection = document.getElementById('filterSection');
const statsSection = document.getElementById('statsSection');
const fileNameSelect = document.getElementById('fileNameSelect');
const processBtn = document.getElementById('processBtn');
const previewTable = document.getElementById('previewTable');
const uniqueCountDisplay = document.getElementById('uniqueCount');

// Event Listeners
fileInput.addEventListener('change', handleFileUpload);
fileNameSelect.addEventListener('change', handleFileNameChange);
processBtn.addEventListener('click', processAndDownloadFiles);

// Handle file upload
async function handleFileUpload(e) {
    const file = e.target.files[0];
    if (!file) return;

    try {
        const data = await file.arrayBuffer();
        workbook = XLSX.read(data);
        
        // Get the first sheet
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        
        // Convert sheet to JSON
        excelData = XLSX.utils.sheet_to_json(firstSheet);
        
        if (excelData.length === 0) {
            alert('No data found in the Excel file');
            return;
        }

        // Check if first column is "File Name"
        const headers = Object.keys(excelData[0]);
        if (!headers[0].toLowerCase().includes('file name')) {
            alert('First column must be "File Name"');
            return;
        }

        // Show filter and stats sections
        filterSection.style.display = 'block';
        statsSection.style.display = 'block';
        
        // Populate file name select and update stats
        populateFileNameSelect();
        updateUniqueFileCount();
        
        // Display preview
        displayPreview(excelData);
        
        // Enable process button
        processBtn.disabled = false;
    } catch (error) {
        console.error('Error reading file:', error);
        alert('Error reading the Excel file. Please make sure it\'s a valid Excel file.');
    }
}

// Convert JSON data to CSV string
function convertToCSV(jsonData) {
    if (jsonData.length === 0) return '';
    
    const headers = Object.keys(jsonData[0]);
    const csvRows = [];

    // Add headers
    csvRows.push(headers.join(','));

    // Add data rows
    for (const row of jsonData) {
        const values = headers.map(header => {
            const value = row[header] ?? '';
            // Escape quotes and wrap in quotes if the value contains commas or quotes
            const escaped = String(value).replace(/"/g, '""');
            return value.toString().includes(',') || value.toString().includes('"') 
                ? `"${escaped}"` 
                : escaped;
        });
        csvRows.push(values.join(','));
    }

    return csvRows.join('\n');
}

// Other helper functions remain the same
function populateFileNameSelect() {
    fileNameSelect.innerHTML = '<option value="">Select a file name...</option>';
    
    const fileNameColumn = Object.keys(excelData[0])[0];
    const uniqueFileNames = [...new Set(excelData.map(row => row[fileNameColumn]))];
    
    uniqueFileNames.forEach(fileName => {
        const option = document.createElement('option');
        option.value = fileName;
        option.textContent = fileName;
        fileNameSelect.appendChild(option);
    });
}

function updateUniqueFileCount() {
    const fileNameColumn = Object.keys(excelData[0])[0];
    const uniqueFileNames = new Set(excelData.map(row => row[fileNameColumn]));
    uniqueCountDisplay.textContent = uniqueFileNames.size;
}

function handleFileNameChange() {
    const selectedFileName = fileNameSelect.value;
    if (!selectedFileName) return;

    const fileNameColumn = Object.keys(excelData[0])[0];
    const filteredData = excelData.filter(row => row[fileNameColumn] === selectedFileName);
    displayPreview(filteredData);
}

function displayPreview(data, limit = 5) {
    if (!data || data.length === 0) {
        previewTable.innerHTML = '<p>No data to preview</p>';
        return;
    }

    const headers = Object.keys(data[0]);
    const previewData = data.slice(0, limit);

    let tableHTML = `
        <table>
            <thead>
                <tr>
                    ${headers.map(header => `<th>${header}</th>`).join('')}
                </tr>
            </thead>
            <tbody>
                ${previewData.map(row => `
                    <tr>
                        ${headers.map(header => `<td>${row[header]}</td>`).join('')}
                    </tr>
                `).join('')}
            </tbody>
        </table>
    `;

    previewTable.innerHTML = tableHTML;
}

function removeFileNameColumn(data) {
    return data.map(row => {
        const newRow = { ...row };
        delete newRow[Object.keys(newRow)[0]];
        return newRow;
    });
}

// Updated process and download function for CSV
// async function processAndDownloadFiles() {
//     const fileNameColumn = Object.keys(excelData[0])[0];
//     const uniqueFileNames = [...new Set(excelData.map(row => row[fileNameColumn]))];
    
//     // Create a new ZIP file
//     const zip = new JSZip();
    
//     // Process each unique file name
//     uniqueFileNames.forEach(fileName => {
//         // Filter data for this file name
//         const filteredData = excelData.filter(row => row[fileNameColumn] === fileName);
        
//         // Remove File Name column from the filtered data
//         const processedData = removeFileNameColumn(filteredData);
        
//         // Convert to CSV
//         const csvContent = convertToCSV(processedData);
        
//         // Add file to ZIP
//         zip.file(`${fileName}.csv`, csvContent);
//     });
    
//     try {
//         // Generate ZIP file
//         const content = await zip.generateAsync({ type: 'blob' });
        
//         // Create download link
//         const downloadLink = document.createElement('a');
//         downloadLink.href = URL.createObjectURL(content);
//         downloadLink.download = 'processed_files.zip';
        
//         // Trigger download
//         document.body.appendChild(downloadLink);
//         downloadLink.click();
//         document.body.removeChild(downloadLink);
//     } catch (error) {
//         console.error('Error creating ZIP file:', error);
//         alert('Error creating ZIP file');
//     }
// }

// ... (previous code remains the same until processAndDownloadFiles function)

async function processAndDownloadFiles() {
    const fileNameColumn = Object.keys(excelData[0])[0];
    const uniqueFileNames = [...new Set(excelData.map(row => row[fileNameColumn]))];
    
    // Create a new ZIP file
    const zip = new JSZip();
    
    // Prepare statistics data
    const statsData = [];
    
    // Process each unique file name
    uniqueFileNames.forEach(fileName => {
        // Filter data for this file name
        const filteredData = excelData.filter(row => row[fileNameColumn] === fileName);
        
        // Remove File Name column from the filtered data
        const processedData = removeFileNameColumn(filteredData);
        
        // Convert to CSV
        const csvContent = convertToCSV(processedData);
        
        // Add file to ZIP
        zip.file(`${fileName}.csv`, csvContent);

        // Add statistics (subtract 1 for header row)
        statsData.push({
            'File Name': fileName,
            'Number of Entries': filteredData.length
        });
    });

    // Create statistics Excel file
    const ws = XLSX.utils.json_to_sheet(statsData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Statistics");
    
    // Convert Excel file to binary
    const excelBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    
    // Add Excel file to ZIP
    zip.file('file_statistics.xlsx', excelBuffer);
    
    try {
        // Generate ZIP file
        const content = await zip.generateAsync({ type: 'blob' });
        
        // Create download link
        const downloadLink = document.createElement('a');
        downloadLink.href = URL.createObjectURL(content);
        downloadLink.download = 'processed_files.zip';
        
        // Trigger download
        document.body.appendChild(downloadLink);
        downloadLink.click();
        document.body.removeChild(downloadLink);
    } catch (error) {
        console.error('Error creating ZIP file:', error);
        alert('Error creating ZIP file');
    }
}