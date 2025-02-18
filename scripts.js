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

// Populate file name select dropdown
function populateFileNameSelect() {
    fileNameSelect.innerHTML = '<option value="">Select a file name...</option>';
    
    const fileNameColumn = Object.keys(excelData[0])[0]; // Get first column name
    const uniqueFileNames = [...new Set(excelData.map(row => row[fileNameColumn]))];
    
    uniqueFileNames.forEach(fileName => {
        const option = document.createElement('option');
        option.value = fileName;
        option.textContent = fileName;
        fileNameSelect.appendChild(option);
    });
}

// Update unique file count
function updateUniqueFileCount() {
    const fileNameColumn = Object.keys(excelData[0])[0];
    const uniqueFileNames = new Set(excelData.map(row => row[fileNameColumn]));
    uniqueCountDisplay.textContent = uniqueFileNames.size;
}

// Handle file name selection
function handleFileNameChange() {
    const selectedFileName = fileNameSelect.value;
    if (!selectedFileName) return;

    const fileNameColumn = Object.keys(excelData[0])[0];
    const filteredData = excelData.filter(row => row[fileNameColumn] === selectedFileName);
    displayPreview(filteredData);
}

// Display preview table
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

// Remove File Name column from data
function removeFileNameColumn(data) {
    return data.map(row => {
        const newRow = { ...row };
        delete newRow[Object.keys(newRow)[0]]; // Remove first column (File Name)
        return newRow;
    });
}

// Process and download files
async function processAndDownloadFiles() {
    const fileNameColumn = Object.keys(excelData[0])[0];
    const uniqueFileNames = [...new Set(excelData.map(row => row[fileNameColumn]))];
    
    // Create a new ZIP file
    const zip = new JSZip();
    
    // Process each unique file name
    uniqueFileNames.forEach(fileName => {
        // Filter data for this file name
        const filteredData = excelData.filter(row => row[fileNameColumn] === fileName);
        
        // Remove File Name column from the filtered data
        const processedData = removeFileNameColumn(filteredData);
        
        // Create a new workbook for this data
        const ws = XLSX.utils.json_to_sheet(processedData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
        
        // Convert workbook to binary string
        const excelBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        
        // Add file to ZIP
        zip.file(`${fileName}.xlsx`, excelBuffer);
    });
    
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