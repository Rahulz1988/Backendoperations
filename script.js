// Global variables to store data
let excelData = null;
let workbook = null;

// DOM Elements
const fileInput = document.getElementById('fileInput');
const filterSection = document.getElementById('filterSection');
const columnSelect = document.getElementById('columnSelect');
const valueSelect = document.getElementById('valueSelect');
const downloadBtn = document.getElementById('downloadBtn');
const previewTable = document.getElementById('previewTable');

// Event Listeners
fileInput.addEventListener('change', handleFileUpload);
columnSelect.addEventListener('change', handleColumnChange);
valueSelect.addEventListener('change', handleValueChange);
downloadBtn.addEventListener('click', downloadFilteredData);

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

        // Show filter section
        filterSection.style.display = 'flex';
        
        // Populate column select
        populateColumnSelect();
        
        // Display preview
        displayPreview(excelData);
    } catch (error) {
        console.error('Error reading file:', error);
        alert('Error reading the Excel file. Please make sure it\'s a valid Excel file.');
    }
}

// Populate column select dropdown
function populateColumnSelect() {
    columnSelect.innerHTML = '<option value="">Select a column...</option>';
    
    const headers = Object.keys(excelData[0]);
    headers.forEach(header => {
        const option = document.createElement('option');
        option.value = header;
        option.textContent = header;
        columnSelect.appendChild(option);
    });
}

// Handle column selection
function handleColumnChange() {
    const selectedColumn = columnSelect.value;
    if (!selectedColumn) {
        valueSelect.innerHTML = '<option value="">Select a value...</option>';
        return;
    }

    // Get unique values for the selected column
    const uniqueValues = [...new Set(excelData.map(row => row[selectedColumn]))];
    
    // Populate value select
    valueSelect.innerHTML = '<option value="">Select a value...</option>';
    uniqueValues.forEach(value => {
        const option = document.createElement('option');
        option.value = value;
        option.textContent = value;
        valueSelect.appendChild(option);
    });
}

// Handle value selection
function handleValueChange() {
    downloadBtn.disabled = !valueSelect.value;
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

// Download filtered data
function downloadFilteredData() {
    const selectedColumn = columnSelect.value;
    const selectedValue = valueSelect.value;

    if (!selectedColumn || !selectedValue) return;

    // Filter data
    const filteredData = excelData.filter(row => row[selectedColumn] == selectedValue);

    if (filteredData.length === 0) {
        alert('No matching data found');
        return;
    }

    // Create new workbook with filtered data
    const ws = XLSX.utils.json_to_sheet(filteredData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Filtered Data');

    // Generate download
    XLSX.writeFile(wb, `filtered_data_${selectedColumn}_${selectedValue}.xlsx`);
}