
let candidateData = null;

let labConfig = null;

 

document.addEventListener('DOMContentLoaded', function() {

    document.getElementById('candidateFile').addEventListener('change', handleCandidateFile);

    document.getElementById('labFile').addEventListener('change', handleLabFile);

    document.getElementById('processButton').addEventListener('click', processAllocation);

    document.getElementById('downloadButton').addEventListener('click', downloadResults);

});

 

function handleCandidateFile(e) {

    const file = e.target.files[0];

    readExcelFile(file, 'candidate');

}

 

function handleLabFile(e) {

    const file = e.target.files[0];

    readExcelFile(file, 'lab');

}

 

function readExcelFile(file, type) {

    const reader = new FileReader();

    reader.onload = function(e) {

        const data = new Uint8Array(e.target.result);

        const workbook = XLSX.read(data, {type: 'array'});

        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];

        const jsonData = XLSX.utils.sheet_to_json(firstSheet);

 

        if (type === 'candidate') {

            candidateData = jsonData;

        } else {

            labConfig = jsonData;

        }

 

        checkEnableProcess();

        showMessage(`${type === 'candidate' ? 'Candidate' : 'Lab'} data loaded successfully`, 'success');

    };

    reader.readAsArrayBuffer(file);

}

 

function checkEnableProcess() {

    const processButton = document.getElementById('processButton');

    processButton.disabled = !(candidateData && labConfig);

}

 

function showMessage(message, type) {

    const messageArea = document.getElementById('messageArea');

    messageArea.textContent = message;

    messageArea.className = type;

    messageArea.style.display = 'block';

}

 

function validateBatchCount(candidates, labConfig) {

    // Group candidates by venue, date, and batch

    const batchGroups = {};

    candidates.forEach(candidate => {

        const key = `${candidate['Venue Code']}_${candidate['Exam Date']}_${candidate.Batch}`;

        if (!batchGroups[key]) {

            batchGroups[key] = {

                count: 0,

                venueCode: candidate['Venue Code'],

                venueName: candidate['Venue Name'],

                date: candidate['Exam Date'],

                batch: candidate.Batch,

                city: candidate.City

            };

        }

        batchGroups[key].count++;

    });

 

    // Calculate total lab capacity for each venue

    const labCapacities = {};

    labConfig.forEach(lab => {

        const key = lab['Centre Code'];

        if (!labCapacities[key]) {

            labCapacities[key] = {

                totalCapacity: 0,

                venueName: lab['Exam Centre Name'],

                city: lab.City,

                labs: []

            };

        }

        labCapacities[key].totalCapacity += parseInt(lab.Count) || 0;

        labCapacities[key].labs.push({

            labNo: lab['Lab No'],

            capacity: parseInt(lab.Count) || 0

        });

    });

 

    // Check for capacity violations

    const errors = [];

    Object.values(batchGroups).forEach(group => {

        const venueCapacity = labCapacities[group.venueCode];

       

        if (!venueCapacity) {

            errors.push({

                venue: group.venueName,

                date: group.date,

                batch: group.batch,

                count: group.count,

                capacity: 0,

                city: group.city,

                message: 'No lab configuration found for this venue'

            });

            return;

        }

 

        if (group.count > venueCapacity.totalCapacity) {

            let labDetails = venueCapacity.labs.map(lab =>

                `Lab ${lab.labNo}: ${lab.capacity} seats`

            ).join(', ');

           

            errors.push({

                venue: group.venueName,

                date: group.date,

                batch: group.batch,

                count: group.count,

                capacity: venueCapacity.totalCapacity,

                city: group.city,

                message: `Batch size (${group.count}) exceeds venue capacity (${venueCapacity.totalCapacity}). Available labs: ${labDetails}`

            });

        }

    });

 

    return errors;

}

 

// function processAllocation() {

//     try {

//         // Sort candidates by City, Date, Batch, and False No

//         const sortedCandidates = [...candidateData].sort((a, b) => {

//             if (a.City !== b.City) return a.City.localeCompare(b.City);

//             if (a['Exam Date'] !== b['Exam Date']) return new Date(a['Exam Date']) - new Date(b['Exam Date']);

//             if (a.Batch !== b.Batch) return a.Batch.localeCompare(b.Batch);

//             return (parseInt(a['False No']) || 0) - (parseInt(b['False No']) || 0);

//         });

 

//         // Validate batch counts

//         const validationErrors = validateBatchCount(sortedCandidates, labConfig);

//         if (validationErrors.length > 0) {

//             let errorMessage = 'Capacity Validation Errors:\n\n';

//             validationErrors.forEach(error => {

//                 errorMessage += `City: ${error.city}\n`;

//                 errorMessage += `Center: ${error.venue}\n`;

//                 errorMessage += `Date: ${error.date}\n`;

//                 errorMessage += `Batch: ${error.batch}\n`;

//                 errorMessage += `Batch Count: ${error.count}\n`;

//                 errorMessage += `Lab Capacity: ${error.capacity}\n`;

//                 errorMessage += `Issue: ${error.message}\n\n`;

//             });

//             throw new Error(errorMessage);

//         }

 

//         // Group and sort labs by venue

//         const venueLabsMap = {};

//         labConfig.forEach(lab => {

//             const key = `${lab['Centre Code']}`;

//             if (!venueLabsMap[key]) {

//                 venueLabsMap[key] = [];

//             }

//             venueLabsMap[key].push({

//                 ...lab,

//                 Count: parseInt(lab.Count) || 0

//             });

//         });

 

//        // Sort labs within each venue by Lab No

// Object.values(venueLabsMap).forEach(labs => {

//     labs.sort((a, b) => {

//         // Convert Lab No to string and remove any non-numeric characters

//         const labNoA = String(a['Lab No']).replace(/\D/g, '');

//         const labNoB = String(b['Lab No']).replace(/\D/g, '');

       

//         // Compare as numbers

//         return parseInt(labNoA) - parseInt(labNoB);

//     });

// });

 

//         // Track current position in each lab

//         const labPositions = {};

//         let currentDate = '';

//         let currentBatch = '';

 

//         // Allocate seats

//         const allocatedResults = sortedCandidates.map(candidate => {

//             // Reset lab positions when date or batch changes

//             if (currentDate !== candidate['Exam Date'] || currentBatch !== candidate.Batch) {

//                 currentDate = candidate['Exam Date'];

//                 currentBatch = candidate.Batch;

//                 Object.keys(labPositions).forEach(key => {

//                     labPositions[key] = 0;

//                 });

//             }

 

//             const venueLabs = venueLabsMap[candidate['Venue Code']] || [];

//             let allocated = false;

//             let allocation = null;

 

//             // Sequential lab allocation

//             for (const lab of venueLabs) {

//                 const key = `${lab['Centre Code']}_${lab['Lab No']}`;

//                 const currentCount = labPositions[key] || 0;

 

//                 if (currentCount < lab.Count) {

//                     // Allocate to this lab

//                     labPositions[key] = currentCount + 1;

//                     allocated = true;

//                     allocation = {

//                         ...candidate,

//                         'Building Name': lab['Building Name'],

//                         'Floor Name': lab['Floor Name'],

//                         'Lab Name': lab['Lab Name'],

//                         'Lab No': lab['Lab No'],

//                         'Server 1': lab['Server 1'],

//                         'Seat No': currentCount + 1

//                     };

//                     break;

//                 }

//             }

 

//             if (!allocated) {

//                 throw new Error(`Unable to allocate seat for candidate ${candidate['Candidate Id']} in ${candidate['Venue Name']}`);

//             }

 

//             return allocation;

//         });

 

//         // Display and enable download

//         displayResults(allocatedResults);

//         showMessage('Seat allocation completed successfully', 'success');

//         document.getElementById('downloadButton').style.display = 'block';

//         document.getElementById('resultTable').style.display = 'table';

//     } catch (error) {

//         showMessage(error.message, 'error');

//         document.getElementById('messageArea').style.whiteSpace = 'pre-line';

//     }

// }

 

function processAllocation() {

    try {

        // Sort candidates by City, Date, Batch, and False No

        const sortedCandidates = [...candidateData].sort((a, b) => {

            if (a.City !== b.City) return a.City.localeCompare(b.City);

            if (a['Exam Date'] !== b['Exam Date']) return new Date(a['Exam Date']) - new Date(b['Exam Date']);

            if (a.Batch !== b.Batch) return a.Batch.localeCompare(b.Batch);

            return (parseInt(a['False No']) || 0) - (parseInt(b['False No']) || 0);

        });

 

        // Validate batch counts

        const validationErrors = validateBatchCount(sortedCandidates, labConfig);

        if (validationErrors.length > 0) {

            let errorMessage = 'Capacity Validation Errors:\n\n';

            validationErrors.forEach(error => {

                errorMessage += `City: ${error.city}\n`;

                errorMessage += `Center: ${error.venue}\n`;

                errorMessage += `Date: ${error.date}\n`;

                errorMessage += `Batch: ${error.batch}\n`;

                errorMessage += `Batch Count: ${error.count}\n`;

                errorMessage += `Lab Capacity: ${error.capacity}\n`;

                errorMessage += `Issue: ${error.message}\n\n`;

            });

            throw new Error(errorMessage);

        }

 

        // Group and sort labs by venue

        const venueLabsMap = {};

        labConfig.forEach(lab => {

            const key = `${lab['Centre Code']}`;

            if (!venueLabsMap[key]) {

                venueLabsMap[key] = [];

            }

            venueLabsMap[key].push({

                ...lab,

                Count: parseInt(lab.Count) || 0

            });

        });

 

        // Sort labs within each venue by Lab No

        Object.values(venueLabsMap).forEach(labs => {

            labs.sort((a, b) => {

                const labNoA = String(a['Lab No']).replace(/\D/g, '');

                const labNoB = String(b['Lab No']).replace(/\D/g, '');

                return parseInt(labNoA) - parseInt(labNoB);

            });

        });

 

        // Track batch-wise seat numbering

        let currentBatchStart = 1;

        let currentDate = '';

        let currentBatch = '';

        let currentVenue = '';

 

        // Allocate seats with continuous numbering within batch

        const allocatedResults = sortedCandidates.map(candidate => {

            // Reset seat numbering when date, batch, or venue changes

            if (currentDate !== candidate['Exam Date'] ||

                currentBatch !== candidate.Batch ||

                currentVenue !== candidate['Venue Code']) {

                currentDate = candidate['Exam Date'];

                currentBatch = candidate.Batch;

                currentVenue = candidate['Venue Code'];

                currentBatchStart = 1;

            }

 

            const venueLabs = venueLabsMap[candidate['Venue Code']] || [];

            let allocated = false;

            let allocation = null;

            let seatRangeStart = currentBatchStart;

 

            // Find the appropriate lab for current seat number

            for (const lab of venueLabs) {

                const labCapacity = lab.Count;

                if (currentBatchStart <= seatRangeStart + labCapacity - 1) {

                    allocated = true;

                    allocation = {

                        ...candidate,

                        'Building Name': lab['Building Name'],

                        'Floor Name': lab['Floor Name'],

                        'Lab Name': lab['Lab Name'],

                        'Lab No': lab['Lab No'],

                        'Server 1': lab['Server'],

                        'Seat No': currentBatchStart

                    };

                    currentBatchStart++;

                    break;

                }

                seatRangeStart += labCapacity;

            }

 

            if (!allocated) {

                throw new Error(

                    `Unable to allocate seat for candidate ${candidate['Candidate Id']} ` +

                    `in ${candidate['Venue Name']} (Batch ${candidate.Batch})`

                );

            }

 

            return allocation;

        });

 

        // Display and enable download

        displayResults(allocatedResults);

        showMessage('Seat allocation completed successfully', 'success');

        document.getElementById('downloadButton').style.display = 'block';

        document.getElementById('resultTable').style.display = 'table';

 

        // Log allocation details for verification

        console.log('Allocation Details:');

        const grouped = groupBy(allocatedResults, 'Lab No');

        Object.entries(grouped).forEach(([labNo, candidates]) => {

            console.log(`Lab ${labNo}: Seats ${Math.min(...candidates.map(c => c['Seat No']))} to ${Math.max(...candidates.map(c => c['Seat No']))}`);

        });

 

    } catch (error) {

        showMessage(error.message, 'error');

        document.getElementById('messageArea').style.whiteSpace = 'pre-line';

    }

}

 

// Utility function to group results

function groupBy(array, key) {

    return array.reduce((result, currentValue) => {

        (result[currentValue[key]] = result[currentValue[key]] || []).push(currentValue);

        return result;

    }, {});

}

 

function displayResults(results) {

    const tbody = document.getElementById('resultBody');

    tbody.innerHTML = '';

 

    results.forEach(result => {

        const row = document.createElement('tr');

        const columns = [

            'Candidate Id', 'Candidate Email', 'Venue Code', 'Venue Name',

            'City', 'Exam Date', 'Exam Day', 'Batch', 'False No',

            'Building Name', 'Floor Name', 'Lab Name', 'Lab No', 'Server 1', 'Seat No'

        ];

       

        columns.forEach(column => {

            const td = document.createElement('td');

            td.textContent = result[column] || '';

            row.appendChild(td);

        });

 

        tbody.appendChild(row);

    });

}





function downloadResults() {

    const table = document.getElementById('resultTable');

    const ws = XLSX.utils.table_to_sheet(table);

    const wb = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(wb, ws, 'Seat Allocation');

    XLSX.writeFile(wb, 'seat_allocation_results.xlsx');

}


