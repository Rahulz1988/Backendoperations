// Track application state
const appState = {
    candidateData: null,
    labConfigData: null,
    labAllocationResult: null,
    seatAllocationResult: null,
  }
  
  // DOM Elements
  const candidateFileInput = document.getElementById("candidateFile")
  const labConfigFileInput = document.getElementById("labConfigFile")
  const labAllocatedFileInput = document.getElementById("labAllocatedFile")
  const allocateLabBtn = document.getElementById("allocateLabBtn")
  const allocateSeatsBtn = document.getElementById("allocateSeatsBtn")
  const downloadLabAllocationBtn = document.getElementById("downloadLabAllocationBtn")
  const downloadSeatAllocationBtn = document.getElementById("downloadSeatAllocationBtn")
  const labAllocationStatus = document.getElementById("labAllocationStatus")
  const seatAllocationStatus = document.getElementById("seatAllocationStatus")
  const labAllocationResults = document.getElementById("labAllocationResults")
  const seatAllocationResults = document.getElementById("seatAllocationResults")
  
  // Event Listeners for file inputs
  candidateFileInput.addEventListener("change", async (e) => {
    const file = e.target.files[0]
    if (!file) return
  
    try {
      showNotification(labAllocationStatus, "Reading candidate data...", "info")
      appState.candidateData = await readExcelFile(file)
      showNotification(labAllocationStatus, "Candidate data loaded successfully!", "success")
      checkEnableLabAllocation()
    } catch (error) {
      console.error("Error reading candidate data:", error)
      showNotification(labAllocationStatus, "Error reading candidate data: " + error.message, "error")
    }
  })
  
  labConfigFileInput.addEventListener("change", async (e) => {
    const file = e.target.files[0]
    if (!file) return
  
    try {
      showNotification(labAllocationStatus, "Reading lab configuration data...", "info")
      appState.labConfigData = await readExcelFile(file)
      showNotification(labAllocationStatus, "Lab configuration data loaded successfully!", "success")
      checkEnableLabAllocation()
    } catch (error) {
      console.error("Error reading lab configuration data:", error)
      showNotification(labAllocationStatus, "Error reading lab configuration data: " + error.message, "error")
    }
  })
  
  labAllocatedFileInput.addEventListener("change", async (e) => {
    const file = e.target.files[0]
    if (!file) return
  
    try {
      showNotification(seatAllocationStatus, "Reading lab allocated data...", "info")
      appState.labAllocationResult = await readExcelFile(file)
      showNotification(seatAllocationStatus, "Lab allocated data loaded successfully!", "success")
      allocateSeatsBtn.disabled = false
    } catch (error) {
      console.error("Error reading lab allocated data:", error)
      showNotification(seatAllocationStatus, "Error reading lab allocated data: " + error.message, "error")
    }
  })
  
  // Lab Allocation Button Click Handler
  allocateLabBtn.addEventListener("click", () => {
    try {
      showNotification(labAllocationStatus, "Allocating labs...", "info")
  
      // Sort candidate data
      const sortedCandidates = sortCandidateData(appState.candidateData)
  
      // Sort lab configuration data
      const sortedLabConfig = sortLabConfigData(appState.labConfigData)
  
      // Perform lab allocation
      appState.labAllocationResult = allocateLabs(sortedCandidates, sortedLabConfig)
  
      // Display results
      displayLabAllocationResults(appState.labAllocationResult)
  
      showNotification(labAllocationStatus, "Lab allocation completed successfully!", "success")
      downloadLabAllocationBtn.disabled = false
    } catch (error) {
      console.error("Error during lab allocation:", error)
      showNotification(labAllocationStatus, "Error during lab allocation: " + error.message, "error")
    }
  })
  
  // Seat Allocation Button Click Handler
  allocateSeatsBtn.addEventListener("click", () => {
    try {
      showNotification(seatAllocationStatus, "Allocating seats...", "info")
  
      // Perform seat allocation
      appState.seatAllocationResult = allocateSeats(appState.labAllocationResult)
  
      // Display results
      displaySeatAllocationResults(appState.seatAllocationResult)
  
      showNotification(seatAllocationStatus, "Seat allocation completed successfully!", "success")
      downloadSeatAllocationBtn.disabled = false
    } catch (error) {
      console.error("Error during seat allocation:", error)
      showNotification(seatAllocationStatus, "Error during seat allocation: " + error.message, "error")
    }
  })
  
  // Download Buttons Click Handlers
  downloadLabAllocationBtn.addEventListener("click", () => {
    downloadExcelFile(appState.labAllocationResult, "Lab_Allocation_Results.xlsx")
  })
  
  downloadSeatAllocationBtn.addEventListener("click", () => {
    downloadExcelFile(appState.seatAllocationResult, "Final_Seat_Allocation.xlsx")
  })
  
  // Helper Functions
  function checkEnableLabAllocation() {
    allocateLabBtn.disabled = !(appState.candidateData && appState.labConfigData)
  }
  
  function showNotification(element, message, type) {
    element.innerHTML = `<div class="notification ${type}">${message}</div>`
  }
  
  async function readExcelFile(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader()
  
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target.result)
          const workbook = XLSX.read(data, { type: "array" })
          const sheetName = workbook.SheetNames[0]
          const worksheet = workbook.Sheets[sheetName]
          const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 })
  
          if (jsonData.length < 2) {
            reject(new Error("File does not contain enough data"))
            return
          }
  
          const headers = jsonData[0]
          const rows = jsonData.slice(1)
  
          const result = rows.map((row) => {
            const obj = {}
            row.forEach((cell, index) => {
              if (index < headers.length) {
                obj[headers[index]] = cell
              }
            })
            return obj
          })
  
          resolve(result)
        } catch (error) {
          reject(error)
        }
      }
  
      reader.onerror = () => {
        reject(new Error("Failed to read file"))
      }
  
      reader.readAsArrayBuffer(file)
    })
  }
  
  function sortCandidateData(data) {
    return [...data].sort((a, b) => {
      // Sort by City (A to Z)
      if ((a.City || "").toLowerCase() < (b.City || "").toLowerCase()) return -1
      if ((a.City || "").toLowerCase() > (b.City || "").toLowerCase()) return 1
  
      // Then by Venue Code (A to Z)
      if ((a["Venue Code"] || "").toLowerCase() < (b["Venue Code"] || "").toLowerCase()) return -1
      if ((a["Venue Code"] || "").toLowerCase() > (b["Venue Code"] || "").toLowerCase()) return 1
  
      // Then by Exam Date (earliest to latest)
      if (a["Exam Date"] < b["Exam Date"]) return -1
      if (a["Exam Date"] > b["Exam Date"]) return 1
  
      // Then by Batch (A to Z)
      if ((a.Batch || "").toLowerCase() < (b.Batch || "").toLowerCase()) return -1
      if ((a.Batch || "").toLowerCase() > (b.Batch || "").toLowerCase()) return 1
  
      // // Then by False No (numerically)
      // const falseNoA = Number.parseInt(a["False No"] || 0)
      // const falseNoB = Number.parseInt(b["False No"] || 0)
      // if (!isNaN(falseNoA) && !isNaN(falseNoB)) {
      //   return falseNoA - falseNoB
      // }
  
      // Then by PWD (A to Z)
      if ((a.PWD || "").toLowerCase() < (b.PWD || "").toLowerCase()) return -1
      if ((a.PWD || "").toLowerCase() > (b.PWD || "").toLowerCase()) return 1
  
      // Then by Candidate ID (A to Z)
      if ((a["Candidate Id"] || "").toLowerCase() < (b["Candidate Id"] || "").toLowerCase()) return -1
      if ((a["Candidate Id"] || "").toLowerCase() > (b["Candidate Id"] || "").toLowerCase()) return 1
  
      return 0
    })
  }
  
  function sortLabConfigData(data) {
    return [...data].sort((a, b) => {
      // Sort by City (A to Z)
      if ((a.City || "").toLowerCase() < (b.City || "").toLowerCase()) return -1
      if ((a.City || "").toLowerCase() > (b.City || "").toLowerCase()) return 1
  
      // Then by Centre Code (A to Z)
      if ((a["Centre Code"] || "").toLowerCase() < (b["Centre Code"] || "").toLowerCase()) return -1
      if ((a["Centre Code"] || "").toLowerCase() > (b["Centre Code"] || "").toLowerCase()) return 1
  
      // Then by Lab No (Smallest to largest)
      const labA = Number.parseInt(a["Lab No"] || 0)
      const labB = Number.parseInt(b["Lab No"] || 0)
      return labA - labB
    })
  }
  
  // function allocateLabs(candidateData, labConfigData) {
  //     // Create a deep copy of lab config data to avoid modifying the original
  //     const labConfigCopy = JSON.parse(JSON.stringify(labConfigData));
  
  //     // Group candidates by city, venue code, date and batch
  //     const candidateGroups = {};
  //     candidateData.forEach(candidate => {
  //         const key = `${candidate.City}-${candidate['Venue Code']}-${candidate['Exam Date']}-${candidate.Batch}`;
  //         if (!candidateGroups[key]) {
  //             candidateGroups[key] = [];
  //         }
  //         candidateGroups[key].push(candidate);
  //     });
  
  //     // DEBUG: Log all venue codes from candidates
  //     console.log('Candidate Venue Codes:',
  //         [...new Set(candidateData.map(c => `${c.City}-${c['Venue Code']}`))]);
  
  //     // DEBUG: Log all centre codes from lab config
  //     console.log('Lab Centre Codes:',
  //         [...new Set(labConfigData.map(l => `${l.City}-${l['Centre Code']}`))]);
  
  //     // Group labs by city and center code
  //     const labGroups = {};
  //     labConfigCopy.forEach(lab => {
  //         // Make sure to handle case sensitivity and trim
  //         const cityKey = (lab.City || '').trim();
  //         const centreKey = (lab['Centre Code'] || '').trim();
  //         const key = `${cityKey}-${centreKey}`;
  
  //         if (!labGroups[key]) {
  //             labGroups[key] = [];
  //         }
  //         labGroups[key].push({
  //             ...lab,
  //             originalCount: parseInt(lab.Count || 0),
  //             availableSeats: parseInt(lab.Count || 0)
  //         });
  //     });
  
  //     // Allocate candidates to labs
  //     const allocatedCandidates = [];
  
  //     // Process each group of candidates
  //     Object.keys(candidateGroups).forEach(key => {
  //         const [city, venueCode, examDate, batch] = key.split('-');
  //         const candidates = candidateGroups[key];
  
  //         console.log(`Processing: ${city} - ${venueCode} - ${examDate} - ${batch}, Candidates: ${candidates.length}`);
  
  //         // Find matching lab group for this city and venue code
  //         // First try direct match
  //         let labKey = Object.keys(labGroups).find(k => {
  //             const [labCity, labCentreCode] = k.split('-');
  //             return labCity === city && labCentreCode === venueCode;
  //         });
  
  //         // If no direct match, try case-insensitive match
  //         if (!labKey) {
  //             labKey = Object.keys(labGroups).find(k => {
  //                 const [labCity, labCentreCode] = k.split('-');
  //                 return labCity.toLowerCase() === city.toLowerCase() &&
  //                        labCentreCode.toLowerCase() === venueCode.toLowerCase();
  //             });
  //         }
  
  //         // If still no match, try matching just by Centre Code/Venue Code
  //         if (!labKey) {
  //             labKey = Object.keys(labGroups).find(k => {
  //                 const [, labCentreCode] = k.split('-');
  //                 return labCentreCode === venueCode;
  //             });
  //         }
  
  //         console.log(`For ${city}-${venueCode}, found lab key: ${labKey}`);
  
  //         if (!labKey || !labGroups[labKey] || labGroups[labKey].length === 0) {
  //             console.error(`Lab groups available:`, Object.keys(labGroups));
  //             throw new Error(`No labs available for candidates in ${city} at venue ${venueCode} for batch ${batch} on ${examDate}`);
  //         }
  
  //         const labs = labGroups[labKey];
  
  //         // Reset lab capacities for each new batch at this venue on this date
  //         // We'll use a more reliable key that includes city, venue, date and batch
  //         const dateAndBatchKey = `${city}-${venueCode}-${examDate}-${batch}`;
  
  //         // Reset lab capacities for this batch
  //         labs.forEach(lab => {
  //             lab.availableSeats = lab.originalCount;
  //         });
  
  //         console.log(`Reset capacities for ${dateAndBatchKey}, available:`,
  //             labs.map(l => `Lab ${l['Lab No']}: ${l.availableSeats}`).join(', '));
  
  //         // Calculate total available seats across all labs for this batch
  //         const totalAvailableSeats = labs.reduce((total, lab) => total + lab.availableSeats, 0);
  
  //         // Check if we have enough seats for all candidates in this batch
  //         if (totalAvailableSeats < candidates.length) {
  //             throw new Error(`Not enough lab capacity for all candidates in ${city} at venue ${venueCode} for batch ${batch} on ${examDate}. Need ${candidates.length} seats but only ${totalAvailableSeats} available.`);
  //         }
  
  //         // Sort labs by Lab No
  //         labs.sort((a, b) => parseInt(a['Lab No']) - parseInt(b['Lab No']));
  
  //         // Allocate candidates sequentially across labs
  //         let currentLabIndex = 0;
  
  //         candidates.forEach((candidate, index) => {
  //             // Find next lab with available seats
  //             while (currentLabIndex < labs.length && labs[currentLabIndex].availableSeats <= 0) {
  //                 currentLabIndex++;
  //             }
  
  //             if (currentLabIndex >= labs.length) {
  //                 throw new Error(`Not enough lab capacity for all candidates in ${city} at venue ${venueCode} for batch ${batch} on ${examDate}`);
  //             }
  
  //             const currentLab = labs[currentLabIndex];
  
  //             // Calculate seat number within the lab
  //             // Start with seat 1 for each lab
  //             const labSeatIndex = currentLab.originalCount - currentLab.availableSeats + 1;
  
  //             // Generate final seat number
  //             let seatNo;
  //             if (candidate['False No']) {
  //                 // Use false number as is or combine with lab seat index
  //                 seatNo = candidate['False No'];
  //             } else {
  //                 // Just use the sequential number
  //                 seatNo = labSeatIndex;
  //             }
  
  //             // Allocate candidate to this lab
  //             allocatedCandidates.push({
  //                 ...candidate,
  //                 'Building Name': currentLab['Building Name'],
  //                 'Floor Name': currentLab['Floor Name'],
  //                 'Lab Name': currentLab['Lab Name'],
  //                 'Lab No': currentLab['Lab No'],
  //                 'Server': currentLab['Server'],
  //                 'Seat No': seatNo
  //             });
  
  //             // Decrease available seats
  //             currentLab.availableSeats--;
  //         });
  
  //         console.log(`Completed allocation for ${dateAndBatchKey}, remaining:`,
  //             labs.map(l => `Lab ${l['Lab No']}: ${l.availableSeats}`).join(', '));
  //     });
  
  //     return allocatedCandidates;
  // }
  function allocateLabs(candidateData, labConfigData) {
    // Create a deep copy of lab config data to avoid modifying the original
    const labConfigCopy = JSON.parse(JSON.stringify(labConfigData))
  
    // Group candidates by city, venue code, date and batch
    const candidateGroups = {}
    candidateData.forEach((candidate) => {
      const key = `${candidate.City}-${candidate["Venue Code"]}-${candidate["Exam Date"]}-${candidate.Batch}`
      if (!candidateGroups[key]) {
        candidateGroups[key] = []
      }
      candidateGroups[key].push(candidate)
    })
  
    // DEBUG: Log all venue codes from candidates
    console.log("Candidate Venue Codes:", [...new Set(candidateData.map((c) => `${c.City}-${c["Venue Code"]}`))])
  
    // DEBUG: Log all centre codes from lab config
    console.log("Lab Centre Codes:", [...new Set(labConfigData.map((l) => `${l.City}-${l["Centre Code"]}`))])
  
    // Group labs by city and center code
    const labGroups = {}
    labConfigCopy.forEach((lab) => {
      // Make sure to handle case sensitivity and trim
      const cityKey = (lab.City || "").trim()
      const centreKey = (lab["Centre Code"] || "").trim()
      const key = `${cityKey}-${centreKey}`
  
      if (!labGroups[key]) {
        labGroups[key] = []
      }
      labGroups[key].push({
        ...lab,
        originalCount: Number.parseInt(lab.Count || 0),
        availableSeats: Number.parseInt(lab.Count || 0),
      })
    })
  
    // Allocate candidates to labs
    const allocatedCandidates = []
  
    // Process each group of candidates
    Object.keys(candidateGroups).forEach((key) => {
      const [city, venueCode, examDate, batch] = key.split("-")
      const candidates = candidateGroups[key]
  
      console.log(`Processing: ${city} - ${venueCode} - ${examDate} - ${batch}, Candidates: ${candidates.length}`)
  
      // Find matching lab group for this city and venue code
      // First try direct match
      let labKey = Object.keys(labGroups).find((k) => {
        const [labCity, labCentreCode] = k.split("-")
        return labCity === city && labCentreCode === venueCode
      })
  
      // If no direct match, try case-insensitive match
      if (!labKey) {
        labKey = Object.keys(labGroups).find((k) => {
          const [labCity, labCentreCode] = k.split("-")
          return labCity.toLowerCase() === city.toLowerCase() && labCentreCode.toLowerCase() === venueCode.toLowerCase()
        })
      }
  
      // If still no match, try matching just by Centre Code/Venue Code
      if (!labKey) {
        labKey = Object.keys(labGroups).find((k) => {
          const [, labCentreCode] = k.split("-")
          return labCentreCode === venueCode
        })
      }
  
      console.log(`For ${city}-${venueCode}, found lab key: ${labKey}`)
  
      if (!labKey || !labGroups[labKey] || labGroups[labKey].length === 0) {
        console.error(`Lab groups available:`, Object.keys(labGroups))
        throw new Error(
          `No labs available for candidates in ${city} at venue ${venueCode} for batch ${batch} on ${examDate}`,
        )
      }
  
      const labs = labGroups[labKey]
  
      // Reset lab capacities for each new batch at this venue on this date
      labs.forEach((lab) => {
        lab.availableSeats = lab.originalCount
      })
  
      console.log(
        `Reset capacities for ${city}-${venueCode}-${examDate}-${batch}, available:`,
        labs.map((l) => `Lab ${l["Lab No"]}: ${l.availableSeats}`).join(", "),
      )
  
      // Calculate total available seats across all labs for this batch
      const totalAvailableSeats = labs.reduce((total, lab) => total + lab.availableSeats, 0)
  
      // Check if we have enough seats for all candidates in this batch
      if (totalAvailableSeats < candidates.length) {
        throw new Error(
          `Not enough lab capacity for all candidates in ${city} at venue ${venueCode} for batch ${batch} on ${examDate}. Need ${candidates.length} seats but only ${totalAvailableSeats} available.`,
        )
      }
  
      // Sort labs by Lab No
      labs.sort((a, b) => Number.parseInt(a["Lab No"]) - Number.parseInt(b["Lab No"]))
  
      // Allocate candidates sequentially across labs
      let currentLabIndex = 0
      let seatCounterGlobal = 1 // Global counter across all labs
  
      // Lab-specific seat counters
      const labSeatCounters = {}
      labs.forEach((lab) => {
        labSeatCounters[lab["Lab No"]] = 1
      })
  
      candidates.forEach((candidate, index) => {
        // Find next lab with available seats
        while (currentLabIndex < labs.length && labs[currentLabIndex].availableSeats <= 0) {
          currentLabIndex++
        }
  
        if (currentLabIndex >= labs.length) {
          throw new Error(
            `Not enough lab capacity for all candidates in ${city} at venue ${venueCode} for batch ${batch} on ${examDate}`,
          )
        }
  
        const currentLab = labs[currentLabIndex]
        const labNumber = currentLab["Lab No"]
  
        // Get the current seat counter for this lab
        const labSeatCounter = labSeatCounters[labNumber]
  
        // Generate seat number in format falsenumber_sequence
        let seatNo
        if (candidate["False No"]) {
          seatNo = `${candidate["False No"]}_${labSeatCounter}`
        } else {
          seatNo = labSeatCounter.toString()
        }
  
        // Allocate candidate to this lab
        allocatedCandidates.push({
          ...candidate,
          "Building Name": currentLab["Building Name"],
          "Floor Name": currentLab["Floor Name"],
          "Lab Name": currentLab["Lab Name"],
          "Lab No": currentLab["Lab No"],
          Server: currentLab["Server"],
          "Seat No": seatNo,
        })
  
        // Decrease available seats and increment counters
        currentLab.availableSeats--
        labSeatCounters[labNumber]++
        seatCounterGlobal++
      })
  
      console.log(
        `Completed allocation for ${city}-${venueCode}-${examDate}-${batch}, remaining:`,
        labs.map((l) => `Lab ${l["Lab No"]}: ${l.availableSeats}`).join(", "),
      )
    })
  
    return allocatedCandidates
  }
 
  // Also update the seat allocation function to maintain the format
  // Replace the current allocateSeats function with this updated version:
  // function allocateLabs(candidateData, labConfigData) {
  //   // Create a deep copy of lab config data to avoid modifying the original
  //   const labConfigCopy = JSON.parse(JSON.stringify(labConfigData))
  
  //   // Group candidates by city, venue code, date and batch
  //   const candidateGroups = {}
  //   candidateData.forEach((candidate) => {
  //     const key = `${candidate.City}-${candidate["Venue Code"]}-${candidate["Exam Date"]}-${candidate.Batch}`
  //     if (!candidateGroups[key]) {
  //       candidateGroups[key] = []
  //     }
  //     candidateGroups[key].push(candidate)
  //   })
  
  //   // Group labs by city and center code
  //   const labGroups = {}
  //   labConfigCopy.forEach((lab) => {
  //     const cityKey = (lab.City || "").trim()
  //     const centreKey = (lab["Centre Code"] || "").trim()
  //     const key = `${cityKey}-${centreKey}`
  
  //     if (!labGroups[key]) {
  //       labGroups[key] = []
  //     }
  //     labGroups[key].push({
  //       ...lab,
  //       originalCount: Number.parseInt(lab.Count || 0),
  //       availableSeats: Number.parseInt(lab.Count || 0),
  //     })
  //   })
  
  //   // Allocate candidates to labs
  //   const allocatedCandidates = []
  
  //   // Process each group of candidates
  //   Object.keys(candidateGroups).forEach((key) => {
  //     const [city, venueCode, examDate, batch] = key.split("-")
  //     const candidates = candidateGroups[key]
  
  //     console.log(`Processing: ${city} - ${venueCode} - ${examDate} - ${batch}, Candidates: ${candidates.length}`)
  
  //     // Find matching lab group
  //     let labKey = Object.keys(labGroups).find((k) => {
  //       const [labCity, labCentreCode] = k.split("-")
  //       return labCity === city && labCentreCode === venueCode
  //     })
  
  //     if (!labKey) {
  //       labKey = Object.keys(labGroups).find((k) => {
  //         const [labCity, labCentreCode] = k.split("-")
  //         return labCity.toLowerCase() === city.toLowerCase() && 
  //                labCentreCode.toLowerCase() === venueCode.toLowerCase()
  //       })
  //     }
  
  //     if (!labKey) {
  //       labKey = Object.keys(labGroups).find((k) => {
  //         const [, labCentreCode] = k.split("-")
  //         return labCentreCode === venueCode
  //       })
  //     }
  
  //     if (!labKey || !labGroups[labKey] || labGroups[labKey].length === 0) {
  //       console.error(`Lab groups available:`, Object.keys(labGroups))
  //       throw new Error(
  //         `No labs available for candidates in ${city} at venue ${venueCode} for batch ${batch} on ${examDate}`
  //       )
  //     }
  
  //     const labs = labGroups[labKey]
  
  //     // Reset lab capacities for each new batch
  //     labs.forEach((lab) => {
  //       lab.availableSeats = lab.originalCount
  //     })
  
  //     // Calculate total available seats
  //     const totalAvailableSeats = labs.reduce((total, lab) => total + lab.availableSeats, 0)
  
  //     // Check if we have enough seats
  //     if (totalAvailableSeats < candidates.length) {
  //       throw new Error(
  //         `Not enough lab capacity for all candidates in ${city} at venue ${venueCode} for batch ${batch} on ${examDate}. Need ${candidates.length} seats but only ${totalAvailableSeats} available.`
  //       )
  //     }
  
  //     // Sort labs by Lab No
  //     labs.sort((a, b) => Number.parseInt(a["Lab No"]) - Number.parseInt(b["Lab No"]))
  
  //     // Allocate candidates sequentially across labs
  //     let currentLabIndex = 0
      
  //     // Use a single counter per batch that continues across all labs
  //     let batchSeatCounter = 1
      
  //     candidates.forEach((candidate) => {
  //       // Find next lab with available seats
  //       while (currentLabIndex < labs.length && labs[currentLabIndex].availableSeats <= 0) {
  //         currentLabIndex++
  //       }
  
  //       if (currentLabIndex >= labs.length) {
  //         throw new Error(
  //           `Not enough lab capacity for all candidates in ${city} at venue ${venueCode} for batch ${batch} on ${examDate}`
  //         )
  //       }
  
  //       const currentLab = labs[currentLabIndex]
  
  //       // Generate seat number using the batch-level counter
  //       let seatNo
  //       if (candidate["False No"]) {
  //         seatNo = `${candidate["False No"]}_${batchSeatCounter}`
  //       } else {
  //         seatNo = batchSeatCounter.toString()
  //       }
  
  //       // Allocate candidate to this lab
  //       allocatedCandidates.push({
  //         ...candidate,
  //         "Building Name": currentLab["Building Name"],
  //         "Floor Name": currentLab["Floor Name"],
  //         "Lab Name": currentLab["Lab Name"],
  //         "Lab No": currentLab["Lab No"],
  //         "Server": currentLab["Server"],
  //         "Seat No": seatNo,
  //       })
  
  //       // Decrease available seats and increment batch counter
  //       currentLab.availableSeats--
  //       batchSeatCounter++ // This counter is never reset within a batch
  //     })
  //   })
  
  //   return allocatedCandidates
  // }
  function displayLabAllocationResults(data) {
    if (!data || data.length === 0) {
      labAllocationResults.innerHTML = "<p>No results to display.</p>"
      return
    }
  
    const headers = [
      "Candidate Id",
      "Candidate Email",
      "Venue Code",
      "Venue Name",
      "City",
      "Exam Date",
      "Exam Day",
      "Batch",
      "False No",
      "PWD",
      "Building Name",
      "Floor Name",
      "Lab Name",
      "Lab No",
      "Server",
      "Seat No",
    ]
  
    let tableHtml = "<table><thead><tr>"
    headers.forEach((header) => {
      tableHtml += `<th>${header}</th>`
    })
    tableHtml += "</tr></thead><tbody>"
  
    data.forEach((row) => {
      tableHtml += "<tr>"
      headers.forEach((header) => {
        tableHtml += `<td>${row[header] !== undefined ? row[header] : ""}</td>`
      })
      tableHtml += "</tr>"
    })
  
    tableHtml += "</tbody></table>"
    labAllocationResults.innerHTML = tableHtml
  }
  
  function displaySeatAllocationResults(data) {
    if (!data || data.length === 0) {
      seatAllocationResults.innerHTML = "<p>No results to display.</p>"
      return
    }
  
    const headers = [
      "Candidate Id",
      "Candidate Email",
      "Venue Code",
      "Venue Name",
      "City",
      "Exam Date",
      "Exam Day",
      "Batch",
      "False No",
      "PWD",
      "Building Name",
      "Floor Name",
      "Lab Name",
      "Lab No",
      "Server",
      "Seat No",
    ]
  
    let tableHtml = "<table><thead><tr>"
    headers.forEach((header) => {
      tableHtml += `<th>${header}</th>`
    })
    tableHtml += "</tr></thead><tbody>"
  
    data.forEach((row) => {
      tableHtml += "<tr>"
      headers.forEach((header) => {
        tableHtml += `<td>${row[header] !== undefined ? row[header] : ""}</td>`
      })
      tableHtml += "</tr>"
    })
  
    tableHtml += "</tbody></table>"
    seatAllocationResults.innerHTML = tableHtml
  }
  
  function downloadExcelFile(data, filename) {
    const worksheet = XLSX.utils.json_to_sheet(data)
    const workbook = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1")
    XLSX.writeFile(workbook, filename)
  }
  
  