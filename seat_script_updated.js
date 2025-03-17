// Track application state
const appState = {
    candidateData: null,
    labConfigData: null,
    finalAllocationResult: null
  }
  
  // DOM Elements
  const candidateFileInput = document.getElementById("candidateFile")
  const labConfigFileInput = document.getElementById("labConfigFile")
  const allocateSeatsBtn = document.getElementById("allocateSeatsBtn")
  const downloadFinalAllocationBtn = document.getElementById("downloadFinalAllocationBtn")
  const allocationStatus = document.getElementById("allocationStatus")
  const allocationResults = document.getElementById("allocationResults")
  
  // Event Listeners for file inputs
  candidateFileInput.addEventListener("change", async (e) => {
    const file = e.target.files[0]
    if (!file) return
  
    try {
      showNotification(allocationStatus, "Reading candidate data...", "info")
      appState.candidateData = await readExcelFile(file)
      showNotification(allocationStatus, "Candidate data loaded successfully!", "success")
      checkEnableAllocation()
    } catch (error) {
      console.error("Error reading candidate data:", error)
      showNotification(allocationStatus, "Error reading candidate data: " + error.message, "error")
    }
  })
  
  labConfigFileInput.addEventListener("change", async (e) => {
    const file = e.target.files[0]
    if (!file) return
  
    try {
      showNotification(allocationStatus, "Reading lab configuration data...", "info")
      appState.labConfigData = await readExcelFile(file)
      showNotification(allocationStatus, "Lab configuration data loaded successfully!", "success")
      checkEnableAllocation()
    } catch (error) {
      console.error("Error reading lab configuration data:", error)
      showNotification(allocationStatus, "Error reading lab configuration data: " + error.message, "error")
    }
  })
  
  // Combined Allocation Button Click Handler
  allocateSeatsBtn.addEventListener("click", () => {
    try {
      showNotification(allocationStatus, "Allocating labs and seats...", "info")
  
      // Step 1: Sort candidate data
      const sortedCandidates = sortCandidateData(appState.candidateData)
  
      // Step 2: Sort lab configuration data
      const sortedLabConfig = sortLabConfigData(appState.labConfigData)
  
      // Step 3: Perform lab allocation
      const labAllocationResult = allocateLabs(sortedCandidates, sortedLabConfig)
  
      // Step 4: Perform seat allocation
      appState.finalAllocationResult = allocateSeats(labAllocationResult)
  
      // Display final results
      //displayAllocationResults(appState.finalAllocationResult)
  
      showNotification(allocationStatus, "Lab and seat allocation completed successfully!", "success")
      downloadFinalAllocationBtn.disabled = false
    } catch (error) {
      console.error("Error during allocation:", error)
      showNotification(allocationStatus, "Error during allocation: " + error.message, "error")
    }
  })
  
  
  // Download Button Click Handler
downloadFinalAllocationBtn.addEventListener("click", () => {
    try {
      // Download the original processed data
      downloadExcelFile(appState.finalAllocationResult, "file_statistics-Final_Allocation.xlsx")
      
      // Generate and download the consolidated Playcard details
      const playcardData = generatePlaycardDetails(appState.finalAllocationResult)
      downloadExcelFile(playcardData, "Placard_Details.xlsx")
      
      // Disable the download button after successful download
      downloadFinalAllocationBtn.disabled = true;
      
      // Show a success message
      showNotification(allocationStatus, "Files downloaded successfully! Upload new files to process another batch.", "success")
    } catch (error) {
      console.error("Error during file download:", error)
      showNotification(allocationStatus, "Error downloading files: " + error.message, "error")
    }
  })
  // Helper Functions
  function checkEnableAllocation() {
    allocateSeatsBtn.disabled = !(appState.candidateData && appState.labConfigData)
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
          const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false })
  
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
  
        // Allocate candidate to this lab
        // Allocate candidate to this lab
      allocatedCandidates.push({
        ...candidate,
        "Building Name": currentLab["Building Name"],
        "Floor Name": currentLab["Floor Name"],
        "Lab Name": currentLab["Lab Name"],
        "Lab No": currentLab["Lab No"],
        "Server": currentLab["Server"],
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

function generatePlaycardDetails(data) {
  if (!data || data.length === 0) return []
  
  // Group candidates by Center Code, Center Name, City, Batch, Building, Floor, Lab, Server
  const groupedData = {}
  
  data.forEach(candidate => {
    // Create a more specific key to ensure proper grouping by batch and lab
    const key = `${candidate["Venue Code"]}|${candidate["Venue Name"]}|${candidate.City}|${candidate.Batch}|${candidate["Building Name"]}|${candidate["Floor Name"]}|${candidate["Lab Name"]}|${candidate["Lab No"]}|${candidate["Server"]}`
    
    if (!groupedData[key]) {
      groupedData[key] = {
        candidates: [],
        "Center Code": candidate["Venue Code"],
        "Center Name": candidate["Venue Name"],
        "City": candidate.City,
        "Batch": candidate.Batch,
        "Building Name": candidate["Building Name"],
        "Floor Name": candidate["Floor Name"], 
        "Lab Name": candidate["Lab Name"],
        "Lab No": candidate["Lab No"],
        "Server": candidate["Server"]
      }
    }
    
    groupedData[key].candidates.push(candidate)
  })
  
  // Generate the consolidated Playcard details
  const playcardDetails = []
  
  Object.values(groupedData).forEach(group => {
    // First, sort candidates by Candidate Id (numerically) to find starting and ending roll numbers
    group.candidates.sort((a, b) => {
      // Convert Candidate Id to numbers for proper numeric sorting
      const candidateIdA = parseInt(a["Candidate Id"].replace(/\D/g, ''), 10) || 0
      const candidateIdB = parseInt(b["Candidate Id"].replace(/\D/g, ''), 10) || 0
      return candidateIdA - candidateIdB
    })
    
    // Get the starting and ending roll numbers
    const startingRollNo = group.candidates[0]["Candidate Id"] || "N/A"
    const endingRollNo = group.candidates[group.candidates.length - 1]["Candidate Id"] || "N/A"
    
    // Now sort by seat number to determine starting and ending seat numbers
    group.candidates.sort((a, b) => {
      const seatA = parseInt(a["Seat No"], 10) || 0
      const seatB = parseInt(b["Seat No"], 10) || 0
      return seatA - seatB
    })
    
    // Get the starting and ending seat values
    const startingSeatNo = group.candidates[0]["Seat No"]
    const endingSeatNo = group.candidates[group.candidates.length - 1]["Seat No"]
    
    // Create the Playcard record
    playcardDetails.push({
      "Center Code": group["Center Code"],
      "Center Name": group["Center Name"],
      "City": group["City"],
      "Batch": group["Batch"],
      "Building Name": group["Building Name"],
      "Floor Name": group["Floor Name"],
      "Lab Name": group["Lab Name"],
      "Lab No": group["Lab No"],
      "Server": group["Server"],
      "Starting Seat No": startingSeatNo,
      "Ending Seat No": endingSeatNo,
      "Starting Roll No.": startingRollNo,
      "Ending Roll No.": endingRollNo,
      "Alloted Count": group.candidates.length
    })
  })
  
  return playcardDetails
}

function allocateSeats(labAllocatedData) {
  // Create a deep copy to avoid modifying the original data
  const data = JSON.parse(JSON.stringify(labAllocatedData))
  
  // First sort by all the requested criteria
  data.sort((a, b) => {
    // Sort by City (A to Z)
    if ((a.City || "").toLowerCase() < (b.City || "").toLowerCase()) return -1
    if ((a.City || "").toLowerCase() > (b.City || "").toLowerCase()) return 1
    
    // Then by Venue Code (A to Z)
    if ((a["Venue Code"] || "").toLowerCase() < (b["Venue Code"] || "").toLowerCase()) return -1
    if ((a["Venue Code"] || "").toLowerCase() > (b["Venue Code"] || "").toLowerCase()) return 1
    
    // Then by Batch (A to Z)
    if ((a.Batch || "").toLowerCase() < (b.Batch || "").toLowerCase()) return -1
    if ((a.Batch || "").toLowerCase() > (b.Batch || "").toLowerCase()) return 1
    
    // Then by Lab No (Smallest to largest)
    const labA = parseInt(a["Lab No"] || 0, 10)
    const labB = parseInt(b["Lab No"] || 0, 10)
    if (labA !== labB) return labA - labB
    
    // Then by PWD (A to Z)
    if ((a.PWD || "").toLowerCase() < (b.PWD || "").toLowerCase()) return -1
    if ((a.PWD || "").toLowerCase() > (b.PWD || "").toLowerCase()) return 1
    
    // Then by False No (Smallest to largest)
    const falseNoA = parseInt(a["False No"] || 0, 10)
    const falseNoB = parseInt(b["False No"] || 0, 10)
    return falseNoA - falseNoB
  })
  
  // Initialize a new structure to track seat counter by venue+batch combination
  const counters = {}
  
  // Process each candidate in the sorted order
  data.forEach(candidate => {
    // Create a unique key for each city+venue+batch combination
    const batchKey = `${candidate.City}-${candidate["Venue Code"]}-${candidate["Exam Date"]}-${candidate.Batch}`
    
    // Initialize counter for this batch if it doesn't exist
    if (!counters[batchKey]) {
      counters[batchKey] = 1
    }
    
    // Get the current counter value for this batch
    const currentCounter = counters[batchKey]
    
    // Assign the seat number with the appropriate format
    if (candidate["False No"]) {
      candidate["Seat No"] = `${currentCounter}`
    } else {
      candidate["Seat No"] = currentCounter.toString()
    }
    
    // Increment the counter for the next candidate in this batch
    counters[batchKey]++
  })
  
  return data
}

// function displayAllocationResults(data) {
//   if (!data || data.length === 0) {
//     allocationResults.innerHTML = "<p>No results to display.</p>"
//     return
//   }

//   // Headers for the results table
//   const headers = [
//     "Candidate Id",
//     "Candidate Email",
//     "Venue Code",
//     "Venue Name",
//     "City",
//     "Exam Date",
//     "Exam Day",
//     "Batch",
//     "False No",
//     "PWD",
//     "Building Name",
//     "Floor Name",
//     "Lab Name",
//     "Lab No",
//     "Server",
//     "Seat No",
//   ]

//   let tableHtml = "<table><thead><tr>"
//   headers.forEach((header) => {
//     tableHtml += `<th>${header}</th>`
//   })
//   tableHtml += "</tr></thead><tbody>"

//   data.forEach((row) => {
//     tableHtml += "<tr>"
//     headers.forEach((header) => {
//       tableHtml += `<td>${row[header] !== undefined ? row[header] : ""}</td>`
//     })
//     tableHtml += "</tr>"
//   })

//   tableHtml += "</tbody></table>"
//   allocationResults.innerHTML = tableHtml
// }

function downloadExcelFile(data, filename) {
  const worksheet = XLSX.utils.json_to_sheet(data)
  const workbook = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1")
  XLSX.writeFile(workbook, filename)
}