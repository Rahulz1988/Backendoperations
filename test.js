// Modify the download seat allocation button event handler
downloadSeatAllocationBtn.addEventListener("click", () => {
    // Download the original processed data
    downloadExcelFile(appState.seatAllocationResult, "Final_Seat_Allocation.xlsx")
    
    // Generate and download the consolidated Playcard details
    const playcardData = generatePlaycardDetails(appState.seatAllocationResult)
    downloadExcelFile(playcardData, "Playcard_Details.xlsx")
  })
  
  // Function to generate Playcard details
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