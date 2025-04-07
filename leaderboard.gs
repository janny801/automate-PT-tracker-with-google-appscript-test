function createLeaderboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const memberSheet = ss.getSheetByName("Member Overview");
  if (!memberSheet) {
    Logger.log("Member Overview sheet not found.");
    return;
  }
  
  // Get all data from the Member Overview sheet.
  const data = memberSheet.getDataRange().getValues();
  if (data.length < 2) {
    Logger.log("Not enough data to create leaderboard.");
    return;
  }
  
  // Get the header row and find the indexes for "Name" and "Total Points".
  const headers = data[0];
  let nameIndex = headers.indexOf("Name");
  let pointsIndex = headers.indexOf("Total Points");
  
  // Fallback: if "Name" isn't found, try combining "First Name" and "Last Name" (assumed to be columns 0 and 1).
  const useFullName = nameIndex === -1;
  if (pointsIndex === -1) {
    // If Total Points isn't found, default to column E (index 4)
    pointsIndex = 4;
  }
  
  // Build an array of [Name, Points] for each member.
  const leaderboardData = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    let name = "";
    if (useFullName) {
      // Assuming First Name is in column A and Last Name is in column B.
      name = row[0] + " " + row[1];
    } else {
      name = row[nameIndex];
    }
    const points = parseFloat(row[pointsIndex]) || 0;
    leaderboardData.push([name, points]);
  }
  
  // Sort the data in descending order by points.
  leaderboardData.sort((a, b) => b[1] - a[1]);
  
  // Create (or clear) the "Leaderboard" sheet.
  let leaderboardSheet = ss.getSheetByName("Leaderboard");
  if (leaderboardSheet) {
    leaderboardSheet.clear();
  } else {
    leaderboardSheet = ss.insertSheet("Leaderboard");
  }
  
  // Set the header row: Rank, Name, Points.
  leaderboardSheet.getRange(1, 1, 1, 3).setValues([["Rank", "Name", "Points"]]);
  
  // Build the output array with ranking.
  const output = [];
  for (let i = 0; i < leaderboardData.length; i++) {
    output.push([i + 1, leaderboardData[i][0], leaderboardData[i][1]]);
  }
  
  // Write the sorted data starting from row 2.
  leaderboardSheet.getRange(2, 1, output.length, 3).setValues(output);
  
  // Optionally auto-resize columns for a neat appearance.
  leaderboardSheet.autoResizeColumns(1, 3);
}

