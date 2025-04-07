function createOrUpdateEventMultipliers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1) Define which sheets to exclude from being treated as events
  const excludeSheets = ["Member Overview", "Home", "EventMultipliers", "Dashboard", "Leaderboard"];

  // 2) Get all sheets and filter out the excluded ones
  const allSheets = ss.getSheets();
  const eventSheetNames = allSheets
    .map(sheet => sheet.getName())
    .filter(name => !excludeSheets.includes(name));

  // 3) Create or clear the "EventMultipliers" sheet
  let multipliersSheet = ss.getSheetByName("EventMultipliers");
  if (multipliersSheet) {
    multipliersSheet.clear();
  } else {
    multipliersSheet = ss.insertSheet("EventMultipliers");
  }

  // 4) Set up a header row
  const headers = ["Event Numbers", "Name", "Attendees", "Points"];
  multipliersSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // 5) For each event sheet, count unique IDs
  const data = [];
  for (let i = 0; i < eventSheetNames.length; i++) {
    const sheetName = eventSheetNames[i];
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) continue; // just in case

    // Read all rows
    const values = sheet.getDataRange().getValues();
    // Create a Set for unique IDs
    const uniqueIDs = new Set();

    // Assuming the ID is in column E (index 4)
    // and row 1 is a header, so start at row index 1
    for (let rowIndex = 1; rowIndex < values.length; rowIndex++) {
      const row = values[rowIndex];
      const id = row[4]; // column E
      // Only add if it's not empty
      if (id) {
        uniqueIDs.add(String(id).trim());
      }
    }

    // The number of unique IDs is the number of unique attendees
    const attendees = uniqueIDs.size;

    // Build one row for the EventMultipliers sheet
    data.push([
      i + 1,        // Event number
      sheetName,    // The name of the event sheet
      attendees,    // Unique attendees
      ""            // Leave points blank for manual entry
    ]);
  }

  // 6) Write the data to EventMultipliers, starting in row 2
  if (data.length > 0) {
    multipliersSheet
      .getRange(2, 1, data.length, headers.length)
      .setValues(data);
  }
}


