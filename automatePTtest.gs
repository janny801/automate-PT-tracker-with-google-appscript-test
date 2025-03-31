// Global variables
var exclusions = [
  "Home",
  "Member Overview",
  "Leaderboard",
  "Bonus"
];
var officers = [
  "Ny Dang", "Eric Lai", "Melody Nguyen", "Nathan Turan", 
  "Han Duong", "Richard Luong", "Randy Le", "Travis Nguyen", 
  "Ethan Tran", "Carter Ung", "Brian Nguyen", "Kelsey Wong"
];

/**
 * Returns a unidirectional array of non-empty values for a given header from a sheet.
 * @param {string} sheetName - The name of the sheet.
 * @param {string} target - The header to search for.
 * @param {number} offset - The starting row (default is 2).
 * @param {number} size - Number of rows to retrieve (default is 400).
 * @returns {Array} Array of values.
 */
function getRangeByName(sheetName, target, offset = 2, size = 400) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var headerRow = sheet.getDataRange().getValues()[0];
  var colIndex = headerRow.indexOf(target) + 1;
  var values = sheet.getRange(offset, colIndex, size, 1).getValues();
  // Filter out empty cells and return a one-dimensional array
  return values.filter(function(item) { return item[0] !== ""; }).map(function(item) { return item[0]; });
}

/**
 * Retrieves event names from the spreadsheet by reading all sheet names,
 * excluding those defined in the exclusions array. When called without a raw flag,
 * the event names are written to the Home sheet; when raw is true, an array is returned.
 * @param {boolean} raw - If true, returns an array of event names.
 * @returns {Array|undefined} Array of event names if raw is true.
 */
function getEventNames(raw) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var homeSheet = ss.getSheetByName("Home");
  if (!raw) {
    // Clear previous event names from Home sheet
    homeSheet.getRange(17, 2, 50, 2).clearContent();
  }
  
  var sheets = ss.getSheets();
  var eventMap = Object.entries(sheets.map(function(sheet) {
    var currentName = sheet.getName();
    if (!(exclusions.indexOf(currentName) > -1)) return currentName;
  }).filter(function(key) { return key !== undefined; }));
  
  if (raw) {
    var clean_ev = [];
    for (var i = 0; i < eventMap.length; i++) {
      clean_ev.push(eventMap[i][1]);
    }
    return clean_ev;
  }
  
  homeSheet.getRange(17, 2, eventMap.length, 2).setValues(eventMap);
  SpreadsheetApp.flush();
}

/**
 * Creates and updates the Leaderboard sheet.
 * Excludes members whose full names (First Name + Last Name) are listed in the officers array.
 */
function createLeaderboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var layout = ["First Name", "Last Name", "Total Points"];
  var unsorted_rang = [];
  var source_data = [];
  
  // Get columns from the "Member Overview" sheet
  layout.forEach(function(header) {
    source_data.push(getRangeByName("Member Overview", header));
  });
  
  // Build the leaderboard data excluding officers
  for (var i = 0; i < source_data[0].length; i++) {
    var fullName = source_data[0][i] + " " + source_data[1][i];
    if (officers.indexOf(fullName) > -1) continue;
    unsorted_rang.push([source_data[0][i], source_data[1][i], source_data[2][i]]);
  }
  
  var leaderboardSheet = ss.getSheetByName("Leaderboard");
  leaderboardSheet.getRange(2, 1, unsorted_rang.length, layout.length).setValues(unsorted_rang);
  // Sort leaderboard by Total Points (column 3) in descending order
  leaderboardSheet.sort(3, false);
  // Clear and then set special markers in column D
  leaderboardSheet.getRange("D:D").clear();
  leaderboardSheet.getRange("D2:D3").setValues([["👑"],["👑"]]);
}

/**
 * Helper function: Updates the Home tab by calling getEventNames.
 */
function updateHomeTab() {
  getEventNames();
}

/**
 * Helper function: Updates the Leaderboard tab by calling createLeaderboard.
 */
function updateLeaderboardTab() {
  createLeaderboard();
}

/**
 * Processes a form submission, updating member information in the "Member Overview" sheet,
 * and then updating the Home and Leaderboard tabs.
 * @param {Object} e - The event object from the form submission.
 */
function onFormSubmit(e) {
  // The specific spreadsheet ID
  var spreadsheetId = "13JHcVgKkPVHWy6Kmc7rWoRS0GyI0QL0BEI4qj03DRbE";
  var ss = SpreadsheetApp.openById(spreadsheetId);

//==========================================================================================================================================================================================
// THE FOLLOWING IS ALL YOU NEED TO TOUCH, DONT TOUCH ANYTHING ELSE!!!! IF YOU HAVE PROBLEMS ASK INTERNS: JANRED OR MUHAMMAD
//==========================================================================================================================================================================================

  // CHANGE THE sourceSheetName PER EVENT
  var sourceSheetName = "eventa"; // The sheet name (when linking from individual google form new tab will be created called something "form response x"by default. 
  //rename that tab, and put that name here ) 

  // ONLY CHANGE THIS IF YOU RENAME THE SHEET THAT THE MEMBER INFO GET COLLECTED AT
  var targetSheetName = "Member Overview"; //This is where all the member info gets collected, i.e the final spot for everything

  //CHANGE THE pointsForThisEvent PER EVENT
  var pointsForThisEvent = 10; // Points to add for this event

//==========================================================================================================================================================================================

  var sourceSheet = ss.getSheetByName(sourceSheetName);
  var targetSheet = ss.getSheetByName(targetSheetName);

  // If the target sheet doesn't exist, create it
  if (!targetSheet) {
    targetSheet = ss.insertSheet(targetSheetName);
    Logger.log("Sheet 'Member Overview' created successfully.");
  }

  // Ensure the header row exists
  var lastColumn = targetSheet.getLastColumn();
  if (lastColumn === 0) {
    lastColumn = 5; // Default to at least 5 columns for headers
  }
  var headers = targetSheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  if (headers.length < 5 || headers[0] !== "First Name" || headers[1] !== "Last Name" || headers[2] !== "Email" || headers[3] !== "UH ID" || headers[4] !== "Total Points") {
    targetSheet.getRange(1, 1, 1, 5).setValues([["First Name", "Last Name", "Email", "UH ID", "Total Points"]]);
    Logger.log("Headers added or fixed in the 'Member Overview' sheet.");
  }

  // Get the last row of form responses (new submission)
  var lastRow = sourceSheet.getLastRow();
  var formData = sourceSheet.getRange(lastRow, 1, 1, sourceSheet.getLastColumn()).getValues()[0];

  // Extract form fields
  var timestamp = formData[0];
  var firstName = formData[1] ? formData[1].trim() : "";
  var lastName = formData[2] ? formData[2].trim() : "";
  var email = formData[3] ? String(formData[3]).trim() : "";
  var uhID = formData[4] ? String(formData[4]).trim() : "";

  // Get all data from the target sheet
  var data = targetSheet.getDataRange().getValues();
  headers = data[0];
  var firstNameColumnIndex = 0;
  var lastNameColumnIndex = 1;
  var emailColumnIndex = 2;
  var uhIDColumnIndex = 3;
  var pointsColumnIndex = 4;
  var eventColumnIndex = headers.indexOf(sourceSheetName);

  // If the event does not exist in the header row, add it as a new column
  if (eventColumnIndex === -1) {
    var newColumn = headers.length + 1;
    targetSheet.getRange(1, newColumn).setValue(sourceSheetName);
    targetSheet.getRange(1, newColumn).setFontWeight("bold").setBackground("#f1f3f4");
    eventColumnIndex = newColumn - 1;
    headers = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0];
  }

  // Locate the row of the user by UH ID and update their info
  var found = false;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][uhIDColumnIndex]).trim() === uhID) {
      found = true;
      if (firstName) targetSheet.getRange(i + 1, firstNameColumnIndex + 1).setValue(firstName);
      if (lastName) targetSheet.getRange(i + 1, lastNameColumnIndex + 1).setValue(lastName);
      var existingEmails = String(data[i][emailColumnIndex] || "").split(", ");
      if (existingEmails.indexOf(email) === -1) {
        existingEmails.push(email);
        targetSheet.getRange(i + 1, emailColumnIndex + 1).setValue(existingEmails.join(", "));
      }
      // Update event points and set background to green
      targetSheet.getRange(i + 1, eventColumnIndex + 1).setValue(pointsForThisEvent);
      targetSheet.getRange(i + 1, eventColumnIndex + 1).setBackground("#b7e1cd");
      
      // Recalculate Total Points starting after the 'Total Points' column header
      var updatedRow = targetSheet.getRange(i + 1, 1, 1, targetSheet.getLastColumn()).getValues()[0];
      var eventStartIndex = headers.indexOf("Total Points") + 1;
      var totalPoints = 0;
      for (var j = eventStartIndex; j < headers.length; j++) {
        var eventPoints = Number(updatedRow[j]) || 0;
        totalPoints += eventPoints;
      }
      targetSheet.getRange(i + 1, pointsColumnIndex + 1).setValue(totalPoints);
      break;
    }
  }

  // If user not found, add a new row with their details
  if (!found) {
    var newRow = data.length + 1;
    targetSheet.getRange(newRow, 1, 1, 5).setValues([[firstName, lastName, email, uhID, pointsForThisEvent]]);
    targetSheet.getRange(newRow, eventColumnIndex + 1).setValue(pointsForThisEvent);
    targetSheet.getRange(newRow, eventColumnIndex + 1).setBackground("#b7e1cd");
  }

  // Resize event columns for better viewing
  for (var col = 6; col <= targetSheet.getLastColumn(); col++) {
    targetSheet.setColumnWidth(col, 100);
  }
  targetSheet.autoResizeColumn(emailColumnIndex + 1);
  targetSheet.getDataRange().setHorizontalAlignment("center");

  Logger.log("Form response processed and added to 'Member Overview'.");

  // Update Home and Leaderboard tabs
  try {
    updateHomeTab();
    updateLeaderboardTab();
    Logger.log("Home and Leaderboard updated successfully.");
  } catch (ex) {
    Logger.log("Error updating Home/Leaderboard: " + ex);
  }
}
