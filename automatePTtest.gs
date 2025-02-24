function onFormSubmit(e) {
  // The specific spreadsheet ID
  var spreadsheetId = "13JHcVgKkPVHWy6Kmc7rWoRS0GyI0QL0BEI4qj03DRbE";

  // Open the spreadsheet
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);

//==========================================================================================================================================================================================
// THE FOLLOWING IS ALL YOU NEED TO TOUCH, DONT TOUCH ANYTHING ELSE!!!! IF YOU HAVE PROBLEMS ASK INTERNS: JANRED OR MUHAMMAD
//==========================================================================================================================================================================================

  // CHANGE THE sourceSheetName PER EVENT
  var sourceSheetName = "eventc"; // The sheet name (when linking from individual google form new tab will be created called something "form response x"by default. 
  //rename that tab, and put that name here ) 

  // ONLY CHANGE THIS IF YOU RENAME THE SHEET THAT THE MEMBER INFO GET COLLECTED AT
  var targetSheetName = "Member Overview"; //This is where all the member info gets collected, i.e the final spot for everything

  //CHANGE THE pointsForThisEvent PER EVENT
  var pointsForThisEvent = 20; // Points to add for this event

//==========================================================================================================================================================================================

  // Open the source and target sheets
  var sourceSheet = spreadsheet.getSheetByName(sourceSheetName);
  var targetSheet = spreadsheet.getSheetByName(targetSheetName);

  // If the target sheet doesn't exist, create it
  if (!targetSheet) {
    targetSheet = spreadsheet.insertSheet(targetSheetName);
    Logger.log("Sheet 'main' created successfully.");
  }

  // Get the last row of data from the form responses (new submission)
  var lastRow = sourceSheet.getLastRow();
  var formData = sourceSheet
    .getRange(lastRow, 1, 1, sourceSheet.getLastColumn())
    .getValues()[0]; // Get the new row

  // Extract form fields
  var timestamp = formData[0];
  var name = formData[1].trim().toLowerCase(); // Normalize name
  var memberStatus = formData[2].trim().toLowerCase() === "yes" ? "Yes" : "No"; // Normalize membership status
  var email = formData[3].trim(); // Trim email input

  // Get all data from the main sheet
  var data = targetSheet.getDataRange().getValues();
  var headers = data[0]; // Get headers (first row)
  var nameColumnIndex = 0; // Name column is A (index 0)
  var pointsColumnIndex = 3; // Points column is D (index 3)
  var eventColumnIndex = headers.indexOf(sourceSheetName);

  // If the event does not exist in the header row, add it as a new column
  if (eventColumnIndex === -1) {
    var newColumn = headers.length + 1; // Determine next available column
    targetSheet.getRange(1, newColumn).setValue(sourceSheetName); // Add event as a header
    targetSheet.getRange(1, newColumn).setFontWeight("bold").setBackground("#f1f3f4"); // Format header
    eventColumnIndex = newColumn - 1; // Update column index
  }

  // Locate the row of the user
  var found = false;
  for (var i = 1; i < data.length; i++) {
    if (data[i][nameColumnIndex].trim().toLowerCase() === name) {
      found = true;

      // Update membership status
      var currentStatus = data[i][1];
      if (memberStatus === "Yes") {
        targetSheet.getRange(i + 1, 2).setValue("Yes");
      } else if (currentStatus === "Yes" && memberStatus === "No") {
        targetSheet.getRange(i + 1, 2).setValue("Previous Member");
      }

      // Update email if not already present
      var existingEmails = data[i][2] ? data[i][2].split(", ") : [];
      if (!existingEmails.includes(email)) {
        existingEmails.push(email);
        targetSheet.getRange(i + 1, 3).setValue(existingEmails.join(", "));
      }

      // Check if the user has already attended this event
      var existingPointsForEvent = data[i][eventColumnIndex] || 0;
      if (existingPointsForEvent === 0) {
        // Add points for the event
        targetSheet.getRange(i + 1, eventColumnIndex + 1).setValue(pointsForThisEvent);
      } else {
        Logger.log("No additional points awarded: Event already recorded.");
      }

      // **Fix: Properly Recalculate Total Points**
      var totalPoints = 0;
      for (var j = 4; j < headers.length; j++) { // Start checking from event columns
        totalPoints += data[i][j] || 0;
      }
      targetSheet.getRange(i + 1, pointsColumnIndex + 1).setValue(totalPoints);

      break;
    }
  }

  // If user not found, add a new row
  if (!found) {
    var newRow = data.length + 1;
    targetSheet.getRange(newRow, 1, 1, 4).setValues([[name, memberStatus, email, pointsForThisEvent]]);
    targetSheet.getRange(newRow, eventColumnIndex + 1).setValue(pointsForThisEvent);
  }

  // Resize all columns to fit content dynamically
  var totalColumns = targetSheet.getLastColumn();
  for (var col = 1; col <= totalColumns; col++) {
    targetSheet.autoResizeColumn(col);
  }

  Logger.log("Form response processed and added to 'main'.");
}
