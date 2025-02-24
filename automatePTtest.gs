function onFormSubmit(e) {
  // The specific spreadsheet ID
  var spreadsheetId = "13JHcVgKkPVHWy6Kmc7rWoRS0GyI0QL0BEI4qj03DRbE";

  // Open the spreadsheet
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);

//==========================================================================================================================================================================================
// THE FOLLOWING IS ALL YOU NEED TO TOUCH, DONT TOUCH ANYTHING ELSE!!!! IF YOU HAVE PROBLEMS ASK INTERNS: JANRED OR MUHAMMAD
//==========================================================================================================================================================================================

  // CHANGE THE sourceSheetName PER EVENT
  var sourceSheetName = "eventa"; // The sheet name (when linking from individual google form new tab will be created called something "form response x"by default. 
  //rename that tab, and put that name here ) 

  // ONLY CHANGE THIS IF YOU RENAME THE SHEET THAT THE MEMBER INFO GET COLLECTED AT
  var targetSheetName = "Member Overview"; //This is where all the member info gets collected, i.e the final spot for everything

  //CHANGE THE pointsForThisEvent PER EVENT
  var pointsForThisEvent = 100; // Points to add for this event

//==========================================================================================================================================================================================

  // Open the source and target sheets
  var sourceSheet = spreadsheet.getSheetByName(sourceSheetName);
  var targetSheet = spreadsheet.getSheetByName(targetSheetName);

  // If the target sheet doesn't exist, create it
  if (!targetSheet) {
    targetSheet = spreadsheet.insertSheet(targetSheetName);
    Logger.log("Sheet 'main' created successfully.");
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

  // Get the last row of data from the form responses (new submission)
  var lastRow = sourceSheet.getLastRow();
  var formData = sourceSheet
    .getRange(lastRow, 1, 1, sourceSheet.getLastColumn())
    .getValues()[0]; // Get the new row

  // Extract form fields
  var timestamp = formData[0];
  var firstName = formData[1] ? formData[1].trim() : ""; // Normalize first name
  var lastName = formData[2] ? formData[2].trim() : ""; // Normalize last name
  var email = formData[3] ? String(formData[3]).trim() : "";
  var uhID = formData[4] ? String(formData[4]).trim() : ""; // Convert UH ID to string before trimming

  // Get all data from the main sheet
  var data = targetSheet.getDataRange().getValues();
  var headers = data[0]; // Get headers (first row)
  var firstNameColumnIndex = 0; // First Name column is A (index 0)
  var lastNameColumnIndex = 1; // Last Name column is B (index 1)
  var emailColumnIndex = 2; // Email column is C (index 2)
  var uhIDColumnIndex = 3; // UH ID column is D (index 3)
  var pointsColumnIndex = 4; // Points column is E (index 4)
  var eventColumnIndex = headers.indexOf(sourceSheetName);

  // If the event does not exist in the header row, add it as a new column
  if (eventColumnIndex === -1) {
    var newColumn = headers.length + 1; // Determine next available column
    targetSheet.getRange(1, newColumn).setValue(sourceSheetName); // Add event as a header
    targetSheet.getRange(1, newColumn).setFontWeight("bold").setBackground("#f1f3f4"); // Format header
    eventColumnIndex = newColumn - 1; // Update column index
    
    // **Re-fetch headers after adding new column**
    headers = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0];
  }

  // Locate the row of the user using UH ID instead of Name
  var found = false;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][uhIDColumnIndex]).trim() === uhID) {
      found = true;

      // **Update First Name and Last Name to the most recent input**
      if (firstName) {
        targetSheet.getRange(i + 1, firstNameColumnIndex + 1).setValue(firstName);
      }
      if (lastName) {
        targetSheet.getRange(i + 1, lastNameColumnIndex + 1).setValue(lastName);
      }

      // Update email if not already present
      var existingEmails = String(data[i][emailColumnIndex] || "").split(", ");
      if (!existingEmails.includes(email)) {
        existingEmails.push(email);
        targetSheet.getRange(i + 1, emailColumnIndex + 1).setValue(existingEmails.join(", "));
      }

      // **Update the points for the event instead of skipping it**
      targetSheet.getRange(i + 1, eventColumnIndex + 1).setValue(pointsForThisEvent);

      // **Make the cell green for attended events**
      targetSheet.getRange(i + 1, eventColumnIndex + 1).setBackground("#b7e1cd");

      // **Now, re-fetch the updated row after modifying event points**
      var updatedRow = targetSheet.getRange(i + 1, 1, 1, targetSheet.getLastColumn()).getValues()[0];

      // Find the index where event columns start (first column after 'Total Points')
      var eventStartIndex = headers.indexOf("Total Points") + 1;

      // **Fix: Properly Recalculate Total Points**
      var totalPoints = 0;
      for (var j = eventStartIndex; j < headers.length; j++) { // Start summing from first event column
        var eventPoints = Number(updatedRow[j]) || 0; // Convert to number, default to 0 if empty
        totalPoints += eventPoints;
      }

      // Update Total Points column
      targetSheet.getRange(i + 1, pointsColumnIndex + 1).setValue(totalPoints);

      break;
    }
  }

  // If user not found, add a new row with First Name, Last Name, Email, UH ID
  if (!found) {
    var newRow = data.length + 1;
    targetSheet.getRange(newRow, 1, 1, 5).setValues([[firstName, lastName, email, uhID, pointsForThisEvent]]);
    targetSheet.getRange(newRow, eventColumnIndex + 1).setValue(pointsForThisEvent);
    
    // **Make the new event cell green for attended events**
    targetSheet.getRange(newRow, eventColumnIndex + 1).setBackground("#b7e1cd");
  }

  // **Resize event columns for better viewing**
  for (var col = 6; col <= targetSheet.getLastColumn(); col++) {
    targetSheet.setColumnWidth(col, 100); // Adjust width for better visibility
  }

  // **Resize Email column dynamically to fit content**
  targetSheet.autoResizeColumn(emailColumnIndex + 1);

  // **Center all values in the spreadsheet**
  targetSheet.getDataRange().setHorizontalAlignment("center");

  Logger.log("Form response processed and added to 'main'.");
}
