
//copy paste this code into google appscript for each form 
//see readme for more instructions

function onFormSubmit(e) {
  // The specific spreadsheet ID
  var spreadsheetId = "13JHcVgKkPVHWy6Kmc7rWoRS0GyI0QL0BEI4qj03DRbE";

  // Open the spreadsheet
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);




 // Change these two variables each form 
  var sourceSheetName = "eventa"; // The sheet name (when linking will be named to "form response x", rename it if you want and put that sheet name here)
  var eventName = "Event A"; // What will display on the main sheet
  // Change two above variables




  var targetSheetName = "main";

  // Open the source sheet
  var sourceSheet = spreadsheet.getSheetByName(sourceSheetName);

  // Open the target sheet (main)
  var targetSheet = spreadsheet.getSheetByName(targetSheetName);

  // If the target sheet doesn't exist, create it and add headers
  if (!targetSheet) {
    targetSheet = spreadsheet.insertSheet(targetSheetName);
    Logger.log("Sheet 'main' created successfully.");
  }

  // Check if the header row is missing, and re-add it if necessary
  var headers = targetSheet.getRange(1, 1, 1, 5).getValues()[0];
  if (
    headers[0] !== "Name" ||
    headers[1] !== "Membership Status" ||
    headers[2] !== "Email(s)" ||
    headers[3] !== "Points" ||
    headers[4] !== "Events Attended"
  ) {
    targetSheet.getRange(1, 1, 1, 5).setValues([
      ["Name", "Membership Status", "Email(s)", "Points", "Events Attended"],
    ]);
    Logger.log("Headers added or fixed in the 'main' sheet.");
  }

  // Format the sheet to look like a table
  targetSheet.getRange(1, 1, 1, 5).setFontWeight("bold").setBackground("#f1f3f4"); // Format headers
  targetSheet.setFrozenRows(1); // Freeze headers
  targetSheet.setColumnWidths(1, 5, 150); // Adjust column widths
  targetSheet.getDataRange().setHorizontalAlignment("center"); // Center-align all content

  // Get the last row of data from the form responses (new submission)
  var lastRow = sourceSheet.getLastRow();
  var formData = sourceSheet
    .getRange(lastRow, 1, 1, sourceSheet.getLastColumn())
    .getValues()[0]; // Get the new row

  // Extract form fields
  var timestamp = formData[0];
  var name = formData[1].trim().toLowerCase(); // Normalize: trim spaces and convert to lowercase
  var memberStatus = formData[2].trim().toLowerCase() === "yes" ? "Yes" : "No"; // Normalize membership status
  var email = formData[3].trim(); // Trim spaces from email

  // Get all data from the main sheet
  var data = targetSheet.getDataRange().getValues();
  var found = false;

  // Find existing or empty row
  for (var i = 1; i < data.length; i++) {
    // Start from 1 to skip headers
    var existingName = data[i][0] ? data[i][0].trim().toLowerCase() : ""; // Normalize existing name
    if (existingName === name) {
      // Check if name matches
      found = true;

      // Handle membership status logic
      var currentStatus = data[i][1];
      if (memberStatus === "Yes") {
        targetSheet.getRange(i + 1, 2).setValue("Yes"); // Set to Yes if they became a member again
      } else if (currentStatus === "Yes" && memberStatus === "No") {
        targetSheet.getRange(i + 1, 2).setValue("Previous Member"); // If they were a member and are now not
      }

      // Add email if not already present
      var existingEmails = data[i][2] ? data[i][2].split(", ") : [];
      if (!existingEmails.includes(email)) {
        existingEmails.push(email);
        targetSheet.getRange(i + 1, 3).setValue(existingEmails.join(", "));
      }

      // Add event name to "Events Attended" if not already present
      var eventsAttended = data[i][4] ? data[i][4].split(", ") : [];
      if (!eventsAttended.includes(eventName)) {
        eventsAttended.push(eventName);
        targetSheet.getRange(i + 1, 5).setValue(eventsAttended.join(", "));

        // Increment points
        var points = data[i][3] || 0; // Default points to 0 if empty
        targetSheet.getRange(i + 1, 4).setValue(points + 10);
      } else {
        Logger.log("No additional points awarded: Event already attended.");
      }

      break;
    }
  }

  // If name not found, add a new row
  if (!found) {
    var initialStatus = memberStatus === "Yes" ? "Yes" : "No";
    var emptyRow = targetSheet.getLastRow() + 1;
    targetSheet
      .getRange(emptyRow, 1, 1, 5)
      .setValues([[name, initialStatus, email, 10, eventName]]);
  }

  // Resize all columns to fit their content dynamically
  var totalColumns = targetSheet.getLastColumn();
  for (var col = 1; col <= totalColumns; col++) {
    targetSheet.autoResizeColumn(col); // Auto-resize each column
  }

  Logger.log("Form response processed and added to 'main'.");
}
