function onFormSubmit(e) {
  // The specific spreadsheet ID
  var spreadsheetId = "1TJ-tA75Dhx2Ti0oWFTuioSOLjbBGxUyYFwfg8FAsaLI";

  // Open the spreadsheet
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);

  //==========================================================================================================================================================================================
  // THE FOLLOWING IS ALL YOU NEED TO TOUCH, DONT TOUCH ANYTHING ELSE!!!! IF YOU HAVE PROBLEMS ASK INTERNS: JANRED OR MUHAMMAD
  //==========================================================================================================================================================================================

  // CHANGE THE sourceSheetName PER EVENT
  var sourceSheetName = "9/4 Icebreaker Social Sign-In"; // The sheet name (when linking from individual google form new tab will be created called something "form response x"by default. 
  //rename that tab, and put that name here ) 

  // ONLY CHANGE THIS IF YOU RENAME THE SHEET THAT THE MEMBER INFO GET COLLECTED AT
  var targetSheetName = "Member Overview"; //This is where all the member info gets collected, i.e the final spot for everything

  //CHANGE THE pointsForThisEvent PER EVENT
  var pointsForThisEvent = 10; // Points to add for this event

  //==========================================================================================================================================================================================

  // Open the source and target sheets
  var sourceSheet = spreadsheet.getSheetByName(sourceSheetName);
  var targetSheet = spreadsheet.getSheetByName(targetSheetName);

  // If the target sheet doesn't exist, create it
  if (!targetSheet) {
    targetSheet = spreadsheet.insertSheet(targetSheetName);
    Logger.log("Sheet 'Member Overview' created successfully.");
  }

  // ---------------------- HEADER SETUP (always ensure Paid Status is after UH ID) ----------------------
  // If empty or missing headers, set them with Paid Status in the correct spot
  var lastColumn = Math.max(1, targetSheet.getLastColumn());
  var headers = targetSheet.getRange(1, 1, 1, lastColumn).getValues()[0].filter(String);

  // Desired base headers in order:
  var desiredBaseHeaders = ["First Name", "Last Name", "Email", "UH ID", "Paid Status", "Total Points"];

  if (headers.length < 6 ||
      headers[0] !== "First Name" ||
      headers[1] !== "Last Name" ||
      headers[2] !== "Email" ||
      headers[3] !== "UH ID") {
    // Reset the first 6 headers exactly as desired
    targetSheet.getRange(1, 1, 1, desiredBaseHeaders.length).setValues([desiredBaseHeaders]);
    headers = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0];
  } else {
    // Ensure "Paid Status" exists AFTER UH ID
    var paidStatusHeader = "Paid Status";
    var uhIdIndexZeroBased = headers.indexOf("UH ID"); // should be 3
    var paidIdx = headers.indexOf(paidStatusHeader);
    if (paidIdx === -1) {
      // Insert a new column after UH ID (1-based index)
      targetSheet.insertColumnAfter(uhIdIndexZeroBased + 1);
      targetSheet.getRange(1, uhIdIndexZeroBased + 2).setValue(paidStatusHeader)
                 .setFontWeight("bold").setBackground("#f1f3f4");
    } else if (paidIdx !== uhIdIndexZeroBased + 1) {
      // If Paid Status exists but is in the wrong place, move it next to UH ID
      var fromCol = paidIdx + 1;                      // 1-based
      var toCol = uhIdIndexZeroBased + 2;             // 1-based, right after UH ID
      targetSheet.moveColumns(targetSheet.getRange(1, fromCol, 1, 1), toCol);
    }
    // Ensure "Total Points" exists (append if missing)
    headers = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0];
    if (headers.indexOf("Total Points") === -1) {
      var newCol = headers.length + 1;
      targetSheet.getRange(1, newCol).setValue("Total Points").setFontWeight("bold").setBackground("#f1f3f4");
    }
    headers = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0];
  }

  // Column indexes (0-based) after enforcing header order
  headers = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0];
  var firstNameColumnIndex = headers.indexOf("First Name"); // 0
  var lastNameColumnIndex  = headers.indexOf("Last Name");  // 1
  var emailColumnIndex     = headers.indexOf("Email");      // 2
  var uhIDColumnIndex      = headers.indexOf("UH ID");      // 3
  var paidStatusColIndex   = headers.indexOf("Paid Status");// 4
  var pointsColumnIndex    = headers.indexOf("Total Points");// 5

  // ---------------------- SOURCE ROW ----------------------
  var lastRow = sourceSheet.getLastRow();
  if (lastRow < 2) {
    Logger.log("No data rows in source sheet.");
    return;
  }
  var formData = sourceSheet.getRange(lastRow, 1, 1, sourceSheet.getLastColumn()).getValues()[0];

  // Extract form fields (adjust if your form columns differ)
  var firstName = formData[1] ? String(formData[1]).trim() : "";
  var lastName  = formData[2] ? String(formData[2]).trim() : "";
  var email     = formData[3] ? String(formData[3]).trim() : "";
  var uhID      = formData[4] ? String(formData[4]).trim() : "";

  // ---------------------- MEMBERSHIP MAP (UHID -> Paid/Unpaid) ----------------------
  var membershipSheet = spreadsheet.getSheetByName("Membership");
  var membershipLastRow = membershipSheet.getLastRow();
  var membershipLastCol = membershipSheet.getLastColumn();
  var membershipValues = membershipLastRow > 1
    ? membershipSheet.getRange(2, 1, membershipLastRow - 1, membershipLastCol).getValues()
    : [];

  var UHID_COL_ABS = 3;   // C
  var PAYABLE_COL_ABS = 26; // Z
  var uhidOffset = UHID_COL_ABS - 1;
  var payableOffset = PAYABLE_COL_ABS - 1;

  var paidMap = {};
  for (var r = 0; r < membershipValues.length; r++) {
    var row = membershipValues[r];
    var memUHID = row[uhidOffset] != null ? String(row[uhidOffset]).trim() : "";
    var payableRaw = row[payableOffset] != null ? String(row[payableOffset]).trim() : "";
    if (memUHID) {
      paidMap[memUHID] = /paid/i.test(payableRaw) ? "Paid" : "Unpaid";
    }
  }
  function getPaidStatusForUHID(id) {
    var key = id != null ? String(id).trim() : "";
    return paidMap[key] || "Unpaid";
  }

  // ---------------------- EVENT COLUMN (create if missing) ----------------------
  var eventColumnIndex = headers.indexOf(sourceSheetName);
  if (eventColumnIndex === -1) {
    var newEventCol = headers.length + 1;
    targetSheet.getRange(1, newEventCol).setValue(sourceSheetName)
               .setFontWeight("bold").setBackground("#f1f3f4");
    headers = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0];
    eventColumnIndex = headers.indexOf(sourceSheetName);
  }

  // ---------------------- UPSERT ROW BY UH ID ----------------------
  var data = targetSheet.getDataRange().getValues();
  var found = false;

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][uhIDColumnIndex]).trim() === uhID) {
      found = true;

      if (firstName) targetSheet.getRange(i + 1, firstNameColumnIndex + 1).setValue(firstName);
      if (lastName)  targetSheet.getRange(i + 1, lastNameColumnIndex + 1).setValue(lastName);

      // email (append unique)
      var existingEmails = String(data[i][emailColumnIndex] || "").split(", ").filter(Boolean);
      if (email && existingEmails.indexOf(email) === -1) {
        existingEmails.push(email);
        targetSheet.getRange(i + 1, emailColumnIndex + 1).setValue(existingEmails.join(", "));
      }

      // event points
      targetSheet.getRange(i + 1, eventColumnIndex + 1).setValue(pointsForThisEvent)
                 .setBackground("#b7e1cd");

      // recompute total points (sum from first event column: column after Total Points)
      var updatedRow = targetSheet.getRange(i + 1, 1, 1, targetSheet.getLastColumn()).getValues()[0];
      var eventStartIndex = headers.indexOf("Total Points") + 1;
      var totalPoints = 0;
      for (var j = eventStartIndex; j < headers.length; j++) {
        totalPoints += Number(updatedRow[j]) || 0;
      }
      targetSheet.getRange(i + 1, pointsColumnIndex + 1).setValue(totalPoints);

      // paid status
      var paidStatus = getPaidStatusForUHID(uhID);
      targetSheet.getRange(i + 1, paidStatusColIndex + 1).setValue(paidStatus)
                 .setBackground(paidStatus === "Paid" ? "#c6efce" : "#ffc7ce");
      break;
    }
  }

  if (!found) {
    var newRow = targetSheet.getLastRow() + 1;
    var paidForNew = getPaidStatusForUHID(uhID);

    // write base columns including Paid Status and Total Points
    targetSheet.getRange(newRow, 1, 1, 6).setValues([[
      firstName, lastName, email, uhID, paidForNew, pointsForThisEvent
    ]]);

    // set event points cell
    targetSheet.getRange(newRow, eventColumnIndex + 1).setValue(pointsForThisEvent)
               .setBackground("#b7e1cd");

    // color paid status cell
    targetSheet.getRange(newRow, paidStatusColIndex + 1)
               .setBackground(paidForNew === "Paid" ? "#c6efce" : "#ffc7ce");
  }

  // Resize: events start at column 7 now (after Total Points)
  for (var col = 7; col <= targetSheet.getLastColumn(); col++) {
    targetSheet.setColumnWidth(col, 100);
  }
  targetSheet.autoResizeColumn(emailColumnIndex + 1);
  targetSheet.getDataRange().setHorizontalAlignment("center");

  Logger.log("Form response processed and added to 'Member Overview'.");

  createOrUpdateEventMultipliers();
  createLeaderboard();
  createOrUpdateDashboard();
}
