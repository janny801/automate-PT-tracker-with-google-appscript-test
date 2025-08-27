function onFormSubmit(e) {
  var lock = LockService.getPublicLock();
  
  try {
    // Use a reasonable timeout (1 hour max)
    if (!lock.tryLock(3600000)) {
      Logger.log('Could not obtain lock after an hour. Form submission skipped.');
      return;
    }
    
    processFormSubmission(e);
    
  } catch (error) {
    Logger.log('Error: ' + error.toString());
  } finally {
    lock.releaseLock();
  }
}

function processFormSubmission(e) {
  var spreadsheetId = "1TJ-tA75Dhx2Ti0oWFTuioSOLjbBGxUyYFwfg8FAsaLI";
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  
  // Configuration
  var sourceSheetName = "9/3 Icebreaker Social Sign-In";
  var targetSheetName = "Member Overview";
  var pointsForThisEvent = 10;

  // Extract form data
  var values = e.values;
  var firstName = values[1] ? String(values[1]).trim() : "";
  var lastName = values[2] ? String(values[2]).trim() : "";
  var email = values[3] ? String(values[3]).trim() : "";
  var uhID = values[4] ? String(values[4]).trim() : "";

  if (!uhID) {
    Logger.log('Skipping submission - missing UH ID');
    return;
  }

  // Get or create target sheet
  var targetSheet = spreadsheet.getSheetByName(targetSheetName) || spreadsheet.insertSheet(targetSheetName);
  
  // Get ALL data at once to minimize reads
  var targetDataRange = targetSheet.getDataRange();
  var targetData = targetDataRange.getValues();
  var targetHeaders = targetData[0] || [];
  
  // Ensure headers exist and get indexes
  var headerIndexes = ensureHeaders(targetSheet, targetHeaders);
  
  // Get membership data (cache this if possible)
  var paidMap = getMembershipPaidMap(spreadsheet);
  var paidStatus = paidMap[uhID] || "Unpaid";
  
  // Ensure event column exists
  var eventColumnIndex = ensureEventColumn(targetSheet, targetHeaders, sourceSheetName);
  
  // Find existing row or create new one
  var uhIDIndex = headerIndexes.uhID;
  var existingRowIndex = -1;
  
  for (var i = 1; i < targetData.length; i++) {
    if (String(targetData[i][uhIDIndex]).trim() === uhID) {
      existingRowIndex = i;
      break;
    }
  }

  if (existingRowIndex !== -1) {
    updateExistingRow(targetSheet, existingRowIndex, headerIndexes, {
      firstName: firstName,
      lastName: lastName,
      email: email,
      uhID: uhID,
      points: pointsForThisEvent,
      paidStatus: paidStatus,
      eventColumnIndex: eventColumnIndex
    });
  } else {
    createNewRow(targetSheet, headerIndexes, {
      firstName: firstName,
      lastName: lastName,
      email: email,
      uhID: uhID,
      points: pointsForThisEvent,
      paidStatus: paidStatus,
      eventColumnIndex: eventColumnIndex
    });
  }

  // Formatting (do this once at the end)
  formatSheet(targetSheet, headerIndexes.email);
  createOrUpdateEventMultipliers();
  createLeaderboard();
  createOrUpdateDashboard();
  Logger.log("Form response processed successfully for UH ID: " + uhID);
}

function ensureHeaders(targetSheet, currentHeaders) {
  var desiredHeaders = ["First Name", "Last Name", "Email", "UH ID", "Paid Status", "Total Points"];
  var headerIndexes = {};
  
  // Create missing headers
  if (currentHeaders.length === 0 || currentHeaders[0] !== "First Name") {
    targetSheet.getRange(1, 1, 1, desiredHeaders.length).setValues([desiredHeaders]);
    currentHeaders = desiredHeaders.concat([]);
  }
  
  // Get indexes
  headerIndexes.firstName = currentHeaders.indexOf("First Name");
  headerIndexes.lastName = currentHeaders.indexOf("Last Name");
  headerIndexes.email = currentHeaders.indexOf("Email");
  headerIndexes.uhID = currentHeaders.indexOf("UH ID");
  headerIndexes.paidStatus = currentHeaders.indexOf("Paid Status");
  headerIndexes.totalPoints = currentHeaders.indexOf("Total Points");
  
  // Ensure Paid Status exists after UH ID
  if (headerIndexes.paidStatus === -1) {
    var uhIdIndex = headerIndexes.uhID;
    targetSheet.insertColumnAfter(uhIdIndex + 1);
    targetSheet.getRange(1, uhIdIndex + 2).setValue("Paid Status")
               .setFontWeight("bold").setBackground("#f1f3f4");
    headerIndexes.paidStatus = uhIdIndex + 1;
  }
  
  // Ensure Total Points exists
  if (headerIndexes.totalPoints === -1) {
    var newCol = currentHeaders.length + 1;
    targetSheet.getRange(1, newCol).setValue("Total Points")
               .setFontWeight("bold").setBackground("#f1f3f4");
    headerIndexes.totalPoints = newCol - 1;
  }
  
  return headerIndexes;
}

function getMembershipPaidMap(spreadsheet) {
  var membershipSheet = spreadsheet.getSheetByName("Membership");
  if (!membershipSheet) return {};
  
  var data = membershipSheet.getDataRange().getValues();
  var paidMap = {};
  
  for (var i = 1; i < data.length; i++) {
    var uhID = data[i][2] ? String(data[i][2]).trim() : ""; // Column C (3)
    var payable = data[i][25] ? String(data[i][25]).trim() : ""; // Column Z (26)
    
    if (uhID) {
      paidMap[uhID] = /paid/i.test(payable) ? "Paid" : "Unpaid";
    }
  }
  
  return paidMap;
}

function ensureEventColumn(targetSheet, headers, eventName) {
  var eventIndex = headers.indexOf(eventName);
  
  if (eventIndex === -1) {
    var newColIndex = headers.length + 1;
    targetSheet.getRange(1, newColIndex).setValue(eventName)
               .setFontWeight("bold").setBackground("#f1f3f4");
    return newColIndex - 1; // Return 0-based index
  }
  
  return eventIndex;
}

function updateExistingRow(targetSheet, rowIndex, headers, data) {
  var rowNum = rowIndex + 1;
  
  // Update basic info if provided
  if (data.firstName) {
    targetSheet.getRange(rowNum, headers.firstName + 1).setValue(data.firstName);
  }
  if (data.lastName) {
    targetSheet.getRange(rowNum, headers.lastName + 1).setValue(data.lastName);
  }
  
  // Update email (append unique)
  var currentEmail = targetSheet.getRange(rowNum, headers.email + 1).getValue();
  var emails = currentEmail ? String(currentEmail).split(", ").filter(Boolean) : [];
  if (data.email && emails.indexOf(data.email) === -1) {
    emails.push(data.email);
    targetSheet.getRange(rowNum, headers.email + 1).setValue(emails.join(", "));
  }
  
  // Set event points
  targetSheet.getRange(rowNum, data.eventColumnIndex + 1)
             .setValue(data.points)
             .setBackground("#b7e1cd");
  
  // Update paid status
  targetSheet.getRange(rowNum, headers.paidStatus + 1)
             .setValue(data.paidStatus)
             .setBackground(data.paidStatus === "Paid" ? "#c6efce" : "#ffc7ce");
  
  // Recalculate total points (including membership bonus if paid)
  recalculateTotalPoints(targetSheet, rowNum, headers, data.paidStatus);
}

function createNewRow(targetSheet, headers, data) {
  var newRow = targetSheet.getLastRow() + 1;
  
  // Calculate total points (event points + membership bonus if paid)
  var totalPoints = data.points;
  if (data.paidStatus === "Paid") {
    totalPoints += 50; // Add membership bonus
  }
  
  // Create base row data
  var rowData = new Array(headers.totalPoints + 1).fill("");
  rowData[headers.firstName] = data.firstName;
  rowData[headers.lastName] = data.lastName;
  rowData[headers.email] = data.email;
  rowData[headers.uhID] = data.uhID;
  rowData[headers.paidStatus] = data.paidStatus;
  rowData[headers.totalPoints] = totalPoints; // Includes bonus if paid
  rowData[data.eventColumnIndex] = data.points; // Event points only
  
  // Set the entire row at once
  targetSheet.getRange(newRow, 1, 1, rowData.length).setValues([rowData]);
  
  // Format specific cells
  targetSheet.getRange(newRow, data.eventColumnIndex + 1).setBackground("#b7e1cd");
  targetSheet.getRange(newRow, headers.paidStatus + 1)
             .setBackground(data.paidStatus === "Paid" ? "#c6efce" : "#ffc7ce");
}

function recalculateTotalPoints(targetSheet, rowNum, headers, paidStatus) {
  var rowData = targetSheet.getRange(rowNum, 1, 1, targetSheet.getLastColumn()).getValues()[0];
  var totalPoints = 0;
  var startIndex = headers.totalPoints + 1; // Columns after Total Points
  
  // Add points from all events
  for (var j = startIndex; j < rowData.length; j++) {
    totalPoints += Number(rowData[j]) || 0;
  }
  
  // Add 50 bonus points if member is paid
  if (paidStatus === "Paid") {
    totalPoints += 50;
  }
  
  targetSheet.getRange(rowNum, headers.totalPoints + 1).setValue(totalPoints);
}

function formatSheet(targetSheet, emailColumnIndex) {
  // Resize columns
  var lastCol = targetSheet.getLastColumn();
  for (var col = 7; col <= lastCol; col++) {
    targetSheet.setColumnWidth(col, 100);
  }
  
  targetSheet.autoResizeColumn(emailColumnIndex + 1);
  targetSheet.getDataRange().setHorizontalAlignment("center");
}





// vv     run func below to sync membership sheet paid status to member overview 

// Sync Membership sheet Paid Status → Member Overview
function syncMembershipToOverview() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var membership = ss.getSheetByName("Membership");
  var overview = ss.getSheetByName("Member Overview");
  if (!membership || !overview) return;

  var memData = membership.getDataRange().getValues();   // Membership data
  var ovData = overview.getDataRange().getValues();      // Member Overview data
  var headers = ovData[0];
  
  var uhIdIndexOV = headers.indexOf("UH ID");
  var paidStatusIndexOV = headers.indexOf("Paid Status");
  var totalPointsIndexOV = headers.indexOf("Total Points");

  if (uhIdIndexOV === -1 || paidStatusIndexOV === -1 || totalPointsIndexOV === -1) {
    Logger.log("Missing UH ID / Paid Status / Total Points in Member Overview");
    return;
  }

  // Build lookup for Overview rows by UH ID
  var ovMap = {};
  for (var i = 1; i < ovData.length; i++) {
    var uh = String(ovData[i][uhIdIndexOV]).trim();
    if (uh) {
      ovMap[uh] = i + 1; // store row number (1-based)
    }
  }

  // Loop Membership rows
  for (var j = 1; j < memData.length; j++) {
    var uhID = String(memData[j][2]).trim();  // UH ID col C (index 2)
    var payable = String(memData[j][25]).trim(); // Payable col Z (index 25)
    if (!uhID) continue;

    var newPaid = /paid/i.test(payable) ? "Paid" : "Unpaid";
    var rowNumOV = ovMap[uhID];

    if (rowNumOV) {
      // Update Paid Status in Overview
      var cell = overview.getRange(rowNumOV, paidStatusIndexOV + 1);
      cell.setValue(newPaid)
          .setBackground(newPaid === "Paid" ? "#c6efce" : "#ffc7ce");

      // Recalculate total points for that row
      recalculateTotalPoints(overview, rowNumOV, {
        totalPoints: totalPointsIndexOV
      }, newPaid);
    }
  }

  Logger.log("Sync completed between Membership → Member Overview");
}








