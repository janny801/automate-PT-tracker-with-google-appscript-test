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



  // var email = "jred8069@gmail.com"; //for testing; remove this line 





  sendEventConfirmationEmail(email, firstName, sourceSheetName, paidStatus, pointsForThisEvent,uhID);

}






/**
 * Sends an event confirmation email after a form submission.
 * Called at the end of processFormSubmission().
 */


function sendEventConfirmationEmail(email, firstName, eventName, paidStatus, points, uhID) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

// Wait to allow point updates (especially right after becoming a member)
Utilities.sleep(1000); 


  var overviewSheet = spreadsheet.getSheetByName("Member Overview");
  if (!overviewSheet) {
    Logger.log("Member Overview sheet not found");
    return;
  }

  var data = overviewSheet.getDataRange().getValues();
  var headers = data[0];
  var uhidIndex = headers.indexOf("UH ID");
  var totalPointsIndex = headers.indexOf("Total Points");

  if (uhidIndex === -1 || totalPointsIndex === -1) {
    Logger.log("Missing UH ID or Total Points column");
    return;
  }

  var totalPoints = null;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][uhidIndex]).trim() === uhID) {
      totalPoints = data[i][totalPointsIndex];
      break;
    }
  }

  var subject = `Thanks for attending ${eventName}!`;
  var membershipFormLink = "https://docs.google.com/forms/d/e/1FAIpQLSeliIy87t1EeOJkRuPmTHPwPCcAs0Pv0sDi5DNk2Vc2fzx12w/viewform";
  var htmlBody = `<p>Hi ${firstName || "there"},</p>`;

  if (paidStatus === "Paid") {
    htmlBody +=
      `<p>Thank you for attending <strong>${eventName}</strong>, you have received <strong>${points} points</strong>.</p>` +
      `<p><strong>Your total amount of points is ${totalPoints}.</strong></p>` +
      `<p>See you again soon!</p>`;
  } else {
    htmlBody +=
      `<p>Thank you for attending <strong>${eventName}</strong>, this event would've earned you <strong>${points} points</strong>.</p>` +
      `<p>Your total amount of points would've been <strong>${totalPoints}</strong>, plus an additional 50 for your membership!</p>` +
      `<p><a href="${membershipFormLink}"><strong>Sign up here for membership</strong></a> to claim your points!</p>`;
  }

  MailApp.sendEmail({
    to: email,
    subject: subject,
    htmlBody: htmlBody
  });
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


/********************************************
 * EMAIL-ONLY POLLER (works with IMPORTRANGE)
 * - NO writes to Member Overview (your other
 *   syncMembershipToOverview() handles that).
 * - Sends welcome email only on real flip -> Paid.
 ********************************************/

/* === CONFIG: adjust header names if needed === */
var MEMBERSHIP_SHEET_NAME = "Membership";   // IMPORTRANGE tab in your tracker
var HDR_EMAIL       = "Email Address";      // must match header text
var HDR_FIRST_NAME  = "First Name";
var HDR_UHID        = "UH ID";
var HDR_STATUS      = "Payable Status";

// statuses that count as paid (tweak if you use others)
var PAID_STATUSES = /^(paid|paid-cash|paid-cheque|paid-other)$/i;

/* === TEST MODE (REMOVE LATER) ==================
 * only send for one UH ID during testing
 * set to null to disable test filter
 * ============================================= */
var TEST_ONLY_UHID = "2190662"; // <— TEST-ONLY: remove/change later


/**
 * RUN ONCE before enabling the timer trigger.
 * Seeds memory so we DO NOT email people already marked Paid right now.
 */

function seedPaidStatusMemory() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(MEMBERSHIP_SHEET_NAME);
  if (!sh) { Logger.log("Membership sheet not found"); return; }

  var rows = sh.getDataRange().getValues();
  if (rows.length < 2) { Logger.log("No data to seed"); return; }

  var hdr = rows[0];
  var iEmail  = hdr.indexOf(HDR_EMAIL);
  var iFirst  = hdr.indexOf(HDR_FIRST_NAME);
  var iUhid   = hdr.indexOf(HDR_UHID);
  var iStatus = hdr.indexOf(HDR_STATUS);
  if ([iEmail, iFirst, iUhid, iStatus].some(function(i){ return i < 0; })) {
    Logger.log("Header mismatch. Found: " + JSON.stringify(hdr));
    return;
  }

  var emailSentMap = {};
  var skipped = [];

  for (var r = 1; r < rows.length; r++) {
    var uhid = String(rows[r][iUhid] || "").trim();
    var email = String(rows[r][iEmail] || "").trim();
    var first = String(rows[r][iFirst] || "").trim();
    var status = String(rows[r][iStatus] || "").trim();
    if (!uhid || !email) continue;

    if (PAID_STATUSES.test(status)) {
      emailSentMap[uhid] = true;
      skipped.push(`${first} (${uhid}) → ${email}`);
    }
  }

  PropertiesService.getScriptProperties()
    .setProperty("emailSentMap", JSON.stringify(emailSentMap));

  Logger.log("Seed complete. The following members are marked as 'already emailed' and will be skipped:");
  skipped.forEach(line => Logger.log(line));
}



/**
 * POLLER — put this on a time-driven trigger (e.g., every 5 minutes).
 * Detects Unpaid -> Paid transitions and sends the welcome email once.
 * DOES NOT modify Member Overview.
 */
function checkPaidStatusAndSendEmails() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(MEMBERSHIP_SHEET_NAME);
    if (!sh) { Logger.log("Membership sheet not found"); return; }

    var rows = sh.getDataRange().getValues();
    if (rows.length < 2) return;

    var hdr = rows[0];
    var iEmail  = hdr.indexOf(HDR_EMAIL);
    var iFirst  = hdr.indexOf(HDR_FIRST_NAME);
    var iUhid   = hdr.indexOf(HDR_UHID);
    var iStatus = hdr.indexOf(HDR_STATUS);
    if ([iEmail, iFirst, iUhid, iStatus].some(function(i){ return i < 0; })) {
      Logger.log("Header mismatch. Found: " + JSON.stringify(hdr));
      return;
    }

    var props   = PropertiesService.getScriptProperties();
    var sentMem = JSON.parse(props.getProperty("emailSentMap") || "{}");







// // used for testing; 

// // clears uhid;2190662 (janred) from the emailsentmap 
// //   allows for me to receive email everytime trigger runs 



// // Clear sent flag for test UHID so you can receive emails again during testing
// if (TEST_ONLY_UHID && sentMem[TEST_ONLY_UHID]) {
//   delete sentMem[TEST_ONLY_UHID];
//   Logger.log("Reset emailSentMap for UHID " + TEST_ONLY_UHID + " to allow resending.");
// } // remove this later this was for testing 




// // Clear sent flag for test UHIDs so you can receive emails again during testing
// if (sentMem["2266573"]) {
//   delete sentMem["2266573"];
//   Logger.log("Reset emailSentMap for UHID 2266573 to allow resending.");
// }

// if (TEST_ONLY_UHID && sentMem[TEST_ONLY_UHID]) {
//   delete sentMem[TEST_ONLY_UHID];
//   Logger.log("Reset emailSentMap for UHID " + TEST_ONLY_UHID + " to allow resending.");
// }
// //same thing as above but testing for ryan 













    for (var r = 1; r < rows.length; r++) {
      var uhid  = String(rows[r][iUhid]  || "").trim();
      var email = String(rows[r][iEmail] || "").trim();
      var firstName = String(rows[r][iFirst] || "").trim();
      var status = String(rows[r][iStatus] || "").trim();

      if (!uhid || !email) continue;
      var isPaidNow = PAID_STATUSES.test(status);
      var alreadySent = !!sentMem[uhid];

      // Log whether email would be sent
      if (isPaidNow && !alreadySent) {
        Logger.log("[TEST] Would send email to " + firstName + " (UHID: " + uhid + ") at " + email);
      }








// //comment this out later for wen ready for prod 
//       // Send email ONLY to test UHID (janred) and ryan 
// if (uhid !== TEST_ONLY_UHID && uhid !== "2266573") continue;











      if (isPaidNow && !alreadySent) {
        sendMembershipEmail(email, firstName);
        sentMem[uhid] = true;
        Logger.log("Welcome email sent to " + email + " (UHID " + uhid + ")");
      }
    }

    props.setProperty("emailSentMap", JSON.stringify(sentMem));
    Logger.log("checkPaidStatusAndSendEmails() complete");
  } catch (err) {
    Logger.log("Poller error: " + err);
  }
}


/** Sends the welcome email when someone becomes a paid member. */
function sendMembershipEmail(email, firstName) {
  email = "jred8069@gmail.com"; // TEST-ONLY: force all sends to me. REMOVE LATER.

  var subject = "Welcome to Membership 🎉 (+50 pts)";

  var htmlBody =
    "<p>Hi " + (firstName || "there") + ",</p>" +
    "<p>Thank you for becoming a paid member of SASE! We are so excited to have you join our community here at UH. By paying for membership, you just received <strong>+50 member points</strong>.</p>" +

    "<p style='font-weight:bold; font-size:16px;'>What are member points?</p>" +
    "<p>As a paid member, you’ll earn points for participating in different events throughout the year. Each event type is worth a different amount of points:</p>" +
    "<ul>" +
    "<li>Becoming a paid member: 50 points</li>" +
    "<li>General meetings: 50 points</li>" +
    "<li>Socials: 30 points</li>" +
    "<li>Professional Development events: 40 points</li>" +
    "<li>CFC events: 40 points</li>" +
    "<li>Volunteering events: 30 points per hour</li>" +
    "<li>SASE National Conference: 100 points</li>" +
    "</ul>" +
    "<p>You can save up your points and redeem them during our Auction Night for amazing prizes!</p>" +

    "<p style='font-weight:bold; font-size:16px;'>Get Involved with Our SASEFam!</p>" +
    "<ul>" +
    "<li><a href='https://docs.google.com/forms/d/156rciwDqCCRm3IWbarzCIod30wIC8c3t2ASdfTnK1s4/viewform?edit_requested=true'>Join a SASE Family! 🧑‍🧑‍🧒‍🧒</a> - Fill out this form to be assigned to a mentor and a family with similar interests to you.</li>" +
    "<li><a href='https://docs.google.com/forms/d/e/1FAIpQLSeqBr9aIj5gN68wzrb_qrtBb7TIsYXxpD9q0-sMCBIFbvE9wQ/viewform'>Join our media team! 📸</a> - Love to be behind the camera? Canva pro? Apply to be a part of our media team to help capture the moments we create at SASE. Apps due 9/20/25!</li>" +
    "<li><a href='https://docs.google.com/forms/d/e/1FAIpQLSdGcU2crNoKmq0XsP6i_G2gM8g2yehUMhGfj4LKW2cZlV7ByA/viewform'>Get sponsored for National Conference! ✈️</a> - Apply to be sponsored for SASE's National Conference in Pittsburgh, PA. You'll have an opportunity to get internships and network with recruiters! Requirements to apply are on the form. Apps due 9/12/25.</li>" +
    "<li><a href='https://linktr.ee/saseuh?fbclid=PAZXh0bgNhZW0CMTEAAaceI5FdYrNwCBI1L2EZoiLC9xSKIkyhIcrJP3FwGK7k6hPAo_OixAjOFPrgdw_aem_9Bersx7ArdUcUF1lGQlbJw'>Explore all our links! 🌐</a> – Check out our Linktree to find all our forms, resources, and socials in one place.</li>" +
    // "</ul>" +

    // "<p><strong>Connect with us!</strong></p>" +
    // "<p>" +
    // "<a href='https://www.instagram.com/uhsase/' style='margin-right:10px'>" +
    //   "<img src='https://cdn-icons-png.flaticon.com/32/2111/2111463.png' alt='Instagram' width='32' height='32'>" +
    // "</a>" +
    // "<a href='https://discord.com/invite/33A7B4f' style='margin-right:10px'>" +
    //   "<img src='https://cdn-icons-png.flaticon.com/32/2111/2111370.png' alt='Discord' width='32' height='32'>" +
    // "</a>" +
    // "<a href='https://groupme.com/join_group/108462129/qcNr5HlU'>" +
    //   "<img src='https://cdn-icons-png.flaticon.com/512/1384/1384056.png' alt='GroupMe' width='32' height='32'>" +
    // "</a>" +
    "</p>";

  MailApp.sendEmail({
    to: email,
    subject: subject,
    htmlBody: htmlBody
  });
}







