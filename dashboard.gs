function createOrUpdateDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1) Create or clear the "Dashboard" sheet
  let dashboardSheet = ss.getSheetByName("Dashboard");
  if (dashboardSheet) {
    // Clear existing content if it already exists
    dashboardSheet.clear();
  } else {
    // Otherwise, create a new sheet named "Dashboard"
    dashboardSheet = ss.insertSheet("Dashboard");
  }
  
  // 2) Put some headers/labels in the Dashboard
  dashboardSheet.getRange("A1").setValue("Dashboard");
  dashboardSheet.getRange("A2").setValue("Total members");
  dashboardSheet.getRange("A3").setValue("Events held");
  dashboardSheet.getRange("A4").setValue("Average attendance");
  dashboardSheet.getRange("A5").setValue("Retention rate");
  
  // 3) Pull data from "Member Overview" to count total members
  let totalMembers = 0;
  const memberSheet = ss.getSheetByName("Member Overview");
  if (memberSheet) {
    const memberData = memberSheet.getDataRange().getValues();
    // Subtract 1 for the header row
    totalMembers = memberData.length > 1 ? memberData.length - 1 : 0;
  }
  
  // 4) Pull data from "EventMultipliers" to count events and sum attendance
  let eventsHeld = 0;
  let totalAttendance = 0;
  const multipliersSheet = ss.getSheetByName("EventMultipliers");
  if (multipliersSheet) {
    const multiData = multipliersSheet.getDataRange().getValues();
    // The first row is a header, so actual events = multiData.length - 1
    eventsHeld = multiData.length > 1 ? multiData.length - 1 : 0;
    
    // Assume "Attendees" is in column C (index 2)
    for (let i = 1; i < multiData.length; i++) {
      const attendees = parseInt(multiData[i][2], 10) || 0;
      totalAttendance += attendees;
    }
  }
  
  // 5) Calculate average attendance
  let avgAttendance = 0;
  if (eventsHeld > 0) {
    avgAttendance = totalAttendance / eventsHeld;
  }
  
  // 6) Example “retention rate”
  //    You can define your own logic, e.g. (avgAttendance / totalMembers) * 100
  let retentionRate = 0;
  if (totalMembers > 0) {
    retentionRate = (avgAttendance / totalMembers) * 100;
  }
  
  // 7) Write these values into the Dashboard sheet
  dashboardSheet.getRange("B2").setValue(totalMembers);
  dashboardSheet.getRange("B3").setValue(eventsHeld);
  dashboardSheet.getRange("B4").setValue(avgAttendance.toFixed(2));
  dashboardSheet.getRange("B5").setValue(retentionRate.toFixed(2) + "%");
  
  // Optional: Format or style the sheet as you wish
  dashboardSheet.autoResizeColumns(1, 2);
}

