/**
 * OUR WALK TRACKER - BACKEND
 * Optimized for Sunday-reset weekly goals, dynamic monthly goals,
 * and robust history retrieval.
 */

function doGet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const now = new Date();
  
  // 1. DYNAMIC GOALS
  const daysInMonth = new Date(now.getFullYear(), now.getMonth() + 1, 0).getDate();
  const monthGoal = daysInMonth * 20;
  const weekGoal = 150;

  // 2. TIME BOUNDARIES
  const startOfMonth = new Date(now.getFullYear(), now.getMonth(), 1);
  const startOfWeek = new Date(now);
  startOfWeek.setDate(now.getDate() - now.getDay()); // Sunday Reset
  startOfWeek.setHours(0,0,0,0);

  let stats = { 
    currentWeek: 0, 
    currentMonth: 0, 
    monthGoal: monthGoal, 
    weekGoal: weekGoal,
    recent: [] 
  };

  // 3. CALCULATE TOTALS & PROCESS HISTORY
  // We'll filter out the header and any empty rows first
  const walkRows = data.slice(1).filter(row => row[0] !== "" && row[2] !== "");

  walkRows.forEach(row => {
    let rowDate = row[0];
    
    // Force conversion to Date object if it's a string or number
    if (!(rowDate instanceof Date)) {
      rowDate = new Date(rowDate);
    }

    // Only process if the date is valid
    if (!isNaN(rowDate.getTime())) {
      const mins = Number(row[2]) || 0;

      if (rowDate >= startOfWeek) stats.currentWeek += mins;
      if (rowDate >= startOfMonth) stats.currentMonth += mins;
    }
  });

  // 4. GENERATE RECENT HISTORY (Last 5)
  const lastFive = walkRows.slice(-5).reverse(); 

  lastFive.forEach(row => {
    let d = row[0];
    if (!(d instanceof Date)) d = new Date(d);

    if (!isNaN(d.getTime())) {
      stats.recent.push({
        date: Utilities.formatDate(d, Session.getScriptTimeZone(), "MMM d"),
        mins: row[2]
      });
    }
  });

  return ContentService.createTextOutput(JSON.stringify(stats))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const data = JSON.parse(e.postData.contents);
    
    // Append Data: [Timestamp, User, Duration, DateString]
    sheet.appendRow([
      new Date(), 
      "Couple", 
      Number(data.duration), 
      data.date
    ]);
    
    return ContentService.createTextOutput(JSON.stringify({"result":"success"}))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (err) {
    console.error('Post Error: ' + err.toString());
    return ContentService.createTextOutput(JSON.stringify({"result":"error", "message": err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
