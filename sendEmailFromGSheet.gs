// This constant is written in column C for rows for which an email
// has been sent successfully.
var EMAIL_SENT = 'EMAIL_SENT';


/**
 * Sends non-duplicate emails with data from the current spreadsheet.
 */
function sendEmails() {

  var startRow = 2; // First row of data to process
  var numRows = 1000; // Number of rows to process
  var spreadsheet = SpreadsheetApp.getActiveSheet();
//  fetch data for calendar id
//  var calendarID = spreadsheet.getRange("G3").getValue();
//  var eventCal = CalendarApp.getCalendarById(calendarID);
//  var startTime = spreadsheet.getRange("E3").getValue();
//  var endTime = spreadsheet.getRange("F3").getValue();
  // Fetch the range of cells A:D
  var dataRange = spreadsheet.getRange(startRow, 1, numRows, 4);
  
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  for (var i = 0; i < data.length; ++i) {
    var row = data[i];
    
    var registerTime = row[0];
    var emailAddress = row[1]; // First column
    var message = "<h2>International Yoga Day 21st June </h2><br/><h3 style='color:grey;'>When? <br/> Sun 21 Jun 2020 5pm – 5:45pm India Standard Time</h3><br/><h3 style='color:grey;'>Joining info:<br/> Join with Google Meet : https://meet.google.com/kjt-aopi-roh</h3>";
    var emailSent = row[3]; // Fourth column
    
    if (emailSent !== EMAIL_SENT && emailAddress!="") { // Prevents sending duplicates
      var subject = 'International Yoga Day 21st June';
      MailApp.sendEmail(emailAddress, subject, message, {htmlBody:message});
      spreadsheet.getRange(startRow + i, 4).setValue(EMAIL_SENT);
      // Make sure the cell is updated right away in case the script is interrupted
      SpreadsheetApp.flush();
    }
  }

}
/**
 * The event handler triggered when opening the spreadsheet.
 * @param {Event} e The onOpen event.
 */
function onOpen() {
  // Add a custom menu to the spreadsheet.
  var ui = SpreadsheetApp.getUi(); // Or DocumentApp, SlidesApp, or FormApp.
  ui.createMenu('Custom Menu')
      .addItem('Send Email', 'sendEmails')
      .addToUi();
}

