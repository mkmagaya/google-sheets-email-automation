
// 1r43c2ndiTTYV55sk  -- spreadsheet ID

function sendDailyReport() {
  try {
    // Access the spreadsheet by ID
    var spreadsheetId = "1r43c2i5NOU4zEIWvR_LYwxYTm0Ysk"; // Replace with your actual Spreadsheet ID
    var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    var sheet = spreadsheet.getActiveSheet();

    // Retrieve the data range
    var range = sheet.getDataRange();
    var data = range.getValues();
    
    // Log the data for debugging
    Logger.log(data);

    // Check if data is retrieved properly
    if (data.length === 0) {
      throw new Error("No data found in the spreadsheet.");
    }

    // Format the data as HTML
    // var dataAsHtml = formatAsHtmlTable(data);

    // Convert the spreadsheet to Blob for attachment
    var spreadsheetBlob = spreadsheet.getBlob().setName(spreadsheet.getName() + ".xlsx");

    // Send the email
    MailApp.sendEmail({
      to: "r@gmail.com",
      cc: "makomagaya05@gmail.com, bte@gmail.com, ze@gmail.com",
      subject: "Rublemak Solutions Daily Report",
      htmlBody: "Good day Rublemak Team <br> Please find the attached Task Schedule documents:<br><br>",// + dataAsHtml,
      attachments: [
        spreadsheet.getAs(MimeType.PDF).setName("Task_Schedule.pdf"),
        spreadsheetBlob
      ]
    });
  } catch (e) {
    Logger.log("Error: " + e.message);
  }
}

// function formatAsHtmlTable(data) {
//   var html = '<table border="1" cellspacing="0" cellpadding="5">';
//   for (var row = 0; row < data.length; row++) {
//     html += '<tr>';
//     for (var col = 0; col < data[row].length; col++) {
//       html += '<td>' + data[row][col] + '</td>';
//     }
//     html += '</tr>';
//   }
//   html += '</table>';
//   return html;
// }
