# Google Sheets Email Automation

This project automates the process of sending Google Sheets data via email every Monday at 8:30 AM Central Africa Time (CAT). The email includes the spreadsheet data formatted as an HTML table, a PDF attachment of the spreadsheet, and the original spreadsheet file.

## Features

- Automatically retrieves data from a specified Google Sheets document.
- Formats the data as an HTML table for the email body.
- Attaches the spreadsheet as both PDF and original format to the email.
- Sends the email to a specified recipient and CCs multiple recipients.
- Schedules the script to run every Monday at 8:30 AM CAT using Google Apps Script.

## Setup

### Prerequisites

- A Google account with access to Google Sheets.
- A GitHub account.

### Step-by-Step Guide

1. **Open Google Sheets Document:**
   - Open the Google Sheets document you want to automate.

2. **Access Google Apps Script:**
   - Navigate to `Extensions` -> `Apps Script` to open the Script Editor.

3. **Write the Email Script:**
   - Copy and paste the following script into the Script Editor:

    ```javascript
    function sendDailyReport() {
      try {
        // Access the spreadsheet by ID
        var spreadsheetId = "YOUR_SPREADSHEET_ID"; // Replace with your actual Spreadsheet ID
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
        var dataAsHtml = formatAsHtmlTable(data);

        // Convert the spreadsheet to Blob for attachment
        var spreadsheetBlob = spreadsheet.getBlob().setName(spreadsheet.getName() + ".xlsx");

        // Send the email
        MailApp.sendEmail({
          to: "ruble@gmail.com",
          cc: "makomagaya05@gmail.com, bt@gmail.com, r@gmail.com",
          subject: "Rublemak Solutions Daily Report",
          htmlBody: "Please find the Task Schedule document:<br><br>" + dataAsHtml,
          attachments: [
            spreadsheet.getAs(MimeType.PDF).setName("Task_Schedule.pdf"),
            spreadsheetBlob
          ]
        });
      } catch (e) {
        Logger.log("Error: " + e.message);
      }
    }

    function formatAsHtmlTable(data) {
      var html = '<table border="1" cellspacing="0" cellpadding="5">';
      for (var row = 0; row < data.length; row++) {
        html += '<tr>';
        for (var col = 0; col < data[row].length; col++) {
          html += '<td>' + data[row][col] + '</td>';
        }
        html += '</tr>';
      }
      html += '</table>';
      return html;
    }
    ```

4. **Create Time-driven Trigger:**
   - Add the following function to set a weekly trigger:

    ```javascript
    function createTimeDrivenTrigger() {
      // Convert 8:30 AM CAT to UTC (CAT is UTC+2)
      var triggerTime = new Date();
      triggerTime.setUTCHours(6); // 8:30 AM CAT is 6:30 AM UTC
      triggerTime.setUTCMinutes(30);

      ScriptApp.newTrigger("sendDailyReport")
        .timeBased()
        .onWeekDay(ScriptApp.WeekDay.MONDAY)
        .at(triggerTime)
        .create();
    }
    ```

5. **Run the Trigger Setup Script:**
   - Select the `createTimeDrivenTrigger` function from the dropdown menu in the toolbar.
   - Click the run button (triangle icon) to create the weekly trigger.

6. **Authorize the Script:**
   - Follow the prompts to grant the necessary permissions the first time you run the script.

## Running the Script Manually

You can manually run the `sendDailyReport` function to test it and ensure that the email is sent correctly:

1. Select the `sendDailyReport` function from the dropdown menu in the toolbar.
2. Click the run button (triangle icon) to execute the function.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.
"# google-sheets-email-automation" 
