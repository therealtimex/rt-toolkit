/**
 * Google Sheets Webhook Trigger
 * 
 * Author: Trung Le at RealTimeX.co
 * Created: 2024-12-14
 * Last Modified: 2024-12-14
 * 
 * Purpose: This script triggers a webhook when:
 * 1. A specific column is edited
 * 2. A new row is inserted
 * It sends the updated row data to a specified webhook URL.
 * 
 * Integration Instructions:
 * 1. Open your Google Sheet
 * 2. Click on Extensions > Apps Script
 * 3. Create a new file named 'webhookTrigger.gs'
 * 4. Copy and paste this entire script
 * 5. Configure the parameters below
 * 6. Click Save (disk icon or Ctrl/Cmd + S)
 * 7. Set up triggers:
 *    - Click on "Triggers" (clock icon) in the left sidebar
 *    - Click "+ Add Trigger" button
 *    - Choose function: triggerGoogleSheetWebhook
 *    - Choose event source: From spreadsheet
 *    - Select event type: On edit
 *    - Add another trigger for "On form submit" if using Google Forms
 * 8. Authorize the script when prompted
 * 9. Test by editing a cell in your watched column and by inserting a new row
 * 
 * Troubleshooting:
 * - View execution logs: View > Execution log
 * - Check trigger runs: View > Execution log history
 * - Verify webhook URL is accessible and accepts POST requests
 */

// User-configurable parameters
const SHEET_NAME = "YourSheetName";  // Name of the sheet to watch
const COLUMN_TO_WATCH = 3;           // Column number to trigger the webhook (e.g., 3 for column C)
const WEBHOOK_URL = "https://your-webhook-url.com";  // Your webhook URL

function triggerGoogleSheetWebhook(e) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(SHEET_NAME);
    const headings = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    let row, triggerWebhook = false;

    if (e.changeType === "INSERT_ROW") {
      row = e.source.getActiveRange().getRow();
      triggerWebhook = true;
    } else if (e.changeType === "EDIT") {
      const activeRange = sheet.getActiveRange();
      row = activeRange.getRow();
      const column = activeRange.getColumn();
      
      // Trigger webhook if the edited column is the one we're watching
      if (column === COLUMN_TO_WATCH) {
        triggerWebhook = true;
      }
    }

    if (!triggerWebhook) return;

    // Get the values of the entire row
    const values = sheet.getRange(row, 1, 1, headings.length).getValues()[0];
    
    // Prepare the payload
    const payload = {
      row_number: row,
      timestamp: new Date().toISOString(),
      change_type: e.changeType,
      ...Object.fromEntries(headings.map((name, i) => [name, values[i]]))
    };
    
    // Webhook configuration
    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    // Send the webhook
    const response = UrlFetchApp.fetch(WEBHOOK_URL, options);
    if (response.getResponseCode() !== 200) {
      console.error(`Failed to send webhook: ${response.getContentText()}`);
    } else {
      console.log(`Webhook sent successfully for ${e.changeType} event on row ${row}`);
    }
  } catch (error) {
    console.error(`Error in triggerGoogleSheetWebhook: ${error.toString()}`);
  }
}
