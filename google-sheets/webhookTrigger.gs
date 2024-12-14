/**
 * Google Sheets Webhook Trigger
 * 
 * Author: Trung Le at RealTimeX.co
 * Created: 2024-12-14
 * Last Modified: 2024-12-14
 * 
 * Purpose: This script triggers a webhook when specific changes are made to a Google Sheet.
 * It sends the updated row data to a specified webhook URL.
 * 
 * Instructions:
 * 1. Set the user-configurable parameters below.
 * 2. Set up a trigger for this function on edit and when a row is inserted.
 * 
 * Credit: Eyal Gershon & PPLX
 */

// User-configurable parameters
const SHEET_NAME = "YourSheetName";  // Name of the sheet to watch
const COLUMN_TO_WATCH = 3;           // Column number to trigger the webhook (e.g., 3 for column C)
const WEBHOOK_URL = "https://your-webhook-url.com";  // Your webhook URL

function triggerGoogleSheetWebhook(e) {
  try {
    // Exit if the change type is not EDIT or INSERT_ROW
    if (e.changeType !== "EDIT" && e.changeType !== "INSERT_ROW") return;
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(SHEET_NAME);
    const headings = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    const activeRange = sheet.getActiveRange();
    const row = activeRange.getRow();
    const column = activeRange.getColumn();
    
    // Exit if it's an EDIT event and the changed column is not the one we're watching
    if (e.changeType === "EDIT" && column !== COLUMN_TO_WATCH) return;
    
    // Get the values of the entire row
    const values = sheet.getRange(row, 1, 1, headings.length).getValues()[0];
    
    // Prepare the payload
    const payload = {
      row_number: row,
      timestamp: new Date().toISOString(),
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
      console.log('Webhook sent successfully');
    }
  } catch (error) {
    console.error(`Error in triggerGoogleSheetWebhook: ${error.toString()}`);
  }
}
