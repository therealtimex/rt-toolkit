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
 * 5. Configure the WEBHOOK_URL parameter below (required)
 * 6. Optionally configure SHEET_NAME and COLUMN_TO_WATCH
 * 7. Click Save (disk icon or Ctrl/Cmd + S)
 * 8. Set up triggers:
 *    - Click on "Triggers" (clock icon) in the left sidebar
 *    - Click "+ Add Trigger" button
 *    - Choose function: triggerGoogleSheetWebhook
 *    - Choose event source: From spreadsheet
 *    - Select event type: On edit
 *    - Add another trigger for "On form submit" if using Google Forms
 * 9. Authorize the script when prompted
 * 10. Test by editing a cell in your watched column and by inserting a new row
 * 
 * Troubleshooting:
 * - View execution logs: View > Execution log
 * - Check trigger runs: View > Execution log history
 * - Verify webhook URL is accessible and accepts POST requests
 */

// User-configurable parameters
/**
 * Enhanced Google Sheets Webhook Trigger
 * 
 * Author: Trung Le at RealTimeX.co
 * Created: 2024-12-14
 * Last Modified: 2024-12-14
 * 
 * Purpose: Triggers a webhook when a specific column is edited or a new row is inserted.
 * Sends the updated row data to a specified webhook URL with improved security, reliability, and performance.
 */

// User-configurable parameters
const WEBHOOK_URL = "https://your-webhook-url.com";  // Your webhook URL (required)
const WEBHOOK_SECRET = "your_secret_token";  // Secret token for webhook authentication
const SHEET_NAME = "";  // Name of the sheet to watch (leave empty for first sheet)
const COLUMN_TO_WATCH = 0;  // Column number to trigger the webhook (0 for last column)

// Constants
const MAX_RETRIES = 3;
const RETRY_DELAY = 1000; // milliseconds
const QUEUE_PROCESS_INTERVAL = 5000; // milliseconds

// Global variables
let webhookQueue = [];
let isProcessingQueue = false;

function triggerGoogleSheetWebhook(e) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = SHEET_NAME ? spreadsheet.getSheetByName(SHEET_NAME) : spreadsheet.getSheets()[0];
    const headings = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const columnToWatch = COLUMN_TO_WATCH || sheet.getLastColumn();
    
    let row, triggerWebhook = false;

    if (e.changeType === "INSERT_ROW") {
      row = e.source.getActiveRange().getRow();
      triggerWebhook = true;
    } else if (e.changeType === "EDIT") {
      const activeRange = sheet.getActiveRange();
      row = activeRange.getRow();
      const column = activeRange.getColumn();
      
      if (column === columnToWatch) {
        triggerWebhook = true;
      }
    }

    if (!triggerWebhook) return;

    const values = sheet.getRange(row, 1, 1, headings.length).getValues()[0];
    
    const payload = {
      row_number: row,
      timestamp: new Date().toISOString(),
      change_type: e.changeType,
      ...Object.fromEntries(headings.map((name, i) => [name, values[i]]))
    };

    addToWebhookQueue(payload);
    processWebhookQueue();

  } catch (error) {
    console.error(`Error in triggerGoogleSheetWebhook: ${error.toString()}`);
    logError(error);
  }
}

function addToWebhookQueue(payload) {
  webhookQueue.push(payload);
  console.log(`Added to queue: ${JSON.stringify(payload)}`);
}

function processWebhookQueue() {
  if (isProcessingQueue || webhookQueue.length === 0) return;

  isProcessingQueue = true;
  processNextWebhook();
}

function processNextWebhook() {
  if (webhookQueue.length === 0) {
    isProcessingQueue = false;
    return;
  }

  const payload = webhookQueue.shift();
  processWebhookWithRetry(payload)
    .then(() => {
      console.log('Webhook processed successfully');
      Utilities.sleep(QUEUE_PROCESS_INTERVAL);
      processNextWebhook();
    })
    .catch(error => {
      console.error('Failed to process webhook after retries');
      logError(error);
      processNextWebhook();
    });
}

async function processWebhookWithRetry(payload, attempt = 1) {
  try {
    const signature = computeHmacSignature(payload);
    const response = await UrlFetchApp.fetch(WEBHOOK_URL, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      headers: {
        'X-Webhook-Signature': signature
      },
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() !== 200) {
      throw new Error(`HTTP error! status: ${response.getResponseCode()}`);
    }
    
    logWebhookSuccess(payload, response);
    return response;
  } catch (error) {
    console.error(`Attempt ${attempt} failed: ${error.toString()}`);
    if (attempt < MAX_RETRIES) {
      await Utilities.sleep(RETRY_DELAY * attempt);
      return processWebhookWithRetry(payload, attempt + 1);
    }
    throw error;
  }
}

function computeHmacSignature(payload) {
  const payloadString = JSON.stringify(payload);
  const signature = Utilities.computeHmacSha256Signature(payloadString, WEBHOOK_SECRET);
  return signature.map(byte => ('0' + (byte & 0xFF).toString(16)).slice(-2)).join('');
}

function logWebhookSuccess(payload, response) {
  console.log(`Webhook sent successfully: ${JSON.stringify({
    payload: payload,
    responseCode: response.getResponseCode(),
    responseBody: response.getContentText()
  })}`);
}

function logError(error) {
  console.error(`Error details: ${JSON.stringify({
    message: error.message,
    stack: error.stack,
    timestamp: new Date().toISOString()
  })}`);
}
