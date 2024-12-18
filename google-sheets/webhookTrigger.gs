/**
 * Google Sheets Webhook Trigger
 * ============================
 * 
 * Author: Trung Le at RealTimeX.co
 * Created: 2024-12-14
 * Last Modified: 2024-12-17
 * Version: 1.1.0
 * 
 * Description:
 * ------------
 * This script monitors specific column changes in a Google Sheet and sends the updated
 * row data to a webhook endpoint. It handles both form submissions and manual edits,
 * with configurable column watching and robust error handling.
 * 
 * Features:
 * ---------
 * - Monitors specific column changes for manual edits
 * - Captures all form submissions automatically
 * - Queued webhook processing with retries
 * - HMAC signature for webhook security
 * - Comprehensive error logging
 * - Configurable sheet and column targeting
 * 
 * Configuration:
 * -------------
 * WEBHOOK_URL: Your webhook endpoint URL (required)
 * WEBHOOK_SECRET: Secret key for HMAC signature (required)
 * SHEET_NAME: Target sheet name (optional, defaults to first sheet)
 * COLUMN_TO_WATCH: Column number to monitor for manual edits (required)
 * 
 * Installation Instructions:
 * ------------------------
 * 1. Open your Google Sheet
 * 2. Click Extensions > Apps Script
 * 3. Copy this entire script into the editor
 * 4. Configure the constants (WEBHOOK_URL, etc.)
 * 5. Save the script
 * 
 * Trigger Setup:
 * -------------
 * Two triggers are required:
 * 
 * 1. Edit Trigger:
 *    - Click "Triggers" in the left sidebar
 *    - Click "+ Add Trigger"
 *    - Choose function: installableOnEdit
 *    - Event source: From spreadsheet
 *    - Event type: On edit
 * 
 * 2. Form Submit Trigger:
 *    - Click "+ Add Trigger" again
 *    - Choose function: onFormSubmit
 *    - Event source: From spreadsheet
 *    - Event type: On form submit
 * 
 * Webhook Payload Format:
 * ---------------------
 * For Edit Events:
 * {
 *   "row_number": number,
 *   "timestamp": string (ISO format),
 *   "change_type": "EDIT",
 *   "edited_column": string,
 *   "old_value": string,
 *   "new_value": string,
 *   ...row_data
 * }
 * 
 * For Form Submit Events:
 * {
 *   "row_number": number,
 *   "timestamp": string (ISO format),
 *   "change_type": "FORM_SUBMIT",
 *   ...form_response_data
 * }
 * 
 * Security:
 * --------
 * - Uses HMAC-SHA256 signatures
 * - Includes webhook secret verification
 * - Runs with authorized user permissions
 * 
 * Error Handling:
 * --------------
 * - Automatic retry mechanism (3 attempts)
 * - Exponential backoff between retries
 * - Detailed error logging
 * - Queue-based processing
 * 
 * Limitations:
 * -----------
 * - Requires UrlFetchApp permissions
 * - Maximum execution time of 6 minutes
 * - Maximum payload size determined by webhook endpoint
 * - Form submissions must be linked to the sheet
 * 
 * Dependencies:
 * ------------
 * - Google Apps Script
 * - UrlFetchApp service
 * - Properties service (for configuration)
 * - Google Forms (for form submission feature)
 * 
 * Support:
 * -------
 * For issues or feature requests, contact:
 * Email: support@realtimex.co
 * 
 * License:
 * -------
 * Copyright (c) 2024 RealTimeX.co
 * All rights reserved.
 */

// User-configurable parameters
const WEBHOOK_URL = "https://your-webhook-url.com";  // Your webhook URL (required)
const WEBHOOK_SECRET = "your_secret_token";  // Secret token for webhook authentication (required)
const SHEET_NAME = "";  // Name of the sheet to watch (leave empty for first sheet)
const COLUMN_TO_WATCH = 0;  // Column number to trigger the webhook (0 for last column)

// Constants
const MAX_RETRIES = 3;
const RETRY_DELAY = 1000;
const QUEUE_PROCESS_INTERVAL = 5000;

// Global variables
let webhookQueue = [];
let isProcessingQueue = false;

function onFormSubmit(e) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = SHEET_NAME ? spreadsheet.getSheetByName(SHEET_NAME) : spreadsheet.getSheets()[0];
    const headings = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Get the actual row number from form response
    const namedValues = e.namedValues;
    const range = e.range;
    const row = range.getRow();
    
    const payload = {
      row_number: row,
      timestamp: new Date().toISOString(),
      change_type: "FORM_SUBMIT",
      ...Object.fromEntries(headings.map((name, i) => [
        name, 
        namedValues[name] ? namedValues[name][0] : ''
      ]))
    };

    addToWebhookQueue(payload);
    processWebhookQueue();

  } catch (error) {
    console.error(`Error in onFormSubmit: ${error.toString()}`);
    logError(error);
  }
}

// Rename onEdit to installableOnEdit to avoid simple triggers
function installableOnEdit(e) {
  // Check if this is a simple trigger (will not have auth)
  if (!e.authMode || e.authMode === ScriptApp.AuthMode.NONE) {
    console.log('Skipping simple trigger execution');
    return;
  }

  try {
    // Skip if no active range or if it's the header row
    if (!e.range || e.range.getRow() === 1) return;

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = SHEET_NAME ? spreadsheet.getSheetByName(SHEET_NAME) : spreadsheet.getSheets()[0];
    
    // Skip if not in the target sheet
    if (SHEET_NAME && e.source.getActiveSheet().getName() !== SHEET_NAME) return;
    
    const headings = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const columnToWatch = COLUMN_TO_WATCH || sheet.getLastColumn();
    
    const activeRange = e.range;
    const row = activeRange.getRow();
    const column = activeRange.getColumn();
    
    // Only trigger if the edited column matches the column to watch
    if (column === columnToWatch) {
      const values = sheet.getRange(row, 1, 1, headings.length).getValues()[0];
      
      const payload = {
        row_number: row,
        timestamp: new Date().toISOString(),
        change_type: "EDIT",
        edited_column: headings[column - 1],
        old_value: e.oldValue || '',
        new_value: e.value || '',
        ...Object.fromEntries(headings.map((name, i) => [name, values[i]]))
      };

      addToWebhookQueue(payload);
      processWebhookQueue();
    }

  } catch (error) {
    console.error(`Error in installableOnEdit: ${error.toString()}`);
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