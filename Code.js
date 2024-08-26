const LABEL_NAME = 'YourLabelName';
const SPREADSHEET_ID = 'YourSpreadsheetID';
const SHEET_NAME = 'Sheet1';

/**
 * Processes emails with a specific label and writes their information to a spreadsheet.
 */
function onLabelAdded() {
  const label = GmailApp.getUserLabelByName(LABEL_NAME);
  const threads = label.getThreads();
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = spreadsheet.getSheetByName(SHEET_NAME);

  // If the sheet is empty, add headers
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Timestamp', 'Message ID', 'From', 'Subject', 'Date']);
  }

  Logger.log(`Processing ${threads.length} threads`);

  let newRowsCount = 0;

  threads.forEach(thread => {
    const messages = thread.getMessages();
    messages.forEach(message => {
      if (shouldProcessMessage(message, sheet)) {
        const rowData = extractMessageData(message);
        sheet.appendRow(rowData);
        message.star();
        newRowsCount++;
        Logger.log(`Added new row for message: ${message.getId()}`);
      }
    });
  });

  Logger.log(`Added ${newRowsCount} new rows to the spreadsheet`);
}

/**
 * Checks if a message should be processed and added to the spreadsheet.
 * @param {GoogleAppsScript.Gmail.GmailMessage} message - The message to check.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to check against.
 * @return {boolean} True if the message should be processed, false otherwise.
 */
function shouldProcessMessage(message, sheet) {
  const messageId = message.getId();
  const data = sheet.getDataRange().getValues();
  const isAlreadyProcessed = data.some(row => row[1] === messageId);

  if (isAlreadyProcessed) {
    Logger.log(`Message ${messageId} already processed, skipping`);
    return false;
  }

  Logger.log(`Message ${messageId} is new, will be processed`);
  return true;
}

/**
 * Extracts relevant data from a message.
 * @param {GoogleAppsScript.Gmail.GmailMessage} message - The message to extract data from.
 * @return {Array} An array containing the extracted data.
 */
function extractMessageData(message) {
  return [
    new Date(), // Current timestamp
    message.getId(),
    message.getFrom(),
    message.getSubject(),
    message.getDate()
  ];
}

/**
 * Creates a time-driven trigger to run the script periodically.
 */
function createTimeDrivenTrigger() {
  // Delete existing triggers
  ScriptApp.getProjectTriggers().forEach(trigger => ScriptApp.deleteTrigger(trigger));
  
  // Create a new trigger
  ScriptApp.newTrigger('onLabelAdded')
    .timeBased()
    .everyMinutes(5)
    .create();
  
  Logger.log('New trigger created to run every 5 minutes');
}

