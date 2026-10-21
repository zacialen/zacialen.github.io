// Sheet Configuration
const SHEET_NAME = 'Sheet1'; // Change this to your sheet name if different
const TEXT_COLUMN = 'A'; // Column containing Ibanag texts
const STATUS_COLUMN = 'B'; // Column containing status (Not Yet, Recorded, Done)
const DATE_COLUMN = 'C'; // Column for date stamp when recorded
const ID_COLUMN = 'D'; // Optional: Column for unique ID

// Fetch 10 random texts with "Not Yet" status
function doGet(e) {
  // Log for debugging
  Logger.log('GET request received. Parameters: ' + JSON.stringify(e.parameter));

  const action = e.parameter.action;
  const callback = e.parameter.callback;

  let jsonData;

  if (action === 'getTexts') {
    const result = getRandomTexts();
    jsonData = result.getContent();
  } else if (action === 'submitRecording') {
    const result = submitRecording(e);
    jsonData = result.getContent();
  } else {
    jsonData = JSON.stringify({error: 'Invalid action', parameters: e.parameter});
  }

  // If callback is provided, return JSONP (for CORS workaround)
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + jsonData + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService.createTextOutput(jsonData)
    .setMimeType(ContentService.MimeType.JSON);
}

// Handle POST requests for submission
function doPost(e) {
  // Log for debugging
  Logger.log('POST request received. Parameters: ' + JSON.stringify(e.parameter));
  Logger.log('POST data: ' + e.postData.contents);

  const action = e.parameter.action;

  if (action === 'submitRecording') {
    return submitRecording(e);
  }

  return ContentService.createTextOutput(JSON.stringify({error: 'Invalid action', parameters: e.parameter}))
    .setMimeType(ContentService.MimeType.JSON);
}

function getRandomTexts() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    const data = sheet.getDataRange().getValues();

    Logger.log('Sheet data rows: ' + data.length);

    // Skip header row if exists
    const startIndex = data[0][0] === 'Text' || data[0][0] === 'Ibanag Text' ? 1 : 0;
    const texts = [];

    // Collect texts with "Not Yet" status
    for (let i = startIndex; i < data.length; i++) {
      const row = data[i];
      const text = row[0]; // Column A
      const status = row[1] ? row[1].toString().trim() : ''; // Column B

      Logger.log('Row ' + i + ': Text=' + text + ', Status=' + status);

      // Only include texts with "Not Yet" status
      if (status === 'Not Yet' || status === '') {
        texts.push({
          text: text,
          row: i + 1, // 1-based row number
          id: row[3] || `text_${i}` // Use ID column (D) or generate one
        });
      }
    }

    Logger.log('Found ' + texts.length + ' available texts');

    // Randomly select 10 texts
    const selectedTexts = [];
    const shuffled = texts.sort(() => 0.5 - Math.random());
    const count = Math.min(10, shuffled.length);

    for (let i = 0; i < count; i++) {
      selectedTexts.push(shuffled[i]);
    }

    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      texts: selectedTexts
    })).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    Logger.log('Error in getRandomTexts: ' + error.toString());
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function submitRecording(e) {
  try {
    const textIds = e.parameter.textIds; // Comma-separated list of text IDs
    const variant = e.parameter.variant;
    const recordingType = e.parameter.recordingType;
    const topic = e.parameter.topic || '';

    Logger.log('Submitting recording. Text IDs: ' + textIds);

    // Update status to "Recorded" and add date stamp
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    const ids = textIds.split(',');
    const currentDate = new Date();
    const formattedDate = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

    ids.forEach(id => {
      // Find the row by ID (search column D)
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) { // Skip header
        if (data[i][3] == id) { // Column D (ID)
          sheet.getRange(i + 1, 2).setValue('Recorded'); // Update status column (B)
          sheet.getRange(i + 1, 3).setValue(formattedDate); // Update date column (C)
          Logger.log('Updated row ' + (i + 1) + ' to Recorded');
          break;
        }
      }
    });

    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      message: 'Recording submitted successfully'
    })).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    Logger.log('Error in submitRecording: ' + error.toString());
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// Helper function to manually update status (you can run this manually)
function updateStatus(textId, newStatus) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][3] == textId) { // Column D (ID)
      sheet.getRange(i + 1, 2).setValue(newStatus); // Update status column (B)
      return `Updated text ${textId} to ${newStatus}`;
    }
  }

  return `Text ${textId} not found`;
}


