# Google Apps Script for Ibanag Speech Collection

## Setup Instructions

1. Open your Google Sheet
2. Create a second sheet tab named "Responses" (this will log all submissions)
3. Go to Extensions → Apps Script
4. Delete any existing code and paste the script below
5. Save the script (Ctrl+S or Cmd+S)
6. Click Deploy → New deployment
7. Select type: Web app
8. Description: "Ibanag Speech Collection API"
9. Execute as: Me (your email)
10. Who has access: Anyone
11. Click Deploy
12. Copy the Web App URL
13. Share the URL with me to integrate into the website

## Google Apps Script Code

```javascript
// Sheet Configuration
const TEXTS_SHEET_NAME = 'ref'; // Main sheet with Ibanag texts
const RESPONSES_SHEET_NAME = 'resp'; // Sheet to log submissions
const ID_COLUMN = 'A'; // Column for sentence ID
const TEXT_COLUMN = 'B'; // Column containing Ibanag texts
const STATUS_COLUMN = 'C'; // Column containing status (Not Yet, Recorded, Done)
const DATE_COLUMN = 'D'; // Column for date stamp when recorded

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

function getRandomTexts() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TEXTS_SHEET_NAME);
    const data = sheet.getDataRange().getValues();

    Logger.log('Sheet data rows: ' + data.length);

    // Skip header row if exists
    const startIndex = data[0][0] === 'Text' || data[0][0] === 'Ibanag Text' ? 1 : 0;
    const texts = [];

    // Collect texts with "Not Yet" status
    for (let i = startIndex; i < data.length; i++) {
      const row = data[i];
      const id = row[0]; // Column A - Sentence ID
      const text = row[1]; // Column B - Ibg Text
      const status = row[2] ? row[2].toString().trim() : ''; // Column C - Status

      Logger.log('Row ' + i + ': ID=' + id + ', Text=' + text + ', Status=' + status);

      // Only include texts with "Not Yet" status
      if (status === 'Not Yet' || status === '') {
        texts.push({
          text: text,
          row: i + 1, // 1-based row number
          id: id || `text_${i}` // Use ID column (A) or generate one
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

    const currentDate = new Date();
    const formattedDate = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

    // Update main sheet: mark texts as "Recorded" and add date stamp
    const textsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TEXTS_SHEET_NAME);
    const ids = textIds.split(',');

    ids.forEach(id => {
      // Find the row by ID (search column A)
      const data = textsSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) { // Skip header
        if (data[i][0] == id) { // Column A (ID)
          textsSheet.getRange(i + 1, 3).setValue('Recorded'); // Update status column (C)
          textsSheet.getRange(i + 1, 4).setValue(formattedDate); // Update date column (D)
          Logger.log('Updated row ' + (i + 1) + ' to Recorded');
          break;
        }
      }
    });

    // Log submission to Responses sheet
    logToResponsesSheet(textIds, variant, recordingType, topic, formattedDate);

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

function logToResponsesSheet(textIds, variant, recordingType, topic, timestamp) {
  try {
    const responsesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RESPONSES_SHEET_NAME);

    // Create sheet if it doesn't exist
    if (!responsesSheet) {
      responsesSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(RESPONSES_SHEET_NAME);
      // Add headers
      responsesSheet.appendRow(['Timestamp', 'Variant', 'Recording Type', 'Topic', 'Sentence IDs', 'Number of Sentences']);
    }

    // Check if headers exist, add if needed
    const headers = responsesSheet.getRange(1, 1, 1, 6).getValues()[0];
    if (headers[0] !== 'Timestamp') {
      responsesSheet.clear();
      responsesSheet.appendRow(['Timestamp', 'Variant', 'Recording Type', 'Topic', 'Sentence IDs', 'Number of Sentences']);
    }

    // Log the submission
    responsesSheet.appendRow([
      timestamp,
      variant,
      recordingType,
      topic,
      textIds,
      textIds.split(',').length
    ]);

    Logger.log('Logged submission to Responses sheet');
  } catch (error) {
    Logger.log('Error logging to Responses sheet: ' + error.toString());
  }
}

// Helper function to manually update status (you can run this manually)
function updateStatus(textId, newStatus) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TEXTS_SHEET_NAME);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == textId) { // Column A (ID)
      sheet.getRange(i + 1, 3).setValue(newStatus); // Update status column (C)
      return `Updated text ${textId} to ${newStatus}`;
    }
  }

  return `Text ${textId} not found`;
}
```

## Google Sheet Structure

### Main Sheet (Sheet1) - Ibanag Texts

| Column A | Column B | Column C | Column D |
|----------|----------|----------|----------|
| Ibanag Text | Status | Date Recorded | Sentence ID |
| Text 1 | Not Yet | | ID001 |
| Text 2 | Not Yet | | ID002 |
| Text 3 | Recorded | 2024-01-15 10:30:00 | ID003 |
| Text 4 | Done | 2024-01-14 15:45:00 | ID004 |
| ... | ... | ... | ... |

**Status values:**
- "Not Yet" - Text is available to be recorded (these will be fetched by the website)
- "Recorded" - Text has been recorded and is pending review (auto-updated on submission)
- "Done" - Recording has been reviewed and approved (you manually update this)

**Sentence ID:**
- Unique identifier for each sentence (e.g., ID001, ID002, etc.)
- Used to cross-reference with the Responses sheet
- Important for tracking which specific sentences were shown/recorded

### Responses Sheet - Submission Log

| Column A | Column B | Column C | Column D | Column E | Column F |
|----------|----------|----------|----------|----------|----------|
| Timestamp | Variant | Recording Type | Topic | Sentence IDs | Number of Sentences |
| 2024-01-15 10:30:00 | Tuguegarao | read | | ID001,ID002,ID003 | 10 |
| 2024-01-15 11:45:00 | Abulug | speak | Daily routine | | 0 |

**Benefits of the Responses sheet:**
- Complete audit trail of all submissions
- Track which sentences were shown to each user
- Analyze submission patterns over time
- Cross-reference sentence IDs with the main sheet
- Know exactly when each submission occurred

## Important: After Updating the Script

1. Save the script
2. Click **Deploy → Manage deployments**
3. Click the **pencil icon** next to your existing deployment
4. Click **Edit**
5. Change version to **"New version"**
6. Click **Deploy**
7. Wait for the deployment to complete (this is crucial!)

## Debugging

If you still get errors:
1. In the Apps Script editor, click **Execution log** at the bottom
2. Run the `doGet` function manually and check the logs
3. Make sure your sheet has the correct column structure
4. Verify the sheet name matches the `TEXTS_SHEET_NAME` constant
5. Ensure you have a "Responses" sheet tab (the script will create it if missing)

## Next Steps

Once you deploy the web app and get the URL, share it with me and I'll update the website to:
1. Fetch 10 random texts with "Not Yet" status from your sheet
2. Submit recordings and update the status to "Recorded" with a date stamp
3. Log all submissions to the Responses sheet with complete details
4. You manually review and update status to "Done"
