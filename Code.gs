// Simple AppScript to read/write Google sheet data from/to Firestore DB
//
// TODO: Firestore db.rules should be changed to writeable only by service account!
// SETUP INSTRUCTIONS:
// 1. After deploying this script, run the 'setupFirebaseConfig' function once to set your Firebase config
// 2. Go to: Apps Script Editor -> Project Settings -> Script Properties
// 3. Add two script properties:
//    - FIREBASE_PROJECT_ID: Your Firebase project ID
//    - FIREBASE_API_KEY: Your Firebase Web API Key
//
// NOTE: Script properties are secure and won't be exposed in your code

// Function to set up Firebase configuration (run once)
function setupFirebaseConfig() {
  const ui = SpreadsheetApp.getUi();
  
  // Prompt for Project ID
  const projectIdResponse = ui.prompt(
    'Firebase Setup (1/2)',
    'Enter your Firebase Project ID:',
    ui.ButtonSet.OK_CANCEL
  );
  if (projectIdResponse.getSelectedButton() == ui.Button.CANCEL) return;
  
  // Prompt for API Key
  const apiKeyResponse = ui.prompt(
    'Firebase Setup (2/2)',
    'Enter your Firebase Web API Key:',
    ui.ButtonSet.OK_CANCEL
  );
  if (apiKeyResponse.getSelectedButton() == ui.Button.CANCEL) return;
  
  // Save to Script Properties
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperties({
    'FIREBASE_PROJECT_ID': projectIdResponse.getResponseText(),
    'FIREBASE_API_KEY': apiKeyResponse.getResponseText()
  });
  
  ui.alert('Setup Complete', 'Firebase configuration has been saved.', ui.ButtonSet.OK);
}

// Get Firebase configuration from Script Properties
function getFirebaseConfig() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const config = {
    projectId: scriptProperties.getProperty('FIREBASE_PROJECT_ID'),
    apiKey: scriptProperties.getProperty('FIREBASE_API_KEY')
  };
  
  if (!config.projectId || !config.apiKey) {
    throw new Error('Firebase configuration not found. Please run the setupFirebaseConfig function first.');
  }
  
  return config;
}

const collectionPath = 'quiz_questions';  // Replace with your collection name

// Define your preferred column order
const COLUMN_ORDER = [
  'id',
  'difficulty',
  'year',
  'audioUrl',
  'answer',
  'choice2',
  'choice3',
  'choice4',
  'videoUrl',
  'updatedAt',
  'createdAt',
  'schemaVersion'  // Add all your preferred columns in order
];

// Define which columns should be written back to Firestore
const WRITABLE_COLUMNS = [
  'difficulty',
  'year',
  'audioUrl',
  'answer',
  'choice2',
  'choice3',
  'choice4',
  'videoUrl'
  // Add any other columns you want to write back
];

// Helper function to extract value from Firestore field based on type
function extractFieldValue(field) {
  if (!field) return '';
  
  // Add logging to see what type of field we're processing
  Logger.log('Processing field:', JSON.stringify(field));
  
  // Check for different Firestore field types
  const possibleTypes = [
    'stringValue',
    'integerValue',
    'doubleValue',
    'booleanValue',
    'timestampValue',
    'arrayValue',
    'mapValue'
  ];
  
  for (const type of possibleTypes) {
    if (field[type] !== undefined) {
      // Handle array values
      if (type === 'arrayValue') {
        return field[type].values ? 
          field[type].values.map(v => extractFieldValue(v)).join(', ') : '';
      }
      // Handle map values
      if (type === 'mapValue') {
        return JSON.stringify(field[type].fields || {});
      }
      // Handle timestamp values specifically
      if (type === 'timestampValue') {
        try {
          Logger.log('Processing timestamp:', field[type]);
          const timestamp = field[type];
          const date = new Date(timestamp);
          const formatted = date.toLocaleString();
          Logger.log('Formatted timestamp:', formatted);
          return formatted;
        } catch (e) {
          Logger.log('Error parsing timestamp:', e);
          return field[type];
        }
      }
      return field[type];
    }
  }
  return '';
}

// Function to read data from Firestore
function readFirestoreData() {
  const config = getFirebaseConfig();
  const sheet = SpreadsheetApp.getActiveSheet();
  
  // Store current column formatting settings before clearing
  const lastColumn = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  const columnFormats = [];
  
  if (lastColumn > 0) {
    for (let col = 1; col <= lastColumn; col++) {
      const columnRange = sheet.getRange(1, col, lastRow || 1, 1);
      columnFormats[col] = {
        wrap: columnRange.getWrap(),
        width: sheet.getColumnWidth(col)
      };
    }
  }
  
  // Construct the URL for Firestore REST API
  const baseUrl = `https://firestore.googleapis.com/v1/projects/${config.projectId}/databases/(default)/documents`;
  const url = `${baseUrl}/${collectionPath}`;
  
  try {
    const options = {
      'method': 'get',
      'contentType': 'application/json',
      'muteHttpExceptions': true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    Logger.log('Response Code:', responseCode);
    
    if (responseCode !== 200) {
      throw new Error(`HTTP Error: ${responseCode} - ${responseText}`);
    }
    
    const data = JSON.parse(responseText);
    
    if (!data.documents || data.documents.length === 0) {
      throw new Error('No documents found in the collection');
    }
    
    // Add debug logging for the first document to see its structure
    Logger.log('First document raw data:');
    Logger.log(JSON.stringify(data.documents[0], null, 2));
    
    // Get all fields from documents
    const allFields = new Set();
    data.documents.forEach(doc => {
      if (doc.fields) {
        Object.keys(doc.fields).forEach(field => allFields.add(field));
      }
    });
    
    // Create headers array based on COLUMN_ORDER and add any missing fields at the end
    const headers = [...COLUMN_ORDER];
    allFields.forEach(field => {
      if (!headers.includes(field)) {
        headers.push(field);
      }
    });
    
    // Prepare rows with ordered fields
    const rows = [headers];
    
    data.documents.forEach(doc => {
      const docId = doc.name.split('/').pop();
      const row = new Array(headers.length).fill('');
      
      // Add debug logging
      Logger.log(`Processing document ${docId}`);
      if (doc.fields) {
        Logger.log(`Available fields: ${Object.keys(doc.fields).join(', ')}`);
        // Add specific logging for timestamp fields
        if (doc.fields.createdAt) {
          Logger.log('createdAt raw value:', JSON.stringify(doc.fields.createdAt));
        }
        if (doc.fields.updatedAt) {
          Logger.log('updatedAt raw value:', JSON.stringify(doc.fields.updatedAt));
        }
      }
      
      // Fill in values according to header order
      headers.forEach((field, index) => {
        if (field === 'id') {
          row[index] = docId;
        } else if (doc.fields && doc.fields[field]) {
          const value = extractFieldValue(doc.fields[field]);
          row[index] = value;
          // Log the field, its raw value, and the extracted value
          Logger.log(`Field ${field}: Raw=${JSON.stringify(doc.fields[field])}, Extracted=${value}`);
        } else {
          Logger.log(`Field ${field} not found in document`);
        }
      });
      
      rows.push(row);
    });
    
    // Clear existing data and write new data
    sheet.clear();
    if (rows.length > 0) {
      // Write the data first
      const dataRange = sheet.getRange(1, 1, rows.length, headers.length);
      dataRange.setValues(rows);
      
      // Format header row
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setBackground('#f3f3f3')
                .setFontWeight('bold');
      
      // Set text clipping for all columns
      const allColumnsRange = sheet.getRange(1, 1, rows.length, headers.length);
      allColumnsRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
      
      // Set reasonable default width for all columns
      headers.forEach((_, index) => {
        const colIndex = index + 1;
        sheet.setColumnWidth(colIndex, 150);
      });
      
      // Add timestamp of last update
      sheet.getRange(rows.length + 2, 1, 1, 2).setValues([
        ['Last Updated:', new Date().toLocaleString()]
      ]);
    }
    
  } catch (error) {
    Logger.log('Error reading from Firestore:', error);
    throw new Error('Failed to read data from Firestore: ' + error.message);
  }
}

// Function to validate columns and write to Firestore
function writeToFirestore() {
  const config = getFirebaseConfig();
  const sheet = SpreadsheetApp.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) {
    throw new Error('No data to write');
  }
  
  const headers = data[0];
  const docIdIndex = headers.indexOf('id');
  
  // Validate id column
  if (docIdIndex === -1) {
    SpreadsheetApp.getActiveSpreadsheet().toast('id column not found', 'Error');
    throw new Error('id column not found');
  }
  
  // Validate writable columns
  const missingColumns = WRITABLE_COLUMNS.filter(col => !headers.includes(col));
  if (missingColumns.length > 0) {
    const errorMessage = `Missing columns in sheet: ${missingColumns.join(', ')}`;
    SpreadsheetApp.getActiveSpreadsheet().toast(errorMessage, 'Error');
    throw new Error(errorMessage);
  }
  
  // Get column indices for writable columns
  const columnIndices = WRITABLE_COLUMNS.reduce((acc, col) => {
    acc[col] = headers.indexOf(col);
    return acc;
  }, {});
  
  const baseUrl = `https://firestore.googleapis.com/v1/projects/${config.projectId}/databases/(default)/documents`;
  let successCount = 0;
  let errorCount = 0;
  
  // Process each row (skip header row and last timestamp row)
  // Find the last real data row (excluding timestamp row)
  let lastDataRowIndex = data.length - 1;
  while (lastDataRowIndex > 0) {
    const row = data[lastDataRowIndex];
    const nonEmptyCells = row.filter(cell => cell !== '').length;
    // If row has only 2 cells (timestamp and value) or fewer, skip it
    if (nonEmptyCells <= 2) {
      lastDataRowIndex--;
    } else {
      break;
    }
  }
  
  // Process rows, excluding header and after lastDataRowIndex
  for (let i = 1; i <= lastDataRowIndex; i++) {
    const row = data[i];
    const docId = row[docIdIndex];
    
    if (!docId) {
      Logger.log(`Skipping row ${i + 1}: No id`);
      continue; // Skip rows without id
    }
    
    const documentUrl = `${baseUrl}/${collectionPath}/${docId}`;
    const fields = {};
    
    // Build the fields object using column indices
    WRITABLE_COLUMNS.forEach(header => {
      const value = row[columnIndices[header]];
      if (value !== '') {
        let fieldValue;
        
        if (typeof value === 'number') {
          if (Number.isInteger(value)) {
            fieldValue = { 'integerValue': value };
          } else {
            fieldValue = { 'doubleValue': value };
          }
        } else if (typeof value === 'boolean') {
          fieldValue = { 'booleanValue': value };
        } else if (value instanceof Date) {
          fieldValue = { 'timestampValue': value.toISOString() };
        } else {
          // Try to parse as JSON if it looks like an array or object
          try {
            if ((value.startsWith('[') && value.endsWith(']')) || 
                (value.startsWith('{') && value.endsWith('}'))) {
              const parsed = JSON.parse(value);
              if (Array.isArray(parsed)) {
                fieldValue = {
                  'arrayValue': {
                    'values': parsed.map(item => ({
                      'stringValue': item.toString()
                    }))
                  }
                };
              } else {
                fieldValue = { 'mapValue': { 'fields': parsed } };
              }
            } else {
              fieldValue = { 'stringValue': value.toString() };
            }
          } catch (e) {
            fieldValue = { 'stringValue': value.toString() };
          }
        }
        
        fields[header] = fieldValue;
      }
    });
    
    // Only make the request if there are fields to update
    if (Object.keys(fields).length > 0) {
      const options = {
        'method': 'PATCH',
        'contentType': 'application/json',
        'muteHttpExceptions': true,
        'payload': JSON.stringify({
          'fields': fields
        })
      };
      
      try {
        const response = UrlFetchApp.fetch(documentUrl, options);
        const responseCode = response.getResponseCode();
        
        if (responseCode === 200) {
          successCount++;
          // Log successful update
          Logger.log(`Updated document ${docId} with fields: ${Object.keys(fields).join(', ')}`);
        } else {
          errorCount++;
          Logger.log(`Error updating document ${docId}: ${response.getContentText()}`);
        }
      } catch (error) {
        errorCount++;
        Logger.log(`Failed to update document ${docId}: ${error.message}`);
      }
    }
  }
  
  // Show completion message with statistics
  const completionMessage = `Update complete:\n${successCount} documents updated successfully\n${errorCount} errors`;
  SpreadsheetApp.getActiveSpreadsheet().toast(completionMessage, 'Status', 10);
}

// Add a function to validate columns without writing
function validateColumns() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  const missingColumns = WRITABLE_COLUMNS.filter(col => !headers.includes(col));
  
  if (missingColumns.length > 0) {
    return {
      valid: false,
      message: `Missing columns: ${missingColumns.join(', ')}`
    };
  }
  
  if (!headers.includes('id')) {
    return {
      valid: false,
      message: 'id column is required'
    };
  }
  
  return {
    valid: true,
    message: 'All required columns are present'
  };
}

// Function to show confirmation dialog and handle write operation
function confirmAndWrite() {
  const ui = SpreadsheetApp.getUi();
  
  // First validate columns
  const validation = validateColumns();
  if (!validation.valid) {
    ui.alert('Validation Error', validation.message, ui.ButtonSet.OK);
    return;
  }
  
  // Show confirmation dialog with details
  const sheet = SpreadsheetApp.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const rowCount = data.length - 1; // Subtract header row
  
  const confirmMessage = 
    `Are you sure you want to write the following to Firestore?\n\n` +
    `Collection: ${collectionPath}\n` +
    `Number of rows to process: ${rowCount}\n` +
    `Fields to write: ${WRITABLE_COLUMNS.join(', ')}\n\n` +
    `This action will update existing documents in Firestore.`;
    
  const response = ui.alert(
    'Confirm Write to Firestore',
    confirmMessage,
    ui.ButtonSet.YES_NO
  );
  
  // Process based on user response
  if (response === ui.Button.YES) {
    try {
      writeToFirestore();
    } catch (error) {
      ui.alert('Error', `Failed to write to Firestore: ${error.message}`, ui.ButtonSet.OK);
    }
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast('Write operation cancelled', 'Cancelled');
  }
}

// Function to write selected row to Firestore
function writeSelectedRowToFirestore() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  const selectedRange = sheet.getActiveRange();
  
  // Get the selected row
  const selectedRow = selectedRange.getRow();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // Validate that a row is selected (not header row)
  if (selectedRow === 1) {
    ui.alert('Error', 'Please select a data row, not the header row.', ui.ButtonSet.OK);
    return;
  }
  
  // Get the row data
  const rowData = sheet.getRange(selectedRow, 1, 1, headers.length).getValues()[0];
  
  // Find id
  const docIdIndex = headers.indexOf('id');
  if (docIdIndex === -1) {
    ui.alert('Error', 'id column not found', ui.ButtonSet.OK);
    return;
  }
  
  const docId = rowData[docIdIndex];
  if (!docId) {
    ui.alert('Error', 'Selected row has no id', ui.ButtonSet.OK);
    return;
  }
  
  // Show confirmation dialog
  const confirmMessage = 
    `Are you sure you want to write this row to Firestore?\n\n` +
    `Collection: ${collectionPath}\n` +
    `id: ${docId}\n` +
    `Fields to write: ${WRITABLE_COLUMNS.join(', ')}\n\n` +
    `This will update the existing document in Firestore.`;
    
  const response = ui.alert(
    'Confirm Write Selected Row',
    confirmMessage,
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Write operation cancelled', 'Cancelled');
    return;
  }
  
  // Prepare the fields to write
  const fields = {};
  WRITABLE_COLUMNS.forEach(header => {
    const colIndex = headers.indexOf(header);
    if (colIndex !== -1) {
      const value = rowData[colIndex];
      if (value !== '') {
        let fieldValue;
        
        if (typeof value === 'number') {
          if (Number.isInteger(value)) {
            fieldValue = { 'integerValue': value };
          } else {
            fieldValue = { 'doubleValue': value };
          }
        } else if (typeof value === 'boolean') {
          fieldValue = { 'booleanValue': value };
        } else if (value instanceof Date) {
          fieldValue = { 'timestampValue': value.toISOString() };
        } else {
          // Try to parse as JSON if it looks like an array or object
          try {
            if ((value.startsWith('[') && value.endsWith(']')) || 
                (value.startsWith('{') && value.endsWith('}'))) {
              const parsed = JSON.parse(value);
              if (Array.isArray(parsed)) {
                fieldValue = {
                  'arrayValue': {
                    'values': parsed.map(item => ({
                      'stringValue': item.toString()
                    }))
                  }
                };
              } else {
                fieldValue = { 'mapValue': { 'fields': parsed } };
              }
            } else {
              fieldValue = { 'stringValue': value.toString() };
            }
          } catch (e) {
            fieldValue = { 'stringValue': value.toString() };
          }
        }
        
        fields[header] = fieldValue;
      }
    }
  });
  
  // Write to Firestore
  const config = getFirebaseConfig();
  const baseUrl = `https://firestore.googleapis.com/v1/projects/${config.projectId}/databases/(default)/documents`;
  const documentUrl = `${baseUrl}/${collectionPath}/${docId}`;
  
  const options = {
    'method': 'PATCH',
    'contentType': 'application/json',
    'muteHttpExceptions': true,
    'payload': JSON.stringify({
      'fields': fields
    })
  };
  
  try {
    const response = UrlFetchApp.fetch(documentUrl, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode === 200) {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `Successfully updated document ${docId}`,
        'Success'
      );
      Logger.log(`Updated document ${docId} with fields: ${Object.keys(fields).join(', ')}`);
    } else {
      ui.alert('Error', `Failed to update document: ${response.getContentText()}`, ui.ButtonSet.OK);
    }
  } catch (error) {
    ui.alert('Error', `Failed to write to Firestore: ${error.message}`, ui.ButtonSet.OK);
  }
}

// Update the menu to include setup option
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Firestore')
    .addItem('Setup Firebase Config', 'setupFirebaseConfig')
    .addItem('Load Data', 'readFirestoreData')
    .addSeparator()
    .addItem('Validate Columns', 'showValidation')
    .addItem('Write Selected Row...', 'writeSelectedRowToFirestore')
    .addItem('Write All Rows...', 'confirmAndWrite')
    .addToUi();
}

// Function to show validation results
function showValidation() {
  const validation = validateColumns();
  SpreadsheetApp.getActiveSpreadsheet().toast(validation.message, 
    validation.valid ? 'Validation Success' : 'Validation Error');
}
