# Google Sheets ↔ Firestore Integration

This Google Apps Script enables bidirectional synchronization between Google Sheets and Firebase Firestore. It allows you to read data from Firestore into a spreadsheet and write data back to Firestore, making it perfect for managing and editing Firestore data through a familiar spreadsheet interface.

## Features

- 📥 **Read from Firestore**: Load Firestore collection data into your spreadsheet
- 📤 **Write to Firestore**: Update Firestore documents directly from your spreadsheet
- ✨ **Smart Data Handling**: Automatically handles various data types (numbers, booleans, arrays, objects, timestamps)
- 🔒 **Secure Configuration**: Firebase credentials stored securely in Script Properties
- ✅ **Data Validation**: Validates data before writing to Firestore
- 🎯 **Selective Updates**: Write individual rows or bulk update all rows
- 📋 **Column Management**: Configurable column order and writable fields

## Prerequisites

1. A Google Firebase project with Firestore database
2. Firebase Web API Key
3. Firebase Project ID
4. Google Sheets access

## Setup Instructions

### 1. Create a New Google Sheet

### 2. Set Up the Apps Script

1. In your Google Sheet, go to `Extensions > Apps Script`
2. Copy the contents of `Code.gs` into the script editor
3. Save the project

### 3. Configure Firebase Credentials

1. In the sheet, you'll see a new menu item called "Firestore"
2. Click `Firestore > Setup Firebase Config`
3. Enter your Firebase Project ID when prompted
4. Enter your Firebase Web API Key when prompted

### 4. Configure Collection Settings (Optional)

In the script, you can modify these constants according to your needs:
- `collectionPath`: The Firestore collection to sync with
- `COLUMN_ORDER`: Preferred order of columns in the sheet
- `WRITABLE_COLUMNS`: Fields that can be written back to Firestore

## Usage

### Reading Data from Firestore

1. Click `Firestore > Load Data`
2. The script will:
   - Fetch all documents from your Firestore collection
   - Display them in the sheet with formatted headers
   - Show a timestamp of the last update

### Writing Data to Firestore

#### Option 1: Write a Single Row
1. Select any cell in the row you want to update
2. Click `Firestore > Write Selected Row...`
3. Confirm the operation in the dialog

#### Option 2: Write All Rows
1. Click `Firestore > Write All Rows...`
2. Review the confirmation dialog showing affected rows
3. Confirm the operation

### Validation

- Click `Firestore > Validate Columns` to check if all required columns are present
- The script automatically validates data before writing to Firestore

## Data Type Support

The integration handles these Firestore data types:
- Strings
- Numbers (Integer/Double)
- Booleans
- Timestamps
- Arrays
- Maps/Objects

## Column Formatting

- Headers are formatted with a gray background
- Text in cells is clipped (showing ... for overflow)
- Default column width is set to 150 pixels

## Troubleshooting

### Common Issues

1. **Configuration Missing**
   - Error: "Firebase configuration not found"
   - Solution: Run the Setup Firebase Config process

2. **Missing Columns**
   - Error: "Missing columns in sheet"
   - Solution: Ensure all required columns from WRITABLE_COLUMNS exist

3. **Invalid Document ID**
   - Error: "Document ID column not found" or "No Document ID"
   - Solution: Ensure the "Document ID" column exists and contains valid IDs

### Error Messages

The script provides feedback through:
- Toast notifications for success/minor issues
- Alert dialogs for critical errors
- Execution logs for detailed debugging

## Security Considerations

- Firebase credentials are stored securely in Script Properties
- The script uses HTTPS for all Firestore API calls
- Consider setting up appropriate Firestore Security Rules

## Limitations

- Maximum 50,000 cells per sheet (Google Sheets limit)
- API quotas apply for both Apps Script and Firestore
- Write operations are limited to configured WRITABLE_COLUMNS

## Contributing

Feel free to submit issues and enhancement requests!

## License

This project is licensed under the MIT License - see the LICENSE file for details 