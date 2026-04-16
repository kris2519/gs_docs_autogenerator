// Function to create the custom menu in the Google Sheets UI
function onOpen() {
  SpreadsheetApp.getUi().createMenu('AutoFill Docs')
    .addItem('Initial Document', 'initialDocument')
    .addToUi();
}

// Function to create a new document
function initialDocument() {
  createGoogleDocs('initialDocument');
}
