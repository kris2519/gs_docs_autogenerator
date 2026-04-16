// Function to create a new Google Docs document based on the selected template
function createGoogleDocs(templateItem) {
  const template = getTemplateConfig(templateItem);
  if (!template) {
    throw new Error("Invalid template item.");
  }

  const sheet = SpreadsheetApp.getActiveSheet();
  const activeRange = sheet.getActiveRange();

  // Check if there are more than three active cells (columns)
  if (activeRange.getNumColumns() <= 3) {
    SpreadsheetApp.getUi().alert("Error: The entire row is not selected. Please select the entire row before running this script.");
    return;
  }

  // Get the row index of the active cell
  const rowIndex = activeRange.getRow();
  const dataRange = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn());
  const data = dataRange.getValues();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const documentName = replacePlaceholders(template.namePattern, template.placeholders, headers, data[0]);
  const copiedDocumentId = copyDocument(template.fileId, documentName);
  if (copiedDocumentId) {
    const url = generateDocument(templateItem, copiedDocumentId, rowIndex);
    moveDocumentToFolder(copiedDocumentId, destinationFolderId);
  }
}

// Function to get the template configuration for the selected template item
function getTemplateConfig(templateItem) {
  return templateConfig[templateItem] || null;
}

// Function to copy a Google Docs template to create a new document
function copyDocument(sourceDocumentId, targetDocumentName) {
  try {
    const sourceDoc = DriveApp.getFileById(sourceDocumentId);
    const copiedFile = sourceDoc.makeCopy(targetDocumentName);
    return copiedFile.getId();
  } catch (error) {
    Logger.log(`Exception while copying the document: ${error}`);
    throw new Error("Failed to copy the document.");
  }
}

// Function to move the generated document to the specified folder
function moveDocumentToFolder(documentId, folderId) {
  try {
    const file = DriveApp.getFileById(documentId);
    const folder = DriveApp.getFolderById(folderId);
    folder.addFile(file);
    DriveApp.getRootFolder().removeFile(file);
  } catch (error) {
    Logger.log(`Exception while moving the document to the folder: ${error}`);
    throw new Error("Failed to move the document to the specified folder.");
  }
}

// Function to generate the Google Docs document from the template and replace placeholders
function generateDocument(templateName, templateId, rowIndex) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];

  try {
    if (!Number.isInteger(rowIndex) || rowIndex <= 0) {
      throw new Error("Invalid row index.");
    }

    // Open the existing Google Docs template by its ID
    const templateDoc = DocumentApp.openById(templateId);
    const body = templateDoc.getBody();

    // Replace placeholders in the document with corresponding data from the spreadsheet
    const placeholders = templateConfig[templateName].placeholders;
    for (const placeholder in placeholders) {
      const placeholderValue = placeholders[placeholder];
      if (placeholderValue === '{Current Date}') {
        const currentDateLocal = getCurrentDate();
        body.replaceText(`{${placeholder}}`, currentDateLocal);
      } else {
        const columnIndex = getColumnByName(headers, placeholderValue);
        if (isValidIndex(columnIndex)) {
          const valueToReplace = row[columnIndex - 1];

          // Prettify dates if the value is a date
          if (valueToReplace instanceof Date) {
            const formattedDate = Utilities.formatDate(valueToReplace, Session.getScriptTimeZone(), 'MM/dd/yyyy');
            body.replaceText(`{${placeholder}}`, formattedDate);
          } else {
            body.replaceText(`{${placeholder}}`, valueToReplace);
          }
        }
      }
    }

    // Save and close the updated document
    templateDoc.saveAndClose();

    // Get the URL of the updated document
    const url = templateDoc.getUrl();

    // Provide a link to open the document
    const linkText = "Open " + templateDoc.getName();
    const htmlOutput = `
      <html>
        <body style="text-align: center; font-size: 18px; font-family: Calibri, Arial, sans-serif;">
          <p>Document has been generated. Click below link to open:</p>
          <a href="${url}" style="color: green;" target="_blank">${linkText}</a>
        </body>
      </html>
    `;

    // Show the link to the user in a custom HTML dialog
    const title = "Document Generated";
    SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(htmlOutput), title);

    return url;
  } catch (error) {
    // Log the error for debugging purposes
    Logger.log(`Error while processing document: ${error}`);

    // Throw a custom error or handle the situation accordingly
    throw new Error("Permission error: The script does not have sufficient permissions to edit the document.");
  }
}

// Function to check if the index is valid (not null and greater than 0)
function isValidIndex(index) {
  return index !== null && index > 0;
}

// Function to get the column index by column name
function getColumnByName(headers, columnName) {
  const columnIndex = headers.indexOf(columnName);
  return columnIndex >= 0 ? columnIndex + 1 : null;
}

// Function to replace placeholders in the text with corresponding values
function replacePlaceholders(text, placeholders, headers, rowData) {
  return text.replace(/\{.*?\}/g, match => {
    const placeholder = match.substring(1, match.length - 1).trim();
    if (placeholder === 'Current Date') {
      return getCurrentDate();
    } else {
      const columnIndex = getColumnByName(headers, placeholders[placeholder]);
      return columnIndex !== null ? rowData[columnIndex - 1] : match;
    }
  });
}

// Function to get the current date in the required format
function getCurrentDate() {
  const currentDate = new Date();
  const formattedDate = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), 'MM/dd/yyyy');
  return formattedDate;
}
