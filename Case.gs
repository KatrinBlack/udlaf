// Define named constants for template IDs and destination folder ID
const TEMPLATE_APPLICATION = '1AILohNfCVQLE_iXNJ39lTdWmUZQvOhHX_NF_Kfjm_5g';      // URL for Application Template response document

const DESTINATION_FOLDER = '1Aujv0XBQ7h38EbfRVv-jdQWVRnfbay1U';

// Define placeholders for the Application template
const PLACEHOLDERS_APPLICATION = {
  'Date': 0,
  'FirstName': 1,
  'LastName': 2,
  'Street': 3,
  'City': 4,
  'State': 5,
  'Postal': 6,
  'Country': 7,
  'DateOfBirth': 8,
  'EmailP': 9,
  'Phone': 10,

  'Embassy/Consulate': 11,
  'JobTitle': 12,  
  'Status': 13,
  'GAP': 14,
  'EmailW': 15,

  'EmailPreference': 16,
  'DuesPreference': 17,
};

/*function onOpen() {
  SpreadsheetApp.getActiveSpreadsheet().addMenu('Approval', [
    { name: 'Approve Applicant', functionName: 'approveApplicant' }
  ]);
}

// Function to approve an applicant
function approveApplicant() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Applications');
  const rows = sheet.getDataRange().getValues();
  
  if (!rows[1]) {
    ui.alert('No Application in row 1');
  } else {
    ui.alert(`Approve Applicant ${rows[1][1]}`, 'caseReport');
  }
}
*/
// Triggered when the spreadsheet is opened
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('Applications');
   
  // Get data from the 'Form Center' sheet
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Applications');
  const rows = sheet.getDataRange().getValues();

  // Add menu items based on case availability
  if (!rows[1]) {
    menu.addItem('No Names found', 'empty');
  } else {
    menu.addItem(`Approve ${rows[1][1]} ${rows[1][2]}`, 'generateentry');
  }
  
  menu.addToUi();
}

// Placeholder for empty function
function empty() {}

// Function to convert datetime string to date only (yyyy-MM-dd format)
function dateOnly(datetimeString) {
  const date = new Date(datetimeString);
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

// Replaces placeholders in the document body with actual values or 'N/A' if empty
function replacePlaceholders(body, placeholders) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Applications');
  const rows = sheet.getDataRange().getValues();
  
  for (const placeholder in placeholders) {
    let value = rows[1][placeholders[placeholder]];
    
    // Convert datetime to date only for 'Date' placeholder
    if (placeholder === 'Date') {
      value = dateOnly(value);
    }
    
    const textToReplace = value !== '' ? value : 'N/A'; // Use 'N/A' if the value is empty
    
    body.replaceText(`{{${placeholder}}}`, textToReplace);
  }
}

// Generates a document from a template and replaces placeholders
function generateDocumentFromTemplate(templateId, destinationFolder, fileName, placeholders) {
  try {
    const googleDocTemplate = DriveApp.getFileById(templateId);

    const copy = googleDocTemplate.makeCopy(fileName, destinationFolder);
    const doc = DocumentApp.openById(copy.getId());
    const body = doc.getBody();

    replacePlaceholders(body, placeholders);

    doc.saveAndClose();
  } catch (error) {
    Logger.log(`Error generating document: ${error}`);
  }
}

// Generates case materials for a student case
function document_generation(rows) {
  const destinationFolder = DriveApp.getFolderById(DESTINATION_FOLDER);
  const fileName = `${rows[1][1]}_${rows[1][2]}`;
  generateDocumentFromTemplate(TEMPLATE_APPLICATION, destinationFolder, fileName, PLACEHOLDERS_APPLICATION);
}

// Entry point for generating case materials
function generateentry() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Applications');
  const rows = sheet.getDataRange().getValues();
  document_generation(rows);
}

