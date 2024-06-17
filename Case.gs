// Define named constants for template IDs and destination folder ID
const TEMPLATE_APPLICATION = '1AILohNfCVQLE_iXNJ39lTdWmUZQvOhHX_NF_Kfjm_5g';      // URL for Application response

const DESTINATION_FOLDER = '112ReA8Ct-bTbl_RzcMX2fJqFI-7k-YRG';

// Define placeholders for student and faculty cases
const PLACEHOLDERS_APPLICATION = {
  'Date': 0,
  'FirstName': 0,
  'LastName': 2,
  'Street': 3,
  'City': 25,
  'State': 26,
  'Postal': 27,
  'Country': 0,
  'DateOfBirth': 0,
  'EmailP': 2,
  'Phone': 3,

  'Embassy': 25,
  'Consulate': 26,
  'JobTitle': 27,  
  'Status': 25,
  'GAP': 26,
  'EmailW': 27,

  'EmailPreference': 25,
  'DuesPreference': 26,
};

// Triggered when the spreadsheet is opened
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('Application Approval');
   
  // Get data from the 'Form Responses' sheet
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form Responses');
  const rows = sheet.getDataRange().getValues();

  // Add menu items based on case availability
  if (!rows[1]) {
    menu.addItem('No Application in row 1', 'empty');
  } else {
    menu.addItem(`Approve Applicant ${rows[1][1]}`, 'caseReport');
  }
  
  menu.addToUi();
}

// Placeholder for empty function
function empty() {}

// Replaces placeholders in the document body with actual values
function replacePlaceholders(body, placeholders) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form Responses');
  const rows = sheet.getDataRange().getValues();
  
  for (const placeholder in placeholders) {
    body.replaceText(`{{${placeholder}}}`, rows[1][placeholders[placeholder]]);
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
function caseStudent(rows) {

  const caseFolder = DriveApp.getFolderById(DESTINATION_FOLDER).createFolder(`${rows[1][0]}`);

  generateDocumentFromTemplate(TEMPLATE_STUDENT_0, caseFolder, `${rows[1][0]} - Case Report`, PLACEHOLDERS_STUDENT);
}


// Entry point for generating case materials
function caseReport() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form Responses');
  const rows = sheet.getDataRange().getValues();
}
