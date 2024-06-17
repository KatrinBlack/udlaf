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
    menu.addItem(`Approve Applicant ${rows[1][0]}`, 'caseReport');
  }
  
  menu.addToUi();
}

// Placeholder for empty function
function empty() {}

// Replaces placeholders in the document body with actual values
function replacePlaceholders(body, placeholders) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form Center');
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
  generateDocumentFromTemplate(TEMPLATE_STUDENT_1, caseFolder, `${rows[1][0]} - Notification Memo`, PLACEHOLDERS_STUDENT);
  generateDocumentFromTemplate(TEMPLATE_STUDENT_2, caseFolder, `${rows[1][0]} - Materials Request`, PLACEHOLDERS_STUDENT);
  generateDocumentFromTemplate(TEMPLATE_STUDENT_3, caseFolder, `${rows[1][0]} - Implicating Student Testimony`, PLACEHOLDERS_STUDENT);
  generateDocumentFromTemplate(TEMPLATE_STUDENT_4, caseFolder, `${rows[1][0]} - Interview Invitation`, PLACEHOLDERS_STUDENT);
  generateDocumentFromTemplate(TEMPLATE_STUDENT_5, caseFolder, `${rows[1][0]} - Case Debriefing`, PLACEHOLDERS_STUDENT);
  generateDocumentFromTemplate(TEMPLATE_STUDENT_6, caseFolder, `${rows[1][0]} - Professor Notification`, PLACEHOLDERS_STUDENT);
}

// Generates case materials for a faculty case
function caseFaculty(rows) {

  const caseFolder = DriveApp.getFolderById(DESTINATION_FOLDER).createFolder(`${rows[1][0]}`);

  generateDocumentFromTemplate(TEMPLATE_FACULTY_0, caseFolder, `${rows[1][0]} - Case Report`, PLACEHOLDERS_FACULTY);
  generateDocumentFromTemplate(TEMPLATE_FACULTY_1, caseFolder, `${rows[1][0]} - Notification Memo`, PLACEHOLDERS_FACULTY);
  generateDocumentFromTemplate(TEMPLATE_FACULTY_2, caseFolder, `${rows[1][0]} - Materials Request`, PLACEHOLDERS_FACULTY);
  generateDocumentFromTemplate(TEMPLATE_FACULTY_4, caseFolder, `${rows[1][0]} - Interview Invitation`, PLACEHOLDERS_FACULTY);
  generateDocumentFromTemplate(TEMPLATE_FACULTY_5, caseFolder, `${rows[1][0]} - Case Debriefing`, PLACEHOLDERS_FACULTY);
  generateDocumentFromTemplate(TEMPLATE_FACULTY_6, caseFolder, `${rows[1][0]} - Professor Notification`, PLACEHOLDERS_FACULTY);
}

// Entry point for generating case materials
function caseReport() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form Center');
  const rows = sheet.getDataRange().getValues();

  // Determine case type and generate materials accordingly
  if (rows[1][5] === 'Student-Implicated: Faculty/staff member reporting a student-implicated Honor System violation.') {
    caseStudent(rows);
  } else if (rows[1][5] === 'Faculty-Implicated: Faculty/staff member reporting a faculty-implicated Honor System violation.') {
    caseFaculty(rows);
  }
}
