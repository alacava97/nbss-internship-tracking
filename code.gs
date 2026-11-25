//to do
//-create a way to reopen submitted forms and repopulate the submitted data.
//-Student facing page to allow students to browse internships
//-Are we supposed to call them something else besieds internships?
//-Create a way for this to expand to other departments, potentially including an admin panel on the spreadsheet page

const SPREADSHEET_ID = '1xguSZeNZD-lXfcFQosPgUqaKvT06f_XNXDQuibg5SYM';
const SHEET_NAME = 'Internship Opportunities';
const SUPPORTING_DOCUMENTS_FOLDER = 'Internship Supporting Documents';

function doGet(e) {
  const token = e && e.parameter ? String(e.parameter.token || '') : '';
  const initialData = token ? getInternshipByEditToken_(token) : null;
  if (token && !initialData) {
    return HtmlService.createHtmlOutput('This edit link is invalid or no longer available.');
  }
  const template = HtmlService.createTemplateFromFile('form');
  template.initialDataJson = JSON.stringify(initialData || {}).replace(/</g, '\\u003c');
  template.isEditing = Boolean(initialData);
  return template.evaluate().setTitle(initialData ? 'Update Internship Information' : 'Internships');
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Internship Form')
    .addItem('Create edit link for selected row', 'createEditLinkForSelectedRow')
    .addToUi();
}

function createEditLinkForSelectedRow() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet || sheet.getActiveRange().getSheet().getName() !== SHEET_NAME) {
    throw new Error('Select a row on the Internship Opportunities sheet.');
  }
  const rowNumber = sheet.getActiveRange().getRow();
  if (rowNumber < 2) throw new Error('Select an internship entry, not the header row.');

  const headers = ensureEditColumns_(sheet);
  const tokenColumn = headers.indexOf('Edit Token') + 1;
  const linkColumn = headers.indexOf('Edit Link') + 1;
  let token = String(sheet.getRange(rowNumber, tokenColumn).getValue() || '');
  if (!token) {
    token = Utilities.getUuid();
    sheet.getRange(rowNumber, tokenColumn).setValue(token);
  }
  const url = ScriptApp.getService().getUrl() + '?token=' + encodeURIComponent(token);
  sheet.getRange(rowNumber, linkColumn).setValue(url);
  SpreadsheetApp.getUi().alert('Edit link created', url, SpreadsheetApp.getUi().ButtonSet.OK);
}

function ensureEditColumns_(sheet) {
  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(String);
  ['Edit Token', 'Edit Link'].forEach(header => {
    if (!headers.includes(header)) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(header);
      headers.push(header);
    }
  });
  return headers;
}

function getInternshipByEditToken_(token) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return null;
  const headers = ensureEditColumns_(sheet);
  const tokenColumn = headers.indexOf('Edit Token') + 1;
  const values = sheet.getDataRange().getValues();
  const rowIndex = values.findIndex((row, index) => index > 0 && String(row[tokenColumn - 1]) === token);
  if (rowIndex < 1) return null;
  const row = values[rowIndex];
  const get = name => row[headers.indexOf(name)] ?? '';
  const dateString = value => value instanceof Date
    ? Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd')
    : '';
  return {
    edit_token: token,
    provider: get('Opportunity Provider'),
    location: get('Location'),
    start_date: dateString(get('Start Date')),
    duration: get('Duration'),
    duration_unit: get('Duration Unit'),
    intern_count: get('Number of Interns'),
    housing_provided: Boolean(get('Housing Provided')),
    meals_provided: Boolean(get('Meals Provided')),
    car_requirement: get('Car Requirement'),
    compensation_frequency: get('Compensation Frequency'),
    compensation_amount: get('Compensation Amount (USD)'),
    unpaid: Boolean(get('Unpaid Internship')),
    overtime_available: Boolean(get('Overtime Available')),
    application_deadline: dateString(get('Application Deadline')),
    contact_name: get('Provider Contact Name'),
    contact_email: get('Provider Contact Email'),
    contact_phone: get('Provider Contact Phone'),
    description: get('Internship Description'),
    supporting_documents: get('Supporting Documents')
  };
}

function saveInternshipOpportunity(data) {
  if (!data || typeof data !== 'object') throw new Error('No internship information was received.');

  const requiredFields = [
    'provider', 'contact_name', 'contact_email', 'contact_phone', 'description', 'location',
    'start_date', 'duration', 'duration_unit', 'intern_count',
    'application_deadline', 'car_requirement'
  ];
  for (const field of requiredFields) {
    if (data[field] == null || String(data[field]).trim() === '') {
      throw new Error('Please complete all required fields.');
    }
  }
  if (!data.unpaid && (!data.compensation_frequency ||
      data.compensation_amount == null || String(data.compensation_amount).trim() === '')) {
    throw new Error('Please provide compensation details or select Unpaid internship.');
  }

  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('The internship spreadsheet was not found.');
  const headers = ensureEditColumns_(sheet);
  const headerMap = Object.fromEntries(headers.map((header, index) => [header, index + 1]));

  const textValue = value => {
    const text = value == null ? '' : String(value).trim();
    return /^[=+@-]/.test(text) ? "'" + text : text;
  };
  const numberValue = value => {
    if (value == null || value === '') return '';
    const number = Number(value);
    if (!Number.isFinite(number) || number < 0) throw new Error('Please enter valid non-negative numbers.');
    return number;
  };
  const dateValue = value => {
    if (!value) return '';
    const match = String(value).match(/^(\\d{4})-(\\d{2})-(\\d{2})$/);
    if (!match) throw new Error('Please enter valid dates.');
    return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]), 12);
  };

  let rowNumber = 0;
  let token = String(data.edit_token || '');
  if (token) {
    const tokenValues = sheet.getRange(2, headerMap['Edit Token'], Math.max(sheet.getLastRow() - 1, 1), 1)
      .getValues().flat().map(String);
    const foundIndex = tokenValues.indexOf(token);
    if (foundIndex < 0) throw new Error('This edit link is invalid or no longer available.');
    rowNumber = foundIndex + 2;
  }

  const existingDocuments = rowNumber
    ? String(sheet.getRange(rowNumber, headerMap['Supporting Documents']).getValue() || '')
    : '';
  const uploadedDocuments = Array.isArray(data.supporting_documents) ? data.supporting_documents : [];
  if (!rowNumber && uploadedDocuments.length === 0) {
    throw new Error('Please attach at least one supporting document.');
  }

  let documentLinks = existingDocuments;
  if (uploadedDocuments.length) {
    const folderIterator = DriveApp.getFoldersByName(SUPPORTING_DOCUMENTS_FOLDER);
    const folder = folderIterator.hasNext() ? folderIterator.next() : DriveApp.createFolder(SUPPORTING_DOCUMENTS_FOLDER);
    const newLinks = uploadedDocuments.map(document => {
      if (!document || !document.name || !document.data) throw new Error('One of the supporting documents could not be read.');
      const bytes = Utilities.base64Decode(document.data);
      const blob = Utilities.newBlob(bytes, document.mimeType || 'application/octet-stream', document.name);
      const file = folder.createFile(blob);
      return file.getName() + ': ' + file.getUrl();
    });
    documentLinks = [existingDocuments, newLinks.join('\\n')].filter(Boolean).join('\\n');
  }

  if (!token) token = Utilities.getUuid();
  const values = {
    'Opportunity Provider': textValue(data.provider),
    'Location': textValue(data.location),
    'Start Date': dateValue(data.start_date),
    'Duration': numberValue(data.duration),
    'Duration Unit': textValue(data.duration_unit),
    'Number of Interns': numberValue(data.intern_count),
    'Housing Provided': Boolean(data.housing_provided),
    'Meals Provided': Boolean(data.meals_provided),
    'Car Requirement': textValue(data.car_requirement),
    'Compensation Frequency': data.unpaid ? '' : textValue(data.compensation_frequency),
    'Compensation Amount (USD)': data.unpaid ? '' : numberValue(data.compensation_amount),
    'Unpaid Internship': Boolean(data.unpaid),
    'Overtime Available': Boolean(data.overtime_available),
    'Application Deadline': dateValue(data.application_deadline),
    'Provider Contact Name': textValue(data.contact_name),
    'Provider Contact Email': textValue(data.contact_email),
    'Provider Contact Phone': textValue(data.contact_phone),
    'Internship Description': textValue(data.description),
    'Supporting Documents': documentLinks,
    'Edit Token': token
  };

  if (rowNumber) {
    Object.entries(values).forEach(([header, value]) => sheet.getRange(rowNumber, headerMap[header]).setValue(value));
    return { success: true, message: 'Your changes have been saved.' };
  }

  const newRow = new Array(headers.length).fill('');
  newRow[headerMap['Submitted At'] - 1] = new Date();
  Object.entries(values).forEach(([header, value]) => { newRow[headerMap[header] - 1] = value; });
  rowNumber = sheet.getLastRow() + 1;
  sheet.getRange(rowNumber, 1, 1, newRow.length).setValues([newRow]);
  const url = ScriptApp.getService().getUrl() + '?token=' + encodeURIComponent(token);
  sheet.getRange(rowNumber, headerMap['Edit Link']).setValue(url);
  return { success: true, message: 'Your submission has been recorded.' };
}
