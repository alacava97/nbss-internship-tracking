const SPREADSHEET_ID = '1xguSZeNZD-lXfcFQosPgUqaKvT06f_XNXDQuibg5SYM';
const SHEET_NAME = 'Internship Opportunities';
const SUPPORTING_DOCUMENTS_FOLDER = 'Internship Supporting Documents';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('form').setTitle('Internships');
}

function saveInternshipOpportunity(data) {
  if (!data || typeof data !== 'object') {
    throw new Error('No internship information was received.');
  }

  const requiredFields = [
    'provider', 'contact_name', 'contact_email', 'contact_phone',
    'description', 'location', 'start_date', 'duration', 'duration_unit',
    'intern_count', 'application_deadline', 'car_requirement'
  ];
  for (const field of requiredFields) {
    if (data[field] == null || String(data[field]).trim() === '') {
      throw new Error('Please complete all required fields.');
    }
  }
  if (!data.unpaid && (
    !data.compensation_frequency ||
    data.compensation_amount == null ||
    String(data.compensation_amount).trim() === ''
  )) {
    throw new Error('Please provide compensation details or select Unpaid internship.');
  }
  if (!Array.isArray(data.supporting_documents) || data.supporting_documents.length === 0) {
    throw new Error('Please attach at least one supporting document.');
  }

  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(SHEET_NAME);

  const headers = [
    'Submitted At',
    'Opportunity Provider',
    'Provider Contact Name',
    'Provider Contact Email',
    'Provider Contact Phone',
    'Internship Description',
    'Supporting Documents',
    'Location',
    'Start Date',
    'Duration',
    'Duration Unit',
    'Number of Interns',
    'Housing Provided',
    'Meals Provided',
    'Car Requirement',
    'Compensation Frequency',
    'Compensation Amount (USD)',
    'Unpaid Internship',
    'Overtime Available',
    'Application Deadline'
  ];

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }

  const textValue = value => {
    const text = value == null ? '' : String(value).trim();
    return /^[=+@-]/.test(text) ? "'" + text : text;
  };
  const numberValue = value => {
    if (value == null || value === '') return '';
    const number = Number(value);
    if (!Number.isFinite(number) || number < 0) {
      throw new Error('Please enter valid non-negative numbers.');
    }
    return number;
  };
  const dateValue = value => {
    if (!value) return '';
    const match = String(value).match(/^(\\d{4})-(\\d{2})-(\\d{2})$/);
    if (!match) throw new Error('Please enter valid dates.');
    return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]), 12);
  };

  const folderIterator = DriveApp.getFoldersByName(SUPPORTING_DOCUMENTS_FOLDER);
  const folder = folderIterator.hasNext()
    ? folderIterator.next()
    : DriveApp.createFolder(SUPPORTING_DOCUMENTS_FOLDER);

  const documentLinks = data.supporting_documents.map(document => {
    if (!document || !document.name || !document.data) {
      throw new Error('One of the supporting documents could not be read.');
    }
    const bytes = Utilities.base64Decode(document.data);
    const blob = Utilities.newBlob(
      bytes,
      document.mimeType || 'application/octet-stream',
      document.name
    );
    const file = folder.createFile(blob);
    return file.getName() + ': ' + file.getUrl();
  });

  const row = [
    new Date(),
    textValue(data.provider),
    textValue(data.contact_name),
    textValue(data.contact_email),
    textValue(data.contact_phone),
    textValue(data.description),
    documentLinks.join('\\n'),
    textValue(data.location),
    dateValue(data.start_date),
    numberValue(data.duration),
    textValue(data.duration_unit),
    numberValue(data.intern_count),
    Boolean(data.housing_provided),
    Boolean(data.meals_provided),
    textValue(data.car_requirement),
    data.unpaid ? '' : textValue(data.compensation_frequency),
    data.unpaid ? '' : numberValue(data.compensation_amount),
    Boolean(data.unpaid),
    Boolean(data.overtime_available),
    dateValue(data.application_deadline)
  ];

  sheet.appendRow(row);
  return { success: true, message: 'Your submission has been recorded.' };
}
