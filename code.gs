//to do
//-create a way to reopen submitted forms and repopulate the submitted data.
//-Student facing page to allow students to browse internships
//-Are we supposed to call them something else besieds internships?
//-Add contact info and supporting documents to form.
//-Create a way for this to expand to other departments, potentially including an admin panel on the spreadsheet page

const SPREADSHEET_ID = '1xguSZeNZD-lXfcFQosPgUqaKvT06f_XNXDQuibg5SYM';
const SHEET_NAME = 'Internship Opportunities';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('form').setTitle('Internships');
}

function saveInternshipOpportunity(data) {
  if (!data || typeof data !== 'object') {
    throw new Error('No internship information was received.');
  }

  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
  }

  const headers = [
    'Submitted At',
    'Opportunity Provider',
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
    const match = String(value).match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!match) throw new Error('Please enter valid dates.');
    return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]), 12);
  };

  const row = [
    new Date(),
    textValue(data.provider),
    textValue(data.location),
    dateValue(data.start_date),
    numberValue(data.duration),
    textValue(data.duration_unit),
    numberValue(data.intern_count),
    Boolean(data.housing_provided),
    Boolean(data.meals_provided),
    textValue(data.car_requirement),
    textValue(data.compensation_frequency),
    numberValue(data.compensation_amount),
    Boolean(data.unpaid),
    Boolean(data.overtime_available),
    dateValue(data.application_deadline)
  ];

  sheet.appendRow(row);
  return { success: true, message: 'Your submission has been recorded.' };
}
