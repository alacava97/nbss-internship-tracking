const SPREADSHEET_ID = '1xguSZeNZD-lXfcFQosPgUqaKvT06f_XNXDQuibg5SYM';
const SHEET_NAME = 'Internship Opportunities';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('form').setTitle('Internships');
}

function saveInternshipOpportunity(data) {
  if (!data || typeof data !== 'object') {
    throw new Error('No internship information was received.');
  }

  const requiredFields = [
    ['provider', 'Opportunity Provider'],
    ['location', 'Location'],
    ['start_date', 'Start Date'],
    ['duration', 'Internship Duration'],
    ['duration_unit', 'Duration Unit'],
    ['intern_count', 'Number of Interns'],
    ['application_deadline', 'Application Deadline'],
    ['car_requirement', 'Car Requirement']
  ];

  requiredFields.forEach(([key, label]) => {
    if (data[key] == null || String(data[key]).trim() === '') {
      throw new Error(label + ' is required.');
    }
  });

  if (!Number.isInteger(Number(data.duration)) || Number(data.duration) < 1) {
    throw new Error('Internship Duration must be a whole number greater than zero.');
  }
  if (!Number.isInteger(Number(data.intern_count)) || Number(data.intern_count) < 1) {
    throw new Error('Number of Interns must be a whole number greater than zero.');
  }
  if (!['days', 'weeks', 'months'].includes(data.duration_unit)) {
    throw new Error('Please select a valid duration unit.');
  }
  if (!['not_required', 'recommended', 'required'].includes(data.car_requirement)) {
    throw new Error('Please select a valid car requirement.');
  }
  if (!data.unpaid) {
    if (!['hourly', 'weekly', 'lump_sum'].includes(data.compensation_frequency)) {
      throw new Error('Pay Period is required unless the internship is unpaid.');
    }
    if (data.compensation_amount == null || String(data.compensation_amount).trim() === '') {
      throw new Error('Compensation Amount is required unless the internship is unpaid.');
    }
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
    const date = new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]), 12);
    if (date.getFullYear() !== Number(match[1]) ||
        date.getMonth() !== Number(match[2]) - 1 ||
        date.getDate() !== Number(match[3])) {
      throw new Error('Please enter valid dates.');
    }
    return date;
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
    data.unpaid ? '' : textValue(data.compensation_frequency),
    data.unpaid ? '' : numberValue(data.compensation_amount),
    Boolean(data.unpaid),
    Boolean(data.overtime_available),
    dateValue(data.application_deadline)
  ];

  sheet.appendRow(row);
  return { success: true, message: 'Internship opportunity saved to the spreadsheet.' };
}
