/** Creates a NEW synthetic workbook and previews reminders without sending email. */
function runDemo() {
  if (!CONFIG.DRY_RUN) throw new Error('Set CONFIG.DRY_RUN to true before running the demo.');
  const ss = createDemoSpreadsheet_();
  const result = previewDemoReminders();
  Logger.log('Open the demo workbook: ' + ss.getUrl());
  return {url: ss.getUrl(), summary: result};
}

function createDemoSpreadsheet_() {
  const ss = SpreadsheetApp.create('Certification reminder tracker - synthetic demo');
  ss.setSpreadsheetTimeZone('America/Los_Angeles');
  const instructions = ss.getSheets()[0];
  instructions.setName('Read me');
  const guidance = [
    ['Certification reminder tracker - portfolio demonstration'],
    ['All names, addresses, and dates are synthetic. No original workbook is needed.'],
    ['Demo Preview contains the messages that would be generated; no email was sent.'],
    ['Red email cells simulate addresses manually flagged as invalid or bounced.'],
    ['Change synthetic dates and run previewDemoReminders again to explore the rules.'],
    ['runDemo creates a new workbook; previewDemoReminders reuses the saved demo workbook.'],
    ['Optional attachments are not needed for the demonstration.']
  ];
  instructions.getRange(1, 1, guidance.length, 1).setValues(guidance);
  instructions.setColumnWidth(1, 780);
  const header = ['Nurse name', 'Nurse email', 'Supervisor email', 'Certification type', 'Expiration date'];
  const today = new Date();
  const fixtures = syntheticRecords_(today);
  for (const config of CONFIG.SHEETS_TO_SCAN) {
    const sheet = ss.insertSheet(config.name);
    const records = fixtures.filter(x => x.sheet === config.name);
    const data = [header].concat(records.map(x => x.values));
    sheet.getRange(1, 1, data.length, header.length).setValues(data);
    sheet.getRange(1, 1, 1, header.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    if (records.length) sheet.getRange(2, 5, records.length, 1).setNumberFormat('yyyy-mm-dd');
    records.forEach((x, i) => {if (x.flagged) sheet.getRange(i + 2, 2).setBackground('#ff0000');});
    sheet.autoResizeColumns(1, 5);
  }
  PropertiesService.getScriptProperties().setProperty('TRACKER_SPREADSHEET_ID', ss.getId());
  return ss;
}

/** Public entry point for previewing edits in the most recently created demo. */
function previewDemoReminders() {
  if (!CONFIG.DRY_RUN) throw new Error('The preview requires CONFIG.DRY_RUN = true.');
  const id = PropertiesService.getScriptProperties().getProperty('TRACKER_SPREADSHEET_ID');
  if (!id) throw new Error('Run runDemo first to create a synthetic workbook.');
  const ss = SpreadsheetApp.openById(id);
  const rows = [['Nurse', 'Recipient', 'Role', 'Credential', 'Expiration', 'Stage', 'Subject', 'Plain text message']];
  const summary = runReminderCheck_(rows);
  let sheet = ss.getSheetByName('Demo Preview');
  if (!sheet) sheet = ss.insertSheet('Demo Preview');
  sheet.clearContents();
  sheet.getRange(1, 1, rows.length, 8).setValues(rows);
  sheet.getRange(1, 1, 1, 8).setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 7);
  sheet.setColumnWidth(8, 650);
  sheet.getRange(1, 8, rows.length, 1).setWrap(true);
  Logger.log('Preview ready. No email sent. ' + JSON.stringify(summary));
  return summary;
}

/** Date offsets keep the demo useful on the day a reviewer runs it. */
function syntheticRecords_(today) {
  function expiry(days) {
    const base = Utilities.formatDate(today, 'America/Los_Angeles', 'yyyy-MM-dd');
    const value = new Date(base + 'T20:00:00Z'); // Noon or 1 p.m. Pacific, away from midnight.
    value.setUTCDate(value.getUTCDate() + days);
    return value;
  }
  return [
    {sheet: 'No Prior Recerts', values: ['Avery Example', 'avery@example.com', 'lead@example.com', 'BLS', expiry(180)]},
    {sheet: 'No Prior Recerts', values: ['Blair Example', 'blair@example.com', 'lead@example.com', 'RN', expiry(45)]},
    {sheet: 'One Prior Recert', values: ['Casey Example', 'casey@example.com', 'supervisor@example.com', 'ACLS', expiry(20)]},
    {sheet: 'One Prior Recert', values: ['Devon Example', 'devon@example.com', 'supervisor@example.com', 'BLS', expiry(90)], flagged: true},
    {sheet: 'Two+ Prior Recerts', values: ['Emery Example', 'emery@example.com', '', 'RN', expiry(240)]},
    {sheet: 'Two+ Prior Recerts', values: ['Finley Example', 'finley@example.com', '', 'BLS', expiry(-5)]},
    {sheet: 'Two+ Prior Recerts', values: ['Gray Example', 'not-an-email', '', 'ACLS', expiry(30)]},
    {sheet: 'Two+ Prior Recerts', values: ['Avery Example', 'avery@example.com', 'lead@example.com', 'BLS', expiry(180)]}
  ];
}
