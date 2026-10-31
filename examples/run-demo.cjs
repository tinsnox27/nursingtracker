const fs = require('fs');
const path = require('path');
const {environment} = require('./mock-apps-script.cjs');
const demo = environment();
const result = demo.run('runDemo()');
console.log('Synthetic portfolio demo. No Google account required. No email is sent.');
console.log('Dry-run preview:', JSON.stringify(result.summary));
const preview = demo.workbook().getSheetByName('Demo Preview').rows;
console.table(preview.slice(1).map(r => ({recipient:r[1], role:r[2], credential:r[3], stage:r[5]})));
// Switch only the in-memory mock environment to simulated sends to demonstrate persistence.
demo.run('CONFIG.DRY_RUN = false');
console.log('First simulated delivery run:', JSON.stringify(demo.run('checkAllCertificationsAndNotify()')));
console.log('Second simulated delivery run:', JSON.stringify(demo.run('checkAllCertificationsAndNotify()')));
console.log('Messages captured in memory:', demo.messages.length);
const output = path.join(__dirname, 'email-preview.html');
fs.writeFileSync(output, '<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>Example certification reminder</title></head><body><p style="font-family:Arial;text-align:center">Synthetic portfolio example. Generated locally; no email sent.</p>' + demo.messages[0].html + '</body></html>');
console.log('Open examples/email-preview.html for the generated message.');
