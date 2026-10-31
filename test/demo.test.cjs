const assert = require('assert/strict');
const {environment} = require('../examples/mock-apps-script.cjs');
let count=0;
function test(name,fn){fn();count++;console.log('PASS '+name);}
test('End-to-end demo creates all tabs and seven preview messages without sending',()=>{
 const e=environment();const r=e.run('runDemo()');
 assert.equal(e.workbook().getSheets().length,5);assert.equal(r.summary.previewed,7);assert.equal(r.summary.invalid,1);
 assert.equal(e.messages.length,0);assert.equal(e.workbook().getSheetByName('Demo Preview').rows.length,8);
 assert.equal(e.workbook().getSheetByName('Reminder Delivery Log v2'),null);
});
test('Preview reruns reuse the workbook and replace rather than append previews',()=>{
 const e=environment();e.run('runDemo()');e.run('previewDemoReminders()');assert.equal(e.books.size,1);assert.equal(e.workbook().getSheetByName('Demo Preview').rows.length,8);
});
test('Synthetic demo exercises all stages and suppresses flagged nurse and duplicate row',()=>{
 const e=environment();e.run('runDemo()');const rows=e.workbook().getSheetByName('Demo Preview').rows.slice(1);
 assert.deepEqual([...new Set(rows.map(r=>r[5]))].sort(),['ONE_MONTH','SIX_MONTHS','THREE_MONTHS']);
 assert.equal(rows.filter(r=>r[1]==='devon@example.com').length,0);assert.equal(rows.filter(r=>r[1]==='avery@example.com').length,1);
 assert.equal(rows.filter(r=>r[0]==='Devon Example'&&r[2]==='supervisor').length,1);
});
test('Simulated delivery persists seven records and suppresses the second run',()=>{
 const e=environment();e.run('runDemo()');e.run('CONFIG.DRY_RUN = false');assert.equal(e.run('checkAllCertificationsAndNotify()').sent,7);
 assert.equal(e.run('checkAllCertificationsAndNotify()').sent,0);assert.equal(e.messages.length,7);
 assert.equal(e.workbook().getSheetByName('Reminder Delivery Log v2').rows.length,8);
});
test('Demo refuses simulated live mode before creating another workbook',()=>{
 const e=environment();e.run('CONFIG.DRY_RUN = false');assert.throws(()=>e.run('runDemo()'),/DRY_RUN/);assert.equal(e.books.size,0);
});
console.log(`${count} demo integration checks passed using in-memory services.`);
