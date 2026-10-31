/* Run with Node.js. Mocks Apps Script services; never sends actual email. */
const fs = require('fs');
const vm = require('vm');
const assert = require('assert/strict');
const code = fs.readFileSync(__dirname + '/../src/Code.gs', 'utf8');
class Sheet {
 constructor(rows,colors){this.rows=rows;this.colors=colors;}
 getDataRange(){return {getValues:()=>this.rows, getBackgrounds:()=>this.rows.map((r,i)=>r.map((_,j)=>this.colors?.[i]?.[j] || '#ffffff'))};}
 appendRow(row){this.rows.push(row);}
 getLastRow(){return this.rows.length;}
 getRange(r,c){return {setValue:v=>{this.rows[r-1][c-1]=v;}};}
}
function setup({dry=false,failSupervisor=false,red=false,quota=50}={}){
 const now=new Date(), expiry=new Date(now);expiry.setUTCDate(expiry.getUTCDate()+90);
 const header=['Name','Email','Supervisor','Credential','Expiration'];
 const row=['A <B>','NURSE@example.com','boss@example.com','BLS & ACLS',expiry];
 const tabs={'No Prior Recerts':new Sheet([header,row],red?[[],['','#ff0000']]:null),'One Prior Recert':new Sheet([header]),'Two+ Prior Recerts':new Sheet([header])};
 let sent=[],released=0;
 const context=vm.createContext({Date,Map,Set,JSON,PropertiesService:{getScriptProperties:()=>({getProperty:()=>null})},Logger:{log:()=>{}},
 LockService:{getScriptLock:()=>({tryLock:()=>true,releaseLock:()=>released++})},
 SpreadsheetApp:{getActiveSpreadsheet:()=>({getSpreadsheetTimeZone:()=> 'UTC',getSheetByName:n=>tabs[n],insertSheet:n=>tabs[n]=new Sheet([])}),flush:()=>{}},
 Utilities:{formatDate:(d,tz,fmt)=>fmt==='yyyy-MM-dd'?d.toISOString().slice(0,10):d.toUTCString()},
 MailApp:{getRemainingDailyQuota:()=>quota},
 GmailApp:{sendEmail:(to,subject,body,opts)=>{if(failSupervisor&&to==='boss@example.com')throw Error('mock rejection');sent.push({to,subject,body,opts});}}
 });
 vm.runInContext(code,context);vm.runInContext('CONFIG.DRY_RUN = '+dry,context);
 return {context,tabs,sent,row,run:()=>vm.runInContext('checkAllCertificationsAndNotify()',context),released:()=>released};
}
let checks=0;function test(name,f){f();checks++;console.log('PASS '+name);}
test('Repeated run sends each recipient once',()=>{const x=setup();assert.equal(x.run().sent,2);assert.equal(x.run().sent,0);assert.equal(x.sent.length,2);assert.equal(x.released(),2);});
test('Renewal cycle creates a new delivery key',()=>{const x=setup();x.run();const next=new Date(x.row[4]);next.setUTCDate(next.getUTCDate()+1);x.row[4]=next;assert.equal(x.run().sent,2);});
test('Supervisor failure never resends successful nurse delivery',()=>{const x=setup({failSupervisor:true});assert.equal(x.run().sent,1);assert.equal(x.run().sent,0);assert.equal(x.sent.filter(s=>s.to==='nurse@example.com').length,1);assert.equal(x.tabs['Reminder Delivery Log v2'].rows[2][1],'ERROR');});
test('Dry run neither sends nor creates delivery log',()=>{const x=setup({dry:true});assert.equal(x.run().previewed,2);assert.equal(x.sent.length,0);assert.equal(x.tabs['Reminder Delivery Log v2'],undefined);});
test('Red nurse cell permits valid supervisor notice',()=>{const x=setup({red:true});assert.equal(x.run().sent,1);assert.equal(x.sent[0].to,'boss@example.com');});
test('Exact bounce colors do not classify magenta as red',()=>{const x=setup();assert.equal(vm.runInContext("isBounceback('#ff00ff')",x.context),false);});
test('Threshold boundaries and missed-run catch-up',()=>{const x=setup();for(const [d,want]of [[181,null],[180,'SIX_MONTHS'],[91,'SIX_MONTHS'],[90,'THREE_MONTHS'],[45,'THREE_MONTHS'],[30,'ONE_MONTH'],[0,'ONE_MONTH'],[-1,null]]){assert.equal(vm.runInContext(`determineThreshold(${d})`,x.context),want);}});
test('Quota exhaustion releases lock and defers unsent recipient',()=>{const x=setup({quota:0});assert.throws(x.run,/quota exhausted/);assert.equal(x.released(),1);assert.equal(x.sent.length,0);});
test('Blank date and malformed email are rejected',()=>{const x=setup();x.row[4]='';assert.equal(x.run().invalid,1);x.row[4]=new Date();x.row[1]='bad,email@example.com';assert.equal(x.run().invalid,1);});
test('HTML values are escaped and plain text is populated',()=>{const x=setup();x.run();assert.match(x.sent[0].opts.htmlBody,/A &lt;B&gt;/);assert.match(x.sent[0].opts.htmlBody,/BLS &amp; ACLS/);assert.ok(x.sent[0].body.length>50);});
test('PENDING records block blind retries after uncertain send',()=>{const x=setup();x.run();x.tabs['Reminder Delivery Log v2'].rows[1][1]='PENDING';assert.equal(x.run().blocked,1);assert.equal(x.sent.length,2);});
test('Duplicate source row does not send twice',()=>{const x=setup();x.tabs['One Prior Recert'].rows.push(x.row);assert.equal(x.run().sent,2);});
test('Missing tab aborts visibly and releases lock',()=>{const x=setup();delete x.tabs['One Prior Recert'];assert.throws(x.run,/Missing source tab/);assert.equal(x.sent.length,0);assert.equal(x.released(),1);});
test('Calendar day arithmetic remains stable across DST boundary',()=>{const x=setup();assert.equal(vm.runInContext("daysBetweenDates('2026-03-07','2026-03-09')",x.context),2);});
console.log(`${checks} checks passed. Local mocked services only; Apps Script validation still required.`);
