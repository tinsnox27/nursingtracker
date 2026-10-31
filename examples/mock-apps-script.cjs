/* Minimal in-memory Apps Script adapter. No Google connection or email delivery. */
const fs = require('fs');
const path = require('path');
const vm = require('vm');
class Sheet {
  constructor(name, rows = []) { this.name = name; this.rows = rows; this.colors = new Map(); }
  getName() { return this.name; }
  setName(name) { this.name = name; return this; }
  getDataRange() {
    const width = Math.max(1, ...this.rows.map(r => r.length));
    return {
      getValues: () => this.rows.length ? this.rows.map(r => Array.from({length: width}, (_, i) => r[i] ?? '')) : [['']],
      getBackgrounds: () => this.rows.map((r, i) => Array.from({length: width}, (_, j) => this.colors.get(`${i + 1}:${j + 1}`) || '#ffffff'))
    };
  }
  getRange(row, col, count = 1, width = 1) {
    const self = this;
    const range = {
      setValues(values) { for (let i = 0; i < count; i++) for (let j = 0; j < width; j++) { self.rows[row + i - 1] ||= []; self.rows[row + i - 1][col + j - 1] = values[i][j]; } return range; },
      setValue(value) { return range.setValues([[value]]); },
      setBackground(color) { self.colors.set(`${row}:${col}`, color); return range; },
      setFontWeight() { return range; }, setNumberFormat() { return range; }, setWrap() { return range; }
    };
    return range;
  }
  appendRow(row) { this.rows.push(row); }
  getLastRow() { return this.rows.length; }
  clearContents() { this.rows = []; return this; }
  setColumnWidth() {} setFrozenRows() {} autoResizeColumns() {}
}
class Workbook {
  constructor(id) { this.id = id; this.sheets = [new Sheet('Sheet1')]; this.timeZone = 'UTC'; }
  getId() { return this.id; }
  getUrl() { return `mock-workbook://${this.id}`; }
  getSheets() { return this.sheets; }
  getSheetByName(name) { return this.sheets.find(s => s.name === name) || null; }
  insertSheet(name) { const s = new Sheet(name); this.sheets.push(s); return s; }
  setSpreadsheetTimeZone(tz) { this.timeZone = tz; }
  getSpreadsheetTimeZone() { return this.timeZone; }
}
function environment() {
  const properties = new Map(), books = new Map(), messages = [], logs = [];
  let id = 0;
  const context = vm.createContext({
    Date, Map, Set, JSON,
    Logger: {log: s => logs.push(String(s))},
    PropertiesService: {getScriptProperties: () => ({getProperty: key => properties.get(key) || null, setProperty: (key, value) => properties.set(key, value)})},
    LockService: {getScriptLock: () => ({tryLock: () => true, releaseLock: () => {}})},
    SpreadsheetApp: {
      create: () => {const b = new Workbook(`demo-${++id}`); books.set(b.id, b); return b;},
      openById: key => {if (!books.has(key)) throw Error('Unknown mock workbook'); return books.get(key);},
      getActiveSpreadsheet: () => null, flush: () => {}
    },
    Utilities: {formatDate: (d, tz, fmt) => {
      const parts = new Intl.DateTimeFormat('en-US', {timeZone:tz, year:'numeric',month:'2-digit',day:'2-digit'}).formatToParts(d);
      const get = type => parts.find(p => p.type === type).value;
      return fmt === 'yyyy-MM-dd' ? `${get('year')}-${get('month')}-${get('day')}` : new Intl.DateTimeFormat('en-US',{timeZone:tz,year:'numeric',month:'long',day:'numeric'}).format(d);
    }},
    MailApp: {getRemainingDailyQuota: () => 100},
    GmailApp: {sendEmail: (recipient, subject, body, options) => messages.push({recipient, subject, body, html: options.htmlBody})}
  });
  for (const name of ['Code.gs', 'Demo.gs']) vm.runInContext(fs.readFileSync(path.join(__dirname, '../src', name), 'utf8'), context, {filename:name});
  return {context, books, properties, messages, logs, run: script => vm.runInContext(script, context), workbook: () => books.get(properties.get('TRACKER_SPREADSHEET_ID'))};
}
module.exports = {environment};
