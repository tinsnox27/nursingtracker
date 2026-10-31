/**
 * Nursing certification reminder tracker - portfolio demonstration, version 2.0.
 * Recreates the workflow described in the original portfolio sample.
 * Defaults to dry run. Does not send email until DRY_RUN is explicitly false.
 * Fixed-day cadence: 180 / 90 / 30 days (approximate months).
 * Uses the spreadsheet timezone for calendar-day comparisons.
 * Red nurse email cell means manually flagged bounce/invalid; no inbox scanning.
 * Delivery log v2 is separate from the legacy log, which lacks expiration dates.
 * Public portfolio recreation with synthetic examples; no original records required.
 */
const CONFIG = {
  DRY_RUN: true,
  SHEETS_TO_SCAN: [
    {name: 'No Prior Recerts', priorRecerts: 0},
    {name: 'One Prior Recert', priorRecerts: 1},
    {name: 'Two+ Prior Recerts', priorRecerts: 2}
  ],
  SENT_LOG_SHEET: 'Reminder Delivery Log v2',
  COLUMNS: {NURSE_NAME: 0, NURSE_EMAIL: 1, SUPERVISOR_EMAIL: 2, CERT_TYPE: 3, EXPIRATION_DATE: 4},
  BOUNCEBACK_COLORS: ['#ff0000', '#f00', 'red'],
  SENDER_NAME: 'Nursing Certification Tracking',
  INFOGRAPHICS: {SIX_MONTHS: '', THREE_MONTHS: '', ONE_MONTH: ''}
};
const LOG_HEADERS = ['Delivery key', 'Status', 'Updated at', 'Recipient', 'Role', 'Credential', 'Expiration', 'Threshold', 'Source', 'Detail'];

function checkAllCertificationsAndNotify() {
  return runReminderCheck_();
}

function runReminderCheck_(previewRows) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) throw new Error('Another reminder run is active.');
  try {
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('TRACKER_SPREADSHEET_ID');
    const ss = spreadsheetId ? SpreadsheetApp.openById(spreadsheetId) : SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error('Bind this script to the tracking spreadsheet.');
    const tz = ss.getSpreadsheetTimeZone();
    const today = calendarDay(new Date(), tz);
    const log = CONFIG.DRY_RUN ? null : getOrCreateSentLog(ss);
    const states = loadDeliveryStates(log || ss.getSheetByName(CONFIG.SENT_LOG_SHEET));
    const summary = {sent: 0, previewed: 0, skipped: 0, invalid: 0, blocked: 0};
    // Missing source tabs are configuration errors, not successful empty runs.
    const sheets = CONFIG.SHEETS_TO_SCAN.map(c => {
      const sheet = ss.getSheetByName(c.name);
      if (!sheet) throw new Error('Missing source tab: ' + c.name);
      return {sheet, config: c};
    });
    for (const item of sheets) {
      const rows = item.sheet.getDataRange().getValues();
      const colors = item.sheet.getDataRange().getBackgrounds();
      for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        if (row.every(v => v === '')) continue;
        const c = CONFIG.COLUMNS;
        const nurse = normalizeEmail(row[c.NURSE_EMAIL]);
        const supervisor = normalizeEmail(row[c.SUPERVISOR_EMAIL]);
        const cert = String(row[c.CERT_TYPE] || '').trim();
        const expiry = row[c.EXPIRATION_DATE];
        const source = item.config.name + ':' + (i + 1);
        if (!validEmail(nurse) || !cert || !(expiry instanceof Date) || isNaN(expiry.getTime())) {
          summary.invalid++; Logger.log('Invalid required fields at ' + source); continue;
        }
        const expiration = calendarDay(expiry, tz);
        const threshold = determineThreshold(daysBetweenDates(today, expiration));
        if (!threshold) {summary.skipped++; continue;}
        const urgency = getUrgencyLevel(threshold);
        const formatted = Utilities.formatDate(expiry, tz, 'MMMM d, yyyy');
        urgency.message = 'Your certification expires on ' + formatted + ' (' + daysBetweenDates(today, expiration) + ' days from today).';
        const recipients = [];
        if (!isBounceback(colors[i][c.NURSE_EMAIL])) {
          recipients.push({email: nurse, role: 'nurse', body: buildNurseEmailBody(String(row[c.NURSE_NAME] || 'Colleague'), cert, formatted, urgency, item.config.priorRecerts)});
        } else {summary.skipped++;}
        // A flagged nurse does not suppress an otherwise valid supervisor notice.
        if (supervisor && supervisor !== nurse) {
          if (validEmail(supervisor) && !isBounceback(colors[i][c.SUPERVISOR_EMAIL])) {
            recipients.push({email: supervisor, role: 'supervisor', body: buildSupervisorEmailBody(String(row[c.NURSE_NAME] || 'Colleague'), cert, formatted, urgency)});
          } else {summary.invalid++;}
        }
        for (const recipient of recipients) {
          const key = deliveryKey(nurse, cert, expiration, threshold, recipient.email, recipient.role);
          const prior = states.get(key);
          if (prior) {
            if (prior.status === 'SENT' || prior.status === 'PREVIEW') summary.skipped++;
            else summary.blocked++; // PENDING/ERROR require reconciliation, never blind retry.
            continue;
          }
          if (CONFIG.DRY_RUN) {
            summary.previewed++; states.set(key, {status: 'PREVIEW'});
            if (previewRows) previewRows.push([String(row[c.NURSE_NAME]), recipient.email, recipient.role, cert, expiration, threshold, urgency.prefix + ' Certification expiring: ' + cert, plainText(recipient.body)]);
            Logger.log('Would send ' + recipient.role + ' reminder at ' + source + ' [' + threshold + ']');
            continue;
          }
          if (MailApp.getRemainingDailyQuota() < 1) {
            throw new Error('Email quota exhausted; remaining recipients deferred.');
          }
          const attachment = recipient.role === 'nurse' ? getInfographicAttachment(threshold) : null;
          // Persist intent before sending. A crash creates a reviewable PENDING record.
          const logRow = appendDelivery(log, [key, 'PENDING', new Date(), recipient.email, recipient.role, cert, expiration, threshold, source, '']);
          states.set(key, {status: 'PENDING', row: logRow});
          SpreadsheetApp.flush();
          try {
            GmailApp.sendEmail(recipient.email,
              (recipient.role === 'supervisor' ? '[Supervisor notice] ' : '') + urgency.prefix + ' Certification expiring: ' + cert,
              plainText(recipient.body), {
                htmlBody: recipient.body, name: CONFIG.SENDER_NAME,
                attachments: attachment ? [attachment] : []
              });
          } catch (error) {
            // Keep an uncertain send blocked even if the status update fails.
            log.getRange(logRow, 2).setValue('ERROR');
            log.getRange(logRow, 10).setValue('Review delivery before retry.');
            SpreadsheetApp.flush();
            summary.blocked++; continue;
          }
          // If this write fails, abort. PENDING prevents automatic duplicate sends.
          log.getRange(logRow, 2).setValue('SENT');
          SpreadsheetApp.flush();
          states.set(key, {status: 'SENT', row: logRow}); summary.sent++;
        }
      }
    }
    Logger.log(JSON.stringify(summary)); return summary;
  } finally {lock.releaseLock();}
}

function getOrCreateSentLog(ss) {
  let sheet = ss.getSheetByName(CONFIG.SENT_LOG_SHEET);
  if (!sheet) {sheet = ss.insertSheet(CONFIG.SENT_LOG_SHEET); sheet.appendRow(LOG_HEADERS);}
  validateLog(sheet); return sheet;
}
function validateLog(sheet) {
  const first = sheet.getDataRange().getValues()[0];
  if (!first || LOG_HEADERS.some((v, i) => first[i] !== v)) throw new Error('Unexpected delivery log schema.');
}
function loadDeliveryStates(sheet) {
  const states = new Map();
  if (!sheet) return states;
  validateLog(sheet);
  sheet.getDataRange().getValues().slice(1).forEach((r, i) => {
    if (!r[0]) return;
    if (!['SENT', 'PENDING', 'ERROR'].includes(r[1])) throw new Error('Unknown delivery status at row ' + (i + 2));
    // Any SENT record takes precedence over an older or duplicate pending record.
    if (!states.has(r[0]) || r[1] === 'SENT') states.set(r[0], {status: r[1], row: i + 2});
  });
  return states;
}
function appendDelivery(sheet, row) {
  sheet.appendRow(row); return sheet.getLastRow();
}
function normalizeEmail(value) {return String(value || '').trim().toLowerCase();}
function validEmail(value) {return /^[^\s@,;<>]+@[^\s@,;<>]+\.[^\s@,;<>]+$/.test(value);}
function isBounceback(color) {return CONFIG.BOUNCEBACK_COLORS.includes(String(color || '').toLowerCase());}
function calendarDay(date, tz) {return Utilities.formatDate(date, tz, 'yyyy-MM-dd');}
function daysBetweenDates(start, end) {
  return Math.round((Date.parse(end + 'T00:00:00Z') - Date.parse(start + 'T00:00:00Z')) / 86400000);
}
function determineThreshold(days) {
  // Catch up within the current band after missed runs; do not send all old stages.
  if (days < 0 || days > 180) return null;
  if (days <= 30) return 'ONE_MONTH';
  if (days <= 90) return 'THREE_MONTHS';
  return 'SIX_MONTHS';
}
function deliveryKey(nurse, cert, expiry, threshold, recipient, role) {
  return JSON.stringify([normalizeEmail(nurse), cert.trim().toLowerCase(), expiry, threshold, normalizeEmail(recipient), role]);
}
function escapeHtml(value) {
  return String(value).replace(/[&<>"']/g, c => ({'&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;'}[c]));
}
function plainText(html) {
  return html.replace(/<[^>]*>/g, ' ').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/\s+/g, ' ').trim();
}
function createDailyTrigger() {
  if (CONFIG.DRY_RUN) throw new Error('Validate the preview and set DRY_RUN to false first.');
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'checkAllCertificationsAndNotify') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('checkAllCertificationsAndNotify').timeBased().atHour(7).everyDays(1).create();
  Logger.log('Daily trigger created for the 7 a.m. hour in the project timezone.');
}

function getUrgencyLevel(threshold) {
  const levels = {
    SIX_MONTHS: {
      prefix: '[Upcoming]',
      color: '#2e7d32',  // Green
      message: 'You have plenty of time to complete your recertification.',
      action: 'Start planning your recertification process.'
    },
    THREE_MONTHS: {
      prefix: '[Action Needed]',
      color: '#f57c00',  // Orange
      message: 'Your certification will expire in approximately 3 months.',
      action: 'Please begin your recertification process soon.'
    },
    ONE_MONTH: {
      prefix: '[URGENT]',
      color: '#c62828',  // Red
      message: 'Your certification expires in approximately 1 month!',
      action: 'Immediate action is required to avoid a lapse in certification.'
    }
  };
  return levels[threshold];
}
 
/**
 * Get recertification guidance based on how many times they've recertified before
 */
function getRecertGuidance(priorRecerts) {
  return {
    header: priorRecerts === 0 ? 'First renewal' : 'Renewal planning',
    message: 'Verify the current renewal requirements with the relevant certification issuer. Requirements vary by credential; this reminder does not establish eligibility or renew the credential.'
  };
}

/**
 * Build personalized HTML email body for nursing staff
 */
function buildNurseEmailBody(name, certType, expirationDate, urgency, priorRecerts) {
  const recertGuidance = getRecertGuidance(priorRecerts);
  name = escapeHtml(name); certType = escapeHtml(certType);
  
  return `
    <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
      <div style="background-color: ${urgency.color}; color: white; padding: 20px; text-align: center;">
        <h1 style="margin: 0;">Nursing Certification Reminder</h1>
      </div>
      
      <div style="padding: 20px; background-color: #f5f5f5;">
        <p>Dear <strong>${name}</strong>,</p>
        
        <p>${urgency.message}</p>
        
        <table style="width: 100%; border-collapse: collapse; margin: 20px 0;">
          <tr>
            <td style="padding: 10px; border: 1px solid #ddd; background-color: white;">
              <strong>Certification:</strong>
            </td>
            <td style="padding: 10px; border: 1px solid #ddd; background-color: white;">
              ${certType}
            </td>
          </tr>
          <tr>
            <td style="padding: 10px; border: 1px solid #ddd; background-color: white;">
              <strong>Expiration Date:</strong>
            </td>
            <td style="padding: 10px; border: 1px solid #ddd; background-color: white;">
              ${expirationDate}
            </td>
          </tr>
        </table>
        
        <div style="background-color: #e3f2fd; padding: 15px; border-left: 4px solid #1565c0; margin: 20px 0;">
          <strong>${recertGuidance.header}</strong><br>
          ${recertGuidance.message}
        </div>
        
        <p style="background-color: #fff3cd; padding: 15px; border-left: 4px solid ${urgency.color};">
          <strong>Next Step:</strong> ${urgency.action}
        </p>
        
        <p>If a guide is attached, use it alongside the current requirements from the relevant issuer. Contact the program coordinator if no guide is attached.</p>
        
        <p>If you have questions, please contact the Staff Development Office or your Nurse Manager.</p>
        
        <p>Thank you,<br>
        <em>Nursing Certification Tracking</em></p>
      </div>
      
      <div style="background-color: #333; color: white; padding: 10px; text-align: center; font-size: 12px;">
        This is an automated message. Please do not reply directly to this email.
      </div>
    </div>
  `;
}
 
/**
 * Build HTML email body for supervisor notification
 */
function buildSupervisorEmailBody(nurseName, certType, expirationDate, urgency) {
  nurseName = escapeHtml(nurseName); certType = escapeHtml(certType);
  return `
    <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
      <div style="background-color: #1565c0; color: white; padding: 20px; text-align: center;">
        <h1 style="margin: 0;">Nurse Manager Notification</h1>
      </div>
      
      <div style="padding: 20px; background-color: #f5f5f5;">
        <p>This is a courtesy notification that a nursing staff member's certification is approaching expiration.</p>
        
        <table style="width: 100%; border-collapse: collapse; margin: 20px 0;">
          <tr>
            <td style="padding: 10px; border: 1px solid #ddd; background-color: white;">
              <strong>Staff Member:</strong>
            </td>
            <td style="padding: 10px; border: 1px solid #ddd; background-color: white;">
              ${nurseName}
            </td>
          </tr>
          <tr>
            <td style="padding: 10px; border: 1px solid #ddd; background-color: white;">
              <strong>Certification:</strong>
            </td>
            <td style="padding: 10px; border: 1px solid #ddd; background-color: white;">
              ${certType}
            </td>
          </tr>
          <tr>
            <td style="padding: 10px; border: 1px solid #ddd; background-color: white;">
              <strong>Expiration Date:</strong>
            </td>
            <td style="padding: 10px; border: 1px solid #ddd; background-color: white;">
              ${expirationDate}
            </td>
          </tr>
          <tr>
            <td style="padding: 10px; border: 1px solid #ddd; background-color: white;">
              <strong>Urgency:</strong>
            </td>
            <td style="padding: 10px; border: 1px solid #ddd; background-color: white; color: ${urgency.color};">
              <strong>${urgency.prefix}</strong>
            </td>
          </tr>
        </table>
        
        <p>Please confirm the staff member's reminder status in the delivery log. Follow up if coverage planning 
        is needed during any required training or if there are concerns about timely renewal.</p>
        
        <p>Thank you,<br>
        <em>Nursing Certification Tracking</em></p>
      </div>
    </div>
  `;
}
 
/**
 * Get infographic attachment from Google Drive
 */
function getInfographicAttachment(threshold) {
  try {
    const fileId = CONFIG.INFOGRAPHICS[threshold];
    if (!fileId || fileId.includes('YOUR_')) {
      return null; // Placeholder ID, skip attachment
    }
    const file = DriveApp.getFileById(fileId);
    return file.getBlob();
  } catch (error) {
    Logger.log(`Could not attach infographic: ${error.message}`);
    return null;
  }
}
 
