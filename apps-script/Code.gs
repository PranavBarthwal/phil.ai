/**
 * COLD DM AUTOMATION
 *
 * Bound to: a Google Sheet (tab CONFIG.SHEET_NAME)
 * Reads templates from: a Google Doc (CONFIG.DOC_ID)
 * Sends via: GmailApp
 * Optional first-email drafts: Groq (key stored in Script Properties, not in this file)
 *
 * SHEET COLUMNS (row 1 = header, any order). Run the menu item
 * "Add missing sheet columns" once to append any new ones:
 *   Email | CC | BCC | Name | Company | Role
 *   Template | Subject Choice | Generate Context
 *   Status | Sequence Step | Scheduled Time | Sent Time
 *   Next Follow-up | Thread ID | Original Subject | Reply | Error
 *
 * DOC FORMAT — multiple templates in one doc:
 *   Template: default
 *   Subject: first option
 *   Subject: second option
 *   Body:
 *   ...email body, bold/links/bullets preserved...
 *
 *   Template: xyz_experience
 *   ...
 *
 *   Template: followup_1
 *   ...
 *
 *   Template: followup_2
 *   ...
 *
 * FIRST EMAIL
 *   Template column          → which Template: block to use (blank = CONFIG.DEFAULT_TEMPLATE_ID)
 *   Subject Choice column    → 1-based index of that template's Subject: lines (blank = 1)
 *   Generate Context column  → if filled, Groq writes the first email using this context
 *                              (and the selected template as a style reference, if present)
 *
 * SEQUENCE (no reply → two fixed follow-ups)
 *   Step 1 = first email (chosen template or Groq)
 *   Step 2 = Template: followup_1  (after FOLLOWUP_DELAY_DAYS[0])
 *   Step 3 = Template: followup_2  (after FOLLOWUP_DELAY_DAYS[1])
 *   A reply from any To/CC address stops the sequence (Reply = Yes, Status = Replied).
 *
 * RECIPIENTS
 *   Email / CC / BCC accept comma- or semicolon-separated addresses.
 *
 * ONE-TIME SETUP
 * 1. Paste your Google Doc ID into CONFIG.DOC_ID.
 * 2. Apps Script editor → Services (+) → add "Gmail API".
 * 3. File → Project settings → Script timezone = your sheet timezone.
 * 4. Menu "Cold DM Automation" → "Add missing sheet columns".
 * 5. Menu → "Save Groq API key" (only if you want Generate Context).
 * 6. In the editor, run setupTriggers once and authorize.
 *
 * Logs: Apps Script editor → Executions.
 */
// ============ CONFIG — EDIT THESE ============
const CONFIG = {
  SHEET_NAME: 'Sheet1',
  DOC_ID: '1xMLNox-lwyIfDFRV4O4NymzEvijemajssSAe8BTH0Uo',
  SENDER_NAME: 'PRANAV BARTHWAL',

  SEND_DAYS: [0, 1, 2, 3, 4, 5, 6],
  START_HOUR: 10,
  END_HOUR: 13,
  INTERVAL_MINUTES: 12,

  SCHEDULE_TRIGGER_MINUTES: 30,
  SEND_TRIGGER_MINUTES: 15,

  APPEND_GMAIL_SIGNATURE: true,

  DEFAULT_TEMPLATE_ID: 'default',
  FOLLOWUP_TEMPLATE_IDS: ['followup_1', 'followup_2'],
  FOLLOWUP_DELAY_DAYS: [3, 7],

  GROQ_MODEL: 'llama-3.3-70b-versatile',
  GROQ_API_URL: 'https://api.groq.com/openai/v1/chat/completions'
};

const REQUIRED_HEADERS = [
  'Email', 'CC', 'BCC', 'Name', 'Company', 'Role',
  'Template', 'Subject Choice', 'Generate Context',
  'Status', 'Sequence Step', 'Scheduled Time', 'Sent Time',
  'Next Follow-up', 'Thread ID', 'Original Subject', 'Reply', 'Error'
];
// ================================================

function log(step, msg) {
  Logger.log('[' + step + '] ' + msg);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Cold DM Automation')
    .addItem('Force-send test email now', 'forceSendTest')
    .addItem('Save Groq API key', 'saveGroqApiKey')
    .addItem('Add missing sheet columns', 'ensureHeadersMenu')
    .addItem('Run diagnostics', 'runDiagnostics')
    .addItem('Reset error rows', 'resetErrors')
    .addToUi();
}

/**
 * ONE-TIME: run this from the Apps Script editor to install background jobs.
 * Safe to re-run — it clears old triggers first.
 */
function setupTriggers() {
  log('SETUP', '--- setupTriggers started ---');
  const fns = ['scheduleEmails', 'sendScheduledEmails'];
  fns.forEach(fn => {
    ScriptApp.getProjectTriggers().forEach(t => {
      if (t.getHandlerFunction() === fn) ScriptApp.deleteTrigger(t);
    });
  });
  ScriptApp.newTrigger('scheduleEmails').timeBased().everyMinutes(CONFIG.SCHEDULE_TRIGGER_MINUTES).create();
  ScriptApp.newTrigger('sendScheduledEmails').timeBased().everyMinutes(CONFIG.SEND_TRIGGER_MINUTES).create();
  log('SETUP', 'Installed: scheduleEmails every ' + CONFIG.SCHEDULE_TRIGGER_MINUTES +
    ' min, sendScheduledEmails every ' + CONFIG.SEND_TRIGGER_MINUTES + ' min');
  log('SETUP', '--- setupTriggers finished ---');
}

function getSheetContext() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) throw new Error('Sheet tab "' + CONFIG.SHEET_NAME + '" not found.');
  ensureHeaders(sheet);
  const data = sheet.getDataRange().getValues();
  const col = mapColumns(data[0]);
  if (col.email === undefined) throw new Error('Sheet is missing an "Email" header.');
  if (col.status === undefined) throw new Error('Sheet is missing a "Status" header.');
  return { sheet: sheet, data: data, col: col };
}

/** Runs automatically — assigns a send slot to Pending rows (new leads + due follow-ups). */
function scheduleEmails() {
  log('SCHEDULE', '--- scheduleEmails started ---');
  try {
    const ctx = getSheetContext();
    markReplies(ctx);
    queueDueFollowUps(ctx);
    ctx.data = ctx.sheet.getDataRange().getValues();

    let lastSlot = getLatestScheduledTime(ctx.data, ctx.col);
    log('SCHEDULE', 'Starting from slot base: ' + lastSlot);
    let count = 0;

    for (let r = 1; r < ctx.data.length; r++) {
      const row = ctx.data[r];
      const email = joinAddresses(parseAddressList(row[ctx.col.email]));
      const status = normalizeStatus(row[ctx.col.status]);

      if (!email) continue;
      if (status && status !== 'pending') continue;

      const candidate = new Date(lastSlot.getTime() + CONFIG.INTERVAL_MINUTES * 60000);
      const slot = nextValidSlot(candidate);
      const step = getSequenceStep(row, ctx.col);
      log('SCHEDULE', 'Row ' + (r + 1) + ' (' + email + '): step ' + step + ' slot ' + slot);

      ctx.sheet.getRange(r + 1, ctx.col.scheduled + 1).setValue(slot);
      ctx.sheet.getRange(r + 1, ctx.col.status + 1).setValue('Scheduled');
      if (ctx.col.step !== undefined) ctx.sheet.getRange(r + 1, ctx.col.step + 1).setValue(step);
      lastSlot = slot;
      count++;
    }
    SpreadsheetApp.flush();
    log('SCHEDULE', '--- scheduleEmails finished: ' + count + ' row(s) newly scheduled ---');
  } catch (e) {
    log('SCHEDULE', 'ERROR: ' + e.message);
  }
}

/** Runs automatically — sends due Scheduled rows, but only inside the send window. */
function sendScheduledEmails() {
  const now = new Date();
  const day = now.getDay();
  const hour = now.getHours();
  const minute = now.getMinutes();

  if (CONFIG.SEND_DAYS.indexOf(day) === -1 || hour < CONFIG.START_HOUR || hour >= CONFIG.END_HOUR) {
    log('SEND', 'Outside send window (' + CONFIG.START_HOUR + ':00–' + CONFIG.END_HOUR + ':00). Current: ' +
      ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'][day] + ' ' + hour + ':' + String(minute).padStart(2, '0') + '. Skipping.');
    return;
  }

  log('SEND', '--- sendScheduledEmails started ---');
  const summary = { scanned: 0, due: 0, sent: 0, errors: 0, skippedReply: 0 };
  try {
    if (!CONFIG.DOC_ID || CONFIG.DOC_ID.indexOf('PASTE_YOUR') === 0) {
      throw new Error('CONFIG.DOC_ID is still the placeholder.');
    }
    const ctx = getSheetContext();
    markReplies(ctx);
    ctx.data = ctx.sheet.getDataRange().getValues();

    const templates = loadTemplates();
    const signature = CONFIG.APPEND_GMAIL_SIGNATURE ? getGmailSignature() : '';

    for (let r = 1; r < ctx.data.length; r++) {
      summary.scanned++;
      const row = ctx.data[r];
      if (normalizeStatus(row[ctx.col.status]) !== 'scheduled') continue;

      const scheduledTime = row[ctx.col.scheduled];
      if (!scheduledTime || new Date(scheduledTime) > now) continue;
      summary.due++;

      if (hasReply(row, ctx.col)) {
        log('SEND', 'Row ' + (r + 1) + ': reply already present, skipping send');
        setReplyStopped(ctx.sheet, r, ctx.col);
        summary.skippedReply++;
        continue;
      }

      const to = joinAddresses(parseAddressList(row[ctx.col.email]));
      if (!to) continue;
      log('SEND', 'Row ' + (r + 1) + ' (' + to + '): due now, sending...');

      try {
        const built = buildEmailForRow(row, ctx.col, templates, signature);
        log('SEND', 'Row ' + (r + 1) + ': step=' + built.step + ' template=' + built.templateId + ' subject="' + built.subject + '"');
        const threadId = deliverEmail(row, ctx.col, built, to);
        afterSuccessfulSend(ctx.sheet, r, ctx.col, built, threadId, now);
        summary.sent++;
      } catch (err) {
        log('SEND', 'Row ' + (r + 1) + ': ERROR — ' + err.message);
        ctx.sheet.getRange(r + 1, ctx.col.status + 1).setValue('Error');
        if (ctx.col.error !== undefined) ctx.sheet.getRange(r + 1, ctx.col.error + 1).setValue(err.message);
        summary.errors++;
      }
    }
    log('SEND', '--- sendScheduledEmails finished --- ' + JSON.stringify(summary));
  } catch (e) {
    log('SEND', 'FATAL ERROR: ' + e.message);
  }
}

/**
 * Manual test only — sends immediately, ignores the day/time window.
 * Does not change Status / Scheduled Time / Sent Time / sequence.
 */
function forceSendTest() {
  const ui = SpreadsheetApp.getUi();
  log('TEST', '--- forceSendTest started ---');
  try {
    const ctx = getSheetContext();
    const resp = ui.prompt(
      'Force-send test email',
      'Enter the sheet row number to send (e.g. 2). Ignores the send window and does not change Status.',
      ui.ButtonSet.OK_CANCEL
    );
    if (resp.getSelectedButton() !== ui.Button.OK) return;

    const rowNum = parseInt(resp.getResponseText().trim(), 10);
    if (isNaN(rowNum) || rowNum < 2 || rowNum > ctx.data.length) {
      throw new Error('Invalid row number: ' + resp.getResponseText());
    }
    const row = ctx.data[rowNum - 1];
    const to = joinAddresses(parseAddressList(row[ctx.col.email]));
    if (!to) throw new Error('Row ' + rowNum + ' has no email address.');
    if (!CONFIG.DOC_ID || CONFIG.DOC_ID.indexOf('PASTE_YOUR') === 0) {
      throw new Error('CONFIG.DOC_ID is still the placeholder.');
    }

    const templates = loadTemplates();
    const signature = CONFIG.APPEND_GMAIL_SIGNATURE ? getGmailSignature() : '';
    const built = buildEmailForRow(row, ctx.col, templates, signature);
    const testSubject = '[TEST] ' + built.subject;
    const options = mailOptions(row, ctx.col, built.htmlBody);
    log('TEST', 'Sending test email to ' + to + (options.cc ? ' cc=' + options.cc : '') + (options.bcc ? ' bcc=' + options.bcc : ''));
    GmailApp.sendEmail(to, testSubject, built.plainBody, options);
    log('TEST', 'Test email sent to ' + to);
    ui.alert('Test email sent to ' + to + '.\n\nSubject: ' + testSubject + '\n\n(Status/Scheduled Time untouched.)');
  } catch (e) {
    log('TEST', 'ERROR: ' + e.message);
    ui.alert('Test send failed: ' + e.message);
  }
}

function saveGroqApiKey() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt(
    'Groq API key',
    'Paste your Groq API key. It is stored in Script Properties (not in the sheet or this file).',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const key = resp.getResponseText().trim();
  if (!key) {
    ui.alert('No key entered.');
    return;
  }
  PropertiesService.getScriptProperties().setProperty('GROQ_API_KEY', key);
  ui.alert('Groq API key saved.');
}

function ensureHeadersMenu() {
  const ctx = getSheetContext();
  SpreadsheetApp.getUi().alert('Sheet columns are ready. Header count: ' + ctx.data[0].length);
}

function resetErrors() {
  const ctx = getSheetContext();
  let count = 0;
  for (let r = 1; r < ctx.data.length; r++) {
    if (normalizeStatus(ctx.data[r][ctx.col.status]) !== 'error') continue;
    ctx.sheet.getRange(r + 1, ctx.col.status + 1).setValue('Pending');
    if (ctx.col.error !== undefined) ctx.sheet.getRange(r + 1, ctx.col.error + 1).setValue('');
    if (ctx.col.scheduled !== undefined) ctx.sheet.getRange(r + 1, ctx.col.scheduled + 1).setValue('');
    count++;
  }
  log('RESET', count + ' row(s) reset to Pending');
  SpreadsheetApp.getUi().alert(count + ' error row(s) reset to Pending.');
}

function runDiagnostics() {
  const ui = SpreadsheetApp.getUi();
  const report = [];
  try {
    const ctx = getSheetContext();
    report.push('✅ Sheet "' + CONFIG.SHEET_NAME + '" opened with ' + (ctx.data.length - 1) + ' data row(s).');
  } catch (e) {
    report.push('❌ Sheet: ' + e.message);
  }

  if (!CONFIG.DOC_ID || CONFIG.DOC_ID.indexOf('PASTE_YOUR') === 0) {
    report.push('❌ DOC_ID is still the placeholder.');
  } else {
    try {
      const templates = loadTemplates();
      const names = Object.keys(templates);
      report.push('✅ Doc templates: ' + (names.length ? names.join(', ') : '(none parsed)'));
      CONFIG.FOLLOWUP_TEMPLATE_IDS.forEach(id => {
        report.push(templates[id] ? '✅ Follow-up template "' + id + '" found' : '❌ Missing follow-up template "' + id + '"');
      });
    } catch (e) {
      report.push('❌ Template parse failed: ' + e.message);
    }
  }

  try {
    const sig = getGmailSignature();
    report.push(sig ? '✅ Gmail signature found (' + sig.length + ' chars)' : '⚠️ No signature found / Gmail API service not added yet.');
  } catch (e) {
    report.push('❌ Signature check failed: ' + e.message);
  }

  const groqKey = PropertiesService.getScriptProperties().getProperty('GROQ_API_KEY');
  report.push(groqKey ? '✅ Groq API key is saved in Script Properties' : '⚠️ No Groq API key (Generate Context will fail until you save one).');

  const triggers = ScriptApp.getProjectTriggers().map(t => t.getHandlerFunction());
  report.push('Installed triggers: ' + (triggers.length ? triggers.join(', ') : 'NONE — run setupTriggers from the editor.'));
  ui.alert('Diagnostics', report.join('\n\n'), ui.ButtonSet.OK);
}

// ============ SEQUENCE ============

function markReplies(ctx) {
  for (let r = 1; r < ctx.data.length; r++) {
    const row = ctx.data[r];
    const status = normalizeStatus(row[ctx.col.status]);
    if (status === 'replied' || status === 'completed' || status === 'pending') continue;
    if (!hasReply(row, ctx.col)) continue;
    log('REPLY', 'Row ' + (r + 1) + ': reply detected, stopping sequence');
    setReplyStopped(ctx.sheet, r, ctx.col);
    ctx.data[r][ctx.col.status] = 'Replied';
    if (ctx.col.reply !== undefined) ctx.data[r][ctx.col.reply] = 'Yes';
  }
}

function setReplyStopped(sheet, r, col) {
  sheet.getRange(r + 1, col.status + 1).setValue('Replied');
  if (col.reply !== undefined) sheet.getRange(r + 1, col.reply + 1).setValue('Yes');
  if (col.scheduled !== undefined) sheet.getRange(r + 1, col.scheduled + 1).setValue('');
  if (col.nextFollowup !== undefined) sheet.getRange(r + 1, col.nextFollowup + 1).setValue('');
}

function queueDueFollowUps(ctx) {
  const now = new Date();
  const maxStep = 1 + CONFIG.FOLLOWUP_TEMPLATE_IDS.length;
  let count = 0;
  for (let r = 1; r < ctx.data.length; r++) {
    const row = ctx.data[r];
    if (normalizeStatus(row[ctx.col.status]) !== 'sent') continue;
    if (hasReply(row, ctx.col)) continue;

    const step = getSequenceStep(row, ctx.col);
    if (step >= maxStep) {
      ctx.sheet.getRange(r + 1, ctx.col.status + 1).setValue('Completed');
      ctx.data[r][ctx.col.status] = 'Completed';
      continue;
    }

    const nextAt = ctx.col.nextFollowup !== undefined ? row[ctx.col.nextFollowup] : '';
    if (!nextAt || new Date(nextAt) > now) continue;

    const nextStep = step + 1;
    log('SEQUENCE', 'Row ' + (r + 1) + ': queueing follow-up step ' + nextStep);
    ctx.sheet.getRange(r + 1, ctx.col.status + 1).setValue('Pending');
    if (ctx.col.step !== undefined) ctx.sheet.getRange(r + 1, ctx.col.step + 1).setValue(nextStep);
    if (ctx.col.scheduled !== undefined) ctx.sheet.getRange(r + 1, ctx.col.scheduled + 1).setValue('');
    if (ctx.col.nextFollowup !== undefined) ctx.sheet.getRange(r + 1, ctx.col.nextFollowup + 1).setValue('');
    ctx.data[r][ctx.col.status] = 'Pending';
    if (ctx.col.step !== undefined) ctx.data[r][ctx.col.step] = nextStep;
    count++;
  }
  if (count) log('SEQUENCE', 'Queued ' + count + ' follow-up(s)');
}

function afterSuccessfulSend(sheet, r, col, built, threadId, now) {
  const maxStep = 1 + CONFIG.FOLLOWUP_TEMPLATE_IDS.length;
  sheet.getRange(r + 1, col.sent + 1).setValue(now);
  if (col.step !== undefined) sheet.getRange(r + 1, col.step + 1).setValue(built.step);
  if (col.thread !== undefined && threadId) sheet.getRange(r + 1, col.thread + 1).setValue(threadId);
  if (col.originalSubject !== undefined && built.step === 1) {
    sheet.getRange(r + 1, col.originalSubject + 1).setValue(built.subject);
  }
  if (built.step >= maxStep) {
    sheet.getRange(r + 1, col.status + 1).setValue('Completed');
    if (col.nextFollowup !== undefined) sheet.getRange(r + 1, col.nextFollowup + 1).setValue('');
    log('SEND', 'Row ' + (r + 1) + ': sequence complete');
    return;
  }
  const delayDays = CONFIG.FOLLOWUP_DELAY_DAYS[built.step - 1] || CONFIG.FOLLOWUP_DELAY_DAYS[CONFIG.FOLLOWUP_DELAY_DAYS.length - 1];
  const nextAt = new Date(now.getTime() + delayDays * 24 * 60 * 60 * 1000);
  sheet.getRange(r + 1, col.status + 1).setValue('Sent');
  if (col.nextFollowup !== undefined) sheet.getRange(r + 1, col.nextFollowup + 1).setValue(nextAt);
  log('SEND', 'Row ' + (r + 1) + ': next follow-up at ' + nextAt);
}

function hasReply(row, col) {
  if (col.reply !== undefined && String(row[col.reply] || '').trim().toLowerCase() === 'yes') return true;
  const threadId = col.thread !== undefined ? String(row[col.thread] || '').trim() : '';
  if (!threadId) return false;
  try {
    const thread = GmailApp.getThreadById(threadId);
    const recipients = parseAddressList(row[col.email])
      .concat(col.cc !== undefined ? parseAddressList(row[col.cc]) : [])
      .map(function (e) { return e.toLowerCase(); });
    const messages = thread.getMessages();
    for (let i = 0; i < messages.length; i++) {
      const from = (messages[i].getFrom() || '').toLowerCase();
      for (let j = 0; j < recipients.length; j++) {
        if (recipients[j] && from.indexOf(recipients[j]) !== -1) return true;
      }
    }
  } catch (e) {
    log('REPLY', 'Could not inspect thread ' + threadId + ': ' + e.message);
  }
  return false;
}

function getSequenceStep(row, col) {
  if (col.step === undefined) return 1;
  const n = parseInt(row[col.step], 10);
  return isNaN(n) || n < 1 ? 1 : n;
}

// ============ BUILD + SEND ============

function loadTemplates() {
  const doc = DocumentApp.openById(CONFIG.DOC_ID);
  const parsed = parseAllTemplates(doc);
  const names = Object.keys(parsed);
  log('PARSE', 'Loaded templates: ' + names.join(', '));
  if (!names.length) throw new Error('No templates found in the Google Doc.');
  return parsed;
}

function buildEmailForRow(row, col, templates, signature) {
  const step = getSequenceStep(row, col);
  const dataMap = rowDataMap(row, col);
  let templateId;
  let template;
  let generated = null;

  if (step === 1) {
    const context = col.context !== undefined ? String(row[col.context] || '').trim() : '';
    templateId = chosenTemplateId(row, col, templates);
    template = templates[templateId] || null;
    if (context) {
      generated = generateWithGroq(context, dataMap, template);
    } else if (!template) {
      throw new Error('Unknown template "' + templateId + '". Known: ' + Object.keys(templates).join(', '));
    }
  } else {
    templateId = CONFIG.FOLLOWUP_TEMPLATE_IDS[step - 2];
    if (!templateId) throw new Error('No follow-up template configured for sequence step ' + step);
    template = templates[templateId];
    if (!template) throw new Error('Follow-up template "' + templateId + '" is missing from the Google Doc.');
  }

  const subjects = generated ? generated.subjects : (template.subjects || []);
  const bodyHtmlSource = generated ? generated.bodyHtml : (template.bodyHtml || '');
  const subject = personalize(pickSubject(subjects, row, col), dataMap);
  if (!subject) throw new Error('No subject line for template "' + templateId + '". Add Subject: lines in the doc or a Subject Choice.');
  let htmlBody = personalize(bodyHtmlSource, dataMap);
  if (signature) htmlBody += '<br><br>' + signature;
  const plainBody = htmlToPlain(htmlBody);
  return {
    step: step,
    templateId: templateId,
    subject: subject,
    htmlBody: htmlBody,
    plainBody: plainBody
  };
}

function chosenTemplateId(row, col, templates) {
  const raw = col.template !== undefined ? String(row[col.template] || '').trim().toLowerCase() : '';
  if (raw) return raw;
  if (templates && templates[CONFIG.DEFAULT_TEMPLATE_ID]) return CONFIG.DEFAULT_TEMPLATE_ID;
  const skip = {};
  CONFIG.FOLLOWUP_TEMPLATE_IDS.forEach(function (id) { skip[id] = true; });
  const first = Object.keys(templates || {}).filter(function (name) { return !skip[name]; })[0];
  return first || CONFIG.DEFAULT_TEMPLATE_ID;
}

function pickSubject(subjects, row, col) {
  const list = (subjects || []).filter(function (s) { return String(s || '').trim(); });
  if (!list.length) return '';
  const raw = col.subjectChoice !== undefined ? String(row[col.subjectChoice] || '').trim() : '';
  if (!raw) return list[0];
  const asNum = parseInt(raw, 10);
  if (!isNaN(asNum) && String(asNum) === raw) {
    if (asNum < 1 || asNum > list.length) {
      throw new Error('Subject Choice ' + asNum + ' is out of range (1–' + list.length + ').');
    }
    return list[asNum - 1];
  }
  const match = list.filter(function (s) {
    return s.toLowerCase() === raw.toLowerCase();
  })[0];
  if (match) return match;
  throw new Error('Subject Choice "' + raw + '" did not match a numbered option or subject text.');
}

function rowDataMap(row, col) {
  const fullName = col.name !== undefined ? String(row[col.name] || '').trim() : '';
  return {
    first_name: fullName.split(' ')[0] || fullName,
    name: fullName,
    company: col.company !== undefined ? row[col.company] || '' : '',
    role: col.role !== undefined ? row[col.role] || '' : '',
    email: col.email !== undefined ? joinAddresses(parseAddressList(row[col.email])) : ''
  };
}

function mailOptions(row, col, htmlBody) {
  const options = { htmlBody: htmlBody, name: CONFIG.SENDER_NAME };
  const cc = col.cc !== undefined ? joinAddresses(parseAddressList(row[col.cc])) : '';
  const bcc = col.bcc !== undefined ? joinAddresses(parseAddressList(row[col.bcc])) : '';
  if (cc) options.cc = cc;
  if (bcc) options.bcc = bcc;
  return options;
}

function deliverEmail(row, col, built, to) {
  const options = mailOptions(row, col, built.htmlBody);
  const existingThread = col.thread !== undefined ? String(row[col.thread] || '').trim() : '';

  if (built.step > 1 && existingThread) {
    try {
      const thread = GmailApp.getThreadById(existingThread);
      thread.reply(built.plainBody, options);
      log('SEND', 'Replied on thread ' + existingThread);
      return existingThread;
    } catch (e) {
      log('SEND', 'Thread reply failed, sending as a new message: ' + e.message);
    }
  }

  let subject = built.subject;
  if (built.step > 1) {
    const orig = col.originalSubject !== undefined ? String(row[col.originalSubject] || '').trim() : '';
    if (orig) subject = orig.indexOf('Re:') === 0 ? orig : 'Re: ' + orig;
  }

  GmailApp.sendEmail(to, subject, built.plainBody, options);
  log('SEND', 'SENT to ' + to);
  return findSentThreadId(to, subject);
}

function findSentThreadId(to, subject) {
  try {
    Utilities.sleep(1500);
    const firstTo = parseAddressList(to)[0] || '';
    const safeSubject = String(subject || '').replace(/"/g, '');
    const query = 'in:sent to:' + firstTo + ' subject:"' + safeSubject + '"';
    const threads = GmailApp.search(query, 0, 1);
    return threads.length ? threads[0].getId() : '';
  } catch (e) {
    log('SEND', 'Could not capture thread id: ' + e.message);
    return '';
  }
}

function getGmailSignature() {
  try {
    const sendAsList = Gmail.Users.Settings.SendAs.list('me');
    const all = sendAsList.sendAs || [];
    const primary = all.find(function (s) { return s.isPrimary; }) || all[0];
    const sig = primary && primary.signature ? primary.signature : '';
    log('SIGNATURE', 'Fetched signature, length ' + sig.length);
    return sig;
  } catch (e) {
    log('SIGNATURE', 'Could not fetch signature (add "Gmail API" under Services in the editor first): ' + e.message);
    return '';
  }
}

// ============ GROQ ============

function generateWithGroq(context, dataMap, template) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GROQ_API_KEY');
  if (!apiKey) {
    throw new Error('Generate Context is set but no Groq API key is saved. Use the menu: Save Groq API key.');
  }

  const style = template
    ? 'Style reference subjects:\n' + (template.subjects || []).map(function (s, i) {
      return (i + 1) + '. ' + s;
    }).join('\n') + '\n\nStyle reference body HTML:\n' + (template.bodyHtml || '')
    : '(no template selected — write a concise cold email from scratch)';

  const userPrompt =
    'Write one cold job-hunting email.\n' +
    'Lead name: ' + dataMap.name + '\n' +
    'Lead first name: ' + dataMap.first_name + '\n' +
    'Company: ' + dataMap.company + '\n' +
    'Role: ' + dataMap.role + '\n' +
    'Sender context / instructions:\n' + context + '\n\n' +
    style + '\n\n' +
    'Keep {{first_name}}, {{name}}, {{company}}, {{role}} placeholders if they still fit.\n' +
    'Return ONLY JSON with this shape:\n' +
    '{"subjects":["option 1","option 2","option 3"],"body_html":"<p>Hi {{first_name}},</p><p>...</p>"}';

  log('GROQ', 'Calling Groq model ' + CONFIG.GROQ_MODEL);
  const response = UrlFetchApp.fetch(CONFIG.GROQ_API_URL, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + apiKey },
    muteHttpExceptions: true,
    payload: JSON.stringify({
      model: CONFIG.GROQ_MODEL,
      temperature: 0.6,
      response_format: { type: 'json_object' },
      messages: [
        {
          role: 'system',
          content: 'You write short, specific, non-cringe cold emails for job hunting. No subject prefixes like Re: unless asked. No markdown fences.'
        },
        { role: 'user', content: userPrompt }
      ]
    })
  });

  const code = response.getResponseCode();
  const text = response.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error('Groq HTTP ' + code + ': ' + text.slice(0, 400));
  }
  const parsed = JSON.parse(text);
  const content = parsed.choices && parsed.choices[0] && parsed.choices[0].message
    ? parsed.choices[0].message.content
    : '';
  const email = parseGroqContent(content);
  if (!email.subjects.length || !email.bodyHtml) {
    throw new Error('Groq returned an empty subject or body.');
  }
  log('GROQ', 'Generated ' + email.subjects.length + ' subject(s), bodyLen=' + email.bodyHtml.length);
  return email;
}

function parseGroqContent(content) {
  let raw = String(content || '').trim();
  raw = raw.replace(/^```(?:json)?\s*/i, '').replace(/\s*```$/, '');
  const obj = JSON.parse(raw);
  let subjects = obj.subjects || obj.subject_lines || [];
  if (typeof subjects === 'string') subjects = [subjects];
  if (obj.subject && !subjects.length) subjects = [obj.subject];
  let bodyHtml = obj.body_html || obj.bodyHtml || obj.body || '';
  if (bodyHtml && bodyHtml.indexOf('<') === -1) bodyHtml = textToHtml(bodyHtml);
  return { subjects: subjects, bodyHtml: bodyHtml };
}

function textToHtml(text) {
  const escaped = escapeHtml(String(text || '').trim());
  if (!escaped) return '';
  return escaped.split(/\n\n+/).map(function (p) {
    return '<p>' + p.replace(/\n/g, '<br>') + '</p>';
  }).join('');
}

// ============ SCHEDULING HELPERS ============

function nextValidSlot(date) {
  let d = new Date(date);
  let guard = 0;
  while (guard++ < 500) {
    const day = d.getDay();
    if (CONFIG.SEND_DAYS.indexOf(day) === -1) { d = advanceToNextDay(d); continue; }
    const hour = d.getHours() + d.getMinutes() / 60;
    if (hour < CONFIG.START_HOUR) { d.setHours(CONFIG.START_HOUR, 0, 0, 0); continue; }
    if (hour >= CONFIG.END_HOUR) { d = advanceToNextDay(d); continue; }
    return d;
  }
  return d;
}

function advanceToNextDay(d) {
  const nd = new Date(d);
  nd.setDate(nd.getDate() + 1);
  nd.setHours(CONFIG.START_HOUR, 0, 0, 0);
  return nd;
}

function getLatestScheduledTime(data, col) {
  let latest = null;
  for (let r = 1; r < data.length; r++) {
    if (normalizeStatus(data[r][col.status]) !== 'scheduled') continue;
    if (!data[r][col.scheduled]) continue;
    const d = new Date(data[r][col.scheduled]);
    if (!latest || d > latest) latest = d;
  }
  const now = new Date();
  return (!latest || latest < now) ? now : latest;
}

function normalizeStatus(value) {
  return String(value || '').trim().toLowerCase();
}

// ============ SHEET COLUMNS ============

function ensureHeaders(sheet) {
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const have = {};
  header.forEach(function (h) {
    const key = h.toString().trim().toLowerCase();
    if (key) have[key] = true;
  });
  const toAdd = REQUIRED_HEADERS.filter(function (name) { return !have[name.toLowerCase()]; });
  if (!toAdd.length) return;
  const start = header.filter(function (h) { return String(h).trim(); }).length + 1;
  sheet.getRange(1, start, 1, toAdd.length).setValues([toAdd]);
  log('SHEET', 'Appended columns: ' + toAdd.join(', '));
}

function mapColumns(headerRow) {
  const map = {};
  headerRow.forEach(function (h, idx) {
    const key = h.toString().trim().toLowerCase();
    if (key === 'email' || key === 'to') map.email = idx;
    else if (key === 'cc') map.cc = idx;
    else if (key === 'bcc') map.bcc = idx;
    else if (key === 'name') map.name = idx;
    else if (key === 'company') map.company = idx;
    else if (key === 'role') map.role = idx;
    else if (key === 'template' || key === 'template id') map.template = idx;
    else if (key === 'subject choice' || key === 'subject' || key === 'subject index') map.subjectChoice = idx;
    else if (key === 'generate context' || key === 'context' || key === 'prompt') map.context = idx;
    else if (key === 'status') map.status = idx;
    else if (key === 'sequence step' || key === 'step') map.step = idx;
    else if (key === 'scheduled time') map.scheduled = idx;
    else if (key === 'sent time') map.sent = idx;
    else if (key === 'next follow-up' || key === 'next follow up' || key === 'next followup') map.nextFollowup = idx;
    else if (key === 'thread id' || key === 'thread') map.thread = idx;
    else if (key === 'original subject') map.originalSubject = idx;
    else if (key === 'reply' || key === 'replied') map.reply = idx;
    else if (key === 'error') map.error = idx;
  });
  return map;
}

function parseAddressList(value) {
  if (value === null || value === undefined) return [];
  return String(value)
    .split(/[;,]/)
    .map(function (s) { return s.trim(); })
    .filter(function (s) { return s && s.indexOf('@') !== -1; });
}

function joinAddresses(list) {
  return (list || []).join(', ');
}

// ============ TEMPLATE (DOC) PARSING ============

function getFlatLines(body) {
  const lines = [];
  const n = body.getNumChildren();
  for (let i = 0; i < n; i++) {
    const el = body.getChild(i);
    const type = el.getType();
    if (type !== DocumentApp.ElementType.PARAGRAPH && type !== DocumentApp.ElementType.LIST_ITEM) continue;
    const raw = type === DocumentApp.ElementType.PARAGRAPH ? el.asParagraph().getText() : el.asListItem().getText();
    raw.split(/[\n\v\r]+/).forEach(function (part) {
      lines.push({ text: part, paraIndex: i });
    });
  }
  return lines;
}

function parseTemplateMarker(text) {
  const trimmed = String(text || '').trim();
  const match = trimmed.match(/^(?:=+\s*)?template\s*:\s*(.+?)\s*(?:=+)?$/i);
  if (!match) return null;
  return match[1].replace(/=+$/, '').trim().toLowerCase();
}

function parseAllTemplates(doc) {
  log('PARSE', 'Parsing doc: ' + doc.getName());
  const body = doc.getBody();
  const lines = getFlatLines(body);
  const starts = [];
  lines.forEach(function (line, idx) {
    const name = parseTemplateMarker(line.text);
    if (name) starts.push({ index: idx, name: name, paraIndex: line.paraIndex });
  });

  const templates = {};
  if (!starts.length) {
    templates[CONFIG.DEFAULT_TEMPLATE_ID] = parseTemplateSection(body, lines, 0, lines.length);
    return templates;
  }

  for (let t = 0; t < starts.length; t++) {
    const end = t + 1 < starts.length ? starts[t + 1].index : lines.length;
    templates[starts[t].name] = parseTemplateSection(body, lines, starts[t].index + 1, end);
  }
  return templates;
}

function parseTemplateSection(body, lines, start, end) {
  const subjects = [];
  let bodyRemainder = '';
  let bodyStartPara = null;
  const sectionEndPara = end < lines.length ? lines[end].paraIndex : body.getNumChildren();

  for (let i = start; i < end; i++) {
    const raw = lines[i].text.trim();
    const lower = raw.toLowerCase();
    if (lower.indexOf('subject:') === 0) {
      const after = raw.substring(lower.indexOf('subject:') + 'subject:'.length).trim();
      if (after) {
        subjects.push(after);
      } else {
        for (let j = i + 1; j < end; j++) {
          const t = lines[j].text.trim();
          const tl = t.toLowerCase();
          if (!t) continue;
          if (tl.indexOf('subject:') === 0 || tl.indexOf('body:') === 0) break;
          if (parseTemplateMarker(t)) break;
          subjects.push(t);
          break;
        }
      }
    } else if (lower.indexOf('body:') === 0 && bodyStartPara === null) {
      bodyRemainder = raw.substring(lower.indexOf('body:') + 'body:'.length).trim();
      bodyStartPara = lines[i].paraIndex + 1;
    }
  }

  if (bodyStartPara === null) {
    bodyStartPara = start < lines.length ? lines[start].paraIndex : 0;
  }

  let bodyHtml = docToBodyHtml(body, bodyStartPara, sectionEndPara);
  if (bodyRemainder) bodyHtml = '<p>' + escapeHtml(bodyRemainder) + '</p>' + bodyHtml;
  log('PARSE', 'Template section subjects=' + subjects.length + ' bodyLen=' + bodyHtml.length);
  return { subjects: subjects, bodyHtml: bodyHtml };
}

function docToBodyHtml(body, startIndex, endIndex) {
  const n = endIndex === undefined ? body.getNumChildren() : endIndex;
  let html = '';
  let inList = false;
  for (let i = startIndex; i < n; i++) {
    const el = body.getChild(i);
    const type = el.getType();
    if (type === DocumentApp.ElementType.LIST_ITEM) {
      if (!inList) { html += '<ul>'; inList = true; }
      html += '<li>' + inlineHtml(el.asListItem().editAsText()) + '</li>';
    } else {
      if (inList) { html += '</ul>'; inList = false; }
      if (type === DocumentApp.ElementType.PARAGRAPH) {
        const t = el.asParagraph().getText().trim();
        if (parseTemplateMarker(t)) continue;
        html += t === '' ? '<br>' : '<p>' + inlineHtml(el.asParagraph().editAsText()) + '</p>';
      }
    }
  }
  if (inList) html += '</ul>';
  return html;
}

function inlineHtml(text) {
  const str = text.getText();
  if (!str) return '';
  const indices = text.getTextAttributeIndices();
  let html = '';
  for (let i = 0; i < indices.length; i++) {
    const start = indices[i];
    const end = (i + 1 < indices.length) ? indices[i + 1] - 1 : str.length - 1;
    if (end < start) continue;
    let chunk = escapeHtml(str.substring(start, end + 1));
    const link = text.getLinkUrl(start);
    if (link) chunk = '<a href="' + link + '">' + chunk + '</a>';
    if (text.isBold(start)) chunk = '<b>' + chunk + '</b>';
    if (text.isItalic(start)) chunk = '<i>' + chunk + '</i>';
    html += chunk;
  }
  return html;
}

function escapeHtml(s) {
  return String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function htmlToPlain(htmlBody) {
  return String(htmlBody || '')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&nbsp;/g, ' ')
    .trim();
}

function personalize(str, data) {
  if (!str) return '';
  return String(str).replace(/\{\{\s*([\w]+)\s*\}\}/g, function (match, key) {
    key = key.toLowerCase();
    return data.hasOwnProperty(key) ? data[key] : match;
  });
}
