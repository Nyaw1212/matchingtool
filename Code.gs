/** Bound to the reconciliation workbook. Enable the advanced Google Sheets service as "Sheets". */
const AUDIT = 'July Home Match Audit 2026';
const LEDGER = 'Home Depot Application Log';
const LOG_HEADERS = ['Audit key', 'Posted at', 'Transaction ID', 'Order number',
  'Signed amount', 'Customer', 'Job type', 'Job start', 'Job end',
  'JULY row at posting', 'Previous total', 'Updated total'];

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Home Depot')
    .addItem('Apply approved candidates', 'applySelectedHomeRows').addToUi();
}

function applySelectedHomeRows() {
  const ui = SpreadsheetApp.getUi();
  try {
    const book = SpreadsheetApp.getActive();
    const audit = book.getActiveSheet();
    if (audit.getName() !== AUDIT) {
      throw new Error('Open the July Home Match Audit 2026 tab before applying candidates.');
    }
    const rows = approvedCandidateRows_(audit);
    const preview = planHomePosting_(book, rows);
    const summary = preview.items.map(item =>
      '#' + item.row + ' ' + item.name + '  ' + item.amount.toFixed(2) +
      (item.recovery ? ' (verify prior posting)' : '')).join('\n');
    const skipped = preview.skipped.join('\n');
    if (!preview.items.length) {
      ui.alert('Nothing to apply', skipped || 'Choose Candidate 1, Candidate 2, or Candidate 3 in column A.', ui.ButtonSet.OK);
      return;
    }
    if (ui.alert('Confirm Home Depot posting',
      summary + (skipped ? '\n\nSkipped\n' + skipped : ''), ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    let message;
    try {
      const live = planHomePosting_(book, approvedCandidateRows_(audit));
      if (JSON.stringify(live) !== JSON.stringify(preview)) {
        throw new Error('The audit or job totals changed during review. Run the preview again.');
      }
      message = commitHomePosting_(book, live);
    } finally {
      lock.releaseLock();
    }
    ui.alert(message);
  } catch (error) {
    ui.alert(error.message);
  }
}

function planHomePosting_(book, rows) {
  const audit = book.getSheetByName(AUDIT);
  const home = book.getSheetByName('Home');
  const july = book.getSheetByName('JULY');
  const log = book.getSheetByName(LEDGER);
  const data = audit.getDataRange().getValues();
  const col = name => {
    const i = data[0].indexOf(name);
    if (i < 0) throw new Error('Missing audit column ' + name);
    return i;
  };
  const sources = home.getRange(8, 1, home.getLastRow() - 7, 17).getDisplayValues();
  const jobs = july.getRange(1, 1, july.getLastRow(), 8).getValues();
  const formulas = july.getRange(1, 8, jobs.length, 1).getFormulas();
  const history = log ? log.getDataRange().getValues() : [LOG_HEADERS];
  if (JSON.stringify(history[0]) !== JSON.stringify(LOG_HEADERS)) {
    throw new Error('The application log headers changed.');
  }
  const records = history.slice(1);
  const known = new Map(records.map(record => [record[0], record]));
  if (known.size !== records.length) throw new Error('The log contains repeated audit keys.');
  const items = [];
  const skipped = [];
  for (const row of rows) {
    try {
      const cell = name => data[row - 1][col(name)];
      const choice = cell('Approval');
      const override = String(cell('Customer override')).trim();
      const rank = /^Candidate ([123])$/.exec(choice)?.[1];
      if (choice === 'APPLIED' || (!rank && !(choice === 'PENDING' && override))) {
        throw new Error('Choose a candidate or enter a customer override on a PENDING row.');
      }
      const key = cell('Audit key');
      if (!key || data.slice(1).filter(r => r[col('Audit key')] === key).length !== 1) {
        throw new Error('Audit key is missing or repeated.');
      }
      const sourceKey = normalizeSourceKey_(key);
      const matching = sources.filter(s => normalizeSourceKey_(
        [s[3],s[4],s[5],s[2],s[0] || s[1],s[12],s[16],parseMoney_(s[11])]
      ) === sourceKey);
      if (matching.length !== 1) throw new Error('Source transaction could not be uniquely reverified.');
      const s = matching[0];
      const [date, amount] = [s[0] || s[1], parseMoney_(s[11])];
      if (date < '2026-07-10' || date > '2026-07-31' || date !== cell('Date used') ||
          !Number.isFinite(amount) || amount !== cell('Signed amount') ||
          s[4] !== String(cell('Transaction ID')) || s[12] !== String(cell('Order number'))) {
        throw new Error('Source date, identifiers or amount changed.');
      }
      if (!s[4] || !s[12] || sources.some(other => other !== s && (other[4] === s[4] || other[12] === s[12]))) {
        throw new Error('Missing or shared supplier identifier; reconcile this receipt manually.');
      }
      const name = override || cell('Candidate ' + rank);
      const jobType = override ? '' : cell('Job type ' + rank);
      const dates = override ? '' : cell('Job dates ' + rank);
      const eligible = jobs.map((job, i) => ({job, i})).filter(({job}) =>
        name && normalize_(job[0]) === normalize_(name) && /^(install|service)$/i.test(job[1]) &&
        (!jobType || job[1] === jobType) &&
        (!dates || [jobDate_(book, job[3]), jobDate_(book, job[4])].filter(Boolean).join(' – ') === dates));
      if (eligible.length !== 1) throw new Error('Destination is missing or ambiguous in JULY.');
      const {job, i} = eligible[0];
      const start = jobDate_(book, job[3]);
      const end = jobDate_(book, job[4]);
      const sameJob = rec => rec[5] === job[0] && rec[6] === job[1] &&
        rec[7] === start && rec[8] === end;
      const jobRecords = records.filter(sameJob);
      const current = job[7] === '' ? 0 : parseMoney_(job[7]);
      const logged = jobRecords.reduce((total, rec) => total + cents_(rec[4]), 0);
      if (!Number.isFinite(current) || cents_(current) !== logged || (formulas[i][0] && formulas[i][0] !== postingFormula_(jobRecords.map(rec => cents_(rec[4]))))) {
        throw new Error('Existing cost cannot be reconciled to the application log.');
      }
      const previous = known.get(key);
      if (previous && (!sameJob(previous) || previous[4] !== amount)) {
        throw new Error('Previously logged posting differs from this selection.');
      }
      items.push({row, key, choice, name: job[0], type: job[1], start, end, dest: i + 1,
        amount, current, terms: jobRecords.map(rec => cents_(rec[4])), transaction: s[4], order: s[12], recovery: Boolean(previous)});
    } catch (error) {
      skipped.push('#' + row + ' ' + error.message);
    }
  }
  return {items, skipped};
}

function commitHomePosting_(book, plan) {
  const audit = book.getSheetByName(AUDIT);
  const july = book.getSheetByName('JULY');
  let log = book.getSheetByName(LEDGER);
  const totals = new Map();
  const entries = [];
  const terms = new Map();
  const now = new Date().toISOString();
  for (const item of plan.items) {
    if (item.recovery) continue;
    const before = totals.get(item.dest) ?? cents_(item.current);
    const after = before + cents_(item.amount);
    totals.set(item.dest, after);
    terms.set(item.dest, [...(terms.get(item.dest) ?? item.terms), cents_(item.amount)]);
    entries.push([item.key, now, item.transaction, item.order, item.amount, item.name, item.type,
      item.start, item.end, item.dest, before / 100, after / 100]);
  }
  if (entries.length) {
    if (!log) {
      log = book.insertSheet(LEDGER);
      log.getRange(1, 1, 1, LOG_HEADERS.length).setValues([LOG_HEADERS]);
      log.setFrozenRows(1);
      SpreadsheetApp.flush();
    }
    // A single Sheets API batch commits both the job totals and their ledger entries.
    const requests = [...terms].map(([row, amounts]) =>
      writeCells_(july.getSheetId(), row - 1, 7, [[{formulaValue: postingFormula_(amounts)}]]));
    requests.push({appendCells: {sheetId: log.getSheetId(), rows: entries.map(cellRow_),
      fields: 'userEnteredValue'}});
    Sheets.Spreadsheets.batchUpdate({requests}, book.getId());
  }
  // Read back through the same API that performed the write to avoid cached Apps Script reads.
  verifyHomePosting_(book.getId(), plan.items, totals, terms);
  const verified = plan.items;
  const headers = audit.getRange(1, 1, 1, audit.getLastColumn()).getValues()[0];
  const pos = name => headers.indexOf(name);
  const updates = verified.items.flatMap(item => [
    writeCells_(audit.getSheetId(), item.row - 1, pos('Selected customer'), [[item.name]]),
    writeCells_(audit.getSheetId(), item.row - 1, pos('Selected JULY row'), [[item.dest]]),
    writeCells_(audit.getSheetId(), item.row - 1, pos('Approval'), [['APPLIED']])
  ]);
  Sheets.Spreadsheets.batchUpdate({requests: updates}, book.getId());
  const statuses = Sheets.Spreadsheets.Values.batchGet(book.getId(), {
    ranges: verified.items.map(item => "'" + AUDIT + "'!A" + item.row)
  }).valueRanges;
  if (statuses.some(range => range.values?.[0]?.[0] !== 'APPLIED')) {
    throw new Error('Some APPLIED statuses could not be verified.');
  }
  const purchases = entries.reduce((total, entry) => total + Math.max(0, cents_(entry[4])), 0);
  const refunds = entries.reduce((total, entry) => total + Math.min(0, cents_(entry[4])), 0);
  const net = purchases + refunds;
  return 'Newly posted: ' + entries.length + '. Net: $' + (net / 100).toFixed(2) +
    '. Purchases: ' + (purchases / 100).toFixed(2) + '. Refunds: ' + (refunds / 100).toFixed(2) + '. Jobs updated: ' + totals.size + '. Prior postings verified: ' + (verified.items.length - entries.length) +
    '. Skipped: ' + plan.skipped.length + '.';
}

function writeCells_(sheetId, rowIndex, columnIndex, values) {
  return {updateCells: {start: {sheetId, rowIndex, columnIndex}, rows: values.map(cellRow_),
    fields: 'userEnteredValue'}};
}
function cellRow_(values) {
  return {values: values.map(value => ({userEnteredValue: typeof value === 'number'
    ? {numberValue: value} : typeof value === 'object' ? value : {stringValue: String(value)}}))};
}
function jobDate_(book, value) {
  return value instanceof Date ? Utilities.formatDate(value, book.getSpreadsheetTimeZone(), 'M/d/yyyy')
    : String(value || '');
}
function normalize_(value) {
  return String(value || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
}
function parseMoney_(value) {
  const cleaned = String(value ?? '').trim().replace(/[$,]/g, '').replace(/^\((.*)\)$/, '-$1');
  return /^-?\d+(\.\d{1,2})?$/.test(cleaned) ? Number(cleaned) : NaN;
}
function cents_(amount) {
  return Math.round(amount * 100);
}



function postingFormula_(amounts) {
  if (!amounts.length) return '';
  return '=' + amounts.map((amount, i) => {
    const sign = amount < 0 ? '-' : i ? '+' : '';
    return sign + (Math.abs(amount) / 100).toFixed(2);
  }).join('');
}

function normalizeSourceKey_(key) {
  const fields = typeof key === 'string' ? JSON.parse(key) : [...key];
  // The audit export uses null for a missing invoice. getDisplayValues uses ''.
  if (fields[6] === null) fields[6] = '';
  return JSON.stringify(fields);
}

function approvedCandidateRows_(audit) {
  const count = audit.getLastRow() - 1;
  if (count < 1) return [];
  return audit.getRange(2, 1, count, 1).getValues().flatMap(([choice], i) =>
    /^Candidate [123]$/.test(choice) ? [i + 2] : []);
}

function verifyHomePosting_(bookId, items, totals, terms) {
  const destinations = [...new Map(items.map(item => [item.dest, item])).values()];
  const ranges = ["'" + LEDGER + "'!A:L", ...destinations.map(item => 'JULY!H' + item.dest)];
  const live = Sheets.Spreadsheets.Values.batchGet(bookId, {
    ranges, valueRenderOption: 'UNFORMATTED_VALUE'
  }).valueRanges;
  const log = (live[0].values || []).slice(1);
  const receiptsVerified = items.every(item => {
    const matches = log.filter(rec => rec[0] === item.key);
    return matches.length === 1 && matches[0].slice(4, 9).every((value, i) =>
      value === [item.amount, item.name, item.type, item.start, item.end][i]);
  });
  const costsVerified = destinations.every((item, i) => {
    const amount = live[i + 1].values?.[0]?.[0];
    return typeof amount === 'number' && cents_(amount) ===
      (totals.get(item.dest) ?? cents_(item.current));
  });
  const formulas = terms.size ? Sheets.Spreadsheets.Values.batchGet(bookId, {
    ranges: [...terms.keys()].map(row => 'JULY!H' + row), valueRenderOption: 'FORMULA'
  }).valueRanges : [];
  const formulasVerified = [...terms.values()].every((amounts, i) =>
    formulas[i]?.values?.[0]?.[0] === postingFormula_(amounts));
  if (!receiptsVerified || !costsVerified || !formulasVerified) {
    throw new Error('Posting verification did not pass. Inspect the application log before retrying.');
  }
}
