/**
 * Google Sheet receiver for OT + GK register extraction.
 * GK schema v15: Drive-backed persistent queue + automatic background processing.
 *
 * STATUS BEHAVIOR:
 * - ONE 6-row status block per PDF/run.
 * - During that PDF, Page / Records / Message update in the SAME top block.
 * - When a NEW PDF starts, the entire previous block is shifted DOWN.
 * - Previous PDF status blocks are NEVER cleared or overwritten.
 * - Newest PDF is always at the top.
 * - Supports quota/status-only updates from otg35.py.
 * - Uses a hidden page-ID index to block duplicate page uploads.
 * - Live Sheet keeps A:J unchanged, K:L blank, and writes File Name in M for new rows.
 * - ALSO supports the GK register workflow when POST body contains dataset='gk'.
 * - GK data goes to separate tabs so the existing OT workflow is not disturbed.
 */

const DATA_TAB = 'Live Sheet';
const STATUS_TAB = 'Status of pdf extration';
const SOURCE_INDEX_TAB = '_OT Source Index';

const STATUS_BLOCK_ROWS = 6;
const STATUS_SEPARATOR_ROWS = 1;

// Permanent date-range controls directly above Column A.
// A1 = From Date label, B1 = From Date picker
// A2 = To Date label,   B2 = To Date picker
// OT headers are on row 3; data begins on row 4.
const FROM_DATE_LABEL_CELL = 'A1';
const FROM_DATE_CELL = 'B1';
const TO_DATE_LABEL_CELL = 'A2';
const TO_DATE_CELL = 'B2';

const DATA_HEADER_ROW = 3;
const DATA_START_ROW = 4;

// Keep two blank spacer columns after the 10 OT columns.
// K and L stay blank; M contains the source PDF file name for NEW rows.
const FILE_NAME_COL = 13; // M
const FILE_NAME_HEADER = 'File Name';


// -----------------------------------------------------------------------------
// GK register dataset (separate tabs; existing OT tabs remain unchanged)
// -----------------------------------------------------------------------------
const GK_DATA_TAB = 'GK Live Sheet';
const GK_STATUS_TAB = 'Status of GK extraction';
const GK_DATA_HEADER_ROW = 3;
const GK_DATA_START_ROW = 4;
const GK_FILE_NAME_COL = 22; // V; T:U intentionally blank.
const GK_HEADERS = [
  'Serial Number',
  'Date',
  'Gamma Knife Number',
  'UHID Number',
  'Patient Name',
  'Mobile Number(s)',
  'Age',
  'Sex',
  'Diagnosis',
  'Mask',
  'Dose',
  'Percentage',
  'Volume (cc)',
  'GURA',
  'Consultant',
  'Physician',
  'Residence',
  'RTT',
  'Technique'
];

const GK_FIELD_KEYS = [
  'sno',
  'date',
  'gk_no',
  'uhid_no',
  'patient_name',
  'mobile_no',
  'age',
  'sex',
  'diagnosis',
  'mask',
  'dose',
  'perc',
  'volume',
  'gura',
  'cons',
  'phy',
  'ro',
  'rtt',
  'technique'
];


// -----------------------------------------------------------------------------
// Persistent Google Drive + background queue
// -----------------------------------------------------------------------------
const GK_QUEUE_TAB = '_GK Queue';
const GK_QUEUE_HEADERS = [
  'Queue ID',
  'Drive File ID',
  'File Name',
  'MIME Type',
  'Drive URL',
  'Content SHA256',
  'Status',
  'Attempts',
  'Added At',
  'Last Update',
  'Next Attempt',
  'Last Error',
  'Records'
];
const GK_QUEUE_START_ROW = 2;
const GK_WORKER_HANDLER = 'processGKQueueTrigger';
const GK_KICK_HANDLER = 'processGKQueueKick';
const GK_WORKER_EVERY_MINUTES = 5;
const GK_WORKER_MAX_SECONDS = 240;
const GK_QUOTA_RETRY_MINUTES = 15;
const GK_RETRY_MINUTES = 5;
const GK_PROCESSING_STALE_MINUTES = 15;



/**
 * One-time setup.
 */
function setupReceiver() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss) {
    throw new Error('Open this Apps Script from the target Google Sheet.');
  }

  const props = PropertiesService.getScriptProperties();

  let key = props.getProperty('RECEIVER_KEY');

  if (!key) {
    key =
      Utilities.getUuid().replace(/-/g, '') +
      Utilities.getUuid().replace(/-/g, '');
  }

  props.setProperty('RECEIVER_KEY', key);
  props.setProperty('SPREADSHEET_ID', ss.getId());

  const data = ensureSheet_(ss, DATA_TAB);
  ensureStatusSheet_(ss);
  ensureSourceIndex_(ss);
  migrateLiveSheetLayout_(data);
  ensureDateControls_(data);
  ensureDataHeaders_(data, []);
  ensureFileNameHeader_(data);

  // New GK tabs. This does not modify or clear existing OT data.
  gkEnsureLayout_(ss);
  gkEnsureQueue_(ss);

  console.log('RECEIVER KEY (keep private): ' + key);
  console.log('Spreadsheet ID saved: ' + ss.getId());

  SpreadsheetApp.getUi().alert(
    'Receiver initialized.\n\n' +
    'Open Executions / execution log and copy the RECEIVER KEY.\n' +
    'Keep that key private.'
  );
}


/**
 * Receiver.
 *
 * A script lock prevents two simultaneous POST requests from changing the
 * same top status block at the same time.
 */
function doPost(e) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000);

    if (!e || !e.postData || !e.postData.contents) {
      return json_({ok: false, error: 'Empty request'});
    }

    const body = JSON.parse(e.postData.contents);

    const props = PropertiesService.getScriptProperties();
    const expectedKey = props.getProperty('RECEIVER_KEY');

    if (!expectedKey || String(body.key || '') !== expectedKey) {
      return json_({ok: false, error: 'Unauthorized'});
    }

    const spreadsheetId = props.getProperty('SPREADSHEET_ID');

    if (!spreadsheetId) {
      return json_({
        ok: false,
        error: 'Receiver not initialized'
      });
    }

    const ss = SpreadsheetApp.openById(spreadsheetId);

    // GK register requests are routed to separate tabs. Existing OT behavior
    // below is deliberately left unchanged.
    if (String(body.dataset || '').toLowerCase() === 'gk') {
      return handleGK_(ss, body, props);
    }

    const data = ensureSheet_(ss, DATA_TAB);
    const status = ensureStatusSheet_(ss);
    const sourceIndex = ensureSourceIndex_(ss);
    migrateLiveSheetLayout_(data);
    ensureDateControls_(data);
    ensureFileNameHeader_(data);

    const action = String(body.action || '');

    // ---------------------------------------------------------------
    // NEW PDF / NEW RUN
    // ---------------------------------------------------------------
    if (action === 'reset') {
      // Preserve existing Live Sheet rows when the Python sender asks for it.
      // Otherwise keep the existing reset behavior.
      if (body.keep_existing === true) {
        ensureDataHeaders_(data, body.headers || []);
      } else {
        resetData_(data, body.headers || []);
      }

      const inputName = String(body.input_name || '');
      const runId = String(body.run_id || '');
      const currentRunId = props.getProperty('CURRENT_RUN_ID') || '';
      const isResume =
        runId !== '' &&
        currentRunId === runId &&
        hasTopStatusBlock_(status);

      props.setProperty('CURRENT_INPUT_NAME', inputName);
      props.setProperty('CURRENT_RUN_ID', runId);

      // A restart of the SAME PDF/run reuses the current top block instead of
      // creating a duplicate status-history block. A genuinely new PDF still
      // archives the previous 6-row block exactly as before.
      if (!isResume) {
        startNewStatusBlock_(status);
      }

      const completedPages = Number(body.page_number || 0);
      const existingRecords = Number(body.total_records || 0);

      writeStatusBlock_(
        status,
        inputName,
        isResume ? 'Resumed' : 'Started',
        completedPages + '/' + Number(body.total_pages || 0),
        existingRecords,
        isResume ? 'Resumed without duplicating completed pages' : 'Started'
      );

      applyDateRangeFilter_(data);
      SpreadsheetApp.flush();

      return json_({
        ok: true,
        action: 'reset'
      });
    }

    // ---------------------------------------------------------------
    // PAGE COMPLETED
    //
    // This DOES NOT create a new status-history row/block.
    // It only updates the current PDF's top 6-row block.
    // ---------------------------------------------------------------
    if (action === 'append') {
      const rows = Array.isArray(body.rows) ? body.rows : [];
      const pageId = String(body.page_id || '').trim();
      const pageNumber = Number(body.page_number || 0);
      const totalPages = Number(body.total_pages || 0);
      const totalRecords = Number(body.total_records || 0);

      // If the exact scanned page was already accepted, never append it again.
      // This also protects against a POST that succeeded but whose HTTP response
      // was lost, causing Python to retry the same request.
      if (pageId && sourcePageExists_(sourceIndex, pageId)) {
        writeStatusBlock_(
          status,
          props.getProperty('CURRENT_INPUT_NAME') || '',
          'Processing',
          pageNumber + '/' + totalPages,
          totalRecords,
          'Page ' + pageNumber + ' already uploaded; duplicate skipped'
        );

        applyDateRangeFilter_(data);
        SpreadsheetApp.flush();

        return json_({
          ok: true,
          action: 'append',
          appended: 0,
          duplicate: true
        });
      }

      let normalized = [];

      if (rows.length) {
        normalized = rows.map(function(row) {
          const r = Array.isArray(row)
            ? row.slice(0, 10)
            : [];

          while (r.length < 10) {
            r.push('');
          }

          return r.map(function(v) {
            return String(v == null ? '' : v);
          });
        });

        const appendStartRow = data.getLastRow() + 1;

        // A:J = OT data. Existing rows remain untouched.
        data.getRange(
          appendStartRow,
          1,
          normalized.length,
          10
        ).setValues(normalized);

        // K:L intentionally remain blank.
        // M = source PDF file name for every newly appended record.
        const sourceFileName =
          String(body.input_name || '') ||
          props.getProperty('CURRENT_INPUT_NAME') || '';

        const fileNameValues = normalized.map(function() {
          return [sourceFileName];
        });

        data.getRange(
          appendStartRow,
          FILE_NAME_COL,
          normalized.length,
          1
        ).setValues(fileNameValues);
      }

      // Mark even a zero-row page as processed so a restart does not repeatedly
      // reprocess/upload an empty or non-table page.
      if (pageId) {
        markSourcePage_(
          sourceIndex,
          pageId,
          props.getProperty('CURRENT_INPUT_NAME') || '',
          pageNumber,
          normalized.length
        );
      }

      writeStatusBlock_(
        status,
        props.getProperty('CURRENT_INPUT_NAME') || '',
        'Processing',
        pageNumber + '/' + totalPages,
        totalRecords,
        'Page ' + pageNumber + ' uploaded'
      );

      // Keep the user's selected date range active as new PDF rows arrive.
      applyDateRangeFilter_(data);

      SpreadsheetApp.flush();

      return json_({
        ok: true,
        action: 'append',
        appended: normalized.length,
        duplicate: false
      });
    }

    // ---------------------------------------------------------------
    // STATUS-ONLY UPDATE
    //
    // Used by otg35.py while Gemini is rate/quota limited. No Live Sheet
    // rows are changed and no status-history block is created.
    // ---------------------------------------------------------------
    if (action === 'status') {
      const inputName =
        String(body.input_name || '') ||
        props.getProperty('CURRENT_INPUT_NAME') || '';

      writeStatusBlock_(
        status,
        inputName,
        String(body.state || 'Processing'),
        Number(body.page_number || 0) +
          '/' +
          Number(body.total_pages || 0),
        Number(body.total_records || 0),
        String(body.message || '')
      );

      SpreadsheetApp.flush();

      return json_({
        ok: true,
        action: 'status'
      });
    }

    // ---------------------------------------------------------------
    // PDF FINISHED
    //
    // Final values stay permanently in this PDF's top block.
    // On the NEXT reset, this complete block is shifted down untouched.
    // ---------------------------------------------------------------
    if (action === 'finish') {
      writeStatusBlock_(
        status,
        props.getProperty('CURRENT_INPUT_NAME') || '',
        'Done',
        Number(body.total_pages || 0) +
          '/' +
          Number(body.total_pages || 0),
        Number(body.total_records || 0),
        'Extraction complete'
      );

      SpreadsheetApp.flush();

      return json_({
        ok: true,
        action: 'finish'
      });
    }

    return json_({
      ok: false,
      error: 'Unknown action'
    });

  } catch (err) {
    return json_({
      ok: false,
      error: String(
        err && err.message
          ? err.message
          : err
      )
    });

  } finally {
    try {
      lock.releaseLock();
    } catch (_err) {
      // No action needed.
    }
  }
}


/**
 * Hidden technical sheet used only to remember which physical scanned pages
 * have already been appended to Live Sheet. Users do not need to touch it.
 */
function ensureSourceIndex_(ss) {
  let sheet = ss.getSheetByName(SOURCE_INDEX_TAB);

  if (!sheet) {
    sheet = ss.insertSheet(SOURCE_INDEX_TAB);
  }

  const headers = [
    'Page ID',
    'File Name',
    'PDF Page',
    'Records',
    'Uploaded At'
  ];

  const current = sheet
    .getRange(1, 1, 1, headers.length)
    .getDisplayValues()[0];

  const hasHeader = current.some(function(v) {
    return String(v || '').trim() !== '';
  });

  if (!hasHeader) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }

  try {
    sheet.hideSheet();
  } catch (_err) {
    // Harmless if Sheets temporarily refuses a hide operation.
  }

  return sheet;
}


function sourcePageExists_(sheet, pageId) {
  if (!pageId || sheet.getLastRow() < 2) {
    return false;
  }

  const finder = sheet
    .getRange(2, 1, sheet.getLastRow() - 1, 1)
    .createTextFinder(pageId)
    .matchEntireCell(true);

  return finder.findNext() !== null;
}


function markSourcePage_(sheet, pageId, fileName, pageNumber, records) {
  if (!pageId || sourcePageExists_(sheet, pageId)) {
    return;
  }

  sheet.appendRow([
    pageId,
    fileName,
    Number(pageNumber || 0),
    Number(records || 0),
    new Date()
  ]);

  try {
    sheet.hideSheet();
  } catch (_err) {
    // No action needed.
  }
}


/**
 * Push the previous PDF's COMPLETE status block down before starting
 * a new PDF.
 *
 * Layout after PDF #3 starts:
 *
 * rows 1..6   -> PDF #3 current status
 * row 7       -> blank separator
 * rows 8..13  -> PDF #2 final status
 * row 14      -> blank separator
 * rows 15..20 -> PDF #1 final status
 *
 * Old blocks are not cleared, rewritten, or deleted.
 */
function startNewStatusBlock_(sheet) {
  if (!hasTopStatusBlock_(sheet)) {
    return;
  }

  sheet.insertRowsBefore(
    1,
    STATUS_BLOCK_ROWS + STATUS_SEPARATOR_ROWS
  );
}


/**
 * Detect whether rows 1..6 contain an existing status block.
 * Formatting alone does not count.
 */
function hasTopStatusBlock_(sheet) {
  const values = sheet
    .getRange(1, 1, STATUS_BLOCK_ROWS, 2)
    .getDisplayValues();

  return values.some(function(row) {
    return row.some(function(value) {
      return String(value || '').trim() !== '';
    });
  });
}


/**
 * Update ONLY the active PDF block at rows 1..6.
 *
 * This is intentionally overwritten while the SAME PDF is running.
 * It never touches any archived PDF block below row 7.
 */
function writeStatusBlock_(
  sheet,
  fileName,
  state,
  page,
  records,
  message
) {
  const values = [
    ['File Name', fileName],
    ['Status', state],
    ['Page', page],
    ['Records uploaded', String(records)],
    ['Last update', new Date()],
    ['Message', message]
  ];

  sheet
    .getRange(1, 1, STATUS_BLOCK_ROWS, 2)
    .setValues(values);

  // Small formatting only; it does not clear any old data.
  sheet
    .getRange(1, 1, STATUS_BLOCK_ROWS, 1)
    .setFontWeight('bold');

  sheet
    .getRange(5, 2)
    .setNumberFormat('M/d/yyyy h:mm:ss AM/PM');

  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 320);
}


/**
 * Clear/rebuild Live Sheet when NOT keeping existing data.
 */
function resetData_(sheet, headers) {
  // Clear ONLY the OT table area. Keep A1:B2 date controls intact.
  const maxRows = Math.max(sheet.getMaxRows(), DATA_HEADER_ROW);

  sheet.getRange(
    DATA_HEADER_ROW,
    1,
    maxRows - DATA_HEADER_ROW + 1,
    10
  ).clearContent();

  // Show all data rows before writing a new dataset.
  if (sheet.getMaxRows() >= DATA_START_ROW) {
    sheet.showRows(
      DATA_START_ROW,
      sheet.getMaxRows() - DATA_START_ROW + 1
    );
  }

  const h = normalizeHeaders_(headers);

  sheet
    .getRange(DATA_HEADER_ROW, 1, 1, 10)
    .setValues([h]);

  // If a full reset is explicitly requested, clear old file-name values too.
  if (maxRows >= DATA_START_ROW) {
    sheet.getRange(
      DATA_START_ROW,
      FILE_NAME_COL,
      maxRows - DATA_START_ROW + 1,
      1
    ).clearContent();
  }

  ensureFileNameHeader_(sheet);
  sheet.setFrozenRows(DATA_HEADER_ROW);
}


/**
 * When keep_existing=true, do not clear Live Sheet.
 * Only add the header if the sheet is empty.
 */
function ensureDataHeaders_(sheet, headers) {
  const current = sheet
    .getRange(DATA_HEADER_ROW, 1, 1, 10)
    .getDisplayValues()[0];

  const hasHeader = current.some(function(value) {
    return String(value || '').trim() !== '';
  });

  if (hasHeader) {
    ensureFileNameHeader_(sheet);
    sheet.setFrozenRows(DATA_HEADER_ROW);
    return;
  }

  const h = normalizeHeaders_(headers);

  sheet
    .getRange(DATA_HEADER_ROW, 1, 1, 10)
    .setValues([h]);

  sheet.setFrozenRows(DATA_HEADER_ROW);
}



/**
 * Visible source-file column on Live Sheet.
 *
 * A:J = existing OT data
 * K:L = intentionally blank spacer columns
 * M   = File Name
 *
 * Existing records are NOT modified/backfilled. New records receive the
 * current PDF filename during the append action.
 */
function ensureFileNameHeader_(sheet) {
  sheet
    .getRange(DATA_HEADER_ROW, FILE_NAME_COL)
    .setValue(FILE_NAME_HEADER)
    .setFontWeight('bold');
}


function normalizeHeaders_(headers) {
  const h =
    Array.isArray(headers) && headers.length
      ? headers.slice(0, 10)
      : [
          'Date',
          'Yearly No.',
          'Monthly No.',
          'Name of Patient',
          'Age/Sex',
          'O.T. No.',
          'Ward',
          'T.C./UHID No.',
          'Diagnosis',
          'Operation'
        ];

  while (h.length < 10) {
    h.push('');
  }

  return h;
}


/**
 * One-time migration for an existing Live Sheet that currently has:
 *   row 1 = OT headers
 *   row 2+ = OT data
 *
 * Inserts two blank rows at the top so the existing header/data moves to:
 *   row 3 = OT headers
 *   row 4+ = OT data
 *
 * Existing data is preserved; nothing is deleted.
 */
function migrateLiveSheetLayout_(sheet) {
  const a1 = String(sheet.getRange('A1').getDisplayValue() || '').trim();
  const b1 = String(sheet.getRange('B1').getDisplayValue() || '').trim();

  // Old layout detection: first cells look like the OT table header.
  const looksLikeOldHeader =
    /^date$/i.test(a1) &&
    /yearly/i.test(b1);

  if (looksLikeOldHeader) {
    sheet.insertRowsBefore(1, 2);
  }
}


/**
 * Creates two permanent date-picker cells on Live Sheet:
 *
 * A1 = From Date     B1 = date picker
 * A2 = To Date       B2 = date picker
 *
 * Users only change B1 and B2. They do not need to create filters or formulas.
 * Clearing either box shows all rows again until both dates are selected.
 */
function ensureDateControls_(sheet) {
  sheet.getRange(FROM_DATE_LABEL_CELL)
    .setValue('From Date')
    .setFontWeight('bold');

  sheet.getRange(TO_DATE_LABEL_CELL)
    .setValue('To Date')
    .setFontWeight('bold');

  const dateRule = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(false)
    .build();

  sheet.getRange(FROM_DATE_CELL)
    .setDataValidation(dateRule)
    .setNumberFormat('dd/MM/yyyy');

  sheet.getRange(TO_DATE_CELL)
    .setDataValidation(dateRule)
    .setNumberFormat('dd/MM/yyyy');

  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 120);
}



/**
 * User-facing behavior:
 * - select From Date in B1
 * - select To Date in B2
 * - rows outside that inclusive date range are hidden automatically
 * - underlying data is never deleted
 * - clear either date box to show all rows
 */
function applyDateRangeFilter_(sheet) {
  const maxRows = sheet.getMaxRows();

  // Always unhide first, so changing the range starts from the full dataset.
  if (maxRows >= DATA_START_ROW) {
    sheet.showRows(
      DATA_START_ROW,
      maxRows - DATA_START_ROW + 1
    );
  }

  const fromDate = parseComparableDate_(
    sheet.getRange(FROM_DATE_CELL).getValue()
  );
  const toDate = parseComparableDate_(
    sheet.getRange(TO_DATE_CELL).getValue()
  );

  // User has not selected both dates yet -> show everything.
  if (!fromDate || !toDate) {
    return;
  }

  let start = fromDate;
  let end = toDate;

  // If dates were selected backwards, handle it gracefully.
  if (start.getTime() > end.getTime()) {
    const temp = start;
    start = end;
    end = temp;
  }

  const lastRow = sheet.getLastRow();

  if (lastRow < DATA_START_ROW) {
    return;
  }

  const dateValues = sheet
    .getRange(
      DATA_START_ROW,
      1,
      lastRow - DATA_START_ROW + 1,
      1
    )
    .getValues();

  const rowsToHide = [];

  dateValues.forEach(function(row, index) {
    const rowNumber = index + DATA_START_ROW;
    const rowDate = parseComparableDate_(row[0]);

    if (
      !rowDate ||
      rowDate.getTime() < start.getTime() ||
      rowDate.getTime() > end.getTime()
    ) {
      rowsToHide.push(rowNumber);
    }
  });

  hideRowsInGroups_(sheet, rowsToHide);
}


/**
 * Parse common OT-register date formats without changing the value stored
 * in column A.
 *
 * Supported examples:
 *   actual Google Sheet Date
 *   21/08/2021
 *   21-08-2021
 *   21.08.2021
 *   2021-08-21
 *   21 Aug 2021
 *   21 August 2021
 *   August 21 2021
 *   January 2021  -> treated as 01/01/2021 for range comparison only
 */
function parseComparableDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return new Date(
      value.getFullYear(),
      value.getMonth(),
      value.getDate()
    );
  }

  const s = String(value == null ? '' : value).trim();

  if (!s) {
    return null;
  }

  let m;

  // dd/mm/yyyy, dd-mm-yyyy, dd.mm.yyyy
  m = s.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2,4})$/);
  if (m) {
    let day = Number(m[1]);
    let month = Number(m[2]);
    let year = Number(m[3]);

    if (year < 100) {
      year += 2000;
    }

    return makeValidDate_(year, month, day);
  }

  // yyyy-mm-dd, yyyy/mm/dd
  m = s.match(/^(\d{4})[\/.\-](\d{1,2})[\/.\-](\d{1,2})$/);
  if (m) {
    return makeValidDate_(
      Number(m[1]),
      Number(m[2]),
      Number(m[3])
    );
  }

  const months = {
    jan: 1, january: 1,
    feb: 2, february: 2,
    mar: 3, march: 3,
    apr: 4, april: 4,
    may: 5,
    jun: 6, june: 6,
    jul: 7, july: 7,
    aug: 8, august: 8,
    sep: 9, sept: 9, september: 9,
    oct: 10, october: 10,
    nov: 11, november: 11,
    dec: 12, december: 12
  };

  // 21 Aug 2021 / 21 August 2021
  m = s.match(
    /^(\d{1,2})\s+([A-Za-z]+)[,\s]+(\d{4})$/
  );
  if (m) {
    const month = months[m[2].toLowerCase()];
    if (month) {
      return makeValidDate_(
        Number(m[3]),
        month,
        Number(m[1])
      );
    }
  }

  // Aug 21 2021 / August 21, 2021
  m = s.match(
    /^([A-Za-z]+)\s+(\d{1,2})[,\s]+(\d{4})$/
  );
  if (m) {
    const month = months[m[1].toLowerCase()];
    if (month) {
      return makeValidDate_(
        Number(m[3]),
        month,
        Number(m[2])
      );
    }
  }

  // January 2021 -> first day of that month, comparison only.
  m = s.match(/^([A-Za-z]+)\s+(\d{4})$/);
  if (m) {
    const month = months[m[1].toLowerCase()];
    if (month) {
      return makeValidDate_(
        Number(m[2]),
        month,
        1
      );
    }
  }

  return null;
}


function makeValidDate_(year, month, day) {
  const d = new Date(year, month - 1, day);

  if (
    d.getFullYear() !== year ||
    d.getMonth() !== month - 1 ||
    d.getDate() !== day
  ) {
    return null;
  }

  return d;
}


/**
 * Hide rows efficiently as contiguous groups rather than one Apps Script
 * call per individual row.
 */
function hideRowsInGroups_(sheet, rows) {
  if (!rows.length) {
    return;
  }

  let start = rows[0];
  let previous = rows[0];

  for (let i = 1; i <= rows.length; i++) {
    const current = i < rows.length ? rows[i] : null;

    if (current !== previous + 1) {
      sheet.hideRows(
        start,
        previous - start + 1
      );

      start = current;
    }

    previous = current;
  }
}


/**
 * Simple edit trigger. Whenever a user changes one of the two date boxes,
 * refresh the visible rows immediately.
 */
function onEdit(e) {
  try {
    if (!e || !e.range) {
      return;
    }

    const sheet = e.range.getSheet();
    const a1 = e.range.getA1Notation();

    if (a1 !== FROM_DATE_CELL && a1 !== TO_DATE_CELL) {
      return;
    }

    if (sheet.getName() === DATA_TAB) {
      applyDateRangeFilter_(sheet);
      return;
    }

    if (sheet.getName() === GK_DATA_TAB) {
      gkApplyDateRangeFilter_(sheet);
      return;
    }
  } catch (err) {
    console.log(
      'Date range filter error: ' +
      String(err && err.message ? err.message : err)
    );
  }
}




// =============================================================================
// GK REGISTER RECEIVER
// =============================================================================

function gkEnsureLayout_(ss) {
  const sheet = ensureSheet_(ss, GK_DATA_TAB);
  const status = ensureSheet_(ss, GK_STATUS_TAB);

  sheet.getRange('A1').setValue('From Date').setFontWeight('bold');
  sheet.getRange('A2').setValue('To Date').setFontWeight('bold');

  const dateRule = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(false)
    .build();

  sheet.getRange('B1').setDataValidation(dateRule).setNumberFormat('dd/MM/yy');
  sheet.getRange('B2').setDataValidation(dateRule).setNumberFormat('dd/MM/yy');

  // One-time, non-destructive migration of the earlier GK layout
  // (Name + phone together, Age/Sex together, old abbreviations).
  gkMigrateLegacyLayout_(sheet);
  gkNormalizeExistingDates_(sheet);

  sheet.getRange(GK_DATA_HEADER_ROW, 1, 1, GK_HEADERS.length)
    .setValues([GK_HEADERS]);
  sheet.getRange(GK_DATA_HEADER_ROW, GK_FILE_NAME_COL)
    .setValue('File Name');

  sheet.getRange(GK_DATA_HEADER_ROW, 1, 1, GK_HEADERS.length)
    .setFontWeight('bold');
  sheet.getRange(GK_DATA_HEADER_ROW, GK_FILE_NAME_COL)
    .setFontWeight('bold');

  sheet.setFrozenRows(GK_DATA_HEADER_ROW);

  // T:U remain empty spacer columns between extracted data and File Name.
  sheet.setColumnWidth(1, 95);   // Serial Number
  sheet.setColumnWidth(2, 100);  // Date
  sheet.setColumnWidth(3, 145);  // Gamma Knife Number
  sheet.setColumnWidth(4, 125);  // UHID Number
  sheet.setColumnWidth(5, 220);  // Patient Name
  sheet.setColumnWidth(6, 190);  // Mobile Number(s)
  sheet.setColumnWidth(7, 70);   // Age
  sheet.setColumnWidth(8, 70);   // Sex
  sheet.setColumnWidth(9, 220);  // Diagnosis
  sheet.setColumnWidth(10, 80);  // Mask
  sheet.setColumnWidth(11, 90);  // Dose
  sheet.setColumnWidth(12, 95);  // Percentage
  sheet.setColumnWidth(13, 110); // Volume (cc)
  sheet.setColumnWidth(14, 90);  // GURA
  sheet.setColumnWidth(15, 120); // Consultant
  sheet.setColumnWidth(16, 120); // Physician
  sheet.setColumnWidth(17, 120); // Residence
  sheet.setColumnWidth(18, 90);  // RTT
  sheet.setColumnWidth(19, 100); // Technique
  sheet.setColumnWidth(GK_FILE_NAME_COL, 220);


  status.setColumnWidth(1, 150);
  status.setColumnWidth(2, 320);
  return sheet;
}


function gkNormalizePhone_(value) {
  const digits = String(value || '').replace(/\D/g, '');
  if (digits.length === 12 && digits.indexOf('91') === 0) {
    return '+91' + digits.substring(2);
  }
  if (digits.length === 11 && digits.charAt(0) === '0') {
    return digits.substring(1);
  }
  return digits;
}


function gkSplitNameMobile_(value) {
  let name = String(value == null ? '' : value).trim();
  const found = [];
  const re = /(?<!\d)(?:(?:\+?91)[\s-]?)?[6-9](?:[\s-]?\d){9}(?!\d)/g;
  const matches = name.match(re) || [];

  matches.forEach(function(m) {
    const normalized = gkNormalizePhone_(m);
    if (normalized && found.indexOf(normalized) === -1) {
      found.push(normalized);
    }
  });

  name = name.replace(re, ' ').replace(/\s+/g, ' ').trim();
  name = name.replace(/^[,;|\-\s]+|[,;|\-\s]+$/g, '');

  return {
    name: name,
    mobile: found.join(', ')
  };
}


function gkSplitAgeSex_(value) {
  const text = String(value == null ? '' : value).trim();
  let age = '';
  let sex = '';

  const ageMatch = text.match(/(?:^|\D)(\d{1,3})(?:\D|$)/);
  if (ageMatch) {
    age = ageMatch[1];
  }

  const sexMatch = text.match(/(?:^|[^A-Za-z])(male|female|m|f)(?:$|[^A-Za-z])/i);
  if (sexMatch) {
    const token = sexMatch[1].toLowerCase();
    sex = (token === 'm' || token === 'male') ? 'M' : 'F';
  }

  return {age: age, sex: sex};
}


function gkVolumeCc_(value) {
  let text = String(value == null ? '' : value).trim();
  if (!text) {
    return '';
  }
  text = text.replace(/\s*(?:c\.?c\.?|cm\s*3|cm\^3)\s*$/i, '').trim();
  return text ? text + ' cc' : '';
}


/**
 * Migrate rows created by the original GK schema:
 * A:S = S.No, Date, GK, UHID, Name, Age/Sex, Diagnosis, Dose, Perc,
 *       Volume, GURA, CONS, PHY, R.O, RTT, Technique, blank, blank, File Name
 *
 * New schema is A:V and keeps the existing data while splitting Name/Mobile and
 * Age/Sex. Mask is blank for legacy rows because the old extraction did not
 * capture it. New uploads will populate it.
 */
function gkMigrateLegacyLayout_(sheet) {
  const oldHeaders = sheet
    .getRange(GK_DATA_HEADER_ROW, 1, 1, 19)
    .getDisplayValues()[0]
    .map(function(v) { return String(v || '').trim().toLowerCase(); });

  const isOld =
    oldHeaders[0] === 's.no' &&
    oldHeaders[5] === 'age/sex' &&
    oldHeaders[15] === 'technique';

  if (!isOld) {
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < GK_DATA_START_ROW) {
    return;
  }

  const count = lastRow - GK_DATA_START_ROW + 1;
  const oldRows = sheet
    .getRange(GK_DATA_START_ROW, 1, count, 19)
    .getDisplayValues();

  const migrated = oldRows.map(function(old) {
    const nm = gkSplitNameMobile_(old[4]);
    const asx = gkSplitAgeSex_(old[5]);

    return [
      old[0],              // A Serial Number
      old[1],              // B Date
      old[2],              // C Gamma Knife Number
      old[3],              // D UHID Number
      nm.name,             // E Patient Name
      nm.mobile,           // F Mobile Number(s)
      asx.age,             // G Age
      asx.sex,             // H Sex
      old[6],              // I Diagnosis
      '',                  // J Mask - unavailable in legacy extraction
      old[7],              // K Dose
      old[8],              // L Percentage
      gkVolumeCc_(old[9]), // M Volume (cc)
      old[10],             // N GURA
      old[11],             // O Consultant
      old[12],             // P Physician
      old[13],             // Q Residence
      old[14],             // R RTT
      old[15],             // S Technique
      '',                  // T spacer
      '',                  // U spacer
      old[18]              // V File Name
    ];
  });

  sheet
    .getRange(GK_DATA_START_ROW, 1, count, GK_FILE_NAME_COL)
    .clearContent();
  sheet
    .getRange(GK_DATA_START_ROW, 1, count, GK_FILE_NAME_COL)
    .setValues(migrated);
}


function handleGK_(ss, body, props) {
  const data = gkEnsureLayout_(ss);
  const status = ensureSheet_(ss, GK_STATUS_TAB);
  const action = String(body.action || '');

  // Save every upload to the configured restricted Drive folder BEFORE any
  // model processing. This makes the queue durable even if Streamlit sleeps.
  if (action === 'enqueue_upload') {
    const result = gkEnqueueUpload_(ss, body, props);
    // Self-heal the recurring worker trigger on every upload. This protects
    // against an interrupted/timeout setup run that saved properties but
    // never reached trigger creation.
    gkEnsureWorkerTrigger_();
    gkKickWorker_();
    SpreadsheetApp.flush();
    return json_(result);
  }

  if (action === 'queue_status') {
    return json_(gkQueueStatus_(ss, body));
  }

  if (action === 'queue_retry_now') {
    return json_(gkRetryQueueNow_(ss, body, props));
  }

  if (action === 'reset') {
    const inputName = String(body.input_name || '');
    const runId = String(body.run_id || '');
    const currentRunId = props.getProperty('GK_CURRENT_RUN_ID') || '';
    const isResume =
      runId !== '' &&
      currentRunId === runId &&
      hasTopStatusBlock_(status);

    props.setProperty('GK_CURRENT_INPUT_NAME', inputName);
    props.setProperty('GK_CURRENT_RUN_ID', runId);

    if (!isResume) {
      startNewStatusBlock_(status);
    }

    writeStatusBlock_(
      status,
      inputName,
      isResume ? 'Resumed' : 'Started',
      Number(body.page_number || 0) + '/' + Number(body.total_pages || 0),
      Number(body.total_records || 0),
      isResume ? 'Resumed; existing S.No rows will be updated, not duplicated' : 'Started'
    );

    gkSortData_(data);
    gkApplyDateRangeFilter_(data);
    SpreadsheetApp.flush();
    return json_({ok: true, action: 'reset', dataset: 'gk'});
  }

  if (action === 'upsert') {
    const rows = Array.isArray(body.rows) ? body.rows : [];
    const sourceFileName =
      String(body.input_name || '') ||
      props.getProperty('GK_CURRENT_INPUT_NAME') || '';

    const result = gkUpsertRows_(data, rows, sourceFileName, String(body.file_url || ''));

    // Keep the GK register in natural register order after every upload:
    // S.No ascending first, then Date ascending as a tie-breaker.
    gkSortData_(data);

    writeStatusBlock_(
      status,
      sourceFileName,
      'Processing',
      Number(body.page_number || 0) + '/' + Number(body.total_pages || 0),
      Number(body.total_records || 0),
      'Inserted ' + result.inserted + ', updated ' + result.updated +
        (result.cleared ? ', crossed-out values cleared ' + result.cleared : '') +
        (result.date_conflicts ? ', protected date conflicts ' + result.date_conflicts : '')
    );

    gkApplyDateRangeFilter_(data);
    SpreadsheetApp.flush();
    return json_({
      ok: true,
      action: 'upsert',
      dataset: 'gk',
      inserted: result.inserted,
      updated: result.updated,
      cleared: result.cleared,
      date_conflicts: result.date_conflicts
    });
  }

  if (action === 'status') {
    const inputName =
      String(body.input_name || '') ||
      props.getProperty('GK_CURRENT_INPUT_NAME') || '';

    writeStatusBlock_(
      status,
      inputName,
      String(body.state || 'Processing'),
      Number(body.page_number || 0) + '/' + Number(body.total_pages || 0),
      Number(body.total_records || 0),
      String(body.message || '')
    );
    SpreadsheetApp.flush();
    return json_({ok: true, action: 'status', dataset: 'gk'});
  }

  if (action === 'finish') {
    writeStatusBlock_(
      status,
      props.getProperty('GK_CURRENT_INPUT_NAME') || String(body.input_name || ''),
      'Done',
      Number(body.total_pages || 0) + '/' + Number(body.total_pages || 0),
      Number(body.total_records || 0),
      'Extraction complete'
    );
    SpreadsheetApp.flush();
    return json_({ok: true, action: 'finish', dataset: 'gk'});
  }

  return json_({ok: false, error: 'Unknown GK action'});
}


function gkNorm_(value) {
  return String(value == null ? '' : value)
    .trim()
    .toLowerCase()
    .replace(/\s+/g, '')
    .replace(/[.]+$/g, '');
}


function gkFallbackKey_(row) {
  // Used only when S.No is missing: Date + GK + UHID + Name.
  return [row[1], row[2], row[3], row[4]].map(gkNorm_).join('|');
}


function gkYearFromDate_(value) {
  const d = parseComparableDate_(value);
  return d ? String(d.getFullYear()) : '';
}


function gkYearSnoKey_(sno, dateValue) {
  const serial = gkNorm_(sno);
  const year = gkYearFromDate_(dateValue);
  return serial && year ? year + '|' + serial : '';
}


function gkNormalizeDateText_(value) {
  const d = parseComparableDate_(value);
  if (!d) {
    return '';
  }
  const dd = String(d.getDate()).padStart(2, '0');
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const yy = String(d.getFullYear()).slice(-2);
  return dd + '/' + mm + '/' + yy;
}


function gkNormalizeExistingDates_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < GK_DATA_START_ROW) {
    return;
  }

  const range = sheet.getRange(
    GK_DATA_START_ROW,
    2,
    lastRow - GK_DATA_START_ROW + 1,
    1
  );
  const values = range.getValues();

  values.forEach(function(row) {
    const raw = String(row[0] == null ? '' : row[0]).trim();
    if (!raw) {
      row[0] = '';
      return;
    }
    const normalized = gkNormalizeDateText_(row[0]);
    if (normalized) {
      row[0] = normalized;
    }
  });

  range.setValues(values);
  range.setNumberFormat('@');
}




function gkSetSourceFileCell_(sheet, rowNumber, sourceFileName, sourceFileUrl) {
  const cell = sheet.getRange(rowNumber, GK_FILE_NAME_COL);
  const label = String(sourceFileName || '').trim();
  const url = String(sourceFileUrl || '').trim();

  if (!url) {
    cell.setValue(label);
    return;
  }

  const rich = SpreadsheetApp.newRichTextValue()
    .setText(label || 'Open source file')
    .setLinkUrl(url)
    .build();
  cell.setRichTextValue(rich);
}


function gkUpsertRows_(sheet, rows, sourceFileName, sourceFileUrl) {
  const lastRow = sheet.getLastRow();
  const existingCount = Math.max(0, lastRow - GK_DATA_START_ROW + 1);
  const existing = existingCount > 0
    ? sheet.getRange(GK_DATA_START_ROW, 1, existingCount, GK_HEADERS.length).getDisplayValues()
    : [];

  // S.No restarts from 1 every calendar year, so S.No alone is NOT a safe ID.
  // Primary identity is now: YEAR(Date) + S.No.
  const byYearSno = {};
  const byFallback = {};

  existing.forEach(function(row, i) {
    const sheetRow = GK_DATA_START_ROW + i;
    const identity = gkYearSnoKey_(row[0], row[1]);

    if (identity && !byYearSno[identity]) {
      byYearSno[identity] = sheetRow;
    }

    const fb = gkFallbackKey_(row);
    if (fb.replace(/\|/g, '') && !byFallback[fb]) {
      byFallback[fb] = sheetRow;
    }
  });

  let inserted = 0;
  let updated = 0;
  let cleared = 0;
  let dateConflicts = 0;

  rows.forEach(function(raw) {
    // New clients send {values: [...], clear_fields: [...]}. For backwards
    // compatibility, older clients that send a plain row array still work.
    let clearFields = [];
    let row = [];

    if (Array.isArray(raw)) {
      row = raw.slice(0, GK_HEADERS.length);
    } else if (raw && typeof raw === 'object') {
      row = Array.isArray(raw.values)
        ? raw.values.slice(0, GK_HEADERS.length)
        : [];
      clearFields = Array.isArray(raw.clear_fields)
        ? raw.clear_fields.map(function(v) {
            return String(v == null ? '' : v).trim().toLowerCase();
          })
        : [];
    }

    while (row.length < GK_HEADERS.length) {
      row.push('');
    }

    row = row.map(function(v) {
      return String(v == null ? '' : v).trim();
    });

    const clearSet = {};
    clearFields.forEach(function(field) {
      if (GK_FIELD_KEYS.indexOf(field) >= 0) {
        clearSet[field] = true;
      }
    });

    if (row[1]) {
      row[1] = gkNormalizeDateText_(row[1]);
    }

    const sno = gkNorm_(row[0]);
    const identity = gkYearSnoKey_(row[0], row[1]);
    const fb = gkFallbackKey_(row);

    // IMPORTANT:
    // - Same YEAR + same S.No => update the existing record.
    // - Same S.No + different YEAR => separate record; never overwrite.
    // - If S.No exists but Date/year is unreadable, never match by S.No alone.
    //   This intentionally prefers a possible duplicate over a destructive
    //   cross-year overwrite.
    // - The fallback key is used only when S.No itself is missing.
    let target = identity && byYearSno[identity]
      ? byYearSno[identity]
      : null;

    if (
      !target &&
      !sno &&
      fb.replace(/\|/g, '') &&
      byFallback[fb]
    ) {
      target = byFallback[fb];
    }

    if (target) {
      const old = sheet
        .getRange(target, 1, 1, GK_HEADERS.length)
        .getDisplayValues()[0];

      const oldDate = String(old[1] || '').trim();
      const newDate = String(row[1] || '').trim();

      // Because target matching is already restricted to the same calendar
      // year + S.No, a conflicting day/month within that year is treated as
      // OCR conflict and the existing date is preserved.
      if (
        !clearSet.date &&
        oldDate &&
        newDate &&
        gkNorm_(oldDate) !== gkNorm_(newDate)
      ) {
        row[1] = oldDate;
        dateConflicts += 1;
      }

      const merged = old.map(function(v, idx) {
        const field = GK_FIELD_KEYS[idx] || '';
        const incoming = String(row[idx] || '').trim();

        // Explicit cross-out is intentional deletion. Ordinary blank OCR still
        // preserves the old value.
        if (field && clearSet[field]) {
          if (String(v || '').trim()) {
            cleared += 1;
          }
          return '';
        }

        return incoming ? incoming : String(v || '');
      });

      sheet.getRange(target, 1, 1, GK_HEADERS.length).setValues([merged]);
      gkSetSourceFileCell_(sheet, target, sourceFileName, sourceFileUrl);
      updated += 1;
    } else {
      const targetRow = Math.max(sheet.getLastRow() + 1, GK_DATA_START_ROW);
      sheet.getRange(targetRow, 1, 1, GK_HEADERS.length).setValues([row]);
      gkSetSourceFileCell_(sheet, targetRow, sourceFileName, sourceFileUrl);
      inserted += 1;

      if (identity) {
        byYearSno[identity] = targetRow;
      }

      // Only maintain fallback matching for rows without S.No.
      if (!sno && fb.replace(/\|/g, '')) {
        byFallback[fb] = targetRow;
      }
    }
  });

  return {
    inserted: inserted,
    updated: updated,
    cleared: cleared,
    date_conflicts: dateConflicts
  };
}

/**
 * Sort GK data rows while keeping rows 1-3 untouched.
 *
 * Primary order: Calendar year from Date (ascending)
 * Secondary order: Serial Number within that year (numeric ascending)
 * Tertiary order: Date (ascending)
 *
 * The full A:S row is moved together, so UHID, patient name, diagnosis,
 * technique and File Name always stay attached to the correct S.No.
 */
function gkSortData_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < GK_DATA_START_ROW) {
    return;
  }

  const rowCount = lastRow - GK_DATA_START_ROW + 1;
  if (rowCount <= 1) {
    return;
  }

  // A:V includes the extracted fields, the blank spacer columns T:U,
  // and the File Name in V.
  const range = sheet.getRange(
    GK_DATA_START_ROW,
    1,
    rowCount,
    GK_FILE_NAME_COL
  );

  const values = range.getValues();
  const richFileCells = sheet
    .getRange(GK_DATA_START_ROW, GK_FILE_NAME_COL, rowCount, 1)
    .getRichTextValues();

  const sortableRows = values.map(function(row, idx) {
    return {
      values: row,
      richFile: richFileCells[idx][0]
    };
  });

  function snoNumber_(value) {
    const text = String(value == null ? '' : value).trim();
    const match = text.match(/\d+/);
    return match ? Number(match[0]) : Number.POSITIVE_INFINITY;
  }

  function dateNumber_(value) {
    const parsed = parseComparableDate_(value);
    return parsed ? parsed.getTime() : Number.POSITIVE_INFINITY;
  }

  sortableRows.sort(function(aItem, bItem) {
    const a = aItem.values;
    const b = bItem.values;

    const aDate = dateNumber_(a[1]);
    const bDate = dateNumber_(b[1]);
    const aYear = isFinite(aDate) ? new Date(aDate).getFullYear() : Number.POSITIVE_INFINITY;
    const bYear = isFinite(bDate) ? new Date(bDate).getFullYear() : Number.POSITIVE_INFINITY;

    if (aYear !== bYear) {
      return aYear - bYear;
    }

    const aSno = snoNumber_(a[0]);
    const bSno = snoNumber_(b[0]);

    if (aSno !== bSno) {
      return aSno - bSno;
    }

    if (aDate !== bDate) {
      return aDate - bDate;
    }

    return String(a[0] || '').localeCompare(String(b[0] || ''));
  });

  range.setValues(sortableRows.map(function(item) { return item.values; }));
  sheet
    .getRange(GK_DATA_START_ROW, GK_FILE_NAME_COL, rowCount, 1)
    .setRichTextValues(sortableRows.map(function(item) { return [item.richFile]; }));
}


/**
 * Run this manually ONCE from Apps Script if you want to immediately sort
 * the rows that are already present, without waiting for the next upload.
 */
function sortGKNow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Prefer the dedicated GK tab used by this receiver.
  let sheet = ss.getSheetByName(GK_DATA_TAB);

  // If this workbook was created earlier and the GK register is in "Live Sheet",
  // use it only when row 3 clearly contains the GK headers.
  if (!sheet) {
    const candidate = ss.getSheetByName(DATA_TAB);
    if (candidate) {
      const h = candidate
        .getRange(GK_DATA_HEADER_ROW, 1, 1, 3)
        .getDisplayValues()[0]
        .map(function(v) { return String(v || '').trim().toLowerCase(); });

      if (
        (h[0] === 's.no' || h[0] === 'serial number') &&
        h[1] === 'date' &&
        (h[2] === 'g.k. no' || h[2] === 'g.k. no.' || h[2] === 'gamma knife number')
      ) {
        sheet = candidate;
      }
    }
  }

  if (!sheet) {
    throw new Error(
      'GK data sheet not found. Expected "' +
      GK_DATA_TAB +
      '" or a GK-formatted "' +
      DATA_TAB +
      '" tab.'
    );
  }

  gkSortData_(sheet);
  gkApplyDateRangeFilter_(sheet);
  SpreadsheetApp.flush();
}


// =============================================================================
// GK DRIVE-BACKED PERSISTENT QUEUE
// =============================================================================

const GK_BACKGROUND_PROMPT = `
You are transcribing a handwritten/scanned Gamma Knife register image or PDF.
Accuracy and row alignment are more important than guessing.

Return JSON ONLY in this exact top-level form:
{"records":[{...}]}

Each patient record must contain ALL of these fields:
source_row, sno, date, gk_no, uhid_no, patient_name, mobile_no, age, sex,
diagnosis, mask, dose, perc, volume, gura, cons, phy, ro, rtt,
crossed_out_fields.

GENERAL:
- One physical patient row = one record.
- Never create two patient records because a row has multiple masks/plans.
- If the whole patient row is clearly cancelled/crossed out, omit it entirely.
- If text wraps inside the same ruled row, combine it into the same record.
- Do not infer unreadable values from nearby rows. Use an empty string instead.
- Ignore camera timestamps, page headings, month titles and printed ruling.
- Handle skewed/rotated/upside-down images visually before reading.

NAME + MOBILE:
- Patient name cell may also contain one or more mobile numbers.
- patient_name = name only.
- mobile_no = all visible mobile numbers from that name cell, comma-separated.
- Never treat S.No, GK No or UHID as a mobile number.

AGE + SEX:
- Split combined values such as 55y/M or 36/F.
- age = age only, sex = M/F/Male/Female only as visibly written.
- Do not infer sex.

DATE:
- Read Date only from that row's Date cell.
- Never carry a date from another row or heading.
- Output every readable date as DD/MM/YY only.
- If unreadable or blank, output "".

GAMMA KNIFE REPEAT:
- Preserve any visible R/(R)/r marker with the number. Post-processing will
  normalize forms such as 7282 (R) to Repeat 7282.

DIAGNOSIS TECHNIQUE MARKER:
- A standalone (M) in Diagnosis means Technique = Masked.
- A standalone (V) in Diagnosis means Technique = Vantage.
- Preserve the visible (M)/(V) inside diagnosis; post-processing removes it.
- M1/M2/M3 mask labels are NOT this diagnosis marker.

MULTIPLE MASK / PLAN -- CRITICAL:
- A patient row may have M1, M2, M3 or similar multiple plan entries.
- Ignore every crossed/cancelled mask entry.
- Keep ONLY the LAST UNCROSSED mask/plan visible in that patient row.
- mask, dose, perc, volume, gura, cons, phy, ro and rtt must ALL come from that
  SAME last uncrossed plan. Never mix values from different plans.
- If there is only one treatment entry and no explicit M label, mask may be blank.

CROSSED / CANCELLED VALUES -- CRITICAL:
- Any value crossed by a pen line, strike-through, scribble, X or clear
  cancellation mark is invalid and must NOT be transcribed.
- If a crossed value has an uncrossed replacement in the same cell/row, return
  only the uncrossed replacement and DO NOT mark that field crossed_out.
- If a field has only a crossed/cancelled value and no replacement, return ""
  for that field and add the canonical field key to crossed_out_fields.
- Allowed crossed_out_fields keys are:
  sno,date,gk_no,uhid_no,patient_name,mobile_no,age,sex,diagnosis,mask,dose,
  perc,volume,gura,cons,phy,ro,rtt,technique.

RIGHT SIDE HEADERS:
- CONS = Consultant
- PHY = Physician
- R.O = Residence
- GURA and RTT: transcribe literally; do not invent expansions.

VOLUME:
- Read volume from the selected last plan. Keep the numeric reading accurately.
  Post-processing will append/normalize cc.

TRANSCRIPTION:
- Copy handwriting literally; do not medically correct names/diagnoses.
- Keep visible units such as Gy and % where legible.
`;


function configureGKBackground() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('Open this Apps Script from the target Google Sheet.');
  }

  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();

  const folderPrompt = ui.prompt(
    'GK Drive folder',
    'Paste the restricted Google Drive folder URL or folder ID. Leave blank to create a new private folder named "GK Register Uploads".',
    ui.ButtonSet.OK_CANCEL
  );
  if (folderPrompt.getSelectedButton() !== ui.Button.OK) return;

  let folderId = gkExtractDriveId_(folderPrompt.getResponseText());
  if (!folderId) {
    const folder = DriveApp.createFolder('GK Register Uploads');
    folderId = folder.getId();
  }
  DriveApp.getFolderById(folderId); // permission/access validation

  const key1Prompt = ui.prompt(
    'API key 1',
    'Paste GEMINI_API_KEY_1. It will be stored in Apps Script Properties, not in the Sheet.',
    ui.ButtonSet.OK_CANCEL
  );
  if (key1Prompt.getSelectedButton() !== ui.Button.OK) return;
  const key1 = String(key1Prompt.getResponseText() || '').trim();
  if (!key1) throw new Error('API key 1 is required.');

  const key2Prompt = ui.prompt(
    'API key 2 (optional)',
    'Paste GEMINI_API_KEY_2, or leave blank if you only use one key.',
    ui.ButtonSet.OK_CANCEL
  );
  if (key2Prompt.getSelectedButton() !== ui.Button.OK) return;
  const key2 = String(key2Prompt.getResponseText() || '').trim();

  const defaultModel = props.getProperty('GEMINI_MODEL') || 'gemini-3.8-flash';
  const modelPrompt = ui.prompt(
    'Model',
    'Model name. Current value: ' + defaultModel + '. Leave blank to keep it.',
    ui.ButtonSet.OK_CANCEL
  );
  if (modelPrompt.getSelectedButton() !== ui.Button.OK) return;
  const model = String(modelPrompt.getResponseText() || '').trim() || defaultModel;

  props.setProperty('SPREADSHEET_ID', ss.getId());
  props.setProperty('GK_DRIVE_FOLDER_ID', folderId);
  props.setProperty('GEMINI_API_KEY_1', key1);
  if (key2) {
    props.setProperty('GEMINI_API_KEY_2', key2);
  } else {
    props.deleteProperty('GEMINI_API_KEY_2');
  }
  props.setProperty('GEMINI_MODEL', model);

  gkEnsureQueue_(ss);
  gkEnsureWorkerTrigger_();

  const folderUrl = DriveApp.getFolderById(folderId).getUrl();
  ui.alert(
    'GK background processing configured.\n\n' +
    'Drive folder: ' + folderUrl + '\n' +
    'Worker checks the persistent queue automatically every ' +
    GK_WORKER_EVERY_MINUTES + ' minutes.'
  );
}


function gkExtractDriveId_(value) {
  const text = String(value || '').trim();
  if (!text) return '';
  const match = text.match(/[-_A-Za-z0-9]{20,}/);
  return match ? match[0] : '';
}


function gkEnsureQueue_(ss) {
  let sheet = ss.getSheetByName(GK_QUEUE_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(GK_QUEUE_TAB);
  }

  sheet
    .getRange(1, 1, 1, GK_QUEUE_HEADERS.length)
    .setValues([GK_QUEUE_HEADERS])
    .setFontWeight('bold');
  sheet.setFrozenRows(1);

  try {
    sheet.hideSheet();
  } catch (_err) {}

  return sheet;
}


function gkWorkerTriggerActive_() {
  return ScriptApp.getProjectTriggers().some(function(trigger) {
    return trigger.getHandlerFunction() === GK_WORKER_HANDLER;
  });
}


function gkUiText_(value) {
  // UI-only wording. Keep internal API/model identifiers unchanged.
  return String(value == null ? '' : value)
    .replace(/gemini-/gi, 'LLM-')
    .replace(/\bGemini\b/gi, 'LLM');
}


function gkEnsureWorkerTrigger_() {
  const triggers = ScriptApp.getProjectTriggers();
  const exists = triggers.some(function(trigger) {
    return trigger.getHandlerFunction() === GK_WORKER_HANDLER;
  });

  if (!exists) {
    ScriptApp.newTrigger(GK_WORKER_HANDLER)
      .timeBased()
      .everyMinutes(GK_WORKER_EVERY_MINUTES)
      .create();
  }
}


/**
 * Manual repair utility. Run this once if jobs remain Pending/Retry with
 * Attempts unchanged. It only touches GK queue worker triggers.
 */
function repairGKWorkerTrigger() {
  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(function(trigger) {
    const handler = trigger.getHandlerFunction();
    if (handler === GK_WORKER_HANDLER || handler === GK_KICK_HANDLER) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger(GK_WORKER_HANDLER)
    .timeBased()
    .everyMinutes(GK_WORKER_EVERY_MINUTES)
    .create();

  // Start one immediate pass so an already-eligible Retry/Pending job does
  // not need to wait for the next periodic window.
  ScriptApp.newTrigger(GK_KICK_HANDLER)
    .timeBased()
    .after(1000)
    .create();

  console.log('GK worker trigger repaired and immediate queue kick scheduled.');
}


function gkKickWorker_() {
  // Any kick also verifies that the permanent 5-minute worker exists.
  gkEnsureWorkerTrigger_();
  const triggers = ScriptApp.getProjectTriggers();
  const exists = triggers.some(function(trigger) {
    return trigger.getHandlerFunction() === GK_KICK_HANDLER;
  });

  if (!exists) {
    ScriptApp.newTrigger(GK_KICK_HANDLER)
      .timeBased()
      .after(1000)
      .create();
  }
}


function processGKQueueKick() {
  processGKQueueTrigger();
}


function gkDriveFolder_() {
  const folderId = PropertiesService.getScriptProperties().getProperty('GK_DRIVE_FOLDER_ID');
  if (!folderId) {
    throw new Error('GK_DRIVE_FOLDER_ID is not configured. Run configureGKBackground().');
  }
  return DriveApp.getFolderById(folderId);
}


function gkEnqueueUpload_(ss, body, props) {
  const fileName = String(body.file_name || '').trim();
  const mimeType = String(body.mime_type || 'application/octet-stream').trim();
  const data64 = String(body.data_base64 || '');
  const sha256 = String(body.content_sha256 || '').trim().toLowerCase();

  if (!fileName) throw new Error('file_name is required');
  if (!data64) throw new Error('data_base64 is required');

  const queue = gkEnsureQueue_(ss);

  // Idempotency: clicking the Streamlit button twice does not create a second
  // Drive copy when the exact same file hash is already queued.
  if (sha256 && queue.getLastRow() >= GK_QUEUE_START_ROW) {
    const rows = queue
      .getRange(GK_QUEUE_START_ROW, 1, queue.getLastRow() - GK_QUEUE_START_ROW + 1, GK_QUEUE_HEADERS.length)
      .getDisplayValues();

    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i][5] || '').toLowerCase() === sha256) {
        const dupStatus = String(rows[i][6] || 'Pending');
        const dupMessage = String(rows[i][11] || '');
        return {
          ok: true,
          action: 'enqueue_upload',
          duplicate_upload: true,
          queue_id: rows[i][0],
          file_name: rows[i][2],
          drive_url: rows[i][4],
          status: dupStatus,
          reason: dupMessage || (dupStatus === 'Pending' ? 'Waiting for background worker' : '')
        };
      }
    }
  }

  const bytes = Utilities.base64Decode(data64);
  const blob = Utilities.newBlob(bytes, mimeType, fileName);
  const file = gkDriveFolder_().createFile(blob);
  const queueId = Utilities.getUuid();
  const now = new Date();

  const quotaBlockedUntil = Number(props.getProperty('GK_QUOTA_BLOCK_UNTIL') || 0);
  const initialReason =
    quotaBlockedUntil && Date.now() < quotaBlockedUntil
      ? 'Waiting because processing is paused for API quota until ' +
        gkFormatIst_(new Date(quotaBlockedUntil))
      : 'Waiting for background worker';

  queue.appendRow([
    queueId,
    file.getId(),
    fileName,
    mimeType,
    file.getUrl(),
    sha256,
    'Pending',
    0,
    now,
    now,
    now,
    initialReason,
    0
  ]);

  return {
    ok: true,
    action: 'enqueue_upload',
    queue_id: queueId,
    file_name: fileName,
    drive_url: file.getUrl(),
    status: 'Pending',
    reason: initialReason
  };
}


function gkQueueStatus_(ss, body) {
  const wanted = {};
  (Array.isArray(body.queue_ids) ? body.queue_ids : []).forEach(function(id) {
    wanted[String(id)] = true;
  });

  // Optional history filter from Streamlit, in IST calendar date format YYYY-MM-DD.
  // This lets a fresh/reloaded browser recover uploads from any selected day
  // without relying on Streamlit session_state.
  const dateIst = String(body.date_ist || '').trim();

  const props = PropertiesService.getScriptProperties();
  const quotaBlockedUntil = Number(props.getProperty('GK_QUOTA_BLOCK_UNTIL') || 0);
  const quotaPauseActive = quotaBlockedUntil && Date.now() < quotaBlockedUntil;
  const workerTriggerActive = ScriptApp.getProjectTriggers().some(function(trigger) {
    return trigger.getHandlerFunction() === GK_WORKER_HANDLER;
  });

  const queue = gkEnsureQueue_(ss);
  const items = [];
  if (queue.getLastRow() < GK_QUEUE_START_ROW) {
    return {ok: true, action: 'queue_status', worker_trigger_active: workerTriggerActive, items: items};
  }

  const rows = queue
    .getRange(GK_QUEUE_START_ROW, 1, queue.getLastRow() - GK_QUEUE_START_ROW + 1, GK_QUEUE_HEADERS.length)
    .getValues();

  rows.forEach(function(row) {
    if (Object.keys(wanted).length && !wanted[String(row[0])]) return;

    const addedAt = row[8] instanceof Date ? row[8] : null;
    if (dateIst) {
      const rowDateIst = addedAt
        ? Utilities.formatDate(addedAt, 'Asia/Kolkata', 'yyyy-MM-dd')
        : '';
      if (rowDateIst !== dateIst) return;
    }

    const status = String(row[6] || '');
    const attempts = Number(row[7] || 0);
    const nextAttempt = row[10] instanceof Date ? row[10] : null;
    const storedMessage = String(row[11] || '');
    const records = Number(row[12] || 0);
    let reason = storedMessage;

    if (status === 'Pending') {
      if (!workerTriggerActive) {
        reason = 'Background worker trigger is missing. Run repairGKWorkerTrigger once.';
      } else if (quotaPauseActive) {
        reason = 'Waiting because processing is paused for API quota until ' +
          gkFormatIst_(new Date(quotaBlockedUntil));
      } else if (!reason) {
        reason = 'Waiting for background worker';
      }
    } else if (status === 'Processing') {
      reason = 'Extraction in progress';
    } else if (status === 'Waiting for quota') {
      reason = storedMessage || 'API quota/rate limit reached';
      if (nextAttempt) {
        reason += ' | Eligible after: ' + gkFormatIst_(nextAttempt);
      }
      if (!workerTriggerActive) {
        reason += ' | Worker trigger missing: run repairGKWorkerTrigger';
      }
    } else if (status === 'Retry') {
      reason = storedMessage || 'Temporary service error';
      if (nextAttempt) {
        reason += ' | Eligible after: ' + gkFormatIst_(nextAttempt);
      }
      if (!workerTriggerActive) {
        reason += ' | Worker trigger missing: run repairGKWorkerTrigger';
      }
    } else if (status === 'Completed') {
      reason = records + ' record(s) extracted';
      if (storedMessage) {
        reason += ' | ' + storedMessage;
      }
    } else if (status === 'Error') {
      reason = storedMessage || 'Processing failed';
    }

    reason = gkUiText_(reason);

    items.push({
      queue_id: String(row[0] || ''),
      drive_file_id: String(row[1] || ''),
      file_name: String(row[2] || ''),
      drive_url: String(row[4] || ''),
      status: status,
      reason: reason,
      attempts: attempts,
      added_at: addedAt ? gkFormatIst_(addedAt) : '',
      last_update: gkFormatIst_(row[9]),
      next_attempt: nextAttempt ? gkFormatIst_(nextAttempt) : '',
      last_error: storedMessage,
      records: records
    });
  });

  return {
    ok: true,
    action: 'queue_status',
    date_ist: dateIst,
    worker_trigger_active: workerTriggerActive,
    items: items
  };
}


function gkFormatIst_(value) {
  if (!(value instanceof Date) || isNaN(value.getTime())) return String(value || '');
  return Utilities.formatDate(value, 'Asia/Kolkata', 'dd/MM/yy HH:mm:ss');
}


function gkQueueSelectJob_(ss) {
  const queue = gkEnsureQueue_(ss);
  const lastRow = queue.getLastRow();
  if (lastRow < GK_QUEUE_START_ROW) return null;

  const values = queue
    .getRange(GK_QUEUE_START_ROW, 1, lastRow - GK_QUEUE_START_ROW + 1, GK_QUEUE_HEADERS.length)
    .getValues();
  const now = new Date();

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const status = String(row[6] || '').trim();
    const lastUpdate = row[9] instanceof Date ? row[9] : null;
    const nextAttempt = row[10] instanceof Date ? row[10] : null;

    const staleProcessing =
      status === 'Processing' &&
      lastUpdate &&
      (now.getTime() - lastUpdate.getTime()) >= GK_PROCESSING_STALE_MINUTES * 60000;

    const eligible =
      status === 'Pending' ||
      status === 'Waiting for quota' ||
      status === 'Retry' ||
      staleProcessing;

    if (!eligible) continue;
    if (nextAttempt && nextAttempt.getTime() > now.getTime()) continue;

    const sheetRow = GK_QUEUE_START_ROW + i;
    const attempts = Number(row[7] || 0) + 1;

    queue.getRange(sheetRow, 7).setValue('Processing');
    queue.getRange(sheetRow, 8).setValue(attempts);
    queue.getRange(sheetRow, 10).setValue(now);
    queue.getRange(sheetRow, 12).setValue('Extraction in progress');

    return {
      row: sheetRow,
      queue_id: String(row[0] || ''),
      drive_file_id: String(row[1] || ''),
      file_name: String(row[2] || ''),
      mime_type: String(row[3] || 'application/octet-stream'),
      drive_url: String(row[4] || ''),
      attempts: attempts
    };
  }

  return null;
}


function gkQueueUpdate_(ss, job, status, message, nextAttempt, records) {
  const queue = gkEnsureQueue_(ss);
  queue.getRange(job.row, 7).setValue(status);
  queue.getRange(job.row, 10).setValue(new Date());
  queue.getRange(job.row, 11).setValue(nextAttempt || '');
  queue.getRange(job.row, 12).setValue(String(message || ''));
  if (records != null) {
    queue.getRange(job.row, 13).setValue(Number(records || 0));
  }
}


function processGKQueueTrigger() {
  const props = PropertiesService.getScriptProperties();
  const spreadsheetId = props.getProperty('SPREADSHEET_ID');
  if (!spreadsheetId) return;

  const quotaBlockedUntil = Number(props.getProperty('GK_QUOTA_BLOCK_UNTIL') || 0);
  if (quotaBlockedUntil && Date.now() < quotaBlockedUntil) {
    return;
  }

  const ss = SpreadsheetApp.openById(spreadsheetId);
  const started = Date.now();

  while ((Date.now() - started) / 1000 < GK_WORKER_MAX_SECONDS) {
    let job = null;
    const pickLock = LockService.getScriptLock();
    try {
      pickLock.waitLock(30000);
      job = gkQueueSelectJob_(ss);
    } finally {
      try { pickLock.releaseLock(); } catch (_err) {}
    }

    if (!job) return;

    try {
      const result = gkProcessDriveJob_(ss, job, props);

      const writeLock = LockService.getScriptLock();
      try {
        writeLock.waitLock(30000);
        const completedMessage = [
          result.model_used ? 'LLM: ' + gkUiText_(result.model_used).replace(/^LLM-/i, '') : '',
          result.api_index ? 'API key ' + result.api_index : ''
        ].filter(Boolean).join(' | ');
        gkQueueUpdate_(ss, job, 'Completed', completedMessage, '', result.records);
        props.deleteProperty('GK_QUOTA_BLOCK_UNTIL');
      } finally {
        try { writeLock.releaseLock(); } catch (_err) {}
      }

    } catch (err) {
      const message = String(err && err.message ? err.message : err);
      const writeLock = LockService.getScriptLock();
      try {
        writeLock.waitLock(30000);

        if (message.indexOf('__GK_QUOTA__') === 0) {
          const next = new Date(Date.now() + GK_QUOTA_RETRY_MINUTES * 60000);
          gkQueueUpdate_(ss, job, 'Waiting for quota', message.replace('__GK_QUOTA__', ''), next, null);
          props.setProperty('GK_QUOTA_BLOCK_UNTIL', String(next.getTime()));
          return;
        }

        if (message.indexOf('__GK_RETRY__') === 0 && job.attempts < 6) {
          const next = new Date(Date.now() + GK_RETRY_MINUTES * 60000);
          gkQueueUpdate_(ss, job, 'Retry', message.replace('__GK_RETRY__', ''), next, null);
        } else {
          gkQueueUpdate_(ss, job, 'Error', message, '', null);
        }
      } finally {
        try { writeLock.releaseLock(); } catch (_err) {}
      }
    }
  }
}


/**
 * Re-queues selected jobs immediately without re-uploading Drive files.
 * Used by the Streamlit "Retry waiting files now" button.
 */
function gkRetryQueueNow_(ss, body, props) {
  const wanted = {};
  (Array.isArray(body.queue_ids) ? body.queue_ids : []).forEach(function(id) {
    wanted[String(id)] = true;
  });

  const queue = gkEnsureQueue_(ss);
  const lastRow = queue.getLastRow();
  let changed = 0;

  if (lastRow >= GK_QUEUE_START_ROW) {
    const rows = queue
      .getRange(
        GK_QUEUE_START_ROW,
        1,
        lastRow - GK_QUEUE_START_ROW + 1,
        GK_QUEUE_HEADERS.length
      )
      .getValues();

    rows.forEach(function(row, i) {
      const queueId = String(row[0] || '');
      if (Object.keys(wanted).length && !wanted[queueId]) return;

      const status = String(row[6] || '').trim();
      if (status === 'Completed' || status === 'Processing') return;

      const sheetRow = GK_QUEUE_START_ROW + i;
      queue.getRange(sheetRow, 7).setValue('Pending');
      queue.getRange(sheetRow, 8).setValue(0);
      queue.getRange(sheetRow, 10).setValue(new Date());
      queue.getRange(sheetRow, 11).setValue(new Date());
      queue.getRange(sheetRow, 12).setValue('Retry requested now');
      changed += 1;
    });
  }

  props.deleteProperty('GK_QUOTA_BLOCK_UNTIL');
  gkEnsureWorkerTrigger_();
  SpreadsheetApp.flush();

  if (changed > 0) {
    gkKickWorker_();
  }

  return {
    ok: true,
    action: 'queue_retry_now',
    changed: changed,
    worker_trigger_active: gkWorkerTriggerActive_()
  };
}


function retryGKErrors() {
  const props = PropertiesService.getScriptProperties();
  const spreadsheetId = props.getProperty('SPREADSHEET_ID');
  if (!spreadsheetId) {
    throw new Error('Receiver is not initialized.');
  }

  const ss = SpreadsheetApp.openById(spreadsheetId);
  const queue = gkEnsureQueue_(ss);
  const lastRow = queue.getLastRow();

  if (lastRow < GK_QUEUE_START_ROW) {
    return;
  }

  const count = lastRow - GK_QUEUE_START_ROW + 1;
  const range = queue.getRange(
    GK_QUEUE_START_ROW,
    1,
    count,
    GK_QUEUE_HEADERS.length
  );
  const values = range.getValues();
  let changed = 0;

  values.forEach(function(row, i) {
    const status = String(row[6] || '').trim();
    if (status !== 'Error') return;

    const sheetRow = GK_QUEUE_START_ROW + i;
    queue.getRange(sheetRow, 7).setValue('Pending');
    queue.getRange(sheetRow, 8).setValue(0);
    queue.getRange(sheetRow, 10).setValue(new Date());
    queue.getRange(sheetRow, 11).setValue('');
    queue.getRange(sheetRow, 12).setValue('');
    changed += 1;
  });

  props.deleteProperty('GK_QUOTA_BLOCK_UNTIL');
  SpreadsheetApp.flush();

  console.log('Reset ' + changed + ' GK queue error job(s) to Pending.');

  if (changed > 0) {
    gkKickWorker_();
  }
}


function gkProcessDriveJob_(ss, job, props) {
  const file = DriveApp.getFileById(job.drive_file_id);
  const blob = file.getBlob();
  const mimeType = job.mime_type || blob.getContentType() || 'application/octet-stream';
  const data64 = Utilities.base64Encode(blob.getBytes());

  const modelResult = gkGenerateStructured_(data64, mimeType, props);
  const rawRecords = Array.isArray(modelResult.records) ? modelResult.records : [];
  const rowPackets = gkRecordsToPackets_(rawRecords);

  const data = gkEnsureLayout_(ss);
  const status = ensureSheet_(ss, GK_STATUS_TAB);
  const result = gkUpsertRows_(data, rowPackets, job.file_name, job.drive_url);
  gkSortData_(data);
  gkApplyDateRangeFilter_(data);

  startNewStatusBlock_(status);
  writeStatusBlock_(
    status,
    job.file_name,
    'Done',
    '1/1',
    rowPackets.length,
    'Background queue: inserted ' + result.inserted +
      ', updated ' + result.updated +
      (result.cleared ? ', crossed-out values cleared ' + result.cleared : '') +
      (result.date_conflicts ? ', protected date conflicts ' + result.date_conflicts : '') +
      (modelResult.model_used ? ', LLM ' + gkUiText_(modelResult.model_used).replace(/^LLM-/i, '') : '') +
      (modelResult.api_index ? ', API key ' + modelResult.api_index : '')
  );

  SpreadsheetApp.flush();
  return {
    records: rowPackets.length,
    model_used: String(modelResult.model_used || ''),
    api_index: Number(modelResult.api_index || 0)
  };
}


function gkGenerateStructured_(data64, mimeType, props) {
  let configuredModel = String(
    props.getProperty('GEMINI_MODEL') || 'gemini-3.8-flash'
  ).trim();

  // The old 3.5 Flash name now routes to 3.6 Flash. Use the current endpoint
  // directly so a stale property does not keep us on the old alias.
  if (configuredModel === 'gemini-3.5-flash') {
    configuredModel = 'gemini-3.6-flash';
  }

  const automaticFallback =
    configuredModel === 'gemini-3.8-flash'
      ? 'gemini-3.6-flash'
      : 'gemini-3.8-flash';

  const models = [
    configuredModel,
    String(props.getProperty('GEMINI_FALLBACK_MODEL') || '').trim(),
    automaticFallback,
    'gemini-3.8-flash',
    'gemini-3.6-flash'
  ].filter(function(v, i, arr) {
    return v && arr.indexOf(v) === i;
  });

  const keys = [
    props.getProperty('GEMINI_API_KEY_1') || '',
    props.getProperty('GEMINI_API_KEY_2') || ''
  ].map(function(v) {
    return String(v || '').trim();
  }).filter(function(v, i, arr) {
    return v && arr.indexOf(v) === i;
  });

  if (!keys.length) {
    throw new Error('No API keys configured. Run configureGKBackground().');
  }

  const schemaProperties = {
    source_row: {type: 'integer'},
    sno: {type: 'string'},
    date: {type: 'string'},
    gk_no: {type: 'string'},
    uhid_no: {type: 'string'},
    patient_name: {type: 'string'},
    mobile_no: {type: 'string'},
    age: {type: 'string'},
    sex: {type: 'string'},
    diagnosis: {type: 'string'},
    mask: {type: 'string'},
    dose: {type: 'string'},
    perc: {type: 'string'},
    volume: {type: 'string'},
    gura: {type: 'string'},
    cons: {type: 'string'},
    phy: {type: 'string'},
    ro: {type: 'string'},
    rtt: {type: 'string'},
    crossed_out_fields: {type: 'array', items: {type: 'string'}}
  };

  const requiredFields = Object.keys(schemaProperties);
  const mediaType =
    String(mimeType || '').toLowerCase() === 'application/pdf'
      ? 'document'
      : 'image';

  const basePayload = {
    input: [
      {
        type: mediaType,
        data: data64,
        mime_type: mimeType
      },
      {
        type: 'text',
        text: GK_BACKGROUND_PROMPT
      }
    ],
    response_format: {
      type: 'text',
      mime_type: 'application/json',
      schema: {
        type: 'object',
        properties: {
          records: {
            type: 'array',
            items: {
              type: 'object',
              properties: schemaProperties,
              required: requiredFields
            }
          }
        },
        required: ['records']
      }
    },
    generation_config: {
      thinking_level: 'low'
    }
  };

  const requestStarted = Date.now();
  const requestBudgetMs = 180000; // leave headroom under Apps Script runtime
  let sawQuota = false;
  let sawRetryable = false;
  let lastError = '';
  let lastQuotaError = '';
  let lastRetryableError = '';
  let nonQuotaNonRetryable = '';

  for (let m = 0; m < models.length; m++) {
    const model = models[m];

    for (let i = 0; i < keys.length; i++) {
      if (Date.now() - requestStarted > requestBudgetMs) {
        break;
      }

      const payload = Object.assign({model: model}, basePayload);
      const response = UrlFetchApp.fetch(
        'https://generativelanguage.googleapis.com/v1beta/interactions',
        {
          method: 'post',
          contentType: 'application/json',
          headers: {
            'x-goog-api-key': keys[i],
            'x-goog-api-client': 'gk-register-apps-script/19'
          },
          payload: JSON.stringify(payload),
          muteHttpExceptions: true
        }
      );

      const code = response.getResponseCode();
      const body = response.getContentText();

      if (code >= 200 && code < 300) {
        const parsed = gkParseInteractionResponse_(body);
        parsed.model_used = model;
        parsed.api_index = i + 1;
        return parsed;
      }

      lastError =
        'Model ' + model +
        ' | API ' + (i + 1) +
        ' | HTTP ' + code + ': ' +
        body.slice(0, 1200);

      if (code === 429 || /quota|rate.?limit|resource.?exhausted/i.test(body)) {
        sawQuota = true;
        lastQuotaError = lastError;
        continue;
      }

      if (code === 500 || code === 502 || code === 503 || code === 504) {
        sawRetryable = true;
        lastRetryableError = lastError;
        // High demand on one key/model should not immediately park the file.
        // Briefly wait, then automatically try the next key or fallback model.
        if (Date.now() - requestStarted < requestBudgetMs - 3000) {
          Utilities.sleep(2000);
        }
        continue;
      }

      // Invalid/restricted key: try the next key and model before giving up.
      if (code === 400 || code === 401 || code === 403) {
        nonQuotaNonRetryable = lastError;
        continue;
      }

      nonQuotaNonRetryable = lastError;
    }

    if (Date.now() - requestStarted > requestBudgetMs) {
      break;
    }
  }

  // If any backend/service call returned a transient 5xx, prefer Retry over a
  // quota label because another configured model was reachable but overloaded.
  if (sawRetryable) {
    throw new Error(
      '__GK_RETRY__' +
      (lastRetryableError || 'Temporary model service error') +
      ' | Tried models: ' + models.join(', ') +
      ' | File stays queued and will retry automatically.'
    );
  }

  if (sawQuota) {
    throw new Error(
      '__GK_QUOTA__' +
      (lastQuotaError || 'HTTP 429: API quota/rate limit reached') +
      ' | Tried models: ' + models.join(', ') +
      ' | File stays queued and will retry automatically.'
    );
  }

  throw new Error(nonQuotaNonRetryable || lastError || 'Model request failed.');
}


function gkParseInteractionResponse_(bodyText) {
  const obj = JSON.parse(bodyText);

  if (obj && Array.isArray(obj.records)) {
    return {records: obj.records};
  }

  const steps = Array.isArray(obj.steps) ? obj.steps : [];
  const texts = [];

  steps.forEach(function(step) {
    if (!step || step.type !== 'model_output') return;
    const content = Array.isArray(step.content) ? step.content : [];
    content.forEach(function(block) {
      if (block && block.type === 'text' && block.text != null) {
        texts.push(String(block.text));
      }
    });
  });

  let outputText = texts.join('').trim();

  if (!outputText && obj.output_text != null) {
    outputText = String(obj.output_text).trim();
  }

  if (!outputText) {
    throw new Error('Model returned an empty structured response.');
  }

  outputText = outputText
    .replace(/^```json\s*/i, '')
    .replace(/^```\s*/i, '')
    .replace(/```$/i, '')
    .trim();

  const parsed = JSON.parse(outputText);

  if (Array.isArray(parsed)) {
    return {records: parsed};
  }

  if (!parsed || !Array.isArray(parsed.records)) {
    throw new Error('Structured response did not contain a records array.');
  }

  return parsed;
}


function gkParseModelResponse_(bodyText) {
  const obj = JSON.parse(bodyText);
  const candidates = Array.isArray(obj.candidates) ? obj.candidates : [];
  if (!candidates.length) throw new Error('Model returned no candidates.');

  const parts = candidates[0] && candidates[0].content && Array.isArray(candidates[0].content.parts)
    ? candidates[0].content.parts
    : [];
  const text = parts
    .map(function(part) { return String(part.text || ''); })
    .join('')
    .trim();

  if (!text) throw new Error('Model returned an empty structured response.');

  let cleaned = text;
  cleaned = cleaned.replace(/^```json\s*/i, '').replace(/^```\s*/i, '').replace(/```$/i, '').trim();
  const parsed = JSON.parse(cleaned);
  if (Array.isArray(parsed)) return {records: parsed};
  if (!parsed || !Array.isArray(parsed.records)) {
    throw new Error('Structured response did not contain a records array.');
  }
  return parsed;
}


function gkRecordsToPackets_(records) {
  const output = [];
  const byYearSno = {};

  (records || []).forEach(function(raw, index) {
    if (!raw || typeof raw !== 'object') return;

    const sno = gkCleanText_(raw.sno).replace(/^[#\s]+/, '').replace(/[.\s]+$/, '');
    const date = gkNormalizeDateText_(raw.date);
    const gkNo = gkNormalizeRepeat_(raw.gk_no);
    const nm = gkSplitNameMobile_(
      [gkCleanText_(raw.patient_name), gkCleanText_(raw.mobile_no)].filter(Boolean).join(' ')
    );
    const asx = gkSplitAgeSex_(
      [gkCleanText_(raw.age), gkCleanText_(raw.sex)].filter(Boolean).join('/')
    );
    const diag = gkDiagnosisTechnique_(raw.diagnosis);
    const volume = gkVolumeCc_(raw.volume);

    const values = [
      sno,
      date,
      gkNo,
      gkCleanText_(raw.uhid_no),
      nm.name,
      nm.mobile,
      asx.age,
      asx.sex,
      diag.diagnosis,
      gkCleanText_(raw.mask),
      gkCleanText_(raw.dose),
      gkCleanText_(raw.perc),
      volume,
      gkCleanText_(raw.gura),
      gkCleanText_(raw.cons),
      gkCleanText_(raw.phy),
      gkCleanText_(raw.ro),
      gkCleanText_(raw.rtt),
      diag.technique
    ];

    const clear = gkCanonicalClearFields_(raw.crossed_out_fields, values);
    const packet = {values: values, clear_fields: clear};

    const key = gkYearSnoKey_(sno, date);
    if (key && byYearSno[key] != null) {
      // Merge duplicate interpretations only when BOTH year and S.No match.
      // S.No can restart from 1 in a new calendar year.
      const prior = output[byYearSno[key]];
      for (let c = 0; c < values.length; c++) {
        if (!prior.values[c] && values[c]) prior.values[c] = values[c];
      }
      prior.clear_fields = prior.clear_fields
        .concat(clear)
        .filter(function(v, i, arr) { return arr.indexOf(v) === i; })
        .filter(function(field) {
          const idx = GK_FIELD_KEYS.indexOf(field);
          return idx < 0 || !prior.values[idx];
        });
    } else {
      if (key) byYearSno[key] = output.length;
      output.push(packet);
    }
  });

  return output;
}


function gkCleanText_(value) {
  return String(value == null ? '' : value).replace(/\s+/g, ' ').trim();
}


function gkNormalizeRepeat_(value) {
  let s = gkCleanText_(value);
  if (!s) return '';
  if (/^repeat\b/i.test(s)) {
    return 'Repeat ' + s.replace(/^repeat\s*[:\-]?\s*/i, '').trim();
  }

  const hasR = /\(\s*r\s*\)/i.test(s) || /(^|\s)r(\s|$)/i.test(s) || /^r\s*[-:]?\s*\d/i.test(s);
  if (!hasR) return s;

  s = s
    .replace(/\(\s*r\s*\)/ig, ' ')
    .replace(/(^|\s)r(?=\s|$)/ig, ' ')
    .replace(/\s+/g, ' ')
    .replace(/^\s*[:\-]+|[:\-]+\s*$/g, '')
    .trim();
  return ('Repeat ' + s).trim();
}


function gkDiagnosisTechnique_(value) {
  let s = gkCleanText_(value);
  let technique = '';
  if (/\(\s*m\s*\)/i.test(s)) {
    technique = 'Masked';
    s = s.replace(/\(\s*m\s*\)/ig, ' ');
  }
  if (/\(\s*v\s*\)/i.test(s)) {
    technique = technique ? technique + ' / Vantage' : 'Vantage';
    s = s.replace(/\(\s*v\s*\)/ig, ' ');
  }
  s = s.replace(/\s+/g, ' ').replace(/^[\s\-;,|]+|[\s\-;,|]+$/g, '');
  return {diagnosis: s, technique: technique};
}


function gkCanonicalClearFields_(rawFields, values) {
  const aliases = {
    'serial_number': 'sno', 's.no': 'sno', 'sno': 'sno',
    'date': 'date',
    'gk_no': 'gk_no', 'gamma_knife_number': 'gk_no',
    'uhid': 'uhid_no', 'uhid_no': 'uhid_no',
    'patient_name': 'patient_name', 'name': 'patient_name',
    'mobile': 'mobile_no', 'mobile_no': 'mobile_no',
    'age': 'age', 'sex': 'sex', 'diagnosis': 'diagnosis',
    'mask': 'mask', 'dose': 'dose', 'perc': 'perc', 'percentage': 'perc',
    'volume': 'volume', 'gura': 'gura',
    'cons': 'cons', 'consultant': 'cons',
    'phy': 'phy', 'physician': 'phy',
    'ro': 'ro', 'r.o': 'ro', 'residence': 'ro',
    'rtt': 'rtt', 'technique': 'technique'
  };

  const result = [];
  (Array.isArray(rawFields) ? rawFields : []).forEach(function(raw) {
    const key = gkCleanText_(raw).toLowerCase().replace(/[^a-z0-9.]+/g, '_').replace(/^_+|_+$/g, '');
    const field = aliases[key];
    if (!field || result.indexOf(field) >= 0) return;
    const idx = GK_FIELD_KEYS.indexOf(field);
    // If an uncrossed replacement is present, do not intentionally clear it.
    if (idx >= 0 && !gkCleanText_(values[idx])) result.push(field);
  });
  return result;
}


function processGKQueueNow() {
  processGKQueueTrigger();
}


function gkApplyDateRangeFilter_(sheet) {
  const maxRows = sheet.getMaxRows();
  if (maxRows >= GK_DATA_START_ROW) {
    sheet.showRows(GK_DATA_START_ROW, maxRows - GK_DATA_START_ROW + 1);
  }

  const fromDate = parseComparableDate_(sheet.getRange('B1').getValue());
  const toDate = parseComparableDate_(sheet.getRange('B2').getValue());
  if (!fromDate || !toDate) {
    return;
  }

  let start = fromDate;
  let end = toDate;
  if (start.getTime() > end.getTime()) {
    const temp = start;
    start = end;
    end = temp;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < GK_DATA_START_ROW) {
    return;
  }

  // Date is column B for the GK dataset.
  const values = sheet
    .getRange(GK_DATA_START_ROW, 2, lastRow - GK_DATA_START_ROW + 1, 1)
    .getValues();

  const rowsToHide = [];
  values.forEach(function(row, idx) {
    const rowNumber = GK_DATA_START_ROW + idx;
    const rowDate = parseComparableDate_(row[0]);
    if (!rowDate || rowDate.getTime() < start.getTime() || rowDate.getTime() > end.getTime()) {
      rowsToHide.push(rowNumber);
    }
  });

  // Hide contiguous groups efficiently.
  if (!rowsToHide.length) {
    return;
  }
  let groupStart = rowsToHide[0];
  let previous = rowsToHide[0];
  for (let i = 1; i <= rowsToHide.length; i++) {
    const current = i < rowsToHide.length ? rowsToHide[i] : null;
    if (current === previous + 1) {
      previous = current;
      continue;
    }
    sheet.hideRows(groupStart, previous - groupStart + 1);
    if (current !== null) {
      groupStart = current;
      previous = current;
    }
  }
}


function ensureStatusSheet_(ss) {
  return (
    ss.getSheetByName(STATUS_TAB) ||
    ss.insertSheet(STATUS_TAB)
  );
}


function ensureSheet_(ss, title) {
  return (
    ss.getSheetByName(title) ||
    ss.insertSheet(title)
  );
}


function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
