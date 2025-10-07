/**
 * Apps Script utility for syncing Google Sheets data into Google Slides tables.
 *
 * This script is designed for media reporting workflows where tables in a slide deck
 * must stay aligned with spreadsheet data.  The script can pull its configuration
 * directly from a "Slide Sync Config" sheet so non-technical users can manage which
 * tables are refreshed without editing the code.
 */

const SLIDE_SYNC_CONFIG = {
  SHEET_NAME: 'Slide Sync Config',
  HEADER_ROW_INDEX: 0,
  ENABLED_COLUMN: 'enabled',
  REQUIRED_HEADERS: [
    'spreadsheet id',
    'sheet name',
    'slide id',
    'page index',
    'match text'
  ],
  OPTIONAL_HEADERS: {
    RANGE: 'range',
    SKIP_HEADER_ROWS: 'skip header rows',
    MAX_ROWS: 'max rows',
    REMOVE_BLANK_ROWS: 'remove blank rows',
    SCHEDULE_ENABLED: 'schedule enabled',
    CRON_EXPRESSION: 'cron expression',
    TEMPLATE: 'template',
    VALIDATION_RULES: 'validation rules'
  }
};

const SLIDE_SYNC_TEMPLATES = {
  'weekly-report': {
    label: 'Weekly Report',
    description: 'Optimised for weekly pacing dashboards.',
    defaults: {
      skipHeaderRows: 1,
      removeBlankRows: true,
      scheduleEnabled: false,
      validationRules: {
        hasHeaderRow: true,
        requiredColumns: ['campaign', 'channel', 'spend'],
        numericColumns: ['impressions', 'spend', 'clicks']
      },
      cronExpression: '0 9 * * 1'
    }
  },
  'monthly-dashboard': {
    label: 'Monthly Dashboard',
    description: 'Aggregated KPI dashboards refreshed once a month.',
    defaults: {
      skipHeaderRows: 1,
      removeBlankRows: true,
      scheduleEnabled: false,
      validationRules: {
        hasHeaderRow: true,
        requiredColumns: ['month'],
        numericColumns: ['spend', 'impressions', 'conversions']
      },
      cronExpression: '0 8 1 * *'
    }
  },
  'daily-metrics': {
    label: 'Daily Metrics',
    description: 'Daily pacing templates with stricter validation.',
    defaults: {
      skipHeaderRows: 1,
      removeBlankRows: true,
      scheduleEnabled: false,
      validationRules: {
        hasHeaderRow: true,
        requiredColumns: ['date'],
        numericColumns: ['impressions', 'clicks', 'spend'],
        maxRows: 31
      },
      cronExpression: '0 7 * * *'
    }
  }
};

const SLIDE_SYNC_SCHEDULER = {
  HANDLER_FUNCTION: 'pollScheduledSlides',
  DEFAULT_POLL_MINUTES: 5,
  LAST_RUN_PROPERTY_PREFIX: 'slideSync.lastRun.',
  MAX_ERRORS_TO_LOG: 5
};

const SLIDE_SYNC_VALIDATION = {
  MAX_ERRORS: 10
};

/**
 * ====================================================
 * Main Controller – Run THIS function only
 * ====================================================
 */
function refreshAllSlides() {
  Logger.log('--- Starting data refresh for all slides ---');

  const updates = getConfiguredUpdates();
  if (!updates.length) {
    Logger.log('No updates were found in the configuration. Aborting.');
    Logger.log('--- Data refresh complete ---');
    return;
  }

  const issues = [];

  updates.forEach((update) => {
    const issue = runParse(update);
    if (issue) issues.push(issue);
  });

  if (issues.length > 0) {
    Logger.log('TL;DR - Issues encountered during the process:');
    issues.forEach((issue, index) => Logger.log(`${index + 1}. ${issue}`));
  } else {
    Logger.log('TL;DR - No issues encountered. All slides updated successfully.');
  }

  Logger.log('--- Data refresh complete ---');
}

/**
 * ====================================================
 * Configuration Loader
 * ====================================================
 */
function getConfiguredUpdates() {
  let spreadsheet;
  try {
    spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {
    Logger.log('No active spreadsheet detected; defaulting to hard-coded updates.');
  }

  if (!spreadsheet) {
    return getFallbackUpdates();
  }

  const configSheet = spreadsheet.getSheetByName(SLIDE_SYNC_CONFIG.SHEET_NAME);
  if (!configSheet) {
    Logger.log(`Configuration sheet "${SLIDE_SYNC_CONFIG.SHEET_NAME}" not found. Using fallback updates.`);
    return getFallbackUpdates();
  }

  const values = configSheet.getDataRange().getValues();
  if (values.length <= 1) {
    Logger.log('Configuration sheet does not contain any rows beyond the header.');
    return [];
  }

  const headers = values[SLIDE_SYNC_CONFIG.HEADER_ROW_INDEX].map((header) =>
    String(header || '').trim().toLowerCase()
  );

  const missingHeaders = SLIDE_SYNC_CONFIG.REQUIRED_HEADERS.filter(
    (header) => !headers.includes(header)
  );

  if (missingHeaders.length) {
    Logger.log(`Configuration sheet is missing required columns: ${missingHeaders.join(', ')}`);
    return [];
  }

  const headerIndex = (name) => headers.indexOf(name);

  const updates = [];
  for (let i = SLIDE_SYNC_CONFIG.HEADER_ROW_INDEX + 1; i < values.length; i++) {
    const row = values[i];
    if (row.every(isEmptyValue)) continue;

    const enabledIndex = headerIndex(SLIDE_SYNC_CONFIG.ENABLED_COLUMN);
    const isEnabled = enabledIndex === -1 ? true : parseBoolean(row[enabledIndex]);
    if (!isEnabled) continue;

    const update = {
      spreadsheetId: String(row[headerIndex('spreadsheet id')]).trim(),
      sheetName: String(row[headerIndex('sheet name')]).trim(),
      slideId: String(row[headerIndex('slide id')]).trim(),
      pageId: parseInt(row[headerIndex('page index')], 10),
      matchText: String(row[headerIndex('match text')]).trim(),
      configRowNumber: i + 1
    };

    if (!update.spreadsheetId || !update.sheetName || !update.slideId || isNaN(update.pageId)) {
      Logger.log(`Skipping configuration row ${i + 1} because it is missing required values.`);
      continue;
    }

    const rangeIndex = headerIndex(SLIDE_SYNC_CONFIG.OPTIONAL_HEADERS.RANGE);
    if (rangeIndex > -1 && !isEmptyValue(row[rangeIndex])) {
      update.rangeA1 = String(row[rangeIndex]).trim();
    }

    const skipHeaderIndex = headerIndex(SLIDE_SYNC_CONFIG.OPTIONAL_HEADERS.SKIP_HEADER_ROWS);
    if (skipHeaderIndex > -1 && !isEmptyValue(row[skipHeaderIndex])) {
      update.skipHeaderRows = Math.max(0, parseInt(row[skipHeaderIndex], 10));
    }

    const maxRowsIndex = headerIndex(SLIDE_SYNC_CONFIG.OPTIONAL_HEADERS.MAX_ROWS);
    if (maxRowsIndex > -1 && !isEmptyValue(row[maxRowsIndex])) {
      update.maxRows = Math.max(0, parseInt(row[maxRowsIndex], 10));
    }

    const removeBlankIndex = headerIndex(SLIDE_SYNC_CONFIG.OPTIONAL_HEADERS.REMOVE_BLANK_ROWS);
    if (removeBlankIndex > -1 && !isEmptyValue(row[removeBlankIndex])) {
      update.removeBlankRows = parseBoolean(row[removeBlankIndex]);
    }

    const scheduleEnabledIndex = headerIndex(SLIDE_SYNC_CONFIG.OPTIONAL_HEADERS.SCHEDULE_ENABLED);
    if (scheduleEnabledIndex > -1 && !isEmptyValue(row[scheduleEnabledIndex])) {
      update.scheduleEnabled = parseBoolean(row[scheduleEnabledIndex]);
    }

    const cronExpressionIndex = headerIndex(SLIDE_SYNC_CONFIG.OPTIONAL_HEADERS.CRON_EXPRESSION);
    if (cronExpressionIndex > -1 && !isEmptyValue(row[cronExpressionIndex])) {
      update.cronExpression = String(row[cronExpressionIndex]).trim();
    }

    const templateIndex = headerIndex(SLIDE_SYNC_CONFIG.OPTIONAL_HEADERS.TEMPLATE);
    if (templateIndex > -1 && !isEmptyValue(row[templateIndex])) {
      update.templateKey = String(row[templateIndex]).trim();
    }

    const validationIndex = headerIndex(SLIDE_SYNC_CONFIG.OPTIONAL_HEADERS.VALIDATION_RULES);
    let validationOverrides = null;
    if (validationIndex > -1 && !isEmptyValue(row[validationIndex])) {
      validationOverrides = parseValidationRules(row[validationIndex], i + 1);
    }

    const templateDefaults = applyTemplateDefaults(update);
    update.validationRules = mergeValidationRules(templateDefaults && templateDefaults.validationRules, validationOverrides);

    if (update.scheduleEnabled === undefined) {
      update.scheduleEnabled = update.cronExpression ? true : false;
    }

    if (!update.cronExpression) {
      update.scheduleEnabled = false;
    }

    update.identifier = buildUpdateIdentifier(update);

    updates.push(update);
  }

  return updates.length ? updates : getFallbackUpdates();
}

function getFallbackUpdates() {
  return [
    {
      spreadsheetId: '1qHZ-BgeaspYIhC4hBkyhOLLQzza20VSwiz8rYynZddg',
      sheetName: 'Totals',
      slideId: '1YMVjyrfgj7jGhKjtjGKr1HVsDjBIsERLLyPT6ytqzsc',
      pageId: 4,
      matchText: 'TOTAL IMPRESSIONS',
    },
    {
      spreadsheetId: '1qHZ-BgeaspYIhC4hBkyhOLLQzza20VSwiz8rYynZddg',
      sheetName: 'Combined Data',
      slideId: '1YMVjyrfgj7jGhKjtjGKr1HVsDjBIsERLLyPT6ytqzsc',
      pageId: 4,
      matchText: 'Source',
    },
    {
      spreadsheetId: '1qHZ-BgeaspYIhC4hBkyhOLLQzza20VSwiz8rYynZddg',
      sheetName: 'Total Cost',
      slideId: '1YMVjyrfgj7jGhKjtjGKr1HVsDjBIsERLLyPT6ytqzsc',
      pageId: 4,
      matchText: 'TOTAL SPEND',
    },
  ].map((update) => ({
    identifier: buildUpdateIdentifier(update),
    removeBlankRows: true,
    scheduleEnabled: false,
    ...update,
  }));
}

/**
 * ====================================================
 * Core Parser Function (now self-healing if called alone)
 * ====================================================
 */
function runParse(update) {
  try {
    const context = buildUpdateContext(update, { logSteps: true });

    if (context.validation && !context.validation.valid) {
      Logger.log(`${context.update.sheetName} - Validation failed: ${context.validation.errors.join('; ')}`);
      return `Validation failed for sheet "${context.update.sheetName}": ${context.validation.errors.join('; ')}`;
    }

    syncTableSize(
      context.table,
      context.normalizedData.length,
      context.normalizedData[0] ? context.normalizedData[0].length : 0
    );
    writeTableValues(context.table, context.normalizedData);

    Logger.log(`${context.update.sheetName} - Table updated successfully on slide.`);
    return null;
  } catch (e) {
    const issue = `Error in sheet "${update?.sheetName || 'Unknown'}": ${e.message}`;
    Logger.log(issue);
    return issue;
  }
}

function previewAllSlides() {
  const updates = getConfiguredUpdates();
  if (!updates.length) {
    Logger.log('No updates available to preview.');
    return [];
  }

  const previews = [];
  updates.forEach((update) => {
    previews.push(previewUpdate(update));
  });
  return previews;
}

function previewUpdate(update) {
  try {
    const context = buildUpdateContext(update, { allowInvalid: true, logSteps: false });
    const summary = summarizeDiff(context.diff);

    Logger.log(`Preview for ${context.update.sheetName}: ${summary}`);
    if (context.validation && !context.validation.valid) {
      Logger.log(`Validation issues: ${context.validation.errors.join('; ')}`);
    }

    return {
      update: context.update,
      summary,
      diff: context.diff,
      currentValues: context.currentValues,
      nextValues: context.normalizedData,
      validation: context.validation,
      truncatedColumns: context.truncatedColumns
    };
  } catch (e) {
    const issue = `Preview failed for sheet "${update?.sheetName || 'Unknown'}": ${e.message}`;
    Logger.log(issue);
    return {
      update,
      error: issue
    };
  }
}

function buildUpdateContext(update, options) {
  const opts = options || {};
  const logEnabled = opts.logSteps !== false;
  const log = (message) => {
    if (logEnabled) Logger.log(message);
  };

  if (!update || typeof update !== 'object') {
    throw new Error('runParse() called without parameters.');
  }

  const context = { update: Object.assign({}, update) };
  const {
    spreadsheetId,
    sheetName,
    slideId,
    pageId,
    matchText,
    rangeA1,
    skipHeaderRows = 0,
    maxRows,
    removeBlankRows = true
  } = update;

  log(`--- Starting update for sheet "${sheetName}" on slide ID "${slideId}" ---`);

  const sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);
  if (!sheet) throw new Error(`${sheetName} - Sheet not found.`);

  let dataRange;
  if (rangeA1) {
    log(`${sheetName} - Pulling data from range ${rangeA1}`);
    dataRange = sheet.getRange(rangeA1);
  } else {
    dataRange = sheet.getDataRange();
  }

  const originalData = clone2DArray(dataRange.getValues());
  log(`${sheetName} - Data retrieved: ${originalData.length} rows`);

  let data = originalData.slice();

  if (skipHeaderRows > 0) {
    data = data.slice(skipHeaderRows);
    log(`${sheetName} - Skipped ${skipHeaderRows} header rows. Remaining rows: ${data.length}`);
  }

  if (maxRows && maxRows > 0) {
    data = data.slice(0, maxRows);
    log(`${sheetName} - Truncated to ${maxRows} rows as requested.`);
  }

  const shouldRemoveBlankRows = removeBlankRows === undefined ? true : removeBlankRows;
  if (shouldRemoveBlankRows) {
    const originalLength = data.length;
    data = data.filter((row) => !row.every(isEmptyValue));
    if (!data.length && originalLength) {
      log(`${sheetName} - All rows were blank; keeping a single empty row to preserve table structure.`);
      data = [[]];
    }
  }

  if (!data.length) {
    data = [[]];
  }

  const slides = SlidesApp.openById(slideId).getSlides();
  if (!slides.length) throw new Error(`${sheetName} - No slides found.`);
  const slide = slides[pageId];
  if (!slide) throw new Error(`${sheetName} - Slide index ${pageId} not found in deck.`);

  const tables = slide.getTables();
  if (!tables.length) throw new Error(`${sheetName} - No tables found on slide.`);

  let table = findMatchingTable(tables, matchText);
  if (!table) {
    log(`${sheetName} - No matching table found with first cell "${matchText}". Using first table as fallback.`);
    table = tables[0];
  }

  const rows = data.length;
  const columns = largestArray(data).length;
  log(`${sheetName} - Found table with ${table.getNumRows()} rows and ${table.getNumColumns()} columns`);
  log(`${sheetName} - Sheet data has ${rows} rows and ${columns} columns`);

  const normalizedData = normalizeDataToTable(data, table.getNumColumns());
  const currentValues = readTableValues(table);
  const diff = computeTableDiff(currentValues, normalizedData);

  const headerRowIndex = Math.max(0, Math.min(skipHeaderRows > 0 ? skipHeaderRows - 1 : 0, originalData.length - 1));
  const headerRow = originalData[headerRowIndex] || [];
  const validation = validateDataSet(normalizedData, headerRow, update.validationRules);

  if (validation && !validation.valid && !opts.allowInvalid) {
    throw new Error(`Validation failed: ${validation.errors.join('; ')}`);
  }

  context.table = table;
  context.slide = slide;
  context.sheet = sheet;
  context.normalizedData = normalizedData;
  context.currentValues = currentValues;
  context.diff = diff;
  context.validation = validation;
  context.truncatedColumns = columns > table.getNumColumns();

  return context;
}

function summarizeDiff(diff) {
  if (!diff || !diff.length) return 'No changes detected.';
  const added = diff.filter((item) => item.type === 'added').length;
  const removed = diff.filter((item) => item.type === 'removed').length;
  const changed = diff.filter((item) => item.type === 'changed').length;
  const parts = [];
  if (added) parts.push(`${added} row(s) added`);
  if (removed) parts.push(`${removed} row(s) removed`);
  if (changed) parts.push(`${changed} row(s) changed`);
  return parts.join(', ');
}

function readTableValues(table) {
  const values = [];
  for (let i = 0; i < table.getNumRows(); i++) {
    const row = [];
    for (let j = 0; j < table.getNumColumns(); j++) {
      row.push(table.getCell(i, j).getText().asString());
    }
    values.push(row);
  }
  return values;
}

function computeTableDiff(currentValues, nextValues) {
  const diff = [];
  const maxRows = Math.max(currentValues.length, nextValues.length);
  for (let i = 0; i < maxRows; i++) {
    const before = currentValues[i] || [];
    const after = nextValues[i] || [];

    if (!before.length && after.length) {
      diff.push({ type: 'added', rowIndex: i, after });
      continue;
    }

    if (before.length && !after.length) {
      diff.push({ type: 'removed', rowIndex: i, before });
      continue;
    }

    const changedColumns = [];
    const maxCols = Math.max(before.length, after.length);
    for (let col = 0; col < maxCols; col++) {
      const previous = before[col] == null ? '' : String(before[col]);
      const next = after[col] == null ? '' : String(after[col]);
      if (previous !== next) {
        changedColumns.push(col);
      }
    }

    if (changedColumns.length) {
      diff.push({ type: 'changed', rowIndex: i, before, after, changedColumns });
    }
  }

  return diff;
}

/**
 * ====================================================
 * Helper: Find the correct table by first-cell text match
 * ====================================================
 */
function findMatchingTable(tables, matchText) {
  if (!matchText) return null;
  for (let i = 0; i < tables.length; i++) {
    try {
      const firstCell = tables[i].getCell(0, 0).getText().asString().trim();
      if (firstCell.toLowerCase().includes(matchText.toLowerCase())) return tables[i];
    } catch (e) {
      continue;
    }
  }
  return null;
}

/**
 * ====================================================
 * Helper: Find the largest array within an array of arrays
 * ====================================================
 */
function largestArray(arrayOfArrays) {
  if (!arrayOfArrays.length) return [];
  return arrayOfArrays.reduce((largestArray, currentArray) =>
    currentArray.length > largestArray.length ? currentArray : largestArray,
  []);
}

/**
 * Ensure that the data fits the table column count.
 */
function normalizeDataToTable(data, tableColumnCount) {
  if (!tableColumnCount) return data;
  return data.map((row) => {
    const normalizedRow = row.slice(0, tableColumnCount);
    while (normalizedRow.length < tableColumnCount) normalizedRow.push('');
    return normalizedRow;
  });
}

/**
 * Sync the table row/column count with the incoming data.
 */
function syncTableSize(table, targetRows, targetColumns) {
  if (targetColumns > table.getNumColumns()) {
    throw new Error('Incoming data contains more columns than the destination table. Expand the table first.');
  }

  while (table.getNumRows() < targetRows) table.appendRow();
  while (table.getNumRows() > targetRows) table.getRow(table.getNumRows() - 1).remove();
}

/**
 * Write normalized data to the slide table.
 */
function writeTableValues(table, data) {
  for (let i = 0; i < data.length; i++) {
    for (let j = 0; j < table.getNumColumns(); j++) {
      const value = data[i] && data[i][j] != null ? data[i][j] : '';
      table.getCell(i, j).getText().setText(String(value));
    }
  }
}

/**
 * Utility helpers
 */
function isEmptyValue(value) {
  return value === '' || value === null || value === undefined;
}

function parseBoolean(value) {
  if (typeof value === 'boolean') return value;
  const normalized = String(value || '').trim().toLowerCase();
  if (!normalized) return false;
  return ['true', 'yes', 'y', '1'].includes(normalized);
}

function clone2DArray(values) {
  if (!values || !values.map) return [];
  return values.map((row) => row.slice());
}

function cloneObject(object) {
  if (object === null || object === undefined) return null;
  try {
    return JSON.parse(JSON.stringify(object));
  } catch (e) {
    return null;
  }
}

function applyTemplateDefaults(update) {
  if (!update.templateKey) return null;
  const templateKey = String(update.templateKey || '').trim().toLowerCase();
  const template = SLIDE_SYNC_TEMPLATES[templateKey];
  if (!template) {
    Logger.log(`Template "${update.templateKey}" was not recognised. Skipping template defaults.`);
    return null;
  }

  const defaults = template.defaults || {};
  const validationDefaults = defaults.validationRules ? cloneObject(defaults.validationRules) : null;

  Object.keys(defaults).forEach((key) => {
    if (key === 'validationRules') return;
    if (update[key] === undefined || update[key] === null || update[key] === '') {
      update[key] = defaults[key];
    }
  });

  update.appliedTemplate = templateKey;
  return { validationRules: validationDefaults };
}

function mergeValidationRules(baseRules, overrides) {
  if (!baseRules && !overrides) return null;
  if (!baseRules) return cloneObject(overrides);
  if (!overrides) return cloneObject(baseRules);
  const merged = cloneObject(baseRules) || {};
  Object.keys(overrides).forEach((key) => {
    merged[key] = overrides[key];
  });
  return merged;
}

function parseValidationRules(value, rowNumber) {
  if (value === null || value === undefined) return null;
  if (typeof value === 'object') return cloneObject(value);
  const normalized = String(value).trim();
  if (!normalized) return null;
  try {
    return JSON.parse(normalized);
  } catch (e) {
    Logger.log(`Row ${rowNumber}: Unable to parse validation rules. Ensure the value is valid JSON.`);
    return null;
  }
}

function buildUpdateIdentifier(update) {
  const parts = [
    update.spreadsheetId || '',
    update.sheetName || '',
    update.slideId || '',
    update.pageId != null ? update.pageId : '',
    update.matchText || ''
  ];
  return Utilities.base64EncodeWebSafe(parts.join('::'));
}

function validateDataSet(data, headerRow, rules) {
  if (!rules) return { valid: true, errors: [], warnings: [] };

  const result = { valid: true, errors: [], warnings: [] };
  const hasHeader = rules.hasHeaderRow === undefined ? true : parseBoolean(rules.hasHeaderRow);
  const workingData = hasHeader ? data.slice(1) : data.slice();
  const normalizedHeader = (headerRow || []).map((value) => String(value || '').trim().toLowerCase());

  const appendError = (message) => {
    if (result.errors.length < SLIDE_SYNC_VALIDATION.MAX_ERRORS) {
      result.errors.push(message);
    } else if (result.errors.length === SLIDE_SYNC_VALIDATION.MAX_ERRORS) {
      result.errors.push('Additional validation errors omitted for brevity.');
    }
  };

  if (rules.minRows != null && workingData.length < Number(rules.minRows)) {
    appendError(`Expected at least ${rules.minRows} data row(s) but found ${workingData.length}.`);
  }

  if (rules.maxRows != null && workingData.length > Number(rules.maxRows)) {
    appendError(`Expected no more than ${rules.maxRows} data row(s) but found ${workingData.length}.`);
  }

  const resolveColumns = (identifiers) => {
    const columns = [];
    const missing = [];
    identifiers.forEach((identifier) => {
      if (typeof identifier === 'number') {
        const index = identifier;
        const label = headerRow && headerRow[index] != null && headerRow[index] !== ''
          ? String(headerRow[index])
          : `Column ${index}`;
        columns.push({ index, label });
        return;
      }
      const normalized = String(identifier || '').trim().toLowerCase();
      const headerIndex = normalizedHeader.indexOf(normalized);
      if (headerIndex === -1) {
        missing.push(identifier);
      } else {
        const label = headerRow && headerRow[headerIndex] != null && headerRow[headerIndex] !== ''
          ? String(headerRow[headerIndex])
          : String(identifier);
        columns.push({ index: headerIndex, label });
      }
    });
    return { columns, missing };
  };

  if (rules.requiredColumns && rules.requiredColumns.length) {
    const { columns, missing } = resolveColumns(rules.requiredColumns);
    if (missing.length) {
      appendError(`Missing required column(s): ${missing.join(', ')}`);
    }
    if (columns.length) {
      workingData.forEach((row, rowIndex) => {
        const missingColumns = columns.filter((column) => !row || isEmptyValue(row[column.index]));
        if (missingColumns.length) {
          appendError(`Row ${rowIndex + 1}${hasHeader ? ' (excluding header)' : ''} is missing required values in ${missingColumns.map((column) => column.label).join(', ')}.`);
        }
      });
    }
  }

  if (rules.numericColumns && rules.numericColumns.length) {
    const { columns, missing } = resolveColumns(rules.numericColumns);
    if (missing.length) {
      appendError(`Numeric validation skipped - column(s) not found: ${missing.join(', ')}`);
    }
    if (columns.length) {
      workingData.forEach((row, rowIndex) => {
        columns.forEach((column) => {
          const value = row ? row[column.index] : null;
          if (value === '' || value === null || value === undefined) return;
          if (isNaN(Number(value))) {
            appendError(`Row ${rowIndex + 1}${hasHeader ? ' (excluding header)' : ''} column "${column.label}" expected numeric data but found "${value}".`);
          }
        });
      });
    }
  }

  result.valid = result.errors.length === 0;
  return result;
}

function enableScheduledRefresh() {
  ensureSchedulerTrigger();
  Logger.log(`Scheduler trigger enabled. Polling every ${SLIDE_SYNC_SCHEDULER.DEFAULT_POLL_MINUTES} minute(s).`);
}

function disableScheduledRefresh() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach((trigger) => {
    if (trigger.getHandlerFunction() === SLIDE_SYNC_SCHEDULER.HANDLER_FUNCTION) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  Logger.log('Scheduler trigger disabled.');
}

function ensureSchedulerTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  const existing = triggers.find((trigger) => trigger.getHandlerFunction() === SLIDE_SYNC_SCHEDULER.HANDLER_FUNCTION);
  if (!existing) {
    ScriptApp.newTrigger(SLIDE_SYNC_SCHEDULER.HANDLER_FUNCTION)
      .timeBased()
      .everyMinutes(SLIDE_SYNC_SCHEDULER.DEFAULT_POLL_MINUTES)
      .create();
  }
}

function pollScheduledSlides() {
  const updates = getConfiguredUpdates().filter((update) => update.scheduleEnabled && update.cronExpression);
  if (!updates.length) {
    Logger.log('No scheduled Slide Sync jobs are currently enabled.');
    return;
  }

  ensureSchedulerTrigger();

  const now = new Date();
  const properties = PropertiesService.getScriptProperties();
  const issues = [];

  updates.forEach((update) => {
    const lastRunKey = getScheduleStorageKey(update);
    const lastRun = properties.getProperty(lastRunKey);
    if (!shouldRunCronNow(update.cronExpression, now, lastRun)) {
      return;
    }

    Logger.log(`Running scheduled sync for ${update.sheetName} (cron: ${update.cronExpression})`);
    const issue = runParse(update);
    if (issue) {
      issues.push(issue);
    } else {
      properties.setProperty(lastRunKey, now.toISOString());
    }
  });

  if (issues.length) {
    Logger.log(`Scheduled refresh completed with ${issues.length} issue(s).`);
    issues.slice(0, SLIDE_SYNC_SCHEDULER.MAX_ERRORS_TO_LOG).forEach((issue) => Logger.log(issue));
  } else {
    Logger.log('Scheduled refresh completed successfully.');
  }
}

function getScheduleStorageKey(update) {
  return `${SLIDE_SYNC_SCHEDULER.LAST_RUN_PROPERTY_PREFIX}${update.identifier || buildUpdateIdentifier(update)}`;
}

function shouldRunCronNow(expression, now, lastRunIso) {
  const cron = parseCronExpression(expression);
  if (!cron) {
    Logger.log(`Invalid cron expression: ${expression}`);
    return false;
  }

  const minute = now.getMinutes();
  const hour = now.getHours();
  const day = now.getDate();
  const month = now.getMonth() + 1;
  const weekday = now.getDay();

  if (!cronFieldMatches(cron.minutes, minute)) return false;
  if (!cronFieldMatches(cron.hours, hour)) return false;
  if (!cronFieldMatches(cron.daysOfMonth, day)) return false;
  if (!cronFieldMatches(cron.months, month)) return false;
  if (!cronFieldMatches(cron.daysOfWeek, weekday)) return false;

  if (lastRunIso) {
    const lastRun = new Date(lastRunIso);
    if (!isNaN(lastRun.getTime())) {
      const lastMinute = Math.floor(lastRun.getTime() / 60000);
      const currentMinute = Math.floor(now.getTime() / 60000);
      if (lastMinute === currentMinute) {
        return false;
      }
    }
  }

  return true;
}

function parseCronExpression(expression) {
  if (!expression) return null;
  const parts = String(expression).trim().split(/\s+/);
  if (parts.length !== 5) return null;

  const [minute, hour, dayOfMonth, month, dayOfWeek] = parts;
  const cron = {
    minutes: parseCronField(minute, 0, 59),
    hours: parseCronField(hour, 0, 23),
    daysOfMonth: parseCronField(dayOfMonth, 1, 31),
    months: parseCronField(month, 1, 12),
    daysOfWeek: parseCronField(dayOfWeek, 0, 7)
  };

  if (cron.daysOfWeek) {
    cron.daysOfWeek = cron.daysOfWeek.map((value) => (value === 7 ? 0 : value));
  }

  if (Object.values(cron).some((value) => value === false)) {
    return null;
  }

  return cron;
}

function parseCronField(field, min, max) {
  const normalized = String(field || '').trim();
  if (!normalized || normalized === '*' || normalized === '?') return null;

  const values = new Set();
  const segments = normalized.split(',');
  for (let i = 0; i < segments.length; i++) {
    const segment = segments[i];
    const [rangePart, stepPart] = segment.split('/');
    const step = stepPart ? parseInt(stepPart, 10) : 1;
    if (!step || isNaN(step) || step < 1) {
      return false;
    }

    let start = min;
    let end = max;
    if (rangePart && rangePart !== '*') {
      if (rangePart.indexOf('-') > -1) {
        const [startStr, endStr] = rangePart.split('-');
        start = parseInt(startStr, 10);
        end = parseInt(endStr, 10);
      } else {
        start = parseInt(rangePart, 10);
        end = start;
      }
    }

    if (isNaN(start) || isNaN(end)) {
      return false;
    }

    start = Math.max(min, start);
    end = Math.min(max, end);
    if (start > end) {
      return false;
    }

    for (let value = start; value <= end; value += step) {
      values.add(value);
    }
  }

  return Array.from(values).sort((a, b) => a - b);
}

function cronFieldMatches(values, current) {
  if (!values || !values.length) return true;
  return values.indexOf(current) > -1;
}

function getSlideSyncJobs() {
  const updates = getConfiguredUpdates().filter((update) => update.configRowNumber);
  return updates.map((update) => ({
    identifier: update.identifier,
    sheetName: update.sheetName,
    spreadsheetId: update.spreadsheetId,
    slideId: update.slideId,
    pageId: update.pageId,
    matchText: update.matchText,
    cronExpression: update.cronExpression || '',
    scheduleEnabled: !!update.scheduleEnabled,
    templateKey: update.templateKey || update.appliedTemplate || '',
    validationRules: update.validationRules || null,
    configRowNumber: update.configRowNumber
  }));
}

function saveSlideSyncJobs(changes) {
  if (!changes || !changes.length) return { updated: 0 };

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) throw new Error('No active spreadsheet found.');
  const configSheet = spreadsheet.getSheetByName(SLIDE_SYNC_CONFIG.SHEET_NAME);
  if (!configSheet) throw new Error('Configuration sheet not found.');

  const values = configSheet.getDataRange().getValues();
  const headers = values[SLIDE_SYNC_CONFIG.HEADER_ROW_INDEX].map((header) => String(header || '').trim().toLowerCase());
  const scheduleColumn = headers.indexOf(SLIDE_SYNC_CONFIG.OPTIONAL_HEADERS.SCHEDULE_ENABLED);
  const cronColumn = headers.indexOf(SLIDE_SYNC_CONFIG.OPTIONAL_HEADERS.CRON_EXPRESSION);

  if (scheduleColumn === -1 && cronColumn === -1) {
    throw new Error('Schedule columns were not found in the configuration sheet.');
  }

  let updated = 0;
  changes.forEach((change) => {
    const rowNumber = Number(change.configRowNumber);
    if (!rowNumber || rowNumber <= SLIDE_SYNC_CONFIG.HEADER_ROW_INDEX) return;

    if (scheduleColumn > -1) {
      configSheet.getRange(rowNumber, scheduleColumn + 1).setValue(parseBoolean(change.scheduleEnabled));
    }
    if (cronColumn > -1) {
      configSheet.getRange(rowNumber, cronColumn + 1).setValue(change.cronExpression || '');
    }
    updated++;
  });

  return { updated };
}

function previewSlideSyncJob(identifier) {
  const update = findUpdateByIdentifier(identifier);
  if (!update) throw new Error('Could not locate configuration for the requested job.');
  return previewUpdate(update);
}

function runSlideSyncJob(identifier) {
  const update = findUpdateByIdentifier(identifier);
  if (!update) throw new Error('Could not locate configuration for the requested job.');
  return runParse(update);
}

function findUpdateByIdentifier(identifier) {
  if (!identifier) return null;
  const updates = getConfiguredUpdates();
  for (let i = 0; i < updates.length; i++) {
    if (updates[i].identifier === identifier) {
      return updates[i];
    }
  }
  return null;
}

function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('Slide Sync')
      .addItem('Refresh all slides', 'refreshAllSlides')
      .addItem('Preview changes (logs)', 'previewAllSlides')
      .addItem('Open manager sidebar', 'showSlideSyncSidebar')
      .addSeparator()
      .addItem('Enable scheduler', 'enableScheduledRefresh')
      .addItem('Disable scheduler', 'disableScheduledRefresh')
      .addToUi();
  } catch (e) {
    Logger.log(`Unable to register menu: ${e.message}`);
  }
}

function showSlideSyncSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('SlideSyncSidebar')
    .setTitle('Slide Sync Manager');
  SpreadsheetApp.getUi().showSidebar(html);
}
