/**
 * Cliniko Sync - Google Apps Script (V8)
 *
 * Setup:
 * 1) In Apps Script: File > Project properties > Script properties, add:
 *    - CLINIKO_API_KEY (required)
 *    - CLINIKO_BASE_URL (required, e.g., https://api.au1.cliniko.com/v1)
 *    - CLINIKO_CLINIC_ID (optional)
 *    - SHEET_ID (optional; defaults to active spreadsheet)
 *    - TIMEZONE (optional; defaults to spreadsheet time zone)
 * 2) Use the custom menu "Cliniko Sync" to run syncs.
 * 3) Use setupTriggers() to schedule daily syncs at 05:00.
 *
 * API key: Cliniko Settings > Integrations > API keys.
 * Base URL: depends on region (e.g., https://api.au1.cliniko.com/v1).
 */

// === Editable constants ===
var APPOINTMENTS_DAYS_PAST = 30;
var APPOINTMENTS_DAYS_FUTURE = 30;
var INVOICES_DAYS_PAST = 90;
var REPORT_WEEKS_BACK = 8;
var REPORT_SHEET_NAME = 'Patients_Without_Upcoming_Appointments';
var REPORT_PRACTITIONER_STATUS_COLUMN = 'Practitioner Follow-up';
var REPORT_KATE_ACTIONS_COLUMN = 'Kate Actions';
var REPORT_PRACTITIONER_STATUS_OPTIONS = [
  'referred to specialist contact and rebook',
  'discharged',
  'lost to follow up',
];
var REPORT_KATE_ACTIONS_OPTIONS = [
  'contacted rebooked',
  'contacted no response',
  'no action required',
];

// Endpoint paths (edit if Cliniko API paths differ)
var ENDPOINTS = {
  appointments: '/appointments',
  patients: '/patients',
  invoices: '/invoices',
};

// Query parameter names for date filtering (edit if needed)
var APPOINTMENT_START_PARAM = 'starts_at';
var APPOINTMENT_END_PARAM = 'ends_at';
var INVOICE_START_PARAM = 'invoice_date';
var INVOICE_END_PARAM = 'invoice_date';

// Curated columns (add or adjust as needed)
var APPOINTMENT_COLUMNS = [
  'id', 'created_at', 'updated_at', 'starts_at', 'ends_at',
  'status', 'appointment_type.name', 'practitioner.id', 'patient.id',
  'appointment_type.id', 'notes', 'cancelled_at', 'cancelled_by.id',
];
var PATIENT_COLUMNS = [
  'id', 'created_at', 'updated_at', 'first_name', 'last_name',
  'preferred_name', 'email', 'phone', 'mobile_phone', 'date_of_birth',
  'gender', 'clinic.id', 'patient_id',
];
var INVOICE_COLUMNS = [
  'id', 'created_at', 'updated_at', 'invoice_date', 'total',
  'status', 'patient.id', 'practitioner.id', 'clinic.id',
];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Cliniko Sync')
    .addItem('Set Config', 'setConfig')
    .addSeparator()
    .addItem('Sync Appointments', 'syncAppointments')
    .addItem('Sync Patients', 'syncPatients')
    .addItem('Sync Invoices', 'syncInvoices')
    .addSeparator()
    .addItem('Run Patients Without Upcoming Appointments Report', 'runPatientsWithoutUpcomingAppointmentsReport')
    .addItem('Setup Monthly Report Trigger', 'setupMonthlyReportTrigger')
    .addSeparator()
    .addItem('Sync All', 'syncAll')
    .addToUi();
}

function setConfig() {
  var ui = SpreadsheetApp.getUi();
  var props = PropertiesService.getScriptProperties();
  var apiKey = promptForConfig(ui, 'Cliniko API Key', props.getProperty('CLINIKO_API_KEY'));
  if (apiKey === null) {
    return;
  }
  var baseUrl = promptForConfig(ui, 'Cliniko Base URL', props.getProperty('CLINIKO_BASE_URL'));
  if (baseUrl === null) {
    return;
  }
  var clinicId = promptForConfig(ui, 'Cliniko Clinic ID (optional)', props.getProperty('CLINIKO_CLINIC_ID'));
  if (clinicId === null) {
    return;
  }
  var sheetId = promptForConfig(ui, 'Sheet ID (optional; blank for active)', props.getProperty('SHEET_ID'));
  if (sheetId === null) {
    return;
  }
  var timezone = promptForConfig(ui, 'Timezone (optional; e.g., Australia/Sydney)', props.getProperty('TIMEZONE'));
  if (timezone === null) {
    return;
  }
  var reportBaseUrl = promptForConfig(ui, 'Report Base URL (e.g., https://clinic.uk1.cliniko.com)', props.getProperty('CLINIKO_REPORT_BASE_URL'));
  if (reportBaseUrl === null) {
    return;
  }
  var reportBusinessId = promptForConfig(ui, 'Report Business ID (optional)', props.getProperty('CLINIKO_REPORT_BUSINESS_ID'));
  if (reportBusinessId === null) {
    return;
  }
  var practitionerIds = promptForConfig(ui, 'Practitioner IDs (comma-separated)', props.getProperty('PRACTITIONER_IDS'));
  if (practitionerIds === null) {
    return;
  }

  props.setProperties({
    CLINIKO_API_KEY: apiKey,
    CLINIKO_BASE_URL: baseUrl,
    CLINIKO_CLINIC_ID: clinicId,
    SHEET_ID: sheetId,
    TIMEZONE: timezone,
    CLINIKO_REPORT_BASE_URL: reportBaseUrl,
    CLINIKO_REPORT_BUSINESS_ID: reportBusinessId,
    PRACTITIONER_IDS: practitionerIds,
  });

  ui.alert('Cliniko configuration saved.');
}

function promptForConfig(ui, label, currentValue) {
  var response = ui.prompt(label, 'Current: ' + (currentValue || '(blank)'), ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) {
    return null;
  }
  return response.getResponseText().trim();
}

function syncAll() {
  syncAppointments();
  syncPatients();
  syncInvoices();
}

function syncAppointments() {
  var endpoint = ENDPOINTS.appointments;
  var start = addDays(new Date(), -APPOINTMENTS_DAYS_PAST);
  var end = addDays(new Date(), APPOINTMENTS_DAYS_FUTURE);
  var params = {};
  params[APPOINTMENT_START_PARAM] = toIsoDate(start);
  params[APPOINTMENT_END_PARAM] = toIsoDate(end);
  runSync(endpoint, 'Appointments', APPOINTMENT_COLUMNS, params);
}

function syncPatients() {
  var endpoint = ENDPOINTS.patients;
  runSync(endpoint, 'Patients', PATIENT_COLUMNS, {});
}

function syncInvoices() {
  var endpoint = ENDPOINTS.invoices;
  var start = addDays(new Date(), -INVOICES_DAYS_PAST);
  var params = {};
  params[INVOICE_START_PARAM] = toIsoDate(start);
  params[INVOICE_END_PARAM] = toIsoDate(new Date());
  runSync(endpoint, 'Invoices', INVOICE_COLUMNS, params);
}

function testConnection() {
  var start = new Date();
  var endpoint = ENDPOINTS.patients;
  var result;
  try {
    result = fetchCliniko(endpoint, { per_page: 1 });
    logRun(endpoint, result.items.length, durationMs(start), '');
  } catch (error) {
    logRun(endpoint, 0, durationMs(start), String(error));
    throw error;
  }
}

function setupTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function (trigger) {
    if (trigger.getHandlerFunction() === 'syncAll') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  var timezone = getConfig().timezone;
  ScriptApp.newTrigger('syncAll')
    .timeBased()
    .everyDays(1)
    .atHour(5)
    .inTimezone(timezone)
    .create();
}

function setupMonthlyReportTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function (trigger) {
    if (trigger.getHandlerFunction() === 'runPatientsWithoutUpcomingAppointmentsReport') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  var timezone = getConfig().timezone;
  ScriptApp.newTrigger('runPatientsWithoutUpcomingAppointmentsReport')
    .timeBased()
    .onMonthDay(1)
    .atHour(6)
    .inTimezone(timezone)
    .create();
}

function runPatientsWithoutUpcomingAppointmentsReport() {
  var config = getConfig();
  if (!config.reportBaseUrl) {
    throw new Error('Missing CLINIKO_REPORT_BASE_URL in script properties.');
  }
  var practitionerIds = config.practitionerIds;
  if (!practitionerIds.length) {
    throw new Error('Missing PRACTITIONER_IDS in script properties.');
  }

  var reportDate = new Date();
  var monthStart = new Date(reportDate.getFullYear(), reportDate.getMonth(), 1);
  var startDate = addDays(monthStart, -(REPORT_WEEKS_BACK * 7));
  var endDate = monthStart;

  var allRows = [];
  var headers = null;
  practitionerIds.forEach(function (practitionerId) {
    var data = fetchPatientsWithoutUpcomingAppointmentsReport(
      practitionerId,
      startDate,
      endDate,
      config.reportBusinessId
    );

    if (!headers) {
      headers = data.headers;
    }

    data.rows.forEach(function (row) {
      if (headers.indexOf('Practitioner Id') === -1) {
        row.push(practitionerId);
      }
      allRows.push(row);
    });
  });

  if (!headers) {
    headers = [];
  }

  if (headers.indexOf('Practitioner Id') === -1) {
    headers.push('Practitioner Id');
  }

  writeReportToSheet(REPORT_SHEET_NAME, headers, allRows);
}

function fetchPatientsWithoutUpcomingAppointmentsReport(practitionerId, startDate, endDate, businessId) {
  var config = getConfig();
  var endpoint = '/reports/patients/without_upcoming_appointments.csv';
  var params = {
    start_date: toReportDate(startDate),
    end_date: toReportDate(endDate),
  };
  if (practitionerId) {
    params['practitioner[id]'] = practitionerId;
  }
  if (businessId) {
    params['business[id]'] = businessId;
  }

  var url = buildUrl(config.reportBaseUrl + endpoint, params);
  var response = fetchWithRetry(url, config.apiKey, 'text/csv');
  var csv = Utilities.parseCsv(response.getContentText());
  if (!csv.length) {
    return { headers: [], rows: [] };
  }

  var headers = csv[0];
  var rows = csv.slice(1);
  return { headers: headers, rows: rows };
}

function writeReportToSheet(sheetName, headers, rows) {
  var sheet = getSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    sheet = getSpreadsheet().insertSheet(sheetName);
  }

  var extendedHeaders = headers.slice();
  if (extendedHeaders.indexOf(REPORT_PRACTITIONER_STATUS_COLUMN) === -1) {
    extendedHeaders.push(REPORT_PRACTITIONER_STATUS_COLUMN);
  }
  if (extendedHeaders.indexOf(REPORT_KATE_ACTIONS_COLUMN) === -1) {
    extendedHeaders.push(REPORT_KATE_ACTIONS_COLUMN);
  }

  var paddedRows = rows.map(function (row) {
    var newRow = row.slice();
    while (newRow.length < extendedHeaders.length) {
      newRow.push('');
    }
    return newRow;
  });

  sheet.clearContents();
  if (extendedHeaders.length > 0) {
    sheet.getRange(1, 1, 1, extendedHeaders.length).setValues([extendedHeaders]);
  }
  if (paddedRows.length > 0) {
    sheet.getRange(2, 1, paddedRows.length, extendedHeaders.length).setValues(paddedRows);
  }

  applyReportDropdowns(sheet, extendedHeaders, paddedRows.length);
}

function applyReportDropdowns(sheet, headers, rowCount) {
  if (!rowCount) {
    return;
  }

  var practitionerColumnIndex = headers.indexOf(REPORT_PRACTITIONER_STATUS_COLUMN) + 1;
  var kateColumnIndex = headers.indexOf(REPORT_KATE_ACTIONS_COLUMN) + 1;

  var practitionerValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(REPORT_PRACTITIONER_STATUS_OPTIONS, true)
    .setAllowInvalid(false)
    .build();
  var kateValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(REPORT_KATE_ACTIONS_OPTIONS, true)
    .setAllowInvalid(false)
    .build();

  if (practitionerColumnIndex > 0) {
    sheet.getRange(2, practitionerColumnIndex, rowCount).setDataValidation(practitionerValidation);
  }
  if (kateColumnIndex > 0) {
    sheet.getRange(2, kateColumnIndex, rowCount).setDataValidation(kateValidation);
  }
}

function runSync(endpoint, sheetName, curatedColumns, params) {
  var start = new Date();
  var rowsWritten = 0;
  var errorMessage = '';
  try {
    var data = fetchCliniko(endpoint, params);
    rowsWritten = writeToSheet(sheetName, data.items, curatedColumns);
  } catch (error) {
    errorMessage = String(error);
    throw error;
  } finally {
    logRun(endpoint, rowsWritten, durationMs(start), errorMessage);
  }
}

function fetchCliniko(endpoint, params) {
  var baseUrl = getConfig().baseUrl;
  var apiKey = getConfig().apiKey;
  var clinicId = getConfig().clinicId;

  var items = [];
  var page = 1;
  var perPage = 100;
  var nextUrl = null;
  var maxPages = 200;

  do {
    var currentParams = Object.assign({}, params);
    if (clinicId) {
      currentParams.clinic_id = clinicId;
    }
    if (!nextUrl) {
      currentParams.page = currentParams.page || page;
      currentParams.per_page = currentParams.per_page || perPage;
    }

    var url = nextUrl || buildUrl(baseUrl + endpoint, currentParams);
    var response = fetchWithRetry(url, apiKey);
    var data = JSON.parse(response.getContentText());
    var batch = extractItems(data);

    items = items.concat(batch);

    nextUrl = getNextUrl(data, baseUrl);
    if (!nextUrl && batch.length === perPage && page < maxPages) {
      page += 1;
    } else if (!nextUrl) {
      break;
    }
  } while (page <= maxPages);

  return { items: items };
}

function fetchWithRetry(url, apiKey, acceptHeader) {
  var maxAttempts = 5;
  var attempt = 0;
  var delay = 500;

  if (!apiKey) {
    throw new Error('Missing CLINIKO_API_KEY. Set it in Script Properties or via Cliniko Sync → Set Config.');
  }

  while (attempt < maxAttempts) {
    try {
      var response = UrlFetchApp.fetch(url, {
        headers: {
          Authorization: 'Basic ' + Utilities.base64Encode(apiKey + ':x'),
          Accept: acceptHeader || 'application/json',
        },
        muteHttpExceptions: true,
      });

      var status = response.getResponseCode();
      if (status === 429 || status === 503) {
        Utilities.sleep(delay);
        delay *= 2;
        attempt += 1;
        continue;
      }
      if (status >= 200 && status < 300) {
        return response;
      }
      if (status === 401) {
        throw new Error(
          'Cliniko API error (401): Unauthorized. Check that CLINIKO_API_KEY is correct and active, ' +
          'and that your CLINIKO_BASE_URL/CLINIKO_REPORT_BASE_URL match the Cliniko region for that key.'
        );
      }
      throw new Error('Cliniko API error (' + status + '): ' + response.getContentText());
    } catch (error) {
      attempt += 1;
      if (attempt >= maxAttempts) {
        throw error;
      }
      Utilities.sleep(delay);
      delay *= 2;
    }
  }
  throw new Error('Cliniko API request failed after retries.');
}

function extractItems(data) {
  if (!data || typeof data !== 'object') {
    return [];
  }
  if (Array.isArray(data.items)) {
    return data.items;
  }
  var keys = Object.keys(data);
  for (var i = 0; i < keys.length; i += 1) {
    var value = data[keys[i]];
    if (Array.isArray(value)) {
      return value;
    }
  }
  return [];
}

function getNextUrl(data, baseUrl) {
  if (!data || typeof data !== 'object') {
    return null;
  }
  var nextLink = null;
  if (data.links && data.links.next) {
    nextLink = data.links.next;
  } else if (data.next) {
    nextLink = data.next;
  }
  if (!nextLink) {
    return null;
  }
  if (nextLink.indexOf('http') === 0) {
    return nextLink;
  }
  return baseUrl + nextLink;
}

function writeToSheet(sheetName, items, curatedColumns) {
  var sheet = getSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    sheet = getSpreadsheet().insertSheet(sheetName);
  }

  var flattened = items.map(function (item) {
    return flattenObject(item);
  });

  var headers = buildHeaders(flattened, curatedColumns);
  var rows = flattened.map(function (row) {
    return headers.map(function (header) {
      return row[header] !== undefined ? row[header] : '';
    });
  });

  sheet.clearContents();
  if (headers.length > 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  return rows.length;
}

function buildHeaders(rows, curatedColumns) {
  var keys = {};
  rows.forEach(function (row) {
    Object.keys(row).forEach(function (key) {
      keys[key] = true;
    });
  });

  var extras = Object.keys(keys).filter(function (key) {
    return curatedColumns.indexOf(key) === -1;
  });
  extras.sort();
  return curatedColumns.concat(extras);
}

function flattenObject(obj) {
  var result = {};
  flattenHelper(obj, '', result);
  return result;
}

function flattenHelper(obj, prefix, result) {
  if (obj === null || obj === undefined) {
    return;
  }
  if (typeof obj !== 'object') {
    result[prefix] = obj;
    return;
  }
  if (Array.isArray(obj)) {
    result[prefix] = obj.map(function (item) {
      return typeof item === 'object' ? JSON.stringify(item) : item;
    }).join(', ');
    return;
  }
  Object.keys(obj).forEach(function (key) {
    var value = obj[key];
    var nextKey = prefix ? prefix + '.' + key : key;
    if (value && typeof value === 'object' && !Array.isArray(value)) {
      flattenHelper(value, nextKey, result);
    } else if (Array.isArray(value)) {
      result[nextKey] = value.map(function (item) {
        return typeof item === 'object' ? JSON.stringify(item) : item;
      }).join(', ');
    } else {
      result[nextKey] = value;
    }
  });
}

function logRun(endpoint, rowsWritten, duration, errorMessage) {
  var sheet = getSpreadsheet().getSheetByName('Sync_Log');
  if (!sheet) {
    sheet = getSpreadsheet().insertSheet('Sync_Log');
    sheet.getRange(1, 1, 1, 5).setValues([
      ['Timestamp', 'Endpoint', 'Rows Written', 'Duration (ms)', 'Error'],
    ]);
  }
  sheet.appendRow([
    new Date(),
    endpoint,
    rowsWritten,
    duration,
    errorMessage || '',
  ]);
}

function getSpreadsheet() {
  var sheetId = getConfig().sheetId;
  if (sheetId) {
    return SpreadsheetApp.openById(sheetId);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getConfig() {
  var props = PropertiesService.getScriptProperties();
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  return {
    apiKey: props.getProperty('CLINIKO_API_KEY') || '',
    baseUrl: props.getProperty('CLINIKO_BASE_URL') || '',
    clinicId: props.getProperty('CLINIKO_CLINIC_ID') || '',
    sheetId: props.getProperty('SHEET_ID') || '',
    timezone: props.getProperty('TIMEZONE') || spreadsheet.getSpreadsheetTimeZone(),
    reportBaseUrl: props.getProperty('CLINIKO_REPORT_BASE_URL') || '',
    reportBusinessId: props.getProperty('CLINIKO_REPORT_BUSINESS_ID') || '',
    practitionerIds: splitConfigList(props.getProperty('PRACTITIONER_IDS')),
  };
}

function addDays(date, days) {
  var copy = new Date(date.getTime());
  copy.setDate(copy.getDate() + days);
  return copy;
}

function toIsoDate(date) {
  return Utilities.formatDate(date, getConfig().timezone, 'yyyy-MM-dd');
}

function toReportDate(date) {
  return Utilities.formatDate(date, getConfig().timezone, 'd MMM yyyy');
}

function buildUrl(base, params) {
  var query = Object.keys(params).map(function (key) {
    return encodeURIComponent(key) + '=' + encodeURIComponent(params[key]);
  }).join('&');
  return query ? base + '?' + query : base;
}

function splitConfigList(value) {
  if (!value) {
    return [];
  }
  return value.split(',').map(function (entry) {
    return entry.trim();
  }).filter(function (entry) {
    return entry.length > 0;
  });
}
