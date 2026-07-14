const SHEETS = {
  EMPLOYEES: 'Employees',
  MASTER: 'Master',
  CONFIG: 'Config'
};

const EMPLOYEE_HEADERS = [
  'Employee ID', 'Employee Name', 'Department', 'Phone', 'Email', 'Joining Date', 'Status',
  'Device Token', 'Device Linked At'
];

const MASTER_HEADERS = [
  'Date', 'Employee ID', 'Employee Name', 'Department', 'In Time', 'Out Time',
  'Status', 'Late Minutes', 'Early Leave Minutes', 'OT Minutes', 'Working Minutes',
  'Remarks', 'Edited By', 'Last Updated'
];

const DEFAULT_CONFIG = {
  'Admin Password': 'admin123',
  'Timezone': 'Asia/Karachi',
  'Work Start Time': '10:30',
  'Work End Time': '20:00',
  'Late Allowed Minutes': '5',
  'Saturday Working': 'Yes',
  'Sunday Holiday': 'Yes',
  'OT Enabled': 'Yes',
  'Company Name': 'Attendance System'
};

function doGet(e) {
  setupDatabase();
  const page = String((e && e.parameter && e.parameter.page) || 'index').toLowerCase();
  const file = page === 'admin' ? 'Admin' : page === 'report' ? 'Report' : 'Index';
  return createTemplateFromFile_(file)
    .evaluate()
    .setTitle('Employee Attendance Management')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return createHtmlOutputFromFile_(filename).getContent();
}

function createTemplateFromFile_(filename) {
  try {
    return HtmlService.createTemplateFromFile(filename);
  } catch (err) {
    return HtmlService.createTemplateFromFile(filename + '.html');
  }
}

function createHtmlOutputFromFile_(filename) {
  try {
    return HtmlService.createHtmlOutputFromFile(filename);
  } catch (err) {
    return HtmlService.createHtmlOutputFromFile(filename + '.html');
  }
}

function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const employeeSheet = getOrCreateSheet_(ss, SHEETS.EMPLOYEES);
  const masterSheet = getOrCreateSheet_(ss, SHEETS.MASTER);
  const configSheet = getOrCreateSheet_(ss, SHEETS.CONFIG);

  ensureConfig_(configSheet);
  migrateEmployeeSheet_(employeeSheet);
  ensureHeaders_(employeeSheet, EMPLOYEE_HEADERS);
  migrateMasterSheet_(masterSheet, configSheet);
  ensureHeaders_(masterSheet, MASTER_HEADERS);

  employeeSheet.setFrozenRows(1);
  masterSheet.setFrozenRows(1);
  configSheet.setFrozenRows(1);
  formatSheets_(employeeSheet, masterSheet, configSheet);
  return { ok: true, message: 'Database is ready.' };
}

function repairMasterAttendance() {
  setupDatabase();
  const config = getConfigMap_();
  const tz = getTimezone_(config);
  const sheet = getSheet_(SHEETS.MASTER);
  const rows = getDataRows_(sheet);
  if (!rows.length) return { ok: true, repairedRows: 0 };

  const repaired = rows.map(row => {
    const date = normalizeDateString_(row[0], tz);
    const inTime = normalizeTimeString_(row[4], tz);
    const outTime = normalizeTimeString_(row[5], tz);
    const status = String(row[6] || (inTime ? 'Present' : '')).trim();
    const remarks = String(row[11] || (!outTime && inTime ? 'Incomplete Checkout' : '')).trim();
    const calc = calculateAttendance_({
      date: date,
      inTime: inTime,
      outTime: outTime,
      status: status,
      remarks: remarks
    }, config);
    return [
      date,
      normalizeId_(row[1]),
      String(row[2] || ''),
      String(row[3] || ''),
      inTime,
      outTime,
      status,
      calc.lateMinutes,
      calc.earlyLeaveMinutes,
      calc.otMinutes,
      calc.workingMinutes,
      remarks,
      String(row[12] || ''),
      row[13] instanceof Date ? timestamp_(tz, row[13]) : String(row[13] || timestamp_(tz))
    ];
  });

  sheet.getRange(2, 1, repaired.length, MASTER_HEADERS.length).setValues(repaired);
  invalidateCache_();
  return { ok: true, repairedRows: repaired.length };
}

function applyAttendanceSettings() {
  setupDatabase();
  const sheet = getSheet_(SHEETS.CONFIG);
  const rows = getDataRows_(sheet);
  const updates = {
    'Work Start Time': '10:30',
    'Work End Time': '20:00',
    'Late Allowed Minutes': '5'
  };
  rows.forEach((row, index) => {
    const key = String(row[0] || '').trim();
    if (updates[key] !== undefined) {
      sheet.getRange(index + 2, 2).setValue(updates[key]);
    }
  });
  invalidateCache_();
  return { ok: true, settings: updates };
}

function getAppBootstrap() {
  setupDatabase();
  const config = getConfigMap_();
  return {
    ok: true,
    companyName: config['Company Name'] || DEFAULT_CONFIG['Company Name'],
    timezone: getTimezone_(config),
    serverDate: formatDate_(new Date(), getTimezone_(config)),
    serverTime: formatTime_(new Date(), getTimezone_(config))
  };
}

function markAttendance(identifier) {
  return markAttendanceWithDevice(identifier, '');
}

function markAttendanceWithDevice(identifier, deviceToken) {
  setupDatabase();
  const cleanIdentifier = String(identifier || '').trim();
  if (!cleanIdentifier) return error_('Enter Employee ID or employee name.');
  const cleanDeviceToken = normalizeDeviceToken_(deviceToken);
  if (!cleanDeviceToken) return error_('Device token is missing. Refresh the page and try again.');

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return error_('The system is busy. Please submit again in a few seconds.');

  try {
    const config = getConfigMap_();
    const tz = getTimezone_(config);
    const now = new Date();
    const today = formatDate_(now, tz);
    const currentTime = formatTime_(now, tz);

    const lookup = findEmployeeByIdentifier_(cleanIdentifier);
    if (lookup.error) return error_(lookup.error);
    const employee = lookup.employee;
    if (!employee) return error_('Employee not found.');
    const employeeId = employee.employeeId;
    if (String(employee.status).toLowerCase() !== 'active') {
      return error_('This employee is inactive and cannot mark attendance.');
    }
    const deviceCheck = ensureEmployeeDevice_(employeeId, cleanDeviceToken, tz);
    if (!deviceCheck.ok) return deviceCheck;
    employee.deviceToken = cleanDeviceToken;
    employee.deviceLinkedAt = deviceCheck.deviceLinkedAt || employee.deviceLinkedAt;

    const cache = CacheService.getScriptCache();
    const duplicateKey = 'entry:' + employeeId;
    if (cache.get(duplicateKey)) {
      return error_('Duplicate entry blocked. Please wait a few seconds before submitting again.');
    }
    cache.put(duplicateKey, '1', 20);

    const holiday = getHolidayForDate_(today, config);
    if (holiday) return error_('Attendance is disabled today: ' + holiday.name + '.');
    if (isWeeklyHoliday_(now, config, tz)) return error_('Attendance is disabled on a configured weekly holiday.');

    const masterSheet = getSheet_(SHEETS.MASTER);
    const values = getDataRows_(masterSheet);
    const rowIndex = values.findIndex(row => sameDateValue_(row[0], today, tz) && normalizeId_(row[1]) === employeeId);

    if (rowIndex === -1) {
      const calc = calculateAttendance_({
        date: today,
        inTime: currentTime,
        outTime: '',
        status: 'Present',
        remarks: 'Incomplete Checkout'
      }, config);
      const record = [
        today,
        employeeId,
        employee.name,
        employee.department,
        currentTime,
        '',
        'Present',
        calc.lateMinutes,
        calc.earlyLeaveMinutes,
        calc.otMinutes,
        calc.workingMinutes,
        'Incomplete Checkout',
        '',
        timestamp_(tz)
      ];
      while (record.length < MASTER_HEADERS.length) record.push('');
      masterSheet.appendRow(record);
      invalidateCache_();
      return {
        ok: true,
        action: 'checkin',
        message: 'Check-in recorded.',
        employee: employee,
        time: currentTime,
        date: today,
        dayState: 'checked_in',
        record: rowToAttendanceObject_(record, tz)
      };
    }

    const sheetRow = rowIndex + 2;
    const existing = values[rowIndex];
    if (String(existing[5] || '').trim()) {
      return error_('Check-out is already recorded for today.');
    }

    const updated = existing.slice(0, MASTER_HEADERS.length);
    while (updated.length < MASTER_HEADERS.length) updated.push('');
    updated[5] = currentTime;
    updated[6] = 'Present';
    updated[11] = '';
    updated[12] = '';
    updated[13] = timestamp_(tz);
    const calc = calculateAttendance_({
      date: today,
      inTime: updated[4],
      outTime: currentTime,
      status: 'Present',
      remarks: ''
    }, config);
    updated[7] = calc.lateMinutes;
    updated[8] = calc.earlyLeaveMinutes;
    updated[9] = calc.otMinutes;
    updated[10] = calc.workingMinutes;

    masterSheet.getRange(sheetRow, 1, 1, MASTER_HEADERS.length).setValues([updated]);
    invalidateCache_();
    return {
      ok: true,
      action: 'checkout',
      message: 'Check-out recorded.',
      employee: employee,
      time: currentTime,
      date: today,
      dayState: 'completed',
      record: rowToAttendanceObject_(updated, tz)
    };
  } catch (err) {
    return error_(err.message || String(err));
  } finally {
    lock.releaseLock();
  }
}

function getDeviceStatus(deviceToken) {
  setupDatabase();
  const cleanDeviceToken = normalizeDeviceToken_(deviceToken);
  if (!cleanDeviceToken) return error_('Device token is missing.');
  const config = getConfigMap_();
  const tz = getTimezone_(config);
  const employee = listEmployeesInternal_().find(e => e.deviceToken === cleanDeviceToken);
  if (!employee) {
    return {
      ok: true,
      linked: false,
      dayState: 'unlinked',
      message: 'This device is not linked yet.'
    };
  }
  if (String(employee.status).toLowerCase() !== 'active') {
    return error_('This linked employee is inactive. Contact admin.');
  }
  const status = getTodayAttendanceStatus_(employee.employeeId, config);
  return {
    ok: true,
    linked: true,
    employee: employee,
    dayState: status.dayState,
    record: status.record,
    message: status.message
  };
}

function adminLogin(password) {
  setupDatabase();
  const config = getConfigMap_();
  const expected = String(config['Admin Password'] || DEFAULT_CONFIG['Admin Password']);
  const supplied = String(password || '');
  if (!supplied || supplied !== expected) return error_('Invalid admin password.');
  return { ok: true, token: makeAdminToken_(), companyName: config['Company Name'] || '' };
}

function getDashboardData(token) {
  requireAdmin_(token);
  const config = getConfigMap_();
  const tz = getTimezone_(config);
  const today = formatDate_(new Date(), tz);
  const employees = listEmployeesInternal_();
  const attendance = listAttendanceInternal_({ date: today });
  const activeCount = employees.filter(e => e.status === 'Active').length;
  const todayPresent = attendance.filter(r => r.inTime).length;
  const incomplete = attendance.filter(r => r.inTime && !r.outTime).length;
  const late = attendance.filter(r => Number(r.lateMinutes) > 0).length;
  const totalOt = attendance.reduce((sum, r) => sum + Number(r.otMinutes || 0), 0);
  return {
    ok: true,
    cards: {
      totalEmployees: employees.length,
      activeEmployees: activeCount,
      todayPresent: todayPresent,
      incompleteRecords: incomplete,
      lateEmployees: late,
      totalOtMinutes: totalOt
    },
    today: today,
    recentAttendance: attendance.slice(0, 20)
  };
}

function listEmployees(token, query) {
  requireAdmin_(token);
  const q = String(query || '').toLowerCase().trim();
  const data = listEmployeesInternal_();
  return {
    ok: true,
    employees: q ? data.filter(e => [
      e.employeeId, e.name, e.department, e.phone, e.email, e.status
    ].join(' ').toLowerCase().indexOf(q) !== -1) : data
  };
}

function resetEmployeeDevice(token, employeeId) {
  requireAdmin_(token);
  setupDatabase();
  const cleanId = normalizeId_(employeeId);
  if (!cleanId) throw new Error('Employee ID is required.');
  const sheet = getSheet_(SHEETS.EMPLOYEES);
  const rows = getDataRows_(sheet);
  const rowIndex = rows.findIndex(row => normalizeId_(row[0]) === cleanId);
  if (rowIndex === -1) throw new Error('Employee not found.');
  sheet.getRange(rowIndex + 2, 8, 1, 2).clearContent();
  invalidateCache_();
  return { ok: true, message: 'Device reset.' };
}

function saveEmployee(token, employee) {
  requireAdmin_(token);
  setupDatabase();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const clean = sanitizeEmployee_(employee);
    const sheet = getSheet_(SHEETS.EMPLOYEES);
    const rows = getDataRows_(sheet);
    const existingIndex = rows.findIndex(row => normalizeId_(row[0]) === clean.employeeId);
    const existing = existingIndex === -1 ? [] : rows[existingIndex];
    const row = [
      clean.employeeId,
      clean.name,
      clean.department,
      clean.phone,
      clean.email,
      clean.joiningDate,
      clean.status,
      existing[7] || '',
      existing[8] || ''
    ];
    if (existingIndex === -1) {
      sheet.appendRow(row);
    } else {
      sheet.getRange(existingIndex + 2, 1, 1, EMPLOYEE_HEADERS.length).setValues([row]);
    }
    invalidateCache_();
    return { ok: true, employee: rowToEmployeeObject_(row) };
  } finally {
    lock.releaseLock();
  }
}

function listAttendance(token, filters) {
  requireAdmin_(token);
  return {
    ok: true,
    records: listAttendanceInternal_(filters || {})
  };
}

function saveAttendanceRecord(token, record) {
  requireAdmin_(token);
  setupDatabase();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const config = getConfigMap_();
    const tz = getTimezone_(config);
    const clean = sanitizeAttendance_(record);
    const employee = findEmployeeById_(clean.employeeId);
    if (!employee) throw new Error('Employee not found.');

    const calc = calculateAttendance_(clean, config);
    const row = [
      clean.date,
      clean.employeeId,
      employee.name,
      employee.department,
      clean.inTime,
      clean.outTime,
      clean.status,
      calc.lateMinutes,
      calc.earlyLeaveMinutes,
      calc.otMinutes,
      calc.workingMinutes,
      clean.remarks,
      clean.editedBy || 'Admin',
      timestamp_(tz)
    ];
    while (row.length < MASTER_HEADERS.length) row.push('');

    const sheet = getSheet_(SHEETS.MASTER);
    const rows = getDataRows_(sheet);
    let rowIndex = -1;
    if (clean.originalDate && clean.originalEmployeeId) {
      rowIndex = rows.findIndex(r => sameDateValue_(r[0], clean.originalDate, tz) && normalizeId_(r[1]) === clean.originalEmployeeId);
    }
    if (rowIndex === -1) {
      rowIndex = rows.findIndex(r => sameDateValue_(r[0], clean.date, tz) && normalizeId_(r[1]) === clean.employeeId);
    }
    if (rowIndex === -1) {
      sheet.appendRow(row);
    } else {
      sheet.getRange(rowIndex + 2, 1, 1, MASTER_HEADERS.length).setValues([row]);
    }
    invalidateCache_();
    return { ok: true, record: rowToAttendanceObject_(row, tz) };
  } finally {
    lock.releaseLock();
  }
}

function getSettings(token) {
  requireAdmin_(token);
  const config = getConfigMap_();
  return {
    ok: true,
    settings: Object.keys(DEFAULT_CONFIG).map(key => ({ key: key, value: config[key] || '' })),
    holidays: getHolidays_(config)
  };
}

function saveSettings(token, settings, holidays) {
  requireAdmin_(token);
  setupDatabase();
  const sheet = getSheet_(SHEETS.CONFIG);
  const rows = [['Key', 'Value', 'Name/Notes']];
  const incoming = {};
  (settings || []).forEach(item => {
    if (item && item.key) incoming[String(item.key)] = String(item.value || '');
  });
  Object.keys(DEFAULT_CONFIG).forEach(key => {
    rows.push([key, incoming[key] || DEFAULT_CONFIG[key], '']);
  });
  (holidays || []).forEach((holiday, index) => {
    if (!holiday || !holiday.date) return;
    rows.push([
      'HOLIDAY_' + String(index + 1).padStart(2, '0'),
      normalizeDateString_(holiday.date, getTimezone_(incoming)),
      String(holiday.name || 'Holiday').trim()
    ]);
  });
  sheet.clearContents();
  sheet.getRange(1, 1, rows.length, 3).setValues(rows);
  invalidateCache_();
  return getSettings(token);
}

function getReportData(token, filters) {
  requireAdmin_(token);
  const config = getConfigMap_();
  return {
    ok: true,
    companyName: config['Company Name'] || DEFAULT_CONFIG['Company Name'],
    generatedAt: timestamp_(getTimezone_(config)),
    records: listAttendanceInternal_(filters || {})
  };
}

function createExcelReport(token, filters) {
  requireAdmin_(token);
  const report = getReportData(token, filters || {});
  const monthLabel = (filters && filters.month) || formatDate_(new Date(), getTimezone_(getConfigMap_())).slice(0, 7);
  const ss = SpreadsheetApp.create('Attendance Report ' + monthLabel);
  const sheet = ss.getActiveSheet();
  sheet.setName('Attendance Report');
  const header = [
    'Date', 'Employee ID', 'Employee Name', 'Department', 'IN Time', 'OUT Time',
    'Status', 'Late Minutes', 'OT Minutes', 'Working Minutes', 'Remarks'
  ];
  const rows = report.records.map(r => [
    r.date, r.employeeId, r.employeeName, r.department, r.inTime, r.outTime,
    r.status, r.lateMinutes, r.otMinutes, r.workingMinutes, r.remarks
  ]);
  sheet.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
  if (rows.length) sheet.getRange(2, 1, rows.length, header.length).setValues(rows);
  sheet.autoResizeColumns(1, header.length);
  return { ok: true, url: ss.getUrl() };
}

function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function getSheet_(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) throw new Error('Missing sheet: ' + name + '. Run setupDatabase first.');
  return sheet;
}

function ensureHeaders_(sheet, headers) {
  const range = sheet.getRange(1, 1, 1, headers.length);
  const existing = range.getValues()[0];
  const hasHeaders = existing.some(v => String(v || '').trim());
  if (!hasHeaders || headers.some((h, i) => String(existing[i] || '') !== h)) {
    range.setValues([headers]);
  }
}

function ensureConfig_(sheet) {
  const first = sheet.getRange(1, 1, 1, 3).getValues()[0];
  if (!first[0]) sheet.getRange(1, 1, 1, 3).setValues([['Key', 'Value', 'Name/Notes']]);
  const values = sheet.getDataRange().getValues();
  const existing = {};
  values.slice(1).forEach(row => {
    if (row[0]) existing[String(row[0])] = true;
  });
  const missing = Object.keys(DEFAULT_CONFIG)
    .filter(key => !existing[key])
    .map(key => [key, DEFAULT_CONFIG[key], '']);
  if (missing.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, missing.length, 3).setValues(missing);
  }
}

function migrateEmployeeSheet_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return;

  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const hasOldQrColumn = header[5] === 'QR Code';
  const missingDeviceColumns = header.indexOf('Device Token') === -1 || header.indexOf('Device Linked At') === -1;
  if (!hasOldQrColumn && !missingDeviceColumns) return;

  const oldRows = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];
  const idIndex = header.indexOf('Employee ID') === -1 ? 0 : header.indexOf('Employee ID');
  const nameIndex = header.indexOf('Employee Name') === -1 ? 1 : header.indexOf('Employee Name');
  const departmentIndex = header.indexOf('Department') === -1 ? 2 : header.indexOf('Department');
  const phoneIndex = header.indexOf('Phone') === -1 ? 3 : header.indexOf('Phone');
  const emailIndex = header.indexOf('Email') === -1 ? 4 : header.indexOf('Email');
  const joiningIndex = header.indexOf('Joining Date') === -1 ? (hasOldQrColumn ? 6 : 5) : header.indexOf('Joining Date');
  const statusIndex = header.indexOf('Status') === -1 ? (hasOldQrColumn ? 7 : 6) : header.indexOf('Status');
  const deviceTokenIndex = header.indexOf('Device Token');
  const deviceLinkedIndex = header.indexOf('Device Linked At');
  const newRows = oldRows
    .filter(row => row[idIndex] || row[nameIndex])
    .map(row => [
      normalizeId_(row[idIndex]),
      String(row[nameIndex] || ''),
      String(row[departmentIndex] || ''),
      String(row[phoneIndex] || ''),
      String(row[emailIndex] || ''),
      row[joiningIndex] instanceof Date ? row[joiningIndex] : normalizeDateString_(row[joiningIndex], DEFAULT_CONFIG.Timezone),
      String(row[statusIndex] || 'Active'),
      deviceTokenIndex === -1 ? '' : String(row[deviceTokenIndex] || ''),
      deviceLinkedIndex === -1 ? '' : (row[deviceLinkedIndex] instanceof Date ? timestamp_(DEFAULT_CONFIG.Timezone, row[deviceLinkedIndex]) : String(row[deviceLinkedIndex] || ''))
    ]);

  sheet.clearContents();
  sheet.getRange(1, 1, 1, EMPLOYEE_HEADERS.length).setValues([EMPLOYEE_HEADERS]);
  if (newRows.length) {
    sheet.getRange(2, 1, newRows.length, EMPLOYEE_HEADERS.length).setValues(newRows);
  }
  const extraColumns = sheet.getMaxColumns() - EMPLOYEE_HEADERS.length;
  if (extraColumns > 0) sheet.deleteColumns(EMPLOYEE_HEADERS.length + 1, extraColumns);
}

function migrateMasterSheet_(sheet, configSheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return;

  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const isOldLayout = header[0] === 'Attendance ID' && header[5] === 'Shift';
  if (!isOldLayout) return;

  const config = configRowsToMap_(getDataRows_(configSheet));
  const tz = getTimezone_(config);
  const oldRows = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];
  const newRows = oldRows
    .filter(row => row[1] || row[2])
    .map(row => {
      const date = normalizeDateString_(row[1], tz);
      const inTime = normalizeTimeString_(row[6], tz);
      const outTime = normalizeTimeString_(row[7], tz);
      const status = String(row[8] || 'Present').trim();
      const remarks = String(row[13] || (!outTime && inTime ? 'Incomplete Checkout' : '')).trim();
      const calc = calculateAttendance_({
        date: date,
        inTime: inTime,
        outTime: outTime,
        status: status,
        remarks: remarks
      }, config);
      return [
        date,
        normalizeId_(row[2]),
        String(row[3] || ''),
        String(row[4] || ''),
        inTime,
        outTime,
        status,
        calc.lateMinutes,
        calc.earlyLeaveMinutes,
        calc.otMinutes,
        calc.workingMinutes,
        remarks,
        String(row[14] || ''),
        row[15] instanceof Date ? timestamp_(tz, row[15]) : String(row[15] || timestamp_(tz))
      ];
    });

  sheet.clearContents();
  sheet.getRange(1, 1, 1, MASTER_HEADERS.length).setValues([MASTER_HEADERS]);
  if (newRows.length) {
    sheet.getRange(2, 1, newRows.length, MASTER_HEADERS.length).setValues(newRows);
  }
  const extraColumns = sheet.getMaxColumns() - MASTER_HEADERS.length;
  if (extraColumns > 0) sheet.deleteColumns(MASTER_HEADERS.length + 1, extraColumns);
}

function formatSheets_(employeeSheet, masterSheet, configSheet) {
  [employeeSheet, masterSheet, configSheet].forEach(sheet => {
    const maxColumns = sheet.getLastColumn();
    if (maxColumns > 0) {
      sheet.getRange(1, 1, 1, maxColumns).setFontWeight('bold').setBackground('#0f8f58').setFontColor('#ffffff');
      sheet.autoResizeColumns(1, maxColumns);
    }
  });
}

function getDataRows_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];
  return sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
}

function getConfigMap_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('config');
  if (cached) return JSON.parse(cached);
  setupDatabase();
  const config = configRowsToMap_(getDataRows_(getSheet_(SHEETS.CONFIG)));
  cache.put('config', JSON.stringify(config), 300);
  return config;
}

function configRowsToMap_(rows) {
  const config = {};
  rows.forEach(row => {
    if (!row[0]) return;
    const key = String(row[0]).trim();
    config[key] = normalizeConfigValue_(key, row[1], config);
    if (row[0] && String(row[0]).indexOf('HOLIDAY_') === 0) {
      config[key + '_NAME'] = String(row[2] || '').trim();
    }
  });
  Object.keys(DEFAULT_CONFIG).forEach(key => {
    if (!config[key]) config[key] = DEFAULT_CONFIG[key];
  });
  return config;
}

function normalizeConfigValue_(key, value, config) {
  const tz = getTimezone_(config || {});
  if (key === 'Work Start Time' || key === 'Work End Time') {
    return normalizeTimeString_(value, tz) || DEFAULT_CONFIG[key];
  }
  if (key.indexOf('HOLIDAY_') === 0) {
    return normalizeDateString_(value, tz);
  }
  if (value instanceof Date) return formatDate_(value, tz);
  return String(value || '').trim();
}

function listEmployeesInternal_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('employees');
  if (cached) return JSON.parse(cached);
  setupDatabase();
  const employees = getDataRows_(getSheet_(SHEETS.EMPLOYEES))
    .filter(row => row[0])
    .map(rowToEmployeeObject_);
  cache.put('employees', JSON.stringify(employees), 180);
  return employees;
}

function listAttendanceInternal_(filters) {
  setupDatabase();
  const config = getConfigMap_();
  const tz = getTimezone_(config);
  const q = String(filters.query || '').toLowerCase().trim();
  const month = String(filters.month || '').trim();
  const date = String(filters.date || '').trim();
  const employee = normalizeId_(filters.employeeId || filters.employee || '');
  const department = String(filters.department || '').toLowerCase().trim();
  const status = String(filters.status || '').toLowerCase().trim();

  return getDataRows_(getSheet_(SHEETS.MASTER))
    .filter(row => row[0])
    .map(row => rowToAttendanceObject_(row, tz))
    .filter(r => {
      if (date && r.date !== normalizeDateString_(date, tz)) return false;
      if (month && r.date.slice(0, 7) !== month) return false;
      if (employee && r.employeeId !== employee) return false;
      if (department && String(r.department || '').toLowerCase() !== department) return false;
      if (status && String(r.status || '').toLowerCase() !== status) return false;
      if (q) {
        const haystack = [
          r.date, r.employeeId, r.employeeName, r.department, r.status, r.remarks
        ].join(' ').toLowerCase();
        if (haystack.indexOf(q) === -1) return false;
      }
      return true;
    })
    .sort((a, b) => String(b.date + b.inTime).localeCompare(String(a.date + a.inTime)));
}

function findEmployeeById_(employeeId) {
  const id = normalizeId_(employeeId);
  return listEmployeesInternal_().find(e => e.employeeId === id) || null;
}

function findEmployeeByIdentifier_(identifier) {
  const text = String(identifier || '').trim();
  const id = normalizeId_(text);
  const employees = listEmployeesInternal_();
  const byId = employees.find(e => e.employeeId === id);
  if (byId) return { employee: byId };

  const name = normalizeName_(text);
  const exactNameMatches = employees.filter(e => normalizeName_(e.name) === name);
  if (exactNameMatches.length === 1) return { employee: exactNameMatches[0] };
  if (exactNameMatches.length > 1) {
    return { error: 'Multiple employees have this name. Please enter Employee ID.' };
  }

  const partialNameMatches = employees.filter(e => normalizeName_(e.name).indexOf(name) !== -1);
  if (partialNameMatches.length === 1) return { employee: partialNameMatches[0] };
  if (partialNameMatches.length > 1) {
    return { error: 'Multiple employees match this name. Please enter Employee ID.' };
  }

  return { employee: null };
}

function ensureEmployeeDevice_(employeeId, deviceToken, tz) {
  const sheet = getSheet_(SHEETS.EMPLOYEES);
  const rows = getDataRows_(sheet);
  const rowIndex = rows.findIndex(row => normalizeId_(row[0]) === employeeId);
  if (rowIndex === -1) return error_('Employee not found.');

  const existingToken = normalizeDeviceToken_(rows[rowIndex][7]);
  if (existingToken && existingToken !== deviceToken) {
    return error_('This employee is already linked to another device. Contact admin to reset device.');
  }

  const otherEmployee = rows.find(row => normalizeId_(row[0]) !== employeeId && normalizeDeviceToken_(row[7]) === deviceToken);
  if (otherEmployee) {
    return error_('This device is already linked to ' + String(otherEmployee[1] || 'another employee') + '.');
  }

  if (!existingToken) {
    const linkedAt = timestamp_(tz);
    sheet.getRange(rowIndex + 2, 8, 1, 2).setValues([[deviceToken, linkedAt]]);
    invalidateCache_();
    return { ok: true, deviceLinkedAt: linkedAt };
  }

  return { ok: true, deviceLinkedAt: rows[rowIndex][8] };
}

function getTodayAttendanceStatus_(employeeId, config) {
  const tz = getTimezone_(config);
  const today = formatDate_(new Date(), tz);
  const rows = getDataRows_(getSheet_(SHEETS.MASTER));
  const row = rows.find(item => sameDateValue_(item[0], today, tz) && normalizeId_(item[1]) === employeeId);
  if (!row) {
    return {
      dayState: 'none',
      record: null,
      message: 'Ready for check-in.'
    };
  }
  const record = rowToAttendanceObject_(row, tz);
  if (record.outTime) {
    return {
      dayState: 'completed',
      record: record,
      message: 'Attendance completed today. Come tomorrow.'
    };
  }
  return {
    dayState: 'checked_in',
    record: record,
    message: 'Check-in recorded. Ready for check-out.'
  };
}

function rowToEmployeeObject_(row) {
  return {
    employeeId: normalizeId_(row[0]),
    name: String(row[1] || ''),
    department: String(row[2] || ''),
    phone: String(row[3] || ''),
    email: String(row[4] || ''),
    joiningDate: row[5] instanceof Date ? formatDate_(row[5], getTimezone_(getConfigMap_())) : String(row[5] || ''),
    status: String(row[6] || 'Active'),
    deviceToken: String(row[7] || ''),
    deviceLinkedAt: row[8] instanceof Date ? timestamp_(getTimezone_(getConfigMap_()), row[8]) : String(row[8] || '')
  };
}

function rowToAttendanceObject_(row, tz) {
  const date = row[0] instanceof Date ? formatDate_(row[0], tz) : normalizeDateString_(row[0], tz);
  const employeeId = normalizeId_(row[1]);
  return {
    recordKey: makeRecordKey_(date, employeeId),
    date: date,
    employeeId: employeeId,
    employeeName: String(row[2] || ''),
    department: String(row[3] || ''),
    inTime: normalizeTimeString_(row[4], tz),
    outTime: normalizeTimeString_(row[5], tz),
    status: String(row[6] || ''),
    lateMinutes: Number(row[7] || 0),
    earlyLeaveMinutes: Number(row[8] || 0),
    otMinutes: Number(row[9] || 0),
    workingMinutes: Number(row[10] || 0),
    remarks: String(row[11] || ''),
    editedBy: String(row[12] || ''),
    lastUpdated: row[13] instanceof Date ? timestamp_(tz, row[13]) : String(row[13] || '')
  };
}

function sanitizeEmployee_(employee) {
  const clean = {
    employeeId: normalizeId_(employee && employee.employeeId),
    name: String(employee && employee.name || '').trim(),
    department: String(employee && employee.department || '').trim(),
    phone: String(employee && employee.phone || '').trim(),
    email: String(employee && employee.email || '').trim(),
    joiningDate: normalizeDateString_(employee && employee.joiningDate || formatDate_(new Date(), getTimezone_(getConfigMap_())), getTimezone_(getConfigMap_())),
    status: String(employee && employee.status || 'Active').trim(),
    deviceToken: normalizeDeviceToken_(employee && employee.deviceToken),
    deviceLinkedAt: String(employee && employee.deviceLinkedAt || '').trim()
  };
  if (!clean.employeeId) throw new Error('Employee ID is required.');
  if (!clean.name) throw new Error('Employee name is required.');
  if (['Active', 'Inactive'].indexOf(clean.status) === -1) clean.status = 'Active';
  return clean;
}

function sanitizeAttendance_(record) {
  const config = getConfigMap_();
  const tz = getTimezone_(config);
  const clean = {
    date: normalizeDateString_(record && record.date || formatDate_(new Date(), tz), tz),
    employeeId: normalizeId_(record && record.employeeId),
    originalDate: normalizeDateString_(record && record.originalDate || '', tz),
    originalEmployeeId: normalizeId_(record && record.originalEmployeeId),
    inTime: normalizeTimeString_(record && record.inTime),
    outTime: normalizeTimeString_(record && record.outTime),
    status: String(record && record.status || 'Present').trim(),
    remarks: String(record && record.remarks || '').trim(),
    editedBy: String(record && record.editedBy || 'Admin').trim()
  };
  if (!clean.date) throw new Error('Date is required.');
  if (!clean.employeeId) throw new Error('Employee is required.');
  if (!clean.inTime && !clean.outTime) throw new Error('At least one time value is required.');
  if (!clean.status) clean.status = 'Present';
  if (clean.inTime && clean.outTime && timeToMinutes_(clean.outTime) < timeToMinutes_(clean.inTime)) {
    throw new Error('Out Time cannot be before In Time.');
  }
  return clean;
}

function calculateAttendance_(record, config) {
  const start = timeToMinutes_(config['Work Start Time'] || DEFAULT_CONFIG['Work Start Time']);
  const end = timeToMinutes_(config['Work End Time'] || DEFAULT_CONFIG['Work End Time']);
  const allowedLate = Number(config['Late Allowed Minutes'] || 0);
  const inMinutes = record.inTime ? timeToMinutes_(record.inTime) : null;
  const outMinutes = record.outTime ? timeToMinutes_(record.outTime) : null;
  const lateMinutes = inMinutes === null ? 0 : Math.max(0, inMinutes - start - allowedLate);
  const earlyLeaveMinutes = outMinutes === null ? 0 : Math.max(0, end - outMinutes);
  const otEnabled = String(config['OT Enabled'] || 'Yes').toLowerCase() === 'yes';
  const otMinutes = otEnabled && outMinutes !== null ? Math.max(0, outMinutes - end) : 0;
  const workingMinutes = inMinutes !== null && outMinutes !== null ? Math.max(0, outMinutes - inMinutes) : 0;
  return {
    lateMinutes: lateMinutes,
    earlyLeaveMinutes: earlyLeaveMinutes,
    otMinutes: otMinutes,
    workingMinutes: workingMinutes
  };
}

function getHolidays_(config) {
  return Object.keys(config)
    .filter(key => /^HOLIDAY_\d+$/.test(key) && config[key])
    .sort()
    .map(key => ({
      key: key,
      date: normalizeDateString_(config[key], getTimezone_(config)),
      name: config[key + '_NAME'] || 'Holiday'
    }));
}

function getHolidayForDate_(date, config) {
  const normalized = normalizeDateString_(date, getTimezone_(config));
  return getHolidays_(config).find(h => h.date === normalized) || null;
}

function isWeeklyHoliday_(dateObj, config, tz) {
  const day = Number(Utilities.formatDate(dateObj, tz, 'u'));
  const saturdayWorking = String(config['Saturday Working'] || 'Yes').toLowerCase() === 'yes';
  const sundayHoliday = String(config['Sunday Holiday'] || 'Yes').toLowerCase() === 'yes';
  if (day === 6 && !saturdayWorking) return true;
  if (day === 7 && sundayHoliday) return true;
  return false;
}

function requireAdmin_(token) {
  const cache = CacheService.getScriptCache();
  if (!token || !cache.get('admin:' + token)) throw new Error('Admin session expired. Please log in again.');
}

function makeAdminToken_() {
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('admin:' + token, '1', 21600);
  return token;
}

function invalidateCache_() {
  const cache = CacheService.getScriptCache();
  cache.remove('employees');
  cache.remove('config');
}

function makeRecordKey_(date, employeeId) {
  return normalizeDateString_(date, DEFAULT_CONFIG.Timezone) + '|' + normalizeId_(employeeId);
}

function getTimezone_(config) {
  return String((config && config.Timezone) || DEFAULT_CONFIG.Timezone || Session.getScriptTimeZone() || 'UTC');
}

function formatDate_(date, tz) {
  return Utilities.formatDate(date, tz || DEFAULT_CONFIG.Timezone, 'yyyy-MM-dd');
}

function formatTime_(date, tz) {
  return Utilities.formatDate(date, tz || DEFAULT_CONFIG.Timezone, 'HH:mm');
}

function timestamp_(tz, date) {
  return Utilities.formatDate(date || new Date(), tz || DEFAULT_CONFIG.Timezone, 'yyyy-MM-dd HH:mm:ss');
}

function normalizeId_(value) {
  return String(value || '').trim().toUpperCase();
}

function normalizeName_(value) {
  return String(value || '').trim().toLowerCase().replace(/\s+/g, ' ');
}

function normalizeDeviceToken_(value) {
  return String(value || '').trim();
}

function normalizeDateString_(value, tz) {
  if (!value) return '';
  if (value instanceof Date) return formatDate_(value, tz || DEFAULT_CONFIG.Timezone);
  const text = String(value).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return text;
  const dash = text.match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/);
  if (dash) return dash[3] + '-' + dash[2].padStart(2, '0') + '-' + dash[1].padStart(2, '0');
  const slash = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (slash) return slash[3] + '-' + slash[2].padStart(2, '0') + '-' + slash[1].padStart(2, '0');
  const parsed = new Date(text);
  if (!isNaN(parsed.getTime())) return formatDate_(parsed, tz || DEFAULT_CONFIG.Timezone);
  return text;
}

function normalizeTimeString_(value, tz) {
  if (!value) return '';
  if (value instanceof Date) return Utilities.formatDate(value, tz || DEFAULT_CONFIG.Timezone, 'HH:mm');
  const text = String(value).trim();
  const match = text.match(/^(\d{1,2}):(\d{2})/);
  if (!match) return '';
  const h = Math.min(23, Math.max(0, Number(match[1])));
  const m = Math.min(59, Math.max(0, Number(match[2])));
  return String(h).padStart(2, '0') + ':' + String(m).padStart(2, '0');
}

function timeToMinutes_(time) {
  const normalized = normalizeTimeString_(time);
  if (!normalized) return 0;
  const parts = normalized.split(':').map(Number);
  return parts[0] * 60 + parts[1];
}

function sameDateValue_(value, date, tz) {
  return normalizeDateString_(value, tz) === normalizeDateString_(date, tz);
}

function error_(message) {
  return { ok: false, message: message };
}
