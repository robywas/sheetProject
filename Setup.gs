function setupWorkbook() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  ensureSheetWithHeader_(
    spreadsheet,
    SHEET_NAMES.PROCEDURES,
    HEADERS.PROCEDURES
  );
  ensureSheetWithHeader_(spreadsheet, SHEET_NAMES.CLIENTS, HEADERS.CLIENTS);
  ensureSheetWithHeader_(
    spreadsheet,
    SHEET_NAMES.EMPLOYEES,
    HEADERS.EMPLOYEES
  );
  ensureSheetWithHeader_(
    spreadsheet,
    SHEET_NAMES.CLIENT_PROCEDURES,
    HEADERS.CLIENT_PROCEDURES
  );
  ensureSheetWithHeader_(
    spreadsheet,
    SHEET_NAMES.ASSIGNMENTS,
    HEADERS.ASSIGNMENTS
  );
  ensureSheetWithHeader_(spreadsheet, SHEET_NAMES.TASKS, HEADERS.TASKS);
  ensureSheetWithHeader_(spreadsheet, SHEET_NAMES.MY_TASKS, HEADERS.MY_TASKS);

  const dashboardSheet = spreadsheet.getSheetByName(SHEET_NAMES.MANAGER_DASHBOARD);
  if (!dashboardSheet) {
    spreadsheet.insertSheet(SHEET_NAMES.MANAGER_DASHBOARD);
  }

  applyFormatting_();
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Struktura arkusza jest gotowa.',
    'Procedury',
    5
  );
}

function seedSampleData() {
  setupWorkbook();

  const proceduresSheet = getSheetOrThrow_(SHEET_NAMES.PROCEDURES);
  const clientsSheet = getSheetOrThrow_(SHEET_NAMES.CLIENTS);
  const employeesSheet = getSheetOrThrow_(SHEET_NAMES.EMPLOYEES);
  const clientProceduresSheet = getSheetOrThrow_(SHEET_NAMES.CLIENT_PROCEDURES);
  const assignmentsSheet = getSheetOrThrow_(SHEET_NAMES.ASSIGNMENTS);

  appendRowsIfOnlyHeader_(proceduresSheet, [
    ['P001', 'Kontrola cisnienia', 'Pomiar i zapis wyniku', 7, 2, true],
    ['P002', 'Kontrola glikemii', 'Pobranie i wpisanie wyniku', 3, 1, true],
    ['P003', 'Ocena rany', 'Kontrola i dokumentacja', 14, 3, true],
  ]);

  appendRowsIfOnlyHeader_(clientsSheet, [
    ['CL001', 'Jan Kowalski', true],
    ['CL002', 'Anna Nowak', true],
    ['CL003', 'Piotr Wisniewski', true],
    ['CL004', 'Maria Zielinska', true],
  ]);

  appendRowsIfOnlyHeader_(employeesSheet, [
    ['E001', 'Agnieszka Opiekun', 'agnieszka.opiekun@example.com', 'pracownik', true],
    ['E002', 'Tomasz Opiekun', 'tomasz.opiekun@example.com', 'pracownik', true],
    ['M001', 'Monika Manager', 'monika.manager@example.com', 'manager', true],
  ]);

  const today = normalizeDate_(new Date());
  const startDate = new Date(today.getTime());
  startDate.setDate(startDate.getDate() - 7);

  appendRowsIfOnlyHeader_(clientProceduresSheet, [
    ['CL001', 'P001', startDate, '', true],
    ['CL001', 'P002', startDate, '', true],
    ['CL002', 'P003', startDate, '', true],
    ['CL003', 'P001', startDate, 10, true],
    ['CL004', 'P002', startDate, '', true],
  ]);

  appendRowsIfOnlyHeader_(assignmentsSheet, [
    ['CL001', 'E001', startDate, '', true],
    ['CL002', 'E001', startDate, '', true],
    ['CL003', 'E002', startDate, '', true],
    ['CL004', 'E002', startDate, '', true],
  ]);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Dodano dane przykladowe.',
    'Procedury',
    5
  );
}

function ensureSheetWithHeader_(spreadsheet, sheetName, headers) {
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  }

  if (sheet.getMaxColumns() < headers.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
  }

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.getRange(1, 1, 1, headers.length).setBackground('#f1f3f4');
  sheet.setFrozenRows(1);
}

function applyFormatting_() {
  const taskSheet = getSheetOrThrow_(SHEET_NAMES.TASKS);
  taskSheet.getRange('E:E').setNumberFormat('yyyy-mm-dd');
  taskSheet.getRange('G:H').setNumberFormat('yyyy-mm-dd hh:mm');

  const clientProceduresSheet = getSheetOrThrow_(SHEET_NAMES.CLIENT_PROCEDURES);
  clientProceduresSheet.getRange('C:C').setNumberFormat('yyyy-mm-dd');

  const assignmentsSheet = getSheetOrThrow_(SHEET_NAMES.ASSIGNMENTS);
  assignmentsSheet.getRange('C:D').setNumberFormat('yyyy-mm-dd');

  const myTasksSheet = getSheetOrThrow_(SHEET_NAMES.MY_TASKS);
  myTasksSheet.getRange('C:C').setNumberFormat('yyyy-mm-dd');
}

function appendRowsIfOnlyHeader_(sheet, rows) {
  if (sheet.getLastRow() > 1) {
    return;
  }
  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}
