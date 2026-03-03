function setupWorkbook() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  migrateLegacySheetNames_(spreadsheet);

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
  applyDataHints_();
  applyDataValidation_();
  try {
    refreshManagerDashboard();
  } catch (error) {
    // Dashboard moze byc odswiezony pozniej z menu.
  }
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
    ['P001', 'Kontrola cisnienia', 'Pomiar i zapis wyniku', 10, 2, true],
    ['P002', 'Kontrola glikemii', 'Pobranie i wpisanie wyniku', SCHEDULE_LAST_DAY_TOKEN, 1, true],
    ['P003', 'Ocena rany', 'Kontrola i dokumentacja', 15, 3, true],
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
  startDate.setDate(1);

  appendRowsIfOnlyHeader_(clientProceduresSheet, [
    ['CL001', 'P001', startDate, true],
    ['CL001', 'P002', startDate, true],
    ['CL002', 'P003', startDate, true],
    ['CL003', 'P001', startDate, true],
    ['CL004', 'P002', startDate, true],
  ]);

  appendRowsIfOnlyHeader_(assignmentsSheet, [
    ['CL001', 'E001', startDate, '', true, 1],
    ['CL001', 'E002', startDate, '', true, 2],
    ['CL002', 'E002', startDate, '', true, 1],
    ['CL002', 'E001', startDate, '', true, 2],
    ['CL003', 'E001', startDate, '', true, 1],
    ['CL003', 'E002', startDate, '', true, 2],
    ['CL004', 'E002', startDate, '', true, 1],
    ['CL004', 'E001', startDate, '', true, 2],
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

function applyDataHints_() {
  const proceduresSheet = getSheetOrThrow_(SHEET_NAMES.PROCEDURES);
  proceduresSheet.getRange('D1').setNote(
    'Podaj dzien miesiaca: 1..31 lub "' + SCHEDULE_LAST_DAY_TOKEN + '".'
  );
}

function applyDataValidation_() {
  const proceduresSheet = getSheetOrThrow_(SHEET_NAMES.PROCEDURES);
  const clientsSheet = getSheetOrThrow_(SHEET_NAMES.CLIENTS);
  const employeesSheet = getSheetOrThrow_(SHEET_NAMES.EMPLOYEES);
  const clientProceduresSheet = getSheetOrThrow_(SHEET_NAMES.CLIENT_PROCEDURES);
  const assignmentsSheet = getSheetOrThrow_(SHEET_NAMES.ASSIGNMENTS);

  const procedureRows = proceduresSheet.getMaxRows() - 1;
  const clientRows = clientsSheet.getMaxRows() - 1;
  const employeeRows = employeesSheet.getMaxRows() - 1;
  const clientProcedureRows = clientProceduresSheet.getMaxRows() - 1;
  const assignmentRows = assignmentsSheet.getMaxRows() - 1;

  const monthDayOptions = [];
  for (let day = 1; day <= 31; day += 1) {
    monthDayOptions.push(String(day));
  }
  monthDayOptions.push(SCHEDULE_LAST_DAY_TOKEN);

  const monthDayRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(monthDayOptions, true)
    .setAllowInvalid(false)
    .setHelpText('Dozwolone: 1..31 lub OSTATNI.')
    .build();
  proceduresSheet.getRange(2, 4, procedureRows, 1).setDataValidation(monthDayRule);

  const formulaSep = getFormulaArgSeparator_();
  const nonNegativeIntegerFormula =
    '=OR(E2=""' +
    formulaSep +
    'AND(ISNUMBER(E2)' +
    formulaSep +
    'E2>=0' +
    formulaSep +
    'E2=INT(E2)))';
  const positiveIntegerFormula =
    '=OR(F2=""' +
    formulaSep +
    'AND(ISNUMBER(F2)' +
    formulaSep +
    'F2>=1' +
    formulaSep +
    'F2=INT(F2)))';

  const nonNegativeIntegerRule = SpreadsheetApp.newDataValidation()
    .requireFormulaSatisfied(nonNegativeIntegerFormula)
    .setAllowInvalid(false)
    .setHelpText('Podaj liczbe calkowita >= 0.')
    .build();
  proceduresSheet.getRange(2, 5, procedureRows, 1).setDataValidation(nonNegativeIntegerRule);

  const positiveIntegerRule = SpreadsheetApp.newDataValidation()
    .requireFormulaSatisfied(positiveIntegerFormula)
    .setAllowInvalid(false)
    .setHelpText('Podaj liczbe calkowita >= 1.')
    .build();
  assignmentsSheet.getRange(2, 6, assignmentRows, 1).setDataValidation(positiveIntegerRule);

  const clientIdRange = clientsSheet.getRange(2, 1, clientRows, 1);
  const procedureIdRange = proceduresSheet.getRange(2, 1, procedureRows, 1);
  const employeeIdRange = employeesSheet.getRange(2, 1, employeeRows, 1);

  const clientIdRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(clientIdRange, true)
    .setAllowInvalid(false)
    .setHelpText('Wybierz client_id z zakladki Klienci.')
    .build();
  const procedureIdRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(procedureIdRange, true)
    .setAllowInvalid(false)
    .setHelpText('Wybierz procedure_id z zakladki Procedury.')
    .build();
  const employeeIdRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(employeeIdRange, true)
    .setAllowInvalid(false)
    .setHelpText('Wybierz employee_id z zakladki Pracownicy.')
    .build();

  clientProceduresSheet.getRange(2, 1, clientProcedureRows, 1).setDataValidation(clientIdRule);
  clientProceduresSheet
    .getRange(2, 2, clientProcedureRows, 1)
    .setDataValidation(procedureIdRule);
  assignmentsSheet.getRange(2, 1, assignmentRows, 1).setDataValidation(clientIdRule);
  assignmentsSheet.getRange(2, 2, assignmentRows, 1).setDataValidation(employeeIdRule);

  proceduresSheet.getRange(2, 6, procedureRows, 1).insertCheckboxes();
  clientsSheet.getRange(2, 3, clientRows, 1).insertCheckboxes();
  clientProceduresSheet.getRange(2, 4, clientProcedureRows, 1).insertCheckboxes();
  assignmentsSheet.getRange(2, 5, assignmentRows, 1).insertCheckboxes();
}

function getFormulaArgSeparator_() {
  const candidates = [',', ';'];
  for (let i = 0; i < candidates.length; i += 1) {
    const sep = candidates[i];
    try {
      SpreadsheetApp.newDataValidation()
        .requireFormulaSatisfied('=AND(1=1' + sep + '1=1)')
        .build();
      return sep;
    } catch (error) {
      // Sprobuj kolejny separator.
    }
  }
  return ',';
}

function migrateLegacySheetNames_(spreadsheet) {
  const legacyToCurrent = [
    [LEGACY_SHEET_NAMES.CLIENTS, SHEET_NAMES.CLIENTS],
    [LEGACY_SHEET_NAMES.CLIENT_PROCEDURES, SHEET_NAMES.CLIENT_PROCEDURES],
  ];

  legacyToCurrent.forEach(([legacyName, currentName]) => {
    const currentSheet = spreadsheet.getSheetByName(currentName);
    const legacySheet = spreadsheet.getSheetByName(legacyName);
    if (!legacySheet) {
      return;
    }

    if (!currentSheet) {
      legacySheet.setName(currentName);
      return;
    }

    if (currentSheet.getLastRow() <= 1 && legacySheet.getLastRow() > 1) {
      const legacyValues = legacySheet.getDataRange().getValues();
      if (currentSheet.getMaxColumns() < legacyValues[0].length) {
        currentSheet.insertColumnsAfter(
          currentSheet.getMaxColumns(),
          legacyValues[0].length - currentSheet.getMaxColumns()
        );
      }
      currentSheet
        .getRange(1, 1, legacyValues.length, legacyValues[0].length)
        .setValues(legacyValues);
    }

    spreadsheet.deleteSheet(legacySheet);
  });
}

function appendRowsIfOnlyHeader_(sheet, rows) {
  if (sheet.getLastRow() > 1) {
    return;
  }
  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}
