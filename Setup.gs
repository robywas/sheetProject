function setupWorkbook() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  migrateLegacySheetNames_(spreadsheet);
  ensureSheetExists_(spreadsheet, SHEET_NAMES.PROCEDURES);
  ensureSheetExists_(spreadsheet, SHEET_NAMES.CLIENTS);
  ensureSheetExists_(spreadsheet, SHEET_NAMES.EMPLOYEES);
  ensureSheetExists_(spreadsheet, SHEET_NAMES.CLIENT_PROCEDURES);
  ensureSheetExists_(spreadsheet, SHEET_NAMES.ASSIGNMENTS);
  ensureSheetExists_(spreadsheet, SHEET_NAMES.TASKS);
  ensureSheetExists_(spreadsheet, SHEET_NAMES.MY_TASKS);
  ensureSheetExists_(spreadsheet, SHEET_NAMES.MANAGER_DASHBOARD);
  const dashboardSheet = getSheetOrThrow_(SHEET_NAMES.MANAGER_DASHBOARD);
  ensureSheetSize_(dashboardSheet, DASHBOARD_MIN_ROWS, 7);
  shrinkSheetToDataBuffer_(dashboardSheet, DASHBOARD_MIN_ROWS, 7);

  migrateIdBasedModelToNameModel_();

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

  applyFormatting_();
  applyDataHints_();
  applyDataValidation_();
  try {
    sortTasksByStatusAndDueDesc_();
  } catch (error) {
    // Sortowanie zadan nie powinno blokowac setupu.
  }
  try {
    refreshManagerDashboard();
  } catch (error) {
    // Dashboard moze byc odswiezony pozniej z menu.
  }
  try {
    refreshClientProceduresControl();
  } catch (error) {
    // Kontrola moze byc odswiezona pozniej z menu.
  }
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Struktura arkusza jest gotowa (build: 2026-03-05d).',
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
    [
      'Kontrola cisnienia',
      'Pomiar i zapis wyniku',
      10,
      2,
      SCHEDULE_MODE.MONTHLY,
      1,
    ],
    [
      'Kontrola glikemii',
      'Pobranie i wpisanie wyniku',
      SCHEDULE_LAST_DAY_TOKEN,
      1,
      SCHEDULE_MODE.MONTHLY,
      1,
    ],
    [
      'Ocena rany',
      'Kontrola i dokumentacja',
      15,
      3,
      SCHEDULE_MODE.MONTHLY,
      1,
    ],
  ]);

  appendRowsIfOnlyHeader_(clientsSheet, [
    ['Jan Kowalski'],
    ['Anna Nowak'],
    ['Piotr Wisniewski'],
    ['Maria Zielinska'],
  ]);

  appendRowsIfOnlyHeader_(employeesSheet, [
    ['Agnieszka Opiekun', 'agnieszka.opiekun@example.com', 'pracownik'],
    ['Tomasz Opiekun', 'tomasz.opiekun@example.com', 'pracownik'],
    ['Monika Manager', 'monika.manager@example.com', 'manager'],
  ]);

  const today = normalizeDate_(new Date());
  const startDate = new Date(today.getTime());
  startDate.setDate(1);

  appendRowsIfOnlyHeader_(clientProceduresSheet, [
    ['Jan Kowalski', 'Kontrola cisnienia', startDate, 'Mierzyc po 10 minutach odpoczynku.'],
    ['Jan Kowalski', 'Kontrola glikemii', startDate, 'Pomiar rano, na czczo.'],
    ['Anna Nowak', 'Ocena rany', startDate, 'Dokumentacja zdjeciowa przy zmianie opatrunku.'],
    ['Piotr Wisniewski', 'Kontrola cisnienia', startDate, 'Uwzglednic druga reke przy odchyleniach.'],
    ['Maria Zielinska', 'Kontrola glikemii', startDate, 'W razie objawow zrobic dodatkowy pomiar.'],
  ]);

  appendRowsIfOnlyHeader_(assignmentsSheet, [
    ['Jan Kowalski', 'Agnieszka Opiekun', startDate, '', 1],
    ['Jan Kowalski', 'Tomasz Opiekun', startDate, '', 2],
    ['Anna Nowak', 'Tomasz Opiekun', startDate, '', 1],
    ['Anna Nowak', 'Agnieszka Opiekun', startDate, '', 2],
    ['Piotr Wisniewski', 'Agnieszka Opiekun', startDate, '', 1],
    ['Piotr Wisniewski', 'Tomasz Opiekun', startDate, '', 2],
    ['Maria Zielinska', 'Tomasz Opiekun', startDate, '', 1],
    ['Maria Zielinska', 'Agnieszka Opiekun', startDate, '', 2],
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

  ensureSheetSize_(sheet, getDefaultMinRowsForSheet_(sheetName), headers.length);

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.getRange(1, 1, 1, headers.length).setBackground('#f1f3f4');
  try {
    const maxRows = Math.max(1, toNumber_(sheet.getMaxRows(), 1));
    sheet.setFrozenRows(Math.min(1, maxRows));
  } catch (error) {
    // Zamrozenie naglowka jest opcjonalne i nie moze blokowac setupu.
  }

  shrinkSheetToDataBuffer_(sheet, getDefaultMinRowsForSheet_(sheetName), headers.length);
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
  myTasksSheet.getRange('B:B').setNumberFormat('yyyy-mm-dd');
}

function applyDataHints_() {
  const proceduresSheet = getSheetOrThrow_(SHEET_NAMES.PROCEDURES);
  const clientProceduresSheet = getSheetOrThrow_(SHEET_NAMES.CLIENT_PROCEDURES);
  const assignmentsSheet = getSheetOrThrow_(SHEET_NAMES.ASSIGNMENTS);
  const tasksSheet = getSheetOrThrow_(SHEET_NAMES.TASKS);
  proceduresSheet.getRange('C1').setNote(
    'Dla trybu miesiecznego podaj dzien: 1..31 lub "' + SCHEDULE_LAST_DAY_TOKEN + '".'
  );
  proceduresSheet.getRange('E1').setNote(
    'Tryb harmonogramu: ' +
      SCHEDULE_MODE.MONTHLY +
      ' (do dnia miesiaca) lub ' +
      SCHEDULE_MODE.DAILY +
      ' (co N dni).'
  );
  proceduresSheet.getRange('F1').setNote(
    'Interwal: dla trybu miesiecznego = co ile miesiecy, dla dziennego = co ile dni.'
  );
  clientProceduresSheet
    .getRange('D1')
    .setNote('Uwagi do konkretnego powiazania klient-procedura. Pokazywane w Moje_zadania.');
  clientProceduresSheet
    .getRange('E1')
    .setNote('Kontrola: OK / Nieprzypisane / Brak zadan. Odswiez z menu Procedury.');
  assignmentsSheet
    .getRange('B1')
    .setNote('Pusty pracownik = automatyczna rotacja miedzy wszystkimi pracownikami.');
  tasksSheet
    .getRange('D1')
    .setNote('Wybierz pracownika z listy (slownik z arkusza Pracownicy).');
}

function applyDataValidation_() {
  const proceduresSheet = getSheetOrThrow_(SHEET_NAMES.PROCEDURES);
  const clientsSheet = getSheetOrThrow_(SHEET_NAMES.CLIENTS);
  const employeesSheet = getSheetOrThrow_(SHEET_NAMES.EMPLOYEES);
  const clientProceduresSheet = getSheetOrThrow_(SHEET_NAMES.CLIENT_PROCEDURES);
  const assignmentsSheet = getSheetOrThrow_(SHEET_NAMES.ASSIGNMENTS);
  const tasksSheet = getSheetOrThrow_(SHEET_NAMES.TASKS);

  const procedureRows = proceduresSheet.getMaxRows() - 1;
  const clientRows = clientsSheet.getMaxRows() - 1;
  const employeeRows = employeesSheet.getMaxRows() - 1;
  const clientProcedureRows = clientProceduresSheet.getMaxRows() - 1;
  const assignmentRows = assignmentsSheet.getMaxRows() - 1;
  const taskRows = tasksSheet.getMaxRows() - 1;

  const monthDayOptions = [''];
  for (let day = 1; day <= 31; day += 1) {
    monthDayOptions.push(String(day));
  }
  monthDayOptions.push(SCHEDULE_LAST_DAY_TOKEN);

  const monthDayRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(monthDayOptions, true)
    .setAllowInvalid(false)
    .setHelpText('Dla trybu miesiecznego: 1..31 lub OSTATNI. Dla dziennego moze byc puste.')
    .build();
  proceduresSheet.getRange(2, 3, procedureRows, 1).setDataValidation(monthDayRule);

  const nonNegativeIntegerRule = SpreadsheetApp.newDataValidation()
    .requireNumberGreaterThanOrEqualTo(0)
    .setAllowInvalid(false)
    .setHelpText('Podaj liczbe >= 0.')
    .build();
  proceduresSheet.getRange(2, 4, procedureRows, 1).setDataValidation(nonNegativeIntegerRule);

  const scheduleModeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList([SCHEDULE_MODE.MONTHLY, SCHEDULE_MODE.DAILY], true)
    .setAllowInvalid(false)
    .setHelpText('Dozwolone: MIESIECZNY albo DZIENNY.')
    .build();
  proceduresSheet.getRange(2, 5, procedureRows, 1).setDataValidation(scheduleModeRule);

  const positiveIntegerRule = SpreadsheetApp.newDataValidation()
    .requireNumberGreaterThanOrEqualTo(1)
    .setAllowInvalid(false)
    .setHelpText('Podaj liczbe >= 1.')
    .build();
  proceduresSheet.getRange(2, 6, procedureRows, 1).setDataValidation(positiveIntegerRule);
  assignmentsSheet.getRange(2, 5, assignmentRows, 1).setDataValidation(positiveIntegerRule);

  const clientNameRange = clientsSheet.getRange(2, 1, clientRows, 1);
  const procedureNameRange = proceduresSheet.getRange(2, 1, procedureRows, 1);
  const employeeNameRange = employeesSheet.getRange(2, 1, employeeRows, 1);

  const clientNameRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(clientNameRange, true)
    .setAllowInvalid(false)
    .setHelpText('Wybierz klienta z zakladki Klienci.')
    .build();
  const procedureNameRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(procedureNameRange, true)
    .setAllowInvalid(false)
    .setHelpText('Wybierz procedure z zakladki Procedury.')
    .build();
  const employeeOptions = [''];
  const employeeSeen = {};
  employeeNameRange.getValues().forEach((row) => {
    const employeeName = normalizeText_(row[0]);
    if (!employeeName) {
      return;
    }
    const key = normalizeLookupKey_(employeeName);
    if (employeeSeen[key]) {
      return;
    }
    employeeSeen[key] = true;
    employeeOptions.push(employeeName);
  });
  const optionalEmployeeNameRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(employeeOptions, true)
    .setAllowInvalid(false)
    .setHelpText(
      'Wybierz pracownika albo zostaw puste (rotacja wszystkich pracownikow).'
    )
    .build();

  clientProceduresSheet.getRange(2, 1, clientProcedureRows, 1).setDataValidation(clientNameRule);
  clientProceduresSheet
    .getRange(2, 2, clientProcedureRows, 1)
    .setDataValidation(procedureNameRule);
  assignmentsSheet.getRange(2, 1, assignmentRows, 1).setDataValidation(clientNameRule);
  assignmentsSheet
    .getRange(2, 2, assignmentRows, 1)
    .setDataValidation(optionalEmployeeNameRule);
  tasksSheet.getRange(2, 4, taskRows, 1).setDataValidation(optionalEmployeeNameRule);

}

function migrateIdBasedModelToNameModel_() {
  const proceduresSnapshot = readSheetSnapshot_(SHEET_NAMES.PROCEDURES);
  const clientsSnapshot = readSheetSnapshot_(SHEET_NAMES.CLIENTS);
  const employeesSnapshot = readSheetSnapshot_(SHEET_NAMES.EMPLOYEES);
  const clientProceduresSnapshot = readSheetSnapshot_(SHEET_NAMES.CLIENT_PROCEDURES);
  const assignmentsSnapshot = readSheetSnapshot_(SHEET_NAMES.ASSIGNMENTS);
  const tasksSnapshot = readSheetSnapshot_(SHEET_NAMES.TASKS);

  const clientIdToName = buildIdToNameMap_(clientsSnapshot, 'client_id', 'klient');
  const procedureIdToName = buildIdToNameMap_(proceduresSnapshot, 'procedure_id', 'procedura');
  const employeeIdToName = buildIdToNameMap_(employeesSnapshot, 'employee_id', 'pracownik');

  const migratedProcedures = proceduresSnapshot.rows
    .map((row) => {
      const isActive = toBoolean_(
        getNamedValue_(row, proceduresSnapshot.indices, 'aktywna', '', true),
        true
      );
      if (!isActive) {
        return null;
      }
      return [
        getNamedValue_(row, proceduresSnapshot.indices, 'procedura', 'procedure_id'),
        getNamedValue_(row, proceduresSnapshot.indices, 'opis'),
        getNamedValue_(row, proceduresSnapshot.indices, 'dzien_miesiaca', '', 1),
        getNamedValue_(row, proceduresSnapshot.indices, 'dni_ostrzezenia'),
        normalizeText_(
          getNamedValue_(
            row,
            proceduresSnapshot.indices,
            'tryb_harmonogramu',
            '',
            SCHEDULE_MODE.MONTHLY
          )
        ).toUpperCase() || SCHEDULE_MODE.MONTHLY,
        Math.max(
          1,
          toNumber_(
            getNamedValue_(row, proceduresSnapshot.indices, 'interwal', '', 1),
            1
          )
        ),
      ];
    })
    .filter(Boolean)
    .filter((row) => normalizeText_(row[0]));

  const migratedClients = clientsSnapshot.rows
    .map((row) => {
      const isActive = toBoolean_(
        getNamedValue_(row, clientsSnapshot.indices, 'aktywny', '', true),
        true
      );
      if (!isActive) {
        return null;
      }
      return [getNamedValue_(row, clientsSnapshot.indices, 'klient', 'client_id')];
    })
    .filter(Boolean)
    .filter((row) => normalizeText_(row[0]));

  const migratedEmployees = employeesSnapshot.rows
    .map((row) => {
      const isActive = toBoolean_(
        getNamedValue_(row, employeesSnapshot.indices, 'aktywny', '', true),
        true
      );
      if (!isActive) {
        return null;
      }
      return [
        getNamedValue_(row, employeesSnapshot.indices, 'pracownik', 'employee_id'),
        getNamedValue_(row, employeesSnapshot.indices, 'email'),
        getNamedValue_(row, employeesSnapshot.indices, 'rola'),
      ];
    })
    .filter(Boolean)
    .filter((row) => normalizeText_(row[0]));

  const migratedClientProcedures = clientProceduresSnapshot.rows
    .map((row) => {
      const isActive = toBoolean_(
        getNamedValue_(row, clientProceduresSnapshot.indices, 'aktywna', '', true),
        true
      );
      if (!isActive) {
        return null;
      }
      return [
        getResolvedNameValue_(
          row,
          clientProceduresSnapshot.indices,
          'klient',
          'client_id',
          clientIdToName
        ),
        getResolvedNameValue_(
          row,
          clientProceduresSnapshot.indices,
          'procedura',
          'procedure_id',
          procedureIdToName
        ),
        getNamedValue_(row, clientProceduresSnapshot.indices, 'data_start'),
        normalizeText_(
          getNamedValue_(row, clientProceduresSnapshot.indices, 'uwagi', 'notatki')
        ) || normalizeText_(getNamedValue_(row, clientProceduresSnapshot.indices, 'notes')),
        '', // kontrola – uzupelni refreshClientProceduresControl
      ];
    })
    .filter(Boolean)
    .filter((row) => normalizeText_(row[0]) && normalizeText_(row[1]));

  const migratedAssignments = assignmentsSnapshot.rows
    .map((row) => {
      const isActive = toBoolean_(
        getNamedValue_(row, assignmentsSnapshot.indices, 'aktywna', '', true),
        true
      );
      if (!isActive) {
        return null;
      }

      return [
        getResolvedNameValue_(
          row,
          assignmentsSnapshot.indices,
          'klient',
          'client_id',
          clientIdToName
        ),
        getResolvedNameValue_(
          row,
          assignmentsSnapshot.indices,
          'pracownik',
          'employee_id',
          employeeIdToName
        ),
        getNamedValue_(row, assignmentsSnapshot.indices, 'data_od'),
        getNamedValue_(row, assignmentsSnapshot.indices, 'data_do'),
        getNamedValue_(row, assignmentsSnapshot.indices, 'kolejnosc', '', 1),
      ];
    })
    .filter(Boolean)
    .filter((row) => normalizeText_(row[0]));

  const migratedTasks = tasksSnapshot.rows
    .map((row) => {
      const taskId =
        getNamedValue_(row, tasksSnapshot.indices, 'task_id') || Utilities.getUuid();
      const clientName = getResolvedNameValue_(
        row,
        tasksSnapshot.indices,
        'klient',
        'client_id',
        clientIdToName
      );
      const procedureName = getResolvedNameValue_(
        row,
        tasksSnapshot.indices,
        'procedura',
        'procedure_id',
        procedureIdToName
      );
      const employeeName = getResolvedNameValue_(
        row,
        tasksSnapshot.indices,
        'pracownik',
        'employee_id',
        employeeIdToName
      );
      const dueDate = getNamedValue_(row, tasksSnapshot.indices, 'due_date');
      const taskKey =
        normalizeText_(getNamedValue_(row, tasksSnapshot.indices, 'task_key')) ||
        (clientName && procedureName && dueDate
          ? buildTaskKey_(clientName, procedureName, dueDate)
          : '');

      return [
        taskId,
        clientName,
        procedureName,
        employeeName,
        dueDate,
        getNamedValue_(row, tasksSnapshot.indices, 'status', '', STATUS.NEW),
        getNamedValue_(row, tasksSnapshot.indices, 'created_at'),
        getNamedValue_(row, tasksSnapshot.indices, 'completed_at'),
        getNamedValue_(row, tasksSnapshot.indices, 'notes'),
        taskKey,
        getNamedValue_(row, tasksSnapshot.indices, 'dni_ostrzezenia', '', 0),
      ];
    })
    .filter((row) => normalizeText_(row[1]) && normalizeText_(row[2]));

  writeMigratedSheet_(proceduresSnapshot.sheet, HEADERS.PROCEDURES, migratedProcedures);
  writeMigratedSheet_(clientsSnapshot.sheet, HEADERS.CLIENTS, migratedClients);
  writeMigratedSheet_(employeesSnapshot.sheet, HEADERS.EMPLOYEES, migratedEmployees);
  writeMigratedSheet_(
    clientProceduresSnapshot.sheet,
    HEADERS.CLIENT_PROCEDURES,
    migratedClientProcedures
  );
  writeMigratedSheet_(assignmentsSnapshot.sheet, HEADERS.ASSIGNMENTS, migratedAssignments);
  writeMigratedSheet_(tasksSnapshot.sheet, HEADERS.TASKS, migratedTasks);
}

function ensureSheetExists_(spreadsheet, sheetName) {
  if (spreadsheet.getSheetByName(sheetName)) {
    return;
  }
  spreadsheet.insertSheet(sheetName);
}

function readSheetSnapshot_(sheetName) {
  const sheet = getSheetOrThrow_(sheetName);
  const values = sheet.getDataRange().getValues();
  const headers = values.length ? values[0].map((cell) => normalizeText_(cell)) : [];
  const indices = {};
  headers.forEach((header, idx) => {
    if (!header) {
      return;
    }
    indices[header] = idx;
  });

  const rows = values
    .slice(1)
    .filter((row) => row.some((cell) => cell !== '' && cell !== null));
  return { sheet, headers, indices, rows };
}

function buildIdToNameMap_(snapshot, idHeader, nameHeader) {
  const map = {};
  snapshot.rows.forEach((row) => {
    const idValue = normalizeText_(getValueByHeader_(row, snapshot.indices, idHeader));
    const nameValue = normalizeText_(getValueByHeader_(row, snapshot.indices, nameHeader));
    if (idValue) {
      map[idValue] = nameValue || idValue;
    }
    if (nameValue) {
      map[nameValue] = nameValue;
    }
  });
  return map;
}

function getNamedValue_(row, indices, primaryHeader, secondaryHeader, fallback) {
  const primaryValue = getValueByHeader_(row, indices, primaryHeader);
  if (primaryValue !== '' && primaryValue !== null && typeof primaryValue !== 'undefined') {
    return primaryValue;
  }
  if (secondaryHeader) {
    const secondaryValue = getValueByHeader_(row, indices, secondaryHeader);
    if (
      secondaryValue !== '' &&
      secondaryValue !== null &&
      typeof secondaryValue !== 'undefined'
    ) {
      return secondaryValue;
    }
  }
  return typeof fallback === 'undefined' ? '' : fallback;
}

function getResolvedNameValue_(row, indices, primaryHeader, legacyIdHeader, idToNameMap) {
  const directValue = normalizeText_(getValueByHeader_(row, indices, primaryHeader));
  if (directValue) {
    return directValue;
  }
  const legacyIdValue = normalizeText_(getValueByHeader_(row, indices, legacyIdHeader));
  if (!legacyIdValue) {
    return '';
  }
  return idToNameMap[legacyIdValue] || legacyIdValue;
}

function getValueByHeader_(row, indices, headerName) {
  const idx = indices[headerName];
  if (typeof idx === 'undefined') {
    return '';
  }
  return row[idx];
}

function writeMigratedSheet_(sheet, headers, rows) {
  ensureSheetSize_(sheet, Math.max(DEFAULT_SHEET_MIN_ROWS, rows.length + 1), headers.length);
  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
  try {
    shrinkSheetToDataBuffer_(
      sheet,
      getDefaultMinRowsForSheet_(sheet.getName()),
      headers.length
    );
  } catch (error) {
    // Migracja nie moze zatrzymac setupu przez problem z redukcja rozmiaru.
  }
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
  ensureSheetSize_(sheet, rows.length + 1, rows[0].length);
  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

function getDefaultMinRowsForSheet_(sheetName) {
  if (sheetName === SHEET_NAMES.MANAGER_DASHBOARD) {
    return DASHBOARD_MIN_ROWS;
  }
  return DEFAULT_SHEET_MIN_ROWS;
}

function shrinkSheetToDataBuffer_(sheet, minRows, minColumns) {
  // Tymczasowo nie zmniejszamy arkuszy podczas setupu/migracji.
  // U niektorych kont Google Sheets deleteRows/deleteColumns potrafi
  // rzucac niestabilny blad "Podaj liczbe >= 0." mimo poprawnych argumentow.
  // Arkusze i tak sa rozszerzane dynamicznie przez ensureSheetSize_.
  if (!sheet || !minRows || !minColumns) {
    return;
  }
}
