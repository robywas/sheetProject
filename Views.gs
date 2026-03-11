function refreshMyTasksView() {
  const employee = resolveCurrentEmployee_();
  if (!employee) {
    throw new Error(
      'Nie znaleziono aktywnego pracownika dla Twojego emaila. Uzuplnij arkusz Pracownicy.'
    );
  }
  refreshMyTasksViewForEmployeeName_(employee.employeeName);
}

function refreshMyTasksViewForEmployeeName_(employeeName) {
  const selectedEmployeeName = normalizeText_(employeeName);
  if (!selectedEmployeeName) {
    return;
  }
  const relationNotesByKey = buildClientProcedureNotesByKey_(
    getObjectRows_(SHEET_NAMES.CLIENT_PROCEDURES)
  );
  const taskRows = getObjectRows_(SHEET_NAMES.TASKS);
  const openTasks = taskRows
    .filter(
      (row) =>
        normalizeLookupKey_(row.pracownik || row.employee_id) ===
        normalizeLookupKey_(selectedEmployeeName)
    )
    .filter((row) => normalizeText_(row.status) !== STATUS.DONE)
    .map((row) => {
      const clientName = normalizeText_(row.klient || row.client_id);
      const procedureName = normalizeText_(row.procedura || row.procedure_id);
      return {
        taskId: normalizeText_(row.task_id),
        clientName,
        procedureName,
        dueDate: toDate_(row.due_date),
        status: normalizeText_(row.status) || STATUS.NEW,
        note: row.notes || '',
        relationNote:
          relationNotesByKey[
            buildClientProcedureRelationKey_(clientName, procedureName)
          ] || '',
      };
    })
    .filter((task) => task.taskId && task.dueDate)
    .sort((left, right) => left.dueDate - right.dueDate);

  const myTasksSheet = getSheetOrThrow_(SHEET_NAMES.MY_TASKS);
  myTasksSheet.getRange(1, 1, 1, HEADERS.MY_TASKS.length).setValues([HEADERS.MY_TASKS]);
  clearSheetBody_(myTasksSheet, HEADERS.MY_TASKS.length);

  if (openTasks.length === 0) {
    myTasksSheet.getRange(2, 1).setValue('Brak otwartych zadan.');
    return;
  }

  const rows = openTasks.map((task) => [
    task.taskId,
    task.dueDate,
    task.clientName,
    task.procedureName,
    task.status,
    task.relationNote,
    task.note,
  ]);

  ensureSheetSize_(myTasksSheet, rows.length + 1, HEADERS.MY_TASKS.length);
  myTasksSheet
    .getRange(2, 1, rows.length, HEADERS.MY_TASKS.length)
    .setValues(rows);

  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList([STATUS.NEW, STATUS.IN_PROGRESS, STATUS.DONE], true)
    .setAllowInvalid(false)
    .build();
  myTasksSheet
    .getRange(2, MY_TASKS_COL.STATUS, rows.length, 1)
    .setDataValidation(statusRule);
  myTasksSheet.getRange(2, MY_TASKS_COL.DUE_DATE, rows.length, 1).setNumberFormat('yyyy-mm-dd');

  const today = normalizeDate_(new Date());
  rows.forEach((row, idx) => {
    const dueDate = toDate_(row[MY_TASKS_COL.DUE_DATE - 1]);
    if (dueDate && dueDate < today) {
      myTasksSheet
        .getRange(idx + 2, 1, 1, HEADERS.MY_TASKS.length)
        .setBackground('#fde7e9');
    }
  });

}

function refreshClientProceduresControl() {
  const sheet = getSheetOrThrow_(SHEET_NAMES.CLIENT_PROCEDURES);
  const relationRows = getObjectRows_(SHEET_NAMES.CLIENT_PROCEDURES).filter(
    (row) =>
      normalizeText_(row.klient || row.client_id) &&
      normalizeText_(row.procedura || row.procedure_id)
  );
  if (relationRows.length === 0) {
    return;
  }

  const today = normalizeDate_(new Date());
  const horizon = new Date(today.getTime());
  horizon.setDate(horizon.getDate() + DEFAULT_GENERATION_DAYS);

  const allTasks = getObjectRows_(SHEET_NAMES.TASKS).map((row) => ({
    klient: normalizeText_(row.klient || row.client_id),
    procedura: normalizeText_(row.procedura || row.procedure_id),
    due_date: toDate_(row.due_date),
    pracownik: normalizeText_(row.pracownik || row.employee_id),
  }));

  const statuses = relationRows.map((row) => {
    const clientKey = normalizeLookupKey_(row.klient || row.client_id);
    const procedureKey = normalizeLookupKey_(row.procedura || row.procedure_id);
    const tasksInWindow = allTasks.filter(
      (t) =>
        normalizeLookupKey_(t.klient) === clientKey &&
        normalizeLookupKey_(t.procedura) === procedureKey &&
        t.due_date &&
        t.due_date >= today &&
        t.due_date <= horizon
    );
    if (tasksInWindow.length === 0) {
      return ['Brak zadań'];
    }
    const unassigned = tasksInWindow.some((t) => !t.pracownik);
    return [unassigned ? 'Nieprzypisane' : 'OK'];
  });

  const startRow = 2;
  ensureSheetSize_(sheet, startRow + statuses.length - 1, HEADERS.CLIENT_PROCEDURES.length);
  sheet.getRange(startRow, 5, statuses.length, 1).setValues(statuses);
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Kontrola Klienci_Procedury zaktualizowana.',
    'Procedury',
    3
  );
}

function buildClientProcedureRelationKey_(clientName, procedureName) {
  return (
    normalizeLookupKey_(clientName) + '|' + normalizeLookupKey_(procedureName)
  );
}

function buildClientProcedureNotesByKey_(clientProcedureRows) {
  const map = {};
  clientProcedureRows.forEach((row) => {
    const clientName = normalizeText_(row.klient || row.client_id);
    const procedureName = normalizeText_(row.procedura || row.procedure_id);
    if (!clientName || !procedureName) {
      return;
    }

    const key = buildClientProcedureRelationKey_(clientName, procedureName);
    const note =
      normalizeText_(row.uwagi || row.notatki || row.notes || row.note) || '';
    if (
      note ||
      !Object.prototype.hasOwnProperty.call(map, key)
    ) {
      map[key] = note;
    }
  });
  return map;
}

function refreshManagerDashboard() {
  const dashboardSheet = getSheetOrThrow_(SHEET_NAMES.MANAGER_DASHBOARD);
  const previousFilters = readManagerFilters_(dashboardSheet);

  const clientMasterRows = getObjectRows_(SHEET_NAMES.CLIENTS);
  const employeeRows = getObjectRows_(SHEET_NAMES.EMPLOYEES).filter((row) =>
    normalizeText_(row.pracownik || row.employee_id)
  );
  const employeeLookups = buildManagerEmployeeLookups_(employeeRows);
  prepareManagerDashboardLayout_(
    dashboardSheet,
    employeeLookups.names,
    previousFilters
  );

  const filters = readManagerFilters_(dashboardSheet);
  const selectedEmployeeName =
    filters.employeeName === MANAGER_FILTER.ALL_EMPLOYEES
      ? ''
      : employeeLookups.byName[filters.employeeName] || filters.employeeName;

  const tasks = getObjectRows_(SHEET_NAMES.TASKS).map((row) => {
    const dueDate = toDate_(row.due_date);
    const completedAt = toDate_(row.completed_at);
    return {
      taskId: normalizeText_(row.task_id),
      clientName: normalizeText_(row.klient || row.client_id),
      procedureName: normalizeText_(row.procedura || row.procedure_id),
      employeeName: normalizeText_(row.pracownik || row.employee_id),
      dueDate,
      status: normalizeText_(row.status) || STATUS.NEW,
      completedAt,
    };
  });
  const allClientNames = getAllClientNames_(clientMasterRows, tasks);

  const tasksInScope = tasks
    .filter((task) =>
      selectedEmployeeName
        ? normalizeLookupKey_(task.employeeName) ===
          normalizeLookupKey_(selectedEmployeeName)
        : true
    )
    .filter((task) => matchesManagerStatusFilter_(task, filters.status));

  const today = normalizeDate_(new Date());
  const dueSoonThreshold = new Date(today.getTime());
  dueSoonThreshold.setDate(dueSoonThreshold.getDate() + filters.horizonDays);
  const riskThreshold = new Date(today.getTime());
  riskThreshold.setDate(riskThreshold.getDate() + filters.riskDays);
  const last30Days = new Date(today.getTime());
  last30Days.setDate(last30Days.getDate() - 30);

  const openTasks = tasksInScope.filter((task) => task.status !== STATUS.DONE);
  const overdueTasks = openTasks.filter((task) => task.dueDate && task.dueDate < today);
  const dueSoonTasks = openTasks.filter(
    (task) => task.dueDate && task.dueDate >= today && task.dueDate <= dueSoonThreshold
  );
  const completedLast30Days = tasksInScope.filter(
    (task) => task.status === STATUS.DONE && task.completedAt && task.completedAt >= last30Days
  );
  const completionRate = tasksInScope.length
    ? Math.round((completedLast30Days.length / tasksInScope.length) * 100)
    : 0;

  ensureSheetSize_(
    dashboardSheet,
    Math.max(
      DASHBOARD_MIN_ROWS,
      tasksInScope.length * 2 + 80,
      allClientNames.length + 120
    ),
    7
  );

  dashboardSheet.getRange('A2').setValue(
    'Aktualizacja: ' +
      Utilities.formatDate(
        new Date(),
        Session.getScriptTimeZone(),
        'yyyy-MM-dd HH:mm:ss'
      )
  );
  dashboardSheet.getRange('A3').setValue(
    'Aktywne filtry: status=' +
      filters.status +
      ', pracownik=' +
      filters.employeeName +
      ', horyzont=' +
      filters.horizonDays +
      ' dni, prog zagrozenia=' +
      filters.riskDays +
      ' dni'
  );

  const kpiTable = [
    ['Wskaznik', 'Wartosc'],
    ['Otwarte zadania', openTasks.length],
    ['Przeterminowane', overdueTasks.length],
    ['Termin <= ' + filters.horizonDays + ' dni', dueSoonTasks.length],
    ['Ukonczone (30 dni)', completedLast30Days.length],
    ['Wskaznik realizacji', completionRate + '%'],
  ];

  const kpiStartRow = 10;
  dashboardSheet.getRange(kpiStartRow, 1).setValue('KPI');
  dashboardSheet.getRange(kpiStartRow, 1).setFontWeight('bold');
  dashboardSheet
    .getRange(kpiStartRow + 1, 1, kpiTable.length, 2)
    .setValues(kpiTable);
  dashboardSheet
    .getRange(kpiStartRow + 1, 1, 1, 2)
    .setFontWeight('bold')
    .setBackground('#f1f3f4');

  const riskTasks = openTasks
    .filter((task) => task.dueDate && task.dueDate <= riskThreshold)
    .sort((left, right) => left.dueDate - right.dueDate)
    .map((task) => {
      const daysDiff = Math.floor((task.dueDate - today) / ONE_DAY_MS);
      return [
        task.taskId,
        task.dueDate,
        task.clientName || '(brak klienta)',
        task.procedureName || '(brak procedury)',
        task.employeeName || '(nieprzypisane)',
        task.status,
        daysDiff,
      ];
    });

  const riskStartRow = kpiStartRow + kpiTable.length + 4;
  const riskHeaders = [
    'task_id',
    'termin',
    'klient',
    'procedura',
    'pracownik',
    'status',
    'dni_do_terminu',
  ];

  dashboardSheet.getRange(riskStartRow, 1).setValue('Zagrozone terminy');
  dashboardSheet.getRange(riskStartRow, 1).setFontWeight('bold');
  dashboardSheet
    .getRange(riskStartRow + 1, 1, 1, riskHeaders.length)
    .setValues([riskHeaders])
    .setFontWeight('bold')
    .setBackground('#f1f3f4');

  if (riskTasks.length > 0) {
    dashboardSheet
      .getRange(riskStartRow + 2, 1, riskTasks.length, riskHeaders.length)
      .setValues(riskTasks);
    dashboardSheet
      .getRange(riskStartRow + 2, 2, riskTasks.length, 1)
      .setNumberFormat('yyyy-mm-dd');
  } else {
    dashboardSheet.getRange(riskStartRow + 2, 1).setValue('Brak zadan zagrozonych.');
  }

  const loadStartRow = riskStartRow + Math.max(riskTasks.length, 1) + 5;
  dashboardSheet.getRange(loadStartRow, 1).setValue('Obciazenie pracownikow');
  dashboardSheet.getRange(loadStartRow, 1).setFontWeight('bold');

  const employeeLoad = {};
  openTasks.forEach((task) => {
    const key = task.employeeName || 'NIEPRZYPISANE';
    if (!employeeLoad[key]) {
      employeeLoad[key] = { open: 0, overdue: 0, dueSoon: 0 };
    }
    employeeLoad[key].open += 1;
    if (task.dueDate && task.dueDate < today) {
      employeeLoad[key].overdue += 1;
    }
    if (task.dueDate && task.dueDate >= today && task.dueDate <= dueSoonThreshold) {
      employeeLoad[key].dueSoon += 1;
    }
  });

  const loadTableHeaders = [
    'pracownik',
    'otwarte',
    'przeterminowane',
    'termin <= ' + filters.horizonDays + ' dni',
  ];
  dashboardSheet
    .getRange(loadStartRow + 1, 1, 1, loadTableHeaders.length)
    .setValues([loadTableHeaders])
    .setFontWeight('bold')
    .setBackground('#f1f3f4');

  const loadRows = Object.keys(employeeLoad)
    .sort()
    .map((employeeName) => [
      employeeName,
      employeeLoad[employeeName].open,
      employeeLoad[employeeName].overdue,
      employeeLoad[employeeName].dueSoon,
    ]);

  if (loadRows.length > 0) {
    dashboardSheet
      .getRange(loadStartRow + 2, 1, loadRows.length, loadTableHeaders.length)
      .setValues(loadRows);
  } else {
    dashboardSheet.getRange(loadStartRow + 2, 1).setValue('Brak danych o obciazeniu.');
  }

  const clientLoadStartRow = loadStartRow + Math.max(loadRows.length, 1) + 5;
  dashboardSheet.getRange(clientLoadStartRow, 1).setValue('Podsumowanie klientow');
  dashboardSheet.getRange(clientLoadStartRow, 1).setFontWeight('bold');

  const clientLoadHeaders = ['klient', 'otwarte', 'przeterminowane', 'termin <= horyzont'];
  dashboardSheet
    .getRange(clientLoadStartRow + 1, 1, 1, clientLoadHeaders.length)
    .setValues([clientLoadHeaders])
    .setFontWeight('bold')
    .setBackground('#f1f3f4');

  const clientLoadMap = {};
  openTasks.forEach((task) => {
    const key = task.clientName || 'NIEPRZYPISANY';
    if (!clientLoadMap[key]) {
      clientLoadMap[key] = { open: 0, overdue: 0, dueSoon: 0 };
    }
    clientLoadMap[key].open += 1;
    if (task.dueDate && task.dueDate < today) {
      clientLoadMap[key].overdue += 1;
    }
    if (task.dueDate && task.dueDate >= today && task.dueDate <= dueSoonThreshold) {
      clientLoadMap[key].dueSoon += 1;
    }
  });

  const clientLoadRows = Object.keys(clientLoadMap)
    .sort()
    .map((clientName) => [
      clientName,
      clientLoadMap[clientName].open,
      clientLoadMap[clientName].overdue,
      clientLoadMap[clientName].dueSoon,
    ]);

  if (clientLoadRows.length > 0) {
    dashboardSheet
      .getRange(clientLoadStartRow + 2, 1, clientLoadRows.length, clientLoadHeaders.length)
      .setValues(clientLoadRows);
  } else {
    dashboardSheet.getRange(clientLoadStartRow + 2, 1).setValue('Brak danych o klientach.');
  }

  const globalClientStatusStartRow =
    clientLoadStartRow + Math.max(clientLoadRows.length, 1) + 5;
  dashboardSheet
    .getRange(globalClientStatusStartRow, 1)
    .setValue('Status wykonania dla wszystkich klientow')
    .setFontWeight('bold');

  const globalClientHeaders = [
    'klient',
    'status_globalny',
    'otwarte',
    'przeterminowane',
    'na_horyzoncie',
    'wykonane_30dni',
    'najblizszy_termin',
  ];
  dashboardSheet
    .getRange(globalClientStatusStartRow + 1, 1, 1, globalClientHeaders.length)
    .setValues([globalClientHeaders])
    .setFontWeight('bold')
    .setBackground('#f1f3f4');

  const globalClientRows = allClientNames.map((clientName) => {
    const nameKey = normalizeLookupKey_(clientName);
    const clientTasks = tasks.filter(
      (task) => normalizeLookupKey_(task.clientName) === nameKey
    );
    const openClientTasks = clientTasks.filter((task) => task.status !== STATUS.DONE);
    const overdueClientTasks = openClientTasks.filter(
      (task) => task.dueDate && task.dueDate < today
    );
    const dueSoonClientTasks = openClientTasks.filter(
      (task) =>
        task.dueDate &&
        task.dueDate >= today &&
        task.dueDate <= dueSoonThreshold
    );
    const completedClientLast30 = clientTasks.filter(
      (task) =>
        task.status === STATUS.DONE &&
        task.completedAt &&
        task.completedAt >= last30Days
    );

    const nearestOpenDue = openClientTasks
      .filter((task) => task.dueDate)
      .sort((left, right) => left.dueDate - right.dueDate)[0];

    let globalStatus = 'BRAK_ZADAN';
    if (overdueClientTasks.length > 0) {
      globalStatus = 'PRZETERMINOWANE';
    } else if (dueSoonClientTasks.length > 0) {
      globalStatus = 'RYZYKO';
    } else if (openClientTasks.length > 0) {
      globalStatus = 'W_TOKU';
    } else if (completedClientLast30.length > 0) {
      globalStatus = 'OK';
    } else if (clientTasks.length > 0) {
      globalStatus = 'BRAK_OTWARTYCH';
    }

    return [
      clientName,
      globalStatus,
      openClientTasks.length,
      overdueClientTasks.length,
      dueSoonClientTasks.length,
      completedClientLast30.length,
      nearestOpenDue ? nearestOpenDue.dueDate : '',
    ];
  });

  if (globalClientRows.length > 0) {
    dashboardSheet
      .getRange(
        globalClientStatusStartRow + 2,
        1,
        globalClientRows.length,
        globalClientHeaders.length
      )
      .setValues(globalClientRows);
    dashboardSheet
      .getRange(globalClientStatusStartRow + 2, 7, globalClientRows.length, 1)
      .setNumberFormat('yyyy-mm-dd');
    applyClientStatusFormatting_(
      dashboardSheet,
      globalClientStatusStartRow + 2,
      globalClientRows.length
    );
  } else {
    dashboardSheet
      .getRange(globalClientStatusStartRow + 2, 1)
      .setValue('Brak klientow do podsumowania.');
  }

  applyManagerConditionalFormatting_(dashboardSheet, riskStartRow + 2, riskTasks.length);
}

function getAllClientNames_(clientRows, taskRows) {
  const map = {};
  clientRows.forEach((row) => {
    const name = normalizeText_(row.klient || row.client_id);
    if (!name) {
      return;
    }
    map[normalizeLookupKey_(name)] = name;
  });
  taskRows.forEach((task) => {
    const name = normalizeText_(task.clientName);
    if (!name) {
      return;
    }
    if (!map[normalizeLookupKey_(name)]) {
      map[normalizeLookupKey_(name)] = name;
    }
  });
  return Object.keys(map)
    .map((key) => map[key])
    .sort();
}

function applyClientStatusFormatting_(sheet, startRow, rowCount) {
  if (!rowCount) {
    return;
  }

  const statusRange = sheet.getRange(startRow, 2, rowCount, 1);
  const statuses = statusRange.getValues();
  statuses.forEach((statusRow, idx) => {
    const status = normalizeText_(statusRow[0]).toUpperCase();
    const cell = statusRange.getCell(idx + 1, 1);
    if (status === 'PRZETERMINOWANE') {
      cell.setBackground('#f8d7da');
      return;
    }
    if (status === 'RYZYKO') {
      cell.setBackground('#fff3cd');
      return;
    }
    if (status === 'OK') {
      cell.setBackground('#d4edda');
      return;
    }
    if (status === 'W_TOKU') {
      cell.setBackground('#e8f0fe');
      return;
    }
    cell.setBackground('#f1f3f4');
  });
}

function readManagerFilters_(sheet) {
  return {
    status: normalizeManagerStatusFilter_(sheet.getRange('B5').getValue()),
    employeeName:
      normalizeText_(sheet.getRange('B6').getValue()) || MANAGER_FILTER.ALL_EMPLOYEES,
    horizonDays: Math.max(
      1,
      toNumber_(sheet.getRange('B7').getValue(), MANAGER_FILTER.DEFAULT_HORIZON_DAYS)
    ),
    riskDays: Math.max(
      0,
      toNumber_(sheet.getRange('B8').getValue(), MANAGER_FILTER.DEFAULT_RISK_DAYS)
    ),
  };
}

function prepareManagerDashboardLayout_(sheet, employeeNames, previousFilters) {
  sheet.clear();
  sheet.setFrozenRows(3);

  sheet.getRange('A1').setValue('Dashboard managera');
  sheet
    .getRange('A1:G1')
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
  sheet.getRange('A1').setFontSize(14);
  sheet.getRange('A2').setFontStyle('italic').setFontColor('#5f6368');
  sheet.getRange('A3').setFontColor('#5f6368');

  sheet.getRange('A4').setValue('Filtry dashboardu').setFontWeight('bold');
  sheet.getRange('A5').setValue('Status');
  sheet.getRange('A6').setValue('Pracownik');
  sheet.getRange('A7').setValue('Horyzont terminu (dni)');
  sheet.getRange('A8').setValue('Prog zagrozenia (dni)');
  sheet.getRange('A4:B8').setBackground('#f8f9fa');

  const statusOptions = [
    MANAGER_FILTER.ALL,
    MANAGER_FILTER.OPEN,
    STATUS.NEW,
    STATUS.IN_PROGRESS,
    STATUS.DONE,
  ];
  const employeeOptions = [MANAGER_FILTER.ALL_EMPLOYEES].concat(employeeNames);

  const selectedStatus = statusOptions.includes(previousFilters.status)
    ? previousFilters.status
    : MANAGER_FILTER.OPEN;
  const selectedEmployee = employeeOptions.includes(previousFilters.employeeName)
    ? previousFilters.employeeName
    : MANAGER_FILTER.ALL_EMPLOYEES;

  sheet.getRange('B5').setValue(selectedStatus);
  sheet.getRange('B6').setValue(selectedEmployee);
  sheet
    .getRange('B7')
    .setValue(
      Math.max(1, toNumber_(previousFilters.horizonDays, MANAGER_FILTER.DEFAULT_HORIZON_DAYS))
    );
  sheet
    .getRange('B8')
    .setValue(Math.max(0, toNumber_(previousFilters.riskDays, MANAGER_FILTER.DEFAULT_RISK_DAYS)));

  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(statusOptions, true)
    .setAllowInvalid(false)
    .build();
  const employeeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(employeeOptions, true)
    .setAllowInvalid(false)
    .build();
  const positiveIntegerRule = SpreadsheetApp.newDataValidation()
    .requireNumberGreaterThanOrEqualTo(1)
    .setAllowInvalid(false)
    .setHelpText('Podaj liczbe >= 1.')
    .build();
  const nonNegativeIntegerRule = SpreadsheetApp.newDataValidation()
    .requireNumberGreaterThanOrEqualTo(0)
    .setAllowInvalid(false)
    .setHelpText('Podaj liczbe >= 0.')
    .build();

  sheet.getRange('B5').setDataValidation(statusRule);
  sheet.getRange('B6').setDataValidation(employeeRule);
  sheet.getRange('B7').setDataValidation(positiveIntegerRule);
  sheet.getRange('B8').setDataValidation(nonNegativeIntegerRule);
}

function buildManagerEmployeeLookups_(employeeRows) {
  const byName = {};
  const names = [];

  employeeRows.forEach((row) => {
    const employeeName = normalizeText_(row.pracownik || row.employee_id);
    if (!employeeName) {
      return;
    }
    names.push(employeeName);
    byName[employeeName] = employeeName;
  });

  names.sort();
  return {
    names,
    byName,
  };
}

function normalizeManagerStatusFilter_(value) {
  const normalized = normalizeText_(value).toUpperCase();
  const allowed = [
    MANAGER_FILTER.ALL,
    MANAGER_FILTER.OPEN,
    STATUS.NEW,
    STATUS.IN_PROGRESS,
    STATUS.DONE,
  ];
  if (allowed.includes(normalized)) {
    return normalized;
  }
  return MANAGER_FILTER.OPEN;
}

function matchesManagerStatusFilter_(task, statusFilter) {
  if (statusFilter === MANAGER_FILTER.ALL) {
    return true;
  }
  if (statusFilter === MANAGER_FILTER.OPEN) {
    return task.status !== STATUS.DONE;
  }
  return task.status === statusFilter;
}

function applyManagerConditionalFormatting_(sheet, startRow, taskCount) {
  if (!taskCount) {
    return;
  }

  const daysColumn = 7;
  const range = sheet.getRange(startRow, 1, taskCount, 7);
  const daysRange = sheet.getRange(startRow, daysColumn, taskCount, 1);
  const values = daysRange.getValues();

  values.forEach((row, idx) => {
    const days = toNumber_(row[0], 9999);
    if (days < 0) {
      range.getCell(idx + 1, 1).offset(0, 0, 1, 7).setBackground('#fde7e9');
      return;
    }
    if (days <= 1) {
      range.getCell(idx + 1, 1).offset(0, 0, 1, 7).setBackground('#fff4e5');
    }
  });
}

function resolveCurrentEmployee_() {
  const email = getCurrentUserEmail_();
  if (!email) {
    return null;
  }

  const employees = getObjectRows_(SHEET_NAMES.EMPLOYEES);
  const matched = employees.find(
    (row) =>
      normalizeText_(row.pracownik || row.employee_id) &&
      normalizeText_(row.email).toLowerCase() === email
  );

  if (!matched) {
    return null;
  }

  return {
    employeeName: normalizeText_(matched.pracownik || matched.employee_id),
    email,
    role: normalizeText_(matched.rola),
  };
}

function getWorkerSummary() {
  const employee = resolveCurrentEmployee_();
  if (!employee) {
    return {
      email: getCurrentUserEmail_(),
      employeeName: '',
      openTasks: 0,
      overdueTasks: 0,
      dueSoonTasks: 0,
      error:
        'Nie znaleziono aktywnego pracownika dla zalogowanego konta. Ustaw email w arkuszu Pracownicy.',
    };
  }

  const today = normalizeDate_(new Date());
  const soonDate = new Date(today.getTime());
  soonDate.setDate(soonDate.getDate() + MANAGER_FILTER.DEFAULT_HORIZON_DAYS);

  const tasks = getObjectRows_(SHEET_NAMES.TASKS)
    .filter(
      (row) =>
        normalizeLookupKey_(row.pracownik || row.employee_id) ===
        normalizeLookupKey_(employee.employeeName)
    )
    .filter((row) => normalizeText_(row.status) !== STATUS.DONE)
    .map((row) => ({
      dueDate: toDate_(row.due_date),
    }))
    .filter((row) => row.dueDate);

  const overdueTasks = tasks.filter((task) => task.dueDate < today).length;
  const dueSoonTasks = tasks.filter(
    (task) => task.dueDate >= today && task.dueDate <= soonDate
  ).length;

  return {
    email: employee.email,
    employeeName: employee.employeeName,
    openTasks: tasks.length,
    overdueTasks,
    dueSoonTasks,
    error: '',
  };
}

function getManagerSummary() {
  const today = normalizeDate_(new Date());
  const soonDate = new Date(today.getTime());
  soonDate.setDate(soonDate.getDate() + MANAGER_FILTER.DEFAULT_HORIZON_DAYS);

  const openTasks = getObjectRows_(SHEET_NAMES.TASKS).filter(
    (row) => normalizeText_(row.status) !== STATUS.DONE
  );
  const overdueTasks = openTasks.filter((row) => {
    const dueDate = toDate_(row.due_date);
    return dueDate && dueDate < today;
  }).length;
  const dueSoonTasks = openTasks.filter((row) => {
    const dueDate = toDate_(row.due_date);
    return dueDate && dueDate >= today && dueDate <= soonDate;
  }).length;

  return {
    openTasks: openTasks.length,
    overdueTasks,
    dueSoonTasks,
  };
}
