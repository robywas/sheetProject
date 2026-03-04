function generateTasks30Days() {
  const createdCount = generateRecurringTasks(DEFAULT_GENERATION_DAYS);
  refreshManagerDashboard();
  try {
    refreshMyTasksView();
  } catch (error) {
    // Brak mapowania pracownika nie powinien blokowac generowania.
  }

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Utworzono ' + createdCount + ' nowych zadan.',
    'Procedury',
    5
  );
}

function generateRecurringTasks(daysAhead) {
  const today = normalizeDate_(new Date());
  const horizon = new Date(today.getTime());
  horizon.setDate(horizon.getDate() + toNumber_(daysAhead, DEFAULT_GENERATION_DAYS));

  const procedures = getObjectRows_(SHEET_NAMES.PROCEDURES);
  const clients = getObjectRows_(SHEET_NAMES.CLIENTS).filter((row) =>
    toBoolean_(row.aktywny, true)
  );
  const clientProcedures = getObjectRows_(SHEET_NAMES.CLIENT_PROCEDURES).filter((row) =>
    toBoolean_(row.aktywna, true)
  );
  const assignments = getObjectRows_(SHEET_NAMES.ASSIGNMENTS).filter((row) =>
    toBoolean_(row.aktywna, true)
  );
  const existingTasks = getObjectRows_(SHEET_NAMES.TASKS);

  const activeClientNames = new Set(
    clients.map((row) => normalizeText_(row.klient)).filter(Boolean)
  );

  const proceduresByName = buildProcedureConfigs_(procedures);

  const assignmentsByClient = buildAssignmentsByClient_(assignments);
  const existingKeys = new Set();
  const employeeByTaskKey = {};
  const lastTaskByPair = {};

  existingTasks.forEach((row) => {
    const clientName = normalizeText_(row.klient);
    const procedureName = normalizeText_(row.procedura);
    const employeeName = normalizeText_(row.pracownik);
    const dueDate = toDate_(row.due_date);
    if (!clientName || !procedureName || !dueDate) {
      return;
    }

    const taskKey =
      normalizeText_(row.task_key) ||
      buildTaskKey_(clientName, procedureName, dueDate);
    existingKeys.add(taskKey);
    employeeByTaskKey[taskKey] = employeeName;

    const pairKey = clientName + '|' + procedureName;
    const current = lastTaskByPair[pairKey];
    if (!current || dueDate > current.dueDate) {
      lastTaskByPair[pairKey] = {
        dueDate,
        employeeName,
      };
    }
  });

  const newRows = [];
  clientProcedures.forEach((relation) => {
    const clientName = normalizeText_(relation.klient);
    const procedureName = normalizeText_(relation.procedura);
    if (!clientName || !procedureName) {
      return;
    }
    if (!activeClientNames.has(clientName)) {
      return;
    }

    const procedure = proceduresByName[procedureName];
    if (!procedure) {
      return;
    }

    const relationStartDate = toDate_(relation.data_start) || today;
    const windowStart = relationStartDate > today ? relationStartDate : today;
    if (windowStart > horizon) {
      return;
    }

    const pairKey = clientName + '|' + procedureName;
    let previousEmployeeName = lastTaskByPair[pairKey]
      ? normalizeText_(lastTaskByPair[pairKey].employeeName)
      : '';

    const months = getMonthStartsBetween_(windowStart, horizon);
    months.forEach((monthStart) => {
      const normalizedDueDate = getDueDateForMonth_(
        monthStart.getFullYear(),
        monthStart.getMonth(),
        procedure.schedule
      );
      if (!normalizedDueDate) {
        return;
      }
      if (normalizedDueDate < windowStart || normalizedDueDate > horizon) {
        return;
      }

      const taskKey = buildTaskKey_(clientName, procedureName, normalizedDueDate);
      if (existingKeys.has(taskKey)) {
        previousEmployeeName = employeeByTaskKey[taskKey] || previousEmployeeName;
        return;
      }

      const employeeName = pickNextEmployeeForDate_(
        assignmentsByClient[clientName] || [],
        normalizedDueDate,
        previousEmployeeName
      );

      newRows.push([
        Utilities.getUuid(),
        clientName,
        procedureName,
        employeeName,
        normalizedDueDate,
        STATUS.NEW,
        new Date(),
        '',
        '',
        taskKey,
        procedure.warningDays || 0,
      ]);
      previousEmployeeName = employeeName || previousEmployeeName;
      existingKeys.add(taskKey);
      employeeByTaskKey[taskKey] = employeeName;
    });
  });

  if (newRows.length > 0) {
    const taskSheet = getSheetOrThrow_(SHEET_NAMES.TASKS);
    taskSheet
      .getRange(taskSheet.getLastRow() + 1, 1, newRows.length, HEADERS.TASKS.length)
      .setValues(newRows);
  }

  return newRows.length;
}

function markTaskAsDone_(taskId) {
  const sheet = getSheetOrThrow_(SHEET_NAMES.TASKS);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return null;
  }

  const taskIdRange = sheet.getRange(2, 1, lastRow - 1, 1);
  const match = taskIdRange
    .createTextFinder(taskId)
    .matchEntireCell(true)
    .findNext();

  if (!match) {
    return null;
  }

  const row = match.getRow();
  const rowValues = sheet.getRange(row, 1, 1, HEADERS.TASKS.length).getValues()[0];
  const completedTask = {
    taskId: normalizeText_(rowValues[0]),
    clientName: normalizeText_(rowValues[1]),
    procedureName: normalizeText_(rowValues[2]),
    employeeName: normalizeText_(rowValues[3]),
    dueDate: toDate_(rowValues[4]),
    status: normalizeText_(rowValues[5]),
  };
  if (completedTask.status === STATUS.DONE) {
    return null;
  }

  sheet.getRange(row, 6).setValue(STATUS.DONE);
  sheet.getRange(row, 8).setValue(new Date());
  return completedTask;
}

function updateTaskNote_(taskId, note) {
  const sheet = getSheetOrThrow_(SHEET_NAMES.TASKS);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return;
  }

  const taskIdRange = sheet.getRange(2, 1, lastRow - 1, 1);
  const match = taskIdRange
    .createTextFinder(taskId)
    .matchEntireCell(true)
    .findNext();

  if (!match) {
    return;
  }

  sheet.getRange(match.getRow(), 9).setValue(note);
}

function buildAssignmentsByClient_(assignmentRows) {
  const map = {};
  assignmentRows.forEach((row) => {
    const clientName = normalizeText_(row.klient);
    const employeeName = normalizeText_(row.pracownik);
    if (!clientName || !employeeName) {
      return;
    }

    if (!map[clientName]) {
      map[clientName] = [];
    }

    map[clientName].push({
      employeeName: employeeName,
      fromDate: toDate_(row.data_od),
      toDate: toDate_(row.data_do),
      order: Math.max(1, toNumber_(row.kolejnosc, 9999)),
    });
  });

  Object.keys(map).forEach((clientName) => {
    map[clientName].sort((left, right) => {
      if (left.order !== right.order) {
        return left.order - right.order;
      }
      return left.employeeName.localeCompare(right.employeeName);
    });
  });

  return map;
}

function pickNextEmployeeForDate_(clientAssignments, dueDate, previousEmployeeName) {
  const eligibleEmployeeNames = getEligibleEmployeeNamesForDate_(
    clientAssignments,
    dueDate
  );
  if (eligibleEmployeeNames.length === 0) {
    return '';
  }

  if (!previousEmployeeName) {
    return eligibleEmployeeNames[0];
  }

  const currentIdx = eligibleEmployeeNames.indexOf(previousEmployeeName);
  if (currentIdx === -1) {
    return eligibleEmployeeNames[0];
  }
  if (eligibleEmployeeNames.length === 1) {
    return eligibleEmployeeNames[0];
  }
  return eligibleEmployeeNames[(currentIdx + 1) % eligibleEmployeeNames.length];
}

function getEligibleEmployeeNamesForDate_(clientAssignments, dueDate) {
  if (!clientAssignments || clientAssignments.length === 0) {
    return [];
  }

  const targetDate = normalizeDate_(dueDate);
  const eligibleAssignments = clientAssignments.filter((assignment) => {
    const startsBeforeOrOn = !assignment.fromDate || assignment.fromDate <= targetDate;
    const endsAfterOrOn = !assignment.toDate || assignment.toDate >= targetDate;
    return startsBeforeOrOn && endsAfterOrOn;
  });

  if (eligibleAssignments.length === 0) {
    return [];
  }

  const sorted = eligibleAssignments.sort((left, right) => {
    if (left.order !== right.order) {
      return left.order - right.order;
    }
    return left.employeeName.localeCompare(right.employeeName);
  });

  const unique = [];
  const seen = new Set();
  sorted.forEach((assignment) => {
    if (seen.has(assignment.employeeName)) {
      return;
    }
    seen.add(assignment.employeeName);
    unique.push(assignment.employeeName);
  });
  return unique;
}

function buildProcedureConfigs_(procedureRows) {
  const map = {};
  procedureRows.forEach((row) => {
    if (!toBoolean_(row.aktywna, true)) {
      return;
    }

    const procedureName = normalizeText_(row.procedura);
    if (!procedureName) {
      return;
    }

    const schedule = parseScheduleDay_(row.dzien_miesiaca);
    if (!schedule) {
      return;
    }

    map[procedureName] = {
      schedule,
      warningDays: Math.max(0, toNumber_(row.dni_ostrzezenia, 2)),
    };
  });
  return map;
}

function createNextTaskFromCompleted_(completedTask) {
  if (
    !completedTask ||
    !completedTask.clientName ||
    !completedTask.procedureName ||
    !completedTask.dueDate
  ) {
    return false;
  }

  const proceduresByName = buildProcedureConfigs_(getObjectRows_(SHEET_NAMES.PROCEDURES));
  const procedureConfig = proceduresByName[completedTask.procedureName];
  if (!procedureConfig) {
    return false;
  }

  const nextDueDate = getNextMonthlyDueDate_(
    completedTask.dueDate,
    procedureConfig.schedule
  );
  if (!nextDueDate) {
    return false;
  }

  const taskKey = buildTaskKey_(
    completedTask.clientName,
    completedTask.procedureName,
    nextDueDate
  );
  const taskRows = getObjectRows_(SHEET_NAMES.TASKS);
  const exists = taskRows.some((row) => {
    const existingClientName = normalizeText_(row.klient);
    const existingProcedureName = normalizeText_(row.procedura);
    const dueDate = toDate_(row.due_date);
    if (!existingClientName || !existingProcedureName || !dueDate) {
      return false;
    }

    const existingKey =
      normalizeText_(row.task_key) ||
      buildTaskKey_(existingClientName, existingProcedureName, dueDate);
    return existingKey === taskKey;
  });
  if (exists) {
    return false;
  }

  const assignments = getObjectRows_(SHEET_NAMES.ASSIGNMENTS).filter((row) =>
    toBoolean_(row.aktywna, true)
  );
  const assignmentsByClient = buildAssignmentsByClient_(assignments);
  const employeeName = pickNextEmployeeForDate_(
    assignmentsByClient[completedTask.clientName] || [],
    nextDueDate,
    completedTask.employeeName
  );

  const row = [
    Utilities.getUuid(),
    completedTask.clientName,
    completedTask.procedureName,
    employeeName,
    nextDueDate,
    STATUS.NEW,
    new Date(),
    '',
    '',
    taskKey,
    procedureConfig.warningDays || 0,
  ];

  const taskSheet = getSheetOrThrow_(SHEET_NAMES.TASKS);
  taskSheet
    .getRange(taskSheet.getLastRow() + 1, 1, 1, HEADERS.TASKS.length)
    .setValues([row]);
  return true;
}

function onEdit(e) {
  if (!e || !e.range) {
    return;
  }

  const sheet = e.range.getSheet();
  if (enforceMasterDataIntegerRulesOnEdit_(sheet, e.range)) {
    return;
  }

  if (sheet.getName() !== SHEET_NAMES.MY_TASKS || e.range.getRow() === 1) {
    return;
  }

  const row = e.range.getRow();
  const col = e.range.getColumn();

  if (col === MY_TASKS_COL.CHECKBOX && e.value === 'TRUE') {
    const taskId = normalizeText_(
      sheet.getRange(row, MY_TASKS_COL.TASK_ID).getValue()
    );
    if (!taskId) {
      return;
    }

    const completedTask = markTaskAsDone_(taskId);
    createNextTaskFromCompleted_(completedTask);
    refreshMyTasksView();
    refreshManagerDashboard();
    return;
  }

  if (col === MY_TASKS_COL.NOTE) {
    const taskId = normalizeText_(
      sheet.getRange(row, MY_TASKS_COL.TASK_ID).getValue()
    );
    if (!taskId) {
      return;
    }
    updateTaskNote_(taskId, e.range.getValue());
  }
}

function enforceMasterDataIntegerRulesOnEdit_(sheet, range) {
  if (range.getRow() === 1) {
    return false;
  }

  const sheetName = sheet.getName();
  const col = range.getColumn();
  const value = range.getValue();

  if (sheetName === SHEET_NAMES.PROCEDURES && col === 4) {
    return validateIntegerCell_(range, value, 0, 'W kolumnie dni_ostrzezenia wpisz liczbe calkowita >= 0.');
  }

  if (sheetName === SHEET_NAMES.ASSIGNMENTS && col === 6) {
    return validateIntegerCell_(range, value, 1, 'W kolumnie kolejnosc wpisz liczbe calkowita >= 1.');
  }

  if (sheetName === SHEET_NAMES.MANAGER_DASHBOARD && (col === 2 && (range.getRow() === 7 || range.getRow() === 8))) {
    const minValue = range.getRow() === 7 ? 1 : 0;
    const label = range.getRow() === 7 ? 'Horyzont terminu' : 'Prog zagrozenia';
    return validateIntegerCell_(
      range,
      value,
      minValue,
      label + ' musi byc liczba calkowita >= ' + minValue + '.'
    );
  }

  return false;
}

function validateIntegerCell_(range, value, minValue, errorMessage) {
  if (value === '' || value === null || typeof value === 'undefined') {
    return false;
  }

  const numericValue = Number(value);
  const isValidInteger =
    Number.isFinite(numericValue) &&
    numericValue >= minValue &&
    numericValue === Math.floor(numericValue);

  if (isValidInteger) {
    return false;
  }

  range.clearContent();
  SpreadsheetApp.getActiveSpreadsheet().toast(errorMessage, 'Walidacja', 5);
  return true;
}
