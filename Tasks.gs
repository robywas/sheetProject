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

  const activeClientIds = new Set(
    clients.map((row) => normalizeText_(row.client_id)).filter(Boolean)
  );

  const proceduresById = buildProcedureConfigs_(procedures);

  const assignmentsByClient = buildAssignmentsByClient_(assignments);
  const existingKeys = new Set();
  const employeeByTaskKey = {};
  const lastTaskByPair = {};

  existingTasks.forEach((row) => {
    const clientId = normalizeText_(row.client_id);
    const procedureId = normalizeText_(row.procedure_id);
    const employeeId = normalizeText_(row.employee_id);
    const dueDate = toDate_(row.due_date);
    if (!clientId || !procedureId || !dueDate) {
      return;
    }

    const taskKey = normalizeText_(row.task_key) || buildTaskKey_(clientId, procedureId, dueDate);
    existingKeys.add(taskKey);
    employeeByTaskKey[taskKey] = employeeId;

    const pairKey = clientId + '|' + procedureId;
    const current = lastTaskByPair[pairKey];
    if (!current || dueDate > current.dueDate) {
      lastTaskByPair[pairKey] = {
        dueDate,
        employeeId,
      };
    }
  });

  const newRows = [];
  clientProcedures.forEach((relation) => {
    const clientId = normalizeText_(relation.client_id);
    const procedureId = normalizeText_(relation.procedure_id);
    if (!clientId || !procedureId) {
      return;
    }
    if (!activeClientIds.has(clientId)) {
      return;
    }

    const procedure = proceduresById[procedureId];
    if (!procedure) {
      return;
    }

    const relationStartDate = toDate_(relation.data_start) || today;
    const windowStart = relationStartDate > today ? relationStartDate : today;
    if (windowStart > horizon) {
      return;
    }

    const pairKey = clientId + '|' + procedureId;
    let previousEmployeeId = lastTaskByPair[pairKey]
      ? normalizeText_(lastTaskByPair[pairKey].employeeId)
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

      const taskKey = buildTaskKey_(clientId, procedureId, normalizedDueDate);
      if (existingKeys.has(taskKey)) {
        previousEmployeeId = employeeByTaskKey[taskKey] || previousEmployeeId;
        return;
      }

      const employeeId = pickNextEmployeeForDate_(
        assignmentsByClient[clientId] || [],
        normalizedDueDate,
        previousEmployeeId
      );

      newRows.push([
        Utilities.getUuid(),
        clientId,
        procedureId,
        employeeId,
        normalizedDueDate,
        STATUS.NEW,
        new Date(),
        '',
        '',
        taskKey,
        procedure.warningDays || 0,
      ]);
      previousEmployeeId = employeeId || previousEmployeeId;
      existingKeys.add(taskKey);
      employeeByTaskKey[taskKey] = employeeId;
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
    clientId: normalizeText_(rowValues[1]),
    procedureId: normalizeText_(rowValues[2]),
    employeeId: normalizeText_(rowValues[3]),
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
    const clientId = normalizeText_(row.client_id);
    const employeeId = normalizeText_(row.employee_id);
    if (!clientId || !employeeId) {
      return;
    }

    if (!map[clientId]) {
      map[clientId] = [];
    }

    map[clientId].push({
      employeeId: employeeId,
      fromDate: toDate_(row.data_od),
      toDate: toDate_(row.data_do),
      order: Math.max(1, toNumber_(row.kolejnosc, 9999)),
    });
  });

  Object.keys(map).forEach((clientId) => {
    map[clientId].sort((left, right) => {
      if (left.order !== right.order) {
        return left.order - right.order;
      }
      return left.employeeId.localeCompare(right.employeeId);
    });
  });

  return map;
}

function pickNextEmployeeForDate_(clientAssignments, dueDate, previousEmployeeId) {
  const eligibleEmployeeIds = getEligibleEmployeeIdsForDate_(
    clientAssignments,
    dueDate
  );
  if (eligibleEmployeeIds.length === 0) {
    return '';
  }

  if (!previousEmployeeId) {
    return eligibleEmployeeIds[0];
  }

  const currentIdx = eligibleEmployeeIds.indexOf(previousEmployeeId);
  if (currentIdx === -1) {
    return eligibleEmployeeIds[0];
  }
  if (eligibleEmployeeIds.length === 1) {
    return eligibleEmployeeIds[0];
  }
  return eligibleEmployeeIds[(currentIdx + 1) % eligibleEmployeeIds.length];
}

function getEligibleEmployeeIdsForDate_(clientAssignments, dueDate) {
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
    return left.employeeId.localeCompare(right.employeeId);
  });

  const unique = [];
  const seen = new Set();
  sorted.forEach((assignment) => {
    if (seen.has(assignment.employeeId)) {
      return;
    }
    seen.add(assignment.employeeId);
    unique.push(assignment.employeeId);
  });
  return unique;
}

function buildProcedureConfigs_(procedureRows) {
  const map = {};
  procedureRows.forEach((row) => {
    if (!toBoolean_(row.aktywna, true)) {
      return;
    }

    const procedureId = normalizeText_(row.procedure_id);
    if (!procedureId) {
      return;
    }

    const schedule = parseScheduleDay_(row.dzien_miesiaca);
    if (!schedule) {
      return;
    }

    map[procedureId] = {
      schedule,
      warningDays: Math.max(0, toNumber_(row.dni_ostrzezenia, 2)),
    };
  });
  return map;
}

function createNextTaskFromCompleted_(completedTask) {
  if (
    !completedTask ||
    !completedTask.clientId ||
    !completedTask.procedureId ||
    !completedTask.dueDate
  ) {
    return false;
  }

  const proceduresById = buildProcedureConfigs_(getObjectRows_(SHEET_NAMES.PROCEDURES));
  const procedureConfig = proceduresById[completedTask.procedureId];
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
    completedTask.clientId,
    completedTask.procedureId,
    nextDueDate
  );
  const taskRows = getObjectRows_(SHEET_NAMES.TASKS);
  const exists = taskRows.some((row) => {
    const existingClientId = normalizeText_(row.client_id);
    const existingProcedureId = normalizeText_(row.procedure_id);
    const dueDate = toDate_(row.due_date);
    if (!existingClientId || !existingProcedureId || !dueDate) {
      return false;
    }

    const existingKey =
      normalizeText_(row.task_key) ||
      buildTaskKey_(existingClientId, existingProcedureId, dueDate);
    return existingKey === taskKey;
  });
  if (exists) {
    return false;
  }

  const assignments = getObjectRows_(SHEET_NAMES.ASSIGNMENTS).filter((row) =>
    toBoolean_(row.aktywna, true)
  );
  const assignmentsByClient = buildAssignmentsByClient_(assignments);
  const employeeId = pickNextEmployeeForDate_(
    assignmentsByClient[completedTask.clientId] || [],
    nextDueDate,
    completedTask.employeeId
  );

  const row = [
    Utilities.getUuid(),
    completedTask.clientId,
    completedTask.procedureId,
    employeeId,
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

  if (sheetName === SHEET_NAMES.PROCEDURES && col === 5) {
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
