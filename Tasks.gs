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

  const procedures = getObjectRows_(SHEET_NAMES.PROCEDURES).filter((row) =>
    toBoolean_(row.aktywna, true)
  );
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

  const proceduresById = {};
  procedures.forEach((row) => {
    const id = normalizeText_(row.procedure_id);
    if (!id) {
      return;
    }

    const frequencyDays = toNumber_(row.czestotliwosc_dni, 0);
    if (frequencyDays < 1) {
      return;
    }

    proceduresById[id] = {
      frequencyDays,
      warningDays: Math.max(0, toNumber_(row.dni_ostrzezenia, 2)),
    };
  });

  const assignmentsByClient = buildAssignmentsByClient_(assignments);
  const existingKeys = new Set();
  const lastDueByPair = {};

  existingTasks.forEach((row) => {
    const clientId = normalizeText_(row.client_id);
    const procedureId = normalizeText_(row.procedure_id);
    const dueDate = toDate_(row.due_date);
    if (!clientId || !procedureId || !dueDate) {
      return;
    }

    const pairKey = clientId + '|' + procedureId;
    if (!lastDueByPair[pairKey] || dueDate > lastDueByPair[pairKey]) {
      lastDueByPair[pairKey] = dueDate;
    }

    const rowKey = normalizeText_(row.task_key) || buildTaskKey_(clientId, procedureId, dueDate);
    existingKeys.add(rowKey);
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

    const frequencyDays = Math.max(
      1,
      toNumber_(relation.czestotliwosc_override, procedure.frequencyDays)
    );
    const relationStartDate = toDate_(relation.data_start) || today;

    const pairKey = clientId + '|' + procedureId;
    let nextDueDate = alignDateToWindow_(relationStartDate, today, frequencyDays);

    if (lastDueByPair[pairKey] && lastDueByPair[pairKey] >= nextDueDate) {
      nextDueDate = new Date(lastDueByPair[pairKey].getTime());
      nextDueDate.setDate(nextDueDate.getDate() + frequencyDays);
    }

    for (
      let dueDate = new Date(nextDueDate.getTime());
      dueDate <= horizon;
      dueDate.setDate(dueDate.getDate() + frequencyDays)
    ) {
      const normalizedDueDate = normalizeDate_(dueDate);
      const taskKey = buildTaskKey_(clientId, procedureId, normalizedDueDate);
      if (existingKeys.has(taskKey)) {
        continue;
      }

      const employeeId = findAssignedEmployeeForDate_(
        assignmentsByClient[clientId] || [],
        normalizedDueDate
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
        procedure.warningDays,
      ]);
      existingKeys.add(taskKey);
    }
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
    return false;
  }

  const taskIdRange = sheet.getRange(2, 1, lastRow - 1, 1);
  const match = taskIdRange
    .createTextFinder(taskId)
    .matchEntireCell(true)
    .findNext();

  if (!match) {
    return false;
  }

  const row = match.getRow();
  sheet.getRange(row, 6).setValue(STATUS.DONE);
  sheet.getRange(row, 8).setValue(new Date());
  return true;
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
    });
  });

  Object.keys(map).forEach((clientId) => {
    map[clientId].sort((left, right) => {
      const leftTime = left.fromDate ? left.fromDate.getTime() : 0;
      const rightTime = right.fromDate ? right.fromDate.getTime() : 0;
      return rightTime - leftTime;
    });
  });

  return map;
}

function findAssignedEmployeeForDate_(clientAssignments, dueDate) {
  if (!clientAssignments || clientAssignments.length === 0) {
    return '';
  }

  const targetDate = normalizeDate_(dueDate);

  const exactMatch = clientAssignments.find((assignment) => {
    const startsBeforeOrOn = !assignment.fromDate || assignment.fromDate <= targetDate;
    const endsAfterOrOn = !assignment.toDate || assignment.toDate >= targetDate;
    return startsBeforeOrOn && endsAfterOrOn;
  });

  if (exactMatch) {
    return exactMatch.employeeId;
  }

  const fallback = clientAssignments.find(
    (assignment) => !assignment.fromDate && !assignment.toDate
  );
  return fallback ? fallback.employeeId : '';
}

function onEdit(e) {
  if (!e || !e.range) {
    return;
  }

  const sheet = e.range.getSheet();
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

    markTaskAsDone_(taskId);
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
