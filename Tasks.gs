function generateTasks30Days() {
  const generationResult = generateRecurringTasks(DEFAULT_GENERATION_DAYS);
  const createdCount = generationResult.createdCount || 0;
  const reassignedCount = generationResult.reassignedCount || 0;
  refreshManagerDashboard();
  try {
    refreshMyTasksView();
  } catch (error) {
    // Brak mapowania pracownika nie powinien blokowac generowania.
  }

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Utworzono ' +
      createdCount +
      ' nowych zadan, uzupelniono przypisanie w ' +
      reassignedCount +
      ' zadaniach.',
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
    normalizeText_(row.klient)
  );
  const employees = getObjectRows_(SHEET_NAMES.EMPLOYEES).filter((row) =>
    normalizeText_(row.pracownik || row.employee_id)
  );
  const assignableEmployeeNames = getAssignableEmployeeNames_(employees);
  const clientProcedures = getObjectRows_(SHEET_NAMES.CLIENT_PROCEDURES).filter(
    (row) => normalizeText_(row.klient || row.client_id) && normalizeText_(row.procedura || row.procedure_id)
  );
  const assignments = getObjectRows_(SHEET_NAMES.ASSIGNMENTS).filter((row) =>
    normalizeText_(row.klient || row.client_id)
  );
  const existingTasks = getObjectRows_(SHEET_NAMES.TASKS);

  const activeClientsByKey = buildNameMapByKey_(clients, ['klient', 'client_id']);
  const employeesByKey = buildNameMapByKey_(employees, ['pracownik', 'employee_id']);

  const proceduresByName = buildProcedureConfigs_(procedures);

  const assignmentsByClient = buildAssignmentsByClient_(
    assignments,
    activeClientsByKey,
    employeesByKey,
    assignableEmployeeNames
  );
  const existingKeys = new Set();
  const employeeByTaskKey = {};
  const rowByTaskKey = {};
  const lastTaskByPair = {};
  const reassignmentUpdates = [];

  existingTasks.forEach((row, idx) => {
    const clientRaw = normalizeText_(row.klient || row.client_id);
    const clientLookupKey = normalizeLookupKey_(clientRaw);
    const clientName = activeClientsByKey[clientLookupKey] || clientRaw;

    const procedureRaw = normalizeText_(row.procedura || row.procedure_id);
    const procedureLookupKey = normalizeLookupKey_(procedureRaw);
    const procedureConfig = proceduresByName[procedureLookupKey];
    const procedureName = procedureConfig ? procedureConfig.procedureName : procedureRaw;

    const employeeRaw = normalizeText_(row.pracownik || row.employee_id);
    let employeeName =
      employeesByKey[normalizeLookupKey_(employeeRaw)] || employeeRaw;
    const dueDate = toDate_(row.due_date);
    if (!clientName || !procedureName || !dueDate) {
      return;
    }

    const taskKey =
      normalizeText_(row.task_key) ||
      buildTaskKey_(clientName, procedureName, dueDate);
    existingKeys.add(taskKey);

    if (!employeeName) {
      const filledEmployeeName = pickNextEmployeeForDate_(
        assignmentsByClient[clientLookupKey] || [],
        dueDate,
        ''
      );
      if (filledEmployeeName) {
        employeeName = filledEmployeeName;
        reassignmentUpdates.push({
          rowNumber: idx + 2,
          employeeName: filledEmployeeName,
        });
      }
    }

    employeeByTaskKey[taskKey] = employeeName;
    rowByTaskKey[taskKey] = idx + 2;

    const pairKey = normalizeLookupKey_(clientName) + '|' + normalizeLookupKey_(procedureName);
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
    const clientRaw = normalizeText_(relation.klient || relation.client_id);
    const clientLookupKey = normalizeLookupKey_(clientRaw);
    const clientName = activeClientsByKey[clientLookupKey];

    const procedureRaw = normalizeText_(relation.procedura || relation.procedure_id);
    const procedureLookupKey = normalizeLookupKey_(procedureRaw);
    const procedure = proceduresByName[procedureLookupKey];
    const procedureName = procedure ? procedure.procedureName : '';

    if (!clientName || !procedureName) {
      return;
    }

    const relationStartDate = toDate_(relation.data_start) || today;
    const windowStart = relationStartDate > today ? relationStartDate : today;
    if (windowStart > horizon) {
      return;
    }

    const pairKey =
      normalizeLookupKey_(clientName) +
      '|' +
      normalizeLookupKey_(procedureName);
    let previousEmployeeName = lastTaskByPair[pairKey]
      ? normalizeText_(lastTaskByPair[pairKey].employeeName)
      : '';

    const dueDates = listScheduledDueDatesForWindow_(
      windowStart,
      horizon,
      relationStartDate,
      procedure
    );
    dueDates.forEach((normalizedDueDate) => {

      const taskKey = buildTaskKey_(clientName, procedureName, normalizedDueDate);
      if (existingKeys.has(taskKey)) {
        const existingEmployeeName = employeeByTaskKey[taskKey] || '';
        if (existingEmployeeName) {
          previousEmployeeName = existingEmployeeName;
          return;
        }

        const reassignedEmployeeName = pickNextEmployeeForDate_(
          assignmentsByClient[clientLookupKey] || [],
          normalizedDueDate,
          previousEmployeeName
        );
        if (!reassignedEmployeeName) {
          return;
        }

        const rowNumber = rowByTaskKey[taskKey];
        if (rowNumber) {
          reassignmentUpdates.push({
            rowNumber,
            employeeName: reassignedEmployeeName,
          });
          employeeByTaskKey[taskKey] = reassignedEmployeeName;
          previousEmployeeName = reassignedEmployeeName;
        }
        return;
      }

      const employeeName = pickNextEmployeeForDate_(
        assignmentsByClient[clientLookupKey] || [],
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
    const startRow = taskSheet.getLastRow() + 1;
    ensureSheetSize_(taskSheet, startRow + newRows.length - 1, HEADERS.TASKS.length);
    taskSheet
      .getRange(startRow, 1, newRows.length, HEADERS.TASKS.length)
      .setValues(newRows);
  }

  if (reassignmentUpdates.length > 0) {
    const taskSheet = getSheetOrThrow_(SHEET_NAMES.TASKS);
    reassignmentUpdates.forEach((update) => {
      taskSheet.getRange(update.rowNumber, 4).setValue(update.employeeName);
    });
  }

  if (newRows.length > 0 || reassignmentUpdates.length > 0) {
    sortTasksByStatusAndDueDesc_();
  }

  try {
    refreshClientProceduresControl();
  } catch (error) {
    // Kontrola nie powinna blokowac generowania.
  }

  return {
    createdCount: newRows.length,
    reassignedCount: reassignmentUpdates.length,
  };
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
  sortTasksByStatusAndDueDesc_();
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

function updateTaskStatus_(taskId, newStatus) {
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

  const rowNumber = match.getRow();
  sheet.getRange(rowNumber, 6).setValue(newStatus);
  if (newStatus === STATUS.DONE) {
    sheet.getRange(rowNumber, 8).setValue(new Date());
  } else {
    sheet.getRange(rowNumber, 8).clearContent();
  }
  sortTasksByStatusAndDueDesc_();
  return true;
}

function buildAssignmentsByClient_(
  assignmentRows,
  clientsByKey,
  employeesByKey,
  assignableEmployeeNames
) {
  const map = {};
  const employeeRankByKey = {};
  (assignableEmployeeNames || []).forEach((employeeName, idx) => {
    employeeRankByKey[normalizeLookupKey_(employeeName)] = idx;
  });

  assignmentRows.forEach((row) => {
    const clientRaw = normalizeText_(row.klient || row.client_id);
    const employeeRaw = normalizeText_(row.pracownik || row.employee_id);
    const clientLookupKey = normalizeLookupKey_(clientRaw);
    const clientName = (clientsByKey && clientsByKey[clientLookupKey]) || clientRaw;
    if (!clientName) {
      return;
    }

    const expandedEmployeeNames = [];
    if (employeeRaw) {
      const employeeLookupKey = normalizeLookupKey_(employeeRaw);
      const employeeName =
        (employeesByKey && employeesByKey[employeeLookupKey]) || employeeRaw;
      if (employeeName) {
        expandedEmployeeNames.push(employeeName);
      }
    } else if (assignableEmployeeNames && assignableEmployeeNames.length > 0) {
      assignableEmployeeNames.forEach((name) => expandedEmployeeNames.push(name));
    }
    if (expandedEmployeeNames.length === 0) {
      return;
    }

    if (!map[clientLookupKey]) {
      map[clientLookupKey] = [];
    }

    const fromDate = toDate_(row.data_od);
    const toDate = toDate_(row.data_do);
    const order = Math.max(1, toNumber_(row.kolejnosc, 9999));

    expandedEmployeeNames.forEach((employeeName) => {
      const rankKey = normalizeLookupKey_(employeeName);
      const rankValue = employeeRankByKey[rankKey];
      map[clientLookupKey].push({
        employeeName: employeeName,
        fromDate,
        toDate,
        order,
        employeeRank:
          typeof rankValue === 'number' ? rankValue : Number.MAX_SAFE_INTEGER,
      });
    });
  });

  Object.keys(map).forEach((clientLookupKey) => {
    map[clientLookupKey].sort((left, right) => {
      if (left.order !== right.order) {
        return left.order - right.order;
      }
      if (left.employeeRank !== right.employeeRank) {
        return left.employeeRank - right.employeeRank;
      }
      return normalizeLookupKey_(left.employeeName).localeCompare(
        normalizeLookupKey_(right.employeeName)
      );
    });
  });

  return map;
}

function getAssignableEmployeeNames_(employeeRows) {
  const names = [];
  const seen = {};

  employeeRows.forEach((row) => {
    const employeeName = normalizeText_(row.pracownik || row.employee_id);
    if (!employeeName) {
      return;
    }
    if (!toBoolean_(row.aktywny, true)) {
      return;
    }
    const key = normalizeLookupKey_(employeeName);
    if (!seen[key]) {
      names.push(employeeName);
      seen[key] = true;
    }
  });

  return names;
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

function getLatestAssignmentEndDateForClient_(assignmentRows) {
  if (!assignmentRows || assignmentRows.length === 0) {
    return null;
  }
  let latest = null;
  assignmentRows.forEach((row) => {
    const d = toDate_(row.data_do);
    if (d) {
      const norm = normalizeDate_(d);
      if (!latest || norm.getTime() > latest.getTime()) {
        latest = norm;
      }
    }
  });
  return latest;
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

  const sourceAssignments =
    eligibleAssignments.length > 0 ? eligibleAssignments : clientAssignments;

  const sorted = sourceAssignments.sort((left, right) => {
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
    const procedureName = normalizeText_(row.procedura || row.procedure_id);
    if (!procedureName) {
      return;
    }

    const rawMode = normalizeText_(row.tryb_harmonogramu).toUpperCase();
    const scheduleMode =
      rawMode === SCHEDULE_MODE.DAILY
        ? SCHEDULE_MODE.DAILY
        : SCHEDULE_MODE.MONTHLY;
    const interval = Math.max(1, toNumber_(row.interwal, 1));
    const schedule = parseScheduleDay_(row.dzien_miesiaca);
    if (scheduleMode === SCHEDULE_MODE.MONTHLY && !schedule) {
      return;
    }

    map[normalizeLookupKey_(procedureName)] = {
      procedureName,
      scheduleMode,
      interval,
      schedule,
      warningDays: Math.max(0, toNumber_(row.dni_ostrzezenia, 2)),
    };
  });
  return map;
}

function listScheduledDueDatesForWindow_(
  windowStart,
  horizon,
  relationStartDate,
  procedureConfig
) {
  if (!procedureConfig) {
    return [];
  }

  const normalizedWindowStart = normalizeDate_(windowStart);
  const normalizedHorizon = normalizeDate_(horizon);
  const normalizedRelationStart = normalizeDate_(relationStartDate);
  if (normalizedWindowStart > normalizedHorizon) {
    return [];
  }

  if (procedureConfig.scheduleMode === SCHEDULE_MODE.DAILY) {
    return getDailyDueDatesBetween_(
      normalizedWindowStart,
      normalizedHorizon,
      normalizedRelationStart,
      procedureConfig.interval
    );
  }

  return getMonthlyDueDatesBetween_(
    normalizedWindowStart,
    normalizedHorizon,
    normalizedRelationStart,
    procedureConfig.schedule,
    procedureConfig.interval
  );
}

function getMonthlyDueDatesBetween_(
  windowStart,
  horizon,
  relationStartDate,
  schedule,
  intervalMonths
) {
  if (!schedule) {
    return [];
  }

  const safeInterval = Math.max(1, toNumber_(intervalMonths, 1));
  const anchorMonthStart = getFirstDayOfMonth_(relationStartDate);
  const monthStarts = getMonthStartsBetween_(windowStart, horizon);
  const dueDates = [];

  monthStarts.forEach((monthStart) => {
    if (monthStart < anchorMonthStart) {
      return;
    }

    const monthsDiff =
      (monthStart.getFullYear() - anchorMonthStart.getFullYear()) * 12 +
      (monthStart.getMonth() - anchorMonthStart.getMonth());
    if (monthsDiff % safeInterval !== 0) {
      return;
    }

    const dueDate = getDueDateForMonth_(
      monthStart.getFullYear(),
      monthStart.getMonth(),
      schedule
    );
    if (!dueDate || dueDate < relationStartDate || dueDate < windowStart || dueDate > horizon) {
      return;
    }
    dueDates.push(dueDate);
  });

  return dueDates;
}

function getDailyDueDatesBetween_(windowStart, horizon, relationStartDate, intervalDays) {
  const safeInterval = Math.max(1, toNumber_(intervalDays, 1));
  const dueDates = [];
  if (relationStartDate > horizon) {
    return dueDates;
  }

  let firstDueDate = new Date(relationStartDate.getTime());
  if (firstDueDate < windowStart) {
    const daysDiff = Math.floor((windowStart.getTime() - firstDueDate.getTime()) / ONE_DAY_MS);
    const steps = Math.ceil(daysDiff / safeInterval);
    firstDueDate = new Date(firstDueDate.getTime());
    firstDueDate.setDate(firstDueDate.getDate() + steps * safeInterval);
    firstDueDate = normalizeDate_(firstDueDate);
  }

  for (
    let cursor = new Date(firstDueDate.getTime());
    cursor <= horizon;
    cursor = new Date(cursor.getFullYear(), cursor.getMonth(), cursor.getDate() + safeInterval)
  ) {
    if (cursor >= windowStart) {
      dueDates.push(normalizeDate_(cursor));
    }
  }

  return dueDates;
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
  const procedureConfig =
    proceduresByName[normalizeLookupKey_(completedTask.procedureName)];
  if (!procedureConfig) {
    return false;
  }

  const nextDueDate = getNextDueDateForProcedure_(
    completedTask.dueDate,
    procedureConfig
  );
  if (!nextDueDate) {
    return false;
  }

  const assignments = getObjectRows_(SHEET_NAMES.ASSIGNMENTS).filter((row) =>
    normalizeText_(row.klient || row.client_id)
  );
  const clientKey = normalizeLookupKey_(completedTask.clientName);
  const latestAssignmentEnd = getLatestAssignmentEndDateForClient_(
    assignments.filter(
      (row) => normalizeLookupKey_(row.klient || row.client_id) === clientKey
    )
  );
  if (
    latestAssignmentEnd &&
    normalizeDate_(nextDueDate).getTime() > normalizeDate_(latestAssignmentEnd).getTime()
  ) {
    return false;
  }

  const taskKey = buildTaskKey_(
    completedTask.clientName,
    procedureConfig.procedureName,
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

  const clients = getObjectRows_(SHEET_NAMES.CLIENTS).filter((row) =>
    normalizeText_(row.klient)
  );
  const employees = getObjectRows_(SHEET_NAMES.EMPLOYEES).filter((row) =>
    normalizeText_(row.pracownik || row.employee_id)
  );
  const assignableEmployeeNames = getAssignableEmployeeNames_(employees);
  const assignmentsByClient = buildAssignmentsByClient_(
    assignments,
    buildNameMapByKey_(clients, ['klient', 'client_id']),
    buildNameMapByKey_(employees, ['pracownik', 'employee_id']),
    assignableEmployeeNames
  );
  const employeeName = pickNextEmployeeForDate_(
    assignmentsByClient[normalizeLookupKey_(completedTask.clientName)] || [],
    nextDueDate,
    completedTask.employeeName
  );

  const row = [
    Utilities.getUuid(),
    completedTask.clientName,
    procedureConfig.procedureName,
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
  ensureSheetSize_(taskSheet, taskSheet.getLastRow() + 1, HEADERS.TASKS.length);
  taskSheet
    .getRange(taskSheet.getLastRow() + 1, 1, 1, HEADERS.TASKS.length)
    .setValues([row]);
  sortTasksByStatusAndDueDesc_();
  return true;
}

function getNextDueDateForProcedure_(dueDate, procedureConfig) {
  if (!procedureConfig) {
    return null;
  }

  const normalizedDueDate = normalizeDate_(dueDate);
  const safeInterval = Math.max(1, toNumber_(procedureConfig.interval, 1));

  if (procedureConfig.scheduleMode === SCHEDULE_MODE.DAILY) {
    const nextDailyDate = new Date(normalizedDueDate.getTime());
    nextDailyDate.setDate(nextDailyDate.getDate() + safeInterval);
    return normalizeDate_(nextDailyDate);
  }

  if (!procedureConfig.schedule) {
    return null;
  }

  const nextMonth = new Date(
    normalizedDueDate.getFullYear(),
    normalizedDueDate.getMonth() + safeInterval,
    1
  );
  return getDueDateForMonth_(
    nextMonth.getFullYear(),
    nextMonth.getMonth(),
    procedureConfig.schedule
  );
}

function sortTasksByStatusAndDueDesc_() {
  const taskSheet = getSheetOrThrow_(SHEET_NAMES.TASKS);
  const lastRow = taskSheet.getLastRow();
  if (lastRow < 2) {
    return;
  }

  const range = taskSheet.getRange(2, 1, lastRow - 1, HEADERS.TASKS.length);
  const rows = range.getValues();
  if (!rows || rows.length < 1) {
    return;
  }

  const toTimestamp = (value) => {
    const dateValue = toDate_(value);
    return dateValue ? dateValue.getTime() : -1;
  };

  rows.sort((left, right) => {
    const leftDone = normalizeText_(left[5]).toUpperCase() === STATUS.DONE;
    const rightDone = normalizeText_(right[5]).toUpperCase() === STATUS.DONE;
    if (leftDone !== rightDone) {
      return leftDone ? 1 : -1;
    }

    const dueDiff = toTimestamp(left[4]) - toTimestamp(right[4]);
    if (dueDiff !== 0) {
      return dueDiff;
    }

    const createdDiff = toTimestamp(right[6]) - toTimestamp(left[6]);
    if (createdDiff !== 0) {
      return createdDiff;
    }

    return normalizeText_(left[0]).localeCompare(normalizeText_(right[0]));
  });

  range.setValues(rows);
  highlightLateCompletedTasks_(taskSheet, rows);
}

function highlightLateCompletedTasks_(taskSheet, rows) {
  if (!rows || rows.length === 0) {
    return;
  }

  const statusBackgrounds = rows.map((row) => {
    const status = normalizeText_(row[5]).toUpperCase();
    const dueDate = toDate_(row[4]);
    const completedAt = toDate_(row[7]);
    const isLateDone =
      status === STATUS.DONE &&
      dueDate &&
      completedAt &&
      completedAt.getTime() > dueDate.getTime();

    return [isLateDone ? '#fde7e9' : ''];
  });

  taskSheet.getRange(2, 6, rows.length, 1).setBackgrounds(statusBackgrounds);
}

function buildNameMapByKey_(rows, fieldCandidates) {
  const map = {};
  rows.forEach((row) => {
    let name = '';
    for (let i = 0; i < fieldCandidates.length; i += 1) {
      name = normalizeText_(row[fieldCandidates[i]]);
      if (name) {
        break;
      }
    }
    if (!name) {
      return;
    }
    map[normalizeLookupKey_(name)] = name;
  });
  return map;
}

function onEdit(e) {
  if (!e || !e.range) {
    return;
  }

  const sheet = e.range.getSheet();
  if (enforceMasterDataIntegerRulesOnEdit_(sheet, e.range)) {
    return;
  }

  const editedSheetName = sheet.getName();
  if (editedSheetName === SHEET_NAMES.TASKS && e.range.getRow() > 1) {
    if (!isCurrentUserManager_()) {
      e.range.setValue(e.oldValue !== undefined ? e.oldValue : '');
      SpreadsheetApp.getActiveSpreadsheet().toast(
        'Tylko manager moze edytowac arkusz Zadania. Uzyj Moje_zadania do zmiany statusu.',
        'Uprawnienia',
        5
      );
      return;
    }
    const editedColumn = e.range.getColumn();
    if (editedColumn === 5) {
      sortTasksByStatusAndDueDesc_();
      return;
    }
    if (editedColumn === 4) {
      try {
        // Odswiezamy widok biezacego uzytkownika, aby od razu zobaczyl
        // efekt przepisania zadania do innego pracownika.
        refreshMyTasksView();
      } catch (error) {
        // Manager bez mapowania pracownika moze przepisywac zadania.
      }
      refreshManagerDashboard();
      return;
    }
  }

  const isMyTasksSheet =
    editedSheetName === SHEET_NAMES.MY_TASKS ||
    editedSheetName.startsWith(MY_TASKS_SHEET_PREFIX);
  if (!isMyTasksSheet || e.range.getRow() === 1) {
    return;
  }

  const row = e.range.getRow();
  const col = e.range.getColumn();

  if (col === MY_TASKS_COL.STATUS) {
    const taskId = normalizeText_(
      sheet.getRange(row, MY_TASKS_COL.TASK_ID).getValue()
    );
    if (!taskId) {
      return;
    }

    const newStatus = normalizeText_(e.range.getValue()).toUpperCase();
    const allowedStatuses = [STATUS.NEW, STATUS.IN_PROGRESS, STATUS.DONE];
    if (!allowedStatuses.includes(newStatus)) {
      refreshMyTasksView();
      return;
    }

    if (newStatus === STATUS.DONE) {
      const completedTask = markTaskAsDone_(taskId);
      createNextTaskFromCompleted_(completedTask);
    } else {
      updateTaskStatus_(taskId, newStatus);
    }

    refreshMyTasksView();
    refreshManagerDashboard();
    if (editedSheetName.startsWith(MY_TASKS_SHEET_PREFIX)) {
      const employeeName = editedSheetName.slice(MY_TASKS_SHEET_PREFIX.length);
      writeMyTasksViewToSheet_(sheet, employeeName);
    }
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
    if (editedSheetName.startsWith(MY_TASKS_SHEET_PREFIX)) {
      const employeeName = editedSheetName.slice(MY_TASKS_SHEET_PREFIX.length);
      writeMyTasksViewToSheet_(sheet, employeeName);
    }
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

  if (sheetName === SHEET_NAMES.PROCEDURES && col === 6) {
    return validateIntegerCell_(range, value, 1, 'W kolumnie interwal wpisz liczbe calkowita >= 1.');
  }

  if (sheetName === SHEET_NAMES.ASSIGNMENTS && col === 5) {
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
