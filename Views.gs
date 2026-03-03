function refreshMyTasksView() {
  const employee = resolveCurrentEmployee_();
  if (!employee) {
    throw new Error(
      'Nie znaleziono aktywnego pracownika dla Twojego emaila. Uzuplnij arkusz Pracownicy.'
    );
  }

  const taskRows = getObjectRows_(SHEET_NAMES.TASKS);
  const openTasks = taskRows
    .filter((row) => normalizeText_(row.employee_id) === employee.employeeId)
    .filter((row) => normalizeText_(row.status) !== STATUS.DONE)
    .map((row) => {
      return {
        taskId: normalizeText_(row.task_id),
        patientId: normalizeText_(row.patient_id),
        procedureId: normalizeText_(row.procedure_id),
        dueDate: toDate_(row.due_date),
        status: normalizeText_(row.status) || STATUS.NEW,
        note: row.notes || '',
      };
    })
    .filter((task) => task.taskId && task.dueDate)
    .sort((left, right) => left.dueDate - right.dueDate);

  const patientNames = getLookupMap_(
    getObjectRows_(SHEET_NAMES.PATIENTS),
    'patient_id',
    'pacjent'
  );
  const procedureNames = getLookupMap_(
    getObjectRows_(SHEET_NAMES.PROCEDURES),
    'procedure_id',
    'procedura'
  );

  const myTasksSheet = getSheetOrThrow_(SHEET_NAMES.MY_TASKS);
  myTasksSheet.getRange(1, 1, 1, HEADERS.MY_TASKS.length).setValues([HEADERS.MY_TASKS]);
  clearSheetBody_(myTasksSheet, HEADERS.MY_TASKS.length);

  if (openTasks.length === 0) {
    myTasksSheet.getRange(2, 1).setValue('Brak otwartych zadan.');
    return;
  }

  const rows = openTasks.map((task) => [
    false,
    task.taskId,
    task.dueDate,
    patientNames[task.patientId] || task.patientId,
    procedureNames[task.procedureId] || task.procedureId,
    task.status,
    task.note,
  ]);

  myTasksSheet
    .getRange(2, 1, rows.length, HEADERS.MY_TASKS.length)
    .setValues(rows);
  myTasksSheet.getRange(2, 1, rows.length, 1).insertCheckboxes();
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

  myTasksSheet.autoResizeColumns(1, HEADERS.MY_TASKS.length);
}

function refreshManagerDashboard() {
  const dashboardSheet = getSheetOrThrow_(SHEET_NAMES.MANAGER_DASHBOARD);
  dashboardSheet.clear();

  const tasks = getObjectRows_(SHEET_NAMES.TASKS).map((row) => {
    const dueDate = toDate_(row.due_date);
    const completedAt = toDate_(row.completed_at);
    return {
      taskId: normalizeText_(row.task_id),
      patientId: normalizeText_(row.patient_id),
      procedureId: normalizeText_(row.procedure_id),
      employeeId: normalizeText_(row.employee_id),
      dueDate,
      status: normalizeText_(row.status) || STATUS.NEW,
      completedAt,
    };
  });

  const today = normalizeDate_(new Date());
  const dueSoonThreshold = new Date(today.getTime());
  dueSoonThreshold.setDate(dueSoonThreshold.getDate() + 7);
  const riskThreshold = new Date(today.getTime());
  riskThreshold.setDate(riskThreshold.getDate() + 2);
  const last30Days = new Date(today.getTime());
  last30Days.setDate(last30Days.getDate() - 30);

  const openTasks = tasks.filter((task) => task.status !== STATUS.DONE);
  const overdueTasks = openTasks.filter((task) => task.dueDate && task.dueDate < today);
  const dueSoonTasks = openTasks.filter(
    (task) => task.dueDate && task.dueDate >= today && task.dueDate <= dueSoonThreshold
  );
  const completedLast30Days = tasks.filter(
    (task) => task.status === STATUS.DONE && task.completedAt && task.completedAt >= last30Days
  );
  const completionRate = tasks.length
    ? Math.round((completedLast30Days.length / tasks.length) * 100)
    : 0;

  dashboardSheet.getRange('A1').setValue('Dashboard managera');
  dashboardSheet.getRange('A2').setValue(
    'Aktualizacja: ' +
      Utilities.formatDate(
        new Date(),
        Session.getScriptTimeZone(),
        'yyyy-MM-dd HH:mm:ss'
      )
  );

  const kpiTable = [
    ['Wskaznik', 'Wartosc'],
    ['Otwarte zadania', openTasks.length],
    ['Przeterminowane', overdueTasks.length],
    ['Termin <= 7 dni', dueSoonTasks.length],
    ['Ukonczone (30 dni)', completedLast30Days.length],
    ['Wskaznik realizacji', completionRate + '%'],
  ];

  dashboardSheet.getRange(4, 1, kpiTable.length, 2).setValues(kpiTable);
  dashboardSheet.getRange(4, 1, 1, 2).setFontWeight('bold').setBackground('#f1f3f4');

  const patientNames = getLookupMap_(
    getObjectRows_(SHEET_NAMES.PATIENTS),
    'patient_id',
    'pacjent'
  );
  const procedureNames = getLookupMap_(
    getObjectRows_(SHEET_NAMES.PROCEDURES),
    'procedure_id',
    'procedura'
  );
  const employeeNames = getLookupMap_(
    getObjectRows_(SHEET_NAMES.EMPLOYEES),
    'employee_id',
    'pracownik'
  );

  const riskTasks = openTasks
    .filter((task) => task.dueDate && task.dueDate <= riskThreshold)
    .sort((left, right) => left.dueDate - right.dueDate)
    .map((task) => {
      const daysDiff = Math.floor((task.dueDate - today) / ONE_DAY_MS);
      return [
        task.taskId,
        task.dueDate,
        patientNames[task.patientId] || task.patientId,
        procedureNames[task.procedureId] || task.procedureId,
        employeeNames[task.employeeId] || task.employeeId || '(nieprzypisane)',
        task.status,
        daysDiff,
      ];
    });

  const riskStartRow = 12;
  const riskHeaders = [
    'task_id',
    'termin',
    'pacjent',
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
    const key = task.employeeId || 'NIEPRZYPISANE';
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

  const loadTableHeaders = ['pracownik', 'otwarte', 'przeterminowane', 'termin <= 7 dni'];
  dashboardSheet
    .getRange(loadStartRow + 1, 1, 1, loadTableHeaders.length)
    .setValues([loadTableHeaders])
    .setFontWeight('bold')
    .setBackground('#f1f3f4');

  const loadRows = Object.keys(employeeLoad)
    .sort()
    .map((employeeId) => [
      employeeNames[employeeId] || employeeId,
      employeeLoad[employeeId].open,
      employeeLoad[employeeId].overdue,
      employeeLoad[employeeId].dueSoon,
    ]);

  if (loadRows.length > 0) {
    dashboardSheet
      .getRange(loadStartRow + 2, 1, loadRows.length, loadTableHeaders.length)
      .setValues(loadRows);
  } else {
    dashboardSheet.getRange(loadStartRow + 2, 1).setValue('Brak danych o obciazeniu.');
  }

  dashboardSheet.autoResizeColumns(1, 7);
}

function resolveCurrentEmployee_() {
  const email = getCurrentUserEmail_();
  if (!email) {
    return null;
  }

  const employees = getObjectRows_(SHEET_NAMES.EMPLOYEES);
  const matched = employees.find(
    (row) =>
      toBoolean_(row.aktywny, true) &&
      normalizeText_(row.email).toLowerCase() === email
  );

  if (!matched) {
    return null;
  }

  return {
    employeeId: normalizeText_(matched.employee_id),
    name: normalizeText_(matched.pracownik),
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
  soonDate.setDate(soonDate.getDate() + 7);

  const tasks = getObjectRows_(SHEET_NAMES.TASKS)
    .filter((row) => normalizeText_(row.employee_id) === employee.employeeId)
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
    employeeName: employee.name || employee.employeeId,
    openTasks: tasks.length,
    overdueTasks,
    dueSoonTasks,
    error: '',
  };
}

function getManagerSummary() {
  const today = normalizeDate_(new Date());
  const soonDate = new Date(today.getTime());
  soonDate.setDate(soonDate.getDate() + 7);

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
