function refreshMyTasksView() {
  const employee = resolveCurrentEmployee_();
  if (!employee) {
    throw new Error(
      'Nie znaleziono aktywnego pracownika dla Twojego emaila. Uzuplnij arkusz Pracownicy.'
    );
  }
  refreshMyTasksViewForEmployeeName_(employee.employeeName);
}

/**
 * Formatuje nowo utworzony arkusz Zadania - X na wzor Zadania:
 * naglowek i dane – format kolumna po kolumnie (Zadania.due_date→termin, procedura→procedura itd.),
 * wysokosc wierszy, szerokosc kolumn, zamrozenie wiersza 1, ukrycie kolumny A (task_id).
 */
function formatNewMyTasksSheetFromTasks_(myTasksSheet) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const tasksSheet = spreadsheet.getSheetByName(SHEET_NAMES.TASKS);
  if (!tasksSheet) {
    return;
  }
  const pasteFormat = SpreadsheetApp.CopyPasteType.PASTE_FORMAT;
  const numCols = TASKS_COL_FOR_MY_TASKS_FORMAT.length;

  for (let c = 0; c < numCols; c += 1) {
    const tasksCol = TASKS_COL_FOR_MY_TASKS_FORMAT[c];
    const myCol = c + 1;
    tasksSheet.getRange(1, tasksCol, 1, tasksCol).copyTo(
      myTasksSheet.getRange(1, myCol, 1, myCol),
      pasteFormat,
      false
    );
    tasksSheet.getRange(2, tasksCol, 2, tasksCol).copyTo(
      myTasksSheet.getRange(2, myCol, 2, myCol),
      pasteFormat,
      false
    );
    myTasksSheet.setColumnWidth(myCol, tasksSheet.getColumnWidth(tasksCol));
  }
  myTasksSheet.setRowHeight(1, tasksSheet.getRowHeight(1));
  myTasksSheet.setRowHeight(2, tasksSheet.getRowHeight(2));
  myTasksSheet.setFrozenRows(1);
  myTasksSheet.hideColumns(1);
}

/**
 * Zwraca arkusz „Zadania - [employeeName]”; tworzy go, jesli nie istnieje.
 * Przy tworzeniu kopiuje formatowanie z arkusza Zadania i ukrywa kolumne A (task_id).
 */
function getOrCreateMyTasksSheetForEmployee_(employeeName) {
  const name = normalizeText_(employeeName);
  if (!name) {
    return { sheet: null, created: false };
  }
  const sheetName = MY_TASKS_SHEET_PREFIX + name;
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(sheetName);
  let created = false;
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    formatNewMyTasksSheetFromTasks_(sheet);
    created = true;
  }
  return { sheet, created };
}

function refreshMyTasksViewForEmployeeName_(employeeName) {
  const selectedEmployeeName = normalizeText_(employeeName);
  if (!selectedEmployeeName) {
    return;
  }
  const { sheet: myTasksSheet, created } =
    getOrCreateMyTasksSheetForEmployee_(selectedEmployeeName);
  if (!myTasksSheet) {
    return;
  }
  writeMyTasksViewToSheet_(myTasksSheet, selectedEmployeeName, {
    applyFullFormat: created,
  });
}

/**
 * Odswieza arkusze Zadania - X dla pracownikow aktywnych (checkbox „aktywny” w Pracownicy).
 * Tylko oni maja zadania przydzielane i tylko dla nich tworzone/odswiezane sa arkusze Zadania - X.
 * Wywolywane z menu (procedura 4) oraz po wygenerowaniu zadan.
 */
function refreshAllMyTasksViews() {
  const employees = getObjectRows_(SHEET_NAMES.EMPLOYEES).filter(
    (row) =>
      normalizeText_(row.pracownik || row.employee_id) &&
      toBoolean_(row.aktywny, true)
  );
  employees.forEach((row) => {
    const name = normalizeText_(row.pracownik || row.employee_id);
    if (name) {
      try {
        refreshMyTasksViewForEmployeeName_(name);
      } catch (e) {
        // Pomin bledy dla pojedynczego pracownika.
      }
    }
  });
}

/**
 * Zapisuje widok zadan pracownika do arkusza Zadania - X (naglowek, wiersze, walidacja, opcjonalnie formatowanie).
 * applyFullFormat: true tylko przy tworzeniu nowego arkusza – kopiuje format z Zadania; przy odswiezeniu tylko dane.
 */
function writeMyTasksViewToSheet_(sheet, employeeName, options) {
  const selectedEmployeeName = normalizeText_(employeeName);
  if (!selectedEmployeeName) {
    return;
  }
  const applyFullFormat = options && options.applyFullFormat === true;
  const preserveFormat = !applyFullFormat;
  const relationNotesByKey = buildClientProcedureNotesByKey_(
    getObjectRows_(SHEET_NAMES.CLIENT_PROCEDURES)
  );
  const taskSheet = getSheetOrThrow_(SHEET_NAMES.TASKS);
  const lastTaskRow = taskSheet.getLastRow();
  const taskRows =
    lastTaskRow < 2
      ? []
      : (() => {
          const values = taskSheet
            .getRange(1, 1, lastTaskRow, HEADERS.TASKS.length)
            .getValues();
          const headers = values[0].map((h) => String(h || '').trim());
          return values
            .slice(1)
            .filter((row) =>
              row.some((cell) => cell !== '' && cell !== null)
            )
            .map((row) => {
              const obj = {};
              headers.forEach((header, idx) => {
                obj[header] = row[idx];
                const normalizedHeader = normalizeLookupKey_(header);
                if (
                  normalizedHeader &&
                  typeof obj[normalizedHeader] === 'undefined'
                ) {
                  obj[normalizedHeader] = row[idx];
                }
              });
              return obj;
            });
        })();
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

  sheet.getRange(1, 1, 1, HEADERS.MY_TASKS.length).setValues([HEADERS.MY_TASKS]);
  clearSheetBody_(sheet, HEADERS.MY_TASKS.length, preserveFormat);

  if (openTasks.length === 0) {
    if (sheet.getSheetName().startsWith(MY_TASKS_SHEET_PREFIX)) {
      SpreadsheetApp.getActiveSpreadsheet().deleteSheet(sheet);
    } else {
      sheet.getRange(2, 1).setValue('Brak otwartych zadan.');
      if (applyFullFormat) {
        applyMyTasksBodyFormatFromTasks_(sheet);
      }
    }
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

  ensureSheetSize_(sheet, rows.length + 1, HEADERS.MY_TASKS.length);
  sheet.getRange(2, 1, rows.length, HEADERS.MY_TASKS.length).setValues(rows);

  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList([STATUS.NEW, STATUS.IN_PROGRESS, STATUS.DONE], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, MY_TASKS_COL.STATUS, rows.length, 1).setDataValidation(statusRule);
  sheet.getRange(2, MY_TASKS_COL.DUE_DATE, rows.length, 1).setNumberFormat('yyyy-mm-dd');

  if (applyFullFormat) {
    applyMyTasksBodyFormatFromTasks_(sheet);
  }

  const today = normalizeDate_(new Date());
  const todayKey = formatDateKey_(today);
  rows.forEach((row, idx) => {
    const dueDate = toDate_(row[MY_TASKS_COL.DUE_DATE - 1]);
    if (!dueDate) {
      return;
    }
    const dueKey = formatDateKey_(dueDate);
    const rowRange = sheet.getRange(idx + 2, 1, idx + 2, HEADERS.MY_TASKS.length);
    if (dueKey < todayKey) {
      rowRange.setBackground('#fde7e9');
    } else if (dueKey === todayKey) {
      rowRange.setBackground('#d4edda');
    }
  });
}

/**
 * Kopiuje formatowanie wierszy danych z Zadania do arkusza Zadania - X kolumna po kolumnie
 * (Zadania.due_date→termin, procedura→procedura itd.) oraz szerokosc kolumn i wysokosc wierszy.
 */
function applyMyTasksBodyFormatFromTasks_(myTasksSheet) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const tasksSheet = spreadsheet.getSheetByName(SHEET_NAMES.TASKS);
  if (!tasksSheet) {
    return;
  }
  const lastRow = myTasksSheet.getLastRow();
  if (lastRow < 2) {
    return;
  }
  const pasteFormat = SpreadsheetApp.CopyPasteType.PASTE_FORMAT;
  const numCols = TASKS_COL_FOR_MY_TASKS_FORMAT.length;

  for (let c = 0; c < numCols; c += 1) {
    const tasksCol = TASKS_COL_FOR_MY_TASKS_FORMAT[c];
    const myCol = c + 1;
    tasksSheet
      .getRange(2, tasksCol, 2, tasksCol)
      .copyTo(
        myTasksSheet.getRange(2, myCol, lastRow, myCol),
        pasteFormat,
        false
      );
    myTasksSheet.setColumnWidth(myCol, tasksSheet.getColumnWidth(tasksCol));
  }
  const bodyRowHeight = tasksSheet.getRowHeight(2);
  for (let r = 2; r <= lastRow; r += 1) {
    myTasksSheet.setRowHeight(r, bodyRowHeight);
  }
}

/**
 * Ustawia arkusz Zadania - X w stan pusty: naglowek + „Brak otwartych zadan.”
 * Uzywane przez procedure awaryjna czyszczenia widokow.
 */
function clearMyTasksViewToEmpty_(sheet) {
  const maxCols = HEADERS.MY_TASKS.length;
  sheet.getRange(1, 1, 1, maxCols).setValues([HEADERS.MY_TASKS]);
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const range = sheet.getRange(2, 1, lastRow, maxCols);
    range.clearContent().clearFormat();
    range.setDataValidation(null);
  }
  sheet.getRange(2, 1).setValue('Brak otwartych zadan.');
}

/**
 * Procedura awaryjna: wymusza usuniecie zadan z widokow Zadania - X dla wszystkich uzytkownikow.
 * Czyści wszystkie arkusze „Zadania - [Imię Nazwisko]”.
 * Jesli w trakcie operacji powstaly dodatkowe arkusze (np. tymczasowe), zostana usuniete po zakonczeniu.
 */
function emergencyClearAllMyTasksViews_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  /** Arkusze utworzone w tej procedurze (np. tymczasowe); do usuniecia po zakonczeniu. */
  const sheetsToRemove = [];

  try {
    spreadsheet.getSheets().forEach((sheet) => {
      const name = sheet.getSheetName();
      if (name.startsWith(MY_TASKS_SHEET_PREFIX)) {
        clearMyTasksViewToEmpty_(sheet);
      }
    });
  } finally {
    sheetsToRemove.forEach((sheet) => {
      try {
        spreadsheet.deleteSheet(sheet);
      } catch (e) {
        // Ignoruj (np. arkusz juz usuniety).
      }
    });
  }
}

/**
 * Procedura awaryjna (menu): wymusza wyczyszczenie widokow Zadania - X u wszystkich uzytkownikow.
 */
function runEmergencyClearAllMyTasksViews() {
  emergencyClearAllMyTasksViews_();
  SpreadsheetApp.getUi().alert('Wyczyszczono widoki Zadania - X u wszystkich pracownikow.');
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

  const allTasks = getObjectRows_(SHEET_NAMES.TASKS).map((row) => ({
    klient: normalizeText_(row.klient || row.client_id),
    procedura: normalizeText_(row.procedura || row.procedure_id),
    due_date: toDate_(row.due_date),
    pracownik: normalizeText_(row.pracownik || row.employee_id),
    status: normalizeText_(row.status),
  }));

  const statuses = relationRows.map((row) => {
    const clientKey = normalizeLookupKey_(row.klient || row.client_id);
    const procedureKey = normalizeLookupKey_(row.procedura || row.procedure_id);
    const tasksForRelation = allTasks.filter(
      (t) =>
        normalizeLookupKey_(t.klient) === clientKey &&
        normalizeLookupKey_(t.procedura) === procedureKey &&
        t.due_date &&
        normalizeText_(t.status) !== STATUS.DONE
    );
    if (tasksForRelation.length === 0) {
      return ['Brak zadań'];
    }
    const unassigned = tasksForRelation.some((t) => !t.pracownik);
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

function isCurrentUserManager_() {
  const employee = resolveCurrentEmployee_();
  return employee && normalizeLookupKey_(employee.role) === 'manager';
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

/**
 * Ukrywa arkusze z listy SHEETS_VISIBLE_ONLY_TO_MANAGER dla uzytkownikow bez roli manager.
 * Dla managera wszystkie arkusze sa widoczne. Wywolywane przy onOpen.
 * Uwaga: osoby z uprawnieniami edycji moga odsłonic arkusze recznie (Widok / Ukryte arkusze) – to ulatwienie, nie zabezpieczenie.
 */
function applySheetVisibilityByRole_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const isManager = isCurrentUserManager_();
  const managerOnlySet = new Set(SHEETS_VISIBLE_ONLY_TO_MANAGER);
  const currentEmployee = resolveCurrentEmployee_();
  const mySheetNameForCurrentUser =
    currentEmployee ? MY_TASKS_SHEET_PREFIX + currentEmployee.employeeName : null;

  spreadsheet.getSheets().forEach((sheet) => {
    const name = sheet.getName();
    try {
      if (managerOnlySet.has(name)) {
        if (isManager) {
          sheet.showSheet();
        } else {
          sheet.hideSheet();
        }
        return;
      }
      if (name.startsWith(MY_TASKS_SHEET_PREFIX)) {
        if (isManager) {
          sheet.showSheet();
        } else if (name === mySheetNameForCurrentUser) {
          sheet.showSheet();
        } else {
          sheet.hideSheet();
        }
      }
    } catch (e) {
      // Ignoruj bledy (np. brak uprawnien do ukrywania).
    }
  });
}

/**
 * Dla panelu Klienci: zwraca nazwe klienta z aktualnie zaznaczonego wiersza na arkuszu Klienci.
 * Gdy aktywny arkusz to nie Klienci lub zaznaczony wiersz to naglowek – zwraca null.
 */
function getSelectedClientFromKlienciSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getActiveSheet();
  if (!sheet || sheet.getName() !== SHEET_NAMES.CLIENTS) {
    return null;
  }
  const range = spreadsheet.getActiveRange();
  if (!range) {
    return null;
  }
  const row = range.getRow();
  if (row < 2) {
    return null;
  }
  const clientName = normalizeText_(sheet.getRange(row, 1).getValue());
  if (!clientName) {
    return null;
  }
  return { clientName };
}

/**
 * Zwraca zadania dla danego klienta w podanym miesiacu (termin, procedura, pracownik, status), posortowane wg terminu.
 * year, month – liczby (month 1–12).
 */
function getTasksForClientInMonth(clientName, year, month) {
  if (!clientName || !year || !month) {
    return [];
  }
  const clientKey = normalizeLookupKey_(clientName);
  const firstDay = new Date(year, month - 1, 1);
  const lastDay = new Date(year, month, 0);
  lastDay.setHours(23, 59, 59, 999);
  const firstTs = firstDay.getTime();
  const lastTs = lastDay.getTime();

  const taskRows = getObjectRows_(SHEET_NAMES.TASKS);
  const tz = Session.getScriptTimeZone();
  const fmt = (d) => (d ? Utilities.formatDate(toDate_(d), tz, 'yyyy-MM-dd') : '');

  const tasks = taskRows
    .filter(
      (row) =>
        normalizeLookupKey_(row.klient || row.client_id) === clientKey
    )
    .map((row) => ({
      dueDate: toDate_(row.due_date),
      procedure: normalizeText_(row.procedura || row.procedure_id) || '',
      employee: normalizeText_(row.pracownik || row.employee_id) || '',
      status: normalizeText_(row.status) || '',
    }))
    .filter((row) => row.dueDate && row.dueDate.getTime() >= firstTs && row.dueDate.getTime() <= lastTs)
    .sort((a, b) => a.dueDate.getTime() - b.dueDate.getTime())
    .map((row) => ({
      dueDate: fmt(row.dueDate),
      procedure: row.procedure,
      employee: row.employee,
      status: row.status,
    }));

  return tasks;
}
