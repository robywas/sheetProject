/**
 * Powiadomienia email: zadania na dzis (ostatni dzien) oraz opoznione.
 * Uruchom z menu Procedury lub ustaw trigger czasowy (np. codziennie rano).
 */

function sendTaskReminderEmails() {
  const today = normalizeDate_(new Date());
  const todayKey = formatDateKey_(today);
  const tasksByEmployee = getTasksDueTodayOrOverdueByEmployee_(todayKey);
  const employeeToEmail = getEmployeeEmailMap_();

  let sentCount = 0;
  let skipCount = 0;
  let totalTasks = 0;

  Object.keys(tasksByEmployee).forEach((employeeName) => {
    const tasks = tasksByEmployee[employeeName];
    totalTasks += tasks.length;
    const key = employeeLookupKey_(employeeName);
    const email =
      employeeToEmail[key] ||
      employeeToEmail[normalizeText_(employeeName)] ||
      employeeToEmail[employeeName];
    if (!email) {
      skipCount += 1;
      writeEmailDiagnostic_(employeeToEmail, tasksByEmployee, key, employeeName);
      return;
    }
    const overdue = tasks.filter((t) => formatDateKey_(t.dueDate) < todayKey);
    const dueToday = tasks.filter((t) => formatDateKey_(t.dueDate) === todayKey);

    const subject =
      'Procedury: zadania na dzis / opoznione (' +
      (overdue.length > 0 ? overdue.length + ' opoznione, ' : '') +
      dueToday.length +
      ' na dzis)';
    const body = buildReminderEmailBody_(overdue, dueToday, today);

    try {
      MailApp.sendEmail(email, subject, body);
      sentCount += 1;
    } catch (err) {
      skipCount += 1;
    }
  });

  let msg = 'Wyslano ' + sentCount + ' wiadomosci.';
  if (totalTasks === 0) {
    msg = 'Brak zadan na dzis ani opoznionych (otwartych).';
  } else if (skipCount > 0) {
    msg += ' Pominieto ' + skipCount + ' (brak email w Pracownicy).';
  }
  SpreadsheetApp.getActiveSpreadsheet().toast(msg, 'Powiadomienia', 5);
}

function getTasksDueTodayOrOverdueByEmployee_(todayKey) {
  const sheet = getSheetOrThrow_(SHEET_NAMES.TASKS);
  const values = sheet.getDataRange().getValues();
  const byEmployee = {};
  if (values.length < 2) {
    return {};
  }
  const headerRow = values[0].map((h) => String(h || '').trim());
  const colStatus = headerRow.findIndex((h) => /status/i.test(h));
  const colDue = headerRow.findIndex((h) => /due_date|termin/i.test(h));
  const colPracownik = headerRow.findIndex((h) => /pracownik|employee/i.test(h));
  const colKlient = headerRow.findIndex((h) => /klient|client/i.test(h));
  const colProcedura = headerRow.findIndex((h) => /procedura|procedure/i.test(h));
  if (colStatus === -1 || colDue === -1 || colPracownik === -1) {
    return {};
  }

  for (let r = 1; r < values.length; r += 1) {
    const row = values[r];
    const status = normalizeText_(row[colStatus]).toUpperCase();
    if (status === STATUS.DONE) {
      continue;
    }
    const dueDate = toDate_(row[colDue]);
    if (!dueDate) {
      continue;
    }
    const dueKey = formatDateKey_(dueDate);
    if (dueKey > todayKey) {
      continue;
    }
    const employeeName = normalizeText_(row[colPracownik]);
    if (!employeeName) {
      continue;
    }

    const key = employeeLookupKey_(employeeName);
    if (!byEmployee[key]) {
      byEmployee[key] = [];
    }
    byEmployee[key].push({
      clientName: colKlient >= 0 ? normalizeText_(row[colKlient]) : '',
      procedureName: colProcedura >= 0 ? normalizeText_(row[colProcedura]) : '',
      dueDate,
      employeeName,
    });
  });

  const byEmployeeName = {};
  Object.keys(byEmployee).forEach((key) => {
    const list = byEmployee[key];
    const name = list[0] ? list[0].employeeName : key;
    byEmployeeName[name] = list;
  });
  return byEmployeeName;
}

function getEmployeeEmailMap_() {
  const sheet = getSheetOrThrow_(SHEET_NAMES.EMPLOYEES);
  const values = sheet.getDataRange().getValues();
  const map = {};
  if (values.length < 2) {
    return map;
  }
  const headerRow = values[0].map((h) => String(h || '').trim());
  const colPracownik = headerRow.findIndex((h) => /pracownik/i.test(h));
  const colEmail = headerRow.findIndex((h) => /^e-?mail$/i.test(h));
  if (colPracownik === -1 || colEmail === -1) {
    return map;
  }
  for (let r = 1; r < values.length; r += 1) {
    const row = values[r];
    const name = normalizeText_(row[colPracownik]);
    const email = normalizeText_(row[colEmail]);
    if (name && email && email.indexOf('@') !== -1) {
      const key = employeeLookupKey_(name);
      map[key] = email;
      if (key !== name) {
        map[name] = email;
      }
    }
  }
  return map;
}

function employeeLookupKey_(name) {
  return normalizeText_(name || '').toLowerCase().replace(/\s+/g, ' ').trim();
}

function writeEmailDiagnostic_(map, tasksByEmployee, searchedKey, searchedName) {
  try {
    const sheet = getSheetOrThrow_(SHEET_NAMES.MANAGER_DASHBOARD);
    const mapKeys = Object.keys(map).join(', ') || '(pusta)';
    const taskNames = Object.keys(tasksByEmployee).join(', ') || '(brak)';
    sheet.getRange('A1').setValue(
      'Diagnoza email: Mapa klucze=[' +
        mapKeys +
        '] Szukany klucz="' +
        searchedKey +
        '" nazwa="' +
        searchedName +
        '" Pracownicy z zadan=[' +
        taskNames +
        ']'
    );
  } catch (e) {
    // ignoruj blad zapisu diagnostyki
  }
}

function buildReminderEmailBody_(overdue, dueToday, today) {
  const tz = Session.getScriptTimeZone();
  const fmt = (d) =>
    Utilities.formatDate(normalizeDate_(d), tz, 'yyyy-MM-dd');

  let body = 'Zadania wymagajace realizacji:\n\n';

  if (overdue.length > 0) {
    body += '--- OPOZNIONE ---\n';
    overdue.forEach((t) => {
      body += '  ' + fmt(t.dueDate) + ' | ' + (t.clientName || '') + ' | ' + (t.procedureName || '') + '\n';
    });
    body += '\n';
  }

  if (dueToday.length > 0) {
    body += '--- OSTATNI DZIEN (dzis) ---\n';
    dueToday.forEach((t) => {
      body += '  ' + fmt(t.dueDate) + ' | ' + (t.clientName || '') + ' | ' + (t.procedureName || '') + '\n';
    });
  }

  body += '\n-- Arkusz Procedury';
  return body;
}
