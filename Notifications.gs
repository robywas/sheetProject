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
    const email = employeeToEmail[employeeLookupKey_(employeeName)];
    if (!email) {
      skipCount += 1;
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
  const taskRows = getObjectRows_(SHEET_NAMES.TASKS);
  const byEmployee = {};

  taskRows.forEach((row) => {
    const status = normalizeText_(row.status).toUpperCase();
    if (status === STATUS.DONE) {
      return;
    }
    const dueDate = toDate_(row.due_date);
    if (!dueDate) {
      return;
    }
    const dueKey = formatDateKey_(dueDate);
    if (dueKey > todayKey) {
      return;
    }

    const employeeName = normalizeText_(row.pracownik || row.employee_id);
    if (!employeeName) {
      return;
    }

    const key = employeeLookupKey_(employeeName);
    if (!byEmployee[key]) {
      byEmployee[key] = [];
    }
    byEmployee[key].push({
      clientName: normalizeText_(row.klient || row.client_id),
      procedureName: normalizeText_(row.procedura || row.procedure_id),
      dueDate,
      employeeName,
    });
  });

  // Zwracamy po nazwie (pierwsza z listy), zeby w wiadomosci bylo czytelne
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
      map[employeeLookupKey_(name)] = email;
    }
  }
  return map;
}

function employeeLookupKey_(name) {
  return normalizeText_(name || '').toLowerCase().replace(/\s+/g, ' ').trim();
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
