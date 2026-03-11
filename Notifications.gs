/**
 * Powiadomienia email: zadania na dzis (ostatni dzien) oraz opoznione.
 * Uruchom z menu Procedury lub ustaw trigger czasowy (np. codziennie rano).
 */

function sendTaskReminderEmails() {
  const today = normalizeDate_(new Date());
  const tasksByEmployee = getTasksDueTodayOrOverdueByEmployee_(today);
  const employeeToEmail = getEmployeeEmailMap_();

  let sentCount = 0;
  let skipCount = 0;

  Object.keys(tasksByEmployee).forEach((employeeName) => {
    const email = employeeToEmail[normalizeLookupKey_(employeeName)];
    if (!email) {
      skipCount += 1;
      return;
    }
    const tasks = tasksByEmployee[employeeName];
    const overdue = tasks.filter((t) => t.dueDate < today);
    const dueToday = tasks.filter((t) => t.dueDate.getTime() === today.getTime());

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
      // Pominieto (np. nieprawidlowy adres) – mozna zalogowac err
    }
  });

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Wyslano ' + sentCount + ' wiadomosci.' + (skipCount > 0 ? ' Pominieto ' + skipCount + ' (brak email).' : ''),
    'Powiadomienia',
    5
  );
}

function getTasksDueTodayOrOverdueByEmployee_(today) {
  const taskRows = getObjectRows_(SHEET_NAMES.TASKS);
  const byEmployee = {};

  taskRows.forEach((row) => {
    const status = normalizeText_(row.status).toUpperCase();
    if (status === STATUS.DONE) {
      return;
    }
    const dueDate = toDate_(row.due_date);
    if (!dueDate || dueDate > today) {
      return;
    }

    const employeeName = normalizeText_(row.pracownik || row.employee_id);
    if (!employeeName) {
      return;
    }

    const key = normalizeLookupKey_(employeeName);
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
  const rows = getObjectRows_(SHEET_NAMES.EMPLOYEES);
  const map = {};
  rows.forEach((row) => {
    const name = normalizeText_(row.pracownik || row.employee_id);
    const email = normalizeText_(row.email);
    if (name && email) {
      map[normalizeLookupKey_(name)] = email;
    }
  });
  return map;
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
