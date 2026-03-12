function onOpen() {
  // Nie ustawiamy widocznosci arkuszy przy otwarciu: w Google Sheets jest ona globalna (dla calego dokumentu).
  // Gdy pracownik otworzy arkusz, ukrycie zakladek zmienialoby widok tez u managera w innej przegladarce.
  SpreadsheetApp.getUi()
    .createMenu('Procedury')
    .addItem('1) Utworz/odswiez strukture', 'setupWorkbook')
    .addItem('2) Dodaj dane przykladowe', 'seedSampleData')
    .addSeparator()
    .addItem('3) Wygeneruj zadania (30 dni)', 'generateTasks30Days')
    .addItem('4) Odswiez moje zadania', 'refreshMyTasksView')
    .addItem('5) Odswiez dashboard managera', 'refreshManagerDashboard')
    .addItem('6) Odswiez kontrole (Klienci_Procedury)', 'refreshClientProceduresControl')
    .addItem('7) Wyslij powiadomienia email (termin / opoznienia)', 'sendTaskReminderEmails')
    .addItem('8) Odswiez Moje_zadania wszystkich pracownikow (manager)', 'refreshAllMyTasksViewsForManager')
    .addItem('9) Oznacz koniecznosc odswiezenia u pracownikow (manager)', 'setAllEmployeesRequireRefresh')
    .addSeparator()
    .addItem('Panel pracownika', 'openWorkerSidebar')
    .addItem('Panel managera', 'openManagerSidebar')
    .addToUi();
}

function onInstall() {
  onOpen();
}

/**
 * Prosty trigger: przy zmianie zaznaczenia (w tym przejsciu na zakladke). Gdy uzytkownik jest na Moje_zadania:
 * – jesli w Pracownicy ma zaznaczone wymaga_odswiezenia, odswieza widok i odznacza;
 * – w przeciwnym razie odswieza co MY_TASKS_AUTO_REFRESH_MINUTES (fallback).
 */
function onSelectionChange(e) {
  if (!e || !e.range) {
    return;
  }
  const sheetName = e.range.getSheet().getName();
  if (sheetName !== SHEET_NAMES.MY_TASKS) {
    return;
  }
  const currentUser = getCurrentUserEmail_();

  if (getEmployeeRequiresRefresh_()) {
    try {
      refreshMyTasksView();
      clearEmployeeRequiresRefresh_(currentUser);
    } catch (err) {}
    return;
  }

  const props = PropertiesService.getDocumentProperties();
  const now = Date.now();
  const lastRefresh = parseInt(props.getProperty('myTasksLastRefresh') || '0', 10);
  const lastUser = props.getProperty('myTasksLastUser') || '';
  const intervalMs = MY_TASKS_AUTO_REFRESH_MINUTES * 60 * 1000;
  if (lastUser === currentUser && now - lastRefresh < intervalMs) {
    return;
  }
  try {
    refreshMyTasksView();
  } catch (err) {
    return;
  }
  props.setProperty('myTasksLastRefresh', String(now));
  props.setProperty('myTasksLastUser', currentUser);
}

function refreshAllViews() {
  refreshManagerDashboard();
  try {
    refreshMyTasksView();
  } catch (error) {
    // Widok pracownika jest zalezny od mapowania email -> pracownik.
  }
}

function openWorkerSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('WorkerSidebar')
    .setTitle('Panel pracownika')
    .setWidth(340);
  SpreadsheetApp.getUi().showSidebar(html);
}

function openManagerSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('ManagerSidebar')
    .setTitle('Panel managera')
    .setWidth(340);
  SpreadsheetApp.getUi().showSidebar(html);
}
