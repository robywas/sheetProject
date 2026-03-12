function onOpen() {
  try {
    applySheetVisibilityByRole_();
  } catch (e) {
    // Widocznosc arkuszy nie moze blokowac menu.
  }
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
    .addSeparator()
    .addItem('Panel pracownika', 'openWorkerSidebar')
    .addItem('Panel managera', 'openManagerSidebar')
    .addToUi();
}

function onInstall() {
  onOpen();
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
