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
    .addItem('9) Pokaz wszystkie arkusze', 'showAllSheets')
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

/** Odkrywa wszystkie arkusze (przywraca widok po otwarciu pliku przez pracownika w innej sesji). */
function showAllSheets() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  spreadsheet.getSheets().forEach((sheet) => {
    try {
      sheet.showSheet();
    } catch (e) {}
  });
  SpreadsheetApp.getActiveSpreadsheet().toast('Wszystkie zakładki są widoczne.', 'Procedury', 3);
}
