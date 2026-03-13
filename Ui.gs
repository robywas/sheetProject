function onOpen() {
  // Menu musi byc dodane najpierw – inaczej przy bledzie dialogu (np. u pracownika) menu by sie nie pojawilo.
  SpreadsheetApp.getUi()
    .createMenu('Procedury')
    .addItem('1) Utworz/odswiez strukture', 'setupWorkbook')
    .addItem('2) Dodaj dane przykladowe', 'seedSampleData')
    .addSeparator()
    .addItem('3) Wygeneruj zadania (30 dni)', 'generateTasks30Days')
    .addItem('4) Odswiez Zadania - X (wszyscy pracownicy)', 'refreshAllMyTasksViews')
    .addItem('5) Odswiez kontrole (Klienci_Procedury)', 'refreshClientProceduresControl')
    .addItem('6) Wyslij powiadomienia email (termin / opoznienia)', 'sendTaskReminderEmails')
    .addSeparator()
    .addItem('Awaryjnie: wyczysc Zadania - X (wszyscy)', 'runEmergencyClearAllMyTasksViews')
    .addSeparator()
    .addItem('Panel pracownika', 'openWorkerSidebar')
    .addItem('Panel managera', 'openManagerSidebar')
    .addItem('Panel Klienci', 'openClientPanel')
    .addToUi();

  // Przy kazdym otwarciu odswiez Zadania - X dla biezacego uzytkownika. Wywolanie przez dialog,
  // bo w prostym triggerze onOpen getActiveUser() bywa pusty; dialog laduje w kontekście otwierajacego.
  try {
    const html = HtmlService.createHtmlOutputFromFile('RefreshMyTasksDialog')
      .setWidth(260)
      .setHeight(80);
    SpreadsheetApp.getUi().showModalDialog(html, 'Zadania');
  } catch (e) {
    // Nie blokuj; uzytkownik moze wybrac z menu „Odswiez Zadania - X (wszyscy pracownicy)” lub procedury 4.
  }
}

function onInstall() {
  onOpen();
}

function refreshAllViews() {
  refreshAllMyTasksViews();
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

function openClientPanel() {
  const html = HtmlService.createHtmlOutputFromFile('ClientPanel')
    .setTitle('Zadania klienta')
    .setWidth(760);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Wywolywane przez trigger „Przy zmianie zaznaczenia”. Otwiera Panel Klienci,
 * gdy uzytkownik zaznaczy wiersz z danymi na arkuszu Klienci.
 * Trigger: Zdarzenia arkusza > Przy zmianie zaznaczenia > onSelectionChange
 */
function onSelectionChange(e) {
  if (!e || !e.range) {
    return;
  }
  const sheet = e.range.getSheet();
  if (sheet.getName() !== SHEET_NAMES.CLIENTS) {
    return;
  }
  if (e.range.getRow() < 2) {
    return;
  }
  openClientPanel();
}
