function onOpen() {
  try {
    var openerEmail = Session.getEffectiveUser().getEmail();
    if (openerEmail) {
      PropertiesService.getDocumentProperties().setProperty('lastOpenerEmail', openerEmail);
    }
  } catch (e) {}
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
    .addSeparator()
    .addItem('8) Zainstaluj auto-otwieranie panelu', 'installPanelOnOpenTrigger')
    .addToUi();
}

/**
 * Wywolywane przez instalowalny trigger przy otwarciu skoroszytu (prosty onOpen nie moze otwierac sidebara).
 * Trigger dziala w kontekście instalatora, wiec rola jest ustalana po emailu osoby otwierajacej (zapisany w prostym onOpen).
 */
function openPanelOnOpen() {
  var openerEmail = (PropertiesService.getDocumentProperties().getProperty('lastOpenerEmail') || getCurrentUserEmail_() || '').toLowerCase();
  try {
    applySheetVisibilityByRole_(openerEmail);
  } catch (e) {}
  try {
    if (isManagerByEmail_(openerEmail)) {
      openManagerSidebar();
    } else {
      openWorkerSidebar();
    }
  } catch (e) {}
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

/**
 * Tworzy instalowalny trigger "On open", ktory przy kazdym otwarciu uruchamia openPanelOnOpen (widocznosc arkuszy + otwarcie sidebara).
 * Uruchom raz z menu (np. jako manager); przy pierwszym uruchomieniu zatwierdz uprawnienia. Od tego momentu panel bedzie sie otwieral przy otwarciu skoroszytu.
 */
function installPanelOnOpenTrigger() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.getProjectTriggers().forEach((trigger) => {
    if (
      trigger.getHandlerFunction() === 'openPanelOnOpen' &&
      trigger.getEventType() === ScriptApp.EventType.ON_OPEN
    ) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  ScriptApp.newTrigger('openPanelOnOpen')
    .forSpreadsheet(spreadsheet)
    .onOpen()
    .create();
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Auto-otwieranie panelu wlaczone. Przy nastepnym otwarciu skoroszytu otworzy sie odpowiedni panel.',
    'Procedury',
    8
  );
}
