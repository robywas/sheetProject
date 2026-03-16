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
    .addItem('Instrukcja: auto-otwieranie panelu Klienci', 'showClientPanelTriggerInstructions')
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
    .setWidth(600)
    .setHeight(520);
  SpreadsheetApp.getUi().showModelessDialog(html, 'Zadania klienta');
}

/** Klucz wlasciwosci: czy panel Klienci jest otwarty (zapisywane z HTML przy otwarciu/zamknieciu). */
const CLIENT_PANEL_OPEN_KEY = 'clientPanelOpen';

/** Ustawia/czyści flagę otwartego panelu Klienci (wywolywane z ClientPanel.html). */
function setClientPanelOpen(isOpen) {
  const props = PropertiesService.getScriptProperties();
  if (isOpen) {
    props.setProperty(CLIENT_PANEL_OPEN_KEY, '1');
  } else {
    props.deleteProperty(CLIENT_PANEL_OPEN_KEY);
  }
}

/**
 * Wywolywane przez trigger „Przy zmianie zaznaczenia” (dodaj recznie w Wykonaj > Triggers).
 * Otwiera Panel Klienci przy zaznaczeniu wiersza z danymi na arkuszu Klienci (tylko jesli panel nie jest juz otwarty).
 * Jesli trigger sie nie wykonuje, otwieraj panel z menu: Procedury > Panel Klienci –
 * okno jest modeless, mozesz zmieniac klienta bez zamykania.
 */
function onSelectionChange(e) {
  try {
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
    if (PropertiesService.getScriptProperties().getProperty(CLIENT_PANEL_OPEN_KEY) === '1') {
      return;
    }
    const ui = (e.source && e.source.getUi) ? e.source.getUi() : SpreadsheetApp.getUi();
    const html = HtmlService.createHtmlOutputFromFile('ClientPanel')
      .setWidth(600)
      .setHeight(520);
    ui.showModelessDialog(html, 'Zadania klienta');
  } catch (err) {
    // Prosty trigger nie moze otwierac UI – wtedy otworz panel z menu
  }
}

/**
 * Pokazuje instrukcje dodania triggera „Przy zmianie zaznaczenia”.
 * Triggeru nie da sie utworzyc z kodu – trzeba go dodac recznie w edytorze Apps Script.
 */
function showClientPanelTriggerInstructions() {
  const msg =
    'Auto-otwieranie panelu przy zaznaczeniu klienta wymaga triggera.\n\n' +
    '1. Otworz edytor Apps Script (Rozszerzenia > Skrypty edytora Apps Script).\n' +
    '2. Po lewej kliknij ikone Zegara (Wykonaj / Triggers).\n' +
    '3. Kliknij „Dodaj trigger” (w prawym dolnym rogu).\n' +
    '4. Funkcja: onSelectionChange\n' +
    '5. Zdarzenie: Z arkusza > Przy zmianie zaznaczenia\n' +
    '6. Zapisz.\n\n' +
    'Po autoryzacji zaznaczenie wiersza na arkuszu Klienci bedzie otwierac panel.';
  SpreadsheetApp.getUi().alert('Trigger – instrukcja', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Ikona przycisku „Panel Klienci” – tylko base64, bez pobierania w skrypcie (brak blokady antywirusa).
 * Aby uzyc wlasnej ikony: pobierz PNG 24x24 (np. list/view) z:
 *   https://icons8.com/icons/set/list  (wybierz ikone → PNG 24px → Pobierz)
 *   lub https://fonts.google.com/icons (Material Icons → wybierz → Download PNG)
 * Potem przekonwertuj plik na base64 (np. https://www.base64-image.de) i wklej ponizej.
 */
var CLIENT_PANEL_BUTTON_ICON_BASE64 = 'iVBORw0KGgoAAAANSUhEUgAAADAAAAAwCAYAAABXAvmHAAAACXBIWXMAAAsTAAALEwEAmpwYAAAAX0lEQVR4nO3VsQnAMBAEwW/lmlajX4CdGicGgUDCM6BI0QbHVwFsKaOv5/v6X/1KwIuAsTgA/i6njzgCWkDcAZh3/IgjoAXEHYB5x484AlpA3AGYd/yII6AFxB0Aanc3L0Cm26Ko+hkAAAAASUVORK5CYII=';

/**
 * Dodaje przycisk „Panel Klienci” w naglowku arkusza Klienci (komorka B1).
 * Zapewnia 2 kolumny, wstawia ikone w B1. Klik uruchamia openClientPanel. Wywolywane z setupWorkbook.
 */
function ensureClientPanelButtonOnKlienciSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    const sheet = getSheetOrThrow_(SHEET_NAMES.CLIENTS);
    const images = sheet.getImages();
    for (let i = images.length - 1; i >= 0; i--) {
      const img = images[i];
      try {
        if (img.getScript() === 'openClientPanel' && img.getAnchorCell().getRow() === 1) {
          img.remove();
        }
      } catch (e) {
        // getScript() moze rzucac jesli nie przypisano
      }
    }
    ensureSheetSize_(sheet, sheet.getMaxRows(), 2);
    const col = 2;
    const row = 1;
    const blob = Utilities.newBlob(
      Utilities.base64Decode(CLIENT_PANEL_BUTTON_ICON_BASE64),
      'image/png'
    );
    const over = sheet.insertImage(blob, col, row);
    over.assignScript('openClientPanel');
    over.setWidth(24);
    over.setHeight(24);
    over.setAltTextTitle('Panel Klienci');
    over.setAltTextDescription('Kliknij, aby otworzyc panel zadan wybranego klienta.');
    ss.toast('Ikona Panel Klienci dodana w B1.', 'Przycisk', 4);
  } catch (err) {
    const msg = (err && err.message) ? err.message : String(err);
    ss.toast('Blad przycisku: ' + msg, 'Panel Klienci', 8);
  }
}
