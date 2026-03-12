function getSheetOrThrow_(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(
      'Brak arkusza "' +
        sheetName +
        '". Uruchom najpierw menu Procedury > 1) Utworz/odswiez strukture.'
    );
  }
  return sheet;
}

function getObjectRows_(sheetName) {
  const sheet = getSheetOrThrow_(sheetName);
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    return [];
  }

  const headers = values[0].map((header) => String(header || '').trim());
  return values
    .slice(1)
    .filter((row) => row.some((cell) => cell !== '' && cell !== null))
    .map((row) => {
      const obj = {};
      headers.forEach((header, idx) => {
        obj[header] = row[idx];
        const normalizedHeader = normalizeLookupKey_(header);
        if (normalizedHeader && typeof obj[normalizedHeader] === 'undefined') {
          obj[normalizedHeader] = row[idx];
        }
      });
      return obj;
    });
}

function normalizeDate_(value) {
  const date = new Date(value);
  date.setHours(0, 0, 0, 0);
  return date;
}

function toDate_(value) {
  if (value === '' || value === null || typeof value === 'undefined') {
    return null;
  }

  if (value instanceof Date && !isNaN(value.getTime())) {
    return normalizeDate_(value);
  }

  const parsed = new Date(value);
  if (isNaN(parsed.getTime())) {
    return null;
  }
  return normalizeDate_(parsed);
}

function toNumber_(value, fallback) {
  if (value === '' || value === null || typeof value === 'undefined') {
    return fallback;
  }
  const num = Number(value);
  return Number.isFinite(num) ? num : fallback;
}

function toBoolean_(value, fallback) {
  if (value === '' || value === null || typeof value === 'undefined') {
    return fallback;
  }
  if (typeof value === 'boolean') {
    return value;
  }
  const normalized = String(value).trim().toLowerCase();
  return ['1', 'true', 'tak', 'yes', 'y'].includes(normalized);
}

function normalizeText_(value) {
  return String(value || '').trim();
}

function normalizeLookupKey_(value) {
  return normalizeText_(value).toLowerCase();
}

function formatDateKey_(date) {
  return Utilities.formatDate(
    normalizeDate_(date),
    Session.getScriptTimeZone(),
    'yyyy-MM-dd'
  );
}

/** Klucz miesiaca YYYY-MM do grupowania zadan (jeden pracownik na klienta w ramach miesiaca). */
function getMonthKey_(date) {
  const d = normalizeDate_(date);
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  return y + '-' + m;
}

/** Wielkanoc (niedziela) dla danego roku – kalendarz gregorianski (algorytm Oudin). */
function getEasterSunday_(year) {
  const Y = year;
  const C = Math.floor(Y / 100);
  const N = Y - 19 * Math.floor(Y / 19);
  const K = Math.floor((C - 17) / 25);
  let I = C - Math.floor(C / 4) - Math.floor((C - K) / 3) + 19 * N + 15;
  I = I - 30 * Math.floor(I / 30);
  I =
    I -
    Math.floor(I / 28) *
      (1 -
        Math.floor(I / 28) *
          Math.floor(29 / (I + 1)) *
          Math.floor((21 - N) / 11));
  let J = Y + Math.floor(Y / 4) + I + 2 - C + Math.floor(C / 4);
  J = J - 7 * Math.floor(J / 7);
  const L = I - J;
  const M = 3 + Math.floor((L + 40) / 44);
  const D = L + 28 - 31 * Math.floor(M / 4);
  return normalizeDate_(new Date(Y, M - 1, D));
}

/** Dni wolne (swieta) w Polsce dla danego roku – Set kluczy YYYY-MM-DD. */
function getPolishHolidaysForYear_(year) {
  const set = new Set();
  const add = (month, day) => {
    const d = new Date(year, month - 1, day);
    set.add(formatDateKey_(d));
  };
  add(1, 1);
  add(1, 6);
  const easter = getEasterSunday_(year);
  add(easter.getMonth() + 1, easter.getDate());
  const easterMonday = new Date(easter.getTime());
  easterMonday.setDate(easterMonday.getDate() + 1);
  set.add(formatDateKey_(easterMonday));
  const corpusChristi = new Date(easter.getTime());
  corpusChristi.setDate(corpusChristi.getDate() + 60);
  set.add(formatDateKey_(corpusChristi));
  add(5, 1);
  add(5, 3);
  add(8, 15);
  add(11, 1);
  add(11, 11);
  add(12, 25);
  add(12, 26);
  return set;
}

/** Czy dzien jest roboczy (nie sobota, nie niedziela, nie swieto). */
function isWorkingDay_(date) {
  const d = normalizeDate_(date);
  const day = d.getDay();
  if (day === 0 || day === 6) {
    return false;
  }
  const year = d.getFullYear();
  const holidays = getPolishHolidaysForYear_(year);
  return !holidays.has(formatDateKey_(d));
}

/** Pierwszy dzien roboczy w dniu podanym lub po nim. */
function getFirstWorkingDayOnOrAfter_(date) {
  let d = normalizeDate_(date);
  while (!isWorkingDay_(d)) {
    d.setDate(d.getDate() + 1);
    d = normalizeDate_(d);
  }
  return d;
}

function buildTaskKey_(clientName, procedureName, dueDate) {
  return [
    normalizeText_(clientName),
    normalizeText_(procedureName),
    formatDateKey_(dueDate),
  ].join('|');
}

function getCurrentUserEmail_() {
  const activeEmail = Session.getActiveUser().getEmail();
  const effectiveEmail = Session.getEffectiveUser().getEmail();
  return normalizeText_(activeEmail || effectiveEmail).toLowerCase();
}

function clearSheetBody_(sheet, maxColumns) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return;
  }
  const range = sheet.getRange(2, 1, lastRow - 1, maxColumns);
  range.clearContent().clearFormat();
  range.setDataValidation(null);
}

function ensureSheetSize_(sheet, requiredRows, requiredColumns) {
  const minRows = Math.max(1, toNumber_(requiredRows, 1));
  const minColumns = Math.max(1, toNumber_(requiredColumns, 1));

  const currentRows = sheet.getMaxRows();
  if (currentRows < minRows) {
    sheet.insertRowsAfter(currentRows, minRows - currentRows);
  }

  const currentColumns = sheet.getMaxColumns();
  if (currentColumns < minColumns) {
    sheet.insertColumnsAfter(currentColumns, minColumns - currentColumns);
  }
}

function getLookupMap_(rows, keyField, valueField) {
  const map = {};
  rows.forEach((row) => {
    const key = normalizeText_(row[keyField]);
    if (!key) {
      return;
    }
    map[key] = row[valueField];
  });
  return map;
}

function parseScheduleDay_(value) {
  const raw = normalizeText_(value).toUpperCase();
  if (!raw) {
    return null;
  }

  if (raw === SCHEDULE_LAST_DAY_TOKEN || raw === 'LAST') {
    return { mode: 'last' };
  }

  const day = toNumber_(raw, 0);
  if (day >= 1 && day <= 31) {
    return { mode: 'day', day };
  }
  return null;
}

function getFirstDayOfMonth_(date) {
  const normalized = normalizeDate_(date);
  return new Date(normalized.getFullYear(), normalized.getMonth(), 1);
}

function getLastDayOfMonth_(year, monthZeroBased) {
  return new Date(year, monthZeroBased + 1, 0).getDate();
}

function getDueDateForMonth_(year, monthZeroBased, schedule) {
  if (!schedule) {
    return null;
  }

  const lastDay = getLastDayOfMonth_(year, monthZeroBased);
  if (schedule.mode === 'last') {
    return normalizeDate_(new Date(year, monthZeroBased, lastDay));
  }

  if (schedule.mode === 'day') {
    const targetDay = Math.min(lastDay, Math.max(1, schedule.day));
    return normalizeDate_(new Date(year, monthZeroBased, targetDay));
  }

  return null;
}

function getNextMonthlyDueDate_(dueDate, schedule) {
  const normalized = normalizeDate_(dueDate);
  const nextMonth = new Date(normalized.getFullYear(), normalized.getMonth() + 1, 1);
  return getDueDateForMonth_(nextMonth.getFullYear(), nextMonth.getMonth(), schedule);
}

function getMonthStartsBetween_(startDate, endDate) {
  const start = getFirstDayOfMonth_(startDate);
  const end = getFirstDayOfMonth_(endDate);
  const months = [];

  for (
    let cursor = new Date(start.getTime());
    cursor <= end;
    cursor = new Date(cursor.getFullYear(), cursor.getMonth() + 1, 1)
  ) {
    months.push(new Date(cursor.getTime()));
  }

  return months;
}
