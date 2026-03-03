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

function formatDateKey_(date) {
  return Utilities.formatDate(
    normalizeDate_(date),
    Session.getScriptTimeZone(),
    'yyyy-MM-dd'
  );
}

function buildTaskKey_(clientId, procedureId, dueDate) {
  return [normalizeText_(clientId), normalizeText_(procedureId), formatDateKey_(dueDate)].join('|');
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
  sheet.getRange(2, 1, lastRow - 1, maxColumns).clearContent().clearFormat();
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
