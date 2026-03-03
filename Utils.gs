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

function buildTaskKey_(patientId, procedureId, dueDate) {
  return [normalizeText_(patientId), normalizeText_(procedureId), formatDateKey_(dueDate)].join('|');
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

function alignDateToWindow_(startDate, windowStartDate, frequencyDays) {
  const aligned = normalizeDate_(startDate);
  const windowStart = normalizeDate_(windowStartDate);
  if (aligned >= windowStart) {
    return aligned;
  }

  const diff = Math.floor((windowStart - aligned) / ONE_DAY_MS);
  const steps = Math.ceil(diff / frequencyDays);
  aligned.setDate(aligned.getDate() + steps * frequencyDays);
  return aligned;
}
