const SHEET_NAMES = Object.freeze({
  PROCEDURES: 'Procedury',
  CLIENTS: 'Klienci',
  EMPLOYEES: 'Pracownicy',
  CLIENT_PROCEDURES: 'Klienci_Procedury',
  ASSIGNMENTS: 'Przypisania',
  TASKS: 'Zadania',
  MANAGER_DASHBOARD: 'Dashboard_managera',
});

/** Prefiks arkusza zadan per pracownik (np. „Zadania - Jan Kowalski”). */
const MY_TASKS_SHEET_PREFIX = 'Zadania - ';

/** Arkusze widoczne tylko dla managera; przy otwarciu przez pracownika są ukrywane (por. applySheetVisibilityByRole_). */
const SHEETS_VISIBLE_ONLY_TO_MANAGER = Object.freeze([
  SHEET_NAMES.TASKS,
  SHEET_NAMES.ASSIGNMENTS,
  SHEET_NAMES.MANAGER_DASHBOARD,
]);

const HEADERS = Object.freeze({
  PROCEDURES: [
    'procedura',
    'opis',
    'dzien_miesiaca',
    'dni_ostrzezenia',
    'tryb_harmonogramu',
    'interwal',
  ],
  CLIENTS: ['klient'],
  EMPLOYEES: ['pracownik', 'email', 'rola', 'aktywny'],
  CLIENT_PROCEDURES: [
    'klient',
    'procedura',
    'data_start',
    'uwagi',
    'kontrola',
  ],
  ASSIGNMENTS: [
    'klient',
    'pracownik',
    'data_od',
    'data_do',
    'kolejnosc',
  ],
  TASKS: [
    'task_id',
    'klient',
    'procedura',
    'pracownik',
    'status',
    'due_date',
    'completed_at',
    'uwagi',
    'notes',
    'dni_ostrzezenia',
    'created_at',
    'task_key',
  ],
  MY_TASKS: [
    'task_id',
    'termin',
    'klient',
    'procedura',
    'status',
    'uwagi',
    'notatka',
  ],
});

const STATUS = Object.freeze({
  NEW: 'NOWE',
  IN_PROGRESS: 'W_TRAKCIE',
  DONE: 'WYKONANE',
});

const SCHEDULE_MODE = Object.freeze({
  MONTHLY: 'MIESIECZNY',
  DAILY: 'DZIENNY',
});

const ROLE_OPTIONS = Object.freeze(['pracownik', 'manager']);

const MANAGER_FILTER = Object.freeze({
  ALL: 'WSZYSTKIE',
  OPEN: 'OTWARTE',
  ALL_EMPLOYEES: 'WSZYSCY',
  DEFAULT_HORIZON_DAYS: 7,
  DEFAULT_RISK_DAYS: 2,
});

/** Indeksy kolumn arkusza Zadania (1-based). Kolejnosc: task_id, klient, procedura, pracownik, status, due_date, completed_at, uwagi, notes, dni_ostrzezenia, created_at, task_key. */
const TASKS_COL = Object.freeze({
  TASK_ID: 1,
  KLIENT: 2,
  PROCEDURA: 3,
  PRACOWNIK: 4,
  STATUS: 5,
  DUE_DATE: 6,
  COMPLETED_AT: 7,
  UWAGI: 8,
  NOTES: 9,
  DNI_OSTRZEZENIA: 10,
  CREATED_AT: 11,
  TASK_KEY: 12,
});

const MY_TASKS_COL = Object.freeze({
  TASK_ID: 1,
  DUE_DATE: 2,
  CLIENT: 3,
  PROCEDURE: 4,
  STATUS: 5,
  RELATION_NOTE: 6,
  NOTE: 7,
});

/**
 * Dla kopiowania formatowania: indeks kolumny w Zadania (1-based) odpowiadajacy kolumnie w Zadania - X.
 * Kolejnosc: task_id, termin(due_date), klient, procedura, status, uwagi, notatka(notes).
 */
const TASKS_COL_FOR_MY_TASKS_FORMAT = Object.freeze([1, 6, 2, 3, 5, 8, 9]);

/** Id wersji (short commit) – ustaw na aktualny po deployu (git rev-parse --short HEAD). */
const DEPLOY_ID = '5e5f5d7';

const DEFAULT_GENERATION_DAYS = 30;
const ONE_DAY_MS = 24 * 60 * 60 * 1000;
const SCHEDULE_LAST_DAY_TOKEN = 'OSTATNI';
const DEFAULT_SHEET_MIN_ROWS = 11;
const DASHBOARD_MIN_ROWS = 120;

const LEGACY_SHEET_NAMES = Object.freeze({
  CLIENTS: 'Pacjenci',
  CLIENT_PROCEDURES: 'Pacjenci_Procedury',
});
