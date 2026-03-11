const SHEET_NAMES = Object.freeze({
  PROCEDURES: 'Procedury',
  CLIENTS: 'Klienci',
  EMPLOYEES: 'Pracownicy',
  CLIENT_PROCEDURES: 'Klienci_Procedury',
  ASSIGNMENTS: 'Przypisania',
  TASKS: 'Zadania',
  MY_TASKS: 'Moje_zadania',
  MANAGER_DASHBOARD: 'Dashboard_managera',
});

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
    'due_date',
    'status',
    'created_at',
    'completed_at',
    'notes',
    'task_key',
    'dni_ostrzezenia',
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

const MY_TASKS_COL = Object.freeze({
  TASK_ID: 1,
  DUE_DATE: 2,
  CLIENT: 3,
  PROCEDURE: 4,
  STATUS: 5,
  RELATION_NOTE: 6,
  NOTE: 7,
});

/** Id wersji (np. short commit) – zmien przy deployu, zeby potwierdzic clasp push. */
const DEPLOY_ID = 'd6370b2';

const DEFAULT_GENERATION_DAYS = 30;
const ONE_DAY_MS = 24 * 60 * 60 * 1000;
const SCHEDULE_LAST_DAY_TOKEN = 'OSTATNI';
const DEFAULT_SHEET_MIN_ROWS = 11;
const DASHBOARD_MIN_ROWS = 120;

const LEGACY_SHEET_NAMES = Object.freeze({
  CLIENTS: 'Pacjenci',
  CLIENT_PROCEDURES: 'Pacjenci_Procedury',
});
