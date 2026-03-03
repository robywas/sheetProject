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
    'procedure_id',
    'procedura',
    'opis',
    'czestotliwosc_dni',
    'dni_ostrzezenia',
    'aktywna',
  ],
  CLIENTS: ['client_id', 'klient', 'aktywny'],
  EMPLOYEES: ['employee_id', 'pracownik', 'email', 'rola', 'aktywny'],
  CLIENT_PROCEDURES: [
    'client_id',
    'procedure_id',
    'data_start',
    'czestotliwosc_override',
    'aktywna',
  ],
  ASSIGNMENTS: ['client_id', 'employee_id', 'data_od', 'data_do', 'aktywna'],
  TASKS: [
    'task_id',
    'client_id',
    'procedure_id',
    'employee_id',
    'due_date',
    'status',
    'created_at',
    'completed_at',
    'notes',
    'task_key',
    'dni_ostrzezenia',
  ],
  MY_TASKS: [
    'oznacz_wykonane',
    'task_id',
    'termin',
    'klient',
    'procedura',
    'status',
    'notatka',
  ],
});

const STATUS = Object.freeze({
  NEW: 'NOWE',
  IN_PROGRESS: 'W_TRAKCIE',
  DONE: 'WYKONANE',
});

const MY_TASKS_COL = Object.freeze({
  CHECKBOX: 1,
  TASK_ID: 2,
  DUE_DATE: 3,
  CLIENT: 4,
  PROCEDURE: 5,
  STATUS: 6,
  NOTE: 7,
});

const DEFAULT_GENERATION_DAYS = 30;
const ONE_DAY_MS = 24 * 60 * 60 * 1000;
