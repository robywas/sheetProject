# sheetProject - zarzadzanie cyklicznymi procedurami w Google Sheets

Gotowy szablon Google Apps Script do obslugi regularnych procedur medyczno-opiekunczych:

- ewidencja procedur cyklicznych,
- powiazanie procedur z klientami,
- przypisania klientow do pracownikow w okresach czasu,
- widok pracownika ("Moje_zadania"),
- widok managera ("Dashboard_managera" + panel boczny).

## 1) Struktura danych (zakladki)

Po uruchomieniu `setupWorkbook()` skrypt zaklada:

1. `Procedury`
2. `Klienci`
3. `Pracownicy`
4. `Klienci_Procedury`
5. `Przypisania`
6. `Zadania`
7. `Moje_zadania`
8. `Dashboard_managera`

## 2) Kluczowe funkcje

- `setupWorkbook()`  
  Zaklada/odswieza zakladki i naglowki oraz migruje stare nazwy zakladek
  `Pacjenci` -> `Klienci` i `Pacjenci_Procedury` -> `Klienci_Procedury`.

- `seedSampleData()`  
  Uzupelnia arkusz przykladowymi danymi (tylko gdy zakladki sa puste poza naglowkiem).

- `generateTasks30Days()`  
  Generuje zadania cykliczne na 30 dni do przodu:
  - uwzglednia aktywne relacje klient-procedura,
  - uwzglednia `dzien_miesiaca` procedury (`1..31` lub `OSTATNI`),
  - przypisuje pracownika wg zakladki `Przypisania` (z rotacja wg `kolejnosc`),
  - nie duplikuje zadan (klucz: `client_id|procedure_id|due_date`).

- `refreshMyTasksView()`  
  Odswieza widok pracownika:
  - pokazuje tylko zadania przypisane do zalogowanego usera (po `email` z zakladki `Pracownicy`),
  - sortuje po terminie,
  - pozwala oznaczac wykonanie checkboxem.

- `refreshManagerDashboard()`  
  Buduje dashboard managera:
  - KPI (otwarte, przeterminowane, termin <= 7 dni, wykonanie 30 dni),
  - lista zagrozonych terminow,
  - obciazenie pracownikow.

- `onEdit(e)`  
  Gdy pracownik zaznaczy checkbox w `Moje_zadania`:
  - zadanie przechodzi do statusu `WYKONANE`,
  - automatycznie tworzy sie kolejne zadanie (nastepny miesiac),
  - nowe zadanie jest przypisywane do kolejnego pracownika z puli klienta.

## 3) Interfejsy

Po odswiezeniu arkusza pojawia sie menu `Procedury`:

1. `Utworz/odswiez strukture`
2. `Dodaj dane przykladowe`
3. `Wygeneruj zadania (30 dni)`
4. `Odswiez moje zadania`
5. `Odswiez dashboard managera`
6. `Panel pracownika`
7. `Panel managera`

Panele boczne:

- `WorkerSidebar.html` - szybkie KPI dla pracownika + odswiezenie widoku.
- `ManagerSidebar.html` - KPI managera + generowanie zadan + odswiezanie dashboardu.

## 4) Jak uruchomic

### Opcja A: bezposrednio w Google Apps Script

1. Otworz docelowy Google Sheet.
2. Wejdz w `Rozszerzenia -> Apps Script`.
3. Wklej pliki `.gs` i `.html` z repo (oraz `appsscript.json`).
4. Zapisz projekt i odswiez arkusz.
5. Z menu `Procedury` uruchom:
   - `Utworz/odswiez strukture`,
   - (opcjonalnie) `Dodaj dane przykladowe`,
   - `Wygeneruj zadania (30 dni)`.

### Opcja B: przez clasp

1. `npm i -g @google/clasp`
2. `clasp login`
3. `clasp clone <scriptId>` lub `clasp create --type sheets`
4. Skopiuj pliki z repo do katalogu clasp
5. `clasp push`

## 5) Ustawienia produkcyjne (zalecane)

1. W `Pracownicy` uzupelnij poprawne emaile kont Google.
2. W `Procedury` ustaw `dzien_miesiaca` jako liczbe `1..31` lub `OSTATNI`.
3. W `Przypisania` utrzymuj zakresy dat przypisania klienta do pracownika
   i (opcjonalnie) `kolejnosc` do sterowania rotacja.
4. Po dodaniu klienta powiaz go recznie z procedurami w `Klienci_Procedury`.
5. Dodaj trigger czasowy (np. codziennie 06:00) dla `generateTasks30Days()`.
6. Ustal workflow statusow (`NOWE`, `W_TRAKCIE`, `WYKONANE`) zgodny z Twoim procesem.

## 6) Rozszerzenia, ktore latwo dodac

- automatyczne powiadomienia email/slack o zadaniach zagrozonych,
- SLA i eskalacje,
- osobny widok audytowy z historia zmian,
- dodatkowe role (koordynator, superwizor),
- walidacje i slowniki danych (lista statusow/procedur).
