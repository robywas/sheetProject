# sheetProject - zarzadzanie cyklicznymi procedurami w Google Sheets

Gotowy szablon Google Apps Script do obslugi regularnych procedur medyczno-opiekunczych:

- ewidencja procedur cyklicznych,
- powiazanie procedur z klientami,
- przypisania klientow do pracownikow w okresach czasu,
- widok pracownika ("Moje_zadania"),
- widok managera ("Dashboard_managera" + panel boczny).

Model danych dziala na nazwach (`klient`, `pracownik`, `procedura`) bez osobnych kolumn ID.

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
  Dodatkowo ustawia twarde walidacje danych i gotowa formatke dashboardu managera.

- `seedSampleData()`  
  Uzupelnia arkusz przykladowymi danymi (tylko gdy zakladki sa puste poza naglowkiem).

- `generateTasks30Days()`  
  Generuje zadania cykliczne na 30 dni do przodu:
  - uwzglednia relacje klient-procedura z arkusza `Klienci_Procedury`,
  - uwzglednia tryb procedury:
    - `MIESIECZNY` + `dzien_miesiaca` (`1..31` lub `OSTATNI`) + `interwal` (np. `3`, `6`),
    - `DZIENNY` + `interwal` (np. `1` codziennie, `2` co 2 dni),
  - przypisuje pracownika wg zakladki `Przypisania` (z rotacja wg `kolejnosc`),
  - nie duplikuje zadan (klucz: `klient|procedura|due_date`),
  - po zapisie sortuje `Zadania`:
    - najpierw niewykonane (`NOWE`, `W_TRAKCIE`) po `due_date` malejaco,
    - potem `WYKONANE` po `due_date` malejaco,
  - podswietla status dla zadan wykonanych po terminie.

- `refreshMyTasksView()`  
  Odswieza widok pracownika:
  - pokazuje tylko zadania przypisane do zalogowanego usera (po `email` z zakladki `Pracownicy`),
  - sortuje po terminie,
  - pozwala zmieniac status (`NOWE`, `W_TRAKCIE`, `WYKONANE`),
  - pokazuje `uwagi` z arkusza `Klienci_Procedury`.

- `refreshManagerDashboard()`  
  Buduje dashboard managera:
  - KPI (otwarte, przeterminowane, termin <= horyzont, wykonanie 30 dni),
  - lista zagrozonych terminow,
  - obciazenie pracownikow,
  - podsumowanie klientow,
  - globalny status wykonania dla wszystkich klientow,
  - filtry: status, pracownik, horyzont i prog zagrozenia.

- `onEdit(e)`  
  Gdy pracownik zmieni status w `Moje_zadania`:
  - status `WYKONANE` zamyka zadanie,
  - automatycznie tworzy sie kolejne zadanie wg trybu i interwalu procedury,
  - nowe zadanie jest przypisywane do kolejnego pracownika z puli klienta.
  Dodatkowo, gdy manager zmieni `pracownik` w arkuszu `Zadania`,
  widok `Moje_zadania` jest odswiezany dla wskazanego pracownika.

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
2. W `Procedury` ustaw:
   - `tryb_harmonogramu`: `MIESIECZNY` albo `DZIENNY`,
   - `interwal`: liczba >= 1 (np. 3 = co 3 miesiace / co 3 dni),
   - `dzien_miesiaca`: wymagany dla trybu miesiecznego (`1..31` lub `OSTATNI`).
3. W `Przypisania` utrzymuj zakresy dat przypisania klienta do pracownika
   i `kolejnosc` do sterowania rotacja.
   Przy pustej kolumnie `pracownik` skrypt potraktuje to jako rotacje miedzy wszystkimi pracownikami.
4. Po dodaniu klienta powiaz go recznie z procedurami w `Klienci_Procedury`.
   W kolumnie `uwagi` mozesz dopisac wskazowki dla wykonawcy (widoczne w `Moje_zadania`).
5. Dodaj trigger czasowy (np. codziennie 06:00) dla `generateTasks30Days()`.
6. Ustal workflow statusow (`NOWE`, `W_TRAKCIE`, `WYKONANE`) zgodny z Twoim procesem.

## 6) Twarde walidacje danych (ustawiane automatycznie)

Po `setupWorkbook()` skrypt ustawia:

- `Procedury!dzien_miesiaca` - dropdown puste/`1..31`/`OSTATNI`,
- `Procedury!dni_ostrzezenia` - liczba calkowita >= 0,
- `Procedury!tryb_harmonogramu` - dropdown `MIESIECZNY` / `DZIENNY`,
- `Procedury!interwal` - liczba calkowita >= 1,
- `Przypisania!kolejnosc` - liczba calkowita >= 1,
- walidacje nazw `klient`, `procedura`, `pracownik` na podstawie slownikow,
- `Zadania!pracownik` - lista pracownikow z mozliwoscia pustej wartosci,
- dodatkowa kontrola przy edycji (onEdit), ktora czyści nieprawidlowe liczby niecalkowite
  w `dni_ostrzezenia`, `interwal` i `kolejnosc`.

Arkusze startuja od niewielkiej liczby wierszy (ok. 10 + naglowek), a skrypt
powieksza je automatycznie w miare potrzeb.

## 7) Rozszerzenia, ktore latwo dodac

- automatyczne powiadomienia email/slack o zadaniach zagrozonych,
- SLA i eskalacje,
- osobny widok audytowy z historia zmian,
- dodatkowe role (koordynator, superwizor),
- walidacje i slowniki danych (lista statusow/procedur).
