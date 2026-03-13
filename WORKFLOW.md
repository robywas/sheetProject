# WORKFLOW - praca na 3 komputerach (GitHub + clasp)

Ten projekt prowadzimy na jednej galezi: `main`.

Cel: uniknac rozjazdow miedzy komputerami i Apps Script.

---

## 1) START PRACY (na dowolnym komputerze)

```bash
cd "<katalog_repo>"
git checkout main
git pull
```

Jesli planujesz zmiany, ktore maja trafic do Google Apps Script:

```bash
clasp push -f
```

---

## 2) W TRAKCIE PRACY

- Edytuj pliki lokalnie.
- Regularnie sprawdzaj status:

```bash
git status
```

---

## 3) KONIEC PRACY

```bash
git add .
git commit -m "krotki opis zmian"
git push
```

Jesli zmiany dotycza GAS (`.gs`, `.html`, `appsscript.json`):

```bash
clasp push -f
```

---

## 4) ZMIANA KOMPUTERA

Po przesiadce na inny komputer zawsze:

```bash
cd "<katalog_repo>"
git checkout main
git pull
```

Dopiero potem zaczynaj nowe zmiany.

---

## 5) ZASADY ANTY-ROZJAZD

1. Zawsze `git pull` przed praca.
2. Nie pracuj rownolegle na 2 komputerach bez push/pull pomiedzy.
3. Trzymaj sie jednej galezi (`main`).
4. `clasp push -f` wykonuj po upewnieniu sie, ze lokalnie masz aktualne `main`.
5. Nie edytuj kodu bezposrednio w Apps Script, chyba ze awaryjnie.
6. **Weryfikacja wgrania:** Przed `clasp push -f` ustaw w `Constants.gs` stalą `DEPLOY_ID` na wynik `git rev-parse --short HEAD` (zcommituj razem z innymi zmianami). Po wgraniu uruchom w arkuszu **Procedury > 1) Utworz/odswiez strukture** – w toastcie zobaczysz **build: <short-commit>**. Ten numer ma byc taki sam jak w podsumowaniu / w git – wtedy potwierdzasz, ze ostatnie zmiany sa wgrane.

---

## 6) SZYBKI RATUNEK PRZY KONFLIKCIE

```bash
git status
git add .
git commit -m "WIP lokalne zmiany"
git pull --rebase
# rozwiaz ewentualne konflikty
git push
```

---

## 7) CHECKLISTA clasp (na nowym komputerze)

```bash
clasp login
```

W katalogu repo utworz `.clasp.json`:

```json
{
  "scriptId": "TU_WKLEJ_SCRIPT_ID",
  "rootDir": "."
}
```

Potem:

```bash
clasp push -f
```

