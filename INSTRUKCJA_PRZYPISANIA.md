# Zasady przypisywania zadań

## Źródło przypisań

Przypisania definiuje arkusz **Przypisania** (kolumny: klient, pracownik, data_od, data_do, kolejnosc). Dla każdego klienta określasz, kto ma wykonywać zadania i w jakiej kolejności.

## Rotacja

- W wierszu Przypisania pole **pracownik** może być **puste**. Wtedy do rotacji wchodzą wszyscy pracownicy z arkusza Pracownicy oznaczani jako **aktywny** (checkbox).
- Gdy **pracownik** jest podany, zadania tego klienta trafiają wyłącznie do niego (bez rotacji).

## Kolejność

Kolumna **kolejnosc** (liczba ≥ 1) ustala priorytet przypisań dla danego klienta: mniejsza wartość = wyższy priorytet. Przy wielu wierszach dla tego samego klienta (np. kilku pracownikach w rotacji) kolejność na liście ustalana jest według **kolejnosc**, a przy tej samej wartości — alfabetycznie po nazwisku.

## Okres ważności

**data_od** i **data_do** ograniczają, dla jakich terminów przypisanie obowiązuje:

- **Przy wyborze pracownika:** zadanie z danym terminem (*due_date*) dostaje pracownika tylko z tych wierszy Przypisań, dla których ten termin mieści się w przedziale [data_od, data_do] (puste data_od/data_do = brak ograniczenia z danej strony).
- **Przy automatycznym tworzeniu „następnego” zadania:** gdy zamykasz zadanie (status WYKONANE), system tworzy kolejne zadanie według harmonogramu procedury. Jeśli termin tego kolejnego zadania **wypada po** ostatniej **data_do** w Przypisaniach dla tego klienta, to zadanie **nie jest tworzone** — data_do traktowana jest jako koniec generowania zadań dla tego przypisania.

## Kiedy ustalany jest pracownik

Pracownik jest ustawiany w momencie **generowania zadań** (np. „Wygeneruj zadania na 30 dni”): dla każdej pary klient–procedura system wybiera **następnego** pracownika z listy (rotacja). Istniejące zadania z pustym polem pracownik są wtedy uzupełniane tą samą logiką.

## Ręczna zmiana

W arkuszu **Zadania** pole **pracownik** może edytować tylko **manager**. Wprowadzone ręcznie przypisanie jest zachowywane — przy kolejnej generacji nie jest nadpisywane.
