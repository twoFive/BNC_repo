# frm_Tutorial - treść stron (v3 - minimalist + setup page)

> **5 stron**, ~80-100 słów każda. User-perspective, zero tech internals.
> **Trigger**: wyłącznie manualnie - `frm_Setup.btn_ShowTutorial` (brak auto-show).
> **Layout**: patrz [`frm_Tutorial.LAYOUT.md`](frm_Tutorial.LAYOUT.md)

---

## Strona 1 z 5 - Powitanie

**Title**: `Witaj w BNC_Sender 👋`

**Body**:
```
BNC_Sender to aplikacja do zbiorczej wysyłki zgłoszeń BNC.

Zamiast wypełniać i wysyłać każde zgłoszenie osobno, wpisujesz je do listy przez cały miesiąc, klikasz jeden przycisk i mail ze wszystkimi zgłoszeniami jest wysyłany przez Outlook do właściwego adresata (do Twojego kierownika, a jeśli sam nim jesteś, wprost na skrzynkę BNC).

Samouczek zajmie ~2 minuty. Naciśnij "Dalej ▶".
```

---

## Strona 2 z 5 - Konfiguracja

**Title**: `Konfiguracja (jednorazowa) ⚙`

**Body**:
```
Przy pierwszym uruchomieniu aplikacji wypełniasz formularz z Twoimi
danymi służbowymi: Imię, Nazwisko, Email, CNA, Nr oddziału,
Email kierownika. Dzięki temu nie musisz już powtarzać wpisywania tych danych w każdym wniosku!

UWAGA: jeśli sam jesteś Kierownikiem, wpisz swój email także w polu
"Email kierownika". Aplikacja rozpozna Twoją rolę automatycznie.

Pola "Email BNC" i "Folder cache" są ustawione fabrycznie. Kliknij "Utwórz folder cache" jeśli folder jeszcze nie istnieje na dysku.

Kliknij "Zapisz". Konfiguracja gotowa - przechodzisz do głównego
ekranu aplikacji (ekranu wysyłki wniosku).
```

---

## Strona 3 z 5 - Codzienna praca

**Title**: `Dodawanie zgłoszeń i wysyłka`

**Body**:
```
DODAJESZ ZGŁOSZENIE:
Wypełnij 3 pola: Klient FK, Nazwa klienta, Miesiąc obrotu
(YYYY-MM, domyślnie bieżący). Kliknij "Dodaj do listy".

POMYŁKA:
Kliknij na wiersz w liście pending → "Usuń zaznaczone" → potwierdź.
Wysłane zgłoszenia są NIEUSUWALNE, sprawdzaj przed wysyłką.

WYSYŁKA (najlepiej raz w miesiącu):
Kliknij "Wyślij Wniosek BNC" → potwierdzenie → mail wychodzi
z Outlooka. Lista pending czyści się, zgłoszenia trafiają
do historii.
```

---

## Strona 4 z 5 - Historia i eksport

**Title**: `Pokaż historię · Eksport`

**Body**:
```
"POKAŻ HISTORIĘ": widzisz wszystkie zgłoszenia (wysłane + pending)
z datą i adresatem. Przydaje się gdy BNC twierdzi "nie dostaliśmy
zgłoszenia" - masz twardy dowód wysyłki.

"EKSPORTUJ DO PLIKU" (w oknie historii): zapisuje wszystko do
osobnego pliku xlsx w wybranej lokalizacji. Backup / dla audytora /
przekazanie zastępcy. Możesz eksportować dowolnie często.
```

---

## Strona 5 z 5 - Pomoc

**Title**: `Gdy coś nie działa · Kontakt 🛟`

**Body**:
```
NAJCZĘSTSZE:
• Excel pyta o makra → klik "Włącz zawartość"
• Outlook pyta o pozwolenie na wysyłkę → klik "Allow"
• "Wyślij" nic nie robi → sprawdź czy Outlook jest otwarty


Inne problemy / Sugestie:
tomasz.pirszel@inter-team.com.pl

Powodzenia! ✉ 📈
Kliknij "Zakończ ✓" żeby zamknąć.
```

---

## Uwagi

**Word count** (target: 80-100 per strona, total ~400 słów):

| Strona | Słów |
|---|---|
| 1 Powitanie | ~55 |
| 2 Konfiguracja | ~90 |
| 3 Codzienna praca | ~85 |
| 4 Historia i eksport | ~75 |
| 5 Pomoc | ~55 |
| **Total** | **~360 słów** |

**Vs oryginalna wersja**: 10 stron × ~150 słów = 1500 słów → 5 stron × ~72 słów = 360 słów (**76% redukcja**).

**Placeholder replacements** (runtime):
- `tomasz.pirszel@inter-team.com.pl` - hardcoded w code-behind (aktualny kontakt product owner + IT support w jednym)

**Świadomie POMINIĘTE** (vs oryginalna wersja):
- Ekspozycja handlowiec vs kierownik jako osobna strona (touched w Setup: "wpisz swój email w Email kierownika")
- Multi-user picker (self-explanatory UI, ktoś widząc dropdown wie że to lista userów)
- Technical internals: xlsx w %TEMP%, 7 kolumn, audit trail internals, snapshot pattern
- Format field validation (walidacja sama powie userowi jeśli błąd)
- Kolumn count w historii (user widzi kolumny, nie musi znać liczby)

**Ton**: imperatywny (2. os. sg: "Wypełnij", "Kliknij", "Sprawdź"). Krótkie zdania.
Bullet lists tylko gdzie faktycznie enumeracja (Strona 5 - najczęstsze problemy).
Emoji minimalnie: 👋 (powitanie), ⚙ (setup), 🛟 (pomoc), ✉ 📈 ✓ (sign-off).

**Typografia**: brak em dash "—", zamiast tego " - " (myślnik z odstępami) LUB dwukropek/przecinek/kropka zależnie od struktury zdania.

**Dynamic content**: **NIE MA** w tej wersji. Rola usera (handlowiec vs kierownik) widoczna w:
- Setup page (Strona 2): instrukcja jak "wpisać swój email jako kierownika" jeśli user jest kierownikiem
- Runtime confirmation MsgBox przy wysyłce (`"Wyslac X zgloszen do kierownika (email)?"` vs `"Wyslac X zgloszen wprost do BNC?"`)

**Setup page decyzja**: dodana jako Strona 2 (post-Powitanie) bo:
- Setup jest pierwszą rzeczą jaką user widzi przy pierwszym otwarciu
- Bez tego user może się zdziwić że aplikacja od razu prosi o dane osobowe
- Kluczowa informacja: "sam jesteś kierownikiem → wpisz swój email w pole Email kierownika" (convention over configuration ADR-005 wyjaśnione w prostych słowach)
- User widzi tutorial PO setup'ie (bo btn_ShowTutorial jest w frm_Setup), więc Setup page pełni rolę potwierdzenia "właśnie to zrobiłeś, tak to działa"
