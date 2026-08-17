# sh_LandingPage - specyfikacja layoutu (Arkusz1)

> **Cel**: Pulpit (hub) aplikacji. Landing page po Workbook_Open + user picker. Zamiast frm_Main jako main screen, mamy Arkusz1 z 5 button-launcher'ami do wszystkich formularzy (hub-and-spoke pattern).
>
> **Widoczność**: `xlSheetVisible` (jedyny widoczny sheet w xlsm). Wszystkie inne (`ws_AppState`, `ws_UsersRegistry`, `ws_DataCache`) pozostają `xlSheetVeryHidden`.
>
> **Trigger show**: automatycznie po `frm_UserPicker.btn_SelectUser_Click → Me.Hide` LUB po `frm_Setup.btn_Save_Click → Me.Hide` (Arkusz1 juz aktywny, formularz modal się chowa).
>
> **Code module**: `Source/Sheets/sh_LandingPage.code.txt` (wklej do code module tego arkusza).
>
> **Plan**: M6/M7

---

## Krok 1 - utwórz + rename Arkusz1 |x

Arkusz1 jest domyślny w każdym nowym xlsm (Excel PL). W Excel EN nazwa
default to `Sheet1` - te same czynności, inna wyświetlana nazwa.

W VBE Project Explorer:
1. Klik na `Arkusz1 (Arkusz1)` → F4 (Properties)
2. `(Name)` (CodeName w VBA) = **`sh_LandingPage`** (locale-safe identifier)
3. `Name` (tab caption) = **`BNC_Sender - Pulpit`**
4. `Visible` = **`-1 - xlSheetVisible`**

> **Ważne**: kod aplikacji odwołuje się do sheeta przez CodeName `sh_LandingPage`,
> nie przez tab name ani lokalizowaną nazwę Arkusz1/Sheet1. Dzięki temu działa
> tak samo na Polish/English Excel bez zmian w kodzie.

**Konwencja naming**:
- Prefix `sh_` = visible sheet (odróżnia od `ws_` = very hidden data sheets)
- `sh_LandingPage` jest **jedyny** sh_* w projekcie na razie

---

## Krok 2 - komórki (statyczne) |x

### Sekcja A: Header (row 1-5)

| Cell | Zawartość | Formatowanie |
|---|---|---|
| `A1` | `BNC_Sender - Pulpit` | Font Segoe UI 20 Bold, ForeColor `&H00804000&` (ciemnoniebieski) |x
| `A3` | `Zalogowany:` | Font Segoe UI 10 Bold |x
| `B3` | *(dynamic, wypełnia `RefreshDashboard`)* | Font Segoe UI 10 |x
| `A4` | `CNA:` | Font Segoe UI 10 Bold |x
| `B4` | *(dynamic — CNA · Oddział)* | Font Segoe UI 10 |x
| `A5` | `Rola:` | Font Segoe UI 10 Bold |x
| `B5` | *(dynamic — HANDLOWIEC / KIEROWNIK + kontekst)* | Font Segoe UI 10, ForeColor per rola (opcjonalne) |x

**Merge cells**: `A1:H1` (title spans across), pozostałe unmerged. |x

### Sekcja B: Buttons (row 7-13) - Form Controls (nie ActiveX)

5 **Form Controls Button** kontrolek. Layout: 2×2 grid + 1 pod spodem.

**Dlaczego Form Controls (nie ActiveX)**: production robustness. ActiveX moze byc zablokowany przez Trust Center IT policies (typowe w korporacjach), nie dziala na Mac/Web. Form Controls dzialaja wszedzie - Windows/Mac/Web, kazda polityka bezpieczenstwa. 20+ handlowcow = 20+ potencjalnych IT policies = ryzyko deployment blockers ActiveX. Form Controls = zero risk. Trade-off: wygladaja starzej, mniej style properties - akceptowalne dla launcher pattern.

W Developer tab → Insert → **Form Controls** (pierwsza sekcja dropdown, NIE ActiveX Controls) → **Button** (pierwsza ikona):

| # | Caption | Pozycja (Excel row/col) | Assigned Macro (Public Sub w mod_LandingPage) |
|---|---|---|---|
| 1 | `Nowe zgloszenie` | Range B7:C9 | `OpenMain` |x
| 2 | `Historia + Log` | Range E7:F9 | `OpenLog` |x
| 3 | `Przelacz usera` | Range B11:C13 | `OpenPicker` |
| 4 | `Ustawienia setup` | Range E11:F13 | `OpenSetup` |
| 5 | `Samouczek` | Range D15:E17 (centered below grid) | `OpenTutorial` |

**Instrukcja krok-po-kroku**:

1. Developer tab → Insert → **Form Controls** dropdown (**NIE** ActiveX)
2. Kliknij ikonę **Button** (pierwsza z lewej)
3. Drag na Arkusz1 w wybranym miejscu (np. Range B7:C9)
4. **Excel otworzy dialog "Assign Macro"** → z listy wybierz **`OpenMain`** → OK
5. Right-click na button → **Edit Text** → wpisz `Nowe zgloszenie` → Enter
6. Powtórz dla 4 pozostałych buttonów (pozycje + macro per tabela wyżej)

**Assigned Macro** = macro wywolywane po klikni. Excel pokazuje w Assign Macro dialog wszystkie Public Subs z standalone modules (nie z Sheet code modules - stad handlery w mod_LandingPage.bas). Wybierz z listy `OpenMain` / `OpenLog` / `OpenPicker` / `OpenSetup` / `OpenTutorial`.

**Zmiana macro po utworzeniu**: right-click button → **Assign Macro...** → wybierz inny → OK.

**Formatowanie tekstu (opcjonalne)**: right-click button → **Format Control...** → tab **Font**. Ustaw Font Size 11, Bold. Form Controls maja ograniczone styling - w praktyce tekst caption wyglada OK domyslnie.

**Rozmiar**: Form Controls skaluja sie z komorkami. Drag krawedzie po utworzeniu zeby dopasowac do siatki 2×2 grid.

**Design Mode dla Form Controls**: NIE potrzebne (w przeciwienstwie do ActiveX). Klikanie w button od razu wolulje macro. Do edycji buttonu (move/resize/re-assign macro) - prawy klik na button.

### Sekcja C: Statystyki (row 15-17 alternatywnie 19-21 zaleznie od tutorial button)

Zakładając Tutorial w row 15-17, statystyki row **19-21**:

| Cell | Zawartość | Formatowanie |
|---|---|---|
| `A19` | `Statystyki:` | Font Segoe UI 12 Bold |
| `A20` | `Pending:` | Font Segoe UI 10 Bold |
| `B20` | *(dynamic — liczba pending records)* | Font Segoe UI 10 |
| `D20` | `Wyslane (all-time):` | Font Segoe UI 10 Bold |
| `E20` | *(dynamic — liczba sent records)* | Font Segoe UI 10 |

**Refresh**: `Worksheet_Activate` w Arkusz1 code odczytuje z `mod_DataCacheSync.GetPendingRecords().Count` + `mod_DataCacheSync.GetAllRecords().Count - pending`.

### Sekcja D: Footer note (opcjonalna, row 23)

| Cell | Zawartość |
|---|---|
| `A23` | `Kliknij dowolny przycisk aby otworzyc formularz. Zamkniecie formularza wraca tu.` |

Font Segoe UI 9 italic, ForeColor gray.

---

## Krok 3 - Column widths + Row heights

Domyślne szerokości kolumn Excela są za małe. Ustaw:

| Column | Width |
|---|---|
| A | 4 (padding lewy) |
| B | 20 |
| C | 20 |
| D | 4 (przerwa między buttonami) |
| E | 20 |
| F | 20 |
| G | 20 |
| H | 4 (padding prawy) |

Row heights:
| Row | Height |
|---|---|
| 1 | 40 (title) |
| 2 | 8 (spacer) |
| 3-5 | 18 (user info) |
| 6 | 12 (spacer) |
| 7-9 | 30 each (buttons row 1) |
| 10 | 8 (spacer) |
| 11-13 | 30 each (buttons row 2) |
| 14 | 8 (spacer) |
| 15-17 | 30 each (tutorial button) |
| 18 | 12 (spacer) |
| 19-20 | 20 each (stats) |
| 23 | 15 (footer note) |

---

## Krok 4 - Wyłącz gridlines + hide row/column headers (aesthetic)

W View tab (Ribbon Excel):
- Uncheck `Gridlines`
- Uncheck `Headings`

Alternatywnie w VBA (przy pierwszym Workbook_Open):
```vba
sh_LandingPage.Activate
ActiveWindow.DisplayGridlines = False
ActiveWindow.DisplayHeadings = False
```

---

## Krok 4b - App-like mode (hide ribbon + lock cells + hide obszaru poza landing) |NEW

Cel: wygląd dedykowanej aplikacji zamiast Excela. Trzy warstwy:

### 1. Ribbon + Excel chrome (kod, automatycznie w `Workbook_Open`)

Nic ręcznie - `ThisWorkbook.Workbook_Open` już to robi:
- `Application.DisplayFullScreen = True` (fullscreen mode)
- `Application.DisplayFormulaBar = False`
- `ActiveWindow.DisplayHeadings = False` (bez A/B/C i 1/2/3)
- `ActiveWindow.DisplayGridlines = False`
- `ActiveWindow.DisplayWorkbookTabs = False` (bez tab bar u dołu)
- `SHOW.TOOLBAR("Ribbon", False)` (ribbon całkowicie hidden, nie tylko zminimalizowany)

Restore w `Workbook_BeforeClose` żeby inne pliki Excel otwarte równolegle nie dziedziczyły.

### 2. Ukryj kolumny i wiersze poza landing page (RĘCZNIE, raz)

Landing page zajmuje `A1:H23`. Wszystko poza tym obszarem = "puste" wizualnie (szare tło Excela).

**Kolumny I:XFD**:
1. Klik na nagłówek kolumny **I**
2. Ctrl+Shift+End (rozszerza selekcję do ostatniej kolumny)
3. Right-click → **Hide**

**Wiersze 24:1048576**:
1. Klik na nagłówek wiersza **24**
2. Ctrl+Shift+End
3. Right-click → **Hide**

**Save workbook** (persistent). Landing area teraz widoczna, reszta zniknięta.

> ⚠ **UWAGA: sam `Hide` NIE wystarcza** - klasyczny bug Excela: user Ctrl+Right / Ctrl+End / scroll wheel może "przejechać" do ukrytego obszaru mimo Hide. Ochronę przed tym daje `ScrollArea` + `EnableSelection = xlNoSelection` z warstwy 3 poniżej. Hide + ScrollArea + xlNoSelection = trzy niezależne warstwy blokady.

### 3. Cell lock + no selection (kod, automatycznie w `Workbook_Open`)

`ThisWorkbook.LockLandingPage()` uruchamiane po `sh_LandingPage.Activate`:
```vba
With sh_LandingPage
    .Unprotect  ' na wypadek istniejącej protekcji
    .Cells.Locked = True
    .EnableSelection = xlNoSelection     ' user w ogóle nie może kliknąć w cell
    .ScrollArea = "A1:H23"               ' scroll ograniczony do landing area
    .Protect Password:="", UserInterfaceOnly:=True, _
             DrawingObjects:=True, Contents:=True, Scenarios:=True
End With
```

**Rola każdej właściwości**:
- `ScrollArea = "A1:H23"` → **hard block scrolla** poza landing area (fix na "hide + Ctrl+End przejeżdża do ukrytych"). NIE persistuje między sesjami - stąd ustawiany w każdym `Workbook_Open`
- `EnableSelection = xlNoSelection` → user w ogóle nie może zaznaczyć cell, więc nawigacja Ctrl+arrow/Tab też nie działa
- `Cells.Locked + Protect` → nawet gdyby jakoś doszedł, edycja zablokowana
- `UserInterfaceOnly:=True` → VBA (`RefreshDashboard`) pisze do B3/B4/B5/B20/E20 mimo protekcji. User nie może
- `DrawingObjects:=True` → **blokuje tylko move/resize/delete buttonów**, macro klik przechodzi (Form Controls dalej działają)

### Rezultat

Wygląd aplikacji "single-window":
- Górna belka: tylko title bar Windows (bez ribbon)
- Landing area: widoczna, wycentrowana, nieklikanie w komórki
- Wszystko poza: szare tło Excela (schowane kolumny/wiersze)
- User może kliknąć tylko przyciski Form Controls

Restore normalnego Excela: zamknięcie xlsm → `Workbook_BeforeClose` przywraca ribbon + chrome.

---

## Krok 5 - code-behind

Pełny code w [`sh_LandingPage.code.txt`](sh_LandingPage.code.txt). Struktura:

- **`Worksheet_Activate`** — refresh user info + stats (called when Arkusz1 becomes active, e.g., after form close)
- **`RefreshDashboard`** (private) — pisze do cells B3, B4, B5, B20, E20
- **5 button click handlers** — każdy `frm_XXX.Show vbModal`
- **Bez SelectionChange** (nie potrzebujemy trackować gdzie user klika)

---

## Krok 6 - Integracja z ThisWorkbook

Workbook_Open (patrz `Source/ThisWorkbook/ThisWorkbook.code.txt`):
- `Application.DisplayFullScreen = True` (app-like feel per user's decision)
- `sh_LandingPage.Activate` (na koniec, żeby user widział Arkusz1 po zamknięciu picker/setup)

Workbook_BeforeClose:
- `Application.DisplayFullScreen = False` (restore dla innych plików Excel otwartych rownolegle)

---

## Wizualne mockup (ASCII)

```
+===================================================================+
|  BNC_Sender - Pulpit                                              |  ← A1 (Segoe UI 20 Bold)
|                                                                    |
|  Zalogowany:  Jan Kowalski                                         |  ← A3:B3
|  CNA:         12345 · Oddzial: W001                                |  ← A4:B4
|  Rola:        HANDLOWIEC (mail leci do kierownika)                 |  ← A5:B5
|                                                                    |
|                                                                    |
|      +--------------------+     +--------------------+             |
|      |  Nowe zgloszenie   |     |  Historia + Log    |             |  ← B7:F9
|      +--------------------+     +--------------------+             |
|                                                                    |
|      +--------------------+     +--------------------+             |
|      |  Przelacz usera    |     |  Ustawienia setup  |             |  ← B11:F13
|      +--------------------+     +--------------------+             |
|                                                                    |
|                    +--------------------+                          |
|                    |     Samouczek      |                          |  ← D15:E17
|                    +--------------------+                          |
|                                                                    |
|  Statystyki:                                                       |
|  Pending: 3         Wyslane (all-time): 47                         |  ← A20/B20 + D20/E20
|                                                                    |
|  Kliknij dowolny przycisk aby otworzyc formularz.                  |  ← A23 (italic)
+===================================================================+
```

## Rozmiar szacowany

Przy ~1600×900 monitorze:
- Arkusz1 wypełnia obszar poniżej ribbon (chyba że DisplayFullScreen=True → wypełnia cały ekran)
- Ustawienie window size przez `ActiveWindow.WindowState = xlMaximized` przy Workbook_Open zapewnia pełny obszar

---

## Diagnostic integration

`mod_Diagnostic.ExpectedSheets` — 3 → **4** (dodać `sh_LandingPage`).
`ExpectedSheetHeaders("sh_LandingPage")` — brak (nie tabela, tylko labels/buttons — skip check).
`AuditSheets` autoNote dla missing case: `"tworzone recznie w VBE - patrz sh_LandingPage.LAYOUT.md"`.

Opcjonalnie w przyszłości: audyt kontrolek ActiveX na Arkusz1 (analogicznie do `AuditFormControls`). Skip dla M6/M7.
