# sh_Dashboard - specyfikacja layoutu (Sheet1)

> **Cel**: Pulpit (hub) aplikacji. Landing page po Workbook_Open + user picker. Zamiast frm_Main jako main screen, mamy Sheet1 z 5 button-launcher'ami do wszystkich formularzy (hub-and-spoke pattern).
>
> **Widoczność**: `xlSheetVisible` (jedyny widoczny sheet w xlsm). Wszystkie inne (`ws_AppState`, `ws_UsersRegistry`, `ws_DataCache`) pozostają `xlSheetVeryHidden`.
>
> **Trigger show**: automatycznie po `frm_UserPicker.btn_SelectUser_Click → Me.Hide` LUB po `frm_Setup.btn_Save_Click → Me.Hide` (Sheet1 juz aktywny, formularz modal się chowa).
>
> **Code module**: `Source/Sheets/sh_Dashboard.code.txt` (wklej do code module tego arkusza).
>
> **Plan**: M6/M7

---

## Krok 1 - utwórz + rename Sheet1

Sheet1 jest domyślny w każdym nowym xlsm. Rename:

W VBE Project Explorer:
1. Klik na `Sheet1 (Sheet1)` → F4 (Properties)
2. `(Name)` (CodeName w VBA) = **`sh_Dashboard`**
3. `Name` (tab caption) = **`BNC_Sender - Pulpit`**
4. `Visible` = **`-1 - xlSheetVisible`**

**Konwencja naming**:
- Prefix `sh_` = visible sheet (odróżnia od `ws_` = very hidden data sheets)
- `sh_Dashboard` jest **jedyny** sh_* w projekcie na razie

---

## Krok 2 - komórki (statyczne)

### Sekcja A: Header (row 1-5)

| Cell | Zawartość | Formatowanie |
|---|---|---|
| `A1` | `BNC_Sender - Pulpit` | Font Segoe UI 20 Bold, ForeColor `&H00804000&` (ciemnoniebieski) |
| `A3` | `Zalogowany:` | Font Segoe UI 10 Bold |
| `B3` | *(dynamic, wypełnia `RefreshDashboard`)* | Font Segoe UI 10 |
| `A4` | `CNA:` | Font Segoe UI 10 Bold |
| `B4` | *(dynamic — CNA · Oddział)* | Font Segoe UI 10 |
| `A5` | `Rola:` | Font Segoe UI 10 Bold |
| `B5` | *(dynamic — HANDLOWIEC / KIEROWNIK + kontekst)* | Font Segoe UI 10, ForeColor per rola (opcjonalne) |

**Merge cells**: `A1:H1` (title spans across), pozostałe unmerged.

### Sekcja B: Buttons (row 7-13)

5 ActiveX **CommandButton** kontrolek. Layout: 2×2 grid + 1 pod spodem.

W Developer tab → Insert → **ActiveX** → **Command Button**:

| # | Button Name (Properties) | Caption | Pozycja przybliżona (Excel row/col) | Wywoluje |
|---|---|---|---|---|
| 1 | `btn_OpenMain` | `Nowe zgloszenie` | Range B7:C9 | `frm_Main.Show vbModal` |
| 2 | `btn_OpenLog` | `Historia + Log` | Range E7:F9 | `frm_Log.Show vbModal` |
| 3 | `btn_OpenPicker` | `Przelacz usera` | Range B11:C13 | `frm_UserPicker.Show vbModal` |
| 4 | `btn_OpenSetup` | `Ustawienia setup` | Range E11:F13 | `frm_Setup.Show vbModal` |
| 5 | `btn_OpenTutorial` | `Samouczek` | Range D15:E17 (centered below grid) | `frm_Tutorial.Show vbModal` |

**Properties dla wszystkich** (jednolicie):
- Font: Segoe UI 11 Bold
- BackColor: `&H00E0E0E0&` (jasnoszare) lub domyślne
- Approx. size: Width=140, Height=40 (dwukrotność wysokości pojedynczej komórki)

**Instrukcja**: kliknij ActiveX button na tabbie → drag na miejsce → Design Mode (Developer tab → Design Mode) → prawy klik button → View Code → automatycznie tworzy handler w Sheet1 module.

### Sekcja C: Statystyki (row 15-17 alternatywnie 19-21 zaleznie od tutorial button)

Zakładając Tutorial w row 15-17, statystyki row **19-21**:

| Cell | Zawartość | Formatowanie |
|---|---|---|
| `A19` | `Statystyki:` | Font Segoe UI 12 Bold |
| `A20` | `Pending:` | Font Segoe UI 10 Bold |
| `B20` | *(dynamic — liczba pending records)* | Font Segoe UI 10 |
| `D20` | `Wyslane (all-time):` | Font Segoe UI 10 Bold |
| `E20` | *(dynamic — liczba sent records)* | Font Segoe UI 10 |

**Refresh**: `Worksheet_Activate` w Sheet1 code odczytuje z `mod_DataCacheSync.GetPendingRecords().Count` + `mod_DataCacheSync.GetAllRecords().Count - pending`.

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
sh_Dashboard.Activate
ActiveWindow.DisplayGridlines = False
ActiveWindow.DisplayHeadings = False
```

---

## Krok 5 - code-behind

Pełny code w [`sh_Dashboard.code.txt`](sh_Dashboard.code.txt). Struktura:

- **`Worksheet_Activate`** — refresh user info + stats (called when Sheet1 becomes active, e.g., after form close)
- **`RefreshDashboard`** (private) — pisze do cells B3, B4, B5, B20, E20
- **5 button click handlers** — każdy `frm_XXX.Show vbModal`
- **Bez SelectionChange** (nie potrzebujemy trackować gdzie user klika)

---

## Krok 6 - Integracja z ThisWorkbook

Workbook_Open (patrz `Source/ThisWorkbook/ThisWorkbook.code.txt`):
- `Application.DisplayFullScreen = True` (app-like feel per user's decision)
- `sh_Dashboard.Activate` (na koniec, żeby user widział Sheet1 po zamknięciu picker/setup)

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
- Sheet1 wypełnia obszar poniżej ribbon (chyba że DisplayFullScreen=True → wypełnia cały ekran)
- Ustawienie window size przez `ActiveWindow.WindowState = xlMaximized` przy Workbook_Open zapewnia pełny obszar

---

## Diagnostic integration

`mod_Diagnostic.ExpectedSheets` — 3 → **4** (dodać `sh_Dashboard`).
`ExpectedSheetHeaders("sh_Dashboard")` — brak (nie tabela, tylko labels/buttons — skip check).
`AuditSheets` autoNote dla missing case: `"tworzone recznie w VBE - patrz sh_Dashboard.LAYOUT.md"`.

Opcjonalnie w przyszłości: audyt kontrolek ActiveX na Sheet1 (analogicznie do `AuditFormControls`). Skip dla M6/M7.
