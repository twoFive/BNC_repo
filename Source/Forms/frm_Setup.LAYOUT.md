# frm_Setup — specyfikacja layoutu

> **Cel**: jednorazowa rejestracja handlowca przy pierwszym otwarciu pliku.
> Pokazywany przez `ThisWorkbook.Workbook_Open`, gdy `mod_UserCacheSync.IsSetupCompleted() = False`.
> **Plan**: M2.2.

---

## Krok 1 — utwórz UserForm

W VBE: **Insert → UserForm**. W oknie Properties (`F4`):

| Property | Wartość |
|---|---|
| `(Name)` | `frm_Setup` |
| `Caption` | `BNC_Sender — Konfiguracja wstępna` |
| `Width` | `480` |
| `Height` | `560` |
| `StartUpPosition` | `1 - CenterOwner` |

---

## Krok 2 — kontrolki

Dodawaj z **Toolbox** (`Ctrl+T` jeśli niewidoczny). Dla każdej kontrolki ustaw `(Name)` dokładnie tak, jak niżej — code-behind się tym posługuje.

### Sekcja: dane podstawowe

| Typ | Name | Caption / Properties |
|---|---|---|
| Label | `lbl_Header` | `Caption = "Witaj. Aby rozpocząć, podaj swoje dane służbowe."` |
| Label | `lbl_Imie` | `Caption = "Imię:"` |
| TextBox | `txt_Imie` | `MaxLength = 100` |
| Label | `lbl_Nazwisko` | `Caption = "Nazwisko:"` |
| TextBox | `txt_Nazwisko` | `MaxLength = 100` |
| Label | `lbl_EmailHandlowca` | `Caption = "Email służbowy:"` |
| TextBox | `txt_EmailHandlowca` | `MaxLength = 200` |
| Label | `lbl_CNA` | `Caption = "CNA (numer handlowca):"` |
| TextBox | `txt_CNA` | `MaxLength = 20` |
| Label | `lbl_NrOddzialu` | `Caption = "Numer oddziału:"` |
| TextBox | `txt_NrOddzialu` | `MaxLength = 20` |

### Sekcja: adresy do wysyłki

| Typ | Name | Caption / Properties |
|---|---|---|
| Label | `lbl_HeaderEmails` | `Caption = "── Adresy do wysyłki ──"` |
| Label | `lbl_EmailKierownika` | `Caption = "Email kierownika:"` |
| TextBox | `txt_EmailKierownika` | `MaxLength = 200` |
| Label | `lbl_HintKierownik` | `Caption = "ℹ Jeśli jesteś kierownikiem, wpisz tu swój własny adres."` (mniejsza czcionka, np. `Font.Size = 8`) |
| Label | `lbl_EmailBNC` | `Caption = "Email zespołu BNC:"` |
| TextBox | `txt_EmailBNC` | `MaxLength = 200` |

### Sekcja: lokalizacja cache

| Typ | Name | Caption / Properties |
|---|---|---|
| Label | `lbl_HeaderCache` | `Caption = "── Lokalizacja plików cache ──"` |
| Label | `lbl_CacheFolderPath` | `Caption = "Folder cache:"` |
| TextBox | `txt_CacheFolderPath` | `MaxLength = 260` |
| CommandButton | `btn_Browse` | `Caption = "Przeglądaj..."` |

### Sekcja: samouczek

| Typ | Name | Caption / Properties |
|---|---|---|
| Label | `lbl_HeaderTutorial` | `Caption = "── Samouczek ──"` |
| TextBox | `txt_Tutorial` | `MultiLine = True`, `ScrollBars = 2 - fmScrollBarsVertical`, `Locked = True`, `Height = 120` |
| CheckBox | `chk_DontShowTutorial` | `Caption = "Nie pokazuj samouczka ponownie"` |

### Sekcja: przyciski (na dole)

| Typ | Name | Caption / Properties |
|---|---|---|
| CommandButton | `btn_Cancel` | `Caption = "Anuluj"`, `Cancel = True` |
| CommandButton | `btn_Save` | `Caption = "Zapisz konfigurację"`, `Default = True` |

---

## Krok 3 — wklej code-behind

W VBE prawy klik na `frm_Setup` w Project Explorer → **View Code**. Wklej zawartość pliku `frm_Setup.code-behind.txt`.

---

## Krok 4 — eksport do `Source/Forms/`

Po sprawdzeniu, że formularz działa: prawy klik na `frm_Setup` → **Export File...** → `C:\Dev\BNC_Sender\FazaA\Source\Forms\frm_Setup.frm` (VBE automatycznie wygeneruje też `frm_Setup.frx`).

Wtedy `git add Source/Forms/ && git commit`.
