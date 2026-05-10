# frm_Setup — specyfikacja layoutu

> **Cel**: jednorazowa rejestracja handlowca przy pierwszym otwarciu pliku.
> Pokazywany przez `ThisWorkbook.Workbook_Open`, gdy `mod_UserCacheSync.IsSetupCompleted() = False`.
> **Plan**: M2.2.

---

## Krok 1 — utwórz UserForm

W VBE: **Insert → UserForm**. W oknie Properties (`F4`):

| Property | Wartość |
|---|---|
| `(Name)` | `frm_Setup` |x
| `Caption` | `BNC_Sender — Konfiguracja wstępna` |x
| `Width` | `480` |x
| `Height` | `560` |x
| `StartUpPosition` | `1 - CenterOwner` |x

---

## Krok 2 — kontrolki

Dodawaj z **Toolbox** (`Ctrl+T` jeśli niewidoczny). Dla każdej kontrolki ustaw `(Name)` dokładnie tak, jak niżej — code-behind się tym posługuje.

### Sekcja: dane podstawowe

| Typ | Name | Caption / Properties |
|---|---|---|
| Label | `lbl_Header` | `Caption = "Witaj. Aby rozpocząć, podaj swoje dane służbowe."` |x
| Label | `lbl_Imie` | `Caption = "Imię:"` |x
| TextBox | `txt_Imie` | `MaxLength = 100` |x
| Label | `lbl_Nazwisko` | `Caption = "Nazwisko:"` |x
| TextBox | `txt_Nazwisko` | `MaxLength = 100` |x
| Label | `lbl_EmailHandlowca` | `Caption = "Email służbowy:"` |x
| TextBox | `txt_EmailHandlowca` | `MaxLength = 200` |x
| Label | `lbl_CNA` | `Caption = "CNA (numer handlowca):"` |X
| TextBox | `txt_CNA` | `MaxLength = 20` |x
| Label | `lbl_NrOddzialu` | `Caption = "Numer oddziału:"` |x
| TextBox | `txt_NrOddzialu` | `MaxLength = 20` |x

### Sekcja: adresy do wysyłki

| Typ | Name | Caption / Properties |
|---|---|---|
| Label | `lbl_HeaderEmails` | `Caption = "── Adresy do wysyłki ──"` |x
| Label | `lbl_EmailKierownika` | `Caption = "Email kierownika:"` |x
| TextBox | `txt_EmailKierownika` | `MaxLength = 200` |x
| Label | `lbl_HintKierownik` | `Caption = "ℹ Jeśli jesteś kierownikiem, wpisz tu swój własny adres."` (mniejsza czcionka, np. `Font.Size = 8`) |x
| Label | `lbl_EmailBNC` | `Caption = "Email zespołu BNC:"` | X
| TextBox | `txt_EmailBNC` | `MaxLength = 200`, `Text = "jessica.cant@swim.omg"`, `Locked = True`, `BackColor = &H8000000F&` (szare tło) |X

> **Polityka**: `txt_EmailBNC` jest **hardcoded i tylko-do-odczytu**. Wartość ustawiana też w code-behind przy `UserForm_Initialize` jako bezpiecznik (gdyby user zmienił `Locked` w designerze). User nie ma możliwości edytować adresu BNC. Patrz ADR-003. 

### Sekcja: lokalizacja cache

> **Polityka**: ścieżka folderu cache jest **hardcoded** na `C:\BNC_CacheFolder\`. User nie wybiera lokalizacji (niezawodność > elastyczność — nie ufamy że user wskaże poprawnie). Zamiast browse pickera dajemy przycisk który tworzy folder skryptem zgodnie z architekturą. Patrz ADR-003.

| Typ | Name | Caption / Properties |
|---|---|---|
| Label | `lbl_HeaderCache` | `Caption = "── Lokalizacja plików cache ──"` |X
| ~~Label `lbl_CacheFolderPath`~~ | — | **USUŃ** label który dodałeś (zastąpiony buttonem niżej) |
| CommandButton | `btn_CreateCacheFolder` | `Caption = "Utwórz folder cache na dysku C:"` | X
| TextBox | `txt_CacheFolderPath` | `MaxLength = 260`, `Text = "C:\BNC_CacheFolder\"`, `Locked = True`, `Enabled = False`, `BackColor = &H8000000F&` |X
| ~~CommandButton `btn_Browse`~~ | — | **POMIŃ** — usunięty (niepotrzebny przy hardcoded path) |

### Sekcja: pomoc + preferencje

> **Design**: samouczek **nie jest inline** w `frm_Setup` — `btn_ShowTutorial` (docelowo) otworzy osobny formularz `frm_Tutorial`. W M2 `frm_Tutorial` jeszcze nie istnieje — przycisk pokazuje MsgBox-placeholder. Implementacja `frm_Tutorial` jest **odłożona do końca M6/M7** (pre-release polish).
>
> Checkbox `chk_DontShowSetupAgain` kontroluje czy `frm_Setup` ma być pokazywany ponownie przy następnym otwarciu pliku. Wartość zapisywana w UserCache jako `DontShowSetupAgain`. Routing rzeczywisty robi `ThisWorkbook.Workbook_Open` (M2.3) — pole `IsSetupCompleted` już dziś gwarantuje skip setupu po pierwszym ukończeniu; ten flag istnieje jako jawna preferencja usera (rezerwa na przyszłe scenariusze typu "force re-setup z menu").

| Typ | Name | Caption / Properties |
|---|---|---|
| Label | `lbl_HeaderTutorial` | `Caption = "── Samouczek ──"` |
| CommandButton | `btn_ShowTutorial` | `Caption = "Pokaż samouczek"` |x
| CheckBox | `chk_DontShowSetupAgain` | `Caption = "Nie pokazuj tego formularza ponownie"` |
| ~~TextBox `txt_Tutorial`~~ | — | **POMIŃ** — wyniesione do osobnego `frm_Tutorial` (M6/M7) |

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
