# frm_Main — specyfikacja layoutu

> **Cel**: główny ekran użytkownika. Dodawanie pojedynczych zgłoszeń do batcha + lista pending + wysyłka batcha.
> Pokazywany przez `ThisWorkbook.Workbook_Open` gdy `IsSetupCompleted() = True`, oraz przez `frm_Setup.btn_Save_Click` po pierwszym setupie.
> **Plan**: M3.

---

## Krok 1 — utwórz UserForm |x

W VBE: **Insert → UserForm**. W oknie Properties (`F4`):

| Property | Wartość |
|---|---|
| `(Name)` | `frm_Main` |
| `Caption` | `BNC_Sender — Wniosek BNC` |
| `Width` | `640` |
| `Height` | `560` |
| `StartUpPosition` | `1 - CenterOwner` |

---

## Krok 2 — kontrolki |

### Sekcja: header (info o userze)  |x

| Typ | Name | Caption / Properties |
|---|---|---|
| Label | `lbl_UserInfo` | wypełnia `UserForm_Initialize` (np. "Zalogowany: Jan Kowalski (CNA: 12345, oddział: W001)") |x
| Label | `lbl_RoleInfo` | wypełnia `UserForm_Initialize` (np. "Tryb: HANDLOWIEC (wnioski wysyłane do kierownika)") |x

### Sekcja: nowe zgłoszenie |x

| Typ | Name | Caption / Properties |
|---|---|---|
| Label | `lbl_HeaderNew` | `Caption = "── Nowe zgłoszenie ──"` |x
| Label | `lbl_KlientFK` | `Caption = "Klient FK:"` |x
| TextBox | `txt_KlientFK` | `MaxLength = 6` | x --edit `MaxLength = 20` na `MaxLength = 6`
| Label | `lbl_NazwaKlienta` | `Caption = "Nazwa klienta:"` |x
| TextBox | `txt_NazwaKlienta` | `MaxLength = 200` |x
| Label | `lbl_MiesiacObrotu` | `Caption = "Miesiąc wykonania obrotu przez klienta:"` — **rename z `lbl_MiesiacZgloszenia`** (Schema v2) |x
| TextBox | `txt_MiesiacObrotu` | `MaxLength = 7` (format YYYY-MM, default = bieżący miesiąc) — **rename z `txt_MiesiacZgloszenia`** (Schema v2) |x
| CommandButton | `btn_Clear` | `Caption = "Wyczyść"` |x
| CommandButton | `btn_AddToList` | `Caption = "Dodaj do listy"`, `Default = True` |x

> **Usunięte w Schema v2** (nie dodawać do projektu):
> - ~~`lbl_Fields` (Caption "Pole dodatkowe:")~~
> - ~~`txt_Fields` (MultiLine TextBox)~~
> Kolumna `Fields` usunięta z `ws_DataCache` bez zastąpienia (YAGNI — analityka pokryta przez `CreatedTimestamp` + `MiesiacObrotu`).

### Sekcja: lista pending

| Typ | Name | Caption / Properties |
|---|---|---|
| Label | `lbl_BatchCount` | wypełnia `RefreshPendingList` (np. "Lista zgłoszeń do wysłania (3)") |x
| Label | `lbl_HdrID` | `Caption = "ID"` — **fake header** nad `lst_PendingBatch`, szerokość ~30pt, wyrównanie do 1. kolumny (`Font.Bold = True` opcjonalnie) |
| Label | `lbl_HdrFK` | `Caption = "Klient FK"` — szerokość ~60pt, 2. kolumna |
| Label | `lbl_HdrNazwa` | `Caption = "Nazwa klienta"` — szerokość ~220pt, 3. kolumna |
| Label | `lbl_HdrMiesiac` | `Caption = "Miesiąc obrotu"` — szerokość ~80pt, 4. kolumna |
| ListBox | `lst_PendingBatch` | `ColumnCount = 4`, **`ColumnHeads = False`**, `ColumnWidths = "30;60;220;80"`, `Height = 200`, `MultiSelect = 0 - fmMultiSelectSingle` (default) — **Schema v2: 5→4 kolumn** (usunięto Fields). Nagłówki przez **fake header bar** (4 labels wyżej) bo `ColumnHeads = True` w VBA UserForm nie działa z `.List = arr` (tylko z RowSource) |
| CommandButton | `btn_DeleteSelected` | `Caption = "Usuń zaznaczone"` — pozycja: pod ListBox po lewej |x

> **Zachowanie**:
> - Najnowszy rekord pojawia się na **górze** listy (newest first); user musi sam kliknąć row żeby zaznaczyć.
> - `btn_DeleteSelected.Enabled = False` gdy lista pusta (`pending.Count = 0`) — szary, nieklikalny.
> - Klik bez zaznaczenia → MsgBox info "Najpierw zaznacz zgłoszenie".
> - Klik z zaznaczeniem → MsgBox potwierdzenia z ID i nazwą klienta → po `Tak` hard delete + refresh listy.
> - Sent records: niemożliwe do usunięcia (defensywny check w `mod_DataCacheSync.DeleteRecord` zwraca `False`, frm_Log nie ma buttona delete). Patrz ADR-006.

### Sekcja: przyciski (na dole)

| Typ | Name | Caption / Properties |
|---|---|---|
| CommandButton | `btn_ShowLog` | `Caption = "Pokaż historię"` |x
| CommandButton | `btn_SendBatch` | `Caption = "Wyślij Wniosek BNC"` |x

---

## Krok 3 — wklej code-behind |x

W VBE prawy klik na `frm_Main` → **View Code**. Wklej zawartość `frm_Main.code-behind.txt`. |x

---

## Krok 4 — eksport do `Source/Forms/`

Po smoke teście: prawy klik → **Export File...** → `frm_Main.frm` (VBE doda `.frx`).
