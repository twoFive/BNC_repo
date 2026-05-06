# frm_Main — specyfikacja layoutu

> **Cel**: główny ekran użytkownika. Dodawanie pojedynczych zgłoszeń do batcha + lista pending + wysyłka batcha.
> Pokazywany przez `ThisWorkbook.Workbook_Open` gdy `IsSetupCompleted() = True`, oraz przez `frm_Setup.btn_Save_Click` po pierwszym setupie.
> **Plan**: M3.

---

## Krok 1 — utwórz UserForm

W VBE: **Insert → UserForm**. W oknie Properties (`F4`):

| Property | Wartość |
|---|---|
| `(Name)` | `frm_Main` |
| `Caption` | `BNC_Sender — Wniosek BNC` |
| `Width` | `640` |
| `Height` | `560` |
| `StartUpPosition` | `1 - CenterOwner` |

---

## Krok 2 — kontrolki

### Sekcja: header (info o userze)

| Typ | Name | Caption / Properties |
|---|---|---|
| Label | `lbl_UserInfo` | wypełnia `UserForm_Initialize` (np. "Zalogowany: Jan Kowalski (CNA: 12345, oddział: W001)") |
| Label | `lbl_RoleInfo` | wypełnia `UserForm_Initialize` (np. "Tryb: HANDLOWIEC (wnioski wysyłane do kierownika)") |

### Sekcja: nowe zgłoszenie

| Typ | Name | Caption / Properties |
|---|---|---|
| Label | `lbl_HeaderNew` | `Caption = "── Nowe zgłoszenie ──"` |
| Label | `lbl_KlientFK` | `Caption = "Klient FK:"` |
| TextBox | `txt_KlientFK` | `MaxLength = 20` |
| Label | `lbl_NazwaKlienta` | `Caption = "Nazwa klienta:"` |
| TextBox | `txt_NazwaKlienta` | `MaxLength = 200` |
| Label | `lbl_MiesiacZgloszenia` | `Caption = "Miesiąc zgłoszenia:"` |
| TextBox | `txt_MiesiacZgloszenia` | `MaxLength = 7` (format YYYY-MM, default = bieżący miesiąc) |
| Label | `lbl_Fields` | `Caption = "Pole dodatkowe:"` |
| TextBox | `txt_Fields` | `MultiLine = True`, `Height = 60`, `MaxLength = 1000` |
| CommandButton | `btn_Clear` | `Caption = "Wyczyść"` |
| CommandButton | `btn_AddToList` | `Caption = "Dodaj do listy"`, `Default = True` |

### Sekcja: lista pending

| Typ | Name | Caption / Properties |
|---|---|---|
| Label | `lbl_BatchCount` | wypełnia `RefreshPendingList` (np. "Lista zgłoszeń do wysłania (3)") |
| ListBox | `lst_PendingBatch` | `ColumnCount = 5`, `ColumnHeads = True`, `ColumnWidths = "30;60;180;60;200"`, `Height = 200` |

### Sekcja: przyciski (na dole)

| Typ | Name | Caption / Properties |
|---|---|---|
| CommandButton | `btn_ShowLog` | `Caption = "Pokaż historię"` |
| CommandButton | `btn_SendBatch` | `Caption = "Wyślij Wniosek BNC"` |

---

## Krok 3 — wklej code-behind

W VBE prawy klik na `frm_Main` → **View Code**. Wklej zawartość `frm_Main.code-behind.txt`.

---

## Krok 4 — eksport do `Source/Forms/`

Po smoke teście: prawy klik → **Export File...** → `frm_Main.frm` (VBE doda `.frx`).
