# frm_Log — specyfikacja layoutu

> **Cel**: historia wszystkich zgłoszeń (pending + sent) z statystykami i eksportem `BNC_DataCache.xlsx` do dowolnej lokalizacji.
> Pokazywany przez `frm_Main.btn_ShowLog_Click`.
> **Plan**: M5.2.

---

## Krok 1 — utwórz UserForm

| Property | Wartość |
|---|---|
| `(Name)` | `frm_Log` |
| `Caption` | `BNC_Sender — Historia zgłoszeń` |
| `Width` | `720` |
| `Height` | `520` |
| `StartUpPosition` | `1 - CenterOwner` |

---

## Krok 2 — kontrolki

| Typ | Name | Caption / Properties |
|---|---|---|
| Label | `lbl_Stats` | wypełnia `LoadRecords` (np. "Wszystkich: 47   \|   Pending: 3   \|   Sent: 44") |
| ListBox | `lst_AllRecords` | `ColumnCount = 6`, `ColumnHeads = False`, `ColumnWidths = "30;60;180;50;180;80"`, `Height = 360` |
| CommandButton | `btn_Export` | `Caption = "Eksportuj do pliku"` |
| CommandButton | `btn_Back` | `Caption = "Powrót do formularza"`, `Cancel = True` |

ListBox kolumny: `ID | KlientFK | Nazwa | Status | Wysłany do | Data wysłania`

---

## Krok 3 — wklej code-behind

`Source/Forms/frm_Log.code-behind.txt`

---

## Krok 4 — eksport

Prawy klik → **Export File...** → `Source/Forms/frm_Log.frm`.
