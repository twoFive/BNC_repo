# frm_Log — specyfikacja layoutu

> **Cel**: historia wszystkich zgłoszeń (pending + sent) z statystykami i eksportem `BNC_DataCache.xlsx` do dowolnej lokalizacji.
> Pokazywany przez `frm_Main.btn_ShowLog_Click`.
> **Plan**: M5.2.

---

## Krok 1 — utwórz UserForm |x

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
| ListBox | `lst_AllRecords` | `ColumnCount = 7`, `ColumnHeads = False`, `ColumnWidths = "30;60;140;50;50;180;80"`, `Height = 360` — **7 kolumn** (post-2026-08-09: dodane CNA po Nazwa, multi-user visibility) |x
| CommandButton | `btn_Export` | `Caption = "Eksportuj do pliku"` |x
| CommandButton | `btn_Back` | `Caption = "Powrót do formularza"`, `Cancel = True` |x

ListBox kolumny: `ID | KlientFK | Nazwa | CNA | Status | Wysłany do | Data wysłania` (7 kolumn, CNA po Nazwa dla multi-user visibility)

---

## Krok 3 — wklej code-behind

`Source/Forms/frm_Log.code-behind.txt` |x

---

## Krok 4 — eksport

Prawy klik → **Export File...** → `Source/Forms/frm_Log.frm`.
