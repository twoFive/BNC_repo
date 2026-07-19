# frm_UserPicker — specyfikacja layoutu

> **Cel**: wybór aktywnego użytkownika przy uruchamianiu aplikacji. Pokazywany przez `ThisWorkbook.Workbook_Open` gdy `mod_UserCacheSync.GetUsersCount() > 0`.
> Przy pierwszym uruchomieniu (`GetUsersCount() = 0`) picker jest **pomijany** — od razu `frm_Setup` (zgodnie z wymaganiem M3.3 Q1: pierwszy user nie widzi picker'a).
> **Plan**: M3.3.

---

## Krok 1 — utwórz UserForm

W VBE: **Insert → UserForm**. W oknie Properties (`F4`):

| Property | Wartość |
|---|---|
| `(Name)` | `frm_UserPicker` |
| `Caption` | `BNC_Sender — Wybór użytkownika` |
| `Width` | `440` |
| `Height` | `380` |
| `StartUpPosition` | `1 - CenterOwner` |

---

## Krok 2 — kontrolki

### Sekcja: header

| Typ | Name | Caption / Properties |
|---|---|---|
| Label | `lbl_Header` | `Caption = "Wybierz użytkownika z listy lub dodaj nowego:"`, `Font.Size = 11` |

### Sekcja: lista użytkowników

| Typ | Name | Caption / Properties |
|---|---|---|
| ListBox | `lst_Users` | `ColumnCount = 3`, `ColumnHeads = True`, `ColumnWidths = "120;120;100"`, `Height = 200`, `MultiSelect = 0 - fmMultiSelectSingle` (default) |

Kolumny (nagłówki ustawiane w code-behind — patrz `PopulateUserList`):
1. **Imię**
2. **Nazwisko**
3. **CNA**

> **Design (M3.3 Q3 — updated)**: 3 kolumny — Imię, Nazwisko, CNA. CNA jest unikalne per handlowca w firmie, więc rozróżnia userów pewniej niż email (który może być roboczy vs personalny). Email i Oddział są w Registry, ale nie w picker'ze.

### Sekcja: przyciski (na dole)

| Typ | Name | Caption / Properties |
|---|---|---|
| CommandButton | `btn_SelectUser` | `Caption = "Wybierz i uruchom"`, `Default = True` |
| CommandButton | `btn_AddNew` | `Caption = "Dodaj nowego użytkownika"` |
| CommandButton | `btn_Cancel` | `Caption = "Anuluj"`, `Cancel = True` |

> **Design (M3.3 Q4)**: `btn_Cancel` zamyka **plik xlsm** (nie całego Excela) z ostrzegawczym `MsgBox`. Zapobiega działaniu aplikacji bez wybranego usera.

---

## Krok 3 — wklej code-behind

W VBE prawy klik na `frm_UserPicker` → **View Code** → wklej zawartość `frm_UserPicker.code-behind.txt`.

---

## Krok 4 — eksport do `Source/Forms/`

Po smoke teście: prawy klik → **Export File...** → `Source/Forms/frm_UserPicker.frm` (VBE doda `.frx`).
