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
| `Height` | `220` |
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
| ComboBox | `cmb_Users` | `Style = 2 - fmStyleDropDownList` (forced selection z listy, brak wpisywania tekstu), `Width = 400` |

Format wyświetlania per pozycja (obliczany w code-behind, `PopulateUserList`):

```
Imię Nazwisko · CNA:<CNA_HandlowcaID>
```

Przykład: `Jan Kowalski · CNA:12345`, `Anna Nowak · CNA:67890`

> **Design (M3.3 Q3 — updated ×2)**: rozwijana lista (ComboBox) zamiast ListBox. Jedno pole obliczeniowe z konkatenacji Imię, Nazwisko, CNA (label) i numeru CNA. Kompaktowa forma — zwiń jeśli 10+ userów, wygodne dla codziennego użycia (jeden klik do rozwinięcia, jeden do wyboru).

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
