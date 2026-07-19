# TODO — dokończenie stanu M3.3 + M4 + M5

> **Stan bazowy**: audit `mod_Diagnostic.AuditFullProject` z **2026-07-19 14:20:44** — 6/8 modułów · 2/4 formularzy OK · 3/3 arkusze fizycznie istnieją (ale ws_UsersRegistry bez nagłówków).
> **Cel**: doprowadzić projekt do stanu **all-green** przed smoke testem multi-user.

---

## 🎯 Pięć kroków (kolejność ma znaczenie)

### ☐ 1. Uzupełnij nagłówki `ws_UsersRegistry`

Arkusz istnieje ale ma pusty wiersz 1. Wywołanie `GetUsersCount` triggeruje `EnsureRegistryHeader`, który uzupełni 13 nagłówków.

**W Immediate Window** (`Ctrl+G`):

```
?mod_UserCacheSync.GetUsersCount
```

Oczekiwane: zwraca `0`, a `ws_UsersRegistry` dostaje w wierszu 1: `UserID | Imie | Nazwisko | EmailHandlowca | CNA_HandlowcaID | NrOddzialu | EmailKierownika | EmailBNC | CacheFolderPath | DataRejestracji | SetupCompleted | DontShowSetupAgain | LastLogin`.

---

### ☐ 2. Utwórz **shell** `frm_Log` (bez kontrolek na razie)

**⚠ Musi być PRZED wklejaniem code-behind do frm_Main** — inaczej compile error na linii `frm_Main.btn_ShowLog_Click` → `frm_Log.Show`.

- VBE → **Insert → UserForm**
- W Properties (F4): `(Name)` → `frm_Log`
- **Nic więcej** — pusty shell wystarczy do kompilacji
- `Ctrl+S`

Verify: `frm_Log` pojawia się w drzewie Project Explorer z 0 lin. kodu.

---

### ☐ 3. Wklej code-behind do `frm_Main`

Shell `frm_Main` już istnieje z 0 linii kodu — dodaj kod-behind.

- Prawy klik na `frm_Main` w Project Explorer → **View Code**
- Otwórz `Source/Forms/frm_Main.code-behind.txt`
- Skopiuj **całą** zawartość (od `' ============` na górze do końca)
- Wklej do okna kodu `frm_Main` (zastąp puste `Option Explicit` jeśli tam jest)
- `Ctrl+S`

Oczekiwane po tym: `frm_Main` ma ~200 lin. + handlery `UserForm_Initialize`, `UserForm_Activate`, `btn_AddToList_Click`, `btn_Clear_Click`, `btn_SendBatch_Click`, `btn_ShowLog_Click`, `btn_DeleteSelected_Click`.

---

### ☐ 4. Import 3 modułów

VBE → **File → Import File**, kolejno:

| # | Plik | Efekt |
|---|---|---|
| 4a | `Source/Modules/mod_MailSender.bas` | Nowy moduł M4 — `SendBatch`, `DetermineRecipient` |
| 4b | `Source/Modules/mod_Export.bas` | Nowy moduł M5.1 — `ExportDataCache`, `GetSuggestedExportFileName` |
| 4c | `Source/Modules/mod_Tests.bas` | **Najpierw Remove** stary (288 lin., 5/7 testów) → **Import** nowy (powinien mieć 7/7 testów) |

`Ctrl+S` po każdym imporcie.

---

### ☐ 5. Rozbuduj `frm_Log` (kontrolki + code-behind)

Otwórz `Source/Forms/frm_Log.LAYOUT.md` — pełny spec kontrolek.

**Kontrolki do dodania**:

| Typ | Name | Właściwości kluczowe |
|---|---|---|
| Label | `lbl_Stats` | Caption pusty (wypełnia `LoadRecords`) |
| ListBox | `lst_AllRecords` | `ColumnCount = 6`, `ColumnHeads = False`, `ColumnWidths = "30;60;180;50;180;80"`, `Height = 360`, `MultiSelect = 0 - fmMultiSelectSingle` |
| CommandButton | `btn_Export` | Caption `"Eksportuj do pliku"` |
| CommandButton | `btn_Back` | Caption `"Powrót do formularza"`, `Cancel = True` |

**Code-behind**:
- Prawy klik `frm_Log` → View Code
- Skopiuj zawartość `Source/Forms/frm_Log.code-behind.txt`
- Wklej, `Ctrl+S`

Oczekiwane: `frm_Log` ~120 lin., handlery `UserForm_Activate`, `btn_Export_Click`, `btn_Back_Click`.

---

## ✅ Weryfikacja końcowa — rerun audit

Po zakończeniu 5 kroków:

```
mod_Diagnostic.AuditFullProject
```

**Oczekiwany all-green output**:

```
1. MODULY
  [OK]      mod_Utils              107 lin.   API: 10/10
  [OK]      mod_Validation         213 lin.   API: 8/8
  [OK]      mod_UserCacheSync      599 lin.   API: 13/13
  [OK]      mod_DataCacheSync      321 lin.   API: 6/6
  [OK]      mod_MailSender         ~200 lin.  API: 2/2
  [OK]      mod_Export             ~80 lin.   API: 2/2
  [OK]      mod_Tests              ~330 lin.  API: 7/7
  [OK]      mod_Diagnostic         708 lin.   API: 9/9
  --> 8/8 modulow obecnych

2. FORMULARZE
  [OK]      frm_Setup              141 lin.   Handlery: 5/5
  [OK]      frm_Main               ~200 lin.  Handlery: 7/7
  [OK]      frm_Log                ~120 lin.  Handlery: 3/3
  [OK]      frm_UserPicker         155 lin.   Handlery: 4/4
  --> 4/4 formularzy OK

3. ARKUSZE
  [OK]      ws_UserCache           (very hidden)   0 kluczy w kol.A
  [OK]      ws_DataCache           (very hidden)   Naglowki: 11/11
  [OK]      ws_UsersRegistry       (very hidden)   Naglowki: 13/13
  --> 3/3 arkuszy OK

4. THISWORKBOOK
  [OK]      Workbook_Open handler obecny
```

Suma linii kodu: ~2900 (z obecnych 2069 dorzucone ~830 z brakujących komponentów).

---

## 🧪 Smoke test po all-green

### ☐ A. Test automatyczny — `mod_Tests.RunAllTests`

Spodziewane: 5 sekcji testów (`mod_Utils`, `mod_UserCacheSync`, `mod_DataCacheSync`, `mod_Validation`, plus placeholders dla `mod_MailSender` i `mod_Export`) — wszystkie PASS, żadnego FAIL.

### ☐ B. Test manualny — flow pierwszego usera

1. **Zamknij plik xlsm**, otwórz ponownie
2. Bo Registry pusty (`GetUsersCount() = 0`) → oczekiwane: **`frm_Setup` od razu**, bez picker'a
3. Wypełnij formularz:
   - Imię: Jan
   - Nazwisko: Kowalski (Twoje dane testowe)
   - Email służbowy: jan@firma.pl
   - CNA: 12345
   - NrOddzialu: W001
   - Email kierownika: kierownik@firma.pl (albo swój email = tryb kierownika)
4. Kliknij `btn_CreateCacheFolder` → utworzy `C:\BNC_CacheFolder\`
5. Kliknij `btn_Save` → walidacja → `AddNewUser` → `UZYTKOWNIK_1_CNA12345` → przejście do `frm_Main`
6. Verify w Immediate:
   ```
   ?mod_UserCacheSync.CurrentUserID
   ?mod_UserCacheSync.GetUsersCount
   ?mod_UserCacheSync.IsUserManager
   ```
   Oczekiwane: `UZYTKOWNIK_1_CNA12345`, `1`, `False` (lub `True` jeśli emaile równe)

### ☐ C. Test manualny — flow drugiego usera + picker

1. W `frm_Main` zamknij formularz (X w rogu)
2. **Zamknij plik xlsm**, otwórz ponownie
3. Registry ma 1 usera → oczekiwane: **`frm_UserPicker`**
4. W ComboBox widzisz: `Jan Kowalski · CNA:12345`
5. Kliknij `btn_AddNew` → `PrepareForNewUser` → `frm_Setup` (pusty formularz)
6. Wypełnij innego usera (inny CNA!): Anna Nowak, CNA 67890, itp.
7. Save → `AddNewUser` → `UZYTKOWNIK_2_CNA67890` → `frm_Main`
8. Zamknij, otwórz plik ponownie
9. Picker pokazuje **2 pozycje** w ComboBox — wybierz jednego, `btn_SelectUser` → `SwitchUser` → `frm_Main`
10. Verify: `?mod_UserCacheSync.CurrentUserID` powinno pokazać wybranego

### ☐ D. Test manualny — btn_Cancel w picker'ze

1. Zamknij i otwórz plik
2. `frm_UserPicker` → kliknij `btn_Cancel`
3. Oczekiwane: MsgBox potwierdzenia "Czy na pewno zamknąć?"
4. Klik "Tak" → xlsm się zamyka (Excel nadal otwarty jeśli miałeś inne pliki)

---

## 🔄 Co dalej po smoke testach

Jeśli wszystko passuje:

1. **Eksport `.frm`/`.frx`** z VBE do `Source/Forms/` dla wszystkich 4 formularzy (Right click → Export File):
   - `frm_Setup.frm/frx`
   - `frm_Main.frm/frx`
   - `frm_Log.frm/frx`
   - `frm_UserPicker.frm/frx`
2. **Eksport `ThisWorkbook.cls`** → `Source/ThisWorkbook/`
3. `git add Source/Forms/*.frm Source/Forms/*.frx Source/ThisWorkbook/*.cls`
4. Commit ekportu — pełne odzwierciedlenie kodu z VBE w Source/
5. Aktualizacja `Working/BNC_Sender_v0.1.0.xlsm` snapshot w repo

Potem dalsze kroki wg `BNC_Sender_PlanWdrozenia_FazaA.md`:
- **M4 UAT** — testy handlowca/kierownika z prawdziwym Outlookiem
- **M5 UAT** — frm_Log + eksport `.xlsx`
- **M6** — polish, tab order, akceleratory, frm_Tutorial
- **M7** — release v1.0.0

---

## 🚨 Jeśli coś się psuje

Compile error po którymś kroku — pierwsze pytanie: **na której linii którego modułu**? VBE zaznaczy problematyczną linię po `Ctrl+F5` (Debug → Compile VBAProject). Sprawdź:

| Symptom | Prawdopodobna przyczyna |
|---|---|
| "Method or data member not found" na `frm_Log` | Krok 2 nie wykonany — shell nie istnieje |
| "Method or data member not found" na `mod_MailSender` / `mod_Export` | Krok 4 nie wykonany — import nie zrobiony |
| Empty audit — 0 modulow | Trust access do VBProject nie włączony (patrz `mod_Diagnostic` header komentarz) |
| `Test_mod_UserCacheSync` FAIL na `SetupCompleted` | Registry lub UserCache w niespójnym stanie — usuń zawartość obu arkuszy ręcznie, rerun test |

Wklej mi output nowego `AuditFullProject` po każdym kroku jeśli chcesz weryfikacji na bieżąco.
