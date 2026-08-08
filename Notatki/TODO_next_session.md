# TODO — dokończenie stanu M3.3 + M4 + M5

> **Stan bazowy**: audit `mod_Diagnostic.AuditFullProject` z **2026-07-19 14:20:44** — 6/8 modułów · 2/4 formularzy OK · 3/3 arkusze fizycznie istnieją (ale ws_UsersRegistry bez nagłówków).
> **Cel**: doprowadzić projekt do stanu **all-green** przed smoke testem multi-user.

---

## 🎯 Pięć kroków (kolejność ma znaczenie)

### ☐ 1. Uzupełnij nagłówki `ws_UsersRegistry` |x

Arkusz istnieje ale ma pusty wiersz 1. Wywołanie `GetUsersCount` triggeruje `EnsureRegistryHeader`, który uzupełni 13 nagłówków.

**W Immediate Window** (`Ctrl+G`):

```
?mod_UserCacheSync.GetUsersCount
```

Oczekiwane: zwraca `0`, a `ws_UsersRegistry` dostaje w wierszu 1: `UserID | Imie | Nazwisko | EmailHandlowca | CNA_HandlowcaID | NrOddzialu | EmailKierownika | EmailBNC | CacheFolderPath | DataRejestracji | SetupCompleted | DontShowSetupAgain | LastLogin`.

---

### ☐ 2. Utwórz **shell** `frm_Log` (bez kontrolek na razie) |x

**⚠ Musi być PRZED wklejaniem code-behind do frm_Main** — inaczej compile error na linii `frm_Main.btn_ShowLog_Click` → `frm_Log.Show`.

- VBE → **Insert → UserForm**
- W Properties (F4): `(Name)` → `frm_Log`
- **Nic więcej** — pusty shell wystarczy do kompilacji
- `Ctrl+S`

Verify: `frm_Log` pojawia się w drzewie Project Explorer z 0 lin. kodu.

---

### ☐ 3. Wklej code-behind do `frm_Main` |x

Shell `frm_Main` już istnieje z 0 linii kodu — dodaj kod-behind.

- Prawy klik na `frm_Main` w Project Explorer → **View Code**
- Otwórz `Source/Forms/frm_Main.code-behind.txt`
- Skopiuj **całą** zawartość (od `' ============` na górze do końca)
- Wklej do okna kodu `frm_Main` (zastąp puste `Option Explicit` jeśli tam jest)
- `Ctrl+S`

Oczekiwane po tym: `frm_Main` ma ~200 lin. + handlery `UserForm_Initialize`, `UserForm_Activate`, `btn_AddToList_Click`, `btn_Clear_Click`, `btn_SendBatch_Click`, `btn_ShowLog_Click`, `btn_DeleteSelected_Click`.

---

### ☐ 4. Import 3 modułów |x

VBE → **File → Import File**, kolejno:

| # | Plik | Efekt |
|---|---|---|
| 4a | `Source/Modules/mod_MailSender.bas` | Nowy moduł M4 — `SendBatch`, `DetermineRecipient` |x
| 4b | `Source/Modules/mod_Export.bas` | Nowy moduł M5.1 — `ExportDataCache`, `GetSuggestedExportFileName` |x
| 4c | `Source/Modules/mod_Tests.bas` | **Najpierw Remove** stary (288 lin., 5/7 testów) → **Import** nowy (powinien mieć 7/7 testów) |x

`Ctrl+S` po każdym imporcie.

---

### ☐ 5. Rozbuduj `frm_Log` (kontrolki + code-behind) |x

Otwórz `Source/Forms/frm_Log.LAYOUT.md` — pełny spec kontrolek.

**Kontrolki do dodania**:

| Typ | Name | Właściwości kluczowe |
|---|---|---|
| Label | `lbl_Stats` | Caption pusty (wypełnia `LoadRecords`) |x
| ListBox | `lst_AllRecords` | `ColumnCount = 6`, `ColumnHeads = False`, `ColumnWidths = "30;60;180;50;180;80"`, `Height = 360`, `MultiSelect = 0 - fmMultiSelectSingle` |x
| CommandButton | `btn_Export` | Caption `"Eksportuj do pliku"` |x
| CommandButton | `btn_Back` | Caption `"Powrót do formularza"`, `Cancel = True` |x

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

---

## 🔧 UPDATE (2026-07-25) — po fixie ComboBox picker

Podczas smoke testu wyszło że `frm_UserPicker` miał ComboBox nazwany `ComboBox1` (VBE default) zamiast `cmb_Users` — compile error na `PopulateUserList`. **Fix**: Properties → `(Name)` = `cmb_Users`. Bug typu "zapomniałem zmienić Name po drag&drop kontrolki z Toolbox".

### ☐ 4d. Ponowny re-import `mod_Tests.bas`

Dodałem **`Test_MultiUser`** (regresja dla M3.3 Registry: `AddNewUser` → `SwitchUser` → `GetAllUsers`). `mod_Tests` który zaimportowałeś w kroku 4c ma 7 procedur — nowy ma **8**.

- Right-click `mod_Tests` w VBE → **Remove mod_Tests** (kliknij **No** przy "Do you want to export?")
- **File → Import File** → `Source/Modules/mod_Tests.bas`
- `Ctrl+S`

Po tym:
- `mod_Diagnostic.ExpectedPublicProcs("mod_Tests")` oczekuje 8/8
- Nowy audit powinien pokazać `[OK] mod_Tests · API: 8/8`

### ☐ Update oczekiwanego all-green output (linia 105)

Zaktualizuj mentalnie:
```
[OK]      mod_Tests              ~430 lin.  API: 8/8   ← było 7/7, dodane Test_MultiUser
```

### ☐ Test E — Smoke test ComboBox picker (specyficzny po fixie)

**Po dokończeniu kroku 5** (frm_Log kontrolki + code-behind):

1. Zamknij i otwórz plik xlsm
2. Oczekiwane: **`frm_UserPicker` otwiera się** (bo `GetUsersCount() ≥ 1`)
3. **Verify ComboBox behavior**:
   - Widzisz swój profil w rozwijanej liście: **`Imietest · CNA:111`** (lub jak nazwałeś testowego usera)
   - Domyślnie **zaznaczony** ten user (wskazówka: `LastLogin` = najnowsze)
   - Kliknij strzałkę w dół — pojawia się lista, użytkownik widoczny
   - **Nie da się wpisać** własnego tekstu do pola ComboBox (dzięki `Style = 2 - fmStyleDropDownList`)
4. Kliknij **"Wybierz i uruchom"** → oczekiwane: przejście do `frm_Main` (już ma code-behind z kroku 3)
5. Verify: `?mod_UserCacheSync.CurrentUserID` — pokazuje wybranego usera

Jeśli któryś krok padnie — wklej dokładnie który.

### ☐ Test F — Smoke test Test_MultiUser (nowy test regresyjny)

W Immediate:
```
mod_Tests.Test_MultiUser
```

Oczekiwane: **~10 asercji PASS**:
- `AddNewUser returns non-empty` = True
- `UserID format UZYTKOWNIK_*_CNA999999` = True
- `GetUsersCount incremented` = origCount+1
- `CurrentUserID = newUserId (auto-switch)` = newUserId
- `UserCache.Imie = _TEST_` = "_TEST_"
- `UserCache.CNA = 999999` = 999999
- `IsUserManager = True (email kierownika=handlowca)` = True
- `GetAllUsers contains new UserID` = True
- `GetAllUsers new user Imie/CNA`
- `SwitchUser back: CurrentUserID` = origUserId (jeśli byl)

⚠ **Test dodaje wiersz do `ws_UsersRegistry` z Imie=`_TEST_` i CNA=`999999`** — NIE kasuje automatycznie. Aby posprzątać:
1. Tymczasowo pokaż arkusz: w VBE → `ws_UsersRegistry` → Properties → `Visible` = `-1 - xlSheetVisible`
2. Znajdź wiersz z Imie `_TEST_` + CNA 999999, usuń go
3. Wróć `Visible` = `2 - xlSheetVeryHidden`

Alternatywnie zostaw jako drugiego testowego usera — nie przeszkadza.

### ☐ (NEW 2026-07-26) Test H — AuditFormControls per formularz

**Motywacja**: Compile error "variable not defined" na `lbl_UserInfo` w `frm_Main.UserForm_Initialize` — identyczny bug jak z `ComboBox1`/`cmb_Users` w `frm_UserPicker`. `AuditFullProject` pokazywał `Handlery: 7/7` ale kontrolka nie istniała w Designerze. **Luka systemowa**: audit nie sprawdzał kontrolek. **Fix**: nowa funkcja `mod_Diagnostic.AuditFormControls`.

**Wymaga re-importu `mod_Diagnostic.bas`** — API wzrosło z 9 do 10 procedur (dodane `AuditFormControls` + prywatne helpery `ExpectedFormControls`, `CheckFormControls`).

**Uruchomienie #1 — full audit z nową sekcją Kontrolki**:
```
mod_Diagnostic.AuditFullProject
```
Teraz per formularz drukuje **dwie linie** — Handlery + Kontrolki:
```
[OK]  frm_Main   230 lin.  Handlery: 7/7
                           Kontrolki: 13/13
```
Jeśli któraś kontrolka jest MISSING w Designerze:
```
[OK]  frm_Main   230 lin.  Handlery: 7/7
                           Kontrolki: 12/13  (brak: lbl_UserInfo)
```

**Uruchomienie #2 — szczegół jednego formularza** (najbardziej użyteczne przy diagnozie compile error):
```
mod_Diagnostic.AuditFormControls "frm_Main"
```
Wypisuje FAKTYCZNE kontrolki z Designera (`Name` + `TypeName`) + OCZEKIWANE + diff. Znajdziesz np. `Label1` zamiast `lbl_UserInfo` — zmień Name w Properties (F4).

**Wymaganie**: "Trust access to the VBA project object model" musi być włączone (ten sam wymóg co inne `Audit*`).

**Lesson learned**: klasa bugów "kontrolka nie została dodana lub ma default name (Label1/ComboBox1)" — bardzo częsta w VBA UserForm dev. Audit code-behind nie wystarczy, bo compile error łapie dopiero w runtime. `AuditFormControls` łapie to statycznie z Designera.

### ☐ Test G — Smoke test drugiego usera flow (rozszerzone)

**Po Test F** (Registry ma teraz 2 userów: Imietest + _TEST_):

1. Zamknij i otwórz plik
2. `frm_UserPicker` pokazuje **2 pozycje** w ComboBox
3. Kliknij `btn_AddNew` → `frm_Setup` (pusty formularz — `PrepareForNewUser` wyczyścił UserCache)
4. Anuluj (lub wypełnij trzeciego usera)
5. Verify multi-user routing działa end-to-end

---

## 🆕 Schema v2 migration (2026-07-26)

**Kontekst**: usunięto kolumnę `Fields` z `ws_DataCache`, rename `MiesiacZgloszenia` → `MiesiacObrotu` (biznesowo jaśniej — "kiedy klient wykonał obrót"). `DataZgloszenia` rozważane i **odrzucone** (redundancja z `CreatedTimestamp` — YAGNI). Schema: 11 → 10 kolumn.

**Wpływ**: breaking change dla `BNC_DataCache.xlsx`. Brak migration code w aplikacji (Faza A, 0 productionowych userów).

### ☐ S1. Manual delete starego cache

1. **Zamknij Excel całkowicie** (proces `EXCEL.EXE`, nie tylko xlsm)
2. Explorer: usuń `C:\BNC_CacheFolder\BNC_DataCache.xlsx`

### ☐ S2. Re-import wszystkich zmienionych modułów

VBE → prawy klik → Remove → No → Import File. Zaimportuj **ponownie**:
- `Source/Modules/mod_DataCacheSync.bas` — schema constants, AppendRecord, EnsureHeader, GetRecordsWhereStatus
- `Source/Modules/mod_Validation.bas` — usunięte `MAX_FIELDS` + Fields validation, rename klucza `MiesiacObrotu`
- `Source/Modules/mod_MailSender.bas` — `GenerateTempFile` 8→7 kolumn
- `Source/Modules/mod_Tests.bas` — testowe rekordy shape v2
- `Source/Modules/mod_Diagnostic.bas` — `ExpectedSheetHeaders("ws_DataCache")` 10 nazw + `ExpectedFormControls("frm_Main")` bez `txt_Fields`

### ☐ S3. `frm_Main` — Designer changes

W VBE otwórz `frm_Main` w Designerze:
- **Usuń** kontrolki:
  - `lbl_Fields`
  - `txt_Fields`
- **Rename** kontrolek (Properties → `(Name)`):
  - `lbl_MiesiacZgloszenia` → `lbl_MiesiacObrotu`
  - `txt_MiesiacZgloszenia` → `txt_MiesiacObrotu`
- **Update Caption** `lbl_MiesiacObrotu`: `"Miesiąc wykonania obrotu przez klienta:"`
- **Update `lst_PendingBatch` properties** (Properties → F4):
  - `ColumnCount = 4` (było 5)
  - `ColumnWidths = "30;60;220;80"` (było `"30;60;180;60;200"`)

### ☐ S4. `frm_Main` — Re-paste code-behind

VBE → View Code na `frm_Main` → `Ctrl+A` → Delete → wklej całą zawartość z `Source/Forms/frm_Main.code-behind.txt`. Zmiany:
- `UserForm_Initialize` — `txt_MiesiacObrotu.Text = ...`
- `btn_AddToList_Click` — usunięto `reportData("Fields")`, rename klucza `MiesiacObrotu`
- `ClearFormFields` — usunięto `txt_Fields.Text = ""`, komentarz `MiesiacObrotu`
- `RefreshPendingList` — ListBox `ReDim arr(0 To ..., 0 To 3)` (4 kolumny), usunięto Fields column

### ☐ S5. `Ctrl+S` (save xlsm)

### ☐ S6. Zamknij i otwórz xlsm

- `Workbook_Open` → `EnsureCacheFileExists` widzi że plik nie istnieje → tworzy **fresh** `BNC_DataCache.xlsx` z 10 kolumnami schema v2.

### ☐ S7. Weryfikacja — Immediate Window

```
mod_Diagnostic.AuditFullProject
```

Oczekiwane:
```
[OK]      ws_DataCache          (very hidden)     Naglowki: 10/10  Rows: 0
[OK]      frm_Main              XXX lin.  Handlery: 7/7
                                          Kontrolki: 12/12   ← było 13/13 (bez txt_Fields)
```

```
mod_Diagnostic.AuditFormControls "frm_Main"
```

Oczekiwane 12/12 obecnych, żadnego MISSING. FAKTYCZNE listuje `txt_MiesiacObrotu` (nie `txt_MiesiacZgloszenia`), brak `txt_Fields`/`lbl_Fields`.

```
mod_Tests.RunAllTests
```

Oczekiwane: ~76 asercji PASS (bez zmiany liczby, bo Fields test usunięty ale nowe testy nie dodane — schema v2 to uproszczenie, nie rozszerzenie).

### ☐ S8. Manual smoke test w `frm_Main`

1. Wpisz KlientFK, NazwaKlienta, MiesiacObrotu (default = bieżący miesiąc) — kliknij "Dodaj do listy"
2. Verify: ListBox pokazuje 4 kolumny (ID, KlientFK, Nazwa, Miesiąc obrotu) — brak Fields
3. Dodaj 2-3 zgłoszenia
4. `btn_DeleteSelected` — usuń jedno pending
5. `btn_ShowLog` → verify frm_Log otwiera się (jego ListBox 6 kolumn: ID, KlientFK, Nazwa, Status, Wysłany, Data — nie dotknięty przez schema v2, patrz LAYOUT)
6. `btn_Back` → wraca do frm_Main
7. Klik "Wyślij Wniosek BNC" (jeśli masz skonfigurowany Outlook) — mail leci z załącznikiem xlsx **7 kolumn** biznesowych
8. `mod_Tests.Test_MultiUser` — nadal PASS (bez zmian, dotyczy Registry nie DataCache)

### 🚫 Co się NIE zmienia

- `frm_Log` code-behind — jego ListBox używa `ReportID`, `KlientFK`, `NazwaKlienta`, `Status`, `EmailRecipient`, `BatchSentTimestamp` — żadne z tych nie zniknęło ani nie zostało rename'owane. **Brak zmian w `frm_Log.code-behind.txt`**.
- `ws_UserCache` schema — bez zmian.
- `ws_UsersRegistry` schema — bez zmian.
- `mod_UserCacheSync` — bez zmian.
- `mod_Utils`, `mod_Export` — bez zmian.
- ADR-y — bez nowego ADR (rename kolumny to zmiana kosmetyczna, usunięcie Fields to YAGNI cleanup — nie decyzja architektoniczna zmieniająca kierunek projektu).

---

## 🆕 Registry xlsx sync (2026-07-26) — symetria 3 warstw cache

**Kontekst**: dodano `BNC_UsersRegistry.xlsx` jako trzeci write-through cache, analogicznie do UserCache/DataCache. Motywacje: symetria (spójność ADR-001/002/008), disaster recovery, admin READ visibility. Kierunek nadal jednostronny (`ws → xlsx`) — admin push (bi-directional) **deferred** do Fazy B lub M7.

### Wymagane re-imports

- `mod_UserCacheSync.bas` — API 13 → 14 (dodane `EnsureRegistryCacheFileExists` + wewnętrzny `SyncRegistryToFile`)
- `mod_Diagnostic.bas` — `ExpectedPublicProcs("mod_UserCacheSync")` update
- `mod_Tests.bas` — nowa asercja w `Test_MultiUser` (xlsx exists check)
- `ThisWorkbook` — dodane `EnsureRegistryCacheFileExists` w `Workbook_Open`

### ☐ R1. Re-import 4 modułów + NEW `mod_UsersRegistrySync`

**Uwaga: sekcja R zaktualizowana po refactorze extract `mod_UsersRegistrySync` (2026-07-26 iteracja 2).** Registry code wyekstrahowany z `mod_UserCacheSync` do osobnego modułu — symetria "sheet ↔ module" per ADR-001. **Musisz zaimportować nowy moduł + ponownie zaimportować odchudzony `mod_UserCacheSync`**.

VBE → Remove → No → Import File dla:
- `Source/Modules/mod_UserCacheSync.bas` — **ponowny re-import** (odchudzony do 8 procedur, bez Registry API)
- `Source/Modules/mod_UsersRegistrySync.bas` — **NOWY moduł** (7 procedur Registry API)
- `Source/Modules/mod_Diagnostic.bas` — ExpectedModules 8→9, ExpectedPublicProcs restructure
- `Source/Modules/mod_Tests.bas` — call sites zaktualizowane na `mod_UsersRegistrySync`

`ThisWorkbook` — dwuklik w Project Explorer → wklej code z `Source/ThisWorkbook/ThisWorkbook.code.txt` (nadpisz current — call sites zaktualizowane).

**Formularze — code-behind również dotknięty**:
- `frm_UserPicker` — call sites `SwitchUser`, `PrepareForNewUser`, `GetAllUsers` teraz przez `mod_UsersRegistrySync`
- `frm_Setup` — call site `AddNewUser` teraz przez `mod_UsersRegistrySync`

W VBE → dwuklik formularz → View Code → Ctrl+A → Delete → wklej nowy code-behind z odpowiedniego pliku.

### ☐ R2. Ctrl+S + Debug → Compile VBAProject

Verify że kompiluje bez błędów. Jeśli błąd — nazwa procedury nie zgadza się (np. literówka w ExpectedPublicProcs) — poprawić.

### ☐ R3. Zamknij i otwórz xlsm

Verify `Workbook_Open` log:
```
[timestamp] INFO: Workbook_Open: N uzytkownikow w Registry, pokazuje frm_UserPicker
```

I sprawdź czy powstał plik:
```
? mod_Utils.FileExists("C:\BNC_CacheFolder\BNC_UsersRegistry.xlsx")
```
→ **`True`** (nowy plik utworzony automatycznie przy pierwszym `Workbook_Open` po deploy)

### ☐ R4. AuditFullProject — expected 9/9 modułów, nowe API counts

```
mod_Diagnostic.AuditFullProject
```

Oczekiwane:
```
[OK]  mod_UserCacheSync      ~230 lin.  API: 8/8    ← było 14, teraz 8 (schudł po ekstrakcji)
[OK]  mod_UsersRegistrySync  ~330 lin.  API: 7/7    ← NOWY modul
--> 9/9 modulow obecnych                             ← było 8, teraz 9
```

Jeśli `[MISSING] mod_UsersRegistrySync` — nie zaimportowałeś nowego modułu (patrz R1).
Jeśli `mod_UserCacheSync API: 14/14 (brak: ClearUserCache)` — masz starą wersję modułu (poprzedni etap Registry sync), nie ponowny re-import po refactor. Wróć do R1 i re-importuj mod_UserCacheSync.

### ☐ R5. `Test_MultiUser` — nowa asercja

```
mod_Tests.Test_MultiUser
```

Oczekiwane (nowa linia): `[PASS] BNC_UsersRegistry.xlsx istnieje po AddNewUser | expected=True actual=True`

### ☐ R6. Manual smoke test admin READ scenario

1. Explorer → `C:\BNC_CacheFolder\`
2. Powinny być 3 pliki: `BNC_UserCache.xlsx`, `BNC_DataCache.xlsx`, **`BNC_UsersRegistry.xlsx`** (nowy)
3. Otwórz `BNC_UsersRegistry.xlsx` w Excelu (nie w VBA, po prostu double-click)
4. Verify: widzisz **wszystkich** userów z Registry — nagłówki 13 kolumn + wiersze per user
5. Zamknij xlsx bez zapisu
6. Wróć do BNC_Sender xlsm → `frm_UserPicker.btn_AddNew` → dodaj testowego usera przez `frm_Setup`
7. Wróć do Explorera → otwórz ponownie `BNC_UsersRegistry.xlsx` — nowy user powinien być widoczny (sync po `AppendUserToRegistry`)

### ☐ R7. Manual test — `SwitchUser` triggers sync (LastLogin update)

1. `frm_UserPicker` otwórz → wybierz **innego** usera (nie current)
2. → `frm_Main` się otwiera
3. Otwórz `BNC_UsersRegistry.xlsx` w Excelu
4. Verify: kolumna 13 (`LastLogin`) dla wybranego usera pokazuje **nowy** timestamp (kilka sekund temu)

### 🚫 Co świadomie NIE działa (deferred)

- **Admin push** — edycja `BNC_UsersRegistry.xlsx` w Excelu nie propaguje się do `ws_UsersRegistry` w xlsm. Read-back deferred do Fazy B (patrz ADR-008 "Persistence" sekcja).
- **Read-only lock** — plik nie ma atrybutu `+R`. Patrz [`HOWTO_readonly_cache_produkcja.md`](HOWTO_readonly_cache_produkcja.md) — implementacja przed M7.
- **Cleanup xlsx przy delete usera** — brak `DeleteUser` API. Manual cleanup przez unhide `ws_UsersRegistry` + delete row.

