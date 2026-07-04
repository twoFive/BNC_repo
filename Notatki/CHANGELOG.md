# CHANGELOG — BNC_Sender

Format: [Keep a Changelog](https://keepachangelog.com/), wersjonowanie [SemVer](https://semver.org/).

## [Unreleased]

### Added
- Inicjalizacja repozytorium i szkielet struktury katalogów (M0)
- `.gitignore`, `README.md`, `Notatki/` (CHANGELOG, TODO, DECISIONS)
- Ekstrakcja dokumentacji PDF z `doc_v2/` do `doc_v2/extracted/*.md`
- **M1.1**: `mod_Utils.bas` — logging, daty, FSO helpery, JoinPath, walidacja typów
- **M1.1**: `mod_Tests.bas` — smoke testy (`RunAllTests` jako entry point)
- **M1.2**: `mod_UserCacheSync.bas` — Repository Pattern dla `ws_UserCache`, write-through cache do `BNC_UserCache.xlsx`
- **M1.3**: `mod_DataCacheSync.bas` — Repository Pattern dla `ws_DataCache`, autoincrement `ReportID`, snapshot `CNA`/`NrOddzialu` przy zapisie
- **M1**: ADR-001 (Repository Pattern), ADR-002 (sync bez clipboard)
- **M2.1**: `mod_Validation.bas` — atomowe walidacje (email, FK, długość, miesiąc, folder) + `ValidateSetupData` / `ValidateReportData`
- **M2.2**: `frm_Setup.LAYOUT.md` (spec layoutu) + `frm_Setup.code-behind.txt` (kod do wklejenia w VBE) — UserForm z rejestracją handlowca, samouczkiem, browse folder dialog
- **M2.3**: `ThisWorkbook.code.txt` — `Workbook_Open` jako entry point z auto-recreate cache i decision setup-vs-main

### Changed
- **M2.2 (policy)**: `txt_EmailBNC` i `txt_CacheFolderPath` w `frm_Setup` są **hardcoded i locked** (ADR-003). Zamiast `btn_Browse` (folder picker) — `btn_CreateCacheFolder` tworzący folder na hardkodowanej ścieżce. Niezawodność > elastyczność.
- **M2.2 (UX)**: samouczek wyniesiony z inline `txt_Tutorial` do osobnego formularza `frm_Tutorial` (placeholder `btn_ShowTutorial` w M2, implementacja `frm_Tutorial` odłożona do M6/M7).
- **Repo**: `Notatki/NOTES.md` w `.gitignore` — lokalny notatnik decyzji autora, nie commitowany.
- **M3**: `frm_Main.LAYOUT.md` + `frm_Main.code-behind.txt` — główny ekran usera (header z info o roli, sekcja nowego zgłoszenia, ListBox pending z 5 kolumnami, przyciski Send/Log/Add/Clear). RefreshPendingList przez 2D array assign (10× szybsze niż `AddItem` pętli).
- **M4**: `mod_MailSender.bas` — `SendBatch()` jako pełny pipeline (GetPending → GenerateTempFile → DetermineRecipient → SendMailWithAttachment → MarkAsSent → CleanupTempFile). `DetermineRecipient` public dla testowalności.
- **M5.1**: `mod_Export.bas` — `ExportDataCache(targetPath)` (literal copy `BNC_DataCache.xlsx`) + `GetSuggestedExportFileName` (`BNC_Eksport_<Nazwisko>_<yyyy-mm-dd>.xlsx`).
- **M5.2**: `frm_Log.LAYOUT.md` + `frm_Log.code-behind.txt` — historia z statystykami pending/sent + Save As dialog dla eksportu.
- **M2.3 + M2.2 (final)**: `ThisWorkbook.code.txt` i `frm_Setup.code-behind.txt` zaktualizowane na direct `frm_Main.Show` (placeholder MsgBox usunięty).
- **M4 + M5**: ADR-004 (plik tymczasowy w %TEMP% jako transient artifact), ADR-005 (centralizacja routingu w mod_MailSender).
- **M3 (extension)**: `mod_DataCacheSync.DeleteRecord(reportID)` — hard delete dla pending, defensywna odmowa dla sent (ADR-006).
- **M3 (UX)**: `btn_DeleteSelected` w `frm_Main` pod ListBox — single-select, confirmation MsgBox z ID + nazwą klienta, `Enabled = False` gdy lista pusta. Newest-first ordering w `RefreshPendingList` (reverse iteration). Brak auto-selekcji po refresh — user musi sam kliknąć row.
- **M3 (tests)**: `Test_mod_DataCacheSync` rozszerzony o asercje delete: odmowa dla sent, zgoda dla pending (round-trip), odmowa dla nieznanego ID.

### Fixed
- **frm_Log (stale data)**: `LoadRecords` przeniesione z `UserForm_Initialize` do `UserForm_Activate`. Bug pre-existing (sprzed M3.2), uwypuklony przez delete feature — drugi `Show` frm_Log po zmianie `ws_DataCache` nie przeładowywał ListBoxa. Patrz ADR-007.
- **frm_Main (symetria)**: `RefreshPendingList` przeniesione z `Initialize` do `Activate`. Aktualnie bug niewidoczny (wszystkie modyfikacje pending dzieją się w frm_Main, więc lokalne handlery odświeżają), ale forward-compatible dla przyszłych ścieżek modyfikacji ws_DataCache spoza frm_Main.

### Documentation
- **`doc_v2/diagrams/03_data_flow_extended.{html,jpg,md}`** — rozszerzenie oryginalnego `03_data_flow.pdf` o **Flow C (hard delete pending z M3.2)**. Wszystkie 3 flows (A: dodanie, B: wysyłka, C: delete) w jednym diagramie + companion MD z opisem każdego kroku, edge cases, cross-refs do ADR-001/002/004/005/006/007.
- **`doc_v2/diagrams/04_data_model.{html,jpg,md}`** — wizualizacja modelu danych (poprzednio tylko PDF + raw extracted MD). Dwie tabele side-by-side (ws_UserCache + ws_DataCache) z field types i ról, cykl życia rekordu (pending → sent ALBO → hard delete), 4 wzorce (hybrid cache, audit trail, convention over config, snapshot), ścieżka migracji do Fazy B (nowy stan `awaiting_acceptance` przez Power Automate).
- **doc_3/diagrams/** — sync nowych plików (03_data_flow_extended + 04_data_model) do konsolidowanej dokumentacji.

### Added — M3.3 Multi-user support
- **`mod_UserCacheSync.bas`** rozszerzony o Registry API: `GetUsersCount`, `GetAllUsers`, `CurrentUserID`, `SwitchUser`, `AddNewUser`, `PrepareForNewUser`. Nowy arkusz **`ws_UsersRegistry`** (very hidden, tabelaryczny, 13 kolumn: `UserID` PK + 11 pól user + `LastLogin`). `EnsureRegistrySheet()` auto-tworzy arkusz przy pierwszym dostępie. Existing API (`GetUserField`, `IsSetupCompleted`, `IsUserManager`) **bez zmian** — nadal czyta z UserCache (semantycznie: aktywny user).
- **`frm_UserPicker.LAYOUT.md`** + **`.code-behind.txt`** — nowy formularz. ListBox 3-kolumnowy (Imię, Nazwisko, Email), przyciski „Wybierz i uruchom" / „Dodaj nowego użytkownika" / „Anuluj" (zamyka xlsm z MsgBox). Domyślna selekcja = user z najnowszym `LastLogin`.
- **`ThisWorkbook.code.txt`** — routing multi-user: `GetUsersCount() = 0` → `frm_Setup` (pierwszy raz), `≥ 1` → `frm_UserPicker`.
- **`frm_Setup.code-behind.txt`** — `btn_Save_Click` używa `AddNewUser(userData)` zamiast `SaveUserData` — każde uruchomienie Setup = nowy user w Registry.
- **`mod_Diagnostic.bas`** — jednorazowy tool developerski (`DumpVBComponents`, `ListPublicProcedures`, `DumpModuleContent`, `CountLinesTotal`) do inspekcji stanu VBProject w VBE. Wymaga `Trust access to the VBA project object model`.
- **ADR-008** (Multi-user Registry pattern) — uzasadnienie architektury z porównaniem alternatyw + `UserID` format `UZYTKOWNIK_<N>_CNA<cna>`.

## [0.1.0] — TBD

- Faza A: implementacja MVP z hybrid cache i workflow kierownika
