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

## [0.1.0] — TBD

- Faza A: implementacja MVP z hybrid cache i workflow kierownika
