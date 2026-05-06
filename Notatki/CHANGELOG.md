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

## [0.1.0] — TBD

- Faza A: implementacja MVP z hybrid cache i workflow kierownika
