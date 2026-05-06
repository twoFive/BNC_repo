# TODO — BNC_Sender Faza A

Bieżąca lista zadań. Szczegółowy plan: [`../BNC_Sender_PlanWdrozenia_FazaA.md`](../BNC_Sender_PlanWdrozenia_FazaA.md).

## M0 — Setup środowiska

- [x] Utworzyć szkielet folderów (`Source/`, `Notatki/`, `doc_v2/extracted/`)
- [x] `.gitignore`, `README.md`, `CHANGELOG.md`, `DECISIONS.md`
- [x] Ekstrakcja dokumentacji PDF
- [x] `git init` w `FazaA\`
- [x] Initial commit
- [ ] Dodać GitHub remote, pierwszy push
- [ ] **(ręcznie w Excelu)** Wykonać kroki I i II z `BNC_srodowiskoDEV_FazaA.pdf`
- [ ] **(ręcznie)** Skopiować `BNC_Sender_v0.1.0.xlsm` z `Releases/` do `Working/`
- [ ] **(ręcznie)** Smoke testy: read userCache / write dataCache / send mail

## M1 — Foundation (`mod_Utils`, `mod_UserCacheSync`, `mod_DataCacheSync`)

- [x] M1.1: `mod_Utils.bas` + `mod_Tests.bas` (Test_mod_Utils)
- [x] M1.2: `mod_UserCacheSync.bas` + Test_mod_UserCacheSync
- [x] M1.3: `mod_DataCacheSync.bas` + Test_mod_DataCacheSync
- [x] ADR-001 (Repository Pattern), ADR-002 (sync bez clipboard)
- [ ] **(ręcznie w VBE)** File > Import File... dla wszystkich 4 plików `.bas`
- [ ] **(ręcznie)** Uruchomić `mod_Tests.RunAllTests` z Immediate Window — wszystkie PASS
- [ ] **(ręcznie)** Sprawdzić, że `BNC_DataCache.xlsx` i `BNC_UserCache.xlsx` powstają w `CacheFolderPath`

## M2..M7

Patrz plan implementacji.
