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

- [ ] Wygenerować szablony `.bas` do importu w VBE
- [ ] M1.1: `mod_Utils` — implementacja + test
- [ ] M1.2: `mod_UserCacheSync` — implementacja + test (sync do xlsx)
- [ ] M1.3: `mod_DataCacheSync` — implementacja + test (sync do xlsx)
- [ ] Eksport modułów do `Source/Modules/`, commit per moduł

## M2..M7

Patrz plan implementacji.
