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

## M2 — Setup form

- [x] M2.1: `mod_Validation.bas` + `Test_mod_Validation`
- [ ] M2.2: `frm_Setup` — layout spec + code-behind template (form tworzony ręcznie w VBE)
- [ ] M2.3: `ThisWorkbook.cls` — `Workbook_Open` jako entry point
- [ ] **(ręcznie w VBE)** Import `mod_Validation.bas`, uruchomić `Test_mod_Validation`
- [ ] **(ręcznie w VBE)** Utworzyć UserForm `frm_Setup` zgodnie z layout spec, wkleić code-behind
- [ ] **(ręcznie w VBE)** Wkleić kod `Workbook_Open` w istniejącą `ThisWorkbook`
- [ ] **(ręcznie)** Eksport `frm_Setup.frm` + `.frx` do `Source/Forms/`
- [ ] **(ręcznie)** Eksport `ThisWorkbook.cls` do `Source/ThisWorkbook/`

## M3 — Main form (kod gotowy)

- [x] M3.1: `frm_Main.LAYOUT.md` + `frm_Main.code-behind.txt`
- [ ] **(ręcznie w VBE)** Utworzyć UserForm `frm_Main` z kontrolkami zgodnie z LAYOUT
- [ ] **(ręcznie w VBE)** Wkleić code-behind
- [ ] **(ręcznie)** Smoke: AddToList × 3, Send (z Outlookiem!), zobacz w `ws_DataCache` że Status=sent

## M4 — Mail sender (kod gotowy)

- [x] M4.1: `mod_MailSender.bas` + `Test_mod_MailSender`
- [ ] **(ręcznie)** Import `mod_MailSender.bas`, re-import `mod_Tests.bas`
- [ ] **(ręcznie)** `mod_Tests.Test_mod_MailSender` — testuje DetermineRecipient bez wysyłki
- [ ] **(ręcznie UAT)** Test handlowca: EmailKierownika≠EmailHandlowca, Send → kierownik dostaje mail
- [ ] **(ręcznie UAT)** Test kierownika: EmailKierownika=EmailHandlowca, Send → BNC dostaje mail
- [ ] **(ręcznie UAT)** Test błędu: wyłącz Outlook, Send → komunikat błędu, plik tymczasowy nie zostaje

## M5 — Log + Export (kod gotowy)

- [x] M5.1: `mod_Export.bas` + `Test_mod_Export`
- [x] M5.2: `frm_Log.LAYOUT.md` + `frm_Log.code-behind.txt`
- [ ] **(ręcznie)** Import `mod_Export.bas`, re-import `mod_Tests.bas`
- [ ] **(ręcznie w VBE)** Utworzyć UserForm `frm_Log` z kontrolkami
- [ ] **(ręcznie w VBE)** Wkleić code-behind
- [ ] **(ręcznie)** Smoke: otwórz frm_Log, sprawdź statystyki, eksportuj, sprawdź plik

## Kolejność import/paste w następnej sesji

> **Cel**: minimalna liczba "compile error" przy pracy w VBE.

1. **Moduły** (kolejność dowolna):
   - `mod_MailSender.bas` — Import File
   - `mod_Export.bas` — Import File
   - `mod_Tests.bas` — **Remove** stary, **Import** nowy (z 6 testami zamiast 4)
2. **UserForms** (utwórz **shells** najpierw — Insert→UserForm × 2, nazwij `frm_Main` i `frm_Log`):
   - Powód: w `frm_Setup.btn_Save` jest `frm_Main.Show`, w `frm_Main.btn_ShowLog` jest `frm_Log.Show`. Nawet pusty shell wystarczy do kompilacji.
3. **Dodaj kontrolki** do `frm_Main` (per `frm_Main.LAYOUT.md`) → wklej code-behind
4. **Dodaj kontrolki** do `frm_Log` (per `frm_Log.LAYOUT.md`) → wklej code-behind
5. **Re-paste**: `frm_Setup.code-behind.txt` (z direct `frm_Main.Show` zamiast placeholdera)
6. **Re-paste**: `ThisWorkbook.code.txt` (z direct `frm_Main.Show`)
7. `Ctrl+S`, `mod_Tests.RunAllTests` w Immediate

## M6 — Polish + UAT

- [ ] Każda procedura `Public` ma `On Error GoTo ErrorHandler`
- [ ] Komunikaty błędów dla usera są zrozumiałe
- [ ] Tab order na formularzach
- [ ] Akceleratory klawiszowe (Alt+S, Alt+A)
- [ ] Implementacja `frm_Tutorial` (odłożona z M2.2)
- [ ] UAT z 3-5 testowymi userami

## M7 — Release v1.0.0

- [ ] Bump version `BNC_Sender_v1.0.0.xlsm`
- [ ] Eksport finalny do `Source/`
- [ ] Git tag `v1.0.0`
- [ ] Wgraj na OneDrive firmowy
