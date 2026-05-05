# BNC_Sender — Faza A

Aplikacja Excel/VBA do zgłaszania nowych klientów do programu BNC (Bonus New Client).
Działa lokalnie na komputerze handlowca, wysyła batche wniosków przez Outlook COM.

> **Status**: w budowie — wersja docelowa `BNC_Sender_v0.1.0.xlsm`
> **Plan implementacji**: [`BNC_Sender_PlanWdrozenia_FazaA.md`](BNC_Sender_PlanWdrozenia_FazaA.md)

---

## Architektura w skrócie

- **Primary storage**: ukryte arkusze `ws_UserCache`, `ws_DataCache` w pliku `.xlsm`
- **Synchronized backup** (write-through cache): pliki `BNC_UserCache.xlsx`, `BNC_DataCache.xlsx` w lokalnym folderze cache
- **Wysyłka**: Outlook COM (MailItem.Send) z plikiem `.xlsx` ad-hoc generowanym z pending zgłoszeń
- **Routing**: jeśli `EmailKierownika == EmailHandlowca` → user jest kierownikiem, mail leci wprost do BNC; inaczej do kierownika z prośbą o weryfikację

Pełna architektura: [`doc_v2/extracted/02_system_architecture.md`](doc_v2/extracted/02_system_architecture.md)

---

## Struktura repozytorium

```
FazaA/
├── Source/                    # eksporty kodu z VBA Editor (Git tracked)
│   ├── Modules/               # .bas
│   ├── Forms/                 # .frm + .frx
│   └── ThisWorkbook/          # .cls
├── Notatki/                   # dokumentacja deweloperska
│   ├── CHANGELOG.md
│   ├── TODO.md
│   └── DECISIONS.md           # ADR
├── doc_v2/                    # dokumentacja projektowa (PDF + extracted MD)
├── Releases/                  # [.gitignore] archiwum wersji .xlsm
├── Working/                   # [.gitignore] aktualnie edytowana wersja
├── CacheFolder/               # [.gitignore] testowy cache (DEV only)
├── BNC_Sender_PlanWdrozenia_FazaA.md
└── README.md
```

## Workflow eksportu kodu z VBA do Source/

Po każdej istotnej zmianie w VBA Editor:

1. W VBA Editor: prawy klik na moduł/formularz → **Export File...**
2. Zapis do `Source/Modules/` lub `Source/Forms/`
3. Terminal: `git add Source/ && git commit -m "feat(Mx): opis"`

## Roadmap (milestones)

| # | Milestone | Czas |
|---|---|---|
| M0 | Setup środowiska + repo | ~1d |
| M1 | Foundation: `mod_Utils`, `mod_UserCacheSync`, `mod_DataCacheSync` | ~3d |
| M2 | `frm_Setup` + `mod_Validation` | ~2d |
| M3 | `frm_Main` (batch + lista pending) | ~3d |
| M4 | `mod_MailSender` (Outlook COM, decision diamond) | ~3d |
| M5 | `frm_Log` + `mod_Export` | ~2d |
| M6 | Polish + UAT | ~3d |
| M7 | Release v1.0.0 | — |

Szczegóły: [`BNC_Sender_PlanWdrozenia_FazaA.md`](BNC_Sender_PlanWdrozenia_FazaA.md)
