# Diagramy architektury — doc_4

Diagramy do prezentacji/dyskusji z biznesem. **Stan post-ADR-009** (2026-08-08) — Faza A code-complete + UAT-ready.

**Kluczowe zmiany vs doc_3**:
- ✅ **Schema v2** — `ws_DataCache` 11→10 kolumn (rename `MiesiacZgloszenia` → `MiesiacObrotu`, usunięto `Fields`)
- ✅ **Registry xlsx sync** — `BNC_UsersRegistry.xlsx` (write-through, jednostronny)
- ✅ **`mod_UsersRegistrySync`** wyekstrahowany z `mod_UserCacheSync` (symetria sheet↔module)
- ✅ **ADR-009** — usunięcie `ws_UserCache`, `ws_AppState` jako session marker, Registry as sole source of truth
- ✅ **9 modułów**: `mod_Utils`, `mod_Validation`, `mod_AppStateSync` (NEW), `mod_UsersRegistrySync`, `mod_DataCacheSync`, `mod_MailSender`, `mod_Export`, `mod_Tests`, `mod_Diagnostic`

**Format wyjścia**: URL + PNG (bez JPG konwersji, bez PDF).

## 📋 Katalog diagramów

| Plik | Cel | Audience |
|---|---|---|
| `01_system_context.png` | **C4 Level 1 — System Context**: aktorzy (Handlowiec, Kierownik, Zespół BNC, Dział IT) + systemy zewnętrzne (Outlook, Exchange, OneDrive). Statyczna mapa "kto styka się z systemem". | biznes + tech |
| `02_system_architecture.png` | **C4 Level 2 — Container**: 4 warstwy (dystrybucja, komputer handlowca, sieciowa, adresaci). Post-ADR-009: **2 xlsx cache** (Registry + DataCache), 9 modułów. | tech / IT |
| `03_data_flow_extended.png` + `.md` | Flow A (insert pending, schema v2) + Flow B (send batch — 7 kolumn biznesowych attachment) + Flow C (hard delete pending, ADR-006) | tech |
| `04_data_model.png` + `.md` | **ERD** — 3 arkusze: `ws_AppState` (session marker) + `ws_UsersRegistry` (13 kol, sole source of truth) + `ws_DataCache` (10 kol schema v2). Relacja SNAPSHOT z Registry (current user) → DataCache. | tech |
| `05_module_architecture.png` | Component Diagram (C4 model · poziom 3) — 4 warstwy modułów: Presentation (Forms) → Business (Validation, MailSender, Export) → Data Access (**3 moduły**: AppStateSync, UsersRegistrySync, DataCacheSync) → Infrastructure (Utils, sheets, Outlook COM) | tech / dev |
| `06_business_process.png` | **Workflow biznesowy v1** — single-page composite: panele Faza A intro + workflow + panele Faza B na jednym obrazie. Dla drukowania A2/A3. | biznes / decyzyjna |
| `06_business_process_v2.png` + `.md` | **Workflow biznesowy v2** — split: PNG czysty workflow, tekstowy opis Faza A + wartości + Faza B w markdown. Dla slajdów + executive summary. | biznes / decyzyjna |
| `*.html` | Źródła diagramów (SVG embedded), regenerowalne | — |

## 🔄 Regeneracja PNG ze źródła HTML

Zmodyfikuj `*.html`, potem uruchom Chrome headless w Git Bash:

```bash
cd doc_4/diagrams

# Wymiary per diagram (dopasowane do zawartości SVG w HTML):
declare -A DIMS=(
  ["01_system_context"]="1400 1020"
  ["02_system_architecture"]="1600 1220"
  ["03_data_flow_extended"]="1400 1080"
  ["04_data_model"]="1400 900"
  ["05_module_architecture"]="1600 1100"
  ["06_business_process"]="1400 1560"
  ["06_business_process_v2"]="1400 1080"
)

for f in "${!DIMS[@]}"; do
  read w h <<< "${DIMS[$f]}"
  "/c/Program Files/Google/Chrome/Application/chrome.exe" \
    --headless --disable-gpu --hide-scrollbars \
    --window-size=$w,$h \
    --screenshot="$(pwd -W)/$f.png" \
    "file:///$(pwd -W)/$f.html"
  echo "✓ $f.png ($w × $h)"
done
```

**Uwaga**: brak JPG konwersji (poprzednie iteracje `doc_v2` / `doc_3` używały PowerShell + System.Drawing do PNG→JPG). W `doc_4` PNG jest **jedynym** output format — mniejsza objętość, brak straty jakości, browsers otwierają natywnie.

## 📐 Wymiary diagramów

- **01 system_context**: 1400 × 1020 px
- **02 system_architecture**: 1600 × 1220 px
- **03 data_flow_extended**: 1400 × 1080 px
- **04 data_model**: 1400 × 900 px
- **05 module_architecture**: 1600 × 1100 px
- **06 business_process (v1)**: 1400 × 1560 px
- **06 business_process_v2**: 1400 × 1080 px

Drukuj na A3 (min. 1600px szeroko) albo prezentuj na monitorze 1080p+.

## 🎯 Post-ADR-009 checklist

Diagramy zaktualizowane by odzwierciedlać:
- [x] **04 data_model** — `ws_AppState` + `ws_UsersRegistry` (13 kol) + `ws_DataCache` (10 kol schema v2), snapshot z Registry current user
- [x] **05 module_architecture** — 3 moduły w Warstwa dostępu do danych (AppState, UsersRegistry, DataCacheSync), 3 arkusze w Warstwa infrastruktury (AppState, UsersRegistry, DataCache), cross-module dep (UsersRegistrySync → AppStateSync)
- [x] **02 system_architecture** — 2 xlsx cache (BNC_UsersRegistry + BNC_DataCache), 9 modułów w komputerze handlowca, 3 ukryte arkusze
- [x] **03 data_flow_extended** — snapshot references (UserCache → Registry current user)
- [x] **01 system_context** — bez zmian (aktorzy/systemy zewnętrzne nie dotknięte)
- [x] **06 business_process** — bez zmian (biznes workflow nie dotknięty ADR-009)

## 🔗 Powiązane

- **ADR-009** — patrz [`Notatki/DECISIONS.md`](../../Notatki/DECISIONS.md)
- **Migration procedure** — patrz [`Notatki/TODO_next_session.md`](../../Notatki/TODO_next_session.md) sekcja T1-T10
- **Historyczne wersje diagramów** — `doc_v2/diagrams/`, `doc_3/diagrams/`
