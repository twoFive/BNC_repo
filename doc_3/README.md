# `doc_3/` — pełna dokumentacja projektu BNC_Sender (faza A)

**Cel**: jedna lokalizacja ze wszystkimi dokumentami projektu — dla łatwego udostępniania (zip, SharePoint, mail) i przeglądania bez nawigowania po całej strukturze repo.

> **Status**: `doc_3/` to **konsolidacja** — kopie plików z `doc_v2/`, `Notatki/`, root projektu i `Source/Forms/`. Źródła prawdy są tam — `doc_3/` to **snapshot** dla wygody konsumenta.

---

## 📁 Struktura `doc_3/`

```
doc_3/
├── README.md                              ← jesteś tutaj
├── 00_index.md                            ← indeks z konwencjami (kopia z doc_v2/)
├── PLAN.md                                ← plan implementacji (z BNC_Sender_PlanWdrozenia_FazaA.md)
├── CHANGELOG.md                           ← historia zmian (z Notatki/)
├── DECISIONS.md                           ← 7 ADR-ów (z Notatki/)
│
├── pdfs/                                  ← oryginalne PDF-y od architekta
│   ├── BNC_fazaA_02_system_architecture.pdf
│   ├── BNC_fazaA_03_data_flow.pdf
│   ├── BNC_fazaA_04_data_model.pdf
│   ├── BNC_fazaA_05_module_architecture.pdf
│   └── BNC_srodowiskoDEV_FazaA.pdf
│
├── extracted/                             ← polished MD (z PDF-ów)
│   ├── 02_system_architecture.md
│   ├── 03_data_flow.md
│   ├── 04_data_model.md
│   └── 05_module_architecture.md
│
├── diagrams/                              ← wizualizacje HTML+SVG → JPG
│   ├── 01_system_context.{html,jpg}       (C4 Level 1)
│   ├── 02_system_architecture.{html,jpg}  (C4 Level 2)
│   ├── 05_module_architecture.{html,jpg}  (C4 Level 3)
│   ├── 06_business_process.{html,jpg}     (BPMN v1 composite)
│   ├── 06_business_process_v2.{html,jpg,md} (BPMN v2 split)
│   ├── README.md
│   └── post-release/                      (wariant docelowy v1.0.0)
│       ├── 02_system_architecture.{html,jpg}
│       ├── 05_module_architecture.{html,jpg}
│       └── README.md
│
└── ui-specs/                              ← UI specyfikacje formularzy
    ├── frm_Setup.LAYOUT.md
    ├── frm_Main.LAYOUT.md
    └── frm_Log.LAYOUT.md
```

---

## 🎯 Audience routing — co czytać przy jakim pytaniu

### Dla decydentów / przełożonych (biznes)

| Pytanie | Dokument |
|---|---|
| Co to za projekt? Jaka wartość? | [`diagrams/06_business_process_v2.md`](diagrams/06_business_process_v2.md) §1 |
| Jak działa proces (workflow)? | [`diagrams/06_business_process_v2.jpg`](diagrams/06_business_process_v2.jpg) |
| Wersja "all-in-one" (wartość + workflow + Faza B) | [`diagrams/06_business_process.jpg`](diagrams/06_business_process.jpg) |
| Plany rozwoju (Faza B — SQL Server + Power Automate) | [`diagrams/06_business_process_v2.md`](diagrams/06_business_process_v2.md) §3 |
| Kto styka się z systemem? Stakeholdery? | [`diagrams/01_system_context.jpg`](diagrams/01_system_context.jpg) (C4 L1) |
| Co IT musi przygotować? | [`extracted/02_system_architecture.md`](extracted/02_system_architecture.md) §"Wymagania dla działu IT" |

### Dla architektów / tech leadów

| Pytanie | Dokument |
|---|---|
| Jak działa technicznie (containers)? | [`diagrams/02_system_architecture.jpg`](diagrams/02_system_architecture.jpg) (C4 L2) |
| Jaka struktura kodu (modules)? | [`diagrams/05_module_architecture.jpg`](diagrams/05_module_architecture.jpg) (C4 L3) |
| Jak będzie wyglądać po release? | [`diagrams/post-release/`](diagrams/post-release/) (v1.0.0 target) |
| Dlaczego konkretnie tak? | [`DECISIONS.md`](DECISIONS.md) (7 ADR-ów) |
| Jak są zorganizowane dane? | [`pdfs/BNC_fazaA_04_data_model.pdf`](pdfs/BNC_fazaA_04_data_model.pdf) + [`extracted/04_data_model.md`](extracted/04_data_model.md) |
| Jak płynie dataflow (Add + Send)? | [`pdfs/BNC_fazaA_03_data_flow.pdf`](pdfs/BNC_fazaA_03_data_flow.pdf) + [`extracted/03_data_flow.md`](extracted/03_data_flow.md) |

### Dla developerów / implementerów

| Pytanie | Dokument |
|---|---|
| Co i w jakiej kolejności robić? | [`PLAN.md`](PLAN.md) |
| Jakie kontrolki w formularzach (UX)? | [`ui-specs/`](ui-specs/) |
| Historia zmian | [`CHANGELOG.md`](CHANGELOG.md) |
| Konwencje dokumentacji (C4, ADR, BPMN, ERD) | [`00_index.md`](00_index.md) |

### Dla operations / IT

| Pytanie | Dokument |
|---|---|
| Jak skonfigurować środowisko DEV? | [`pdfs/BNC_srodowiskoDEV_FazaA.pdf`](pdfs/BNC_srodowiskoDEV_FazaA.pdf) |
| Wymagania od IT (Trust Locations, Outlook polityki) | [`extracted/02_system_architecture.md`](extracted/02_system_architecture.md) §"Wymagania dla działu IT" — gotowy ticket do skopiowania |

---

## 🗺 Standardy użyte w portfolio

| Standard | Liczba plików | Lokalizacje |
|---|---|---|
| **C4 Model** (Simon Brown) — L1/L2/L3 | 10 | `pdfs/`, `extracted/`, `diagrams/`, `diagrams/post-release/` |
| **BPMN-light** (swim lanes / flow) | 5 | `pdfs/03_*`, `extracted/03_*`, `diagrams/06_*` |
| **ERD-light** | 2 | `pdfs/04_*`, `extracted/04_*` |
| **ADR** (Michael Nygard format) | 1 plik z 7 wpisami | `DECISIONS.md` |
| **Keep a Changelog** | 1 | `CHANGELOG.md` |
| **UI Control Spec** (custom) | 3 | `ui-specs/` |

Pełna analiza: [`00_index.md`](00_index.md) §1.

---

## 🔄 Co zmieniono podczas konsolidacji

| Plik | Zmiana | Powód |
|---|---|---|
| `extracted/02_system_architecture.md` | `../../Notatki/DECISIONS.md` → `../DECISIONS.md` | `DECISIONS.md` jest teraz w `doc_3/` root |
| `extracted/05_module_architecture.md` | j.w. | j.w. |
| `extracted/0[25]_*.md` | `Źródło: doc_v2/BNC_fazaA_*.pdf` → `Źródło: ../pdfs/BNC_fazaA_*.pdf` | PDF-y są teraz w `doc_3/pdfs/` |
| `diagrams/README.md` | `doc_v2/diagrams` → `doc_3/diagrams` w pipeline regeneracji | Self-consistency |
| `diagrams/post-release/README.md` | `doc_v2/diagrams` → `doc_3/diagrams` | j.w. |

---

## 🌐 Linki które **zostają na zewnątrz** (świadomie)

Niektóre linki w skopiowanych dokumentach wskazują poza `doc_3/` (do `Source/`, `Notatki/`, root projektu). To **świadomy wybór**:

- `extracted/*.md` linki do `../../Source/Modules/*.bas` — **kod żyje poza dokumentacją**, w `Source/Modules/`. Linki działają gdy `doc_3/` jest w repo. W zip-snapshocie linki będą martwe — to akceptowalne dla "doc-only" share'u.
- `PLAN.md`, `DECISIONS.md`, `CHANGELOG.md` opisują strukturę całego projektu (`Notatki/`, `Source/`) — to **opis** struktury repo, nie hyperlinki do nawigacji.
- `00_index.md` — kopia indeksu opisującego `doc_v2/` (zachowuje swoją wartość jako snapshot konwencji, paths-as-strings nie hyperlinki).

---

## 🔁 Synchronizacja z `doc_v2/`, `Notatki/`, etc.

`doc_3/` to **snapshot** kopiowany z innych lokalizacji. Po zmianach w **źródłach** (`doc_v2/`, `Notatki/`, `Source/Forms/*.LAYOUT.md`, `BNC_Sender_PlanWdrozenia_FazaA.md`) trzeba **ręcznie** zaktualizować `doc_3/`:

```bash
# Pełny refresh doc_3 ze źródeł (zachowuje tylko README.md doc_3)
cd C:/Dev/BNC_Sender/FazaA
cp doc_v2/00_index.md doc_3/
cp doc_v2/BNC_fazaA_*.pdf doc_v2/BNC_srodowiskoDEV_FazaA.pdf doc_3/pdfs/
cp doc_v2/extracted/*.md doc_3/extracted/
cp -r doc_v2/diagrams/* doc_3/diagrams/
cp Notatki/CHANGELOG.md Notatki/DECISIONS.md doc_3/
cp BNC_Sender_PlanWdrozenia_FazaA.md doc_3/PLAN.md
cp Source/Forms/frm_*.LAYOUT.md doc_3/ui-specs/

# Re-apply path fixes
sed -i 's|\.\./\.\./Notatki/DECISIONS\.md|../DECISIONS.md|g' doc_3/extracted/02_*.md doc_3/extracted/05_*.md
sed -i 's|doc_v2/BNC_fazaA_02_system_architecture\.pdf|../pdfs/BNC_fazaA_02_system_architecture.pdf|g' doc_3/extracted/02_*.md
sed -i 's|doc_v2/BNC_fazaA_05_module_architecture\.pdf|../pdfs/BNC_fazaA_05_module_architecture.pdf|g' doc_3/extracted/05_*.md
sed -i 's|doc_v2/diagrams|doc_3/diagrams|g' doc_3/diagrams/README.md doc_3/diagrams/post-release/README.md
```

> **Recommendation**: synchronizuj `doc_3/` **przed** każdym istotnym share'em (przegląd z przełożonym, audit, onboarding nowego członka zespołu). Między synchronizacjami `doc_3/` może być nieaktualne.

---

## 📊 Liczby

- **5** PDF-ów (oryginały) — ~1.9 MB łącznie
- **4** extracted MD (polished z PDF-ów)
- **8** HTML+JPG par diagramów (4 main + 2 post-release + 2 business v1/v2)
- **3** UI specs (LAYOUT.md)
- **5** plików meta/tekstowych w root (README, index, PLAN, CHANGELOG, DECISIONS)

Łącznie: ~25 plików dokumentacji w jednej lokalizacji.
