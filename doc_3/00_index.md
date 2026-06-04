# `doc_v2/` — indeks dokumentacji projektu BNC_Sender

> Entry point do dokumentacji projektu. Mapa wszystkich plików, konwencje, audience per dokument.
> Repo: [github.com/twoFive/BNC_repo](https://github.com/twoFive/BNC_repo)

---

## 1. Konwencje dokumentacji

### 1.1 Frameworki i standardy

| Standard | Zastosowanie | Pliki |
|---|---|---|
| **[C4 Model](https://c4model.com/)** (Simon Brown) | Architektura systemu w 4 poziomach abstrakcji | `01` (Context), `02` (Container), `05` (Component) |
| **BPMN-light** (Business Process Model) | Przepływy biznesowe i danych (swim lanes, decision diamonds) | `03` (Data flow), `06` (Business process) |
| **ERD-light** (Entity Relationship Diagram) | Model danych — pola, relacje, snapshoty | `04` (Data model) |
| **ADR** (Architectural Decision Records, Michael Nygard) | Decyzje architektoniczne z uzasadnieniem | `Notatki/DECISIONS.md` (poza `doc_v2/`) |

### 1.2 Konwencje nazewnicze

| Element | Format | Przykład |
|---|---|---|
| Numeracja sekwencyjna | `NN_topic` | `01_system_context`, `02_system_architecture` |
| Stały sufiks tematyczny | po `NN_` | `_system_context`, `_module_architecture` |
| PDF (oryginał od architekta) | `BNC_fazaA_NN_topic.pdf` | `BNC_fazaA_02_system_architecture.pdf` |
| Markdown (treść tekstowa) | `NN_topic.md` | `02_system_architecture.md` |
| Diagram (HTML źródło) | `NN_topic.html` | `02_system_architecture.html` |
| Diagram (JPG render) | `NN_topic.jpg` | `02_system_architecture.jpg` |
| Wariant docelowy | `post-release/NN_topic.{html,jpg}` | `post-release/02_system_architecture.jpg` |

### 1.3 Meta-pattern: **Docs-as-Code**

| Etap | Forma |
|---|---|
| **Źródło** | `*.html` (SVG embedded) + `*.md` (tekst) — version-controlled w Git |
| **Render** | Chrome headless (`--screenshot`) → PowerShell `System.Drawing` (JPG) |
| **Output** | `*.jpg` w repo (do prezentacji, druku) + `*.pdf` historyczne (oryginał od architekta) |
| **Regeneracja** | Pipeline udokumentowany w [`diagrams/README.md`](diagrams/README.md) |

Wszystko **tekstowe i regenerowalne** — nic w stanie "tylko obraz bez źródła".

### 1.4 Konwencje treściowe

- **Polish-first** — wszystkie tytuły, opisy, etykiety na diagramach po polsku z diakrytykami. Kod-komentarze w `.bas/.cls` są bez diakrytyków (bezpieczne kodowanie cp1250/ASCII).
- **Cross-linki repo-relative** — `[02_system_architecture.md](02_system_architecture.md)`, nie absolutne URL-e.
- **Sekcja `## Cel`** na początku każdego MD — 1-2 zdania zakresu.
- **Footer z wersją** — `BNC_Sender_v0.1.0 · Faza A · stan po MX`.
- **Audience marker** — każdy dokument oznaczony jako "biznes/decyzyjna" lub "techniczna".

### 1.5 Kolorystyka diagramów (przekrojowa)

| Kolor | Znaczenie |
|---|---|
| 🟦 Niebieski (`#1565c0`) | System w zakresie projektu, warstwa prezentacji, lub główny aktor |
| 🟧 Pomarańczowy (`#ef6c00`) | Logika biznesowa, decyzje, control gates |
| 🟪 Fioletowy (`#7b1fa2`) | Warstwa dostępu do danych (Repository Pattern), lub plany Faza B |
| ⬜ Szary (`#455a64`) | Infrastruktura, systemy zewnętrzne |
| 🟩 Zielony (`#388e3c`) | Aktor/persona, akcje pozytywne ("co dostarczamy") |
| 🟨 Żółty (`#ffd54f`) | Highlight, cel końcowy (np. "klient w programie BNC") |

---

## 2. Indeks dokumentów

### 2.1 Diagramy (`doc_v2/diagrams/`)

| Plik | Standard | Zawartość | Audience |
|---|---|---|---|
| **`01_system_context.{html,jpg}`** | C4 Level 1 | Aktorzy (Handlowiec, Kierownik, Zespół BNC, Dział IT) + systemy zewnętrzne (Outlook, Exchange, OneDrive) wokół BNC_Sender. Statyczna mapa zależności. | biznes + tech |
| **`02_system_architecture.{html,jpg}`** | C4 Level 2 (Container) | 4 warstwy systemu: dystrybucja, komputer handlowca, sieciowa, adresaci. Ticket do IT. Wzorce: hybrid cache, %TEMP%, transient artifact. | tech |
| **`05_module_architecture.{html,jpg}`** | C4 Level 3 (Component) | Podział aplikacji VBA na 7 modułów w 4 warstwach. Reguły zależności (strzałki tylko w dół). `mod_MailSender` jako serce decyzji kierownik vs handlowiec. | tech |
| **`06_business_process.{html,jpg}`** | BPMN swim lanes (v1) | Single-page composite: Faza A intro + workflow + Faza B roadmap (1400×1560). Do drukowania A2/A3. | biznes |
| **`06_business_process_v2.{html,jpg,md}`** | BPMN swim lanes (v2) | Split: JPG to czysty workflow (1400×1080), opis biznesowy w `.md` (Faza A wartość, Faza B plany). Do slajdów + executive summary. | biznes |
| **`post-release/02_*.{html,jpg}`** | C4 Level 2 (v1.0.0 target) | Wariant docelowy — po wdrożeniu wszystkich M1-M7. Footer "Stan docelowy po Release". | tech |
| **`post-release/05_*.{html,jpg}`** | C4 Level 3 (v1.0.0 target) | Wariant docelowy — `frm_Tutorial` jako pełny moduł, strzałka `frm_Setup → frm_Tutorial`. | tech |
| **`README.md`** | Meta | Tabela diagramów, pipeline regeneracji (Chrome headless + System.Drawing). | tech |

### 2.2 Treść tekstowa (`doc_v2/extracted/`)

Polished markdowny wyekstraktowane z PDF-ów + ręcznie dopisane sekcje.

| Plik | Zawartość |
|---|---|
| **`02_system_architecture.md`** | Architektura systemu jako tekst: 4 warstwy, hybrid cache, Outlook COM, ticket do IT, mechanizm %TEMP%, ścieżka migracji do Fazy B. Cross-linki do ADR-ów. |
| **`03_data_flow.md`** | Przepływy danych: Flow A (dodawanie zgłoszenia), Flow B (wysyłka batcha z decision diamond). Mechanizm pliku tymczasowego. |
| **`04_data_model.md`** | Model danych: ws_UserCache (key-value), ws_DataCache (tabela), kopia xlsx, snapshot CNA/NrOddzialu przy zapisie, EmailRecipient jako audit trail. |
| **`05_module_architecture.md`** | Architektura modułów: 4 warstwy aplikacji, mod_MailSender jako serce logiki, wzorzec sync, reguła warstw, 5+2 moduły, czego brakuje vs Faza B. |

### 2.3 Oryginalne PDF-y (`doc_v2/`)

Bezpośrednio od architekta projektu — **historyczna prawda** (źródło ekstraktów w `extracted/`).

| Plik | Zawartość |
|---|---|
| **`BNC_fazaA_02_system_architecture.pdf`** | Oryginalny diagram + opis architektury systemu |
| **`BNC_fazaA_03_data_flow.pdf`** | Oryginalny diagram + opis przepływów danych |
| **`BNC_fazaA_04_data_model.pdf`** | Oryginalny diagram + opis modelu danych |
| **`BNC_fazaA_05_module_architecture.pdf`** | Oryginalny diagram + opis architektury modułów |
| **`BNC_srodowiskoDEV_FazaA.pdf`** | Setup środowiska deweloperskiego (referencowane w `Plan implementacji`) |

> PDF-y są **immutable** — nie modyfikujemy. Aktualizacje treści idą do `extracted/*.md` (z notą "rozszerzone vs oryginał").

---

## 3. Dokumenty poza `doc_v2/`

Inne dokumenty projektu zlokalizowane w innych miejscach repo — wymienione dla pełności obrazu.

### 3.1 Plan implementacji (root projektu)

| Plik | Zawartość |
|---|---|
| **`BNC_Sender_PlanWdrozenia_FazaA.md`** | Plan wdrożenia krok-po-kroku: 7 milestones (M0-M7), pseudokod referencyjny dla każdego modułu/formularza, zadania per milestone, kryteria akceptacji. Aktualizowany w trakcie implementacji o feat'y poza pierwotnym planem (np. M3.2 hard delete). |

### 3.2 Notatki deweloperskie (`Notatki/`)

| Plik | Zawartość |
|---|---|
| **`CHANGELOG.md`** | Keep a Changelog format. Sekcje per milestone (M0-M7) + Fixed/Changed. |
| **`TODO.md`** | Bieżąca lista zadań — checklist do/done, kolejność import/paste w VBE. |
| **`DECISIONS.md`** | ADR-y (Architectural Decision Records). Aktualnie 7 ADR-ów: Repository Pattern, sync bez clipboard, hardcoded EmailBNC+CacheFolderPath, %TEMP% transient artifact, centralizacja routingu, hard delete pending, Activate pattern. |
| **`NOTES.md`** *(w `.gitignore`)* | Lokalny notatnik autora z decyzjami pre-implementacji. Nie commitowany. |

### 3.3 Kod źródłowy (`Source/`)

| Folder | Zawartość |
|---|---|
| **`Source/Modules/*.bas`** | Moduły VBA do importu w VBE (mod_Utils, mod_Validation, mod_UserCacheSync, mod_DataCacheSync, mod_MailSender, mod_Export, mod_Tests). |
| **`Source/Forms/*.LAYOUT.md`** | Specyfikacje UserForms — tabele kontrolek z `(Name)`, `Caption`, properties. |
| **`Source/Forms/*.code-behind.txt`** | Kod VBA do wklejenia w View Code formularza w VBE. |
| **`Source/ThisWorkbook/ThisWorkbook.code.txt`** | Kod `Workbook_Open` do wklejenia w klasę ThisWorkbook w VBE. |

---

## 4. Audience routing — jak prezentować

| Pytanie/scenariusz | Najlepszy dokument |
|---|---|
| "Co to za projekt? Po co go robimy?" | `06_business_process_v2.md` §1 (lub `06_business_process.jpg` v1 panel "Faza A") |
| "Jak działa proces zgłaszania klienta?" | `06_business_process_v2.jpg` (czysty workflow) |
| "Kto i co styka się z systemem?" | `01_system_context.jpg` ⭐ |
| "Jakie zewnętrzne zależności (do dyskusji z IT)?" | `01_system_context.jpg` + ticket do IT z `02_system_architecture.md` §IT |
| "Jak to jest zbudowane technicznie?" | `02_system_architecture.jpg` |
| "Jak wygląda struktura kodu (warstwy, moduły)?" | `05_module_architecture.jpg` |
| "Jak działa przepływ danych w aplikacji?" | `03_data_flow.md` lub `extracted/03_data_flow.md` |
| "Jak są zorganizowane dane (schema)?" | `04_data_model.md` lub `extracted/04_data_model.md` |
| "Dokąd zmierzamy w Fazie B?" | `06_business_process_v2.md` §3 + `post-release/05_module_architecture.jpg` |
| "Jak to będzie wyglądać po release?" | `post-release/05_module_architecture.jpg` |
| "Co już zrobione, co jeszcze do zrobienia?" | `BNC_Sender_PlanWdrozenia_FazaA.md` + `Notatki/TODO.md` |
| "Dlaczego konkretnie tak, a nie inaczej?" | `Notatki/DECISIONS.md` (ADR-y) |

---

## 5. Co planowane do dopisania

Lista **opcjonalna** — pokrywanie luk względem dojrzałego portfolio C4. **Brak blokerów**, ale dopełniłoby standard.

- [ ] **`extracted/01_system_context.md`** — tekstowy opis kontekstu (komplement do `01_system_context.jpg`)
- [ ] **`extracted/06_business_process.md`** — ekstrakt MD z pełnego workflow z perspektywy biznesu (kopia treści z `diagrams/06_business_process_v2.md` w `extracted/`)
- [ ] **`post-release/01_system_context.{html,jpg}`** — wariant po Fazie B (z MS SQL Server + Power Automate jako nowymi systemami zewnętrznymi)
- [ ] **`glossary.md`** — słownik skrótów branżowych (CNA, FK, BNC, DW, batch, decision diamond, audit trail, %TEMP%)
- [ ] **C4 Level 4 (Code)** — diagramy klas dla najbardziej krytycznych modułów (`mod_MailSender.DetermineRecipient`). Standardowo pomijane bo VBA tools słabe dla autogenu.

---

## 6. Workflow utrzymania dokumentacji

1. **Pisanie**: edytujesz `*.md` lub `*.html` w `doc_v2/`.
2. **Render** (jeśli zmieniony HTML): pipeline z `diagrams/README.md`:
   ```bash
   # Chrome headless → PNG
   "/c/Program Files/Google/Chrome/Application/chrome.exe" \
     --headless --disable-gpu --hide-scrollbars \
     --window-size=W,H \
     --screenshot="$(pwd -W)/NAME.png" \
     "file:///$(pwd -W)/NAME.html"
   ```
   ```powershell
   # PNG → JPG (quality 92, PowerShell System.Drawing)
   Add-Type -AssemblyName System.Drawing
   $png = [System.Drawing.Image]::FromFile("NAME.png")
   $enc = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | ? { $_.MimeType -eq 'image/jpeg' }
   $params = New-Object System.Drawing.Imaging.EncoderParameters(1)
   $params.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter([System.Drawing.Imaging.Encoder]::Quality, 92L)
   $png.Save("NAME.jpg", $enc, $params)
   $png.Dispose()
   Remove-Item "NAME.png"
   ```
3. **Update indeksu**: jeśli nowy plik — dopisz wpis do tabeli w §2 tego dokumentu.
4. **Commit + push**: typowo `docs(diagrams): ...` lub `docs(extracted): ...`.

---

**Wersja dokumentu**: utrzymywany na bieżąco z każdą zmianą `doc_v2/`. Ostatnia aktualizacja po commit'cie dodającym `01_system_context.{html,jpg}` + ten plik (`00_index.md`).
