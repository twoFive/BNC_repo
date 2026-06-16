# System Context — BNC_Sender (Faza A)

> **Cel diagramu**: pokazać kto i co styka się z systemem BNC_Sender — aktorów (osoby/zespoły) i systemy zewnętrzne. Widok statyczny, **nie pokazuje** kolejności akcji (do tego służy `06_business_process_v2.jpg` lub `03_data_flow_extended.jpg`).
>
> 📊 **Graf**: [`01_system_context.jpg`](01_system_context.jpg) (C4 Level 1)

---

## 🟢 BIZNES

### Dla decydentów / komitetu wdrożeniowego

**Co dokładnie wdrażamy**: aplikację wewnętrzną dla zespołu handlowców i ich kierowników. Aplikacja zastępuje obecny manualny proces zgłaszania nowych klientów do programu BNC, w którym handlowcy ręcznie wypełniają plik Excel i wysyłają mailem do kierownika, a ten ręcznie forwarduje do zespołu BNC. Po wdrożeniu, aplikacja sama waliduje zgłoszenia, sama wybiera adresata (kierownik vs zespół BNC) i sama zapisuje audit trail.

**Skala wdrożenia w Fazie A**: około 20 handlowców (po 1 instalacji per laptop), 5 kierowników (po 1 oddziale firmy), 1 zespół BNC jako adresat docelowy. Przewidywany wolumen to ~30 zgłoszeń na handlowca miesięcznie, czyli ~600 zgłoszeń miesięcznie dla całego zespołu. Wartość biznesowa to przede wszystkim oszczędność czasu (~40 godzin manualnej pracy miesięcznie w skali zespołu) oraz audit trail w razie sporów typu „BNC nie dostało mojego zgłoszenia".

**Co Państwo zatwierdzacie podejmując decyzję o wdrożeniu**: zgodę na konfigurację po stronie IT (Trusted Locations w Excel, polityka Outlook Programmatic Access, wyjątek antywirusowy) oraz publikację gotowego pliku na firmowym OneDrive jako miejsce dystrybucji. Faza A **nie wymaga** żadnej infrastruktury serwerowej — nie potrzeba nowych licencji, nowych serwerów, nowych baz danych. Wszystko działa lokalnie na komputerze użytkownika z wykorzystaniem już istniejącej infrastruktury firmowej (Exchange do maili).

**Co znajduje się poza zakresem Fazy A**: nie ma centralnej bazy danych — każdy handlowiec ma lokalną historię zgłoszeń na swoim laptopie. Nie ma raportów cross-team (np. „ilu klientów pozyskał oddział W001 w Q2") — wymagałyby agregacji z 20 laptopów. Nie ma automatyzacji akceptacji po stronie kierownika — kierownik nadal forwarduje wnioski ręcznie. Te trzy braki są **świadomie odłożone do Fazy B** (zakres do dyskusji jako odrębna decyzja inwestycyjna — patrz `06_business_process_v2.md` §3).

---

### Dla użytkowników końcowych (handlowcy + kierownicy)

**Jak instaluję aplikację**: dział IT publikuje plik `BNC_Sender_v1.0.0.xlsm` na firmowym OneDrive. Pobierasz go na swój laptop (jednorazowo) i zapisujesz w lokalizacji wskazanej przez IT (typowo `C:\Aplikacje\BNC\`). Otwierasz plik w Excelu — przy pierwszym uruchomieniu zobaczysz formularz konfiguracyjny.

**Jednorazowa konfiguracja**: przy pierwszym otwarciu wpisujesz swoje dane służbowe (imię, nazwisko, email służbowy, CNA, numer oddziału) oraz adres email swojego kierownika. Aplikacja sama rozpozna, czy jesteś handlowcem (osobny adres kierownika), czy kierownikiem (wpisujesz swój własny adres jako „kierownika"). Konfiguracja zajmuje ~2 minuty i robi się tylko raz.

**Jak codziennie z tego korzystam**: otwierasz plik, pojawia się główny ekran „Wniosek BNC". Wprowadzasz dane klienta (FK, nazwa, miesiąc), klikasz „Dodaj do listy". Możesz dodać 10-30 zgłoszeń w trakcie miesiąca (kumulujące się w batchu). Pod koniec miesiąca klikasz „Wyślij wnioski BNC" — aplikacja jednym kliknięciem wysyła wszystkie pending zgłoszenia do odpowiedniego adresata (kierownika lub zespołu BNC). Pomyłkowe zgłoszenie można usunąć z listy przed wysyłką klikając „Usuń zaznaczone".

**Co gdy coś nie działa**: gdy plik nie otwiera makr albo Outlook pyta o pozwolenie przy wysyłce, zgłaszasz do działu IT (oni mają gotowy ticket z konfiguracją do zastosowania). Gdy nie pamiętasz, co dokładnie wysłałeś w danym miesiącu — używasz przycisku „Pokaż historię" w aplikacji, widzisz wszystkie wysłane wnioski z dokładną datą i adresatem. Gdy BNC twierdzi, że nie dostało zgłoszenia — pokazujesz im wpis z historii: „wysłałem 5 maja do `kierownik@firma.pl`" — twardy dowód, sprawa zamknięta.

---

## 🔵 TECH

### Dla architektów / tech leadów

**Diagram w klasyfikacji C4**: poziom 1 (System Context Diagram) zgodnie ze standardem Simona Browna. Pokazuje **statyczny** układ aktorów i systemów zewnętrznych wokół naszego systemu (`BNC_Sender`). Nie ma temporalności — strzałki oznaczają **relacje** („wysyła maile do", „konfiguruje"), nie sekwencję działań. Sekwencja jest w `03_data_flow_extended.jpg` (Flow A/B/C) i `06_business_process_v2.jpg` (workflow biznesowy).

**Granica systemu w fazie A**: w środku — sam plik `BNC_Sender_v*.xlsm` z UserForms (frm_Setup/Main/Log), modułami VBA i ukrytymi arkuszami (`ws_UserCache`, `ws_DataCache`) jako primary storage. Poza systemem — wszystko inne, czyli aktorzy (Handlowiec, Kierownik handlowca, Zespół BNC, Dział IT) oraz systemy zewnętrzne (OneDrive firmowy, Outlook klient lokalny, Exchange firmowy). Aplikacja w fazie A **nie ma własnego backendu** — to świadoma decyzja architektoniczna, redukująca operational overhead (brak serwera = brak monitoringu, brak deploy pipeline'u, brak SLA do utrzymania).

**Integracje zewnętrzne**: OneDrive firmowy służy wyłącznie do dystrybucji release'u (`xlsm` jako blob), aplikacja **nigdy** nie czyta z OneDrive w runtime. Outlook klient lokalny jest wywoływany przez COM (`CreateObject("Outlook.Application")`, late binding bez wymaganej referencji), używamy tylko `MailItem.Send`. Exchange firmowy nie jest dotykany bezpośrednio — leci tam SMTP via Outlook. Brak innych integracji w fazie A (brak ADO, brak REST, brak SharePoint API, brak Teams). To upraszcza testowanie i zmniejsza powierzchnię ataku.

**Constraints projektowe wpływające na C4 L1**: jeden user per laptop (brak współbieżności, brak conflict resolution), brak własnego serwera (brak centralnej historii, brak dashboardu BNC team), Outlook COM jako jedyny mechanizm wysyłki (zależność od Trust Center policy). Rozwiązanie kompromisowe — szybkie do wdrożenia, niska wartość operacyjna, wymaga inwestycji w Fazę B dla scale-up.

**Migracja do Fazy B i jej wpływ na C4 L1**: dochodzą dwa nowe systemy zewnętrzne — **MS SQL Server** (centralna baza wszystkich zgłoszeń, integracja z CRM, raporty cross-team) i **Power Automate** (flow akceptacji „1 klikiem" eliminujący ręczne forwardowanie przez kierownika). Aktorzy się **nie zmieniają** — wciąż handlowcy, kierownicy, zespół BNC, dział IT. Zmienia się natomiast czerwona dashed strzałka na diagramie (kierownik → BNC team „manualnie forwarduje") — w Fazie B zastąpiona przez integrację Power Automate → SQL Server → mail do BNC, eliminując manualny krok. Patrz wariant w `post-release/` (jeśli istnieje) lub `06_business_process_v2.md` §3.

---

### Dla działu IT / Ops

**Co IT musi przygotować przed wdrożeniem**: trzy konfiguracje po stronie environment, gotowy ticket do skopiowania znajduje się w `extracted/02_system_architecture.md` w sekcji „Wymagania dla działu IT". W skrócie: (1) dodać folder z plikiem xlsm do Trusted Locations w Excel Trust Center — bez tego makra zostaną zablokowane przy otwarciu; (2) wyjątek antywirusowy dla procesu `EXCEL.EXE` tworzącego pliki `.xlsx` w lokalizacji `%TEMP%` — bez tego niektóre polityki AV blokują generowanie pliku do wysyłki; (3) polityka „Trust access to the Outlook object model" w Outlooku — bez tego user dostaje security popup przy każdym wywołaniu `MailItem.Send`.

**Deploy i dystrybucja**: jeden plik `BNC_Sender_v*.xlsm` publikowany na firmowym OneDrive w lokalizacji dostępnej dla wszystkich handlowców. Update'y robione przez podmianę pliku (semver w nazwie pliku — `_v1.0.0.xlsm`, `_v1.0.1.xlsm`, ...) z komunikatem do userów. Brak auto-update — userzy ręcznie pobierają nową wersję. Akceptowalne przy ~20 osobach, w Fazie B (większa skala) warto rozważyć Group Policy lub Intune deployment.

**Monitoring i diagnostyka**: brak centralnego monitoringu w fazie A (brak własnego backendu = nie ma czego monitorować). Diagnostyka po stronie usera: aplikacja loguje do Immediate Window Excela (`Ctrl+G` w VBE) z timestampami przez `mod_Utils.LogInfo` / `LogError`. Synchronizowane pliki `BNC_UserCache.xlsx` i `BNC_DataCache.xlsx` w lokalnym folderze cache (`C:\BNC_CacheFolder\`) są **read-only backupem** stanu — można je otworzyć w Excelu i zweryfikować poprawność. Gdy aplikacja nie wysyła maili, w pierwszej kolejności sprawdzasz konfigurację Outlook Trust Center (punkt 3 powyżej).

**Co IT nie musi robić w Fazie A**: nie ma serwera do utrzymania (brak SQL, brak IIS, brak custom backend). Nie ma SLA do dotrzymania (lokalna aplikacja, awaria u jednego usera nie wpływa na innych). Nie ma certyfikatów, kluczy API, secret-ów do rotacji. Nie ma backupów do robienia poza standardowymi backupami laptopów (które i tak robicie). Faza B (jeśli zatwierdzona) doda MS SQL Server (typowo HA cluster) i instalację Power Automate flow — wtedy zakres pracy IT rośnie znacząco i wymaga osobnej decyzji.

---

## Powiązane dokumenty

- 🖼 [`01_system_context.jpg`](01_system_context.jpg) — wizualizacja (C4 Level 1)
- 📋 [`extracted/02_system_architecture.md`](../extracted/02_system_architecture.md) — szczegółowa architektura systemu (4 warstwy: dystrybucji, komputera handlowca, sieciowa, adresatów) + gotowy ticket do IT
- 💼 [`06_business_process_v2.md`](06_business_process_v2.md) — workflow biznesowy + opis Fazy B
- 🏗 [`02_system_architecture.jpg`](02_system_architecture.jpg) — C4 Level 2 (Container Diagram)
- 🧩 [`05_module_architecture.jpg`](05_module_architecture.jpg) — C4 Level 3 (Component Diagram)
- 📐 [`DECISIONS.md`](../../Notatki/DECISIONS.md) — Architectural Decision Records (ADR-001…007)
