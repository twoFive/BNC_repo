# Architectural Decision Records — BNC_Sender Faza A

Format ADR: krótki opis decyzji, uzasadnienie, konsekwencje.

---

## ADR-000: Inicjalizacja repozytorium w `FazaA/` (nie w `BNC_Sender/`)

**Data**: 2026-05-06
**Status**: zatwierdzona

**Decyzja**: Repozytorium Git zainicjalizowane w `C:\Dev\BNC_Sender\FazaA\`, nie poziom wyżej w `C:\Dev\BNC_Sender\`.

**Uzasadnienie**: Faza A jest izolowanym etapem projektu. Faza B będzie mieć osobne repo (lub osobny katalog z osobnym `git init`) — granice fazy = granice repo. Plan implementacji dopuszczał obie opcje; wybór `FazaA/` upraszcza ścieżki w `.gitignore` i unika trackowania potencjalnych artefaktów Fazy B.

**Konsekwencja**: Wszystkie ścieżki względne w repo zaczynają się od katalogu `FazaA/`, nie od `BNC_Sender/`.

---

## ADR-001: Repository Pattern dla cache

**Data**: 2026-05-06
**Status**: zatwierdzona

**Decyzja**: Każdy ukryty arkusz (`ws_UserCache`, `ws_DataCache`) ma dedykowany moduł (`mod_UserCacheSync`, `mod_DataCacheSync`), który jest **jedynym miejscem dostępu** do tego arkusza. Reszta aplikacji woła publiczne funkcje tych modułów, nigdy nie czyta/pisze bezpośrednio do worksheets.

**Uzasadnienie**: Hermetyzacja danych. Reszta aplikacji nie wie, czy dane są w worksheet, w pliku xlsx, czy gdzieś indziej. W fazie B podmienimy implementację na bazę Access (`mod_DataAccess`) bez zmian w warstwach wyższych.

**Konsekwencja**: Trochę więcej kodu (każda operacja ma function w `mod_*Sync`), ale zero duplikacji i pełna kontrola nad zmianami. Wszystkie wywołania `ws_UserCache.Cells(...)` w kodzie poza `mod_UserCacheSync` traktuję jako **smell** do refaktora.

---

## ADR-002: Sync `worksheet → xlsx` bez clipboard

**Data**: 2026-05-06
**Status**: zatwierdzona

**Decyzja**: `SyncToFile` w obu modułach Sync używa przypisania `destWs.Range(...).Value = srcWs.UsedRange.Value` zamiast `Range.Copy` + `PasteSpecial`.

**Uzasadnienie**: Schowek systemowy jest zasobem współdzielonym. Jeśli user zrobi Ctrl+C w trakcie sync, paste może odpalić się na cudzych danych albo VBA dostanie błąd "PasteSpecial method of Range class failed". Range.Value = Range.Value to operacja in-process, deterministyczna, race-free i ~10× szybsza dla małych arkuszy.

**Konsekwencja**: Tracimy formatowanie komórek (kolory, czcionki) — ale w pliku backupu nie potrzebujemy formatowania, tylko wartości.

---

## ADR-003: `txt_EmailBNC` i `CacheFolderPath` jako hardcoded (niezawodność > elastyczność)

**Data**: 2026-05-06
**Status**: zatwierdzona (na podstawie `Notatki/NOTES.md`)

**Decyzja**: W `frm_Setup` dwa pola, które logicznie należą do "konfiguracji":
1. `txt_EmailBNC` — adres zespołu BNC (testowo: `jessica.cant@swim.omg`)
2. `txt_CacheFolderPath` — ścieżka folderu cache (`C:\BNC_CacheFolder\`)

są **hardcoded i tylko-do-odczytu** (`Locked = True`, `Enabled = False`). User nie może ich edytować w UI. Zamiast Windows folder picker (`btn_Browse`) — przycisk `btn_CreateCacheFolder` który **tworzy folder na hardkodowanej ścieżce** wywołując `mod_Utils.EnsureFolderExists`.

**Uzasadnienie**: trust-based system w którym handlowcy (~50-100 userów) konfigurują aplikację samodzielnie ma inherent risk błędnej konfiguracji:
- Adres BNC wpisany z literówką → maile lecą w eter, audit trail w `EmailRecipient` pokazuje błędny adres
- Folder cache na sieciowym dysku nieosiągalnym z drugiego biura → sync padnie cyklicznie, user zauważy po miesiącu
- Folder cache w OneDrive/Dropbox → dwustronna synchronizacja może powodować conflicts z lokalną kopią xlsx

Wartość elastyczności (każdy user może mieć inny folder/adres) jest niska — wszyscy handlowcy wysyłają do tego samego BNC, wszyscy potrzebują tej samej struktury cache. **Niezawodność wygrywa**.

Wartości są nadal zapisywane w `ws_UserCache` (przez `SaveUserData`), więc reszta aplikacji (`mod_MailSender`, `mod_DataCacheSync.SyncToFile`) odczytuje je standardową ścieżką przez `mod_UserCacheSync.GetUserField` — bez świadomości polityki hardcoded.

**Konsekwencja**:
- Zmiana `EmailBNC` (np. produkcyjne wdrożenie z prawdziwym adresem) wymaga edycji constanta `HARDCODED_EMAIL_BNC` w code-behind `frm_Setup` + redeploy xlsm. To **nie jest** zmiana konfiguracji per-user, to zmiana wersji aplikacji.
- Constanty (`HARDCODED_EMAIL_BNC`, `HARDCODED_CACHE_FOLDER`) są w jednym miejscu — `frm_Setup` code-behind. W razie potrzeby można je przenieść do `mod_Utils` jako Public Const, ale nie ma na to potrzeby w fazie A.
- Bezpiecznik: code-behind nadpisuje `txt_EmailBNC.Text` i `txt_CacheFolderPath.Text` w `UserForm_Initialize` — nawet jeśli ktoś przypadkiem odznaczy `Locked` w designerze, wartości i tak są ustawiane z `Const`-ów.

---

## ADR-004: Plik tymczasowy w `%TEMP%` jako mail attachment (transient artifact)

**Data**: 2026-05-06
**Status**: zatwierdzona

**Decyzja**: `mod_MailSender.GenerateTempFile` tworzy plik `BNC_Wniosek_<yyyymmdd_hhnnss>.xlsx` w folderze `%TEMP%` użytkownika, dodaje jako załącznik do maila przez Outlook COM, i **kasuje natychmiast po wysyłce** (`CleanupTempFile` w success path + w `ErrorHandler`).

**Uzasadnienie**: 
- `%TEMP%` (`C:\Users\[login]\AppData\Local\Temp\`) to obszar pisany przez użytkownika bez UAC, wykluczony z synchronizacji OneDrive i z backupów korporacyjnych.
- Plik jest **transient** — służy tylko jednej operacji (attachment do `MailItem.Send`). Pozostawienie go zaśmiecałoby dysk i konfundowało usera ("po co mam plik `BNC_Wniosek_20260506_142311.xlsx`?").
- Historia zgłoszeń ma swoje miejsce w `BNC_DataCache.xlsx` (zarządzane przez `mod_DataCacheSync.SyncToFile`) — to **prawdziwy** backup. Plik tymczasowy to artefakt wysyłki, nie kopia bezpieczeństwa.

**Konsekwencje**:
- `CleanupTempFile` musi być w error handlerze — jeśli `Send` padnie, plik nadal się posprząta (chronimy przed zaśmiecaniem przy retry).
- Jeśli antywirus blokuje EXCEL.EXE tworzący pliki w `%TEMP%`, send padnie. To pokryte w ticket do IT (sekcja 2 z `02_system_architecture.md`).

---

## ADR-005: Centralizacja decyzji "kierownik vs handlowiec" w `mod_MailSender.DetermineRecipient`

**Data**: 2026-05-06
**Status**: zatwierdzona

**Decyzja**: `mod_MailSender` jest **jedynym** modułem implementującym logikę decyzji `IsUserManager() ? wprost-do-BNC : do-kierownika`. Reszta aplikacji (UserForms, inne moduły) **nie wie** o roli usera w kontekście wysyłki — zna tylko `mod_UserCacheSync.IsUserManager()` jako odczyt rozróżnienia, ale routing jest enkapsulowany.

**Uzasadnienie**: 
- Encapsulation: zmiana logiki routingu (np. dodanie CC, zmiana subject) ma jeden punkt edycji.
- Testowalność: `DetermineRecipient` jest **public** (wyjątek od preferencji `Private` dla helperów) — bo to czysta funkcja bez side effects, a publiczność umożliwia automated test (`Test_mod_MailSender` weryfikuje obie gałęzie decision diamond bez wysyłania prawdziwego maila).
- Zgodność z planem: dokumentacja `05_module_architecture.md` jasno pisze _"mod_MailSender jest jedynym modulem, w ktorym jest implementowana decyzja kierownik vs handlowiec"_ — to wymaganie projektowe, nie tylko nasz wybór.

**Konsekwencje**:
- `frm_Main.btn_SendBatch_Click` woła `mod_MailSender.SendBatch()` bez znajomości routingu — pokazuje confirmation z `IsUserManager()` tylko dla UX (info kogo email leci), ale nie podejmuje decyzji.
- `EmailRecipient` w `ws_DataCache` jest zapisywany przez `mod_DataCacheSync.MarkAsSent` z wartością otrzymaną od `mod_MailSender` — **rzeczywisty** adresat (nie "logiczny"), bo to fundament audit trail dla reklamacji "BNC nie dostało zgłoszenia" (z `04_data_model.md`).

---

## ADR-006: Hard delete dla pending, sent immutable

**Data**: 2026-05-10
**Status**: zatwierdzona

**Decyzja**: `mod_DataCacheSync.DeleteRecord(reportID)`:
1. **Pending records** — usuwane przez **hard delete** (`ws.Rows(r).Delete`). `ReportID` przepada definitywnie, brak śladu w `ws_DataCache`, frm_Log nigdy ich nie pokazuje.
2. **Sent records** — **niemożliwe do usunięcia**. Defensywny check w `DeleteRecord` odmawia (return `False` + `LogError`). frm_Log nie ma buttona delete dla tych rekordów.

**Uzasadnienie**:

**Hard delete (vs soft delete `Status=cancelled`)**:
- Prostota: brak trzeciego statusu, brak statystyk pending/sent/cancelled w frm_Log, brak `MarkAsCancelled` API.
- Pending = draft = pełna kontrola usera. "Skasowane" znaczy "nigdy nie istniało" z perspektywy aplikacji.
- Soft delete miałby sens gdyby był audyt compliance ("ile rekordów user kasował?"). W tej fazie brak takiego wymagania.

**Sent immutable**:
- Zgodne z filozofią audit trail całej aplikacji (`EmailRecipient` snapshot w `ws_DataCache` istnieje **właśnie po to**, żeby user mógł odpowiedzieć BNC "wysłałem dnia X do Y, sprawdź u niego" — patrz ADR `EmailRecipient` w `04_data_model.md`).
- Usunięcie sent = zniszczenie dowodu wysyłki. Niedopuszczalne w trust-based systemie.
- Jeśli user się pomylił i wysłał błędne zgłoszenie — odpowiednia ścieżka to mail "anuluj poprzednie" do tego samego adresata, **nie** kasowanie historii.

**Konsekwencje**:
- `GetNextReportID` może **reuse** ID skasowanego pending recordu (gdy był to ostatni record). Pragmatycznie akceptowalne — hard-deleted pending nie ma śladu, więc nie ma konfliktu z reusowanym ID.
- Nigdy nie wprowadzimy "undo" dla delete — gone is gone. User dostaje confirmation MsgBox przed kasowaniem jako jedyną przeszkodę.
- `Test_mod_DataCacheSync` test cleanup nadal używa direct `ws.Rows(r).Delete` dla sent records (DeleteRecord odmówiłby) — to świadome, oznaczone komentarzem "bypassing API, test cleanup only".

---

## ADR-007: Activate pattern dla widoków `ws_DataCache`

**Data**: 2026-05-10
**Status**: zatwierdzona

**Decyzja**: UserForms które wyświetlają dane z `ws_DataCache` (`frm_Main`, `frm_Log`) ładują dane w handlerze `UserForm_Activate()`, **nie** `UserForm_Initialize()`.

`Initialize` zostaje dla **statycznego setupu** (identity captions, role detection, default values, ListBox headers) — rzeczy które nie zmieniają się w trakcie sesji.

**Uzasadnienie**: VBA UserForms lifecycle:
- `Initialize` fires **raz**, przy pierwszym instantiate'owaniu formularza (pierwszy `Show` lub `Load`)
- `Activate` fires **przy każdym** `Show`, w tym po powrocie z `Hide`

Workflow użytkownika: `frm_Main` → "Pokaż historię" → `frm_Log` → "Powrót" → `frm_Main` → delete record → "Pokaż historię" → `frm_Log` (drugi raz).

Gdyby `LoadRecords` siedział w `Initialize`, drugi `Show` `frm_Log` nie wywołałby przeładowania — ListBox pokazałby **stale data** (sprzed delete'a). Bug wystąpił już przed M3.2 (np. po `SendBatch` zmienia `Status=pending → sent`), ale delete go uwypukla — user **częściej** wraca do log po sprawdzenie "no i co, usunął się?".

**Konsekwencje**:
- `frm_Log.UserForm_Initialize` znika (był pusty poza `LoadRecords`).
- `frm_Main` ma teraz **dwa** lifecycle handlery: `Initialize` (static) + `Activate` (dynamic refresh).
- Kolejne UserForms które będą czytać z `ws_DataCache` muszą stosować ten sam pattern (np. `frm_Tutorial` w M6/M7 — nie dotyczy, samouczek nie ma danych dynamicznych).
- `frm_Setup` zostaje z `Initialize` only — operuje na `ws_UserCache` które user może zmienić tylko sam przez ten formularz; brak scenariusza stale data.

**Reguła ogólna**: jeśli formularz pokazuje dane z arkusza który **może** być zmieniony spoza tego formularza w trakcie życia instancji — refresh w `Activate`. Inaczej `Initialize` wystarczy.

---

<!-- Kolejne ADR-y dodawaj poniżej w trakcie implementacji modułów -->
