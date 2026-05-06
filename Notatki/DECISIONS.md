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

<!-- Kolejne ADR-y dodawaj poniżej w trakcie implementacji modułów -->
