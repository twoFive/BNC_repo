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

<!-- Kolejne ADR-y dodawaj poniżej w trakcie implementacji modułów -->
