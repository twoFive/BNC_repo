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

<!-- Kolejne ADR-y dodawaj poniżej w trakcie implementacji modułów -->
