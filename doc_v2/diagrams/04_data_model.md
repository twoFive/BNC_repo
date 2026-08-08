# Model danych — Faza A

> **Aktualizacja oryginału**: `pdfs/BNC_fazaA_04_data_model.pdf` opisuje strukturę baseline. Ten dokument **uzupełnia** o zmiany z M2.2 (rename `DontShowTutorial` → `DontShowSetupAgain`), M3.2 (hard delete pending — ADR-006) i **Schema v2** (2026-07-26 — rename `MiesiacZgloszenia` → `MiesiacObrotu`, usunięto kolumnę `Fields`; breaking change wymagający hard reset `BNC_DataCache.xlsx`, brak production data w Fazie A).
>
> 📊 **Graf**: [`04_data_model.jpg`](04_data_model.jpg)

---

## 1. `ws_UserCache` — tożsamość handlowca

> **Lokalizacja**: ukryty arkusz (`Visible = xlSheetVeryHidden`) wewnątrz pliku xlsm.
> **Format**: key-value — kolumna A = klucz, kolumna B = wartość.
> **Skala**: 1 wiersz danych usera (~11 pól), zmiany rzadkie.

### Schema

| Klucz (kolumna A) | Typ | Rola |
|---|---|---|
| `Imie` | String | Tożsamość handlowca |
| `Nazwisko` | String | Tożsamość handlowca |
| `EmailHandlowca` | String (email) | **Detekcja roli** ↓ |
| `CNA_HandlowcaID` | Long | Numer handlowca w firmie |
| `NrOddzialu` | String | Identyfikator oddziału |
| `EmailKierownika` | String (email) | **Detekcja roli** ↑ |
| `EmailBNC` | String (email) | Hardcoded `jessica.cant@swim.omg` ([ADR-003](../../Notatki/DECISIONS.md)) |
| `CacheFolderPath` | String (path) | Hardcoded `C:\BNC_CacheFolder\` ([ADR-003](../../Notatki/DECISIONS.md)) |
| `DataRejestracji` | Date | Timestamp setupu |
| `SetupCompleted` | Boolean | `True` → pomiń `frm_Setup` przy `Workbook_Open` |
| `DontShowSetupAgain` | Boolean | Preferencja usera (M2.2 — rename z `DontShowTutorial`) |

### Convention over configuration — detekcja roli

Brak osobnego pola `IsKierownik`. Rola usera **detektowana** przez porównanie:

```vba
If EmailKierownika = EmailHandlowca Then
    ' user jest kierownikiem (sam siebie wpisał) → mail wprost do BNC
Else
    ' user jest handlowcem → mail do kierownika z prośbą o forward
End If
```

Korzyść: jedna prawda o roli (nie da się mieć sprzecznych pól), brak nowego widoku w setup, brak nowych test casów. W Fazie B (SQL Server) zastąpione formalnym polem `Role` lub relacją do `tbl_Roles`.

### Synchronized backup

→ `BNC_UserCache.xlsx` (write-through cache, 1:1, jednostronny sync). ADR-001 (Repository Pattern), ADR-002 (sync bez clipboard).

---

## 2. `ws_DataCache` — historia zgłoszeń BNC

> **Lokalizacja**: ukryty arkusz (`Visible = xlSheetVeryHidden`).
> **Format**: tabela (**10 kolumn** — schema v2, 1 wiersz = 1 zgłoszenie BNC).
> **Skala**: typowo 10–30 records/miesiąc/user, 120–360/rok/user.

### Schema v2 (10 kolumn, 2026-07-26)

| # | Kolumna | Typ | Rola |
|---|---|---|---|
| 1 | `ReportID` | Long | **PK** · autoincrement (logika w VBA, `GetNextReportID = max(ID) + 1`) |
| 2 | `KlientFK` | Long | FK klienta (free-text w Fazie A, w Fazie B FK do `tbl_Clients`) |
| 3 | `NazwaKlienta` | String(3..200) | Wpisana przez handlowca |
| 4 | `MiesiacObrotu` | String "yyyy-MM" | **Miesiąc wykonania obrotu przez klienta** (business time klienta). Default = bieżący miesiąc. Rename z `MiesiacZgloszenia` w schema v2 dla jasności semantyki. |
| 5 | `CNA_HandlowcaID` | Long | **SNAPSHOT** z `ws_UserCache` przy INSERT |
| 6 | `NrOddzialu` | String | **SNAPSHOT** z `ws_UserCache` przy INSERT |
| 7 | `CreatedTimestamp` | Date | `Now()` w momencie INSERT (**system time** — audit fact, kiedy record trafił do systemu) |
| 8 | `Status` | Enum: `pending` \| `sent` | Stan wniosku · UPDATE przez `MarkAsSent` |
| 9 | `EmailRecipient` | String (email) | **AUDIT** — rzeczywisty adresat wysyłki |
| 10 | `BatchSentTimestamp` | Date | Timestamp momentu `MarkAsSent` (**system time**) |

### Schema v1 → v2 diff

| Zmiana | Powód |
|---|---|
| **Rename** `MiesiacZgloszenia` → `MiesiacObrotu` (kolumna 4) | Jaśniejsza semantyka — pole opisuje **kiedy klient wykonał obrót** (fakt biznesowy klienta), nie **kiedy handlowiec zgłasza** (co jest już w `CreatedTimestamp`). |
| **Usunięcie** `Fields` (kolumna 5, wraz z `MAX_FIELDS=1000`) | YAGNI — nieużywane pole "dodatkowe" bez konkretnego business case. Kolumny 6-11 przesunięte na 5-10. |
| **NIE dodane** `DataZgloszenia` | Rozważane, odrzucone — redundancja z `CreatedTimestamp`. Bez konkretnego scenariusza backdatingu nie ma wartości analitycznej ponad tym co już mamy. |

**Migration Faza A**: brak production data (0 userów productionowych). Manual delete starego `BNC_DataCache.xlsx` przed pierwszym otwarciem nowej wersji → `EnsureCacheFileExists` tworzy fresh 10-kolumnowy plik. Zero migration code w aplikacji (YAGNI).

### Cykl życia rekordu

```
       INSERT (AppendRecord)
              │
              ▼
        ┌──────────┐         MarkAsSent          ┌─────────┐
        │ pending  │ ─────────────────────────► │  sent   │
        └──────────┘                            └─────────┘
              │                                     (immutable
   DeleteRecord │                                   po ADR-006)
   (hard delete)│
              ▼
         ┌────────┐
         │ row    │ ◄── gone forever, brak śladu
         │ deleted│     (M3.2 nowe)
         └────────┘
```

**Sent records immutable**: `DeleteRecord` zwraca `False` dla `Status=sent`, log błędu. Audit trail chroniony — nikt nie skasuje dowodu wysyłki.

### Audit trail — fundament reklamacji

Pola **`EmailRecipient`** + **`BatchSentTimestamp`** są kluczowe dla scenariusza "BNC twierdzi, że nie dostało zgłoszenia X":

- Handlowiec mówi: "Wysłałem 5 maja do mojego kierownika `kierownik@firma.pl`" → przerzucenie odpowiedzialności do kierownika.
- Lub: "Wysłałem 5 maja wprost do BNC (`bnc@firma.pl`)" → twardy dowód kontaktu.

Z perspektywy aplikacji "wysłane do kierownika" to `Status=sent` — ale dzięki `EmailRecipient` mamy informację **komu faktycznie** wysłaliśmy.

### Snapshot przy zapisie — wzorzec ochrony historii

Pola `CNA_HandlowcaID`, `NrOddzialu`, `EmailRecipient` są **kopiami** z `ws_UserCache` (lub z `mod_MailSender.DetermineRecipient`) w momencie zapisu.

**Po co?** Jeśli handlowiec **zmieni oddział** miesiąc po wysyłce, jego stare zgłoszenia muszą **pamiętać stary oddział**. Bez snapshot byłoby: "wszystkie moje historyczne zgłoszenia z W001 nagle są z W007" — chaos audytowy.

To klasyczny wzorzec **temporal data**: separacja "current state" (`ws_UserCache`) od "historical fact" (`ws_DataCache` snapshot).

### `ReportID` po hard delete

`GetNextReportID` używa `max(ID) + 1` na bieżących wierszach.

- Usunięcie wiersza z **najwyższym** ID → kolejny INSERT dostaje to samo ID (reuse).
- Usunięcie wiersza z **niższym** ID → następny INSERT dostaje max+1 (gap zostaje).

Akceptowalne, ponieważ:
- Usunięty record nie ma `EmailRecipient` ani `BatchSentTimestamp` (był pending).
- Nie ma konfliktu audytowego z reusem.
- Local-only identifier, niewidoczny zewnętrznie.

### Synchronized backup

→ `BNC_DataCache.xlsx` (write-through cache, sync po każdej zmianie). ADR-001/002.

---

## 3. `ws_UsersRegistry` — lista wszystkich handlowców (M3.3)

> **Lokalizacja**: ukryty arkusz (`Visible = xlSheetVeryHidden`).
> **Format**: tabela (13 kolumn, 1 wiersz = 1 handlowiec).
> **Skala**: 1-5 userów per laptop w Fazie A (dev), 20+ w M7 rollout.

### Schema (13 kolumn)

| # | Kolumna | Typ | Rola |
|---|---|---|---|
| 1 | `UserID` | String | **PK** — format `UZYTKOWNIK_<N>_CNA<cna>` (patrz [ADR-008](../../Notatki/DECISIONS.md)) |
| 2-12 | canonical fields | (miks) | Kopia 11 pól z `ws_UserCache` (Imie, Nazwisko, EmailHandlowca, CNA_HandlowcaID, NrOddzialu, EmailKierownika, EmailBNC, CacheFolderPath, DataRejestracji, SetupCompleted, DontShowSetupAgain) |
| 13 | `LastLogin` | Date | Timestamp ostatniego `SwitchUser` |

### Relacja z `ws_UserCache`

- **Registry** = *storage of record* (persistent, wszyscy userzy).
- **UserCache** = *working memory* (aktywny user, szybki dostęp).
- Analogia: L1 cache + main memory w architekturze CPU.
- `SwitchUser` = zapisuje UserCache → Registry (poprzedni user), potem ładuje nowego z Registry → UserCache.

### Synchronized backup (post-M3.3, 2026-07-26)

→ `BNC_UsersRegistry.xlsx` (write-through cache, sync po każdej mutacji). ADR-001/002/008.

**Owner moduł**: `mod_UsersRegistrySync` (osobny od `mod_UserCacheSync` — refactor 2026-07-26 dla symetrii "sheet ↔ module" z ADR-001).

**Sync trigger points**:
- `AppendUserToRegistry` — nowy user dodany przez `AddNewUser`
- `UpdateLastLoginInRegistry` — `SwitchUser` update LastLogin

**Sync direction**: jednostronny `ws → xlsx`. Read-back **nie zaimplementowany** — admin może READ xlsx (visibility bez ingerencji w usera xlsm), ale nie push zmian. Bi-directional deferred do Fazy B (SQL Server) lub M7.

---

## 4. Kluczowe wzorce modelu danych

### 🗄 Hybrid cache (ADR-001, ADR-002)
- **Primary**: `ws_*Cache` / `ws_UsersRegistry` w xlsm (very hidden) — szybki dostęp, in-memory.
- **Backup**: `BNC_*Cache.xlsx` / `BNC_UsersRegistry.xlsx` w `CacheFolderPath` — bezpieczna kopia poza xlsm.
- **Sync jednostronny** worksheet → xlsx (nigdy odwrotnie). Eliminuje conflict resolution.
- **Trzy warstwy** (post-M3.3): UserCache (current user), Registry (wszyscy userzy), DataCache (zgłoszenia). Symetryczne pipeline.

### 🔒 Audit trail (ADR-005, ADR-006)
- `EmailRecipient` + `BatchSentTimestamp` jako dowód kontaktu.
- Sent records IMMUTABLE — `DeleteRecord` odmawia.
- Tylko pending mogą być usunięte (przed wysyłką, brak audytu do ochrony).

### 🎭 Convention over configuration
- Rola usera dedukowana z `EmailKierownika = EmailHandlowca`.
- Brak duplikacji informacji w schema.
- W Fazie B: formalne pole `Role` w SQL.

### 📸 Snapshot przy zapisie
- `CNA_HandlowcaID`, `NrOddzialu`, `EmailRecipient` jako kopie z UserCache.
- Zmiany w UserCache nie wpływają na historyczne wpisy.
- Ochrona przed retroactive rewrite historii.

---

## 5. Ścieżka migracji do Fazy B

### `ws_UserCache` → zostaje
Mała tabela (jeden user per laptop), mała wartość przeniesienia do SQL. Może zostać jako single-source-of-truth dla identity (faza A) lub przejść do `tbl_Users` w SQL przy okazji migracji innych komponentów.

### `ws_DataCache` → `tbl_Reports` w MS SQL Server
- Centralna baza wszystkich zgłoszeń (wcześniej rozproszone xlsx u każdego handlowca).
- Agregacja cross-team — np. "ilu nowych klientów pozyskał oddział W001 w Q2?" jako jedno zapytanie.
- Integracja z CRM / systemem finansowym (księgowość bonusów).
- Backup poza laptopem handlowca.

### Nowy stan `Status` — `awaiting_acceptance`

```
pending → awaiting_acceptance → sent
          (mail leci, kierownik    (po kliknięciu
           dostaje przycisk         "Akceptuj" w mailu)
           Akceptuj)
```

`awaiting_acceptance` to nowy state wprowadzany przez **Power Automate flow** w Fazie B:
- Aplikacja wysyła mail do kierownika z przyciskiem `[Akceptuj]` (Outlook Adaptive Cards).
- Kierownik klika `[Akceptuj]` → flow zapisuje akceptację w SQL Server + automatycznie forwarduje do BNC + notyfikuje handlowca.

### Architektonicznie

`ws_*Cache` (Faza A) → `mod_DataAccess` (Repository Pattern z ADO/EF, Faza B). **Bez zmian** w warstwach wyższych (UserForms, mod_Validation, mod_MailSender) — to siła pattern'u Repository.

---

## Powiązane ADR-y

| ADR | Dotyczy |
|---|---|
| [ADR-001](../../Notatki/DECISIONS.md) Repository Pattern | Encapsulacja dostępu do `ws_*Cache` |
| [ADR-002](../../Notatki/DECISIONS.md) Sync bez clipboard | `SyncToFile` w obu modułach Sync |
| [ADR-003](../../Notatki/DECISIONS.md) Hardcoded EmailBNC + CacheFolderPath | Pola w UserCache locked w UI |
| [ADR-005](../../Notatki/DECISIONS.md) Centralizacja routingu | Wpływa na `EmailRecipient` snapshot |
| [ADR-006](../../Notatki/DECISIONS.md) Hard delete pending | Definiuje cykl życia rekordu (pending może zniknąć, sent immutable) |

---

## Cross-reference

- Oryginalny PDF (baseline od architekta): [`../BNC_fazaA_04_data_model.pdf`](../BNC_fazaA_04_data_model.pdf)
- Surowy extracted MD (raw pdftotext): [`../extracted/04_data_model.md`](../extracted/04_data_model.md)
- Implementacja: [`mod_UserCacheSync.bas`](../../Source/Modules/mod_UserCacheSync.bas), [`mod_DataCacheSync.bas`](../../Source/Modules/mod_DataCacheSync.bas)
- Flows operujące na tych danych: [`03_data_flow_extended.jpg`](03_data_flow_extended.jpg) + [`03_data_flow_extended.md`](03_data_flow_extended.md)
