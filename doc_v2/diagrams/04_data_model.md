# Model danych — Faza A

> **Aktualizacja oryginału**: `pdfs/BNC_fazaA_04_data_model.pdf` opisuje strukturę baseline. Ten dokument **uzupełnia** o zmiany z M2.2 (rename `DontShowTutorial` → `DontShowSetupAgain`), M3.2 (hard delete pending — ADR-006), **Schema v2** (2026-07-26 — rename `MiesiacZgloszenia` → `MiesiacObrotu`, usunięto kolumnę `Fields`) oraz **ADR-009** (2026-08-08 — **usunięcie `ws_UserCache`**, Registry as sole source of truth, `ws_AppState` dla session marker).
>
> 📊 **Graf**: [`04_data_model.jpg`](04_data_model.jpg)

---

## 1. `ws_AppState` — session state marker (post-ADR-009)

> **Lokalizacja**: ukryty arkusz (`Visible = xlSheetVeryHidden`) wewnątrz pliku xlsm.
> **Format**: key-value — kolumna A = klucz, kolumna B = wartość.
> **Skala**: ~1-5 kluczy meta-stanu aplikacji, zmiany rzadkie.

### Schema (obecnie 1 klucz)

| Klucz (kolumna A) | Typ | Rola |
|---|---|---|
| `_CurrentUserID` | String | UserID aktualnie zalogowanego usera (lookup w Registry). Format `UZYTKOWNIK_<N>_CNA<cna>`. |

**Extensible** dla przyszłych session/version markers: `_LastVersion`, `_InstallDate`, `_FeatureFlags`, itp.

### Convention: `_` prefix = pole systemowe

Klucze z podkreślnikiem-prefixem to **meta-stan aplikacji** (nie business data usera). Analogicznie do konwencji w wielu językach (private/internal).

### Brak xlsx sync (świadome YAGNI)

W przeciwieństwie do UserCache/DataCache/Registry, `ws_AppState` **nie ma** `BNC_AppState.xlsx`. Powód: `_CurrentUserID` to marker sesji, nie critical data — w razie utraty user przejdzie przez picker. Symmetryczny sync xlsx nie wnosiłby wartości ([ADR-002](../../Notatki/DECISIONS.md)).

### Zniknięcie `ws_UserCache` (historyczne)

Wcześniej w tym miejscu (do 2026-08-08) był `ws_UserCache` — 11 kluczy z danymi current usera. **Usunięty przez ADR-009** jako duplikacja z Registry. Dane current usera teraz odczytywane bezpośrednio z Registry przez `mod_UsersRegistrySync.GetCurrentUserField(fieldName)` (lookup po `_CurrentUserID`).

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
| 5 | `CNA_HandlowcaID` | Long | **SNAPSHOT** z Registry (current user) przy INSERT |
| 6 | `NrOddzialu` | String | **SNAPSHOT** z Registry (current user) przy INSERT |
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

Pola `CNA_HandlowcaID`, `NrOddzialu`, `EmailRecipient` są **kopiami** z `ws_UsersRegistry` (current user row, przez `mod_UsersRegistrySync.GetCurrentUserField`) lub z `mod_MailSender.DetermineRecipient` w momencie zapisu.

**Po co?** Jeśli handlowiec **zmieni oddział** miesiąc po wysyłce, jego stare zgłoszenia muszą **pamiętać stary oddział**. Bez snapshot byłoby: "wszystkie moje historyczne zgłoszenia z W001 nagle są z W007" — chaos audytowy.

To klasyczny wzorzec **temporal data**: separacja "current state" (`ws_UsersRegistry` current user row) od "historical fact" (`ws_DataCache` snapshot).

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

## 3. `ws_UsersRegistry` — lista wszystkich handlowców (M3.3, sole source post-ADR-009)

> **Lokalizacja**: ukryty arkusz (`Visible = xlSheetVeryHidden`).
> **Format**: tabela (13 kolumn, 1 wiersz = 1 handlowiec).
> **Skala**: 1-5 userów per laptop w Fazie A (dev), 20+ w M7 rollout.
> **Post-ADR-009**: **jedyne źródło prawdy** dla danych userów. Current user data odczytywana bezpośrednio z Registry (lookup po `_CurrentUserID` z `ws_AppState`).

### Schema (13 kolumn)

| # | Kolumna | Typ | Rola |
|---|---|---|---|
| 1 | `UserID` | String | **PK** — format `UZYTKOWNIK_<N>_CNA<cna>` (patrz [ADR-008](../../Notatki/DECISIONS.md)) |
| 2 | `Imie` | String | Tożsamość handlowca |
| 3 | `Nazwisko` | String | Tożsamość handlowca |
| 4 | `EmailHandlowca` | String (email) | **Detekcja roli** ↓ |
| 5 | `CNA_HandlowcaID` | Long | Numer handlowca w firmie |
| 6 | `NrOddzialu` | String | Identyfikator oddziału |
| 7 | `EmailKierownika` | String (email) | **Detekcja roli** ↑ |
| 8 | `EmailBNC` | String (email) | Hardcoded `jessica.cant@swim.omg` ([ADR-003](../../Notatki/DECISIONS.md)) |
| 9 | `CacheFolderPath` | String (path) | Hardcoded `C:\BNC_CacheFolder\` ([ADR-003](../../Notatki/DECISIONS.md)) |
| 10 | `DataRejestracji` | Date | Timestamp setupu |
| 11 | `SetupCompleted` | Boolean | (M3.3: zawsze True dla userów w Registry) |
| 12 | `DontShowSetupAgain` | Boolean | Preferencja usera (M2.2 — rename z `DontShowTutorial`) |
| 13 | `LastLogin` | Date | Timestamp ostatniego `SwitchUser` |

### Convention over configuration — detekcja roli (ADR-005)

Brak osobnego pola `IsKierownik`. Rola usera **detektowana** przez porównanie:

```vba
If EmailKierownika = EmailHandlowca Then
    ' user jest kierownikiem (sam siebie wpisał) → mail wprost do BNC
Else
    ' user jest handlowcem → mail do kierownika z prośbą o forward
End If
```

Implementacja: `mod_UsersRegistrySync.IsUserManager()` — czyta oba pola z Registry dla current usera (post-ADR-009 — było w `mod_UserCacheSync`).

### Current user access pattern (post-ADR-009)

- `mod_UsersRegistrySync.CurrentUserID()` → czyta `_CurrentUserID` z `ws_AppState`
- `mod_UsersRegistrySync.GetCurrentUserField("Imie")` → lookup wiersza po CurrentUserID → return column value
- `mod_UsersRegistrySync.SetCurrentUserField("EmailBNC", newValue)` → lookup + write + save + sync xlsx
- Zero duplikacji z osobnym cache — Registry pełni obie role (storage + read/write path).

### Synchronized backup (post-M3.3, 2026-07-26)

→ `BNC_UsersRegistry.xlsx` (write-through cache, sync po każdej mutacji). ADR-001/002/008.

**Owner moduł**: `mod_UsersRegistrySync`.

**Sync trigger points**:
- `AppendUserToRegistry` — nowy user dodany przez `AddNewUser`
- `UpdateLastLoginInRegistry` — `SwitchUser` update LastLogin
- `SetCurrentUserField` / `UpdateCurrentUserFields` — mutation current user

**Sync direction**: jednostronny `ws → xlsx`. Read-back **nie zaimplementowany** — admin może READ xlsx (visibility bez ingerencji w usera xlsm), ale nie push zmian. Bi-directional deferred do Fazy B (SQL Server) lub M7.

---

## 4. Kluczowe wzorce modelu danych

### 🗄 Hybrid cache (ADR-001, ADR-002, ADR-009)
- **Primary storage**: `ws_UsersRegistry` / `ws_DataCache` w xlsm (very hidden) — szybki dostęp, in-memory.
- **Backup xlsx**: `BNC_UsersRegistry.xlsx` / `BNC_DataCache.xlsx` w `CacheFolderPath` — bezpieczna kopia poza xlsm.
- **Sync jednostronny** worksheet → xlsx (nigdy odwrotnie). Eliminuje conflict resolution.
- **Dwie warstwy z persistence** (post-ADR-009): Registry (wszyscy userzy + current user API), DataCache (zgłoszenia).
- **Trzeci sheet** — `ws_AppState` (session marker `_CurrentUserID`) bez xlsx sync (YAGNI — marker sesji).

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

### `ws_AppState` → zostaje jako session marker (albo migruje do SQL session)
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
- Implementacja: [`mod_AppStateSync.bas`](../../Source/Modules/mod_AppStateSync.bas), [`mod_UsersRegistrySync.bas`](../../Source/Modules/mod_UsersRegistrySync.bas), [`mod_DataCacheSync.bas`](../../Source/Modules/mod_DataCacheSync.bas)
- Flows operujące na tych danych: [`03_data_flow_extended.jpg`](03_data_flow_extended.jpg) + [`03_data_flow_extended.md`](03_data_flow_extended.md)
