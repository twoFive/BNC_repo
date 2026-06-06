# Model danych — Faza A

> **Aktualizacja oryginału**: `pdfs/BNC_fazaA_04_data_model.pdf` opisuje strukturę baseline. Ten dokument **uzupełnia** o zmiany z M2.2 (rename `DontShowTutorial` → `DontShowSetupAgain`) i M3.2 (hard delete pending — ADR-006).
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
> **Format**: tabela (11 kolumn, 1 wiersz = 1 zgłoszenie BNC).
> **Skala**: typowo 10–30 records/miesiąc/user, 120–360/rok/user.

### Schema (11 kolumn)

| # | Kolumna | Typ | Rola |
|---|---|---|---|
| 1 | `ReportID` | Long | **PK** · autoincrement (logika w VBA, `GetNextReportID = max(ID) + 1`) |
| 2 | `KlientFK` | Long | FK klienta (free-text w Fazie A, w Fazie B FK do `tbl_Clients`) |
| 3 | `NazwaKlienta` | String(3..200) | Wpisana przez handlowca |
| 4 | `MiesiacZgloszenia` | String "yyyy-MM" | Default = bieżący miesiąc |
| 5 | `Fields` | String(0..1000) | Pole dodatkowe (opcjonalne) |
| 6 | `CNA_HandlowcaID` | Long | **SNAPSHOT** z `ws_UserCache` przy INSERT |
| 7 | `NrOddzialu` | String | **SNAPSHOT** z `ws_UserCache` przy INSERT |
| 8 | `CreatedTimestamp` | Date | `Now()` w momencie INSERT |
| 9 | `Status` | Enum: `pending` \| `sent` | Stan wniosku · UPDATE przez `MarkAsSent` |
| 10 | `EmailRecipient` | String (email) | **AUDIT** — rzeczywisty adresat wysyłki |
| 11 | `BatchSentTimestamp` | Date | Timestamp momentu `MarkAsSent` |

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

## 3. Kluczowe wzorce modelu danych

### 🗄 Hybrid cache (ADR-001, ADR-002)
- **Primary**: `ws_*Cache` w xlsm (very hidden) — szybki dostęp, in-memory.
- **Backup**: `BNC_*Cache.xlsx` w `CacheFolderPath` — bezpieczna kopia poza xlsm.
- **Sync jednostronny** worksheet → xlsx (nigdy odwrotnie). Eliminuje conflict resolution.

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

## 4. Ścieżka migracji do Fazy B

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
