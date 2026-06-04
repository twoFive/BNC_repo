# Architektura modułów — faza A

> **Component Diagram (C4) — podział kodu i zależności**
> Źródło: `../pdfs/BNC_fazaA_05_module_architecture.pdf`
> Powiązane: [`02_system_architecture.md`](02_system_architecture.md), [`03_data_flow.md`](03_data_flow.md), [`04_data_model.md`](04_data_model.md)

---

## Cel

Diagram zaprojektowany zgodnie ze standardem **[C4 Model](https://c4model.com/)** autorstwa Simona Browna. Konkretnie jest to **poziom 3 (Component Diagram)**, który pokazuje wewnętrzny podział aplikacji VBA na moduły i zależności między nimi w fazie A.

---

## Cztery warstwy aplikacji

```
┌────────────────────────────────────────────────────────────────┐
│  Warstwa prezentacji                                           │
│  ─────────────────────                                         │
│   frm_Setup       frm_Main         frm_Log                     │
│  (rejestracja)   (batch + send)   (historia + export)          │
└────────────────────┬───────────────────────────────────────────┘
                     │
┌────────────────────▼───────────────────────────────────────────┐
│  Warstwa logiki biznesowej                                     │
│  ──────────────────────────                                    │
│   mod_Validation    mod_MailSender    mod_Export               │
│  (czysta, bezstanowa)  (Outlook COM)   (literal copy)          │
└────────────────────┬───────────────────────────────────────────┘
                     │
┌────────────────────▼───────────────────────────────────────────┐
│  Warstwa dostępu do danych (Repository Pattern)                │
│  ──────────────────────────────────────────────                │
│   mod_UserCacheSync          mod_DataCacheSync                 │
│  (DAO dla ws_UserCache)     (DAO dla ws_DataCache)             │
└────────────────────┬───────────────────────────────────────────┘
                     │
┌────────────────────▼───────────────────────────────────────────┐
│  Warstwa infrastruktury                                        │
│  ──────────────────────                                        │
│   mod_Utils    ws_UserCache    ws_DataCache    Outlook COM     │
│  (helpery)    (very hidden)   (very hidden)   (CreateObject)   │
└────────────────────────────────────────────────────────────────┘
```

| Warstwa | Komponenty | Rola |
|---|---|---|
| **Prezentacji** | `frm_Setup`, `frm_Main`, `frm_Log` | Interakcja z userem, wprowadzanie danych, prezentacja stanu |
| **Logiki biznesowej** | `mod_Validation`, `mod_MailSender`, `mod_Export` | Reguły walidacji, decyzja routingu maili, eksport danych |
| **Dostępu do danych** | `mod_UserCacheSync`, `mod_DataCacheSync` | Hermetyzacja arkuszy, sync do xlsx (Repository Pattern) |
| **Infrastruktury** | `mod_Utils`, ukryte arkusze, Outlook COM | Helpery (logowanie, daty, FSO), primary storage, transport poczty |

---

## `mod_MailSender` — serce logiki "kierownik vs handlowiec"

`mod_MailSender` jest **jedynym** modułem, w którym jest implementowana decyzja kierownik vs handlowiec.

Ten moduł:
1. Czyta z `ws_UserCache` pola `EmailKierownika`, `EmailHandlowca`, `EmailBNC`.
2. Porównuje je (`IsUserManager()` w `mod_UserCacheSync`).
3. **Wybiera adresata** w `DetermineRecipient()` (public dla testowalności).
4. Generuje treść body maila zależnie od adresata.
5. Wpisuje **rzeczywistego adresata** do `ws_DataCache` jako `EmailRecipient` (audit trail).

**Reszta aplikacji nie wie o roli usera** w kontekście wysyłki — to właśnie tworzy hermetyzację (encapsulation) logiki w jednym miejscu.

ADR: [ADR-005 (centralizacja routingu w mod_MailSender)](../DECISIONS.md).

Implementacja: [`mod_MailSender.bas`](../../Source/Modules/mod_MailSender.bas).

---

## Wzorzec synchronizacji — dwa moduły Sync zamiast jednego DataAccess

W fazie B (z bazą danych) będzie istniał **jeden** moduł `mod_DataAccess` — klasyczny Repository Pattern z ADO.

W fazie A nie ma bazy, ale jest **hybrid cache**, dlatego mamy dwa moduły synchronizacji. Każdy odpowiada za swoją parę `worksheet + xlsx` i implementuje wzorzec **write-through cache**:

| Operacja | Status |
|---|---|
| Zapis do worksheet | Prymarna, zawsze wykonywana |
| `ThisWorkbook.Save` | Następuje od razu po zapisie do worksheet |
| Sync do `.xlsx` | **Wtórna i best-effort** — błąd loguje, ale nie blokuje usera |

**Migracja do fazy B nie wymaga modyfikacji** warstwy prezentacji ani logiki biznesowej — tylko podmiana modułów Sync na `mod_DataAccess`.

ADR: [ADR-001 (Repository Pattern)](../DECISIONS.md), [ADR-002 (sync bez clipboard)](../DECISIONS.md).

---

## Reguła warstw

> **Strzałki idą tylko w dół, nigdy w górę.**

- UserForms wywołują **moduły logiki**.
- Moduły logiki wywołują **moduły synchronizacji**.
- Moduły synchronizacji operują **na worksheets**.

`mod_Utils` jest **dostępny dla wszystkich** (helpery, logowanie, daty) — jest poniżej wszystkich warstw, jak biblioteka.

`mod_Validation` **nie wywołuje synchronizacji** — walidacja jest **czysta, bezstanowa**. To umożliwia trywialne testy jednostkowe (`Test_mod_Validation` w `mod_Tests`) bez setupu workspace.

### Dlaczego ta reguła ma znaczenie

Złamanie reguły warstw powoduje **cykliczne zależności**:
- "UserForm wywołuje sync, który wywołuje UserForm" → cykl.
- "Logika biznesowa pyta UserForm o dane" → łamanie inversji zależności.

Architektura warstwowa z **jednokierunkowym przepływem** to klasyczny wzorzec, który ułatwia migrację do fazy B (zamiana `mod_*Sync` na `mod_DataAccess`) bez efektu domino.

---

## Liczba modułów — 5 logicznych + 2 sync

Ten podział ma **7 modułów VBA**:

| Lp. | Moduł | Warstwa | Status |
|---|---|---|---|
| 1 | `mod_Utils` | infrastruktura | ✓ M1.1 |
| 2 | `mod_UserCacheSync` | dostęp do danych | ✓ M1.2 |
| 3 | `mod_DataCacheSync` | dostęp do danych | ✓ M1.3 |
| 4 | `mod_Validation` | logika biznesowa | ✓ M2.1 |
| 5 | `mod_MailSender` | logika biznesowa | ✓ M4 |
| 6 | `mod_Export` | logika biznesowa | ✓ M5.1 |
| 7 | `ThisWorkbook` (klasa) | entry point | ✓ M2.3 |

Możliwe byłoby ich złączenie (np. połączyć oba sync), ale **rozdzielanie pełni dwie role**:
1. Jaśniej widoczna **odpowiedzialność każdego modułu** (single responsibility).
2. Łatwiejsza migracja do fazy B — sync UserCache może pozostać, sync DataCache zostanie zastąpiony przez `mod_DataAccess`.

Plus oddzielny `mod_Tests` z testami smoke dla każdego modułu (nie liczony do produkcyjnych 7).

---

## Czego brakuje vs faza B

W fazie A **nie ma**:

- **`mod_EmergencyBuffer`** — bufor awaryjny przy padzie VPN. Bo nie ma VPN; użytkownicy są lokalnie z Outlookiem.
- **`mod_DataAccess`** — bo nie ma bazy. Zastąpienie obu `mod_*Sync` przy migracji.
- **Cache klientów** (`tbl_Clients`) — bo brak słownika klientów. W fazie A `KlientFK` to free-text liczba; faza B doda walidację referencyjną.

Wszystko to jest **celowo proste**, ale architektonicznie **przygotowane pod kierunek rozbudowy** — dlatego warstwy są ostre, granice respectowane, a Repository Pattern wprowadzony już teraz.
