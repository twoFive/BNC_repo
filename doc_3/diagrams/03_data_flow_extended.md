# Przepływy danych — Faza A (extended z M3.2)

> **Aktualizacja oryginału**: `pdfs/BNC_fazaA_03_data_flow.pdf` opisuje Flow A i Flow B. Ten dokument **rozszerza** o **Flow C — hard delete pending records** dodany w M3.2 (commit `fd4b3bc`, ADR-006).
>
> 📊 **Graf**: [`03_data_flow_extended.jpg`](03_data_flow_extended.jpg)

---

## 1. Flow A — Dodanie zgłoszenia do batcha

> **Trigger**: user klika `btn_AddToList` w `frm_Main`.

### Kroki

1. **User wpisuje dane** w `frm_Main`: `KlientFK`, `NazwaKlienta`, `MiesiacZgloszenia`, `Fields` (opcjonalne).
2. **Walidacja** — `mod_Validation.ValidateReportData(reportData)` — sprawdza FK jako Long > 0, długość nazwy klienta 3–200, format miesiąca `yyyy-MM`, długość fields ≤ 1000.
3. **INSERT pending** — `mod_DataCacheSync.AppendRecord(reportData)`:
   - Auto-generuje `ReportID` (max+1).
   - **Snapshot** `CNA_HandlowcaID`, `NrOddzialu` z `ws_UserCache`.
   - `Status = "pending"`, `EmailRecipient = ""`, `BatchSentTimestamp = ""`.
   - `CreatedTimestamp = Now()`.
4. **`ThisWorkbook.Save`** — persyst zmian w pliku xlsm.
5. **SyncToFile** — kopia do `BNC_DataCache.xlsx` (write-through cache, ADR-001/002). Best-effort — błąd loguje, nie blokuje.
6. **UI refresh** — `RefreshPendingList` w `UserForm_Activate` (ADR-007). ListBox pokazuje record na **górze** (newest first).

### Stan końcowy

Record w `ws_DataCache` ze `Status=pending`, widoczny na liście batcha, kopia w `BNC_DataCache.xlsx`. User może powtórzyć krok 2 (typowo 10–30 zgłoszeń w miesiącu).

---

## 2. Flow B — Wysyłka batcha (Wniosek BNC)

> **Trigger**: user klika `btn_SendBatch` w `frm_Main`.

### Kroki

1. **Confirmation MsgBox** — pokazuje liczbę pending + adresat (zależnie od roli).
2. **GetPendingRecords** — `mod_DataCacheSync.GetPendingRecords()` zwraca wszystkie wiersze ze `Status=pending`.
3. **GenerateTempFile** — `mod_MailSender` tworzy `%TEMP%\BNC_Wniosek_yyyymmdd_hhmmss.xlsx` z 8 polami biznesowymi (bez `Status`/`EmailRecipient` — to pola wewnętrzne aplikacji). ADR-004.
4. **DECISION DIAMOND** — `mod_MailSender.DetermineRecipient()`:
   - `EmailKierownika = EmailHandlowca` → **user = kierownik** → `To = EmailBNC` + body "do weryfikacji".
   - `EmailKierownika ≠ EmailHandlowca` → **user = handlowiec** → `To = EmailKierownika` + body "do weryfikacji i przekazania do BNC".
5. **SendMailWithAttachment** — Outlook COM (`CreateObject("Outlook.Application")`, `MailItem.Send`). Late binding bez wymaganej referencji.
6. **MarkAsSent** — `mod_DataCacheSync.MarkAsSent(reportIDs, recipient)`:
   - UPDATE `Status = "sent"`.
   - UPDATE `EmailRecipient` = rzeczywisty adresat z kroku 4 (audit trail!).
   - UPDATE `BatchSentTimestamp = Now()`.
7. **CleanupTempFile** — kasuje `%TEMP%\BNC_Wniosek_*.xlsx`. Wykonywane też w `ErrorHandler` — żadne pliki nie zostają po awarii.
8. **UI** — MsgBox sukces + `RefreshPendingList` (lista pending teraz pusta).

### Stan końcowy

- Wszystkie pending → sent · audit trail kompletny (`EmailRecipient` + `BatchSentTimestamp` zapisane).
- `BNC_DataCache.xlsx` zsynchronizowany · plik tymczasowy w `%TEMP%` wykasowany.
- Sent records **immutable** (ADR-006) — nie da się ich usunąć przez UI ani API.

---

## 3. Flow C — Hard delete pending **[NEW M3.2 · ADR-006]**

> **Trigger**: user klika `btn_DeleteSelected` w `frm_Main`. Funkcja dodana w M3.2 dla scenariusza "user dodał typo, chce skasować przed wysyłką".

### Kroki

1. **User zaznacza pending** w `lst_PendingBatch` (single-select). Newest jest na górze listy.
2. **Click `btn_DeleteSelected`** — button enabled tylko gdy `pending.Count > 0`. Jeśli `ListIndex < 0` → MsgBox "Najpierw zaznacz zgłoszenie".
3. **Confirmation MsgBox**:
   ```
   Potwierdzenie usunięcia zgłoszenia BNC
   
   ID: 47
   Klient: Acme Sp. z o.o.
   
   Operacja nieodwracalna. Kontynuować?
   ```
   `vbDefaultButton2` — domyślnie zaznaczone "Nie" (safety net).
4. **DeleteRecord** — `mod_DataCacheSync.DeleteRecord(reportID) As Boolean`:
   - **DECISION**: znajduje wiersz po `reportID`, sprawdza `Status`.
     - `Status = "pending"` → **TAK**: `ws.Rows(r).Delete` + `ThisWorkbook.Save` + `SyncToFile` → return `True`.
     - `Status = "sent"` → **NIE**: `mod_Utils.LogError "Odmowa - status='sent'..."` → return `False`. **Sent records immutable** (ADR-006).
5. **UI refresh** — sukces: `LogInfo "Usunięto zgłoszenie ID=X"` + `RefreshPendingList` (ListBox bez tego ID, btn disabled gdy lista pusta).

### Stan końcowy

- Pending record **znika** z `ws_DataCache` (brak śladu, gone forever).
- `BNC_DataCache.xlsx` zsynchronizowany.
- Sent records nietknięte (próba `DeleteRecord(sentID)` zwraca `False`).
- `ReportID` **może być reusowany** jeśli usunięty był ostatnim — `GetNextReportID = max(ID) + 1`, więc jeśli max-owy zniknął, kolejny dostaje to samo ID. To akceptowalne — usunięty record nie ma audit trail, więc nie ma konfliktu.

### Edge cases

| Scenariusz | Zachowanie |
|---|---|
| User klika "Usuń" bez zaznaczenia | MsgBox info "Najpierw zaznacz zgłoszenie" |
| Lista pusta | `btn_DeleteSelected.Enabled = False` (szary) |
| User klika "Nie" w confirmation | Exit Sub, nic się nie dzieje |
| Próba usunięcia sent | `DeleteRecord` zwraca `False` + log błędu, MsgBox "Sprawdź Immediate Window" |
| Race condition (concurrent delete) | Nie dotyczy — single-user na single-machine |

---

## Powiązane ADR-y

| ADR | Dotyczy |
|---|---|
| [ADR-001](../DECISIONS.md) Repository Pattern | Cała data layer dla cache |
| [ADR-002](../DECISIONS.md) Sync bez clipboard | `SyncToFile` w Flow A, B, C |
| [ADR-004](../DECISIONS.md) Plik tymczasowy `%TEMP%` | Flow B krok 3 + 7 (cleanup) |
| [ADR-005](../DECISIONS.md) Centralizacja routingu | Flow B krok 4 (decision diamond) |
| [ADR-006](../DECISIONS.md) Hard delete pending | Flow C cały — definiuje politykę |
| [ADR-007](../DECISIONS.md) Activate pattern | Flow A, C — `RefreshPendingList` w `UserForm_Activate` |

---

## Cross-reference

- Oryginalny PDF (Flow A + B, bez Flow C): [`../pdfs/BNC_fazaA_03_data_flow.pdf`](../pdfs/BNC_fazaA_03_data_flow.pdf)
- Surowy extracted MD (raw pdftotext): [`../extracted/03_data_flow.md`](../extracted/03_data_flow.md)
- Implementacja: [`mod_DataCacheSync.bas`](../../Source/Modules/mod_DataCacheSync.bas), [`mod_MailSender.bas`](../../Source/Modules/mod_MailSender.bas), [`frm_Main.code-behind.txt`](../../Source/Forms/frm_Main.code-behind.txt)
- Model danych: [`04_data_model.jpg`](04_data_model.jpg) + [`04_data_model.md`](04_data_model.md)
