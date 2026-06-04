# BNC_Sender — proces, wartość i rozwój

Dokument biznesowy do prezentacji architektury z przełożonym / komitetem decyzyjnym. Workflow w osobnym pliku graficznym (`06_business_process_v2.jpg`), opis i kontekst poniżej w tekście.

---

## 1. O projekcie BNC_Sender (Faza A)

**Cel**: aplikacja dla zespołu handlowców do zgłaszania nowych klientów do programu **BNC** (Bonus New Client).

### 1.1 Co zastępujemy

Dotychczasowy workflow w firmie opierał się na **ręcznych operacjach na plikach Excel**:

- Handlowiec **manualnie wypełniał** plik xlsx zgłoszenia.
- Wysyłał plik mailem do swojego kierownika.
- Kierownik **manualnie przeglądał** każdy plik, decydował o akceptacji, **ręcznie forwardował** do zespołu BNC.
- Brak ujednoliconego formatu — każdy handlowiec inaczej wypełniał kolumny, kierownik tracił czas na "porządkowanie".
- **Brak audit trail** — przy sporze "BNC twierdzi że nie dostało zgłoszenia" handlowiec mógł pokazać tylko swój wysłany mail, ale **nie miał dowodu** że kierownik faktycznie przekazał wniosek dalej.

### 1.2 Co dostarczamy w Fazie A

Aplikacja **BNC_Sender** (plik Excel z makrami VBA, instalowany lokalnie u każdego handlowca):

- **Jednolity format zgłoszenia** — TextBox-y w UserForm, walidacja CNA / email / Klient FK / długości pól.
- **Auto-routing**: aplikacja sama rozpoznaje czy user jest handlowcem czy kierownikiem (porównanie emaili w setupie) — kieruje mail wprost do BNC albo do kierownika.
- **Batch wysyłka 1 klikiem**: handlowiec może dodać 10-30 zgłoszeń w trakcie miesiąca, wysłać jednym wnioskiem.
- **Audit trail w pliku cache**: każdy wniosek zapisywany z `Status`, `EmailRecipient` (do kogo faktycznie), `BatchSentTimestamp` — dowód kontaktu w razie sporu.
- **Historia + eksport**: osobny ekran „Historia" pokazuje wszystkie wysłane wnioski; eksport do xlsx do wykorzystania w raportach.

### 1.3 Wartość biznesowa

| Aspekt | Przed (manualnie) | Po (BNC_Sender) |
|---|---|---|
| **Format** | Każdy inny — kierownik traci czas na sortowanie | Jednolity, walidowany przed wysyłką |
| **Wysyłka** | Mail per zgłoszenie | Batch 1 klikiem |
| **Routing** | Handlowiec sam pamięta gdzie wysłać | Auto-decyzja: handlowiec→kierownik, kierownik→BNC |
| **Audit** | Brak dowodu kontaktu | Zapisany adresat + timestamp w cache |
| **Skala** | ~5 minut per zgłoszenie | ~30 sekund per zgłoszenie + 1 klik na batch |

**Typowy wolumen**: ~30 zgłoszeń/handlowiec/miesiąc. Dla zespołu 20 handlowców i 5 kierowników: ~600 zgłoszeń/m-c × ~4 min oszczędności = **40 godzin manualnej pracy oszczędzonej miesięcznie**.

---

## 2. Workflow biznesowy (Faza A)

Diagram pokazuje **trzy role w procesie** (nie tytuły stanowisk) i **decision diamond** różnicujący wysyłkę zależnie od roli usera aplikacji.

📊 **Graf**: [`06_business_process_v2.jpg`](06_business_process_v2.jpg)

### Trzy role w workflow

| Lane | Co reprezentuje | Kto wchodzi w tę rolę |
|---|---|---|
| **UŻYTKOWNIK APLIKACJI** | Osoba która wprowadza zgłoszenia i klika "Wyślij" | **Handlowiec** (typowo) lub **Kierownik** (gdy zgłasza klienta osobiście) |
| **KIEROWNIK (jako odbiorca)** | Osoba która weryfikuje i forwarduje wnioski od handlowca do BNC | Kierownik handlowca — **tylko gdy user aplikacji ≠ kierownik** |
| **ZESPÓŁ BNC** | Jednostka centralna firmy rejestrująca klientów w programie | Zespół BNC |

> **Kluczowe**: Lane 1 to **rola**, nie konkretne stanowisko. Kierownik może być w tej samej osobie userem aplikacji (gdy sam zgłasza klienta) — wtedy mail leci wprost do BNC, omijając rolę "odbiorca".

### Krótki opis przepływu

1. **User pozyskuje nowego klienta** (spotkanie biznesowe).
2. **User wprowadza zgłoszenie** do aplikacji (dane klienta + miesiąc).
   - Krok 2 może powtórzyć — typowo 10–30 zgłoszeń w trakcie miesiąca.
3. **User wysyła wnioski BNC** (pojedynczo lub batchem z miesiąca).
4. **Decision diamond**: czy user aplikacji jest kierownikiem?
   - **NIE** (typowo — user jest handlowcem) → mail leci do kierownika jako "odbiorca akceptujący".
   - **TAK** (user JEST kierownikiem) → mail leci wprost do BNC, **pomija lane "odbiorca"**.
5. **Kierownik (jako odbiorca) otrzymuje wnioski mailem** od swojego handlowca.
6. **Kierownik weryfikuje i akceptuje** — **control gate**, kierownik ponosi odpowiedzialność za poprawność.
7. **Zespół BNC otrzymuje wnioski** (od kierownika-odbiorcy LUB wprost od usera-kierownika).
8. **Zespół BNC weryfikuje dane klienta** (kontrola jakości).
9. **Zespół BNC rejestruje klienta w programie BNC** — klient kwalifikuje się do bonusu.

### Mechanika decyzji w aplikacji

Aplikacja rozpoznaje rolę userka przez porównanie pól w setupie:
- `EmailKierownika` ≠ `EmailHandlowca` → user jest handlowcem → mail leci do `EmailKierownika`
- `EmailKierownika` = `EmailHandlowca` → user jest jednocześnie kierownikiem → mail leci wprost do `EmailBNC`

Brak osobnego pola "IsKierownik" — to jest **convention over configuration** (patrz `ADR` w `04_data_model.md`).

### Audit trail dla reklamacji

Każdy wniosek zapisywany jest w lokalnym pliku cache z **adresatem** (do kogo faktycznie wysłano) i **datą wysyłki**. W razie sporu „BNC nie dostało zgłoszenia X" handlowiec ma **dowód kontaktu**:
- „Wysłałem 5 maja do mojego kierownika `kierownik@firma.pl` — sprawdź u niego." → przerzucenie odpowiedzialności do kierownika.
- „Wysłałem 5 maja wprost do BNC (`bnc@firma.pl`)" → twardy dowód kontaktu z zespołem BNC.

---

## 3. Faza B — planowany rozwój

Faza A to **MVP** automatyzacji. Faza B inwestuje w **infrastrukturę centralną** i **dalszą eliminację manualnej pracy**.

### 3.1 Baza danych MS SQL Server

**Co dochodzi**: centralna baza wszystkich zgłoszeń BNC, hostowana na firmowym serwerze SQL.

**Zmiana w stosunku do Fazy A**:
- Faza A: każdy handlowiec ma **lokalną kopię** historii zgłoszeń (pliki xlsx w `C:\BNC_CacheFolder\`).
- Faza B: wszystkie zgłoszenia w **jednej bazie** — handlowiec wpisuje, baza przechowuje. Zespół BNC widzi **wszystkie wnioski** w jednym dashboardzie.

**Korzyści biznesowe**:
- **Raporty cross-team** — np. „ilu nowych klientów pozyskał oddział W001 w Q2?". W Fazie A trzeba zbierać ręcznie z 10 plików; w Fazie B jedno zapytanie SQL.
- **Integracja z CRM / systemem finansowym** — baza SQL może być źródłem danych dla automatyzacji księgowości bonusów.
- **Centralna historia** — w razie awarii laptopa handlowca historia zgłoszeń jest **bezpieczna na serwerze**.

### 3.2 Akceptacja „1 klikiem" przez kierownika (Power Automate)

**Kluczowa zmiana w workflow**: kierownik **nie forwarduje** już ręcznie wniosku do BNC. Zamiast tego:

1. **Mail przychodzący do kierownika** ma przycisk `[Akceptuj]` (technologia Outlook Adaptive Cards + Power Automate).
2. **Kliknięcie [Akceptuj]** uruchamia automatyzację Power Automate:
   1. Zapisuje akceptację w bazie SQL Server (timestamp + kierownik + status).
   2. Automatycznie przekazuje wniosek do zespołu BNC (mail + zapis w bazie).
   3. Zwraca status do handlowca — notyfikacja "Twój wniosek został zaakceptowany przez `kierownik@firma.pl` dnia X i przekazany do BNC".

**Co eliminujemy z Fazy A**:
- Manualne forwardowanie maila przez kierownika.
- Ręczne sprawdzanie czy kierownik faktycznie przekazał dalej (bo Power Automate gwarantuje).
- Pingowanie kierownika "czy już wysłałeś?".

**Korzyść biznesowa**:
- **Czas kierownika ↓ ~70%** — z ~5 minut/wniosek (otwórz, przeczytaj, zaznacz wszystkich, kliknij Forward, sprawdź adresatów, wyślij) do **~30 sekund** (otwórz mail, kliknij Akceptuj).
- **Eliminacja błędów forward'owania** — kierownik nie zapomni o wniosku, nie wyśle do złego adresata.
- **Real-time status** dla handlowca — wie od razu kiedy wniosek wyleciał do BNC.

### 3.3 Ścieżka migracji Faza A → Faza B

**Bez disruption dla handlowca**: aplikacja BNC_Sender w Fazie B **wygląda tak samo** dla handlowca (te same UserForms). Pod spodem:
- Warstwa danych zmienia się z lokalnych plików xlsx **na zapis do bazy SQL** (via ADO / połączenie sieciowe).
- Workflow z mailem dla kierownika zostaje, ale **format maila** zmienia się — zawiera przycisk akceptacji zamiast prośby o ręczne forward.

**Realizacja**:
- **Fundament** (warstwa danych): wprowadzić bazę SQL, zmienić `mod_DataCacheSync` (z Fazy A) → `mod_DataAccess` (z ADO).
- **Integracja** (Power Automate): osobny komponent na poziomie firmowym, nie w pliku xlsm. Konfigurowany przez IT/admina automatyzacji.
- **UI**: bez zmian. Handlowiec nie zauważy migracji.

---

## 4. Dane techniczne (dla referencji)

> Sekcja techniczna — można pominąć w prezentacji biznesowej.

- **Wersja aktualna**: `BNC_Sender_v0.1.0.xlsm` (faza A, w trakcie wdrożenia)
- **Stack**: Excel + VBA (faza A); + MS SQL Server + Power Automate (faza B)
- **Pełna architektura systemu**: [`02_system_architecture.jpg`](02_system_architecture.jpg)
- **Architektura modułów**: [`05_module_architecture.jpg`](05_module_architecture.jpg)
- **Wariant docelowy v1.0.0**: [`post-release/`](post-release/)
- **Repozytorium kodu**: [github.com/twoFive/BNC_repo](https://github.com/twoFive/BNC_repo)

---

## Zmiany w stosunku do `06_business_process.jpg` (v1)

**v1** (`06_business_process.{html,jpg}`): jednoplikowa kompozycja — panele Faza A intro + workflow + panele Faza B na jednym obrazie (1400×1560 px). Wszystko w jednym pliku do drukowania na A2.

**v2** (`06_business_process_v2.{html,jpg,md}`): rozdzielenie — **JPG to tylko workflow** (1400×1080 px, czytelny na ekranie), **opis biznesowy w tym MD**. Lepsze do dokumentów wieloplikowych — JPG do slajdów, MD do tekstowych materiałów (executive summary, dokument w SharePoint, etc.).

Obie wersje **opisują dokładnie ten sam projekt** — różni je tylko **forma prezentacji**.
