# Architektura systemu — faza A

> **Komponenty aplikacji BNC_Sender i ich rozmieszczenie**
> Źródło: `../pdfs/BNC_fazaA_02_system_architecture.pdf`
> Powiązane: [`05_module_architecture.md`](05_module_architecture.md), [`03_data_flow.md`](03_data_flow.md), [`04_data_model.md`](04_data_model.md)

---

## Cel

Architektura systemu pokazuje, **z jakich komponentów** składa się aplikacja BNC_Sender w fazie A oraz **gdzie każdy komponent fizycznie żyje**.

Faza A nie wymaga żadnej infrastruktury serwerowej — cała aplikacja działa **na komputerze handlowca**, a jedyny serwer w grze to firmowy Exchange do dystrybucji maili.

---

## Cztery warstwy systemu

| Warstwa | Co zawiera | Lokalizacja |
|---|---|---|
| **Dystrybucji** | Release pliku `BNC_Sender_v0.1.0.xlsm` | OneDrive firmowy |
| **Komputera handlowca** | Plik xlsm (UserForms, moduły, ukryte arkusze `ws_UserCache`/`ws_DataCache` jako primary storage) + lokalny cache (`BNC_UserCache.xlsx`, `BNC_DataCache.xlsx`) + Outlook lokalny | `C:\…\BNC_Sender_v*.xlsm` + `C:\BNC_CacheFolder\` |
| **Sieciowa** | Standardowy serwer Exchange firmowy | Infrastruktura IT |
| **Adresatów** | Skrzynki zespołu BNC oraz kierownika handlowca (skonfigurowane w setup) | Outlook poza siecią |

---

## Hybrid cache — kluczowy wzorzec architektoniczny

Aplikacja stosuje **write-through cache**:

- **Primary storage**: ukryte arkusze (`Visible = xlSheetVeryHidden`) wewnątrz pliku xlsm.
- **Synchronized backup**: dwa osobne pliki `.xlsx` w lokalnym folderze cache.

Po **każdej** zmianie w worksheet aplikacja synchronizuje zmianę do odpowiedniego xlsx. To daje:
- **Szybkie operacje** — worksheet jest w pamięci.
- **Bezpieczeństwo** — kopia poza plikiem xlsm na wypadek korupcji macro-enabled workbooka.

Implementacja: [`mod_UserCacheSync`](../../Source/Modules/mod_UserCacheSync.bas), [`mod_DataCacheSync`](../../Source/Modules/mod_DataCacheSync.bas). ADR: [ADR-001](../DECISIONS.md), [ADR-002](../DECISIONS.md).

---

## Wysyłka batch przez Outlook COM

Mail z załącznikiem xlsx jest wysyłany przez **Outlook Application COM** (`CreateObject("Outlook.Application")`). Pipeline:

1. Plik xlsx do wysyłki **generowany ad-hoc** z aktualnych pending zgłoszeń w `ws_DataCache`.
2. Zapisany w **folderze tymczasowym** (`%TEMP%`).
3. **Dodany jako załącznik** do nowego `MailItem`.
4. **Kasowany** po wysłaniu (`CleanupTempFile` w success path **i** w error handlerze).

Adresat decydowany **dynamicznie** (decision diamond):

```
If EmailKierownika == EmailHandlowca Then
    ' user jest kierownikiem -> mail wprost do BNC
Else
    ' user jest handlowcem -> mail do kierownika z prośbą o przekazanie
End If
```

Implementacja: [`mod_MailSender`](../../Source/Modules/mod_MailSender.bas). ADR: [ADR-004](../DECISIONS.md), [ADR-005](../DECISIONS.md).

---

## Wymagania dla działu IT

> Sekcja sformułowana jako **gotowy ticket** do działu IT — można skopiować w całości i wysłać jako mail lub wkleić do systemu zgłoszeń.

```
Temat: Konfiguracja środowiska dla aplikacji BNC_Sender

Witam,

Wdrażamy aplikację wewnętrzną BNC_Sender (plik Excel z makrami VBA) dla zespołu
handlowców. Aplikacja działa lokalnie na komputerze każdego usera, bez własnego
serwera — jedynym używanym zasobem firmowym jest Outlook do wysyłki maili.

Proszę o następujące ustawienia po stronie IT:

1. Trusted Locations w Excel
   Dodanie folderu, w którym userzy będą trzymać plik aplikacji
   (np. C:\Aplikacje\BNC) jako Trusted Location w Excel Trust Center.
   Bez tego makra w pliku xlsm zostaną zablokowane przy otwarciu, a aplikacja
   przestanie działać.

2. Wyjątek antywirusa dla EXCEL.EXE w %TEMP%
   Aplikacja generuje pliki xlsx w folderze tymczasowym usera (%TEMP%) i kasuje
   je po wysłaniu maila. Niektóre polityki antywirusowe blokują takie operacje.
   Proszę o dodanie wyjątku dla procesu EXCEL.EXE tworzącego pliki xlsx
   w lokalizacji %TEMP%.

3. Outlook Programmatic Access
   Aplikacja wysyła maile przez Outlook COM (MailItem.Send). Domyślnie Outlook
   może pokazywać userowi popup security warning przy każdym wywołaniu.
   Proszę o ustawienie polityki "Trust access to the Outlook object model"
   dla aplikacji firmowych lub konkretnie dla folderu z BNC_Sender.

4. Sterownik ACE OLEDB (rezerwowo)
   W fazie A aplikacja nie korzysta z ADO ani baz danych Access. W fazie B
   planowana jest migracja do bazy danych — wtedy będzie potrzebny ACE OLEDB
   w wersji zgodnej z bitowością Office (najczęściej 64-bit). Proszę o
   weryfikację, że sterownik jest lub będzie dostępny przez GPO.
```

---

## Mechanizm pliku tymczasowego w `%TEMP%`

`%TEMP%` to zmienna środowiskowa Windows wskazująca na folder tymczasowych plików konkretnego usera. Standardowa lokalizacja:

```
C:\Users\<login>\AppData\Local\Temp\
```

Charakterystyki istotne dla aplikacji:

- **Writable** dla zalogowanego usera bez UAC (to jego własny obszar).
- **Auto-czyszczony** przez Windows Storage Sense (po kilkunastu dniach).
- **Wykluczony** z synchronizacji OneDrive domyślnie.
- **Wykluczony** z backupów w większości polityk korporacyjnych.
- **Niewidoczny** dla usera w Eksploratorze (bo `AppData` jest hidden).

Czyli **dokładnie to**, czego potrzebujemy dla plików tymczasowych: brak konfliktów synchronizacji, brak zaśmiecania backupów, automatyczny cleanup w razie czego.

ADR: [ADR-004 (`%TEMP%` jako transient artifact)](../DECISIONS.md).

---

## Ścieżka migracji do fazy B

Architektura została zaprojektowana tak, że migracja do **fazy B** (dodanie bazy danych Access oraz wprowadzenie `tbl_Clients` ze słownikiem klientów) wymaga modyfikacji **tylko warstwy synchronizacji**:

- `mod_DataCacheSync` zostanie **zastąpiony** lub uzupełniony przez `mod_DataAccess` (klasyczny Repository Pattern z ADO).
- `mod_UserCacheSync` może pozostać (UserCache to single-row identity, baza nie daje tu wartości).

Reszta architektury pozostaje **bez zmian**:
- UserForms (`frm_Setup`, `frm_Main`, `frm_Log`) — bez modyfikacji.
- Logika biznesowa (`mod_Validation`, `mod_MailSender`, `mod_Export`) — bez modyfikacji.
- `mod_Utils` — bez modyfikacji.

To jest siła wzorca **Repository Pattern** — granica między warstwami pozwala podmienić storage backend bez efektu domino.
