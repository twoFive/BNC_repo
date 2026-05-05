Architektura systemu - faza A

Komponenty aplikacji BNC_Sender i ich rozmieszczenie

                                                       Diagram: Architektura systemu

Cel diagramu

Diagram architektury systemu pokazuje, z jakich komponentow sklada sie aplikacja
BNC_Sender w fazie A oraz gdzie kazdy komponent fizycznie zyje. Faza A nie wymaga
zadnej infrastruktury serwerowej - cala aplikacja dziala na komputerze handlowca, a
jedyny serwer w grze to firmowy Exchange do dystrybucji maili.

BNC faza A - Architektura systemu  Strona 1
Cztery warstwy systemu

Warstwa dystrybucji to OneDrive firmowy zawierajacy release pliku
BNC_Sender_v0.1.0.xlsm. Warstwa komputera handlowca zawiera plik xlsm (z
UserForms, modulami i ukrytymi arkuszami ws_UserCache, ws_DataCache jako primary
storage), lokalny folder cache C:\BNC_CacheFolder\ z synchronizowanymi plikami xlsx
(BNC_UserCache.xlsx, BNC_DataCache.xlsx) oraz Outlook lokalny do wysylki maili.
Warstwa sieciowa to standardowy serwer Exchange firmowy. Warstwa adresatow to
skonfigurowane podczas setup skrzynki: zespolu BNC oraz kierownika handlowca.

Hybrid cache - kluczowy wzorzec architektoniczny

Aplikacja stosuje wzorzec write-through cache. Primary storage to ukryte arkusze (very
hidden) wewnatrz pliku xlsm. Synchronized backup to dwa osobne pliki xlsx zapisywane
w lokalnym folderze. Po kazdej zmianie w worksheet aplikacja synchronizuje zmiane do
odpowiedniego xlsx. To daje szybkie operacje (worksheet w pamieci) i bezpieczenstwo
(kopia poza plikiem xlsm na wypadek korupcji).

Wysylka batch przez Outlook COM

Mail z zalacznikiem xlsx jest wysylany przez Outlook Application COM. Plik xlsx do
wysylki jest generowany ad-hoc z aktualnych pending zgloszen w ws_DataCache,
zapisany w folderze tymczasowym (%TEMP%), dodany jako zalacznik, i kasowany po
wyslaniu. Adresat jest decydowany dynamicznie - jezeli user jest kierownikiem
(EmailKierownika == EmailHandlowca), mail leci wprost do BNC. Jezeli user jest
handlowcem, mail leci do jego kierownika z prosba o weryfikacje i przekazanie do BNC.

BNC faza A - Architektura systemu  Strona 2
Wymagania dla dzialu IT

Ponizsza sekcja jest sformulowana jako gotowy ticket do dzialu IT. Mozesz skopiowac ja
w calosci i wyslac jako mail lub wkleic do systemu zgloszen.

Temat: Konfiguracja srodowiska dla aplikacji BNC_Sender

Witam,

Wdrazamy aplikacje wewnetrzna BNC_Sender (plik Excel z makrami VBA) dla zespolu
handlowcow. Aplikacja dziala lokalnie na komputerze kazdego usera, bez wlasnego serwera -
jedynym uzywanym zasobem firmowym jest Outlook do wysylki maili.

Prosze o nastepujace ustawienia po stronie IT:

1. Trusted Locations w Excel
Dodanie folderu, w ktorym userzy beda trzymac plik aplikacji (np. C:\Aplikacje\BNC) jako
Trusted Location w Excel Trust Center. Bez tego makra w pliku xlsm zostana zablokowane
przy otwarciu, a aplikacja przestanie dzialac.

2. Wyjatek antywirusa dla EXCEL.EXE w %TEMP%
Aplikacja generuje pliki xlsx w folderze tymczasowym usera (%TEMP%) i kasuje je po
wyslaniu maila. Niektore polityki antywirusowe blokuja takie operacje. Prosze o dodanie
wyjatku dla procesu EXCEL.EXE tworzacego pliki xlsx w lokalizacji %TEMP%.

3. Outlook Programmatic Access
Aplikacja wysyla maile przez Outlook COM (MailItem.Send). Domyslnie Outlook moze
pokazywac userowi popup security warning przy kazdym wywolaniu. Prosze o ustawienie
polityki 'Trust access to the Outlook object model' dla aplikacji firmowych lub konkretnie dla
folderu z BNC_Sender.

4. Sterownik ACE OLEDB (rezerwowo)
W fazie A aplikacja nie korzysta z ADO ani baz danych Access. W fazie B planowana jest
migracja do bazy danych - wtedy bedzie potrzebny ACE OLEDB w wersji zgodnej z bitowoscia
Office (najczesciej 64-bit). Prosze o weryfikacje, ze sterownik jest lub bedzie dostepny przez
GPO.

Mechanizm pliku tymczasowego w %TEMP%

%TEMP% to zmienna srodowiskowa Windows wskazujaca na folder tymczasowych plikow
konkretnego usera. Standardowo: C:\Users\[login]\AppData\Local\Temp\. Folder jest
writable dla zalogowanego usera bez dodatkowych uprawnien admin, jest
auto-czyszczony przez Windows Storage Sense, jest wykluczony z synchronizacji
OneDrive i z backupow korporacyjnych - co jest dokladnie tym, czego potrzebujemy dla
plikow tymczasowych.

BNC faza A - Architektura systemu  Strona 3
Sciezka migracji do fazy B

Architektura zostala zaprojektowana tak, ze migracja do fazy B (dodanie bazy danych
Access oraz wprowadzenie tbl_Clients ze slownikiem klientow) wymaga modyfikacji tylko
warstwy synchronizacji - mod_DataCacheSync zostanie zastapiony lub uzupelniony przez
mod_DataAccess. Reszta architektury pozostaje bez zmian.

BNC faza A - Architektura systemu  Strona 4
