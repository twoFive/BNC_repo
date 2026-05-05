Model danych - faza A

Struktura ws_UserCache, ws_DataCache i ich kopii xlsx

                                            Diagram: Model danych z polami audytowymi

Cel diagramu

Klasyczny ERD nie pasuje do fazy A, bo nie ma bazy danych. Model danych w fazie A to
dwa ukryte arkusze (very hidden) wewnatrz pliku xlsm jako primary storage oraz ich
synchronizowane kopie jako pliki xlsx. Diagram pokazuje strukture obu warstw, z
naciskiem na role kazdego pola - bez konkretnych przykladowych wartosci, ktore
moglyby sugerowac ze jakas wartosc jest 'wlasciwa' lub 'wzorcowa'.

BNC faza A - Model danych  Strona 1
ws_UserCache - tozsamosc handlowca

Ukryty arkusz w pliku xlsm zawierajacy 1 wiersz z danymi handlowca. Format key-value:
kolumna A to nazwa parametru, kolumna B to wartosc. Pola obejmuja podstawowe dane
tozsamosci (Imie, Nazwisko, EmailHandlowca, CNA_HandlowcaID, NrOddzialu) oraz
konfiguracje aplikacji (EmailKierownika, EmailBNC, CacheFolderPath). Pole
SetupCompleted=True oznacza, ze user przeszedl przez frm_Setup - przy nastepnym
otwarciu aplikacja omija setup i otwiera od razu frm_Main.

Detekcja roli przez convention over configuration

Aplikacja nie ma osobnego pola IsKierownik - role usera rozpoznaje przez porownanie
EmailKierownika z EmailHandlowca. Jezeli sa rowne, user jest kierownikiem (sam siebie
wpisal jako swojego kierownika). To jest swiadome uproszczenie projektowe - wzorzec
convention over configuration. Brak duplikacji informacji (jedna prawda o roli zamiast
dwoch potencjalnie sprzecznych pol), brak nowego widoku w setup, brak nowych test
casow. W fazie B z baza danych zastapimy to formalnym polem Role lub relacja do
tbl_Roles.

ws_DataCache - historia zgloszen z audit trail

Ukryty arkusz z tabela zgloszen. Kazdy wiersz to jedno zgloszenie BNC. ReportID to
autoincrement (logika ladowana z VBA bo Excel nie ma natywnego autoincrement). Pola
CNA_HandlowcaID i NrOddzialu sa snapshotami z ws_UserCache w momencie zapisu - to
celowy duplikat danych chroniacy historie. Status moze byc pending lub sent.
BatchSentTimestamp jest pusty dla pending, wypelniony dla sent. Pole EmailRecipient to
audit trail rejestrujacy do KOGO faktycznie wyslano dany rekord.

EmailRecipient - audit trail dla reklamacji

Pole EmailRecipient w ws_DataCache jest kluczowe dla scenariusza reklamacji 'BNC nie
dostalo zgloszenia'. Bez tego pola handlowiec wie tylko ze jego aplikacja zaznaczyla
Status=sent. Z polem widzi: 'wyslalem do EmailKierownika dnia X' - moze sprawdzic u
kierownika, czy ten przekazal dalej. Lub: 'wyslalem wprost do EmailBNC dnia X' - ma
dowod kontaktu. To jest fundament zaufania w trust-based system.

BNC_UserCache.xlsx - synchronized backup tozsamosci

Plik xlsx w lokalnym folderze cache (sciezka zdefiniowana w ws_UserCache w polu
CacheFolderPath - na produkcji standardowo C:\BNC_CacheFolder\). Format pliku xlsx,
nie xlsm - bez makr, bezpieczny dla antywirusow korporacyjnych. Plik ma identyczna
strukture jak ws_UserCache i jest synchronizowany 1:1 po kazdej zmianie ws_UserCache
(rzadko, glownie podczas setup).

BNC_DataCache.xlsx - synchronized backup historii

Plik xlsx w tym samym folderze co BNC_UserCache.xlsx. Identyczna struktura jak
ws_DataCache - tabela zgloszen ze statusami, recipientami i timestampami.
Synchronizowany po kazdej zmianie. Auto-recreate: jezeli plik nie istnieje przy starcie
aplikacji, aplikacja odtwarza go z aktualnej zawartosci ws_DataCache bez bledu i bez
utraty danych.

BNC faza A - Model danych  Strona 2
Kierunek synchronizacji - jednostronny

Sync jest tylko w jedna strone: worksheet do xlsx. Aplikacja nigdy nie czyta z xlsx -
traktuje go wylacznie jako write-only backup. To upraszcza logike i eliminuje wszelkie
problemy konfliktow. Dwustronna synchronizacja byla rozwazana, ale wprowadzaby
ryzyko race condition i niespojnosci.

Snapshot przy zapisie - wzorzec ochrony historycznej

ws_DataCache duplikuje pola CNA_HandlowcaID i NrOddzialu z ws_UserCache. To nie jest
blad ani redundancja - to wzorzec snapshot przy zapisie. Jezeli handlowiec zmieni
oddzial, jego stare zgloszenia musza pamietac stary oddzial. Pole EmailRecipient rowniez
jest snapshotem - rejestruje rzeczywistego adresata w momencie wysylki, nawet jezeli
pozniej user zmieni EmailKierownika lub EmailBNC w setupie.

BNC faza A - Model danych  Strona 3
