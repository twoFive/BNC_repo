Architektura modulow - faza A

Component Diagram (C4) - podzial kodu i zaleznosci

                                                    Diagram: Architektura modulow C4

Cel diagramu

Diagram zostal zaprojektowany zgodnie ze standardem C4 Model autorstwa Simona
Browna. Konkretnie jest to poziom 3 (Component Diagram), ktory pokazuje wewnetrzny
podzial aplikacji VBA na moduly i zaleznosci miedzy nimi w fazie A.

BNC faza A - Architektura modulow  Strona 1
Cztery warstwy aplikacji

Warstwa prezentacji zawiera trzy UserForms: frm_Setup (jednorazowa rejestracja z
samouczkiem i checkboxem 'Nie pokazuj ponownie'), frm_Main (wprowadzanie zgloszen i
wysylka batcha), frm_Log (historia i eksport literal copy). Warstwa logiki biznesowej
zawiera mod_Validation, mod_MailSender, mod_Export. Warstwa dostepu do danych to
mod_UserCacheSync i mod_DataCacheSync. Warstwa infrastruktury to mod_Utils, trzy
ukryte arkusze (ws_UserCache, ws_DataCache) oraz Outlook COM.

mod_MailSender - serce logiki kierownik vs handlowiec

mod_MailSender jest jedynym modulem, w ktorym jest implementowana decyzja
kierownik vs handlowiec. Ten modul czyta z ws_UserCache pola EmailKierownika,
EmailHandlowca i EmailBNC, porownuje je i wybiera adresata. Modul takze generuje
tresc body maila zaleznie od adresata oraz wpisuje rzeczywistego adresata do
ws_DataCache jako EmailRecipient. Reszta aplikacji nie wie o roli usera - to wlasnie
tworzy hermetyzacje (encapsulation) logiki w jednym miejscu.

Wzorzec synchronizacji - dwa moduly Sync zamiast jednego
DataAccess

W fazie B (z baza danych) bedzie istnial jeden modul mod_DataAccess - klasyczny
Repository Pattern. W fazie A nie ma bazy, ale jest hybrid cache, dlatego mamy dwa
moduly synchronizacji. Kazdy odpowiada za swoja pare worksheet+xlsx i implementuje
wzorzec write-through cache: zapis do worksheet to operacja prymarna, sync do xlsx to
operacja wtorna i best-effort. Migracja do fazy B nie wymaga modyfikacji warstwy
prezentacji ani logiki biznesowej - tylko podmiana modulow Sync na mod_DataAccess.

Regula warstw

Strzalki ida tylko w dol, nigdy w gore. UserForms wywoluja moduly logiki, moduly logiki
wywoluja moduly synchronizacji, moduly synchronizacji operuja na worksheets.
mod_Utils jest dostepny dla wszystkich (helpery, logowanie, daty) - jest ponizej
wszystkich warstw, jak biblioteka. mod_Validation nie wywoluje synchronizacji -
walidacja jest czysta, bezstanowa.

Liczba modulow - 5 logicznych + 2 sync

Ten podzial ma 7 modulow VBA. Mozliwe byloby ich zlaczenie (np. polaczyc oba sync),
ale rozdzielanie pelni dwie role: 1) jasniej widoczna odpowiedzialnosc kazdego modulu
(single responsibility), 2) latwiejsza migracja do fazy B - sync userCache moze pozostac,
sync dataCache zostanie zastapiony przez DataAccess.

Czego brakuje vs faza B

W fazie A nie ma mod_EmergencyBuffer (bufor awaryjny przy padzie VPN - bo nie ma
VPN), nie ma mod_DataAccess (bo nie ma bazy), nie ma cache klientow (bo brak
tbl_Clients). Wszystko to jest celowo proste, ale architektonicznie przygotowane pod
kierunek rozbudowy.

BNC faza A - Architektura modulow  Strona 2
