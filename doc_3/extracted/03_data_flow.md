Przeplyw danych - faza A

Sciezki: dodawanie zgloszen i wysylka batcha (Wniosku BNC)

                                           Diagram: Przeplyw danych z decision diamond

Cel diagramu

Diagram pokazuje dwa podstawowe przeplyw danych w aplikacji. Flow A to dodawanie
pojedynczego zgloszenia do batcha. Flow B to wysylka batcha pending zgloszen jako
Wniosek BNC. Wewnatrz Flow B znajduje sie decision diamond rozdzielajacy logike
wysylki w zaleznosci od roli usera.

BNC faza A - Przeplyw danych  Strona 1
Flow A - dodawanie zgloszenia do batcha

User wypelnia main_UserForm wprowadzajac KlientFK, NazwaKlienta i pozostale pola. Po
kliknieciu Dodaj do listy aplikacja waliduje lokalnie format pol. Po pomyslnej walidacji
INSERT do ws_DataCache z polami Status=pending i EmailRecipient=pusty. Nastepnie
ThisWorkbook.Save persystuje zmiane w pliku xlsm na dysku. Dopiero po tym
mod_DataCacheSync probuje skopiowac stan do BNC_DataCache.xlsx - jezeli sie nie
udaje, blad jest logowany ale nie blokuje usera.

Flow B - wysylka Wniosku BNC z decision diamond

User klika Wyslij Wniosek BNC. Aplikacja wybiera z ws_DataCache wszystkie wiersze ze
statusem pending. Generuje plik tymczasowy xlsx w %TEMP%. Nastepnie nastepuje
decision diamond - sprawdzenie czy user jest kierownikiem.

Logika kierownik vs handlowiec - convention over configuration

Aplikacja rozpoznaje role usera bez dodatkowych pol konfiguracyjnych - wykorzystuje
porownanie EmailKierownika z EmailHandlowca. Jezeli sa rowne, oznacza to ze user
wpisal sam siebie jako swojego kierownika - czyli jest kierownikiem. To jest wzorzec
convention over configuration. Brak nowego pola, brak nowej tabeli, jedna regula w
kodzie. Konkretnie:

 If ws_UserCache.EmailKierownika = ws_UserCache.EmailHandlowca Then
        REM User jest kierownikiem - wysylamy wprost do BNC
        To = ws_UserCache.EmailBNC
        Body = "Zalaczam wniosek BNC. Prosze o weryfikacje."
        EmailRecipient = ws_UserCache.EmailBNC

 Else
        REM User jest handlowcem - wysylamy do kierownika
        To = ws_UserCache.EmailKierownika
        Body = "Zalaczam wniosek BNC. Prosze o weryfikacje i przekazanie do " _
                 & ws_UserCache.EmailBNC
        EmailRecipient = ws_UserCache.EmailKierownika

 End If

Status sent w kontekscie audytu

Po pomyslnej wysylce UPDATE w ws_DataCache - wszystkie wyslane wiersze dostaja
Status=sent oraz BatchSentTimestamp. Dodatkowo aplikacja zapisuje kolumne
EmailRecipient z rzeczywistym adresatem. To jest audit trail - przy reklamacji BNC nie
dostalo handlowiec moze pokazac wyslalem do mojego kierownika dnia X, sprawdz u
niego. Z perspektywy aplikacji wyslane do kierownika to status sent - ale dzieki polu
EmailRecipient mamy pelna informacje, KOMU faktycznie wyslalismy.

BNC faza A - Przeplyw danych  Strona 2
Mechanizm pliku tymczasowego w %TEMP%

Czym jest %TEMP% i jak Windows go zarzadza

%TEMP% to zmienna srodowiskowa Windows wskazujaca na folder tymczasowych plikow
konkretnego usera. Standardowa lokalizacja:

 C:\Users\[login_usera]\AppData\Local\Temp\

Kazdy user Windows ma swoj wlasny folder TEMP. Charakterystyki istotne dla aplikacji:
    · Zawsze writable dla zalogowanego usera - to jego wlasny obszar, bez UAC
    · Auto-czyszczony przez Storage Sense w Windows 10/11 (po kilkunastu dniach)
    · Wykluczony z synchronizacji OneDrive domyslnie
    · Wykluczony z backupow w wiekszosci polityk korporacyjnych
    · Niewidoczny dla usera w Eksploratorze (bo AppData jest hidden)

Czy Excel potrzebuje dodatkowych pozwolen?

Krotka odpowiedz: NIE. Folder %TEMP% jest wlasnoscia usera, kazda aplikacja dzialajaca
w jego kontekscie ma do niego dostep bez dodatkowych uprawnien. VBA + Excel COM
do operacji na plikach (Workbooks.Add, SaveAs, Close, FileSystemObject.DeleteFile) to
standardowy workflow, ktory dziala out of the box.

BNC faza A - Przeplyw danych  Strona 3
Cykl zycia pliku tymczasowego krok po kroku

 Public Sub SendBatchAsAttachment()
        Dim tempFolderPath As String
        Dim tempFileName As String
        Dim tempFilePath As String

        REM 1. Sciezka do folderu TEMP
        tempFolderPath = Environ("TEMP")

        REM 2. Unikalna nazwa pliku (uniknij kolizji)
        tempFileName = "BNC_Wniosek_" _

                                & Format(Now(), "yyyymmdd_hhmmss") _
                                & ".xlsx"
        tempFilePath = tempFolderPath & "\\" & tempFileName

        REM 3. Stworz plik xlsx z aktualnym batchem
        Dim wb As Workbook
        Set wb = Workbooks.Add
        REM ... wypelnij dane z ws_DataCache WHERE Status='pending' ...
        wb.SaveAs Filename:=tempFilePath, FileFormat:=xlOpenXMLWorkbook
        wb.Close SaveChanges:=False

        REM 4. Zalacz do maila i wyslij (decision diamond - kierownik vs BNC)
        REM ... mod_MailSender.SendWithAttachment(tempFilePath) ...

        REM 5. Posprzataj
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FileExists(tempFilePath) Then

               fso.DeleteFile tempFilePath
        End If
 End Sub

Co sie dzieje gdy mail sie nie wysl

W kroku 4 wystapi blad. Plik tymczasowy zostaje na dysku (krok 5 nie zostal wykonany).
Aplikacja powinna obsluzyc blad: pokazac komunikat userowi, zaproponowac ponowna
probe, oraz wyczyscic plik tymczasowy w error handler. Klasyczny VBA error handling
uzywa wzorca On Error GoTo Cleanup z sekcja Cleanup zawierajaca DeleteFile w finally.

Dlaczego nie zostawiamy pliku jako historie

Bo historia zgloszen jest juz w BNC_DataCache.xlsx - to jest prawdziwy backup. Plik
tymczasowy w %TEMP% to transient artifact - sluzy tylko do jednej operacji (wysylki jako
zalacznik) i jest natychmiast niepotrzebny. Pozostawienie go zamiecaloby dysk i
mogloby konfundowac usera (po co mam plik BNC_Wniosek_20260505_142311.xlsx?).

BNC faza A - Przeplyw danych  Strona 4
