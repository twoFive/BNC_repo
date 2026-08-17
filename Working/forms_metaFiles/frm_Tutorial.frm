VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Tutorial 
   Caption         =   "Samouczek"
   ClientHeight    =   8316.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   10980
   OleObjectBlob   =   "frm_Tutorial.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_Tutorial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Patrz: frm_Tutorial.LAYOUT.md
' ============================================================================

Option Explicit

Private Const TOTAL_PAGES As Long = 5
Private m_currentPage As Long  ' 1..TOTAL_PAGES

' ----- Lifecycle -----------------------------------------------------------

Private Sub UserForm_Initialize()
    m_currentPage = 1
    RenderPage
End Sub

' ----- Buttons -------------------------------------------------------------

Private Sub btn_Next_Click()
    If m_currentPage < TOTAL_PAGES Then
        m_currentPage = m_currentPage + 1
        RenderPage
    Else
        Me.Hide  ' na ostatniej stronie: Next == Zakoncz
    End If
End Sub

Private Sub btn_Prev_Click()
    If m_currentPage > 1 Then
        m_currentPage = m_currentPage - 1
        RenderPage
    End If
End Sub

Private Sub btn_Skip_Click()
    Me.Hide
End Sub

' ----- Rendering -----------------------------------------------------------

' Dispatcher - wolany po kazdej zmianie strony. Aktualizuje:
'   - lbl_PageIndicator (numer strony)
'   - btn_Prev.Enabled (wylaczony na 1. stronie)
'   - btn_Next.Caption (zmiana na "Zakoncz" na ostatniej stronie)
'   - lbl_PageTitle + txt_PageBody (per-page renderer wybrany przez Select Case)
Private Sub RenderPage()
    lbl_PageIndicator.Caption = "Strona " & m_currentPage & " z " & TOTAL_PAGES
    btn_Prev.Enabled = (m_currentPage > 1)

    If m_currentPage = TOTAL_PAGES Then
        btn_Next.Caption = "Zakoncz"
    Else
        btn_Next.Caption = "Dalej ->"
    End If

    Select Case m_currentPage
        Case 1: RenderPage1_Welcome
        Case 2: RenderPage2_Setup
        Case 3: RenderPage3_DailyWork
        Case 4: RenderPage4_HistoryExport
        Case 5: RenderPage5_Help
    End Select
End Sub

' ----- Per-page renderers (content z frm_Tutorial.TEXT.md, ASCII-only) -----

Private Sub RenderPage1_Welcome()
    lbl_PageTitle.Caption = "Witaj w BNC_Sender"

    Dim s As String
    s = "BNC_Sender to aplikacja do zbiorczej wysy≥ki zg≥oszeÒ BNC." & vbCrLf & vbCrLf
    s = s & "Zamiast wype≥niaÊ i wysy≥aÊ kaøde zg≥oszenie osobno, " & _
        "wpisujesz je do listy przez ca≥y miesiπc, klikasz jeden przycisk " & _
        "i mail ze wszystkimi zg≥oszeniami jest wysy≥any przez Outlook do " & _
        "w≥aúciwego adresata (do Twojego kierownika, a jeúli sam nim jesteú, " & _
        "wprost na skrzynkÍ BNC)." & vbCrLf & vbCrLf
    s = s & "Samouczek zajmie ~2 minuty. Naciúnij " & Chr(34) & "Dalej ->" & Chr(34) & "."

    txt_PageBody.text = s
End Sub

Private Sub RenderPage2_Setup()
    lbl_PageTitle.Caption = "Konfiguracja (jednorazowa)"

    Dim s As String
    s = "Przy pierwszym uruchomieniu aplikacji wype≥niasz formularz z Twoimi " & _
        "danymi s≥uøbowymi: ImiÍ, Nazwisko, Email, CNA, Nr oddzia≥u, " & _
        "Email kierownika. DziÍki temu nie musisz juø powtarzaÊ wpisywania " & _
        "tych danych w kaødym wniosku!" & vbCrLf & vbCrLf
    s = s & "UWAGA: jeúli sam jesteú Kierownikiem, wpisz swÛj email takøe w polu " & _
        Chr(34) & "Email kierownika" & Chr(34) & ". " & _
        "Aplikacja rozpozna Twojπ rolÍ automatycznie." & vbCrLf & vbCrLf
    s = s & "Pola " & Chr(34) & "Email BNC" & Chr(34) & " i " & _
        Chr(34) & "Folder cache" & Chr(34) & " sπ ustawione fabrycznie. " & _
        "Kliknij " & Chr(34) & "UtwÛrz folder cache" & Chr(34) & _
        " jeúli folder jeszcze nie istnieje na dysku." & vbCrLf & vbCrLf
    s = s & "Kliknij " & Chr(34) & "Zapisz" & Chr(34) & ". Konfiguracja gotowa - " & _
        "przechodzisz do g≥Ûwnego ekranu aplikacji (ekranu wysy≥ki wniosku)."

    txt_PageBody.text = s
End Sub

Private Sub RenderPage3_DailyWork()
    lbl_PageTitle.Caption = "Dodawanie zg≥oszeÒ i wysy≥ka"

    Dim s As String
    s = "DODAJESZ ZG£OSZENIE:" & vbCrLf
    s = s & "Wype≥nij 3 pola: Klient FK, Nazwa klienta, Miesiπc obrotu " & _
        "(YYYY-MM, domyúlnie bieøπcy). Kliknij " & Chr(34) & _
        "Dodaj do listy" & Chr(34) & "." & vbCrLf & vbCrLf
    s = s & "POMY£KA:" & vbCrLf
    s = s & "Kliknij na wiersz w liúcie pending -> " & Chr(34) & _
        "UsuÒ zaznaczone" & Chr(34) & " -> potwierdü. " & _
        "Wys≥ane zg≥oszenia sπ NIEUSUWALNE, sprawdzaj przed wysy≥kπ." & _
        vbCrLf & vbCrLf
    s = s & "WYSY£KA (najlepiej raz w miesiπcu):" & vbCrLf
    s = s & "Kliknij " & Chr(34) & "Wyúlij Wniosek BNC" & Chr(34) & _
        " -> potwierdzenie -> mail wychodzi z Outlooka. " & _
        "Lista pending czyúci siÍ, zg≥oszenia trafiajπ do historii."

    txt_PageBody.text = s
End Sub

Private Sub RenderPage4_HistoryExport()
    lbl_PageTitle.Caption = "Pokaø historiÍ - Eksport"

    Dim s As String
    s = Chr(34) & "POKAØ HISTORI " & Chr(34) & ": widzisz wszystkie zg≥oszenia " & _
        "(wys≥ane + pending) z datπ i adresatem. Przydaje siÍ gdy BNC twierdzi " & _
        Chr(34) & "nie dostaliúmy zg≥oszenia" & Chr(34) & _
        " - masz twardy dowÛd wysy≥ki." & vbCrLf & vbCrLf
    s = s & Chr(34) & "EKSPORTUJ DO PLIKU" & Chr(34) & " (w oknie historii): " & _
        "zapisuje wszystko do osobnego pliku xlsx w wybranej lokalizacji. " & _
        "Backup / dla audytora / przekazanie zastÍpcy. " & _
        "Moøesz eksportowaÊ dowolnie czÍsto."

    txt_PageBody.text = s
End Sub

Private Sub RenderPage5_Help()
    lbl_PageTitle.Caption = "Gdy coú nie dzia≥a - Kontakt"

    Dim s As String
    s = "NAJCZ STSZE:" & vbCrLf
    s = s & "- Excel pyta o makra -> klik " & Chr(34) & "W≥πcz zawartoúÊ" & _
        Chr(34) & vbCrLf
    s = s & "- Outlook pyta o pozwolenie na wysy≥kÍ -> klik " & _
        Chr(34) & "Allow" & Chr(34) & vbCrLf
    s = s & "- " & Chr(34) & "Wyúlij" & Chr(34) & _
        " nic nie robi -> sprawdü czy Outlook jest otwarty" & _
        vbCrLf & vbCrLf & vbCrLf
    s = s & "Inne problemy / Sugestie:" & vbCrLf
    s = s & "tomasz.pirszel@inter-team.com.pl" & vbCrLf & vbCrLf
    s = s & "Powodzenia!" & vbCrLf
    s = s & "Kliknij " & Chr(34) & "Zakoncz" & Chr(34) & " øeby zamknπÊ."

    txt_PageBody.text = s
End Sub


