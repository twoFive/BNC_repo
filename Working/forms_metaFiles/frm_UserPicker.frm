VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_UserPicker 
   Caption         =   "BNC_Sender - Wybór u¿ytkownika"
   ClientHeight    =   3552
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5016
   OleObjectBlob   =   "frm_UserPicker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_UserPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ============================================================================
'  frm_UserPicker - code-behind (wklej do okna kodu UserForm w VBE)
'  Wybor aktywnego usera z listy zarejestrowanych + opcja dodania nowego.
'  Pokazywany przez ThisWorkbook.Workbook_Open gdy GetUsersCount() > 0.
'  Patrz: Notatki/DECISIONS.md (ADR-008 - multi-user Registry pattern)
' ============================================================================

Option Explicit

' UserID kazdego item'u w cmb_Users - potrzebne do SwitchUser po kliknieciu
' btn_SelectUser (ComboBox pokazuje concat "Imie Nazwisko · CNA:num", nie UserID).
' Storage: rownolegly array indexed jak cmb_Users.ListIndex.
Private m_UserIDs() As String
Private m_UserCount As Long

' ----- Lifecycle -----------------------------------------------------------

Private Sub UserForm_Activate()
    PopulateUserList
End Sub

' ----- Buttons -------------------------------------------------------------

' Wybor zaznaczonego usera z listy. Wywoluje SwitchUser -> laduje UserCache
' z Registry, ustawia jako aktywnego, aktualizuje LastLogin.
Private Sub btn_SelectUser_Click()
    If cmb_Users.ListIndex < 0 Then
        MsgBox "Najpierw wybierz uzytkownika z listy rozwijanej.", vbInformation, "Brak wyboru"
        Exit Sub
    End If

    Dim selectedUserId As String
    selectedUserId = m_UserIDs(cmb_Users.ListIndex)

    On Error GoTo SwitchError
    mod_UsersRegistrySync.SwitchUser selectedUserId
    On Error GoTo 0

    Me.Hide
    frm_Main.Show
    Exit Sub

SwitchError:
    mod_Utils.LogError "frm_UserPicker.btn_SelectUser_Click", Err.Number, Err.Description
    MsgBox "Nie udalo sie zaladowac uzytkownika. Sprawdz Immediate Window.", _
           vbExclamation, "Blad"
End Sub

' Otwiera frm_Setup dla dodania nowego usera.
' Post-ADR-009: PrepareForNewUser (clear UserCache) usuniete - UserCache
' juz nie istnieje, frm_Setup.UserForm_Initialize nie prefilluje niczego
' (dead code prefill z pre-M3.3 usuniete), wiec po prostu Show.
' frm_Setup.btn_Save wywola AddNewUser -> nowy wiersz w Registry.
Private Sub btn_AddNew_Click()
    Me.Hide
    frm_Setup.Show
End Sub

' Anuluj: zamyka plik xlsm (nie caly Excel). Bez wybranego usera aplikacja
' nie ma sensu, ale user moze miec inne pliki otwarte - zamykamy tylko nasz.
Private Sub btn_Cancel_Click()
    Dim answer As VbMsgBoxResult
    answer = MsgBox( _
        "Aplikacja BNC_Sender wymaga wyboru uzytkownika." & vbCrLf & vbCrLf & _
        "Czy na pewno zamknac aplikacje?", _
        vbYesNo + vbQuestion, "Potwierdzenie zamkniecia")

    If answer = vbYes Then
        Me.Hide
        ThisWorkbook.Close SaveChanges:=False
    End If
End Sub

' ----- Helpers -------------------------------------------------------------

' Ladowanie listy userow z Registry do ComboBox jako concat "Imie Nazwisko · CNA:num".
' Rownolegly array m_UserIDs pamieta UserID kazdej pozycji (potrzebne
' do SwitchUser bo ComboBox pokazuje tylko display string, nie UserID).
Private Sub PopulateUserList()
    Dim users As Collection
    Set users = mod_UsersRegistrySync.GetAllUsers()

    cmb_Users.Clear

    If users.Count = 0 Then
        ReDim m_UserIDs(0 To 0)
        m_UserCount = 0
        btn_SelectUser.Enabled = False
        Exit Sub
    End If

    btn_SelectUser.Enabled = True
    ReDim m_UserIDs(0 To users.Count - 1)
    m_UserCount = users.Count

    Dim i As Long
    Dim u As Object
    Dim displayText As String
    i = 0
    For Each u In users
        ' Format: "Imie Nazwisko · CNA:num"  (M3.3 update)
        displayText = CStr(u("Imie")) & " " & _
                      CStr(u("Nazwisko")) & " · CNA:" & _
                      CStr(u("CNA_HandlowcaID"))
        cmb_Users.AddItem displayText
        m_UserIDs(i) = CStr(u("UserID"))
        i = i + 1
    Next u

    ' Domyslnie zaznaczony ostatnio zalogowany (LastLogin max) - konwencja UX
    Dim defaultIdx As Long
    defaultIdx = FindLastLoggedInIndex(users)
    If defaultIdx >= 0 Then
        cmb_Users.ListIndex = defaultIdx
    Else
        cmb_Users.ListIndex = 0  ' fallback: pierwszy user na liscie
    End If
End Sub

' Znajduje index usera z najnowszym LastLogin (dla domyslnej selekcji).
' Returns -1 gdy wszyscy maja empty LastLogin.
Private Function FindLastLoggedInIndex(users As Collection) As Long
    Dim maxDate As Date
    Dim maxIdx As Long
    Dim i As Long
    Dim u As Object
    Dim ll As Variant

    maxIdx = -1
    maxDate = DateSerial(1900, 1, 1)

    i = 0
    For Each u In users
        ll = u("LastLogin")
        If IsDate(ll) Then
            If CDate(ll) > maxDate Then
                maxDate = CDate(ll)
                maxIdx = i
            End If
        End If
        i = i + 1
    Next u

    FindLastLoggedInIndex = maxIdx
End Function


