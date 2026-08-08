Attribute VB_Name = "mod_UsersRegistrySync"
Option Explicit

' ============================================================================
'  mod_UsersRegistrySync - Repository Pattern dla ws_UsersRegistry.
'  Post-ADR-009 refactor (2026-08-08): Registry to SOLE source of truth dla
'  user data. UserCache usuniete (bylo materialized view, duplikacja).
'  Current user access przez GetCurrentUserField/SetCurrentUserField -
'  bezposredni odczyt/zapis do wiersza Registry dla aktualnego usera
'  (identyfikowanego przez _CurrentUserID w ws_AppState).
'
'  Cross-module dependencies:
'  - mod_AppStateSync (marker aktywnego usera: _CurrentUserID)
'  - Zero zaleznosci od mod_UserCacheSync (moduł usuniety)
'
'  Format UserID: UZYTKOWNIK_<autoinc>_CNA<cna> (ADR-008, Q2 decyzja).
'  Patrz: DECISIONS.md (ADR-001, ADR-008, ADR-009)
'         doc_v2/diagrams/04_data_model.md
' ============================================================================

Private Const REGISTRY_SHEET As String = "ws_UsersRegistry"
Private Const REGISTRY_CACHE_FILE_NAME As String = "BNC_UsersRegistry.xlsx"

Private Const REG_USER_ID As Long = 1
Private Const REG_IMIE As Long = 2
Private Const REG_NAZWISKO As Long = 3
Private Const REG_EMAIL_HANDLOWCA As Long = 4
Private Const REG_CNA As Long = 5
Private Const REG_NR_ODDZIALU As Long = 6
Private Const REG_EMAIL_KIEROWNIKA As Long = 7
Private Const REG_EMAIL_BNC As Long = 8
Private Const REG_CACHE_FOLDER As Long = 9
Private Const REG_DATA_REJESTRACJI As Long = 10
Private Const REG_SETUP_COMPLETED As Long = 11
Private Const REG_DONT_SHOW_SETUP As Long = 12
Private Const REG_LAST_LOGIN As Long = 13
Private Const REG_TOTAL_COLS As Long = 13

' Klucz w ws_AppState przechowujacy UserID aktualnie zalogowanego usera.
Private Const CURRENT_USER_KEY As String = "_CurrentUserID"

' ============================================================================
'  Public API - Registry (users list)
' ============================================================================

' Liczba zarejestrowanych userow. 0 = pierwsze uruchomienie.
Public Function GetUsersCount() As Long
    Dim ws As Worksheet
    Set ws = EnsureRegistrySheet()
    If ws Is Nothing Then
        GetUsersCount = 0
        Exit Function
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, REG_USER_ID).End(xlUp).row
    If lastRow < 2 Then
        GetUsersCount = 0
    Else
        GetUsersCount = lastRow - 1  ' wiersz 1 = naglowek
    End If
End Function

' UserID aktualnie zalogowanego usera. Pusty gdy nikt nie wybrany.
' Cross-module: _CurrentUserID mieszka w ws_AppState (session marker).
Public Function CurrentUserID() As String
    CurrentUserID = CStr(mod_AppStateSync.GetAppValue(CURRENT_USER_KEY))
End Function

' Wszystkie zarejestrowani userzy jako Collection of Scripting.Dictionary.
' Uzywane przez frm_UserPicker do budowy listy wyboru.
Public Function GetAllUsers() As Collection
    Set GetAllUsers = New Collection

    Dim ws As Worksheet
    Set ws = EnsureRegistrySheet()
    If ws Is Nothing Then Exit Function

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, REG_USER_ID).End(xlUp).row
    If lastRow < 2 Then Exit Function

    Dim r As Long
    For r = 2 To lastRow
        GetAllUsers.Add RowToUserDict(ws, r)
    Next r
End Function

' Przelacza aktywnego usera (trivial - ADR-009 uproszczenie).
' Rejestr sam jest source of truth, wystarczy zmienic marker.
Public Sub SwitchUser(userId As String)
    mod_AppStateSync.SetAppValue CURRENT_USER_KEY, userId
    UpdateLastLoginInRegistry userId
    mod_Utils.LogInfo "SwitchUser: aktywny user = " & userId
End Sub

' Dodaje nowego usera do Registry i przelacza na niego. Wywolywane z
' frm_Setup.btn_Save gdy pierwszy uzytkownik (Registry pusty) lub gdy
' user przyszedl z frm_UserPicker.btn_AddNew.
' Format UserID: UZYTKOWNIK_<autoinc>_CNA<cna>  (ADR-008 Q2)
' Returns: wygenerowany UserID.
Public Function AddNewUser(userData As Object) As String
    Dim newUserId As String
    newUserId = GenerateUserID(userData)
    userData("UserID") = newUserId

    AppendUserToRegistry userData
    mod_AppStateSync.SetAppValue CURRENT_USER_KEY, newUserId
    UpdateLastLoginInRegistry newUserId

    mod_Utils.LogInfo "AddNewUser: dodano " & newUserId
    AddNewUser = newUserId
End Function

' Auto-recreate: jesli BNC_UsersRegistry.xlsx nie istnieje przy starcie
' aplikacji, tworzy go z aktualnej zawartosci ws_UsersRegistry.
Public Sub EnsureRegistryCacheFileExists()
    Dim folderPath As String
    folderPath = CStr(GetCurrentUserField("CacheFolderPath"))
    If Len(folderPath) = 0 Then Exit Sub  ' setup jeszcze nieukonczony

    mod_Utils.EnsureFolderExists folderPath

    Dim fullPath As String
    fullPath = mod_Utils.JoinPath(folderPath, REGISTRY_CACHE_FILE_NAME)

    If Not mod_Utils.FileExists(fullPath) Then SyncRegistryToFile
End Sub

' ============================================================================
'  Public API - Current user access (post-ADR-009 replaces UserCache API)
' ============================================================================

' Czyta pole aktualnego usera z jego wiersza w Registry. Zwraca "" gdy:
' - brak aktywnego usera (_CurrentUserID puste)
' - user nie znaleziony w Registry
' - unknown fieldName
' Zastapienie starego mod_UserCacheSync.GetUserField.
Public Function GetCurrentUserField(fieldName As String) As Variant
    Dim userId As String
    userId = CurrentUserID()
    If Len(userId) = 0 Then
        GetCurrentUserField = ""
        Exit Function
    End If

    Dim ws As Worksheet
    Set ws = EnsureRegistrySheet()
    If ws Is Nothing Then
        GetCurrentUserField = ""
        Exit Function
    End If

    Dim r As Long
    r = FindRegistryRow(ws, userId)
    If r = 0 Then
        GetCurrentUserField = ""
        Exit Function
    End If

    Dim col As Long
    col = FieldNameToColumn(fieldName)
    If col = 0 Then
        GetCurrentUserField = ""
        Exit Function
    End If

    GetCurrentUserField = ws.Cells(r, col).Value
End Function

' Zapisuje pojedyncze pole aktualnego usera w Registry + save + sync xlsx.
' Zastapienie starego mod_UserCacheSync.SetUserField.
' No-op gdy brak aktywnego usera lub unknown fieldName.
Public Sub SetCurrentUserField(fieldName As String, value As Variant)
    Dim userId As String
    userId = CurrentUserID()
    If Len(userId) = 0 Then Exit Sub

    Dim ws As Worksheet
    Set ws = EnsureRegistrySheet()
    If ws Is Nothing Then Exit Sub

    Dim r As Long
    r = FindRegistryRow(ws, userId)
    If r = 0 Then Exit Sub

    Dim col As Long
    col = FieldNameToColumn(fieldName)
    If col = 0 Then Exit Sub

    ws.Cells(r, col).Value = value
    ThisWorkbook.Save
    SyncRegistryToFile
End Sub

' Zwraca Scripting.Dictionary z wszystkimi polami aktualnego usera.
' Zastapienie starego mod_UserCacheSync.GetUserData.
Public Function GetCurrentUserData() As Object
    Set GetCurrentUserData = CreateObject("Scripting.Dictionary")

    Dim userId As String
    userId = CurrentUserID()
    If Len(userId) = 0 Then Exit Function

    Dim ws As Worksheet
    Set ws = EnsureRegistrySheet()
    If ws Is Nothing Then Exit Function

    Dim r As Long
    r = FindRegistryRow(ws, userId)
    If r = 0 Then Exit Function

    Set GetCurrentUserData = RowToUserDict(ws, r)
End Function

' Batch update - zapisuje wiele pol aktualnego usera w Registry jednym save+sync.
' Uzywane po scenariuszach gdzie kilka pol zmienia sie naraz (np. profile edit).
' Zastapienie starego mod_UserCacheSync.SaveUserData (dla current user context).
Public Sub UpdateCurrentUserFields(fieldsDict As Object)
    Dim userId As String
    userId = CurrentUserID()
    If Len(userId) = 0 Then Exit Sub
    If fieldsDict Is Nothing Then Exit Sub

    Dim ws As Worksheet
    Set ws = EnsureRegistrySheet()
    If ws Is Nothing Then Exit Sub

    Dim r As Long
    r = FindRegistryRow(ws, userId)
    If r = 0 Then Exit Sub

    Dim k As Variant
    For Each k In fieldsDict.Keys
        Dim col As Long
        col = FieldNameToColumn(CStr(k))
        If col > 0 Then
            ws.Cells(r, col).Value = fieldsDict(k)
        End If
    Next k

    ThisWorkbook.Save
    SyncRegistryToFile
End Sub

' Convention over configuration (ADR-005): user jest kierownikiem gdy sam
' siebie wpisal jako EmailKierownika. Zastapienie starego
' mod_UserCacheSync.IsUserManager - teraz czyta bezposrednio z Registry.
Public Function IsUserManager() As Boolean
    Dim handlowca As String, kierownika As String
    handlowca = LCase$(Trim$(CStr(GetCurrentUserField("EmailHandlowca"))))
    kierownika = LCase$(Trim$(CStr(GetCurrentUserField("EmailKierownika"))))

    If Len(handlowca) = 0 Or Len(kierownika) = 0 Then
        IsUserManager = False
        Exit Function
    End If

    IsUserManager = (handlowca = kierownika)
End Function

' ============================================================================
'  Private - Registry helpers
' ============================================================================

' Field name -> column index mapping (dla GetCurrentUserField/SetCurrentUserField).
' 0 = unknown field name.
Private Function FieldNameToColumn(fieldName As String) As Long
    Select Case fieldName
        Case "UserID": FieldNameToColumn = REG_USER_ID
        Case "Imie": FieldNameToColumn = REG_IMIE
        Case "Nazwisko": FieldNameToColumn = REG_NAZWISKO
        Case "EmailHandlowca": FieldNameToColumn = REG_EMAIL_HANDLOWCA
        Case "CNA_HandlowcaID": FieldNameToColumn = REG_CNA
        Case "NrOddzialu": FieldNameToColumn = REG_NR_ODDZIALU
        Case "EmailKierownika": FieldNameToColumn = REG_EMAIL_KIEROWNIKA
        Case "EmailBNC": FieldNameToColumn = REG_EMAIL_BNC
        Case "CacheFolderPath": FieldNameToColumn = REG_CACHE_FOLDER
        Case "DataRejestracji": FieldNameToColumn = REG_DATA_REJESTRACJI
        Case "SetupCompleted": FieldNameToColumn = REG_SETUP_COMPLETED
        Case "DontShowSetupAgain": FieldNameToColumn = REG_DONT_SHOW_SETUP
        Case "LastLogin": FieldNameToColumn = REG_LAST_LOGIN
        Case Else: FieldNameToColumn = 0
    End Select
End Function

' Buduje Dictionary z wiersza Registry. Uzywane przez GetAllUsers +
' GetCurrentUserData - zero duplikacji mapping code.
Private Function RowToUserDict(ws As Worksheet, r As Long) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("UserID") = CStr(ws.Cells(r, REG_USER_ID).Value)
    d("Imie") = CStr(ws.Cells(r, REG_IMIE).Value)
    d("Nazwisko") = CStr(ws.Cells(r, REG_NAZWISKO).Value)
    d("EmailHandlowca") = CStr(ws.Cells(r, REG_EMAIL_HANDLOWCA).Value)
    d("CNA_HandlowcaID") = ws.Cells(r, REG_CNA).Value
    d("NrOddzialu") = CStr(ws.Cells(r, REG_NR_ODDZIALU).Value)
    d("EmailKierownika") = CStr(ws.Cells(r, REG_EMAIL_KIEROWNIKA).Value)
    d("EmailBNC") = CStr(ws.Cells(r, REG_EMAIL_BNC).Value)
    d("CacheFolderPath") = CStr(ws.Cells(r, REG_CACHE_FOLDER).Value)
    d("DataRejestracji") = ws.Cells(r, REG_DATA_REJESTRACJI).Value
    d("SetupCompleted") = ws.Cells(r, REG_SETUP_COMPLETED).Value
    d("DontShowSetupAgain") = ws.Cells(r, REG_DONT_SHOW_SETUP).Value
    d("LastLogin") = ws.Cells(r, REG_LAST_LOGIN).Value
    Set RowToUserDict = d
End Function

' Zwraca ws_UsersRegistry, tworzac go jesli nie istnieje. Sheet very hidden.
' Ustawia naglowek w wierszu 1. Idempotentne.
Private Function EnsureRegistrySheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(REGISTRY_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        Application.ScreenUpdating = False
        On Error GoTo CreateError
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = REGISTRY_SHEET
        ws.Visible = xlSheetVeryHidden
        On Error GoTo 0
        Application.ScreenUpdating = True
        mod_Utils.LogInfo "EnsureRegistrySheet: utworzono " & REGISTRY_SHEET
    End If

    EnsureRegistryHeader ws
    Set EnsureRegistrySheet = ws
    Exit Function

CreateError:
    mod_Utils.LogError "mod_UsersRegistrySync.EnsureRegistrySheet", Err.Number, Err.Description
    Application.ScreenUpdating = True
    Set EnsureRegistrySheet = Nothing
End Function

' Ustawia naglowek w wierszu 1 Registry. Idempotentne (skip gdy juz jest).
Private Sub EnsureRegistryHeader(ws As Worksheet)
    If Len(CStr(ws.Cells(1, REG_USER_ID).Value)) > 0 Then Exit Sub

    ws.Cells(1, REG_USER_ID).Value = "UserID"
    ws.Cells(1, REG_IMIE).Value = "Imie"
    ws.Cells(1, REG_NAZWISKO).Value = "Nazwisko"
    ws.Cells(1, REG_EMAIL_HANDLOWCA).Value = "EmailHandlowca"
    ws.Cells(1, REG_CNA).Value = "CNA_HandlowcaID"
    ws.Cells(1, REG_NR_ODDZIALU).Value = "NrOddzialu"
    ws.Cells(1, REG_EMAIL_KIEROWNIKA).Value = "EmailKierownika"
    ws.Cells(1, REG_EMAIL_BNC).Value = "EmailBNC"
    ws.Cells(1, REG_CACHE_FOLDER).Value = "CacheFolderPath"
    ws.Cells(1, REG_DATA_REJESTRACJI).Value = "DataRejestracji"
    ws.Cells(1, REG_SETUP_COMPLETED).Value = "SetupCompleted"
    ws.Cells(1, REG_DONT_SHOW_SETUP).Value = "DontShowSetupAgain"
    ws.Cells(1, REG_LAST_LOGIN).Value = "LastLogin"
End Sub

' Format: UZYTKOWNIK_<N>_CNA<cna>  gdzie N = kolejny autoinc, cna z userData.
Private Function GenerateUserID(userData As Object) As String
    Dim nextN As Long
    nextN = GetUsersCount() + 1

    Dim cnaStr As String
    If userData Is Nothing Then
        cnaStr = "0"
    ElseIf userData.Exists("CNA_HandlowcaID") Then
        cnaStr = CStr(userData("CNA_HandlowcaID"))
    Else
        cnaStr = "0"
    End If

    GenerateUserID = "UZYTKOWNIK_" & nextN & "_CNA" & cnaStr
End Function

' Znajduje wiersz w Registry po UserID. Zwraca 0 gdy nie znaleziono.
Private Function FindRegistryRow(ws As Worksheet, userId As String) As Long
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, REG_USER_ID).End(xlUp).row
    If lastRow < 2 Then
        FindRegistryRow = 0
        Exit Function
    End If

    Dim r As Long
    For r = 2 To lastRow
        If StrComp(CStr(ws.Cells(r, REG_USER_ID).Value), userId, vbBinaryCompare) = 0 Then
            FindRegistryRow = r
            Exit Function
        End If
    Next r
    FindRegistryRow = 0
End Function

' Dopisuje nowego usera do Registry + save + sync xlsx.
Private Sub AppendUserToRegistry(userData As Object)
    Dim ws As Worksheet
    Set ws = EnsureRegistrySheet()
    If ws Is Nothing Then Exit Sub

    Dim r As Long
    r = ws.Cells(ws.Rows.Count, REG_USER_ID).End(xlUp).row + 1
    If r < 2 Then r = 2

    ws.Cells(r, REG_USER_ID).Value = SafeGet(userData, "UserID")
    ws.Cells(r, REG_IMIE).Value = SafeGet(userData, "Imie")
    ws.Cells(r, REG_NAZWISKO).Value = SafeGet(userData, "Nazwisko")
    ws.Cells(r, REG_EMAIL_HANDLOWCA).Value = SafeGet(userData, "EmailHandlowca")
    ws.Cells(r, REG_CNA).Value = SafeGet(userData, "CNA_HandlowcaID")
    ws.Cells(r, REG_NR_ODDZIALU).Value = SafeGet(userData, "NrOddzialu")
    ws.Cells(r, REG_EMAIL_KIEROWNIKA).Value = SafeGet(userData, "EmailKierownika")
    ws.Cells(r, REG_EMAIL_BNC).Value = SafeGet(userData, "EmailBNC")
    ws.Cells(r, REG_CACHE_FOLDER).Value = SafeGet(userData, "CacheFolderPath")
    ws.Cells(r, REG_DATA_REJESTRACJI).Value = SafeGet(userData, "DataRejestracji")
    ws.Cells(r, REG_SETUP_COMPLETED).Value = SafeGet(userData, "SetupCompleted")
    ws.Cells(r, REG_DONT_SHOW_SETUP).Value = SafeGet(userData, "DontShowSetupAgain")
    ws.Cells(r, REG_LAST_LOGIN).Value = Now()

    ThisWorkbook.Save
    SyncRegistryToFile
End Sub

' Update LastLogin dla usera w Registry + save + sync xlsx.
Private Sub UpdateLastLoginInRegistry(userId As String)
    Dim ws As Worksheet
    Set ws = EnsureRegistrySheet()
    If ws Is Nothing Then Exit Sub

    Dim r As Long
    r = FindRegistryRow(ws, userId)
    If r = 0 Then Exit Sub

    ws.Cells(r, REG_LAST_LOGIN).Value = Now()
    ThisWorkbook.Save
    SyncRegistryToFile
End Sub

' Best-effort sync ws_UsersRegistry -> BNC_UsersRegistry.xlsx.
' Bez clipboard (ADR-002). Wywolywany po kazdej mutacji Registry.
Private Sub SyncRegistryToFile()
    Dim wbOut As Workbook
    Dim restoreScreen As Boolean
    Dim restoreAlerts As Boolean
    On Error GoTo Cleanup

    Dim folderPath As String
    folderPath = CStr(GetCurrentUserField("CacheFolderPath"))
    If Len(folderPath) = 0 Then Exit Sub

    mod_Utils.EnsureFolderExists folderPath

    Dim fullPath As String
    fullPath = mod_Utils.JoinPath(folderPath, REGISTRY_CACHE_FILE_NAME)

    Dim srcWs As Worksheet
    On Error Resume Next
    Set srcWs = ThisWorkbook.Worksheets(REGISTRY_SHEET)
    On Error GoTo Cleanup
    If srcWs Is Nothing Then Exit Sub

    Application.ScreenUpdating = False
    restoreScreen = True
    Application.DisplayAlerts = False
    restoreAlerts = True

    Set wbOut = Workbooks.Add

    Dim destWs As Worksheet
    Set destWs = wbOut.Worksheets(1)
    destWs.Name = REGISTRY_SHEET

    Dim usedRange As Range
    Set usedRange = srcWs.UsedRange
    If usedRange.Cells.Count > 0 Then
        destWs.Range( _
            destWs.Cells(1, 1), _
            destWs.Cells(usedRange.Rows.Count, usedRange.Columns.Count) _
        ).Value = usedRange.Value
    End If

    If mod_Utils.FileExists(fullPath) Then Kill fullPath
    wbOut.SaveAs Filename:=fullPath, FileFormat:=xlOpenXMLWorkbook
    wbOut.Close SaveChanges:=False
    Set wbOut = Nothing

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

Cleanup:
    mod_Utils.LogError "mod_UsersRegistrySync.SyncRegistryToFile", Err.Number, Err.Description
    On Error Resume Next
    If Not wbOut Is Nothing Then wbOut.Close SaveChanges:=False
    If restoreAlerts Then Application.DisplayAlerts = True
    If restoreScreen Then Application.ScreenUpdating = True
End Sub

' Bezpieczny dostep do pol Dictionary - "" jesli klucz nie istnieje.
Private Function SafeGet(d As Object, key As String) As Variant
    If d Is Nothing Then
        SafeGet = ""
    ElseIf d.Exists(key) Then
        SafeGet = d(key)
    Else
        SafeGet = ""
    End If
End Function
