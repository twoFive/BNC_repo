Attribute VB_Name = "mod_UsersRegistrySync"
Option Explicit

' ============================================================================
'  mod_UsersRegistrySync - Repository Pattern dla ws_UsersRegistry.
'  Wyekstrahowany z mod_UserCacheSync (2026-07-26 refactor) dla symetrii:
'  trzy warstwy cache = trzy moduly *Sync (ADR-001).
'
'  Registry = tabelaryczny arkusz z lista wszystkich zarejestrowanych userow
'  (1 wiersz = 1 user, 13 kolumn). UserCache reprezentuje AKTYWNEGO usera,
'  Registry - PELNA liste. Synchronizuje stan do BNC_UsersRegistry.xlsx
'  (write-through, jednostronny sync ws -> xlsx, best-effort).
'
'  Cross-module dependency: SwitchUser + AddNewUser + LoadUserFromRegistry
'  wolaja mod_UserCacheSync do zapisu/odczytu UserCache aktywnego usera.
'  UserCache NIE ma zaleznosci zwrotnej - Registry zawsze wie o UserCache,
'  UserCache nic nie wie o Registry.
'
'  Format UserID: UZYTKOWNIK_<autoinc>_CNA<cna> (ADR-008, Q2 decyzja).
'  Patrz: DECISIONS.md (ADR-001, ADR-008)
'         doc_v2/diagrams/04_data_model.md (sekcja 3)
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

' Klucz w UserCache przechowujacy UserID aktualnie zalogowanego usera.
' Konwencja: podkreslnik-prefix = pole systemowe (nie kanoniczne UserCache).
Private Const CURRENT_USER_KEY As String = "_CurrentUserID"

' ============================================================================
'  Public API
' ============================================================================

' Liczba zarejestrowanych userow w ws_UsersRegistry. 0 = pierwsze uruchomienie.
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

' Zwraca UserID aktualnie zalogowanego usera. Pusty gdy nikt nie wybrany.
' Cross-module read: _CurrentUserID mieszka w UserCache (semantycznie Registry-owned).
Public Function CurrentUserID() As String
    CurrentUserID = CStr(mod_UserCacheSync.GetUserField(CURRENT_USER_KEY))
End Function

' Wszystkie zarejestrowani userzy jako Collection of Scripting.Dictionary.
' Kazdy Dict zawiera pola: UserID, Imie, Nazwisko, EmailHandlowca,
' CNA_HandlowcaID, NrOddzialu, EmailKierownika, EmailBNC, CacheFolderPath,
' DataRejestracji, SetupCompleted, DontShowSetupAgain, LastLogin.
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
        GetAllUsers.Add d
    Next r
End Function

' Przelacza aktywnego usera: zapisuje aktualny stan UserCache do Registry
' dla poprzedniego usera, ladowuje nowego z Registry do UserCache, aktualizuje
' LastLogin. Wywolywane z frm_UserPicker.
Public Sub SwitchUser(userId As String)
    ' 1. Zapisz aktualny stan aktywnego usera z UserCache do Registry
    Dim previousUserId As String
    previousUserId = CurrentUserID()
    If Len(previousUserId) > 0 Then SaveCurrentUserToRegistry previousUserId

    ' 2. Zaladuj nowego usera z Registry do UserCache
    LoadUserFromRegistry userId

    ' 3. Ustaw UserID jako aktualny (w UserCache jako _CurrentUserID)
    mod_UserCacheSync.SetUserField CURRENT_USER_KEY, userId

    ' 4. Zaktualizuj LastLogin w Registry
    UpdateLastLoginInRegistry userId

    mod_Utils.LogInfo "SwitchUser: aktywny user = " & userId
End Sub

' Dodaje nowego usera do Registry i przelacza na niego. Wywolywane z
' frm_Setup.btn_Save gdy pierwszy uzytkownik (Registry pusty) lub gdy
' user przyszedl z frm_UserPicker.btn_AddNew.
' Format UserID: UZYTKOWNIK_<autoinc>_CNA<cna>  (M3.3 Q2 decyzja)
' Returns: wygenerowany UserID.
Public Function AddNewUser(userData As Object) As String
    Dim ws As Worksheet
    Set ws = EnsureRegistrySheet()

    Dim newUserId As String
    newUserId = GenerateUserID(userData)
    userData("UserID") = newUserId

    ' 1. Zapisz nowego usera do Registry
    AppendUserToRegistry userData

    ' 2. Skopiuj do UserCache jako aktywnego (cross-module)
    mod_UserCacheSync.SaveUserData userData

    ' 3. Ustaw jako aktualny
    mod_UserCacheSync.SetUserField CURRENT_USER_KEY, newUserId

    ' 4. LastLogin = Now()
    UpdateLastLoginInRegistry newUserId

    mod_Utils.LogInfo "AddNewUser: dodano " & newUserId
    AddNewUser = newUserId
End Function

' Czysci UserCache aby przygotowac frm_Setup do dodania nowego usera.
' Wywolywane z frm_UserPicker.btn_AddNew_Click przed frm_Setup.Show.
' Cross-module: deleguje do UserCache-owner.
Public Sub PrepareForNewUser()
    mod_UserCacheSync.ClearUserCache
    mod_Utils.LogInfo "PrepareForNewUser: UserCache wyczyszczony dla nowego usera"
End Sub

' Auto-recreate: jesli BNC_UsersRegistry.xlsx nie istnieje przy starcie
' aplikacji, tworzy go z aktualnej zawartosci ws_UsersRegistry.
' Symetria z mod_UserCacheSync.EnsureCacheFileExists i
' mod_DataCacheSync.EnsureCacheFileExists.
Public Sub EnsureRegistryCacheFileExists()
    Dim folderPath As String
    folderPath = CStr(mod_UserCacheSync.GetUserField("CacheFolderPath"))
    If Len(folderPath) = 0 Then Exit Sub  ' setup jeszcze nieukonczony

    mod_Utils.EnsureFolderExists folderPath

    Dim fullPath As String
    fullPath = mod_Utils.JoinPath(folderPath, REGISTRY_CACHE_FILE_NAME)

    If Not mod_Utils.FileExists(fullPath) Then SyncRegistryToFile
End Sub

' ============================================================================
'  Private - Registry helpers
' ============================================================================

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
' Q2 decyzja: identifier human-readable, latwy do debugu w Immediate Window.
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

' Dopisuje nowego usera do Registry.
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

' Kopiuje pola z wiersza Registry do UserCache (cross-module - wola
' mod_UserCacheSync.SaveUserData ktora sama robi ClearContents + write + sync).
Private Sub LoadUserFromRegistry(userId As String)
    Dim wsReg As Worksheet
    Set wsReg = EnsureRegistrySheet()
    If wsReg Is Nothing Then Exit Sub

    Dim r As Long
    r = FindRegistryRow(wsReg, userId)
    If r = 0 Then
        mod_Utils.LogError "mod_UsersRegistrySync.LoadUserFromRegistry", 0, _
            "UserID nie znaleziony w Registry: " & userId
        Exit Sub
    End If

    ' Wyczysc UserCache przed loadem (cross-module)
    mod_UserCacheSync.ClearUserCache

    ' Zbuduj Dictionary z wiersza Registry
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("Imie") = wsReg.Cells(r, REG_IMIE).Value
    d("Nazwisko") = wsReg.Cells(r, REG_NAZWISKO).Value
    d("EmailHandlowca") = wsReg.Cells(r, REG_EMAIL_HANDLOWCA).Value
    d("CNA_HandlowcaID") = wsReg.Cells(r, REG_CNA).Value
    d("NrOddzialu") = wsReg.Cells(r, REG_NR_ODDZIALU).Value
    d("EmailKierownika") = wsReg.Cells(r, REG_EMAIL_KIEROWNIKA).Value
    d("EmailBNC") = wsReg.Cells(r, REG_EMAIL_BNC).Value
    d("CacheFolderPath") = wsReg.Cells(r, REG_CACHE_FOLDER).Value
    d("DataRejestracji") = wsReg.Cells(r, REG_DATA_REJESTRACJI).Value
    d("SetupCompleted") = wsReg.Cells(r, REG_SETUP_COMPLETED).Value
    d("DontShowSetupAgain") = wsReg.Cells(r, REG_DONT_SHOW_SETUP).Value

    ' Zapisz do UserCache (cross-module - SaveUserData robi save + sync)
    mod_UserCacheSync.SaveUserData d
End Sub

' Kopiuje pola z UserCache aktywnego usera do wiersza Registry dla podanego userId.
' Cross-module: czyta UserCache przez public API.
Private Sub SaveCurrentUserToRegistry(userId As String)
    Dim wsReg As Worksheet
    Set wsReg = EnsureRegistrySheet()
    If wsReg Is Nothing Then Exit Sub

    Dim r As Long
    r = FindRegistryRow(wsReg, userId)
    If r = 0 Then Exit Sub  ' user nie zarejestrowany - nic nie zapisujemy

    wsReg.Cells(r, REG_IMIE).Value = mod_UserCacheSync.GetUserField("Imie")
    wsReg.Cells(r, REG_NAZWISKO).Value = mod_UserCacheSync.GetUserField("Nazwisko")
    wsReg.Cells(r, REG_EMAIL_HANDLOWCA).Value = mod_UserCacheSync.GetUserField("EmailHandlowca")
    wsReg.Cells(r, REG_CNA).Value = mod_UserCacheSync.GetUserField("CNA_HandlowcaID")
    wsReg.Cells(r, REG_NR_ODDZIALU).Value = mod_UserCacheSync.GetUserField("NrOddzialu")
    wsReg.Cells(r, REG_EMAIL_KIEROWNIKA).Value = mod_UserCacheSync.GetUserField("EmailKierownika")
    wsReg.Cells(r, REG_EMAIL_BNC).Value = mod_UserCacheSync.GetUserField("EmailBNC")
    wsReg.Cells(r, REG_CACHE_FOLDER).Value = mod_UserCacheSync.GetUserField("CacheFolderPath")
    wsReg.Cells(r, REG_SETUP_COMPLETED).Value = mod_UserCacheSync.GetUserField("SetupCompleted")
    wsReg.Cells(r, REG_DONT_SHOW_SETUP).Value = mod_UserCacheSync.GetUserField("DontShowSetupAgain")
    ' DataRejestracji, LastLogin - nie ruszamy (LastLogin update robi
    ' UpdateLastLoginInRegistry, ktore tez robi Save + Sync)
End Sub

' Update LastLogin dla usera w Registry.
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
' Bez clipboard (ADR-002). Wywolywany po kazdej mutacji Registry:
' AppendUserToRegistry, UpdateLastLoginInRegistry. Bledy tylko do logu.
Private Sub SyncRegistryToFile()
    Dim wbOut As Workbook
    Dim restoreScreen As Boolean
    Dim restoreAlerts As Boolean
    On Error GoTo Cleanup

    Dim folderPath As String
    folderPath = CStr(mod_UserCacheSync.GetUserField("CacheFolderPath"))
    If Len(folderPath) = 0 Then Exit Sub

    mod_Utils.EnsureFolderExists folderPath

    Dim fullPath As String
    fullPath = mod_Utils.JoinPath(folderPath, REGISTRY_CACHE_FILE_NAME)

    Dim srcWs As Worksheet
    On Error Resume Next
    Set srcWs = ThisWorkbook.Worksheets(REGISTRY_SHEET)
    On Error GoTo Cleanup
    If srcWs Is Nothing Then Exit Sub  ' Registry jeszcze nie utworzony

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
