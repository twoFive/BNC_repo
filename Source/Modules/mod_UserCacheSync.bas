Attribute VB_Name = "mod_UserCacheSync"
Option Explicit

' ============================================================================
'  mod_UserCacheSync - Repository Pattern dla ws_UserCache.
'  Jedyne miejsce dostepu do ukrytego arkusza ws_UserCache. Synchronizuje
'  stan do BNC_UserCache.xlsx (write-through cache, jednostronny sync
'  worksheet -> xlsx, best-effort).
'
'  Format ws_UserCache: key-value (kolumna A = klucz, kolumna B = wartosc).
'  Detekcja roli: IsUserManager() porownuje EmailKierownika z EmailHandlowca.
'  Patrz: BNC_Sender_PlanWdrozenia_FazaA.md (M1.2)
'         doc_v2/extracted/04_data_model.md
' ============================================================================

Private Const SHEET_NAME As String = "ws_UserCache"
Private Const CACHE_FILE_NAME As String = "BNC_UserCache.xlsx"

' Multi-user Registry (M3.3) - tabelaryczny arkusz z lista wszystkich userow.
' UserCache reprezentuje AKTYWNEGO usera, Registry - PELNA liste.
Private Const REGISTRY_SHEET As String = "ws_UsersRegistry"
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
Private Const CURRENT_USER_KEY As String = "_CurrentUserID"

' Kanoniczna lista pol UserCache (uzywana przez GetUserData do pelnego dumpa).
Private Function GetFieldKeys() As Variant
    GetFieldKeys = Array( _
        "Imie", _
        "Nazwisko", _
        "EmailHandlowca", _
        "CNA_HandlowcaID", _
        "NrOddzialu", _
        "EmailKierownika", _
        "EmailBNC", _
        "CacheFolderPath", _
        "DataRejestracji", _
        "SetupCompleted", _
        "DontShowSetupAgain" _
    )
End Function

' ============================================================================
'  Public API
' ============================================================================

' Czyta wartosc dla podanego klucza. Zwraca "" jesli klucza nie ma.
Public Function GetUserField(fieldKey As String) As Variant
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)

    Dim r As Long
    r = FindKeyRow(ws, fieldKey)
    If r = 0 Then
        GetUserField = ""
        Exit Function
    End If
    GetUserField = ws.Cells(r, 2).Value
End Function

' Zwraca Scripting.Dictionary z wszystkimi kanonicznymi polami.
Public Function GetUserData() As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")

    Dim keys As Variant
    keys = GetFieldKeys()

    Dim i As Long
    For i = LBound(keys) To UBound(keys)
        result(CStr(keys(i))) = GetUserField(CStr(keys(i)))
    Next i

    Set GetUserData = result
End Function

' Zapisuje pojedyncza wartosc + ThisWorkbook.Save + SyncToFile.
Public Sub SetUserField(fieldKey As String, value As Variant)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)

    Dim r As Long
    r = FindKeyRow(ws, fieldKey)

    If r = 0 Then r = NextEmptyRow(ws)
    ws.Cells(r, 1).Value = fieldKey
    ws.Cells(r, 2).Value = value

    ThisWorkbook.Save
    SyncToFile
End Sub

' Pisze caly zestaw pol (np. po setup). Wymusza SetupCompleted=True.
Public Sub SaveUserData(userData As Object)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)

    If Not userData.Exists("SetupCompleted") Then userData("SetupCompleted") = True

    Dim k As Variant
    For Each k In userData.Keys
        Dim r As Long
        r = FindKeyRow(ws, CStr(k))
        If r = 0 Then r = NextEmptyRow(ws)
        ws.Cells(r, 1).Value = k
        ws.Cells(r, 2).Value = userData(k)
    Next k

    ThisWorkbook.Save
    SyncToFile
End Sub

Public Function IsSetupCompleted() As Boolean
    Dim v As Variant
    v = GetUserField("SetupCompleted")

    If IsEmpty(v) Or VarType(v) = vbNull Then
        IsSetupCompleted = False
    ElseIf VarType(v) = vbBoolean Then
        IsSetupCompleted = CBool(v)
    ElseIf VarType(v) = vbString Then
        Dim s As String
        s = LCase$(Trim$(CStr(v)))
        IsSetupCompleted = (s = "true" Or s = "1" Or s = "prawda")
    Else
        ' Numeryczne 0/1
        On Error Resume Next
        IsSetupCompleted = (CDbl(v) <> 0)
        On Error GoTo 0
    End If
End Function

' Convention over configuration: jezeli EmailKierownika == EmailHandlowca,
' user jest kierownikiem (sam siebie wpisal jako kierownika).
Public Function IsUserManager() As Boolean
    Dim handlowca As String, kierownika As String
    handlowca = LCase$(Trim$(CStr(GetUserField("EmailHandlowca"))))
    kierownika = LCase$(Trim$(CStr(GetUserField("EmailKierownika"))))

    If Len(handlowca) = 0 Or Len(kierownika) = 0 Then
        IsUserManager = False
        Exit Function
    End If

    IsUserManager = (handlowca = kierownika)
End Function

' Auto-recreate: jesli BNC_UserCache.xlsx nie istnieje przy starcie aplikacji,
' tworzy go z aktualnej zawartosci ws_UserCache.
Public Sub EnsureCacheFileExists()
    Dim folderPath As String
    folderPath = CStr(GetUserField("CacheFolderPath"))
    If Len(folderPath) = 0 Then Exit Sub  ' setup jeszcze nieukonczony

    mod_Utils.EnsureFolderExists folderPath

    Dim fullPath As String
    fullPath = mod_Utils.JoinPath(folderPath, CACHE_FILE_NAME)

    If Not mod_Utils.FileExists(fullPath) Then SyncToFile
End Sub

' ============================================================================
'  Multi-user API (M3.3)
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
Public Function CurrentUserID() As String
    CurrentUserID = CStr(GetUserField(CURRENT_USER_KEY))
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

    ' 3. Ustaw UserID jako aktualny
    SetUserField CURRENT_USER_KEY, userId

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

    ' 2. Skopiuj do UserCache jako aktywnego
    SaveUserData userData

    ' 3. Ustaw jako aktualny
    SetUserField CURRENT_USER_KEY, newUserId

    ' 4. LastLogin = Now()
    UpdateLastLoginInRegistry newUserId

    mod_Utils.LogInfo "AddNewUser: dodano " & newUserId
    AddNewUser = newUserId
End Function

' Czysci UserCache aby przygotowac frm_Setup do dodania nowego usera.
' Wywolywane z frm_UserPicker.btn_AddNew_Click przed frm_Setup.Show.
Public Sub PrepareForNewUser()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)

    ' Wyczysc wszystkie pola UserCache (nie kasujemy arkusza, tylko wartosci)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    If lastRow >= 1 Then
        ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, 2)).ClearContents
    End If

    mod_Utils.LogInfo "PrepareForNewUser: UserCache wyczyszczony dla nowego usera"
End Sub

' ============================================================================
'  Private
' ============================================================================

Private Function FindKeyRow(ws As Worksheet, fieldKey As String) As Long
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

    Dim r As Long
    For r = 1 To lastRow
        If StrComp(CStr(ws.Cells(r, 1).Value), fieldKey, vbTextCompare) = 0 Then
            FindKeyRow = r
            Exit Function
        End If
    Next r
    FindKeyRow = 0
End Function

' Pierwszy wolny wiersz z perspektywy klucza (kolumna A).
Private Function NextEmptyRow(ws As Worksheet) As Long
    If Len(CStr(ws.Cells(1, 1).Value)) = 0 Then
        NextEmptyRow = 1
    Else
        NextEmptyRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row + 1
    End If
End Function

' Best-effort sync ws_UserCache -> BNC_UserCache.xlsx.
' Bez clipboard (Range.Value = Range.Value, brak race condition ze schowkiem).
Private Sub SyncToFile()
    Dim wbOut As Workbook
    Dim restoreScreen As Boolean
    Dim restoreAlerts As Boolean
    On Error GoTo Cleanup

    Dim folderPath As String
    folderPath = CStr(GetUserField("CacheFolderPath"))
    If Len(folderPath) = 0 Then Exit Sub

    mod_Utils.EnsureFolderExists folderPath

    Dim fullPath As String
    fullPath = mod_Utils.JoinPath(folderPath, CACHE_FILE_NAME)

    Dim srcWs As Worksheet
    Set srcWs = ThisWorkbook.Worksheets(SHEET_NAME)

    Application.ScreenUpdating = False
    restoreScreen = True
    Application.DisplayAlerts = False
    restoreAlerts = True

    Set wbOut = Workbooks.Add

    Dim destWs As Worksheet
    Set destWs = wbOut.Worksheets(1)
    destWs.Name = SHEET_NAME

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
    mod_Utils.LogError "mod_UserCacheSync.SyncToFile", Err.Number, Err.Description
    On Error Resume Next
    If Not wbOut Is Nothing Then wbOut.Close SaveChanges:=False
    If restoreAlerts Then Application.DisplayAlerts = True
    If restoreScreen Then Application.ScreenUpdating = True
End Sub

' ============================================================================
'  Private - Registry helpers (M3.3)
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
    mod_Utils.LogError "mod_UserCacheSync.EnsureRegistrySheet", Err.Number, Err.Description
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
End Sub

' Kopiuje pola z wiersza Registry do ws_UserCache (key-value).
Private Sub LoadUserFromRegistry(userId As String)
    Dim wsReg As Worksheet
    Set wsReg = EnsureRegistrySheet()
    If wsReg Is Nothing Then Exit Sub

    Dim r As Long
    r = FindRegistryRow(wsReg, userId)
    If r = 0 Then
        mod_Utils.LogError "mod_UserCacheSync.LoadUserFromRegistry", 0, _
            "UserID nie znaleziony w Registry: " & userId
        Exit Sub
    End If

    ' Wyczysc UserCache przed loadem
    Dim wsCache As Worksheet
    Set wsCache = ThisWorkbook.Worksheets(SHEET_NAME)
    Dim lastRow As Long
    lastRow = wsCache.Cells(wsCache.Rows.Count, 1).End(xlUp).row
    If lastRow >= 1 Then wsCache.Range(wsCache.Cells(1, 1), wsCache.Cells(lastRow, 2)).ClearContents

    ' Wpisz kazde pole jako klucz-wartosc
    SetUserFieldRaw wsCache, "Imie", wsReg.Cells(r, REG_IMIE).Value
    SetUserFieldRaw wsCache, "Nazwisko", wsReg.Cells(r, REG_NAZWISKO).Value
    SetUserFieldRaw wsCache, "EmailHandlowca", wsReg.Cells(r, REG_EMAIL_HANDLOWCA).Value
    SetUserFieldRaw wsCache, "CNA_HandlowcaID", wsReg.Cells(r, REG_CNA).Value
    SetUserFieldRaw wsCache, "NrOddzialu", wsReg.Cells(r, REG_NR_ODDZIALU).Value
    SetUserFieldRaw wsCache, "EmailKierownika", wsReg.Cells(r, REG_EMAIL_KIEROWNIKA).Value
    SetUserFieldRaw wsCache, "EmailBNC", wsReg.Cells(r, REG_EMAIL_BNC).Value
    SetUserFieldRaw wsCache, "CacheFolderPath", wsReg.Cells(r, REG_CACHE_FOLDER).Value
    SetUserFieldRaw wsCache, "DataRejestracji", wsReg.Cells(r, REG_DATA_REJESTRACJI).Value
    SetUserFieldRaw wsCache, "SetupCompleted", wsReg.Cells(r, REG_SETUP_COMPLETED).Value
    SetUserFieldRaw wsCache, "DontShowSetupAgain", wsReg.Cells(r, REG_DONT_SHOW_SETUP).Value

    ThisWorkbook.Save
    SyncToFile
End Sub

' Kopiuje pola z ws_UserCache do wiersza Registry dla podanego userId.
Private Sub SaveCurrentUserToRegistry(userId As String)
    Dim wsReg As Worksheet
    Set wsReg = EnsureRegistrySheet()
    If wsReg Is Nothing Then Exit Sub

    Dim r As Long
    r = FindRegistryRow(wsReg, userId)
    If r = 0 Then Exit Sub  ' user nie zarejestrowany - nic nie zapisujemy

    wsReg.Cells(r, REG_IMIE).Value = GetUserField("Imie")
    wsReg.Cells(r, REG_NAZWISKO).Value = GetUserField("Nazwisko")
    wsReg.Cells(r, REG_EMAIL_HANDLOWCA).Value = GetUserField("EmailHandlowca")
    wsReg.Cells(r, REG_CNA).Value = GetUserField("CNA_HandlowcaID")
    wsReg.Cells(r, REG_NR_ODDZIALU).Value = GetUserField("NrOddzialu")
    wsReg.Cells(r, REG_EMAIL_KIEROWNIKA).Value = GetUserField("EmailKierownika")
    wsReg.Cells(r, REG_EMAIL_BNC).Value = GetUserField("EmailBNC")
    wsReg.Cells(r, REG_CACHE_FOLDER).Value = GetUserField("CacheFolderPath")
    wsReg.Cells(r, REG_SETUP_COMPLETED).Value = GetUserField("SetupCompleted")
    wsReg.Cells(r, REG_DONT_SHOW_SETUP).Value = GetUserField("DontShowSetupAgain")
    ' DataRejestracji, LastLogin - nie ruszamy
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
End Sub

' Set key-value pair na konkretnym arkuszu bez auto-save. Uzywany w
' LoadUserFromRegistry gdzie zbieramy wiele setow w batch.
Private Sub SetUserFieldRaw(ws As Worksheet, fieldKey As String, value As Variant)
    Dim r As Long
    r = FindKeyRow(ws, fieldKey)
    If r = 0 Then r = NextEmptyRow(ws)
    ws.Cells(r, 1).Value = fieldKey
    ws.Cells(r, 2).Value = value
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
