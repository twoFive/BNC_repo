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
        "DontShowTutorial" _
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
