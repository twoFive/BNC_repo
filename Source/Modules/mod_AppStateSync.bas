Attribute VB_Name = "mod_AppStateSync"
Option Explicit

' ============================================================================
'  mod_AppStateSync
'  Repository Pattern dla ws_AppState - key-value store dla session state
'  aplikacji (ADR-009). NIE trzyma user data, tylko meta-stan.
'
'  Obecne klucze: _CurrentUserID (UserID aktywnego usera).
'  Brak xlsx sync - marker sesji nie jest krytyczny, w razie utraty user
'  przechodzi przez picker.
' ============================================================================

Private Const SHEET_NAME As String = "ws_AppState"

' ============================================================================
'  Public API
' ============================================================================

' Czyta wartoœæ dla podanego klucza. Zwraca "" jeœli klucza nie ma.
Public Function GetAppValue(fieldKey As String) As Variant
    Dim ws As Worksheet
    Set ws = EnsureAppStateSheet()
    If ws Is Nothing Then
        GetAppValue = ""
        Exit Function
    End If

    Dim r As Long
    r = FindKeyRow(ws, fieldKey)
    If r = 0 Then
        GetAppValue = ""
        Exit Function
    End If
    GetAppValue = ws.Cells(r, 2).Value
End Function

' Zapisuje wartoœæ dla podanego klucza (create-or-update) + ThisWorkbook.Save.
' Bez xlsx sync (per ADR-009 - AppState nie ma persistence layer).
Public Sub SetAppValue(fieldKey As String, value As Variant)
    Dim ws As Worksheet
    Set ws = EnsureAppStateSheet()
    If ws Is Nothing Then Exit Sub

    Dim r As Long
    r = FindKeyRow(ws, fieldKey)

    If r = 0 Then r = NextEmptyRow(ws)
    ws.Cells(r, 1).Value = fieldKey
    ws.Cells(r, 2).Value = value

    ThisWorkbook.Save
End Sub

' Zwraca ws_AppState, tworz¹c go jeœli nie istnieje. Sheet very hidden.
' Idempotentne. Wywo³ywane z Workbook_Open dla auto-create przy pierwszym
' uruchomieniu nowej wersji xlsm.
Public Function EnsureAppStateSheet() As Worksheet
    On Error Resume Next
    Set EnsureAppStateSheet = ThisWorkbook.Worksheets(SHEET_NAME)
    On Error GoTo 0

    If EnsureAppStateSheet Is Nothing Then
        Application.ScreenUpdating = False
        On Error GoTo CreateError
        Set EnsureAppStateSheet = ThisWorkbook.Worksheets.Add
        EnsureAppStateSheet.Name = SHEET_NAME
        EnsureAppStateSheet.Visible = xlSheetVeryHidden
        On Error GoTo 0
        Application.ScreenUpdating = True
        mod_Utils.LogInfo "EnsureAppStateSheet: utworzono " & SHEET_NAME
    End If
    Exit Function

CreateError:
    mod_Utils.LogError "mod_AppStateSync.EnsureAppStateSheet", Err.Number, Err.Description
    Application.ScreenUpdating = True
    Set EnsureAppStateSheet = Nothing
End Function

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

Private Function NextEmptyRow(ws As Worksheet) As Long
    If Len(CStr(ws.Cells(1, 1).Value)) = 0 Then
        NextEmptyRow = 1
    Else
        NextEmptyRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row + 1
    End If
End Function
