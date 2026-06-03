Attribute VB_Name = "mod_DataCacheSync"
Option Explicit

' ============================================================================
'  mod_DataCacheSync - Repository Pattern dla ws_DataCache.
'  Tabela zgloszen BNC: kazdy wiersz = jedno zgloszenie. ReportID jest
'  autoincrement (logika w VBA, Excel nie ma natywnego autoincrement).
'  Pola CNA_HandlowcaID i NrOddzialu sa snapshotami z UserCache w momencie
'  zapisu - chronia historie przed pozniejszymi zmianami w setupie.
'
'  Synchronizuje stan do BNC_DataCache.xlsx (write-through, jednostronny).
'  Patrz: BNC_Sender_PlanWdrozenia_FazaA.md (M1.3)
'         doc_v2/extracted/04_data_model.md
' ============================================================================

Private Const SHEET_NAME As String = "ws_DataCache"
Private Const CACHE_FILE_NAME As String = "BNC_DataCache.xlsx"

' Schemat kolumn (1-indexed). Naglowki w wierszu 1, dane od wiersza 2.
Private Const COL_REPORT_ID As Long = 1
Private Const COL_KLIENT_FK As Long = 2
Private Const COL_NAZWA_KLIENTA As Long = 3
Private Const COL_MIESIAC As Long = 4
Private Const COL_FIELDS As Long = 5
Private Const COL_CNA As Long = 6
Private Const COL_NR_ODDZIALU As Long = 7
Private Const COL_CREATED As Long = 8
Private Const COL_STATUS As Long = 9
Private Const COL_RECIPIENT As Long = 10
Private Const COL_BATCH_SENT As Long = 11
Private Const TOTAL_COLS As Long = 11

Public Const STATUS_PENDING As String = "pending"
Public Const STATUS_SENT As String = "sent"

' ============================================================================
'  Public API
' ============================================================================

' Dodaje nowe zgloszenie ze Status=pending. Aplikacja sama uzupelnia:
'   ReportID, CNA_HandlowcaID, NrOddzialu (snapshot), CreatedTimestamp,
'   Status, EmailRecipient="", BatchSentTimestamp="".
' Zwraca: ReportID nowego rekordu.
Public Function AppendRecord(reportData As Object) As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)

    EnsureHeader ws

    Dim newID As Long
    newID = GetNextReportID(ws)

    Dim r As Long
    r = ws.Cells(ws.Rows.Count, COL_REPORT_ID).End(xlUp).row + 1
    If r < 2 Then r = 2  ' nigdy nie nadpisuj naglowka

    ws.Cells(r, COL_REPORT_ID).Value = newID
    ws.Cells(r, COL_KLIENT_FK).Value = SafeGet(reportData, "KlientFK")
    ws.Cells(r, COL_NAZWA_KLIENTA).Value = SafeGet(reportData, "NazwaKlienta")
    ws.Cells(r, COL_MIESIAC).Value = SafeGet(reportData, "MiesiacZgloszenia")
    ws.Cells(r, COL_FIELDS).Value = SafeGet(reportData, "Fields")

    ' Snapshot z UserCache - chroni historie przed zmianami w setupie.
    ws.Cells(r, COL_CNA).Value = mod_UserCacheSync.GetUserField("CNA_HandlowcaID")
    ws.Cells(r, COL_NR_ODDZIALU).Value = mod_UserCacheSync.GetUserField("NrOddzialu")

    ws.Cells(r, COL_CREATED).Value = Now()
    ws.Cells(r, COL_STATUS).Value = STATUS_PENDING
    ws.Cells(r, COL_RECIPIENT).Value = ""
    ws.Cells(r, COL_BATCH_SENT).Value = ""

    ThisWorkbook.Save
    SyncToFile

    AppendRecord = newID
End Function

' Wszystkie wiersze ze Status=pending. Collection of Scripting.Dictionary.
Public Function GetPendingRecords() As Collection
    Set GetPendingRecords = GetRecordsWhereStatus(STATUS_PENDING)
End Function

' Wszystkie wiersze (do log_UserForm). Collection of Scripting.Dictionary.
Public Function GetAllRecords() As Collection
    Set GetAllRecords = GetRecordsWhereStatus("")  ' "" = bez filtra
End Function

' UPDATE Status=sent + EmailRecipient + BatchSentTimestamp dla podanych ReportID.
Public Sub MarkAsSent(reportIDs As Collection, recipient As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_REPORT_ID).End(xlUp).row
    If lastRow < 2 Then Exit Sub

    ' Zbuduj set ID-ow dla O(1) lookup.
    Dim sentSet As Object
    Set sentSet = CreateObject("Scripting.Dictionary")
    Dim id As Variant
    For Each id In reportIDs
        sentSet(CStr(id)) = True
    Next id

    Dim sentTime As Date
    sentTime = Now()

    Dim r As Long
    For r = 2 To lastRow
        If sentSet.Exists(CStr(ws.Cells(r, COL_REPORT_ID).Value)) Then
            ws.Cells(r, COL_STATUS).Value = STATUS_SENT
            ws.Cells(r, COL_RECIPIENT).Value = recipient
            ws.Cells(r, COL_BATCH_SENT).Value = sentTime
        End If
    Next r

    ThisWorkbook.Save
    SyncToFile
End Sub

' Hard delete pojedynczego pending recordu (ADR-006).
' Defensywnie: tylko Status=pending. Sent records sa immutable (audit trail).
' Returns True jesli usunieto, False jesli ID nieznane LUB Status != pending.
Public Function DeleteRecord(reportID As Long) As Boolean
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_REPORT_ID).End(xlUp).row
    If lastRow < 2 Then
        DeleteRecord = False
        Exit Function
    End If

    Dim r As Long
    For r = 2 To lastRow
        If IsNumeric(ws.Cells(r, COL_REPORT_ID).Value) Then
            If CLng(ws.Cells(r, COL_REPORT_ID).Value) = reportID Then
                ' Defensywny check: tylko pending mozna usuwac (ADR-006).
                Dim status As String
                status = CStr(ws.Cells(r, COL_STATUS).Value)
                If status <> STATUS_PENDING Then
                    mod_Utils.LogError "mod_DataCacheSync.DeleteRecord", 0, _
                        "Odmowa - status='" & status & "', tylko pending mozna usunac. ID=" & reportID
                    DeleteRecord = False
                    Exit Function
                End If

                ws.Rows(r).Delete
                ThisWorkbook.Save
                SyncToFile
                DeleteRecord = True
                Exit Function
            End If
        End If
    Next r

    ' Nie znaleziono
    DeleteRecord = False
End Function

' Auto-recreate: tworzy BNC_DataCache.xlsx jesli nie istnieje (Workbook_Open).
Public Sub EnsureCacheFileExists()
    Dim folderPath As String
    folderPath = CStr(mod_UserCacheSync.GetUserField("CacheFolderPath"))
    If Len(folderPath) = 0 Then Exit Sub

    mod_Utils.EnsureFolderExists folderPath

    Dim fullPath As String
    fullPath = mod_Utils.JoinPath(folderPath, CACHE_FILE_NAME)

    If Not mod_Utils.FileExists(fullPath) Then
        EnsureHeader ThisWorkbook.Worksheets(SHEET_NAME)
        SyncToFile
    End If
End Sub

' ============================================================================
'  Private
' ============================================================================

Private Function GetNextReportID(ws As Worksheet) As Long
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_REPORT_ID).End(xlUp).row

    If lastRow < 2 Then
        GetNextReportID = 1
        Exit Function
    End If

    Dim lastID As Variant
    lastID = ws.Cells(lastRow, COL_REPORT_ID).Value
    If IsNumeric(lastID) Then
        GetNextReportID = CLng(lastID) + 1
    Else
        GetNextReportID = 1
    End If
End Function

' Tworzy naglowki tylko jesli wiersz 1 jest pusty - idempotentne.
Private Sub EnsureHeader(ws As Worksheet)
    If Len(CStr(ws.Cells(1, COL_REPORT_ID).Value)) > 0 Then Exit Sub

    ws.Cells(1, COL_REPORT_ID).Value = "ReportID"
    ws.Cells(1, COL_KLIENT_FK).Value = "KlientFK"
    ws.Cells(1, COL_NAZWA_KLIENTA).Value = "NazwaKlienta"
    ws.Cells(1, COL_MIESIAC).Value = "MiesiacZgloszenia"
    ws.Cells(1, COL_FIELDS).Value = "Fields"
    ws.Cells(1, COL_CNA).Value = "CNA_HandlowcaID"
    ws.Cells(1, COL_NR_ODDZIALU).Value = "NrOddzialu"
    ws.Cells(1, COL_CREATED).Value = "CreatedTimestamp"
    ws.Cells(1, COL_STATUS).Value = "Status"
    ws.Cells(1, COL_RECIPIENT).Value = "EmailRecipient"
    ws.Cells(1, COL_BATCH_SENT).Value = "BatchSentTimestamp"
End Sub

' statusFilter="" - bez filtra (wszystkie wiersze).
Private Function GetRecordsWhereStatus(statusFilter As String) As Collection
    Dim result As New Collection

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_REPORT_ID).End(xlUp).row
    If lastRow < 2 Then
        Set GetRecordsWhereStatus = result
        Exit Function
    End If

    Dim r As Long
    Dim status As String
    For r = 2 To lastRow
        status = CStr(ws.Cells(r, COL_STATUS).Value)
        If statusFilter = "" Or status = statusFilter Then
            Dim rec As Object
            Set rec = CreateObject("Scripting.Dictionary")
            rec("ReportID") = ws.Cells(r, COL_REPORT_ID).Value
            rec("KlientFK") = ws.Cells(r, COL_KLIENT_FK).Value
            rec("NazwaKlienta") = ws.Cells(r, COL_NAZWA_KLIENTA).Value
            rec("MiesiacZgloszenia") = ws.Cells(r, COL_MIESIAC).Value
            rec("Fields") = ws.Cells(r, COL_FIELDS).Value
            rec("CNA_HandlowcaID") = ws.Cells(r, COL_CNA).Value
            rec("NrOddzialu") = ws.Cells(r, COL_NR_ODDZIALU).Value
            rec("CreatedTimestamp") = ws.Cells(r, COL_CREATED).Value
            rec("Status") = status
            rec("EmailRecipient") = ws.Cells(r, COL_RECIPIENT).Value
            rec("BatchSentTimestamp") = ws.Cells(r, COL_BATCH_SENT).Value
            result.Add rec
        End If
    Next r

    Set GetRecordsWhereStatus = result
End Function

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

' Best-effort sync ws_DataCache -> BNC_DataCache.xlsx.
Private Sub SyncToFile()
    Dim wbOut As Workbook
    Dim restoreScreen As Boolean
    Dim restoreAlerts As Boolean
    On Error GoTo Cleanup

    Dim folderPath As String
    folderPath = CStr(mod_UserCacheSync.GetUserField("CacheFolderPath"))
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
    mod_Utils.LogError "mod_DataCacheSync.SyncToFile", Err.Number, Err.Description
    On Error Resume Next
    If Not wbOut Is Nothing Then wbOut.Close SaveChanges:=False
    If restoreAlerts Then Application.DisplayAlerts = True
    If restoreScreen Then Application.ScreenUpdating = True
End Sub
