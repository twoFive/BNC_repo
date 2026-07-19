Attribute VB_Name = "mod_Diagnostic"
Option Explicit

' ============================================================================
'  mod_Diagnostic - one-off diagnostyka stanu VB Project.
'  Do inspekcji po recznej edycji w VBE (co jest, co brakuje, jakie procedury).
'
'  WYMAGANE: "Trust access to the VBA project object model" w Excel Trust
'  Center -> Macro Settings. Bez tego dostep do ThisWorkbook.VBProject rzuca
'  Run-time error 1004 "Programmatic access to Visual Basic Project is not
'  trusted".
'
'  Uruchamianie z Immediate Window (Ctrl+G):
'
'  === TOP-LEVEL (rekomendowane) ===
'    mod_Diagnostic.AuditFullProject       - pelny audit wszystkiego naraz
'
'  === POJEDYNCZE SEKCJE ===
'    mod_Diagnostic.AuditModules           - moduly + Public API
'    mod_Diagnostic.AuditForms             - formularze (present + puste)
'    mod_Diagnostic.AuditSheets            - arkusze + naglowki
'    mod_Diagnostic.AuditThisWorkbook      - klasa ThisWorkbook
'
'  === LEGACY (basic listing) ===
'    mod_Diagnostic.DumpVBComponents       - flat lista modulow
'    mod_Diagnostic.ListPublicProcedures "mod_Utils"  - proc w module
'    mod_Diagnostic.DumpModuleContent "mod_Utils"     - pelna tresc modulu
'    mod_Diagnostic.CountLinesTotal        - suma linii kodu
' ============================================================================

' ============================================================================
'  EXPECTED PROJECT STATE
'  Edytuj gdy zmienia sie architektura projektu.
' ============================================================================

Private Function ExpectedModules() As Variant
    ExpectedModules = Array( _
        "mod_Utils", _
        "mod_Validation", _
        "mod_UserCacheSync", _
        "mod_DataCacheSync", _
        "mod_MailSender", _
        "mod_Export", _
        "mod_Tests", _
        "mod_Diagnostic" _
    )
End Function

Private Function ExpectedForms() As Variant
    ExpectedForms = Array( _
        "frm_Setup", _
        "frm_Main", _
        "frm_Log", _
        "frm_UserPicker" _
    )
End Function

Private Function ExpectedSheets() As Variant
    ExpectedSheets = Array( _
        "ws_UserCache", _
        "ws_DataCache", _
        "ws_UsersRegistry" _
    )
End Function

' Public procedury oczekiwane per modul.
Private Function ExpectedPublicProcs(moduleName As String) As Variant
    Select Case moduleName
        Case "mod_Utils"
            ExpectedPublicProcs = Array( _
                "LogInfo", "LogError", "FormatTimestampISO", "GetCurrentMonthYear", _
                "FileExists", "FolderExists", "EnsureFolderExists", "JoinPath", _
                "IsValidEmail", "IsValidLong")
        Case "mod_Validation"
            ExpectedPublicProcs = Array( _
                "ValidateEmail", "ValidateClientFK", "ValidateNonEmpty", _
                "ValidateLength", "ValidateMonthYear", "ValidateFolderPath", _
                "ValidateSetupData", "ValidateReportData")
        Case "mod_UserCacheSync"
            ExpectedPublicProcs = Array( _
                "GetUserField", "GetUserData", "SetUserField", "SaveUserData", _
                "IsSetupCompleted", "IsUserManager", "EnsureCacheFileExists", _
                "GetUsersCount", "CurrentUserID", "GetAllUsers", "SwitchUser", _
                "AddNewUser", "PrepareForNewUser")
        Case "mod_DataCacheSync"
            ExpectedPublicProcs = Array( _
                "AppendRecord", "GetPendingRecords", "GetAllRecords", _
                "MarkAsSent", "DeleteRecord", "EnsureCacheFileExists")
        Case "mod_MailSender"
            ExpectedPublicProcs = Array("SendBatch", "DetermineRecipient")
        Case "mod_Export"
            ExpectedPublicProcs = Array("ExportDataCache", "GetSuggestedExportFileName")
        Case "mod_Tests"
            ExpectedPublicProcs = Array( _
                "RunAllTests", "Test_mod_Utils", "Test_mod_UserCacheSync", _
                "Test_mod_DataCacheSync", "Test_mod_Validation", _
                "Test_mod_MailSender", "Test_mod_Export")
        Case "mod_Diagnostic"
            ExpectedPublicProcs = Array( _
                "AuditFullProject", "AuditModules", "AuditForms", "AuditSheets", _
                "AuditThisWorkbook", "DumpVBComponents", "ListPublicProcedures", _
                "DumpModuleContent", "CountLinesTotal")
        Case Else
            ExpectedPublicProcs = Array()
    End Select
End Function

' Naglowki oczekiwane w wierszu 1 arkusza (kolumny 1..N).
Private Function ExpectedSheetHeaders(sheetName As String) As Variant
    Select Case sheetName
        Case "ws_UserCache"
            ' Key-value - naglowka nie ma, klucze w kolumnie A
            ExpectedSheetHeaders = Array()  ' skip check
        Case "ws_DataCache"
            ExpectedSheetHeaders = Array( _
                "ReportID", "KlientFK", "NazwaKlienta", "MiesiacZgloszenia", _
                "Fields", "CNA_HandlowcaID", "NrOddzialu", "CreatedTimestamp", _
                "Status", "EmailRecipient", "BatchSentTimestamp")
        Case "ws_UsersRegistry"
            ExpectedSheetHeaders = Array( _
                "UserID", "Imie", "Nazwisko", "EmailHandlowca", "CNA_HandlowcaID", _
                "NrOddzialu", "EmailKierownika", "EmailBNC", "CacheFolderPath", _
                "DataRejestracji", "SetupCompleted", "DontShowSetupAgain", "LastLogin")
        Case Else
            ExpectedSheetHeaders = Array()
    End Select
End Function

' Handlery event'ow oczekiwane per formularz.
Private Function ExpectedFormHandlers(formName As String) As Variant
    Select Case formName
        Case "frm_Setup"
            ExpectedFormHandlers = Array( _
                "UserForm_Initialize", "btn_Save_Click", "btn_Cancel_Click", _
                "btn_CreateCacheFolder_Click", "btn_ShowTutorial_Click")
        Case "frm_Main"
            ExpectedFormHandlers = Array( _
                "UserForm_Initialize", "UserForm_Activate", "btn_AddToList_Click", _
                "btn_Clear_Click", "btn_SendBatch_Click", "btn_ShowLog_Click", _
                "btn_DeleteSelected_Click")
        Case "frm_Log"
            ExpectedFormHandlers = Array( _
                "UserForm_Activate", "btn_Export_Click", "btn_Back_Click")
        Case "frm_UserPicker"
            ExpectedFormHandlers = Array( _
                "UserForm_Activate", "btn_SelectUser_Click", "btn_AddNew_Click", _
                "btn_Cancel_Click")
        Case Else
            ExpectedFormHandlers = Array()
    End Select
End Function

' ============================================================================
'  MAIN AUDIT - pelny audit projektu
' ============================================================================

Public Sub AuditFullProject()
    On Error GoTo TrustError

    Debug.Print String(90, "=")
    Debug.Print "  FULL PROJECT AUDIT - " & ThisWorkbook.Name
    Debug.Print "  " & Format(Now(), "yyyy-mm-dd hh:nn:ss")
    Debug.Print String(90, "=")

    AuditModules
    AuditForms
    AuditSheets
    AuditThisWorkbook

    Debug.Print String(90, "=")
    Debug.Print "  AUDIT DONE - jesli powyzej wystapily [MISSING] lub [EMPTY], patrz TODO."
    Debug.Print String(90, "=")
    Exit Sub

TrustError:
    PrintTrustErrorHelp
End Sub

' ============================================================================
'  Audit poszczegolnych sekcji
' ============================================================================

' Sekcja MODULY: obecnosc + Public API surface per modul.
Public Sub AuditModules()
    Debug.Print vbNewLine & String(90, "-")
    Debug.Print "1. MODULY (.bas)"
    Debug.Print String(90, "-")

    Dim expected As Variant
    expected = ExpectedModules()

    Dim missingCount As Long, presentCount As Long
    Dim i As Long
    For i = LBound(expected) To UBound(expected)
        Dim modName As String
        modName = CStr(expected(i))

        Dim vbc As Object
        Set vbc = FindVBComponent(modName)

        If vbc Is Nothing Then
            Debug.Print "  [MISSING] " & modName & " - importuj z Source/Modules/" & modName & ".bas"
            missingCount = missingCount + 1
        Else
            Dim lines As Long
            lines = 0
            On Error Resume Next
            lines = vbc.CodeModule.CountOfLines
            On Error GoTo 0

            Debug.Print "  [OK]      " & PadRight(modName, 22) & lines & " lin.   " & _
                        CheckPublicProcs(vbc, modName)
            presentCount = presentCount + 1
        End If
    Next i

    Debug.Print "  --> " & presentCount & "/" & (UBound(expected) - LBound(expected) + 1) & _
                " modulow obecnych" & IIf(missingCount > 0, "  ·  " & missingCount & " MISSING", "")
End Sub

' Sekcja FORMULARZE: obecnosc + code line count + handlery.
Public Sub AuditForms()
    Debug.Print vbNewLine & String(90, "-")
    Debug.Print "2. FORMULARZE (UserForms)"
    Debug.Print String(90, "-")

    Dim expected As Variant
    expected = ExpectedForms()

    Dim missingCount As Long, emptyCount As Long, presentCount As Long
    Dim i As Long
    For i = LBound(expected) To UBound(expected)
        Dim formName As String
        formName = CStr(expected(i))

        Dim vbc As Object
        Set vbc = FindVBComponent(formName)

        If vbc Is Nothing Then
            Debug.Print "  [MISSING] " & formName & " - Insert->UserForm w VBE + wklej " & _
                        "Source/Forms/" & formName & ".code-behind.txt"
            missingCount = missingCount + 1
        Else
            Dim lines As Long
            lines = 0
            On Error Resume Next
            lines = vbc.CodeModule.CountOfLines
            On Error GoTo 0

            If lines = 0 Then
                Debug.Print "  [EMPTY]   " & PadRight(formName, 22) & _
                            "shell istnieje, code-behind NIE wklejony - wklej " & _
                            "Source/Forms/" & formName & ".code-behind.txt"
                emptyCount = emptyCount + 1
            Else
                Debug.Print "  [OK]      " & PadRight(formName, 22) & lines & " lin.   " & _
                            CheckFormHandlers(vbc, formName)
                presentCount = presentCount + 1
            End If
        End If
    Next i

    Debug.Print "  --> " & presentCount & "/" & (UBound(expected) - LBound(expected) + 1) & _
                " formularzy OK" & _
                IIf(emptyCount > 0, "  ·  " & emptyCount & " EMPTY", "") & _
                IIf(missingCount > 0, "  ·  " & missingCount & " MISSING", "")
End Sub

' Sekcja ARKUSZE: obecnosc + naglowki w wierszu 1.
Public Sub AuditSheets()
    Debug.Print vbNewLine & String(90, "-")
    Debug.Print "3. ARKUSZE (Worksheets)"
    Debug.Print String(90, "-")

    Dim expected As Variant
    expected = ExpectedSheets()

    Dim missingCount As Long, headerBadCount As Long, presentCount As Long
    Dim i As Long
    For i = LBound(expected) To UBound(expected)
        Dim sheetName As String
        sheetName = CStr(expected(i))

        Dim ws As Worksheet
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(sheetName)
        On Error GoTo 0

        If ws Is Nothing Then
            Dim autoNote As String
            If sheetName = "ws_UsersRegistry" Then
                autoNote = "  (auto-created przy pierwszym GetUsersCount/AddNewUser)"
            ElseIf sheetName = "ws_UserCache" Or sheetName = "ws_DataCache" Then
                autoNote = "  (utworzyc recznie w Excelu, tab-name = klucz)"
            End If
            Debug.Print "  [MISSING] " & sheetName & autoNote
            missingCount = missingCount + 1
        Else
            Dim headerCheck As String
            headerCheck = CheckSheetHeaders(ws, sheetName)
            Dim visibleTag As String
            Select Case ws.Visible
                Case xlSheetVisible: visibleTag = "visible"
                Case xlSheetHidden: visibleTag = "hidden"
                Case xlSheetVeryHidden: visibleTag = "very hidden"
            End Select

            Debug.Print "  [OK]      " & PadRight(sheetName, 22) & _
                        PadRight("(" & visibleTag & ")", 16) & headerCheck
            presentCount = presentCount + 1
        End If
    Next i

    Debug.Print "  --> " & presentCount & "/" & (UBound(expected) - LBound(expected) + 1) & _
                " arkuszy OK" & IIf(missingCount > 0, "  ·  " & missingCount & " MISSING", "")
End Sub

' Sekcja THISWORKBOOK: obecnosc handlera Workbook_Open.
Public Sub AuditThisWorkbook()
    Debug.Print vbNewLine & String(90, "-")
    Debug.Print "4. THISWORKBOOK"
    Debug.Print String(90, "-")

    Dim vbc As Object
    Set vbc = FindVBComponent("ThisWorkbook")

    If vbc Is Nothing Then
        Debug.Print "  [ERR]     ThisWorkbook nie znaleziony - to nie powinno sie zdarzyc"
        Exit Sub
    End If

    Dim lines As Long
    lines = vbc.CodeModule.CountOfLines
    Debug.Print "  [OK]      ThisWorkbook   " & lines & " lin."

    If ProcedureExists(vbc, "Workbook_Open") Then
        Debug.Print "  [OK]      Workbook_Open handler obecny"
    Else
        Debug.Print "  [MISSING] Workbook_Open - wklej Source/ThisWorkbook/ThisWorkbook.code.txt"
    End If
End Sub

' ============================================================================
'  Helpery audytu
' ============================================================================

' Znajduje VBComponent po nazwie. Zwraca Nothing gdy nie ma.
Private Function FindVBComponent(compName As String) As Object
    On Error GoTo NotFound
    Set FindVBComponent = ThisWorkbook.VBProject.VBComponents(compName)
    Exit Function
NotFound:
    Set FindVBComponent = Nothing
End Function

' Sprawdza czy modul zawiera wszystkie oczekiwane Public procedury.
' Zwraca string typu "API: 10/10" lub "API: 8/10 (brak: X, Y)".
Private Function CheckPublicProcs(vbc As Object, moduleName As String) As String
    Dim expected As Variant
    expected = ExpectedPublicProcs(moduleName)

    Dim expectedCount As Long
    If IsArray(expected) Then
        On Error Resume Next
        expectedCount = UBound(expected) - LBound(expected) + 1
        On Error GoTo 0
    End If

    If expectedCount = 0 Then
        CheckPublicProcs = ""
        Exit Function
    End If

    Dim foundCount As Long
    Dim missing As String
    Dim i As Long
    For i = LBound(expected) To UBound(expected)
        If ProcedureExists(vbc, CStr(expected(i))) Then
            foundCount = foundCount + 1
        Else
            missing = missing & CStr(expected(i)) & ", "
        End If
    Next i

    If foundCount = expectedCount Then
        CheckPublicProcs = "API: " & foundCount & "/" & expectedCount
    Else
        If Len(missing) > 0 Then missing = Left$(missing, Len(missing) - 2)
        CheckPublicProcs = "API: " & foundCount & "/" & expectedCount & _
                           "  (brak: " & missing & ")"
    End If
End Function

' Sprawdza czy formularz zawiera wszystkie oczekiwane handlery event'ow.
Private Function CheckFormHandlers(vbc As Object, formName As String) As String
    Dim expected As Variant
    expected = ExpectedFormHandlers(formName)

    Dim expectedCount As Long
    If IsArray(expected) Then
        On Error Resume Next
        expectedCount = UBound(expected) - LBound(expected) + 1
        On Error GoTo 0
    End If

    If expectedCount = 0 Then
        CheckFormHandlers = ""
        Exit Function
    End If

    Dim foundCount As Long
    Dim missing As String
    Dim i As Long
    For i = LBound(expected) To UBound(expected)
        If ProcedureExists(vbc, CStr(expected(i))) Then
            foundCount = foundCount + 1
        Else
            missing = missing & CStr(expected(i)) & ", "
        End If
    Next i

    If foundCount = expectedCount Then
        CheckFormHandlers = "Handlery: " & foundCount & "/" & expectedCount
    Else
        If Len(missing) > 0 Then missing = Left$(missing, Len(missing) - 2)
        CheckFormHandlers = "Handlery: " & foundCount & "/" & expectedCount & _
                            "  (brak: " & missing & ")"
    End If
End Function

' Czy procedura o danej nazwie istnieje w VBComponent.
Private Function ProcedureExists(vbc As Object, procName As String) As Boolean
    Dim cm As Object
    Set cm = vbc.CodeModule

    Dim procStartLine As Long
    Dim procKind As Long

    ' Sprobuj 3 kindy: vbext_pk_Proc(0), vbext_pk_Let(1), vbext_pk_Get(3)
    On Error Resume Next

    procStartLine = 0
    procStartLine = cm.ProcStartLine(procName, 0)
    If procStartLine > 0 Then
        ProcedureExists = True
        Exit Function
    End If

    Err.Clear
    procStartLine = cm.ProcStartLine(procName, 1)
    If procStartLine > 0 Then
        ProcedureExists = True
        Exit Function
    End If

    Err.Clear
    procStartLine = cm.ProcStartLine(procName, 3)
    If procStartLine > 0 Then
        ProcedureExists = True
        Exit Function
    End If

    On Error GoTo 0
    ProcedureExists = False
End Function

' Sprawdza czy naglowki w wierszu 1 arkusza pasuja do oczekiwanych.
Private Function CheckSheetHeaders(ws As Worksheet, sheetName As String) As String
    Dim expected As Variant
    expected = ExpectedSheetHeaders(sheetName)

    Dim expectedCount As Long
    If IsArray(expected) Then
        On Error Resume Next
        expectedCount = UBound(expected) - LBound(expected) + 1
        On Error GoTo 0
    End If

    If expectedCount = 0 Then
        ' Special case ws_UserCache - key-value, brak naglowka
        If sheetName = "ws_UserCache" Then
            Dim keyCount As Long
            keyCount = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
            If ws.Cells(1, 1).Value = "" Then keyCount = 0
            CheckSheetHeaders = "  " & keyCount & " kluczy w kol.A"
        Else
            CheckSheetHeaders = ""
        End If
        Exit Function
    End If

    Dim foundCount As Long
    Dim missing As String
    Dim i As Long
    For i = LBound(expected) To UBound(expected)
        Dim col As Long
        col = i - LBound(expected) + 1
        If StrComp(CStr(ws.Cells(1, col).Value), CStr(expected(i)), vbTextCompare) = 0 Then
            foundCount = foundCount + 1
        Else
            missing = missing & CStr(expected(i)) & "(col " & col & "), "
        End If
    Next i

    If foundCount = expectedCount Then
        CheckSheetHeaders = "  Naglowki: " & foundCount & "/" & expectedCount & _
                            IIf(sheetName = "ws_DataCache", "  Rows: " & DataRowsCount(ws), "")
    Else
        If Len(missing) > 0 Then missing = Left$(missing, Len(missing) - 2)
        CheckSheetHeaders = "  Naglowki: " & foundCount & "/" & expectedCount & _
                            "  (brak: " & missing & ")"
    End If
End Function

' Liczba wierszy danych w arkuszu tabelarycznym (bez naglowka).
Private Function DataRowsCount(ws As Worksheet) As Long
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    DataRowsCount = IIf(lastRow < 2, 0, lastRow - 1)
End Function

Private Sub PrintTrustErrorHelp()
    Debug.Print "BLAD: Brak dostepu do VBProject. Wlacz w Excelu:"
    Debug.Print "  File -> Options -> Trust Center -> Trust Center Settings"
    Debug.Print "  -> Macro Settings -> [x] Trust access to the VBA project object model"
End Sub

' ============================================================================
'  LEGACY - proste listowanie (bez audytu vs Expected)
' ============================================================================

' Lista modulow z projektu z typem, liczba linii, nazwa pierwszej procedury.
Public Sub DumpVBComponents()
    Dim vbc As Object
    Dim typeStr As String
    Dim lineCount As Long
    Dim firstProc As String
    Dim totalLines As Long
    Dim countModules As Long, countForms As Long, countClasses As Long, countDocs As Long

    On Error GoTo TrustError

    Debug.Print String(90, "=")
    Debug.Print "VBComponents diagnostic - " & ThisWorkbook.Name
    Debug.Print "Data: " & Format(Now(), "yyyy-mm-dd hh:nn:ss")
    Debug.Print String(90, "=")
    Debug.Print PadRight("Name", 32) & PadRight("Type", 10) & PadRight("Lines", 8) & "First procedure"
    Debug.Print String(90, "-")

    For Each vbc In ThisWorkbook.VBProject.VBComponents
        Select Case vbc.Type
            Case 1: typeStr = "Module": countModules = countModules + 1
            Case 2: typeStr = "Class": countClasses = countClasses + 1
            Case 3: typeStr = "Form": countForms = countForms + 1
            Case 100: typeStr = "Doc": countDocs = countDocs + 1
            Case Else: typeStr = "?(" & vbc.Type & ")"
        End Select

        lineCount = 0
        On Error Resume Next
        lineCount = vbc.CodeModule.CountOfLines
        On Error GoTo TrustError

        totalLines = totalLines + lineCount
        firstProc = FirstProcedureName(vbc)

        Debug.Print PadRight(vbc.Name, 32) & PadRight(typeStr, 10) & PadRight(CStr(lineCount), 8) & firstProc
    Next vbc

    Debug.Print String(90, "-")
    Debug.Print "PODSUMOWANIE: " & _
                countModules & " modul(y), " & _
                countForms & " form(y), " & _
                countClasses & " klas(y), " & _
                countDocs & " dokument(y)   ·   " & _
                totalLines & " lin. lacznie"
    Debug.Print String(90, "=")
    Exit Sub

TrustError:
    PrintTrustErrorHelp
End Sub

' Lista Public procedur (Sub, Function, Property) w danym module.
Public Sub ListPublicProcedures(moduleName As String)
    Dim vbc As Object
    On Error GoTo NotFound
    Set vbc = ThisWorkbook.VBProject.VBComponents(moduleName)
    On Error GoTo 0

    Debug.Print "Public procedures in " & moduleName & ":"
    Debug.Print String(60, "-")

    Dim cm As Object
    Set cm = vbc.CodeModule

    Dim i As Long, procName As String, procKind As Long, procLines As Long
    Dim procStartLine As Long, headerLine As String
    Dim visibility As String

    i = 1
    Do While i <= cm.CountOfLines
        On Error Resume Next
        procName = cm.ProcOfLine(i, procKind)
        On Error GoTo 0

        If Len(procName) > 0 Then
            procStartLine = cm.ProcStartLine(procName, procKind)
            procLines = cm.ProcCountLines(procName, procKind)

            Dim j As Long
            For j = procStartLine To procStartLine + procLines - 1
                headerLine = cm.Lines(j, 1)
                If InStr(headerLine, "Sub ") > 0 Or InStr(headerLine, "Function ") > 0 Or InStr(headerLine, "Property ") > 0 Then
                    Exit For
                End If
            Next j

            If InStr(headerLine, "Public") > 0 Then
                visibility = "Public "
            ElseIf InStr(headerLine, "Private") > 0 Then
                visibility = "Private"
            Else
                visibility = "(pub)  "
            End If

            Debug.Print "  " & visibility & vbTab & procName & vbTab & "(" & procLines & " lin.)"
            i = procStartLine + procLines
        Else
            i = i + 1
        End If
    Loop
    Debug.Print String(60, "-")
    Exit Sub

NotFound:
    Debug.Print "Module not found: " & moduleName & _
        "  ·  uzyj DumpVBComponents zeby zobaczyc dostepne moduly"
End Sub

' Wypisuje pelna zawartosc modulu do Immediate Window.
Public Sub DumpModuleContent(moduleName As String)
    Dim vbc As Object
    On Error GoTo NotFound
    Set vbc = ThisWorkbook.VBProject.VBComponents(moduleName)
    On Error GoTo 0

    Dim cm As Object
    Set cm = vbc.CodeModule

    Debug.Print String(80, "=")
    Debug.Print "= " & moduleName & "  (" & cm.CountOfLines & " linii)"
    Debug.Print String(80, "=")

    If cm.CountOfLines > 0 Then
        Debug.Print cm.Lines(1, cm.CountOfLines)
    End If

    Debug.Print String(80, "=")
    Debug.Print "= END " & moduleName
    Debug.Print String(80, "=")
    Exit Sub

NotFound:
    Debug.Print "Module not found: " & moduleName
End Sub

Public Function CountLinesTotal() As Long
    Dim vbc As Object
    Dim total As Long
    On Error GoTo TrustError
    For Each vbc In ThisWorkbook.VBProject.VBComponents
        If vbc.Type = 1 Or vbc.Type = 2 Or vbc.Type = 3 Then
            total = total + vbc.CodeModule.CountOfLines
        End If
    Next vbc
    CountLinesTotal = total
    Debug.Print "Suma linii (modules + classes + forms, bez Sheet/ThisWorkbook): " & total
    Exit Function

TrustError:
    PrintTrustErrorHelp
    CountLinesTotal = -1
End Function

' ============================================================================
'  Wewnetrzne helpery
' ============================================================================

Private Function PadRight(s As String, width As Long) As String
    If Len(s) >= width Then
        PadRight = s & " "
    Else
        PadRight = s & Space(width - Len(s))
    End If
End Function

Private Function FirstProcedureName(vbc As Object) As String
    Dim i As Long, procName As String, procKind As Long
    On Error Resume Next
    For i = 1 To vbc.CodeModule.CountOfLines
        procName = vbc.CodeModule.ProcOfLine(i, procKind)
        If Len(procName) > 0 Then
            FirstProcedureName = procName
            Exit Function
        End If
    Next i
    FirstProcedureName = "(brak)"
End Function
