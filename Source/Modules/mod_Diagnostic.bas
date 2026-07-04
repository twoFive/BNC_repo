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
'    mod_Diagnostic.DumpVBComponents           - lista wszystkich modulow
'    mod_Diagnostic.ListPublicProcedures "mod_Utils"  - procedury w module
'    mod_Diagnostic.DumpModuleContent "mod_Utils"     - pelna tresc modulu
'    mod_Diagnostic.CountLinesTotal            - suma linii kodu w projekcie
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
    Debug.Print "BLAD: Brak dostepu do VBProject. Wlacz w Excelu:"
    Debug.Print "  File -> Options -> Trust Center -> Trust Center Settings"
    Debug.Print "  -> Macro Settings -> [x] Trust access to the VBA project object model"
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

            ' Znajdz linie z deklaracja Sub/Function/Property
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
' UWAGA: Immediate Window ma limit, dla modulow > ~200 linii wynik jest obciety.
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

' Zwraca sume linii kodu w calym projekcie (bez dokumentow Sheet/ThisWorkbook).
Public Function CountLinesTotal() As Long
    Dim vbc As Object
    Dim total As Long
    On Error GoTo TrustError
    For Each vbc In ThisWorkbook.VBProject.VBComponents
        If vbc.Type = 1 Or vbc.Type = 2 Or vbc.Type = 3 Then  ' Module, Class, Form
            total = total + vbc.CodeModule.CountOfLines
        End If
    Next vbc
    CountLinesTotal = total
    Debug.Print "Suma linii (modules + classes + forms, bez Sheet/ThisWorkbook): " & total
    Exit Function

TrustError:
    Debug.Print "BLAD: Brak dostepu do VBProject (patrz DumpVBComponents)."
    CountLinesTotal = -1
End Function

' Helper - padowanie w prawo do stalej szerokosci.
Private Function PadRight(s As String, width As Long) As String
    If Len(s) >= width Then
        PadRight = s & " "
    Else
        PadRight = s & Space(width - Len(s))
    End If
End Function

' Helper - pierwsza nazwa procedury w module (dla overview).
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
