Attribute VB_Name = "mod_Utils"
Option Explicit

' ============================================================================
'  mod_Utils - helpery dostepne z calej aplikacji.
'  Logowanie do Immediate Window, formatowanie dat, operacje na plikach,
'  walidacja typow, JoinPath. Nizej niz wszystkie warstwy - nie wola innych
'  modulow aplikacji.
'  Patrz: BNC_Sender_PlanWdrozenia_FazaA.md (M1.1)
' ============================================================================

' ----- Logowanie -----------------------------------------------------------

Public Sub LogInfo(message As String)
    Debug.Print "[" & Format(Now(), "yyyy-mm-dd hh:nn:ss") & "] INFO: " & message
End Sub

Public Sub LogError(source As String, errNumber As Long, errDescription As String)
    Debug.Print "[" & Format(Now(), "yyyy-mm-dd hh:nn:ss") & "] ERROR in " & _
                source & ": #" & errNumber & " - " & errDescription
End Sub

' ----- Daty ----------------------------------------------------------------

Public Function FormatTimestampISO(dt As Date) As String
    FormatTimestampISO = Format(dt, "yyyy-mm-dd") & "T" & Format(dt, "hh:nn:ss")
End Function

' Pierwszy dzien biezacego miesiaca (np. dla MiesiacZgloszenia).
Public Function GetCurrentMonthYear() As Date
    GetCurrentMonthYear = DateSerial(Year(Now()), Month(Now()), 1)
End Function

' ----- Pliki ---------------------------------------------------------------

Public Function FileExists(filePath As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    FileExists = fso.FileExists(filePath)
End Function

Public Function FolderExists(folderPath As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    FolderExists = fso.FolderExists(folderPath)
End Function

' Tworzy folder rekursywnie - przechodzi w gore drzewa az do istniejacego rodzica.
Public Sub EnsureFolderExists(folderPath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Len(folderPath) = 0 Then Exit Sub
    If fso.FolderExists(folderPath) Then Exit Sub

    Dim parent As String
    parent = fso.GetParentFolderName(folderPath)
    If Len(parent) > 0 And Not fso.FolderExists(parent) Then
        EnsureFolderExists parent
    End If
    fso.CreateFolder folderPath
End Sub

' Buduje sciezke folderPath\fileName niezaleznie od separatora na koncu folderu.
Public Function JoinPath(folderPath As String, fileName As String) As String
    If Len(folderPath) = 0 Then
        JoinPath = fileName
    ElseIf Right$(folderPath, 1) = "\" Or Right$(folderPath, 1) = "/" Then
        JoinPath = folderPath & fileName
    Else
        JoinPath = folderPath & "\" & fileName
    End If
End Function

' ----- Walidacja typow -----------------------------------------------------

' Prosty test: cos@cos.cos - bez pelnej regex RFC 5322, wystarczy dla UI.
Public Function IsValidEmail(text As String) As Boolean
    Dim t As String
    t = Trim$(text)
    If Len(t) < 5 Then Exit Function

    Dim atPos As Long, dotPos As Long
    atPos = InStr(t, "@")
    If atPos < 2 Then Exit Function

    dotPos = InStrRev(t, ".")
    If dotPos < atPos + 2 Then Exit Function
    If dotPos = Len(t) Then Exit Function

    ' Brak bialych znakow wewnatrz
    If InStr(t, " ") > 0 Then Exit Function

    IsValidEmail = True
End Function

Public Function IsValidLong(text As String) As Boolean
    On Error GoTo ErrorHandler
    Dim t As String
    t = Trim$(text)
    If Len(t) = 0 Then Exit Function

    Dim n As Long
    n = CLng(t)
    IsValidLong = True
    Exit Function
ErrorHandler:
    IsValidLong = False
End Function
