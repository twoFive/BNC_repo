Attribute VB_Name = "mod_Validation"
Option Explicit

' ============================================================================
'  mod_Validation
'  Regu³y walidacji w jednym miejscu (warstwa logiki). Stateless, nie wo³a
'  sync ani UserForms.
'
'  Konwencja:
'    Validate*Data(...) As String  -> "" jeœli OK, komunikat b³êdu w p.p.
'    Validate*(...)     As Boolean -> True/False (atomowe)
' ============================================================================

' Limity d³ugoœci
Private Const MIN_NAZWA_KLIENTA As Long = 3
Private Const MAX_NAZWA_KLIENTA As Long = 200
Private Const MIN_IMIE_NAZWISKO As Long = 2
Private Const MAX_IMIE_NAZWISKO As Long = 100

' ----- Walidacje atomowe ---------------------------------------------------

Public Function ValidateEmail(email As String) As Boolean
    ValidateEmail = mod_Utils.IsValidEmail(email)
End Function

Public Function ValidateClientFK(fk As String) As Boolean
    ' KlientFK musi byæ liczb¹ ca³kowit¹ dodatni¹ (klucz obcy w s³owniku klientów).
    If Not mod_Utils.IsValidLong(fk) Then
        ValidateClientFK = False
        Exit Function
    End If
    ValidateClientFK = (CLng(Trim$(fk)) > 0)
End Function

Public Function ValidateNonEmpty(text As String) As Boolean
    ValidateNonEmpty = (Len(Trim$(text)) > 0)
End Function

Public Function ValidateLength(text As String, minLen As Long, maxLen As Long) As Boolean
    Dim n As Long
    n = Len(Trim$(text))
    ValidateLength = (n >= minLen And n <= maxLen)
End Function

' Format YYYY-MM (np. "2026-05") lub dowolna data parsowalna przez CDate.
Public Function ValidateMonthYear(text As String) As Boolean
    Dim t As String
    t = Trim$(text)
    If Len(t) = 0 Then Exit Function

    ' Wzorzec YYYY-MM
    If Len(t) = 7 Then
        If Mid$(t, 5, 1) = "-" Then
            Dim yearPart As String, monthPart As String
            yearPart = Left$(t, 4)
            monthPart = Right$(t, 2)
            If mod_Utils.IsValidLong(yearPart) And mod_Utils.IsValidLong(monthPart) Then
                Dim m As Long
                m = CLng(monthPart)
                If m >= 1 And m <= 12 Then
                    ValidateMonthYear = True
                    Exit Function
                End If
            End If
        End If
    End If

    ' Fallback: próbuj sparsowaæ jako datê
    On Error GoTo NotADate
    Dim d As Date
    d = CDate(t)
    ValidateMonthYear = True
    Exit Function
NotADate:
    ValidateMonthYear = False
End Function

' Œcie¿ka folderu - prosty test syntaktyczny (nie sprawdza istnienia).
Public Function ValidateFolderPath(path As String) As Boolean
    Dim t As String
    t = Trim$(path)
    If Len(t) < 3 Then Exit Function

    ' Drive letter (C:\) lub UNC (\\server\share)
    If Mid$(t, 2, 2) = ":\" Then
        ValidateFolderPath = True
    ElseIf Left$(t, 2) = "\\" Then
        ValidateFolderPath = True
    End If
End Function

' ----- Walidacje z³o¿one (zwracaj¹ komunikat b³êdu albo "") ---------------

' Walidacja danych z frm_Setup. userData = Scripting.Dictionary z polami:
'   Imie, Nazwisko, EmailHandlowca, CNA_HandlowcaID, NrOddzialu,
'   EmailKierownika, EmailBNC, CacheFolderPath
Public Function ValidateSetupData(userData As Object) As String
    If userData Is Nothing Then
        ValidateSetupData = "Brak danych do walidacji."
        Exit Function
    End If

    If Not ValidateLength(GetField(userData, "Imie"), MIN_IMIE_NAZWISKO, MAX_IMIE_NAZWISKO) Then
        ValidateSetupData = "Imiê musi mieæ od " & MIN_IMIE_NAZWISKO & _
                            " do " & MAX_IMIE_NAZWISKO & " znaków."
        Exit Function
    End If
    If Not ValidateLength(GetField(userData, "Nazwisko"), MIN_IMIE_NAZWISKO, MAX_IMIE_NAZWISKO) Then
        ValidateSetupData = "Nazwisko musi mieæ od " & MIN_IMIE_NAZWISKO & _
                            " do " & MAX_IMIE_NAZWISKO & " znaków."
        Exit Function
    End If

    If Not ValidateEmail(GetField(userData, "EmailHandlowca")) Then
        ValidateSetupData = "Niepoprawny format adresu Email s³u¿bowy."
        Exit Function
    End If

    If Not ValidateClientFK(GetField(userData, "CNA_HandlowcaID")) Then
        ValidateSetupData = "CNA (numer handlowca) musi byæ liczb¹ dodatni¹."
        Exit Function
    End If

    If Not ValidateNonEmpty(GetField(userData, "NrOddzialu")) Then
        ValidateSetupData = "Numer oddzia³u nie mo¿e byæ pusty."
        Exit Function
    End If

    If Not ValidateEmail(GetField(userData, "EmailKierownika")) Then
        ValidateSetupData = "Niepoprawny format adresu Email kierownika." & vbCrLf & _
                            "Je¿eli jesteœ kierownikiem, wpisz tu swój w³asny adres s³u¿bowy."
        Exit Function
    End If

    If Not ValidateEmail(GetField(userData, "EmailBNC")) Then
        ValidateSetupData = "Niepoprawny format adresu Email zespo³u BNC."
        Exit Function
    End If

    If Not ValidateFolderPath(GetField(userData, "CacheFolderPath")) Then
        ValidateSetupData = "Niepoprawna œcie¿ka folderu cache." & vbCrLf & _
                            "Wymagany format: C:\Folder\ lub \\server\share\."
        Exit Function
    End If

    ValidateSetupData = ""  ' OK
End Function

' Walidacja danych z frm_Main (jedno zg³oszenie). reportData = Dictionary:
'   KlientFK, NazwaKlienta, MiesiacObrotu
Public Function ValidateReportData(reportData As Object) As String
    If reportData Is Nothing Then
        ValidateReportData = "Brak danych do walidacji."
        Exit Function
    End If

    If Not ValidateClientFK(GetField(reportData, "KlientFK")) Then
        ValidateReportData = "Klient FK musi byæ liczb¹ dodatni¹."
        Exit Function
    End If

    If Not ValidateLength(GetField(reportData, "NazwaKlienta"), _
                          MIN_NAZWA_KLIENTA, MAX_NAZWA_KLIENTA) Then
        ValidateReportData = "Nazwa klienta musi mieæ od " & MIN_NAZWA_KLIENTA & _
                             " do " & MAX_NAZWA_KLIENTA & " znaków."
        Exit Function
    End If

    If Not ValidateMonthYear(GetField(reportData, "MiesiacObrotu")) Then
        ValidateReportData = "Miesi¹c wykonania obrotu przez klienta w niepoprawnym formacie." & vbCrLf & _
                             "Wymagany format: YYYY-MM (np. 2026-05)."
        Exit Function
    End If

    ValidateReportData = ""  ' OK
End Function

' ----- Private helper ------------------------------------------------------

Private Function GetField(d As Object, key As String) As String
    If d Is Nothing Then
        GetField = ""
    ElseIf d.Exists(key) Then
        GetField = CStr(d(key))
    Else
        GetField = ""
    End If
End Function
