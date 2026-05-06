Attribute VB_Name = "mod_Validation"
Option Explicit

' ============================================================================
'  mod_Validation - reguly walidacji w jednym miejscu (warstwa logiki).
'  Stateless, nie wola sync ani UserForms. UI pyta "czy te dane sa poprawne",
'  ten modul odpowiada.
'
'  Konwencja:
'    Validate*Data(...) As String  -> "" jesli OK, w przeciwnym razie komunikat
'    Validate*(...)     As Boolean -> True/False (atomowe)
'  Patrz: BNC_Sender_PlanWdrozenia_FazaA.md (M2.1)
' ============================================================================

' Limity dlugosci
Private Const MIN_NAZWA_KLIENTA As Long = 3
Private Const MAX_NAZWA_KLIENTA As Long = 200
Private Const MIN_IMIE_NAZWISKO As Long = 2
Private Const MAX_IMIE_NAZWISKO As Long = 100
Private Const MAX_FIELDS As Long = 1000

' ============================================================================
'  Walidacje atomowe
' ============================================================================

Public Function ValidateEmail(email As String) As Boolean
    ValidateEmail = mod_Utils.IsValidEmail(email)
End Function

Public Function ValidateClientFK(fk As String) As Boolean
    ' KlientFK musi byc liczba calkowita dodatnia (klucz obcy w slowniku klientow).
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

    ' Fallback: probuj sparsowac jako date
    On Error GoTo NotADate
    Dim d As Date
    d = CDate(t)
    ValidateMonthYear = True
    Exit Function
NotADate:
    ValidateMonthYear = False
End Function

' Sciezka folderu - prosty test syntaktyczny (nie sprawdza istnienia).
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

' ============================================================================
'  Walidacje zlozone (zwracaja komunikat bledu albo "")
' ============================================================================

' Walidacja danych z frm_Setup. userData = Scripting.Dictionary z polami:
'   Imie, Nazwisko, EmailHandlowca, CNA_HandlowcaID, NrOddzialu,
'   EmailKierownika, EmailBNC, CacheFolderPath
Public Function ValidateSetupData(userData As Object) As String
    If userData Is Nothing Then
        ValidateSetupData = "Brak danych do walidacji."
        Exit Function
    End If

    ' Imie / Nazwisko
    If Not ValidateLength(GetField(userData, "Imie"), MIN_IMIE_NAZWISKO, MAX_IMIE_NAZWISKO) Then
        ValidateSetupData = "Imie musi miec od " & MIN_IMIE_NAZWISKO & _
                            " do " & MAX_IMIE_NAZWISKO & " znakow."
        Exit Function
    End If
    If Not ValidateLength(GetField(userData, "Nazwisko"), MIN_IMIE_NAZWISKO, MAX_IMIE_NAZWISKO) Then
        ValidateSetupData = "Nazwisko musi miec od " & MIN_IMIE_NAZWISKO & _
                            " do " & MAX_IMIE_NAZWISKO & " znakow."
        Exit Function
    End If

    ' Email handlowca - obowiazkowy
    If Not ValidateEmail(GetField(userData, "EmailHandlowca")) Then
        ValidateSetupData = "Niepoprawny format adresu Email sluzbowy."
        Exit Function
    End If

    ' CNA - liczba dodatnia
    If Not ValidateClientFK(GetField(userData, "CNA_HandlowcaID")) Then
        ValidateSetupData = "CNA (numer handlowca) musi byc liczba dodatnia."
        Exit Function
    End If

    ' Nr oddzialu - niepuste
    If Not ValidateNonEmpty(GetField(userData, "NrOddzialu")) Then
        ValidateSetupData = "Numer oddzialu nie moze byc pusty."
        Exit Function
    End If

    ' Email kierownika
    If Not ValidateEmail(GetField(userData, "EmailKierownika")) Then
        ValidateSetupData = "Niepoprawny format adresu Email kierownika." & vbCrLf & _
                            "Jezeli jestes kierownikiem, wpisz tu swoj wlasny adres sluzbowy."
        Exit Function
    End If

    ' Email BNC
    If Not ValidateEmail(GetField(userData, "EmailBNC")) Then
        ValidateSetupData = "Niepoprawny format adresu Email zespolu BNC."
        Exit Function
    End If

    ' Folder cache - format sciezki
    If Not ValidateFolderPath(GetField(userData, "CacheFolderPath")) Then
        ValidateSetupData = "Niepoprawna sciezka folderu cache." & vbCrLf & _
                            "Wymagany format: C:\Folder\ lub \\server\share\."
        Exit Function
    End If

    ValidateSetupData = ""  ' OK
End Function

' Walidacja danych z frm_Main (jedno zgloszenie). reportData = Dictionary:
'   KlientFK, NazwaKlienta, MiesiacZgloszenia, Fields
Public Function ValidateReportData(reportData As Object) As String
    If reportData Is Nothing Then
        ValidateReportData = "Brak danych do walidacji."
        Exit Function
    End If

    ' KlientFK - liczba dodatnia
    If Not ValidateClientFK(GetField(reportData, "KlientFK")) Then
        ValidateReportData = "Klient FK musi byc liczba dodatnia."
        Exit Function
    End If

    ' NazwaKlienta - 3-200 znakow
    If Not ValidateLength(GetField(reportData, "NazwaKlienta"), _
                          MIN_NAZWA_KLIENTA, MAX_NAZWA_KLIENTA) Then
        ValidateReportData = "Nazwa klienta musi miec od " & MIN_NAZWA_KLIENTA & _
                             " do " & MAX_NAZWA_KLIENTA & " znakow."
        Exit Function
    End If

    ' MiesiacZgloszenia - format YYYY-MM lub data
    If Not ValidateMonthYear(GetField(reportData, "MiesiacZgloszenia")) Then
        ValidateReportData = "Miesiac zgloszenia w niepoprawnym formacie." & vbCrLf & _
                             "Wymagany format: YYYY-MM (np. 2026-05)."
        Exit Function
    End If

    ' Fields - opcjonalne, ale jesli jest to max 1000 znakow
    Dim fields As String
    fields = GetField(reportData, "Fields")
    If Len(fields) > MAX_FIELDS Then
        ValidateReportData = "Pole dodatkowe za dlugie (max " & MAX_FIELDS & " znakow)."
        Exit Function
    End If

    ValidateReportData = ""  ' OK
End Function

' ============================================================================
'  Private helper
' ============================================================================

Private Function GetField(d As Object, key As String) As String
    If d Is Nothing Then
        GetField = ""
    ElseIf d.Exists(key) Then
        GetField = CStr(d(key))
    Else
        GetField = ""
    End If
End Function
