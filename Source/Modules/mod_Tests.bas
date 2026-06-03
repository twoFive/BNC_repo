Attribute VB_Name = "mod_Tests"
Option Explicit

' ============================================================================
'  mod_Tests - smoke testy dla modulow Foundation.
'  Uruchamiaj z Immediate Window:
'      mod_Tests.RunAllTests
'      mod_Tests.Test_mod_Utils
'  Wyniki ladowane do Debug.Print (Ctrl+G w VBE).
' ============================================================================

Public Sub RunAllTests()
    Debug.Print "==================== RunAllTests START ===================="
    Test_mod_Utils
    Test_mod_UserCacheSync
    Test_mod_DataCacheSync
    Test_mod_Validation
    Test_mod_MailSender
    Test_mod_Export
    Debug.Print "==================== RunAllTests END   ===================="
End Sub

' ----- mod_Utils -----------------------------------------------------------

Public Sub Test_mod_Utils()
    Debug.Print "----- Test_mod_Utils -----"

    ' LogInfo / LogError - tylko sprawdz ze nie pada
    mod_Utils.LogInfo "Test message from Test_mod_Utils"
    mod_Utils.LogError "Test_mod_Utils", 999, "Symulowany blad"

    ' FormatTimestampISO
    Debug.Print "FormatTimestampISO: " & mod_Utils.FormatTimestampISO(Now())

    ' GetCurrentMonthYear
    Debug.Print "GetCurrentMonthYear: " & Format(mod_Utils.GetCurrentMonthYear(), "yyyy-mm-dd")

    ' IsValidEmail
    AssertEqual "IsValidEmail OK", True, mod_Utils.IsValidEmail("test@example.com")
    AssertEqual "IsValidEmail no-at", False, mod_Utils.IsValidEmail("not-email")
    AssertEqual "IsValidEmail no-tld", False, mod_Utils.IsValidEmail("test@example")
    AssertEqual "IsValidEmail empty", False, mod_Utils.IsValidEmail("")
    AssertEqual "IsValidEmail spaces", False, mod_Utils.IsValidEmail("a b@c.de")

    ' IsValidLong
    AssertEqual "IsValidLong int", True, mod_Utils.IsValidLong("12345")
    AssertEqual "IsValidLong neg", True, mod_Utils.IsValidLong("-1")
    AssertEqual "IsValidLong text", False, mod_Utils.IsValidLong("abc")
    AssertEqual "IsValidLong empty", False, mod_Utils.IsValidLong("")
    AssertEqual "IsValidLong float", False, mod_Utils.IsValidLong("1.5")

    ' FileExists - aktualny xlsm musi istniec
    AssertEqual "FileExists self", True, mod_Utils.FileExists(ThisWorkbook.FullName)
    AssertEqual "FileExists fake", False, mod_Utils.FileExists("C:\__nope__\nope.xyz")

    ' FolderExists
    AssertEqual "FolderExists C:\", True, mod_Utils.FolderExists("C:\")
    AssertEqual "FolderExists fake", False, mod_Utils.FolderExists("C:\__nope_folder__")

    ' JoinPath
    AssertEqual "JoinPath no-slash", "C:\foo\bar.txt", mod_Utils.JoinPath("C:\foo", "bar.txt")
    AssertEqual "JoinPath with-slash", "C:\foo\bar.txt", mod_Utils.JoinPath("C:\foo\", "bar.txt")

    ' EnsureFolderExists - utworz tymczasowy podfolder w %TEMP%
    Dim testFolder As String
    testFolder = mod_Utils.JoinPath(Environ("TEMP"), "BNC_Test_" & Format(Now(), "yyyymmddhhnnss"))
    mod_Utils.EnsureFolderExists testFolder
    AssertEqual "EnsureFolderExists created", True, mod_Utils.FolderExists(testFolder)
    On Error Resume Next
    RmDir testFolder
    On Error GoTo 0

    Debug.Print "----- Test_mod_Utils DONE -----"
End Sub

' ----- mod_UserCacheSync ---------------------------------------------------

' UWAGA: ten test pisze i odczytuje z ws_UserCache. Operuje na tymczasowych
' kluczach z prefiksem "_TEST_" zeby nie zepsuc setupu uzytkownika - na koncu
' kasuje uzyte klucze.
Public Sub Test_mod_UserCacheSync()
    Debug.Print "----- Test_mod_UserCacheSync -----"

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("ws_UserCache")
    On Error GoTo 0
    If ws Is Nothing Then
        Debug.Print "  [SKIP] arkusz ws_UserCache nie istnieje - wymagany manualny setup z M0"
        Exit Sub
    End If

    ' --- read API ---
    Debug.Print "GetUserField('Imie') = '" & CStr(mod_UserCacheSync.GetUserField("Imie")) & "'"
    Debug.Print "IsSetupCompleted    = " & mod_UserCacheSync.IsSetupCompleted()
    Debug.Print "IsUserManager       = " & mod_UserCacheSync.IsUserManager()

    ' --- write API na tymczasowym kluczu ---
    Const TEST_KEY As String = "_TEST_RoundTrip_"
    Const TEST_VAL As String = "smoke_test_value_42"

    mod_UserCacheSync.SetUserField TEST_KEY, TEST_VAL
    AssertEqual "SetUserField round-trip", TEST_VAL, mod_UserCacheSync.GetUserField(TEST_KEY)

    ' --- IsUserManager: przemiennosc ---
    ' Tylko log - nie wymuszamy konkretnej wartosci, bo zalezy od stanu UserCache.
    Debug.Print "  [info] EmailHandlowca='" & _
        CStr(mod_UserCacheSync.GetUserField("EmailHandlowca")) & "'"
    Debug.Print "  [info] EmailKierownika='" & _
        CStr(mod_UserCacheSync.GetUserField("EmailKierownika")) & "'"

    ' --- GetUserData zwraca Dictionary z wszystkimi kanonicznymi polami ---
    Dim d As Object
    Set d = mod_UserCacheSync.GetUserData()
    AssertEqual "GetUserData has Imie", True, d.Exists("Imie")
    AssertEqual "GetUserData has SetupCompleted", True, d.Exists("SetupCompleted")

    ' --- cleanup tymczasowego klucza ---
    Dim r As Long
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    For r = 1 To lastRow
        If CStr(ws.Cells(r, 1).Value) = TEST_KEY Then
            ws.Cells(r, 1).Value = ""
            ws.Cells(r, 2).Value = ""
            Exit For
        End If
    Next r
    ThisWorkbook.Save

    Debug.Print "----- Test_mod_UserCacheSync DONE -----"
End Sub

' ----- mod_DataCacheSync ---------------------------------------------------

' UWAGA: ten test pisze do ws_DataCache (jeden tymczasowy rekord) i kasuje
' go na koncu. Operuje na realnej tabeli, ale czysci po sobie.
Public Sub Test_mod_DataCacheSync()
    Debug.Print "----- Test_mod_DataCacheSync -----"

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("ws_DataCache")
    On Error GoTo 0
    If ws Is Nothing Then
        Debug.Print "  [SKIP] arkusz ws_DataCache nie istnieje - wymagany manualny setup z M0"
        Exit Sub
    End If

    ' --- AppendRecord ---
    Dim newRecord As Object
    Set newRecord = CreateObject("Scripting.Dictionary")
    newRecord("KlientFK") = 99999
    newRecord("NazwaKlienta") = "_TEST_Klient_DoUsuniecia_"
    newRecord("MiesiacZgloszenia") = mod_Utils.GetCurrentMonthYear()
    newRecord("Fields") = "smoke test fields"

    Dim newID As Long
    newID = mod_DataCacheSync.AppendRecord(newRecord)
    Debug.Print "  AppendRecord -> ReportID=" & newID
    AssertEqual "AppendRecord returned >0", True, (newID > 0)

    ' --- GetPendingRecords zawiera nowy rekord ---
    Dim pending As Collection
    Set pending = mod_DataCacheSync.GetPendingRecords()
    Dim found As Boolean
    Dim rec As Object
    For Each rec In pending
        If CLng(rec("ReportID")) = newID Then
            found = True
            AssertEqual "pending.Status", "pending", CStr(rec("Status"))
            AssertEqual "pending.NazwaKlienta", "_TEST_Klient_DoUsuniecia_", CStr(rec("NazwaKlienta"))
            Exit For
        End If
    Next rec
    AssertEqual "GetPendingRecords contains new ID", True, found

    ' --- MarkAsSent zmienia status ---
    Dim ids As New Collection
    ids.Add newID
    mod_DataCacheSync.MarkAsSent ids, "test_recipient@example.com"

    Dim allRecs As Collection
    Set allRecs = mod_DataCacheSync.GetAllRecords()
    Dim verifiedSent As Boolean
    For Each rec In allRecs
        If CLng(rec("ReportID")) = newID Then
            AssertEqual "MarkAsSent.Status", "sent", CStr(rec("Status"))
            AssertEqual "MarkAsSent.EmailRecipient", "test_recipient@example.com", CStr(rec("EmailRecipient"))
            verifiedSent = True
            Exit For
        End If
    Next rec
    AssertEqual "MarkAsSent verified", True, verifiedSent

    ' --- DeleteRecord: odmowa dla sent (ADR-006) ---
    AssertEqual "DeleteRecord rejects sent", False, mod_DataCacheSync.DeleteRecord(newID)

    ' Zweryfikuj ze record nadal istnieje
    Set allRecs = mod_DataCacheSync.GetAllRecords()
    Dim stillThere As Boolean
    For Each rec In allRecs
        If CLng(rec("ReportID")) = newID Then stillThere = True : Exit For
    Next rec
    AssertEqual "Sent record still exists after refused delete", True, stillThere

    ' --- DeleteRecord: zgoda dla pending (round-trip) ---
    Dim pendingRecord As Object
    Set pendingRecord = CreateObject("Scripting.Dictionary")
    pendingRecord("KlientFK") = 88888
    pendingRecord("NazwaKlienta") = "_TEST_Klient_DoDelete_"
    pendingRecord("MiesiacZgloszenia") = mod_Utils.GetCurrentMonthYear()
    pendingRecord("Fields") = "smoke test delete"

    Dim deleteID As Long
    deleteID = mod_DataCacheSync.AppendRecord(pendingRecord)
    AssertEqual "DeleteRecord accepts pending", True, mod_DataCacheSync.DeleteRecord(deleteID)

    ' Zweryfikuj ze record znikl
    Set allRecs = mod_DataCacheSync.GetAllRecords()
    Dim foundDeleted As Boolean
    For Each rec In allRecs
        If CLng(rec("ReportID")) = deleteID Then foundDeleted = True : Exit For
    Next rec
    AssertEqual "Pending record gone after delete", False, foundDeleted

    ' --- DeleteRecord: odmowa dla nieistniejacego ID ---
    AssertEqual "DeleteRecord rejects unknown ID", False, mod_DataCacheSync.DeleteRecord(999999)

    ' --- cleanup pierwszego (sent) rekordu - bezposredni dostep, bo DeleteRecord
    '     z definicji odmawia sent records (ADR-006 enforcement). Bypassing API
    '     jest celowe tylko w test cleanup, nigdy w produkcji.
    Dim r As Long
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    For r = lastRow To 2 Step -1
        If CStr(ws.Cells(r, 1).Value) = CStr(newID) Then
            ws.Rows(r).Delete
            Exit For
        End If
    Next r
    ThisWorkbook.Save

    Debug.Print "----- Test_mod_DataCacheSync DONE -----"
End Sub

' ----- mod_Validation -----------------------------------------------------

Public Sub Test_mod_Validation()
    Debug.Print "----- Test_mod_Validation -----"

    ' --- Walidacje atomowe ---
    AssertEqual "ValidateEmail OK", True, mod_Validation.ValidateEmail("test@example.com")
    AssertEqual "ValidateEmail bad", False, mod_Validation.ValidateEmail("not-email")

    AssertEqual "ValidateClientFK pos", True, mod_Validation.ValidateClientFK("12345")
    AssertEqual "ValidateClientFK zero", False, mod_Validation.ValidateClientFK("0")
    AssertEqual "ValidateClientFK neg", False, mod_Validation.ValidateClientFK("-1")
    AssertEqual "ValidateClientFK text", False, mod_Validation.ValidateClientFK("abc")

    AssertEqual "ValidateNonEmpty ok", True, mod_Validation.ValidateNonEmpty("hello")
    AssertEqual "ValidateNonEmpty spaces", False, mod_Validation.ValidateNonEmpty("   ")
    AssertEqual "ValidateNonEmpty empty", False, mod_Validation.ValidateNonEmpty("")

    AssertEqual "ValidateLength in", True, mod_Validation.ValidateLength("abcde", 3, 10)
    AssertEqual "ValidateLength too short", False, mod_Validation.ValidateLength("ab", 3, 10)
    AssertEqual "ValidateLength too long", False, mod_Validation.ValidateLength("abcdefghijk", 3, 10)

    AssertEqual "ValidateMonthYear yyyy-mm", True, mod_Validation.ValidateMonthYear("2026-05")
    AssertEqual "ValidateMonthYear bad month", False, mod_Validation.ValidateMonthYear("2026-13")
    AssertEqual "ValidateMonthYear empty", False, mod_Validation.ValidateMonthYear("")
    AssertEqual "ValidateMonthYear text", False, mod_Validation.ValidateMonthYear("not-a-date")

    AssertEqual "ValidateFolderPath drive", True, mod_Validation.ValidateFolderPath("C:\Foo\")
    AssertEqual "ValidateFolderPath unc", True, mod_Validation.ValidateFolderPath("\\server\share")
    AssertEqual "ValidateFolderPath bad", False, mod_Validation.ValidateFolderPath("Foo")

    ' --- ValidateSetupData (zlozone) ---
    Dim setup As Object
    Set setup = CreateObject("Scripting.Dictionary")
    setup("Imie") = "Jan"
    setup("Nazwisko") = "Kowalski"
    setup("EmailHandlowca") = "jan@firma.pl"
    setup("CNA_HandlowcaID") = "12345"
    setup("NrOddzialu") = "W001"
    setup("EmailKierownika") = "kier@firma.pl"
    setup("EmailBNC") = "bnc@firma.pl"
    setup("CacheFolderPath") = "C:\BNC_CacheFolder\"
    AssertEqual "ValidateSetupData OK", "", mod_Validation.ValidateSetupData(setup)

    setup("EmailHandlowca") = "broken-email"
    AssertEqual "ValidateSetupData bad email", _
        True, (Len(mod_Validation.ValidateSetupData(setup)) > 0)
    setup("EmailHandlowca") = "jan@firma.pl"  ' przywroc

    setup("CNA_HandlowcaID") = "abc"
    AssertEqual "ValidateSetupData bad CNA", _
        True, (Len(mod_Validation.ValidateSetupData(setup)) > 0)

    ' --- ValidateReportData ---
    Dim rep As Object
    Set rep = CreateObject("Scripting.Dictionary")
    rep("KlientFK") = "999"
    rep("NazwaKlienta") = "Acme Sp. z o.o."
    rep("MiesiacZgloszenia") = "2026-05"
    rep("Fields") = "extra info"
    AssertEqual "ValidateReportData OK", "", mod_Validation.ValidateReportData(rep)

    rep("NazwaKlienta") = "ab"  ' za krotkie
    AssertEqual "ValidateReportData short name", _
        True, (Len(mod_Validation.ValidateReportData(rep)) > 0)

    Debug.Print "----- Test_mod_Validation DONE -----"
End Sub

' ----- mod_MailSender ------------------------------------------------------

' UWAGA: ten test NIE wysyla maila - testuje tylko DetermineRecipient
' (czysta funkcja). Pelna sciezka wysylki SendBatch wymaga manualnego UAT
' z Outlookiem (M4.1.3-4.1.5 w planie).
'
' Test manipuluje EmailKierownika w UserCache zeby sprawdzic obie galezie
' decision diamond, na koncu przywraca oryginalna wartosc.
Public Sub Test_mod_MailSender()
    Debug.Print "----- Test_mod_MailSender -----"

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("ws_UserCache")
    On Error GoTo 0
    If ws Is Nothing Then
        Debug.Print "  [SKIP] arkusz ws_UserCache nie istnieje"
        Exit Sub
    End If

    ' Backup oryginalnych wartosci
    Dim origKierownika As String, origHandlowca As String
    origKierownika = CStr(mod_UserCacheSync.GetUserField("EmailKierownika"))
    origHandlowca = CStr(mod_UserCacheSync.GetUserField("EmailHandlowca"))

    ' Scenariusz 1: HANDLOWIEC (kierownika != handlowca) -> mail do kierownika
    mod_UserCacheSync.SetUserField "EmailHandlowca", "handlowiec@firma.pl"
    mod_UserCacheSync.SetUserField "EmailKierownika", "kierownik@firma.pl"

    Dim r As Object
    Set r = mod_MailSender.DetermineRecipient()
    AssertEqual "Handlowiec.To = kierownik", "kierownik@firma.pl", CStr(r("To"))
    AssertEqual "Handlowiec.Subject ma 'akceptacji'", _
        True, (InStr(CStr(r("Subject")), "akceptacji") > 0)
    AssertEqual "Handlowiec.Body ma EmailBNC", _
        True, (InStr(CStr(r("Body")), CStr(mod_UserCacheSync.GetUserField("EmailBNC"))) > 0)

    ' Scenariusz 2: KIEROWNIK (kierownika == handlowca) -> mail wprost do BNC
    mod_UserCacheSync.SetUserField "EmailKierownika", "handlowiec@firma.pl"

    Set r = mod_MailSender.DetermineRecipient()
    AssertEqual "Kierownik.To = EmailBNC", _
        CStr(mod_UserCacheSync.GetUserField("EmailBNC")), CStr(r("To"))
    AssertEqual "Kierownik.Subject NIE ma 'akceptacji'", _
        False, (InStr(CStr(r("Subject")), "akceptacji") > 0)

    ' Przywroc oryginalne wartosci
    mod_UserCacheSync.SetUserField "EmailHandlowca", origHandlowca
    mod_UserCacheSync.SetUserField "EmailKierownika", origKierownika

    Debug.Print "----- Test_mod_MailSender DONE -----"
End Sub

' ----- mod_Export ----------------------------------------------------------

Public Sub Test_mod_Export()
    Debug.Print "----- Test_mod_Export -----"

    ' GetSuggestedExportFileName - format BNC_Eksport_<Nazwisko>_<data>.xlsx
    Dim suggested As String
    suggested = mod_Export.GetSuggestedExportFileName()
    Debug.Print "  GetSuggestedExportFileName: " & suggested
    AssertEqual "Suggested startsWith BNC_Eksport_", _
        True, (Left$(suggested, 12) = "BNC_Eksport_")
    AssertEqual "Suggested endsWith .xlsx", _
        True, (Right$(suggested, 5) = ".xlsx")

    ' ExportDataCache - probuje skopiowac, jesli zrodlo nie istnieje -> False
    Dim folderPath As String
    folderPath = CStr(mod_UserCacheSync.GetUserField("CacheFolderPath"))
    If Len(folderPath) = 0 Then
        Debug.Print "  [SKIP] CacheFolderPath nie ustawiony - wymagany setup"
        Exit Sub
    End If

    ' Test: kopia do %TEMP% - bezpieczna lokalizacja
    Dim targetPath As String
    targetPath = mod_Utils.JoinPath(Environ("TEMP"), "BNC_Test_Export_" & _
                 Format(Now(), "yyyymmddhhnnss") & ".xlsx")

    ' Upewnij sie ze BNC_DataCache.xlsx istnieje (auto-recreate jesli nie).
    mod_DataCacheSync.EnsureCacheFileExists

    Dim ok As Boolean
    ok = mod_Export.ExportDataCache(targetPath)
    AssertEqual "ExportDataCache success", True, ok
    AssertEqual "Target file exists after export", True, mod_Utils.FileExists(targetPath)

    ' Cleanup pliku testowego
    On Error Resume Next
    Kill targetPath
    On Error GoTo 0

    Debug.Print "----- Test_mod_Export DONE -----"
End Sub

' ----- Helpery testow ------------------------------------------------------

Private Sub AssertEqual(name As String, expected As Variant, actual As Variant)
    Dim okText As String
    If CStr(expected) = CStr(actual) Then
        okText = "PASS"
    Else
        okText = "FAIL"
    End If
    Debug.Print "  [" & okText & "] " & name & _
                " | expected=" & CStr(expected) & " actual=" & CStr(actual)
End Sub
