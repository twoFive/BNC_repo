Attribute VB_Name = "mod_Tests"
Option Explicit

' ============================================================================
'  mod_Tests
'  Smoke testy dla modu³ów Foundation. Uruchamiaj z Immediate Window:
'      mod_Tests.RunAllTests
'      mod_Tests.Test_mod_Utils
'
'  Wyniki -> Debug.Print (Ctrl+G w VBE).
' ============================================================================

Public Sub RunAllTests()
    Debug.Print "==================== RunAllTests START ===================="
    Test_mod_Utils
    Test_IsFormOpen
    Test_mod_AppStateSync
    Test_mod_DataCacheSync
    Test_mod_Validation
    Test_mod_MailSender
    Test_mod_Export
    Test_MultiUser
    Debug.Print "==================== RunAllTests END   ===================="
End Sub

' ----- mod_Utils ----------------------------------------------------------

Public Sub Test_mod_Utils()
    Debug.Print "----- Test_mod_Utils -----"

    mod_Utils.LogInfo "Test message from Test_mod_Utils"
    mod_Utils.LogError "Test_mod_Utils", 999, "Symulowany blad"

    Debug.Print "FormatTimestampISO: " & mod_Utils.FormatTimestampISO(Now())
    Debug.Print "GetCurrentMonthYear: " & Format(mod_Utils.GetCurrentMonthYear(), "yyyy-mm-dd")

    AssertEqual "IsValidEmail OK", True, mod_Utils.IsValidEmail("test@example.com")
    AssertEqual "IsValidEmail no-at", False, mod_Utils.IsValidEmail("not-email")
    AssertEqual "IsValidEmail no-tld", False, mod_Utils.IsValidEmail("test@example")
    AssertEqual "IsValidEmail empty", False, mod_Utils.IsValidEmail("")
    AssertEqual "IsValidEmail spaces", False, mod_Utils.IsValidEmail("a b@c.de")

    AssertEqual "IsValidLong int", True, mod_Utils.IsValidLong("12345")
    AssertEqual "IsValidLong neg", True, mod_Utils.IsValidLong("-1")
    AssertEqual "IsValidLong text", False, mod_Utils.IsValidLong("abc")
    AssertEqual "IsValidLong empty", False, mod_Utils.IsValidLong("")
    AssertEqual "IsValidLong float", False, mod_Utils.IsValidLong("1.5")

    AssertEqual "FileExists self", True, mod_Utils.FileExists(ThisWorkbook.FullName)
    AssertEqual "FileExists fake", False, mod_Utils.FileExists("C:\__nope__\nope.xyz")

    AssertEqual "FolderExists C:\", True, mod_Utils.FolderExists("C:\")
    AssertEqual "FolderExists fake", False, mod_Utils.FolderExists("C:\__nope_folder__")

    AssertEqual "JoinPath no-slash", "C:\foo\bar.txt", mod_Utils.JoinPath("C:\foo", "bar.txt")
    AssertEqual "JoinPath with-slash", "C:\foo\bar.txt", mod_Utils.JoinPath("C:\foo\", "bar.txt")

    ' EnsureFolderExists - utwórz tymczasowy podfolder w %TEMP%
    Dim testFolder As String
    testFolder = mod_Utils.JoinPath(Environ("TEMP"), "BNC_Test_" & Format(Now(), "yyyymmddhhnnss"))
    mod_Utils.EnsureFolderExists testFolder
    AssertEqual "EnsureFolderExists created", True, mod_Utils.FolderExists(testFolder)
    On Error Resume Next
    RmDir testFolder
    On Error GoTo 0

    Debug.Print "----- Test_mod_Utils DONE -----"
End Sub

' ----- IsFormOpen (regression guard: zombie hidden forms) ------------------

' Regression test dla bugfix 2026-08-10: IsFormOpen sprawdza .Visible,
' nie tylko istnienie w UserForms. Bez testu regresja mog³aby cicho
' przywróciæ bug (blokada re-Show po Me.Hide).
'
' Nie testujemy visible=True case - wymaga³oby to Show vbModal blokuj¹cego
' wykonanie tego suba (modal loop). Manual testing dla visible case.
Public Sub Test_IsFormOpen()
    Debug.Print "----- Test_IsFormOpen -----"

    AssertEqual "IsFormOpen nonexistent", False, _
        mod_Utils.IsFormOpen("frm_NoSuchForm_XYZ")

    ' Loaded ale niewidoczny (emulacja stanu po Me.Hide) -> False (zombie check).
    Load frm_Setup
    AssertEqual "IsFormOpen loaded-but-hidden (zombie)", False, _
        mod_Utils.IsFormOpen("frm_Setup")

    Unload frm_Setup

    AssertEqual "IsFormOpen after Unload", False, _
        mod_Utils.IsFormOpen("frm_Setup")

    Debug.Print "----- Test_IsFormOpen DONE -----"
End Sub

' ----- mod_AppStateSync ---------------------------------------------------

' Pisze i odczytuje z ws_AppState pod tymczasowym kluczem "_TEST_RoundTrip_" -
' na koñcu kasuje. NIE dotyka _CurrentUserID ¿eby nie zepsuæ sesji.
Public Sub Test_mod_AppStateSync()
    Debug.Print "----- Test_mod_AppStateSync -----"

    Const TEST_KEY As String = "_TEST_RoundTrip_"
    Const TEST_VAL As String = "smoke_test_value_42"

    mod_AppStateSync.SetAppValue TEST_KEY, TEST_VAL
    AssertEqual "SetAppValue round-trip", TEST_VAL, _
        CStr(mod_AppStateSync.GetAppValue(TEST_KEY))

    AssertEqual "GetAppValue missing key returns empty", "", _
        CStr(mod_AppStateSync.GetAppValue("_TEST_NonExistent_"))

    Dim currentId As String
    currentId = CStr(mod_AppStateSync.GetAppValue("_CurrentUserID"))
    Debug.Print "  [info] _CurrentUserID = '" & currentId & "'"
    AssertEqual "_CurrentUserID non-empty (session ma aktywnego usera)", True, _
        (Len(currentId) > 0)

    Dim ws As Worksheet
    Set ws = mod_AppStateSync.EnsureAppStateSheet()
    AssertEqual "EnsureAppStateSheet returns sheet", True, Not (ws Is Nothing)
    AssertEqual "EnsureAppStateSheet ws.Name", "ws_AppState", ws.Name

    ' Cleanup tymczasowego klucza
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

    Debug.Print "----- Test_mod_AppStateSync DONE -----"
End Sub

' ----- mod_DataCacheSync --------------------------------------------------

' Pisze do ws_DataCache jeden tymczasowy rekord i kasuje na koñcu.
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

    ' AppendRecord
    Dim newRecord As Object
    Set newRecord = CreateObject("Scripting.Dictionary")
    newRecord("KlientFK") = 99999
    newRecord("NazwaKlienta") = "_TEST_Klient_DoUsuniecia_"
    newRecord("MiesiacObrotu") = Format(mod_Utils.GetCurrentMonthYear(), "yyyy-mm")

    Dim newID As Long
    newID = mod_DataCacheSync.AppendRecord(newRecord)
    Debug.Print "  AppendRecord -> ReportID=" & newID
    AssertEqual "AppendRecord returned >0", True, (newID > 0)

    ' GetPendingRecords zawiera nowy rekord
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

    ' MarkAsSent
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

    ' DeleteRecord: odmowa dla sent (ADR-006)
    AssertEqual "DeleteRecord rejects sent", False, mod_DataCacheSync.DeleteRecord(newID)

    Set allRecs = mod_DataCacheSync.GetAllRecords()
    Dim stillThere As Boolean
    For Each rec In allRecs
        If CLng(rec("ReportID")) = newID Then stillThere = True : Exit For
    Next rec
    AssertEqual "Sent record still exists after refused delete", True, stillThere

    ' DeleteRecord: zgoda dla pending (round-trip)
    Dim pendingRecord As Object
    Set pendingRecord = CreateObject("Scripting.Dictionary")
    pendingRecord("KlientFK") = 88888
    pendingRecord("NazwaKlienta") = "_TEST_Klient_DoDelete_"
    pendingRecord("MiesiacObrotu") = Format(mod_Utils.GetCurrentMonthYear(), "yyyy-mm")

    Dim deleteID As Long
    deleteID = mod_DataCacheSync.AppendRecord(pendingRecord)
    AssertEqual "DeleteRecord accepts pending", True, mod_DataCacheSync.DeleteRecord(deleteID)

    Set allRecs = mod_DataCacheSync.GetAllRecords()
    Dim foundDeleted As Boolean
    For Each rec In allRecs
        If CLng(rec("ReportID")) = deleteID Then foundDeleted = True : Exit For
    Next rec
    AssertEqual "Pending record gone after delete", False, foundDeleted

    AssertEqual "DeleteRecord rejects unknown ID", False, mod_DataCacheSync.DeleteRecord(999999)

    ' Cleanup sent rekordu bezpoœrednio (bypass API - DeleteRecord odmawia
    ' sent per ADR-006; bypass akceptowalny tylko w test cleanup).
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

    ' ValidateSetupData (z³o¿one)
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
    setup("EmailHandlowca") = "jan@firma.pl"

    setup("CNA_HandlowcaID") = "abc"
    AssertEqual "ValidateSetupData bad CNA", _
        True, (Len(mod_Validation.ValidateSetupData(setup)) > 0)

    ' ValidateReportData
    Dim rep As Object
    Set rep = CreateObject("Scripting.Dictionary")
    rep("KlientFK") = "999"
    rep("NazwaKlienta") = "Acme Sp. z o.o."
    rep("MiesiacObrotu") = "2026-05"
    AssertEqual "ValidateReportData OK", "", mod_Validation.ValidateReportData(rep)

    rep("NazwaKlienta") = "ab"
    AssertEqual "ValidateReportData short name", _
        True, (Len(mod_Validation.ValidateReportData(rep)) > 0)

    Debug.Print "----- Test_mod_Validation DONE -----"
End Sub

' ----- mod_MailSender -----------------------------------------------------

' NIE wysy³a maila - testuje tylko DetermineRecipient (czysta funkcja).
' Manipuluje EmailKierownika w Registry ¿eby sprawdziæ obie ga³êzie
' decision diamond, na koñcu przywraca oryginaln¹ wartoœæ.
Public Sub Test_mod_MailSender()
    Debug.Print "----- Test_mod_MailSender -----"

    If mod_UsersRegistrySync.GetUsersCount() = 0 Then
        Debug.Print "  [SKIP] Registry pusty - wymagany aktywny user"
        Exit Sub
    End If

    ' Backup oryginalnych wartoœci
    Dim origKierownika As String, origHandlowca As String
    origKierownika = CStr(mod_UsersRegistrySync.GetCurrentUserField("EmailKierownika"))
    origHandlowca = CStr(mod_UsersRegistrySync.GetCurrentUserField("EmailHandlowca"))

    ' HANDLOWIEC (kierownika != handlowca) -> mail do kierownika
    mod_UsersRegistrySync.SetCurrentUserField "EmailHandlowca", "handlowiec@firma.pl"
    mod_UsersRegistrySync.SetCurrentUserField "EmailKierownika", "kierownik@firma.pl"

    Dim r As Object
    Set r = mod_MailSender.DetermineRecipient()
    AssertEqual "Handlowiec.To = kierownik", "kierownik@firma.pl", CStr(r("To"))
    AssertEqual "Handlowiec.Subject ma 'akceptacji'", _
        True, (InStr(CStr(r("Subject")), "akceptacji") > 0)
    AssertEqual "Handlowiec.Body ma EmailBNC", _
        True, (InStr(CStr(r("Body")), CStr(mod_UsersRegistrySync.GetCurrentUserField("EmailBNC"))) > 0)

    ' KIEROWNIK (kierownika == handlowca) -> mail wprost do BNC
    mod_UsersRegistrySync.SetCurrentUserField "EmailKierownika", "handlowiec@firma.pl"

    Set r = mod_MailSender.DetermineRecipient()
    AssertEqual "Kierownik.To = EmailBNC", _
        CStr(mod_UsersRegistrySync.GetCurrentUserField("EmailBNC")), CStr(r("To"))
    AssertEqual "Kierownik.Subject NIE ma 'akceptacji'", _
        False, (InStr(CStr(r("Subject")), "akceptacji") > 0)

    ' Przywróæ oryginalne wartoœci
    mod_UsersRegistrySync.SetCurrentUserField "EmailHandlowca", origHandlowca
    mod_UsersRegistrySync.SetCurrentUserField "EmailKierownika", origKierownika

    Debug.Print "----- Test_mod_MailSender DONE -----"
End Sub

' ----- mod_Export ---------------------------------------------------------

Public Sub Test_mod_Export()
    Debug.Print "----- Test_mod_Export -----"

    Dim suggested As String
    suggested = mod_Export.GetSuggestedExportFileName()
    Debug.Print "  GetSuggestedExportFileName: " & suggested
    AssertEqual "Suggested startsWith BNC_Eksport_", _
        True, (Left$(suggested, 12) = "BNC_Eksport_")
    AssertEqual "Suggested endsWith .xlsx", _
        True, (Right$(suggested, 5) = ".xlsx")

    Dim folderPath As String
    folderPath = CStr(mod_UsersRegistrySync.GetCurrentUserField("CacheFolderPath"))
    If Len(folderPath) = 0 Then
        Debug.Print "  [SKIP] CacheFolderPath nie ustawiony - wymagany setup"
        Exit Sub
    End If

    Dim targetPath As String
    targetPath = mod_Utils.JoinPath(Environ("TEMP"), "BNC_Test_Export_" & _
                 Format(Now(), "yyyymmddhhnnss") & ".xlsx")

    mod_DataCacheSync.EnsureCacheFileExists

    Dim ok As Boolean
    ok = mod_Export.ExportDataCache(targetPath)
    AssertEqual "ExportDataCache success", True, ok
    AssertEqual "Target file exists after export", True, mod_Utils.FileExists(targetPath)

    On Error Resume Next
    Kill targetPath
    On Error GoTo 0

    Debug.Print "----- Test_mod_Export DONE -----"
End Sub

' ----- Multi-user Registry (M3.3, ADR-008) --------------------------------

' Test regresyjny dla multi-user flow: AddNewUser, GenerateUserID format,
' GetAllUsers, SwitchUser back, xlsx sync po AddNewUser.
'
' Testowy user oznaczony Imie="_TEST_" + CNA=999999 - pozostaje w Registry
' (test nie kasuje). Usuñ manualnie jeœli nie potrzeba historii.
Public Sub Test_MultiUser()
    Debug.Print "----- Test_MultiUser (M3.3) -----"

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("ws_UsersRegistry")
    On Error GoTo 0

    Dim origUserId As String
    Dim origUserCount As Long
    origUserId = mod_UsersRegistrySync.CurrentUserID()
    origUserCount = mod_UsersRegistrySync.GetUsersCount()
    Debug.Print "  [info] Stan wejsciowy: " & origUserCount & " userow, current=" & origUserId

    ' AddNewUser
    Dim testData As Object
    Set testData = CreateObject("Scripting.Dictionary")
    testData("Imie") = "_TEST_"
    testData("Nazwisko") = "MultiUser"
    testData("EmailHandlowca") = "test.multiuser@example.com"
    testData("CNA_HandlowcaID") = 999999
    testData("NrOddzialu") = "TEST"
    testData("EmailKierownika") = "test.multiuser@example.com"  ' = handlowca -> tryb kierownik
    testData("EmailBNC") = "test-bnc@example.com"
    testData("CacheFolderPath") = "C:\BNC_CacheFolder\"
    testData("DataRejestracji") = Now()
    testData("SetupCompleted") = True
    testData("DontShowSetupAgain") = False

    Dim newUserId As String
    newUserId = mod_UsersRegistrySync.AddNewUser(testData)
    Debug.Print "  AddNewUser -> " & newUserId

    AssertEqual "AddNewUser returns non-empty", True, (Len(newUserId) > 0)
    AssertEqual "UserID format UZYTKOWNIK_*_CNA999999", True, _
                (InStr(newUserId, "UZYTKOWNIK_") = 1 And InStr(newUserId, "_CNA999999") > 0)
    AssertEqual "GetUsersCount incremented", origUserCount + 1, mod_UsersRegistrySync.GetUsersCount()
    AssertEqual "CurrentUserID = newUserId (auto-switch)", newUserId, mod_UsersRegistrySync.CurrentUserID()

    AssertEqual "GetCurrentUserField.Imie = _TEST_", "_TEST_", _
                CStr(mod_UsersRegistrySync.GetCurrentUserField("Imie"))
    AssertEqual "GetCurrentUserField.CNA = 999999", 999999, _
                mod_UsersRegistrySync.GetCurrentUserField("CNA_HandlowcaID")
    AssertEqual "IsUserManager = True (email kierownika=handlowca)", True, _
                mod_UsersRegistrySync.IsUserManager()

    ' GetAllUsers zawiera nowego
    Dim users As Collection
    Set users = mod_UsersRegistrySync.GetAllUsers()
    Dim found As Boolean
    Dim u As Object
    For Each u In users
        If CStr(u("UserID")) = newUserId Then
            found = True
            AssertEqual "GetAllUsers new user Imie", "_TEST_", CStr(u("Imie"))
            AssertEqual "GetAllUsers new user CNA", 999999, u("CNA_HandlowcaID")
            Exit For
        End If
    Next u
    AssertEqual "GetAllUsers contains new UserID", True, found

    ' Registry xlsx sync
    Dim regCachePath As String
    regCachePath = mod_Utils.JoinPath( _
        CStr(mod_UsersRegistrySync.GetCurrentUserField("CacheFolderPath")), _
        "BNC_UsersRegistry.xlsx")
    AssertEqual "BNC_UsersRegistry.xlsx istnieje po AddNewUser", True, _
                mod_Utils.FileExists(regCachePath)

    ' SwitchUser back do poprzedniego (jeœli by³)
    If Len(origUserId) > 0 Then
        mod_UsersRegistrySync.SwitchUser origUserId
        AssertEqual "SwitchUser back: CurrentUserID", origUserId, mod_UsersRegistrySync.CurrentUserID()
    Else
        Debug.Print "  [info] SKIP SwitchUser back - brak previousUserId"
    End If

    Debug.Print "  [info] Testowy user " & newUserId & " ZOSTAJE w Registry."
    Debug.Print "         Aby usunac: otworz ws_UsersRegistry (Visible=Visible)"
    Debug.Print "         i usun wiersz gdzie Imie='_TEST_' + CNA=999999"

    Debug.Print "----- Test_MultiUser DONE -----"
End Sub

' ----- Helpery testów -----------------------------------------------------

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
