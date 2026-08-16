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
    Test_IsFormOpen
    Test_mod_AppStateSync
    Test_mod_DataCacheSync
    Test_mod_Validation
    Test_mod_MailSender
    Test_mod_Export
    Test_MultiUser
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

' ----- IsFormOpen (regression guard: zombie hidden forms) ------------------

' Regression test dla bugfix 2026-08-10: IsFormOpen sprawdza .Visible, nie
' tylko istnienie w UserForms collection. Bez tego testu, przyszly refactor
' (np. cofniecie do "For Each f In UserForms: If f.Name = ...: True") wprowadzi
' bug ktory zablokuje re-Show po Me.Hide (zombie forms).
'
' UWAGA: NIE testujemy visible=True case bo wymagaloby to Show vbModal ktory
' blokuje wykonanie tego suba (modal loop). Manual testing dla visible case.
' Testujemy tylko sciezki bez modal: nieistniejacy form + loaded-but-hidden.
Public Sub Test_IsFormOpen()
    Debug.Print "----- Test_IsFormOpen -----"

    ' Nieistniejacy form -> False
    AssertEqual "IsFormOpen nonexistent", False, _
        mod_Utils.IsFormOpen("frm_NoSuchForm_XYZ")

    ' Loaded (w UserForms collection) ale niewidoczny -> False (zombie check).
    ' To glowna sciezka regresji - Load bez Show emuluje stan po Me.Hide.
    Load frm_Setup
    AssertEqual "IsFormOpen loaded-but-hidden (zombie)", False, _
        mod_Utils.IsFormOpen("frm_Setup")

    ' Cleanup - usun z UserForms
    Unload frm_Setup

    ' Po Unload -> ponownie False (bo forma znikla z kolekcji)
    AssertEqual "IsFormOpen after Unload", False, _
        mod_Utils.IsFormOpen("frm_Setup")

    Debug.Print "----- Test_IsFormOpen DONE -----"
End Sub

' ----- mod_AppStateSync (post-ADR-009 replaces Test_mod_UserCacheSync) -----

' UWAGA: ten test pisze i odczytuje z ws_AppState. Operuje na tymczasowym
' kluczu "_TEST_RoundTrip_" - na koncu kasuje. NIE dotyka _CurrentUserID
' zeby nie zepsuc sesji.
Public Sub Test_mod_AppStateSync()
    Debug.Print "----- Test_mod_AppStateSync -----"

    ' Round-trip write/read
    Const TEST_KEY As String = "_TEST_RoundTrip_"
    Const TEST_VAL As String = "smoke_test_value_42"

    mod_AppStateSync.SetAppValue TEST_KEY, TEST_VAL
    AssertEqual "SetAppValue round-trip", TEST_VAL, _
        CStr(mod_AppStateSync.GetAppValue(TEST_KEY))

    ' Missing key returns ""
    AssertEqual "GetAppValue missing key returns empty", "", _
        CStr(mod_AppStateSync.GetAppValue("_TEST_NonExistent_"))

    ' _CurrentUserID present (assuming session has active user)
    Dim currentId As String
    currentId = CStr(mod_AppStateSync.GetAppValue("_CurrentUserID"))
    Debug.Print "  [info] _CurrentUserID = '" & currentId & "'"
    AssertEqual "_CurrentUserID non-empty (session ma aktywnego usera)", True, _
        (Len(currentId) > 0)

    ' EnsureAppStateSheet is idempotent
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

    ' --- AppendRecord (schema v2: bez Fields, MiesiacObrotu zamiast MiesiacZgloszenia) ---
    Dim newRecord As Object
    Set newRecord = CreateObject("Scripting.Dictionary")
    newRecord("KlientFK") = 99999
    newRecord("NazwaKlienta") = "_TEST_Klient_DoUsuniecia_"
    newRecord("MiesiacObrotu") = Format(mod_Utils.GetCurrentMonthYear(), "yyyy-mm")

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
    pendingRecord("MiesiacObrotu") = Format(mod_Utils.GetCurrentMonthYear(), "yyyy-mm")

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

    ' --- ValidateReportData (schema v2: bez Fields, MiesiacObrotu) ---
    Dim rep As Object
    Set rep = CreateObject("Scripting.Dictionary")
    rep("KlientFK") = "999"
    rep("NazwaKlienta") = "Acme Sp. z o.o."
    rep("MiesiacObrotu") = "2026-05"
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

    If mod_UsersRegistrySync.GetUsersCount() = 0 Then
        Debug.Print "  [SKIP] Registry pusty - wymagany aktywny user do test scenariuszy"
        Exit Sub
    End If

    ' Backup oryginalnych wartosci (post-ADR-009: przez Registry API)
    Dim origKierownika As String, origHandlowca As String
    origKierownika = CStr(mod_UsersRegistrySync.GetCurrentUserField("EmailKierownika"))
    origHandlowca = CStr(mod_UsersRegistrySync.GetCurrentUserField("EmailHandlowca"))

    ' Scenariusz 1: HANDLOWIEC (kierownika != handlowca) -> mail do kierownika
    mod_UsersRegistrySync.SetCurrentUserField "EmailHandlowca", "handlowiec@firma.pl"
    mod_UsersRegistrySync.SetCurrentUserField "EmailKierownika", "kierownik@firma.pl"

    Dim r As Object
    Set r = mod_MailSender.DetermineRecipient()
    AssertEqual "Handlowiec.To = kierownik", "kierownik@firma.pl", CStr(r("To"))
    AssertEqual "Handlowiec.Subject ma 'akceptacji'", _
        True, (InStr(CStr(r("Subject")), "akceptacji") > 0)
    AssertEqual "Handlowiec.Body ma EmailBNC", _
        True, (InStr(CStr(r("Body")), CStr(mod_UsersRegistrySync.GetCurrentUserField("EmailBNC"))) > 0)

    ' Scenariusz 2: KIEROWNIK (kierownika == handlowca) -> mail wprost do BNC
    mod_UsersRegistrySync.SetCurrentUserField "EmailKierownika", "handlowiec@firma.pl"

    Set r = mod_MailSender.DetermineRecipient()
    AssertEqual "Kierownik.To = EmailBNC", _
        CStr(mod_UsersRegistrySync.GetCurrentUserField("EmailBNC")), CStr(r("To"))
    AssertEqual "Kierownik.Subject NIE ma 'akceptacji'", _
        False, (InStr(CStr(r("Subject")), "akceptacji") > 0)

    ' Przywroc oryginalne wartosci
    mod_UsersRegistrySync.SetCurrentUserField "EmailHandlowca", origHandlowca
    mod_UsersRegistrySync.SetCurrentUserField "EmailKierownika", origKierownika

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
    folderPath = CStr(mod_UsersRegistrySync.GetCurrentUserField("CacheFolderPath"))
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

' ----- Multi-user Registry (M3.3, ADR-008) --------------------------------

' Test regresyjny dla multi-user flow. Sprawdza:
' - AddNewUser -> nowy wiersz w Registry + switch na nowego usera
' - GenerateUserID format: UZYTKOWNIK_<N>_CNA<cna>
' - GetAllUsers zwraca nowego usera
' - SwitchUser back do poprzedniego -> UserCache poprawnie przywrocony
'
' UWAGA: test dodaje wiersz do ws_UsersRegistry i NIE kasuje go automatycznie.
' Test rekordy oznaczone Imie="_TEST_" + CNA=999999 - latwe do rozpoznania
' i manualnej usuwki. Manualnie usun po testach jesli nie potrzeba historii.
Public Sub Test_MultiUser()
    Debug.Print "----- Test_MultiUser (M3.3) -----"

    ' Registry musi istniec - jesli nie, AddNewUser go stworzy
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("ws_UsersRegistry")
    On Error GoTo 0

    ' --- Backup aktualnego stanu ---
    Dim origUserId As String
    Dim origUserCount As Long
    origUserId = mod_UsersRegistrySync.CurrentUserID()
    origUserCount = mod_UsersRegistrySync.GetUsersCount()
    Debug.Print "  [info] Stan wejsciowy: " & origUserCount & " userow, current=" & origUserId

    ' --- AddNewUser: dodaje testowego usera ---
    Dim testData As Object
    Set testData = CreateObject("Scripting.Dictionary")
    testData("Imie") = "_TEST_"
    testData("Nazwisko") = "MultiUser"
    testData("EmailHandlowca") = "test.multiuser@example.com"
    testData("CNA_HandlowcaID") = 999999
    testData("NrOddzialu") = "TEST"
    testData("EmailKierownika") = "test.multiuser@example.com"  ' = handlowca → tryb kierownik
    testData("EmailBNC") = "test-bnc@example.com"
    testData("CacheFolderPath") = "C:\BNC_CacheFolder\"
    testData("DataRejestracji") = Now()
    testData("SetupCompleted") = True
    testData("DontShowSetupAgain") = False

    Dim newUserId As String
    newUserId = mod_UsersRegistrySync.AddNewUser(testData)
    Debug.Print "  AddNewUser -> " & newUserId

    ' --- Assertions ---
    AssertEqual "AddNewUser returns non-empty", True, (Len(newUserId) > 0)
    AssertEqual "UserID format UZYTKOWNIK_*_CNA999999", True, _
                (InStr(newUserId, "UZYTKOWNIK_") = 1 And InStr(newUserId, "_CNA999999") > 0)
    AssertEqual "GetUsersCount incremented", origUserCount + 1, mod_UsersRegistrySync.GetUsersCount()
    AssertEqual "CurrentUserID = newUserId (auto-switch)", newUserId, mod_UsersRegistrySync.CurrentUserID()

    ' --- Current user data (post-ADR-009: bezposredni odczyt z Registry) ---
    AssertEqual "GetCurrentUserField.Imie = _TEST_", "_TEST_", _
                CStr(mod_UsersRegistrySync.GetCurrentUserField("Imie"))
    AssertEqual "GetCurrentUserField.CNA = 999999", 999999, _
                mod_UsersRegistrySync.GetCurrentUserField("CNA_HandlowcaID")
    AssertEqual "IsUserManager = True (email kierownika=handlowca)", True, _
                mod_UsersRegistrySync.IsUserManager()

    ' --- GetAllUsers zawiera nowego ---
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

    ' --- Registry xlsx sync (post-M3.3 symetria z DataCache) ---
    Dim regCachePath As String
    regCachePath = mod_Utils.JoinPath( _
        CStr(mod_UsersRegistrySync.GetCurrentUserField("CacheFolderPath")), _
        "BNC_UsersRegistry.xlsx")
    AssertEqual "BNC_UsersRegistry.xlsx istnieje po AddNewUser", True, _
                mod_Utils.FileExists(regCachePath)

    ' --- SwitchUser back do poprzedniego (jesli byl) ---
    If Len(origUserId) > 0 Then
        mod_UsersRegistrySync.SwitchUser origUserId
        AssertEqual "SwitchUser back: CurrentUserID", origUserId, mod_UsersRegistrySync.CurrentUserID()
    Else
        Debug.Print "  [info] SKIP SwitchUser back - brak previousUserId (Test w pustym Registry)"
    End If

    ' --- Note: testowy user PONIESIE w Registry, oznaczony _TEST_ + CNA=999999 ---
    Debug.Print "  [info] Testowy user " & newUserId & " ZOSTAJE w Registry."
    Debug.Print "         Aby usunac: otworz ws_UsersRegistry (Visible=Visible tymczasowo)"
    Debug.Print "         i usun wiersz gdzie Imie='_TEST_' + CNA=999999"

    Debug.Print "----- Test_MultiUser DONE -----"
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
