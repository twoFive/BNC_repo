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
