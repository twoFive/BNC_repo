Attribute VB_Name = "mod_Export"
Option Explicit

' ============================================================================
'  mod_Export - literal copy BNC_DataCache.xlsx do wybranej lokalizacji.
'  "Literal copy" znaczy: dokladnie ten sam plik xlsx co w cache - nie
'  generujemy nic specjalnego, kopiujemy bit-perfect. Plik backupu w
'  CacheFolderPath jest jedynym zrodlem prawdy o historii.
'
'  Patrz: BNC_Sender_PlanWdrozenia_FazaA.md (M5.1)
' ============================================================================

Private Const SOURCE_FILE_NAME As String = "BNC_DataCache.xlsx"

' ============================================================================
'  Public API
' ============================================================================

' Kopiuje BNC_DataCache.xlsx z CacheFolderPath do targetPath (overwrite=True).
' Returns: True jesli sukces, False jesli zrodlo nie istnieje albo bald I/O.
Public Function ExportDataCache(targetPath As String) As Boolean
    Dim sourcePath As String
    sourcePath = GetSourcePath()

    If Not mod_Utils.FileExists(sourcePath) Then
        mod_Utils.LogError "mod_Export.ExportDataCache", 0, _
            "Plik zrodlowy nie istnieje: " & sourcePath
        ExportDataCache = False
        Exit Function
    End If

    On Error GoTo ErrorHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile sourcePath, targetPath, True  ' True = overwrite

    mod_Utils.LogInfo "ExportDataCache: " & sourcePath & " -> " & targetPath
    ExportDataCache = True
    Exit Function

ErrorHandler:
    mod_Utils.LogError "mod_Export.ExportDataCache", Err.Number, Err.Description
    ExportDataCache = False
End Function

' Sugerowana nazwa pliku eksportu, np. "BNC_Eksport_Kowalski_2026-05-06.xlsx".
' Sanityzuje nazwisko - usuwa spacje i znaki specjalne dla bezpiecznej nazwy.
Public Function GetSuggestedExportFileName() As String
    Dim nazwisko As String
    nazwisko = SanitizeFileName(CStr(mod_UserCacheSync.GetUserField("Nazwisko")))
    If Len(nazwisko) = 0 Then nazwisko = "User"

    GetSuggestedExportFileName = "BNC_Eksport_" & nazwisko & "_" & _
                                 Format(Now(), "yyyy-mm-dd") & ".xlsx"
End Function

' ============================================================================
'  Private
' ============================================================================

Private Function GetSourcePath() As String
    Dim folderPath As String
    folderPath = CStr(mod_UserCacheSync.GetUserField("CacheFolderPath"))
    GetSourcePath = mod_Utils.JoinPath(folderPath, SOURCE_FILE_NAME)
End Function

' Zamienia wszystkie znaki niedozwolone w nazwach plikow Windows na "_".
Private Function SanitizeFileName(s As String) As String
    Dim invalid As String
    invalid = "\/:*?""<>| "
    Dim result As String
    result = s

    Dim i As Long, ch As String
    For i = 1 To Len(invalid)
        ch = Mid$(invalid, i, 1)
        result = Replace(result, ch, "_")
    Next i

    SanitizeFileName = Trim$(result)
End Function
