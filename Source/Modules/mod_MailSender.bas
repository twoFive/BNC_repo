Attribute VB_Name = "mod_MailSender"
Option Explicit

' ============================================================================
'  mod_MailSender
'  Serce logiki "kierownik vs handlowiec" (ADR-005). Decyduje o adresacie,
'  generuje plik tymczasowy xlsx w %TEMP% (ADR-004), wysyła mail przez
'  Outlook COM, zapisuje adresata do EmailRecipient (audit trail).
' ============================================================================

' Ol* constanty - definiowane lokalnie (nie wymaga referencji Outlook).
Private Const olMailItem As Long = 0

' ----- Public API ---------------------------------------------------------

' Główna funkcja: czyta pending z DataCache, generuje plik tymczasowy,
' decyduje adresata, wysyła, oznacza jako sent, sprząta plik tymczasowy.
Public Function SendBatch() As Boolean
    Dim tempFilePath As String
    Dim recipientInfo As Object
    Dim pending As Collection
    Dim sentIDs As Collection
    Dim record As Object

    On Error GoTo ErrorHandler

    Set pending = mod_DataCacheSync.GetPendingRecords()
    If pending.Count = 0 Then
        mod_Utils.LogInfo "SendBatch: brak pending zgloszen, exit."
        SendBatch = False
        Exit Function
    End If

    tempFilePath = GenerateTempFile(pending)
    mod_Utils.LogInfo "SendBatch: wygenerowany plik tymczasowy: " & tempFilePath

    Set recipientInfo = DetermineRecipient()
    mod_Utils.LogInfo "SendBatch: adresat = " & CStr(recipientInfo("To"))

    SendMailWithAttachment _
        recipient:=CStr(recipientInfo("To")), _
        subject:=CStr(recipientInfo("Subject")), _
        body:=CStr(recipientInfo("Body")), _
        attachmentPath:=tempFilePath

    Set sentIDs = New Collection
    For Each record In pending
        sentIDs.Add record("ReportID")
    Next record
    mod_DataCacheSync.MarkAsSent sentIDs, CStr(recipientInfo("To"))

    CleanupTempFile tempFilePath

    mod_Utils.LogInfo "SendBatch: pipeline OK, wyslano " & pending.Count & " zgloszen."
    SendBatch = True
    Exit Function

ErrorHandler:
    mod_Utils.LogError "mod_MailSender.SendBatch", Err.Number, Err.Description
    If Len(tempFilePath) > 0 Then CleanupTempFile tempFilePath
    SendBatch = False
End Function

' Decision diamond - kierownik vs handlowiec. Public dla testowalności
' (czysta funkcja bez side effects).
'
' Convention over configuration: jeżeli EmailKierownika == EmailHandlowca,
' user jest kierownikiem -> mail wprost do BNC. W p.p. -> mail do kierownika.
Public Function DetermineRecipient() As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")

    Dim emailHandlowca As String
    Dim emailKierownika As String
    Dim emailBNC As String
    emailHandlowca = CStr(mod_UsersRegistrySync.GetCurrentUserField("EmailHandlowca"))
    emailKierownika = CStr(mod_UsersRegistrySync.GetCurrentUserField("EmailKierownika"))
    emailBNC = CStr(mod_UsersRegistrySync.GetCurrentUserField("EmailBNC"))

    Dim dateTag As String
    dateTag = Format(Now(), "yyyy-mm-dd")

    If mod_UsersRegistrySync.IsUserManager() Then
        ' KIEROWNIK - wprost do BNC.
        result("To") = emailBNC
        result("Subject") = "Wniosek BNC - " & dateTag
        result("Body") = "Dzień dobry," & vbCrLf & vbCrLf & _
            "W załączeniu wniosek BNC. Proszę o weryfikację." & vbCrLf & vbCrLf & _
            "Pozdrawiam," & vbCrLf & _
            CStr(mod_UsersRegistrySync.GetCurrentUserField("Imie")) & " " & _
            CStr(mod_UsersRegistrySync.GetCurrentUserField("Nazwisko"))
    Else
        ' HANDLOWIEC - do kierownika z prośbą o przekazanie do BNC.
        result("To") = emailKierownika
        result("Subject") = "Wniosek BNC do akceptacji - " & dateTag
        result("Body") = "Dzień dobry," & vbCrLf & vbCrLf & _
            "W załączeniu wniosek BNC. Proszę o weryfikację i przekazanie do " & _
            emailBNC & "." & vbCrLf & vbCrLf & _
            "Pozdrawiam," & vbCrLf & _
            CStr(mod_UsersRegistrySync.GetCurrentUserField("Imie")) & " " & _
            CStr(mod_UsersRegistrySync.GetCurrentUserField("Nazwisko"))
    End If

    Set DetermineRecipient = result
End Function

' ----- Private - plik tymczasowy + Outlook COM ----------------------------

' Tworzy plik xlsx w %TEMP% z aktualnym batchem (ADR-004).
' Returns: pełna ścieżka do utworzonego pliku.
Private Function GenerateTempFile(records As Collection) As String
    Dim tempFolder As String
    Dim fileName As String
    Dim fullPath As String

    tempFolder = Environ("TEMP")
    fileName = "BNC_Wniosek_" & Format(Now(), "yyyymmdd_hhnnss") & ".xlsx"
    fullPath = mod_Utils.JoinPath(tempFolder, fileName)

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim restoreScreen As Boolean
    Dim restoreAlerts As Boolean

    Application.ScreenUpdating = False
    restoreScreen = True
    Application.DisplayAlerts = False
    restoreAlerts = True

    Set wb = Workbooks.Add
    Set ws = wb.Worksheets(1)
    ws.Name = "BNC_Wniosek"

    ' Nagłówki wniosku - subset z DataCache (bez pól wewnętrznych typu Status).
    ws.Cells(1, 1).Value = "ReportID"
    ws.Cells(1, 2).Value = "KlientFK"
    ws.Cells(1, 3).Value = "NazwaKlienta"
    ws.Cells(1, 4).Value = "CNA"
    ws.Cells(1, 5).Value = "NrOddzialu"
    ws.Cells(1, 6).Value = "MiesiacObrotu"
    ws.Cells(1, 7).Value = "CreatedTimestamp"

    Dim r As Long
    Dim record As Object
    r = 2
    For Each record In records
        ws.Cells(r, 1).Value = record("ReportID")
        ws.Cells(r, 2).Value = record("KlientFK")
        ws.Cells(r, 3).Value = record("NazwaKlienta")
        ws.Cells(r, 4).Value = record("CNA_HandlowcaID")
        ws.Cells(r, 5).Value = record("NrOddzialu")
        ws.Cells(r, 6).Value = record("MiesiacObrotu")
        ws.Cells(r, 7).Value = record("CreatedTimestamp")
        r = r + 1
    Next record

    ws.Columns.AutoFit

    wb.SaveAs Filename:=fullPath, FileFormat:=xlOpenXMLWorkbook
    wb.Close SaveChanges:=False
    Set wb = Nothing

    If restoreAlerts Then Application.DisplayAlerts = True
    If restoreScreen Then Application.ScreenUpdating = True

    GenerateTempFile = fullPath
End Function

' Wysyłka maila przez Outlook COM. Wymaga zaufanego dostępu do Outlook
' (Trust Center -> "Trust access to the Outlook object model" lub polityka IT).
Private Sub SendMailWithAttachment(recipient As String, _
                                    subject As String, _
                                    body As String, _
                                    attachmentPath As String)
    Dim outlookApp As Object
    Dim mailItem As Object

    Set outlookApp = CreateObject("Outlook.Application")
    Set mailItem = outlookApp.CreateItem(olMailItem)

    With mailItem
        .To = recipient
        .Subject = subject
        .Body = body
        .Attachments.Add attachmentPath
        .Send
    End With

    Set mailItem = Nothing
    Set outlookApp = Nothing
End Sub

' Best-effort delete pliku tymczasowego - błędy ignorowane.
Private Sub CleanupTempFile(filePath As String)
    On Error Resume Next
    If Len(filePath) = 0 Then Exit Sub

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(filePath) Then
        fso.DeleteFile filePath, True  ' True = force, ignoruj read-only
    End If
End Sub
