Attribute VB_Name = "mod_MailSender"
Option Explicit

' ============================================================================
'  mod_MailSender - serce logiki "kierownik vs handlowiec" (ADR-005).
'  Jedyny modul, ktory:
'  - decyduje o adresacie wysylki (decision diamond)
'  - generuje plik tymczasowy xlsx w %TEMP% z aktualnym batchem (ADR-004)
'  - wysyla mail przez Outlook COM
'  - zapisuje rzeczywistego adresata do EmailRecipient (audit trail)
'
'  Public API:
'    SendBatch()           -> Boolean   (glowna funkcja, wywolana z frm_Main)
'    DetermineRecipient()  -> Object    (Dictionary z To/Subject/Body, public
'                                        dla testow + przegladnosci)
'  Patrz: BNC_Sender_PlanWdrozenia_FazaA.md (M4.1)
'         doc_v2/extracted/03_data_flow.md (Flow B z decision diamond)
' ============================================================================

' Olxxx constanty - definiowane lokalnie (nie wymaga referencji Outlook).
Private Const olMailItem As Long = 0

' ============================================================================
'  Public API
' ============================================================================

' Glowna funkcja: czyta pending z DataCache, generuje plik tymczasowy,
' decyduje adresata, wysyla, oznacza jako sent, sprzata plik tymczasowy.
' Returns: True jesli pelny pipeline OK, False jesli error (z log + cleanup).
Public Function SendBatch() As Boolean
    Dim tempFilePath As String
    Dim recipientInfo As Object
    Dim pending As Collection
    Dim sentIDs As Collection
    Dim record As Object

    On Error GoTo ErrorHandler

    ' 1. Pobierz pending
    Set pending = mod_DataCacheSync.GetPendingRecords()
    If pending.Count = 0 Then
        mod_Utils.LogInfo "SendBatch: brak pending zgloszen, exit."
        SendBatch = False
        Exit Function
    End If

    ' 2. Plik tymczasowy w %TEMP%
    tempFilePath = GenerateTempFile(pending)
    mod_Utils.LogInfo "SendBatch: wygenerowany plik tymczasowy: " & tempFilePath

    ' 3. Decision diamond - kierownik vs handlowiec
    Set recipientInfo = DetermineRecipient()
    mod_Utils.LogInfo "SendBatch: adresat = " & CStr(recipientInfo("To"))

    ' 4. Wyslij mail przez Outlook COM
    SendMailWithAttachment _
        recipient:=CStr(recipientInfo("To")), _
        subject:=CStr(recipientInfo("Subject")), _
        body:=CStr(recipientInfo("Body")), _
        attachmentPath:=tempFilePath

    ' 5. UPDATE statusu w ws_DataCache
    Set sentIDs = New Collection
    For Each record In pending
        sentIDs.Add record("ReportID")
    Next record
    mod_DataCacheSync.MarkAsSent sentIDs, CStr(recipientInfo("To"))

    ' 6. Cleanup plik tymczasowy
    CleanupTempFile tempFilePath

    mod_Utils.LogInfo "SendBatch: pipeline OK, wyslano " & pending.Count & " zgloszen."
    SendBatch = True
    Exit Function

ErrorHandler:
    mod_Utils.LogError "mod_MailSender.SendBatch", Err.Number, Err.Description
    ' Cleanup nawet po bledzie - plik nie powinien zostac na dysku.
    If Len(tempFilePath) > 0 Then CleanupTempFile tempFilePath
    SendBatch = False
End Function

' Decision diamond - centralna logika kierownik vs handlowiec.
' Public dla testowalnosci - to czysta funkcja bez side effects.
' Returns Scripting.Dictionary z polami: To, Subject, Body.
'
' Convention over configuration: jezeli EmailKierownika == EmailHandlowca,
' user jest kierownikiem -> mail wprost do BNC.
' W przeciwnym razie -> mail do kierownika z prosba o przekazanie.
Public Function DetermineRecipient() As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")

    Dim emailHandlowca As String
    Dim emailKierownika As String
    Dim emailBNC As String
    emailHandlowca = CStr(mod_UserCacheSync.GetUserField("EmailHandlowca"))
    emailKierownika = CStr(mod_UserCacheSync.GetUserField("EmailKierownika"))
    emailBNC = CStr(mod_UserCacheSync.GetUserField("EmailBNC"))

    Dim dateTag As String
    dateTag = Format(Now(), "yyyy-mm-dd")

    If mod_UserCacheSync.IsUserManager() Then
        ' KIEROWNIK - wprost do BNC.
        result("To") = emailBNC
        result("Subject") = "Wniosek BNC - " & dateTag
        result("Body") = "Dzien dobry," & vbCrLf & vbCrLf & _
            "W zalaczeniu wniosek BNC. Prosze o weryfikacje." & vbCrLf & vbCrLf & _
            "Pozdrawiam," & vbCrLf & _
            CStr(mod_UserCacheSync.GetUserField("Imie")) & " " & _
            CStr(mod_UserCacheSync.GetUserField("Nazwisko"))
    Else
        ' HANDLOWIEC - do kierownika z prosba o przekazanie do BNC.
        result("To") = emailKierownika
        result("Subject") = "Wniosek BNC do akceptacji - " & dateTag
        result("Body") = "Dzien dobry," & vbCrLf & vbCrLf & _
            "W zalaczeniu wniosek BNC. Prosze o weryfikacje i przekazanie do " & _
            emailBNC & "." & vbCrLf & vbCrLf & _
            "Pozdrawiam," & vbCrLf & _
            CStr(mod_UserCacheSync.GetUserField("Imie")) & " " & _
            CStr(mod_UserCacheSync.GetUserField("Nazwisko"))
    End If

    Set DetermineRecipient = result
End Function

' ============================================================================
'  Private - plik tymczasowy + Outlook COM
' ============================================================================

' Tworzy plik xlsx w %TEMP% z aktualnym batchem (ADR-004).
' Returns: pelna sciezka do utworzonego pliku.
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

    ' Naglowki wlasciwe dla wniosku (subset z DataCache, bez Status/Recipient
    ' bo to sa pola wewnetrzne aplikacji - BNC dostaje tylko dane biznesowe).
    ws.Cells(1, 1).Value = "ReportID"
    ws.Cells(1, 2).Value = "KlientFK"
    ws.Cells(1, 3).Value = "NazwaKlienta"
    ws.Cells(1, 4).Value = "CNA_HandlowcaID"
    ws.Cells(1, 5).Value = "NrOddzialu"
    ws.Cells(1, 6).Value = "MiesiacZgloszenia"
    ws.Cells(1, 7).Value = "Fields"
    ws.Cells(1, 8).Value = "CreatedTimestamp"

    Dim r As Long
    Dim record As Object
    r = 2
    For Each record In records
        ws.Cells(r, 1).Value = record("ReportID")
        ws.Cells(r, 2).Value = record("KlientFK")
        ws.Cells(r, 3).Value = record("NazwaKlienta")
        ws.Cells(r, 4).Value = record("CNA_HandlowcaID")
        ws.Cells(r, 5).Value = record("NrOddzialu")
        ws.Cells(r, 6).Value = record("MiesiacZgloszenia")
        ws.Cells(r, 7).Value = record("Fields")
        ws.Cells(r, 8).Value = record("CreatedTimestamp")
        r = r + 1
    Next record

    ' Auto-fit dla czytelnosci.
    ws.Columns.AutoFit

    wb.SaveAs Filename:=fullPath, FileFormat:=xlOpenXMLWorkbook
    wb.Close SaveChanges:=False
    Set wb = Nothing

    If restoreAlerts Then Application.DisplayAlerts = True
    If restoreScreen Then Application.ScreenUpdating = True

    GenerateTempFile = fullPath
End Function

' Wysylka maila przez Outlook COM. Wymaga zaufanego dostepu do Outlook
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

' Bezpieczne usuwanie pliku tymczasowego. Best-effort - bledy zignorowane.
Private Sub CleanupTempFile(filePath As String)
    On Error Resume Next
    If Len(filePath) = 0 Then Exit Sub

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(filePath) Then
        fso.DeleteFile filePath, True  ' True = force, ignoruj read-only
    End If
End Sub
