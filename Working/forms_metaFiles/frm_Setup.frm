VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Setup 
   Caption         =   "BNC_Sender — Konfiguracja wstêpna"
   ClientHeight    =   10632
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7488
   OleObjectBlob   =   "frm_Setup.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_Setup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ----- Hardcoded constants (polityka projektowa, ADR-003) ------------------
' EmailBNC i CacheFolderPath sa staloymi, niedostepnymi userowi do edycji.
' Ustawiane w Initialize jako bezpiecznik na wypadek gdyby ktos zmienil
' Locked=False w designerze.
Private Const HARDCODED_EMAIL_BNC As String = "gitara101warszawa@gmail.com"

Private Const HARDCODED_CACHE_FOLDER As String = "C:\BNC_CacheFolder\"

' ----- Lifecycle -----------------------------------------------------------

' Post-ADR-009 refactor: frm_Setup zawsze tworzy NOWEGO usera (nigdy nie
' edytuje istniejacego). Dawny prefill code dla 'edit mode' usuniety -
' byl dead code od M3.3 (AddNewUser wymusza pusty formularz).
' Edycja profilu = future feature (poza zakresem Fazy A) - patrz ADR-008.
Private Sub UserForm_Initialize()
    ' Pola hardcoded - polityka projektowa (ADR-003).
    txt_EmailBNC.text = HARDCODED_EMAIL_BNC
    txt_EmailBNC.Locked = True
    txt_CacheFolderPath.text = HARDCODED_CACHE_FOLDER
    txt_CacheFolderPath.Locked = True
End Sub

' ----- Buttons -------------------------------------------------------------

' Samouczek - 5-page wizard (M6/M7). Modal blokuje frm_Setup do zamkniecia
' tutoriala - user nie zgubi kontekstu setupu.
' Patrz: frm_Tutorial.LAYOUT.md + frm_Tutorial.TEXT.md
Private Sub btn_ShowTutorial_Click()
    frm_Tutorial.Show vbModal
End Sub

' Tworzy folder cache na dysku zgodnie z polityka projektowa.
' Wywolanie EnsureFolderExists jest idempotentne - jesli folder juz istnieje,
' nic sie nie dzieje (i to jest OK, user dostaje neutralne potwierdzenie).
Private Sub btn_CreateCacheFolder_Click()
    Dim path As String
    path = Trim$(txt_CacheFolderPath.text)

    On Error GoTo CreateError

    Dim alreadyExisted As Boolean
    alreadyExisted = mod_Utils.FolderExists(path)

    mod_Utils.EnsureFolderExists path

    If alreadyExisted Then
        MsgBox "Folder cache juz istnial:" & vbCrLf & path, _
               vbInformation, "Folder cache"
    Else
        MsgBox "Folder cache utworzony:" & vbCrLf & path, _
               vbInformation, "Folder cache"
    End If
    Exit Sub

CreateError:
    mod_Utils.LogError "frm_Setup.btn_CreateCacheFolder_Click", _
                       Err.Number, Err.Description
    MsgBox "Nie udalo sie utworzyc folderu cache:" & vbCrLf & path & vbCrLf & vbCrLf & _
           "Blad: " & Err.Description, vbExclamation, "Blad folderu cache"
End Sub

Private Sub btn_Save_Click()
    ' 1. Zbierz dane z pol (EmailBNC i CacheFolderPath sa hardcoded).
    Dim userData As Object
    Set userData = CreateObject("Scripting.Dictionary")
    userData("Imie") = Trim$(txt_Imie.text)
    userData("Nazwisko") = Trim$(txt_Nazwisko.text)
    userData("EmailHandlowca") = Trim$(txt_EmailHandlowca.text)
    userData("CNA_HandlowcaID") = Trim$(txt_CNA.text)
    userData("NrOddzialu") = Trim$(txt_NrOddzialu.text)
    userData("EmailKierownika") = Trim$(txt_EmailKierownika.text)
    userData("EmailBNC") = HARDCODED_EMAIL_BNC
    userData("CacheFolderPath") = HARDCODED_CACHE_FOLDER
    userData("DataRejestracji") = Now()
    userData("SetupCompleted") = True
    userData("DontShowSetupAgain") = CBool(chk_DontShowSetupAgain.value)

    ' 2. Walidacja
    Dim errMsg As String
    errMsg = mod_Validation.ValidateSetupData(userData)
    If Len(errMsg) > 0 Then
        MsgBox errMsg, vbExclamation, "Blad walidacji"
        Exit Sub
    End If

    ' 3. Folder cache - jesli nie istnieje, utworz (defensywnie, gdyby user
    '    zapomnial kliknac btn_CreateCacheFolder).
    On Error GoTo FolderError
    mod_Utils.EnsureFolderExists CStr(userData("CacheFolderPath"))
    On Error GoTo 0

    ' 4. Dodaj nowego usera do Registry (M3.3) + zapis do UserCache jako
    '    aktywnego + auto-sync do xlsx. AddNewUser wywolywane zawsze -
    '    frm_Setup jest tylko flow "dodania nowego usera" (pierwszy raz
    '    LUB z frm_UserPicker.btn_AddNew). Edycja istniejacego usera
    '    to inny scenariusz, poza zakresem M3.3.
    Dim newUserId As String
    newUserId = mod_UsersRegistrySync.AddNewUser(userData)

    ' 5. Auto-recreate plikow cache (DataCache jeszcze nie istnieje przy pierwszym setupie)
    mod_DataCacheSync.EnsureCacheFileExists

    mod_Utils.LogInfo "Setup zakonczony - user=" & CStr(userData("EmailHandlowca")) & _
                     " · UserID=" & newUserId

    ' 6. Zamknij setup, otworz frm_Main.
    Me.Hide
    frm_Main.Show
    Exit Sub

FolderError:
    MsgBox "Nie udalo sie utworzyc folderu cache:" & vbCrLf & _
           CStr(userData("CacheFolderPath")) & vbCrLf & vbCrLf & _
           "Blad: " & Err.Description, vbExclamation, "Blad folderu cache"
End Sub

Private Sub btn_Cancel_Click()
    If MsgBox("Czy na pewno przerwac konfiguracje? Aplikacja nie uruchomi sie poprawnie.", _
              vbYesNo + vbQuestion, "Potwierdzenie") = vbYes Then
        Me.Hide
    End If
End Sub

' Helper GetTutorialText() zostal usuniety w M2.2 - tresc samouczka bedzie
' w code-behind frm_Tutorial po jego implementacji (M6/M7).


