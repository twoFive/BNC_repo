# SECURITY_OPTIONS - opcje bezpieczenstwa BNC_Sender

> Katalog opcji zabezpieczen aplikacji, z metodami implementacji i tradeoff'ami.
> Zero PL diakrytykow w kodzie snippets (patrz feedback_vba_encoding_cp1250).

---

## Stan obecny (juz wdrozone)

| Warstwa | Implementacja | Plik |
|---|---|---|
| Ribbon hidden | `SHOW.TOOLBAR("Ribbon",False)` w Workbook_Open | ThisWorkbook |
| Full screen | `Application.DisplayFullScreen = True` | ThisWorkbook |
| Bez formula bar / headings / gridlines / tabs | Application/ActiveWindow properties | ThisWorkbook |
| Cell locked + no selection | `Cells.Locked=True + EnableSelection=xlNoSelection + Protect UserInterfaceOnly=True` | ThisWorkbook.LockLandingPage |
| Scroll ograniczony do A1:H23 | `ScrollArea = "A1:H23"` | ThisWorkbook.LockLandingPage |
| Landing area only visible | Hide cols I:XFD + rows 24:1048576 (MANUAL) | sh_LandingPage |
| Buttons non-modifiable | `DrawingObjects:=True` w Protect | ThisWorkbook.LockLandingPage |
| VBA writes mimo protekcji | `UserInterfaceOnly:=True` | ThisWorkbook.LockLandingPage |
| Restore chrome przy close | Workbook_BeforeClose | ThisWorkbook |
| Save prompt suppress (commented) | `' ThisWorkbook.Saved = True` | ThisWorkbook.Workbook_BeforeClose |

---

## Opcja 1: Block otwieranie innych plikow Excel podczas sesji BNC_Sender

**Cel**: user nie moze otworzyc rownolegle innych xlsx/xlsm (koncentracja na app).

### Wariant A - Application-level event handler (rekomendowane)

Wymagane: WithEvents klasa + hook w ThisWorkbook. Reaguje na kazde otwarcie
workbooka w tym samym Excel instance.

**Setup**:

1. Insert -> **Class Module** w VBE, Name = `cls_AppEvents`
2. Wklej:
```vba
Option Explicit

Public WithEvents App As Application

Private Sub App_WorkbookOpen(ByVal Wb As Workbook)
    ' Ignoruj otwieranie samego BNC_Sender (edge case przy reopen)
    If Wb.Name = ThisWorkbook.Name Then Exit Sub

    ' Zamknij inne pliki bez zapisu, pokaz komunikat
    Dim wbName As String
    wbName = Wb.Name
    Application.EnableEvents = False
    Wb.Close SaveChanges:=False
    Application.EnableEvents = True

    MsgBox "BNC_Sender jest aktywny - nie mozna otwierac innych plikow." & vbCrLf & vbCrLf & _
           "Aby otworzyc '" & wbName & "', zamknij najpierw BNC_Sender.", _
           vbExclamation, "Blokada wielu plikow"
End Sub

Private Sub App_WorkbookActivate(ByVal Wb As Workbook)
    ' Defensywnie: jesli jakos inny wb sie aktywowal, zamknij
    If Wb.Name <> ThisWorkbook.Name Then
        Application.EnableEvents = False
        Wb.Close SaveChanges:=False
        Application.EnableEvents = True
        ThisWorkbook.Activate
    End If
End Sub
```

3. W `ThisWorkbook` dodaj:
```vba
Private m_AppEvents As cls_AppEvents

Private Sub Workbook_Open()
    ' ... (istniejacy kod) ...

    ' Hookup Application events
    Set m_AppEvents = New cls_AppEvents
    Set m_AppEvents.App = Application
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' ... (istniejacy kod) ...

    Set m_AppEvents = Nothing  ' unhook
End Sub
```

**Tradeoff**:
- + Solidna blokada, dziala automatycznie
- + User dostaje jasny komunikat
- - Wymaga dodatkowego class module
- - Uzytkownik moze byc zmuszony do zamkniecia BNC_Sender zeby otworzyc np. attachment z maila
- ! Jesli BNC_Sender wywoluje `Workbooks.Add` (mod_MailSender.GenerateTempFile,
    mod_DataCacheSync.SyncToFile, mod_UsersRegistrySync.SyncRegistryToFile) -
    event tez fire'uje na te internal xlsx! Trzeba dodac flage guard.

**Fix dla internal Workbooks.Add** (dodac do cls_AppEvents):
```vba
Public InternalOperationInProgress As Boolean

Private Sub App_WorkbookOpen(ByVal Wb As Workbook)
    If InternalOperationInProgress Then Exit Sub
    If Wb.Name = ThisWorkbook.Name Then Exit Sub
    ' ... reszta jak wyzej
End Sub
```

I w mod_MailSender / mod_DataCacheSync / mod_UsersRegistrySync przed Workbooks.Add:
```vba
' Guard internal xlsx operations
On Error Resume Next
If Not m_AppEvents Is Nothing Then m_AppEvents.InternalOperationInProgress = True
On Error GoTo 0

Set wbOut = Workbooks.Add
' ... praca z wbOut ...
wbOut.Close SaveChanges:=False

On Error Resume Next
If Not m_AppEvents Is Nothing Then m_AppEvents.InternalOperationInProgress = False
On Error GoTo 0
```

### Wariant B - Prosciej: dedykowana instancja Excel

Zamiast blokowania w BNC_Sender, uruchamiaj xlsm w OSOBNEJ instancji Excela.
Wtedy inne pliki naturalnie ida do glownej instancji, izolacja.

**Setup**: shortcut do Excela z flaga `/x`:
```
"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE" /x "C:\path\BNC_Sender_v0.1.0.xlsm"
```

**Tradeoff**:
- + Zero kodu, zero maintenance
- + Naturalna izolacja - kazdy inny plik ide do glownej instancji
- - Wymaga specjalnego shortcuta / instrukcji dla usera
- - User moze omijac otwierajac xlsm bezposrednio
- - Nie enforced w kodzie

### Wariant C - Nic nie robimy, polegamy na UX

Jesli user otworzy inny plik, on i tak nie zaklocaja BNC_Sender (rozne
workbooki, rozne code modules). Ryzyko: user moze przypadkiem zapisac zmiany
w niewlasciwym pliku.

**Tradeoff**: prostota vs zero enforcement.

---

## Opcja 2: VBA Project password

**Cel**: user (lub IT) nie moze podejrzec/zmienic kodu VBA.

**Setup**:
1. VBE -> Tools -> **VBAProject Properties**
2. Tab **Protection**
3. Checkbox **Lock project for viewing**
4. Password + Confirm password
5. OK, save xlsm, close + reopen -> kod niedostepny bez hasla

**Tradeoff**:
- + Ochrona przed reverse engineering / tampering
- - Utrata hasla = brak dostepu do kodu (backup!)
- - Nie chroni przed run - tylko przed view/edit
- - Trivial do zlamania (many online tools) - security through obscurity
- - Znacznie utrudnia debugowanie w produkcji

**Nie rekomendowane** dla wewnetrznego IT-team narzedzia. Warto dla external
distribution jesli chcemy obfuscation.

---

## Opcja 3: Workbook-level password (open password)

**Cel**: xlsm mozna otworzyc tylko z haslem.

**Setup**: File -> Info -> Protect Workbook -> Encrypt with Password.

**Tradeoff**:
- + Silny protection przy dobrym hasle (AES-256 w nowszym Excel)
- - User musi znac haslo -> logistyka distribution
- - Utrata hasla = brak dostepu do danych
- - Nie chroni po otwarciu (kod / dane wolne)

**Nie rekomendowane** dla 20+ handlowcow (koszt logistyki > benefit).

---

## Opcja 4: Digital signature dla VBA project

**Cel**: user widzi ze kod jest podpisany przez zaufanego wydawce, brak
"Enable Macros?" prompt przy kazdym otwarciu.

**Setup**:
1. Kup / wygeneruj self-signed certificate (SelfCert.exe z Office tools)
2. VBE -> Tools -> **Digital Signature** -> Choose -> wybierz cert
3. Save xlsm

**Tradeoff**:
- + Silny UX (no macro prompt dla trusted publisher)
- + Detekcja tampering (podpis znika po edycji kodu)
- - Self-signed = user musi rejst cert lokalnie (dep loy overhead)
- - Certyfikat od CA = koszt
- - Kazda edycja kodu wymaga re-sign

**Rekomendowane** jesli app idzie do 20+ handlowcow (jednorazowy setup CA cert,
long-term better UX). Poza scope Fazy A.

---

## Opcja 5: Trust access do VBA project object model

**Cel**: mod_Diagnostic wymaga tego (dostep do VBProject dla audit).

**Setup**: File -> Options -> Trust Center -> Trust Center Settings ->
Macro Settings -> checkbox **Trust access to the VBA project object model**.

**Tradeoff**: konieczne dla mod_Diagnostic, ale ryzyko: inne malicious macros
tez dostana access. Wlaczaj tylko na dev machines.

---

## Opcja 6: File format enforcement

**Cel**: user nie moze zapisac jako .xlsx (utrata makr).

**Metody**:

### 6a. Zablokuj Save As
```vba
' W ThisWorkbook
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    If SaveAsUI Then
        MsgBox "Save As jest zablokowany dla BNC_Sender. Uzyj Ctrl+S dla zwyklego save.", _
               vbExclamation, "Save As blocked"
        Cancel = True
    End If
End Sub
```

### 6b. Wymuszaj xlsm przy save
```vba
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    If LCase$(Right$(ThisWorkbook.FullName, 5)) <> ".xlsm" Then
        MsgBox "BNC_Sender musi byc zapisany jako .xlsm (macro-enabled).", _
               vbExclamation, "Wrong format"
        Cancel = True
    End If
End Sub
```

**Tradeoff**: proste, uzytecznie dla end-users. Rekomendowane dla produkcji.

---

## Opcja 7: Anty-tampering: hidden sheets protection

**Cel**: user nie moze un-hide `ws_AppState`, `ws_UsersRegistry`, `ws_DataCache`.

**Setup w kodzie** (dodac do LockLandingPage LUB nowy sub w ThisWorkbook):
```vba
Private Sub LockAllHiddenSheets()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible = xlSheetVeryHidden Then
            On Error Resume Next
            ws.Unprotect
            ws.Cells.Locked = True
            ws.Protect Password:="", UserInterfaceOnly:=True
            On Error GoTo 0
        End If
    Next ws
End Sub
```

Wywolac w Workbook_Open po ensureAppStateSheet + AddNewUser (kiedy juz sheets
istnieja).

**Tradeoff**: rzadko potrzebne bo sheets sa xlSheetVeryHidden (VBA-only unhide).
Chyba ze user zna VBE i klika Immediate `ws.Visible = xlSheetVisible`.

---

## Opcja 8: Kill switch / disable app

**Cel**: mozliwosc zdalnego wylaczenia app w razie krytycznego buga.

**Metoda**: sprawdz version na serwerze przy Workbook_Open, jesli disabled
-> MsgBox + close.

```vba
' W Workbook_Open, przed rounting:
If IsAppDisabled() Then
    MsgBox "Aplikacja BNC_Sender jest tymczasowo wylaczona." & vbCrLf & _
           "Skontaktuj sie z: tomasz.pirszel@inter-team.com.pl", _
           vbExclamation, "App disabled"
    ThisWorkbook.Close SaveChanges:=False
    Exit Sub
End If

Private Function IsAppDisabled() As Boolean
    On Error Resume Next
    Dim path As String
    path = "\\shared\bnc_sender\disabled.txt"
    IsAppDisabled = mod_Utils.FileExists(path)
End Function
```

**Tradeoff**: wymaga shared file location + policy. Nice-to-have dla enterprise.

---

## Rekomendacja per etap

**MVP (Faza A production)**:
- [x] Ribbon/chrome hidden (done)
- [x] Cell lock + no selection (done)
- [x] Hide cols/rows poza landing (done)
- [ ] **Opcja 6a/6b** - Save As guard + xlsm enforcement (low effort, high benefit)

**Faza B (multi-user, 20+ handlowcow)**:
- [ ] **Opcja 4** - Digital signature (zamiast per-user macro prompt)
- [ ] Opcja 8 - Kill switch (crisis mgmt)

**Nie rekomendowane teraz**:
- Opcja 1 (Application event block) - moze frustrowac userow, complexity high, internal xlsx guards messy
- Opcja 2 (VBA project password) - utrudnia debug, security-through-obscurity
- Opcja 3 (workbook password) - logistyka > benefit

---

## Related

- `sh_LandingPage.LAYOUT.md` sekcja **Krok 4b** - app-like mode details
- `ThisWorkbook.code.txt` - Workbook_Open + LockLandingPage source
- `feedback_vba_encoding_cp1250` (memory) - encoding rules dla wszystkich .bas edits
