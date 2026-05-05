# Plan implementacji BNC_Sender v0.1.0 - faza A

> **Dokument**: Plan wdrożenia krok po kroku  
> **Autor**: PM + Senior Engineer  
> **Wersja docelowa**: BNC_Sender_v0.1.0.xlsm  
> **Czas implementacji**: ~17 dni roboczych (3.5 tygodnia full-time lub 6-8 tygodni przy 50% angażowaniu)  
> **Stan przed startem**: Środowisko DEV skonfigurowane zgodnie z `BNC_srodowiskoDEV_FazaA.pdf`

---

## Spis treści

1. [Filozofia tego planu](#1-filozofia-tego-planu)
2. [Struktura repozytorium](#2-struktura-repozytorium)
3. [Roadmap - 7 milestones](#3-roadmap---7-milestones)
4. [M0 - Setup środowiska](#m0---setup-środowiska)
5. [M1 - Foundation (mod_Utils + Sync)](#m1---foundation)
6. [M2 - Setup form (frm_Setup + mod_Validation)](#m2---setup-form)
7. [M3 - Main form (frm_Main)](#m3---main-form)
8. [M4 - Mail sender (mod_MailSender)](#m4---mail-sender)
9. [M5 - Log + Export (frm_Log + mod_Export)](#m5---log--export)
10. [M6 - Polish + UAT](#m6---polish--uat)
11. [M7 - Release v1.0.0](#m7---release-v100)
12. [Dziennik decyzji architektonicznych](#12-dziennik-decyzji-architektonicznych)
13. [Lista kontrolna gotowości v1.0.0](#13-lista-kontrolna-gotowości-v100)

---

## 1. Filozofia tego planu

### Dlaczego "od dołu do góry"

Implementujemy moduły w kolejności **od fundamentów do warstwy prezentacji**. To znaczy: najpierw `mod_Utils` i `mod_*Sync` (warstwa data access), potem moduły logiki, na końcu UserForms.

**Dlaczego nie odwrotnie?** Bo gdy zaczniesz od `frm_Main`, to przy każdym kliknięciu przycisku będziesz pisał "TODO: tutaj zapisać dane". Stos TODO będzie rosnął, wpadniesz w paraliż decyzyjny. Z fundamentami w miejscu - kliknięcie przycisku to po prostu wywołanie gotowej funkcji `mod_DataCacheSync.AppendRecord(...)`.

### Dlaczego małe milestones

Każdy milestone (M0-M7) ma **deliverable** - coś, co działa i co możesz pokazać. Po M1 nie masz jeszcze aplikacji, ale masz "działający sync między worksheet a xlsx", co możesz przetestować w Immediate Window. To buduje zaufanie do kodu i pozwala wcześnie wyłapać błędy.

### Konwencja zatwierdzania kroków

Każdy krok w tym dokumencie ma checkbox `- [ ]`. Po ukończeniu zmieniasz na `- [x]`. **Nie idziesz dalej, dopóki kryteria akceptacji nie są spełnione**. Senior nigdy nie idzie do następnego modułu, jeśli poprzedni nie działa - bo bug w fundamentach kosztuje 10x więcej do naprawy niż bug w warstwie prezentacji.

### Pseudokod vs gotowy kod

W tym planie znajdziesz **pseudokod** (intencje, struktura), nie gotowy kod do skopiowania. Powody dwa:

1. **Uczysz się** - wpisując kod ręcznie, rozumiesz co robisz. Skopiowany kod to kod, który "działa magicznie" - i nie wiesz, jak go naprawić, gdy padnie.
2. **Lepiej dopasujesz do swojego stylu** - VBA ma kilka konwencji (Dim wszędzie vs Dim na początku, error handling On Error vs Try-Catch via classes itp.). Pseudokod pozwala wybrać.

---

## 2. Struktura repozytorium

```
C:\Dev\BNC_Sender\
└── FazaA\
    ├── Releases\                        # archiwum wersji xlsm
    │   ├── BNC_Sender_v0.1.0.xlsm
    │   └── BNC_Sender_v0.2.0.xlsm       # kolejne wersje
    │
    ├── Working\                         # aktualnie edytowana wersja
    │   └── BNC_Sender_v0.X.0.xlsm
    │
    ├── CacheFolder\                     # testowy folder cache (DEV only)
    │   ├── BNC_UserCache.xlsx
    │   └── BNC_DataCache.xlsx
    │
    ├── Source\                          # eksporty kodu z VBA Editor (Git tracked!)
    │   ├── Modules\                     # moduły .bas
    │   │   ├── mod_Utils.bas
    │   │   ├── mod_Validation.bas
    │   │   ├── mod_UserCacheSync.bas
    │   │   ├── mod_DataCacheSync.bas
    │   │   ├── mod_MailSender.bas
    │   │   └── mod_Export.bas
    │   │
    │   ├── Forms\                       # UserForms .frm + .frx
    │   │   ├── frm_Setup.frm            # tekst (Git diff działa)
    │   │   ├── frm_Setup.frx            # binarny (kontrolki, ikony)
    │   │   ├── frm_Main.frm
    │   │   ├── frm_Main.frx
    │   │   ├── frm_Log.frm
    │   │   └── frm_Log.frx
    │   │
    │   └── ThisWorkbook\                # eventy aplikacji
    │       └── ThisWorkbook.cls         # Workbook_Open i inne
    │
    ├── Notatki\                         # dokumentacja deweloperska
    │   ├── CHANGELOG.md                 # co zmieniło się w każdej wersji
    │   ├── TODO.md                      # bieżąca lista zadań
    │   └── DECISIONS.md                 # Architectural Decision Records
    │
    ├── README.md                        # opis projektu, jak uruchomić
    └── .gitignore                       # ignoruj Working\, Releases\, CacheFolder\
```

### Co commitujesz do Git, a co nie

**Commituj** (`git add`):
- `Source/` - cały folder z eksportami `.bas`, `.frm`, `.frx`, `.cls`
- `Notatki/` - changelog, todo, decisions
- `README.md`, `.gitignore`

**Nie commituj** (w `.gitignore`):
- `Releases/` - pliki xlsm (binarne, duże, można odbudować z Source)
- `Working/` - aktualnie edytowana wersja
- `CacheFolder/` - dane testowe
- `~$*.xlsm` - tymczasowe pliki Excela

### `.gitignore` do skopiowania

```gitignore
# Pliki Excel - binarne, nie do Git
*.xlsm
*.xlsx
*.xls

# Tymczasowe pliki Office
~$*

# Foldery z plikami binarnymi
Releases/
Working/
CacheFolder/

# Wyjątek: chcemy commitować .frx mimo że są binarne (są częścią eksportu UserForm)
!Source/Forms/*.frx
```

### Workflow eksportu kodu z VBA do Source/

Po każdej istotnej zmianie kodu w VBA Editor:

1. W VBA Editor klikasz prawym na moduł lub formularz
2. Wybierasz **Export File...**
3. Zapisujesz w odpowiednim folderze (`Source/Modules/` lub `Source/Forms/`)
4. W terminalu: `git add Source/`, `git commit -m "feat(M2): dodano walidację email w mod_Validation"`

**To jest robota manualna, ale niezbędna.** Bez tego Git nie widzi zmian w kodzie VBA.

---

## 3. Roadmap - 7 milestones

| # | Milestone | Co dostarczasz | Czas |
|---|---|---|---|
| **M0** | Setup środowiska | Działający xlsm + Outlook + ścieżki | ~1 dzień |
| **M1** | Foundation | mod_Utils, mod_UserCacheSync, mod_DataCacheSync | ~3 dni |
| **M2** | Setup form | frm_Setup + mod_Validation + samouczek | ~2 dni |
| **M3** | Main form | frm_Main + dodawanie do batcha + lista pending | ~3 dni |
| **M4** | Mail sender | mod_MailSender + decision diamond + plik tymczasowy | ~3 dni |
| **M5** | Log + Export | frm_Log + mod_Export | ~2 dni |
| **M6** | Polish + UAT | Error handling, UX polish, testy z 3-5 userami | ~3 dni |
| **M7** | Release v1.0.0 | Finalna wersja gotowa do wdrożenia | - |

**Łącznie**: ~17 dni roboczych = 3.5 tygodnia full-time lub 6-8 tygodni przy 50% angażowaniu.

---

## M0 - Setup środowiska

> **Czas**: ~1 dzień  
> **Wynik**: Środowisko DEV gotowe, plik xlsm z arkuszami, smoke testy przechodzą

### Zadania

- [ ] **0.1** Wykonaj wszystkie kroki z `BNC_srodowiskoDEV_FazaA.pdf` (etapy I i II)
- [ ] **0.2** Zainicjalizuj Git w `C:\Dev\BNC_Sender\`
  ```bash
  cd C:\Dev\BNC_Sender
  git init
  git add .gitignore README.md
  git commit -m "chore: initial commit"
  ```
- [ ] **0.3** Utwórz `Source/Modules/` i `Source/Forms/` (puste foldery, ale z `.gitkeep` żeby Git je trackował)
- [ ] **0.4** Utwórz pierwszy `CHANGELOG.md`:
  ```markdown
  # CHANGELOG - BNC_Sender
  
  ## v0.1.0 (in progress)
  - Faza A: implementacja MVP z hybrid cache i workflow kierownika
  ```
- [ ] **0.5** Utwórz `DECISIONS.md` (pusty na razie - będziesz dodawać ADR w trakcie)
- [ ] **0.6** Skopiuj `BNC_Sender_v0.1.0.xlsm` z `Releases/` do `Working/`

### Kryterium akceptacji M0

- [ ] Trzy smoke testy z dokumentu DEV przechodzą (read userCache, write dataCache, send mail)
- [ ] Git widzi `Source/`, `Notatki/`, `README.md`, `.gitignore`
- [ ] Możesz uruchomić `git status` i widzisz "nothing to commit, working tree clean"

---

## M1 - Foundation

> **Czas**: ~3 dni  
> **Wynik**: Działający sync między worksheets a plikami xlsx, helpery dostępne dla reszty kodu

To jest najważniejszy milestone całego planu. Gdy te trzy moduły działają, reszta aplikacji budowana jest "na nich" i ma solidne fundamenty.

### M1.1 - mod_Utils

**Lokalizacja w VBA**: Module (Insert > Module, nazwa: `mod_Utils`)

**Po co istnieje**: Helpery używane przez wszystkie inne moduły. Centralizacja typowych operacji żeby uniknąć duplikacji.

**Public API (procedury widoczne z innych modułów)**:

```vba
' ==================== mod_Utils ====================

' Logowanie do Immediate Window z timestampem
Public Sub LogInfo(message As String)
    ' Output: [2026-05-05 14:23:11] INFO: <message>
End Sub

Public Sub LogError(source As String, errNumber As Long, errDescription As String)
    ' Output: [2026-05-05 14:23:11] ERROR in <source>: #<errNumber> - <errDescription>
End Sub

' Helpery do dat - bo Format() ma w VBA dziwne quirki z lokalizacją
Public Function FormatTimestampISO(dt As Date) As String
    ' Wynik: "2026-05-05T14:23:11"
End Function

Public Function GetCurrentMonthYear() As Date
    ' Wynik: pierwszy dzień bieżącego miesiąca, np. 2026-05-01
End Function

' Helpery do plików - opakowania FileSystemObject
Public Function FileExists(filePath As String) As Boolean
End Function

Public Function FolderExists(folderPath As String) As Boolean
End Function

Public Sub EnsureFolderExists(folderPath As String)
    ' Tworzy folder rekursywnie jeśli nie istnieje
End Sub

' Helpery do walidacji typów
Public Function IsValidEmail(text As String) As Boolean
    ' Prosta regex: cos@cos.cos
End Function

Public Function IsValidLong(text As String) As Boolean
    ' Czy text jest poprawną liczbą całkowitą Long
End Function
```

### Zadania M1.1

- [ ] **1.1.1** Utwórz moduł `mod_Utils` w VBA Editor
- [ ] **1.1.2** Zaimplementuj wszystkie funkcje z Public API powyżej
- [ ] **1.1.3** Napisz test w nowym module `mod_Tests`:
  ```vba
  Sub Test_mod_Utils()
      ' Test LogInfo
      mod_Utils.LogInfo "Test message"
      
      ' Test IsValidEmail
      Debug.Print mod_Utils.IsValidEmail("test@example.com")  ' True
      Debug.Print mod_Utils.IsValidEmail("not-email")          ' False
      
      ' Test FileExists
      Debug.Print mod_Utils.FileExists(ThisWorkbook.FullName)  ' True
      
      ' itd.
  End Sub
  ```
- [ ] **1.1.4** Wszystkie testy zwracają oczekiwane wyniki
- [ ] **1.1.5** Eksport: `Source/Modules/mod_Utils.bas` (File > Export)
- [ ] **1.1.6** Git commit: `git commit -m "feat(M1): mod_Utils z helperami logowania, dat, plików, walidacji"`

### M1.2 - mod_UserCacheSync

**Lokalizacja w VBA**: Module (`mod_UserCacheSync`)

**Po co istnieje**: Hermetyzuje cały dostęp do `ws_UserCache`. Reszta aplikacji nie czyta bezpośrednio z arkusza - woła funkcje tego modułu. Plus: synchronizuje stan z plikiem `BNC_UserCache.xlsx`.

**Wzorzec architektoniczny**: Repository Pattern. `ws_UserCache` to "tabela", a `mod_UserCacheSync` to "DAO" (Data Access Object).

**Public API**:

```vba
' ==================== mod_UserCacheSync ====================

' Czytanie pojedynczych wartości
Public Function GetUserField(fieldKey As String) As Variant
    ' Czyta wartość z ws_UserCache dla podanego klucza
    ' fieldKey: "Imie", "EmailHandlowca", itd.
End Function

' Czytanie całego usera jako struktury
Public Function GetUserData() As Object
    ' Zwraca Scripting.Dictionary z wszystkimi polami
    ' Kluczowe pole: SetupCompleted (Boolean)
End Function

' Zapis pojedynczej wartości
Public Sub SetUserField(fieldKey As String, value As Variant)
    ' Zapisuje do ws_UserCache
    ' Po zapisie: ThisWorkbook.Save + SyncToFile
End Sub

' Zapis całej struktury usera (np. po setupie)
Public Sub SaveUserData(userData As Object)
    ' Pisze wszystkie pola, ustawia SetupCompleted=True
    ' Po zapisie: ThisWorkbook.Save + SyncToFile
End Sub

' Czy user ukończył setup?
Public Function IsSetupCompleted() As Boolean
End Function

' Wykrywanie roli (kluczowa funkcja!)
Public Function IsUserManager() As Boolean
    ' Zwraca True jeśli EmailKierownika == EmailHandlowca
End Function

' Synchronizacja do pliku xlsx
Private Sub SyncToFile()
    ' Kopiuje ws_UserCache do <CacheFolderPath>\BNC_UserCache.xlsx
    ' Best-effort: jeśli nie uda się, log błędu, ale nie throw
End Sub

' Auto-recreate (wywoływane z Workbook_Open)
Public Sub EnsureCacheFileExists()
    ' Jeśli BNC_UserCache.xlsx nie istnieje, tworzy go z ws_UserCache
End Sub
```

### Zadania M1.2

- [ ] **1.2.1** Utwórz moduł `mod_UserCacheSync`
- [ ] **1.2.2** Zaimplementuj wszystkie publiczne funkcje
- [ ] **1.2.3** Implementacja `SyncToFile`:
  - Otwórz nowy workbook (`Workbooks.Add`)
  - Skopiuj zakres z `ws_UserCache`
  - SaveAs jako xlsx (FileFormat=xlOpenXMLWorkbook) w `<CacheFolderPath>\BNC_UserCache.xlsx`
  - Close bez zapisu nowego workbooka
  - Wszystko w `On Error GoTo Cleanup` - jeśli padnie, log + cleanup
- [ ] **1.2.4** Test:
  ```vba
  Sub Test_mod_UserCacheSync()
      ' Test odczytu
      Debug.Print mod_UserCacheSync.GetUserField("Imie")
      Debug.Print mod_UserCacheSync.IsSetupCompleted()
      Debug.Print mod_UserCacheSync.IsUserManager()
      
      ' Test zapisu
      mod_UserCacheSync.SetUserField "Imie", "TestImie"
      
      ' Sprawdź czy BNC_UserCache.xlsx zaktualizował się
      ' (otwórz plik w Eksploratorze i podejrzyj)
  End Sub
  ```
- [ ] **1.2.5** Eksport: `Source/Modules/mod_UserCacheSync.bas`
- [ ] **1.2.6** Git commit

### M1.3 - mod_DataCacheSync

**Lokalizacja w VBA**: Module (`mod_DataCacheSync`)

**Po co istnieje**: Identycznie jak `mod_UserCacheSync`, ale dla `ws_DataCache`. Operuje na zgłoszeniach (wiele wierszy), nie na single user.

**Public API**:

```vba
' ==================== mod_DataCacheSync ====================

' Dodanie nowego zgłoszenia
Public Function AppendRecord(reportData As Object) As Long
    ' reportData to Scripting.Dictionary z polami:
    '   KlientFK, NazwaKlienta, MiesiacZgloszenia, Fields
    ' Aplikacja sama dodaje:
    '   ReportID (autoincrement), CNA_HandlowcaID, NrOddzialu (snapshot z UserCache)
    '   CreatedTimestamp = Now(), Status = "pending", EmailRecipient = "", BatchSentTimestamp = ""
    ' Zwraca: ReportID nowego rekordu
End Function

' Pobranie wszystkich pending recordów (do wysyłki batcha)
Public Function GetPendingRecords() As Collection
    ' Zwraca Collection z Dictionary (każdy = 1 rekord)
End Function

' Pobranie wszystkich recordów (do log_UserForm)
Public Function GetAllRecords() As Collection
End Function

' UPDATE statusu i recipientu po wysyłce
Public Sub MarkAsSent(reportIDs As Collection, recipient As String)
    ' Dla każdego ReportID z reportIDs:
    '   - Status = "sent"
    '   - EmailRecipient = recipient
    '   - BatchSentTimestamp = Now()
    ' Po update: ThisWorkbook.Save + SyncToFile
End Sub

' Synchronizacja do pliku xlsx
Private Sub SyncToFile()
    ' Identyczna logika jak w mod_UserCacheSync, ale dla ws_DataCache
End Sub

' Auto-recreate
Public Sub EnsureCacheFileExists()
End Sub

' Pomocnicze - następny ReportID
Private Function GetNextReportID() As Long
    ' Czyta ostatni ReportID z ws_DataCache, dodaje 1
    ' Jeśli arkusz pusty - zwraca 1
End Function
```

### Zadania M1.3

- [ ] **1.3.1** Utwórz moduł `mod_DataCacheSync`
- [ ] **1.3.2** Zaimplementuj wszystkie publiczne funkcje
- [ ] **1.3.3** Test:
  ```vba
  Sub Test_mod_DataCacheSync()
      ' Test AppendRecord
      Dim newRecord As Object
      Set newRecord = CreateObject("Scripting.Dictionary")
      newRecord("KlientFK") = 12345
      newRecord("NazwaKlienta") = "Test Klient"
      newRecord("MiesiacZgloszenia") = mod_Utils.GetCurrentMonthYear()
      newRecord("Fields") = "Test fields"
      
      Dim newID As Long
      newID = mod_DataCacheSync.AppendRecord(newRecord)
      Debug.Print "Dodano rekord ID: " & newID
      
      ' Test GetPendingRecords
      Dim pending As Collection
      Set pending = mod_DataCacheSync.GetPendingRecords()
      Debug.Print "Pending count: " & pending.Count
      
      ' Test MarkAsSent
      Dim ids As New Collection
      ids.Add newID
      mod_DataCacheSync.MarkAsSent ids, "test@example.com"
      
      ' Sprawdź w arkuszu
  End Sub
  ```
- [ ] **1.3.4** Eksport + commit

### Kryterium akceptacji M1

- [ ] Wszystkie 3 testy modułów zwracają oczekiwane wyniki
- [ ] Możesz dodać 5 rekordów testowych do `ws_DataCache`, oznaczyć 3 jako sent, i widzisz prawidłowy stan w pliku `BNC_DataCache.xlsx`
- [ ] Pliki w `Source/Modules/` są zsynchronizowane z VBA Editor
- [ ] Git: 3 commity (po jednym na moduł), wszystko na branchu main

### Decyzja architektoniczna do zapisu w DECISIONS.md

```markdown
## ADR-001: Repository Pattern dla cache

**Decyzja**: Każdy ukryty arkusz (ws_UserCache, ws_DataCache) ma dedykowany moduł 
(mod_*Sync), który jest jedynym miejscem dostępu do tego arkusza.

**Uzasadnienie**: Hermetyzacja danych. Reszta aplikacji nie wie, czy dane są 
w worksheet, w pliku xlsx, czy gdzie indziej. W fazie B podmienimy implementację 
na bazę Access bez zmian w warstwach wyższych.

**Konsekwencja**: Trochę więcej kodu (każda operacja ma function w mod_*Sync), 
ale zero duplikacji i pełna kontrola nad zmianami.
```

---

## M2 - Setup form

> **Czas**: ~2 dni  
> **Wynik**: Pierwszy uruchomieniowy ekran aplikacji - rejestracja usera

### M2.1 - mod_Validation

**Po co istnieje**: Reguły walidacji w jednym miejscu. Reszta kodu pyta "czy te dane są poprawne?", a ten moduł odpowiada.

**Public API**:

```vba
' ==================== mod_Validation ====================

' Walidacja danych setupu
Public Function ValidateSetupData(userData As Object) As String
    ' userData: Scripting.Dictionary z polami z frm_Setup
    ' Zwraca: "" jeśli OK, w przeciwnym razie komunikat błędu
End Function

' Walidacja danych zgłoszenia (w main_UserForm)
Public Function ValidateReportData(reportData As Object) As String
    ' Sprawdza:
    '   - KlientFK to liczba (Long)
    '   - NazwaKlienta nie puste, długość 3-200
    '   - Fields nie puste (lub pole opcjonalne, do decyzji)
    ' Zwraca: "" jeśli OK, komunikat błędu
End Function

' Walidacje atomowe (też używane bezpośrednio z UI)
Public Function ValidateEmail(email As String) As Boolean
End Function

Public Function ValidateClientFK(fk As String) As Boolean
End Function

Public Function ValidateNonEmpty(text As String) As Boolean
End Function

Public Function ValidateLength(text As String, minLen As Long, maxLen As Long) As Boolean
End Function
```

### Zadania M2.1

- [ ] **2.1.1** Utwórz moduł `mod_Validation`
- [ ] **2.1.2** Zaimplementuj funkcje walidacyjne (używaj `mod_Utils.IsValidEmail`, `mod_Utils.IsValidLong`)
- [ ] **2.1.3** Test:
  ```vba
  Sub Test_mod_Validation()
      Debug.Print mod_Validation.ValidateEmail("test@example.com")  ' True
      Debug.Print mod_Validation.ValidateClientFK("12345")           ' True
      Debug.Print mod_Validation.ValidateClientFK("abc")             ' False
      
      Dim setupData As Object
      Set setupData = CreateObject("Scripting.Dictionary")
      setupData("Imie") = "Jan"
      setupData("Nazwisko") = "Kowalski"
      setupData("EmailHandlowca") = "jan@firma.pl"
      ' ... reszta pól ...
      Debug.Print mod_Validation.ValidateSetupData(setupData)  ' "" jeśli OK
  End Sub
  ```
- [ ] **2.1.4** Eksport + commit

### M2.2 - frm_Setup

**Lokalizacja w VBA**: UserForm (Insert > UserForm, nazwa: `frm_Setup`)

**Po co istnieje**: Jednorazowa rejestracja handlowca. Pokazuje się tylko gdy `mod_UserCacheSync.IsSetupCompleted()` zwraca False.

**Layout (kontrolki na formularzu)**:

```
┌─────────────────────────────────────────────────────────────┐
│ BNC_Sender - Konfiguracja wstępna                       [X] │
├─────────────────────────────────────────────────────────────┤
│                                                             │
│  Witaj w aplikacji BNC_Sender. Aby rozpocząć, podaj         │
│  swoje dane służbowe.                                       │
│                                                             │
│  Imię:                  [____________________]              │
│  Nazwisko:              [____________________]              │
│  Email służbowy:        [____________________]              │
│  CNA (numer handlowca): [____________________]              │
│  Numer oddziału:        [____________________]              │
│                                                             │
│  ─── Adresy do wysyłki ───────────────────────              │
│                                                             │
│  Email kierownika:      [____________________]              │
│   ℹ️ Jeśli jesteś kierownikiem, wpisz swój własny adres -   │
│      wnioski będą wysyłane wprost do BNC.                   │
│                                                             │
│  Email zespołu BNC:     [____________________]              │
│                                                             │
│  ─── Lokalizacja plików cache ────────────────              │
│                                                             │
│  Folder cache:          [C:\BNC_CacheFolder\______]         │
│                         [Przeglądaj...]                     │
│                                                             │
│  ─── Samouczek ────────────────────────────────              │
│                                                             │
│  [scrollable text area z instrukcją obsługi]                │
│                                                             │
│  [✓] Nie pokazuj tego samouczka ponownie                    │
│                                                             │
│                          [Anuluj]  [Zapisz konfigurację]    │
└─────────────────────────────────────────────────────────────┘
```

**Kontrolki (nazwy)**:
- TextBox: `txt_Imie`, `txt_Nazwisko`, `txt_EmailHandlowca`, `txt_CNA`, `txt_NrOddzialu`, `txt_EmailKierownika`, `txt_EmailBNC`, `txt_CacheFolderPath`
- Label: `lbl_Imie`, `lbl_Nazwisko`, ... (etykiety obok każdego TextBox)
- TextBox (multiline, scrollable): `txt_Tutorial`
- CheckBox: `chk_DontShowTutorial`
- Button: `btn_Browse`, `btn_Cancel`, `btn_Save`

**Kod formularza** (w `frm_Setup` jako Code Behind):

```vba
' ==================== frm_Setup ====================

Private Sub UserForm_Initialize()
    ' Wypełnij txt_Tutorial treścią samouczka
    txt_Tutorial.Text = GetTutorialText()
    
    ' Domyślna lokalizacja cache
    txt_CacheFolderPath.Text = "C:\BNC_CacheFolder\"
    
    ' Sprawdź, czy user już ma jakieś dane (np. powtarza setup)
    If mod_UserCacheSync.GetUserField("SetupCompleted") = True Then
        ' Wczytaj istniejące dane do pól
        txt_Imie.Text = mod_UserCacheSync.GetUserField("Imie")
        ' ... reszta pól ...
    End If
End Sub

Private Sub btn_Browse_Click()
    ' Otwórz Windows folder picker
    Dim folderPath As String
    folderPath = SelectFolder()
    If folderPath <> "" Then
        txt_CacheFolderPath.Text = folderPath
    End If
End Sub

Private Sub btn_Save_Click()
    ' 1. Zbierz dane z pól
    Dim userData As Object
    Set userData = CreateObject("Scripting.Dictionary")
    userData("Imie") = Trim(txt_Imie.Text)
    userData("Nazwisko") = Trim(txt_Nazwisko.Text)
    userData("EmailHandlowca") = Trim(txt_EmailHandlowca.Text)
    userData("CNA_HandlowcaID") = Trim(txt_CNA.Text)
    userData("NrOddzialu") = Trim(txt_NrOddzialu.Text)
    userData("EmailKierownika") = Trim(txt_EmailKierownika.Text)
    userData("EmailBNC") = Trim(txt_EmailBNC.Text)
    userData("CacheFolderPath") = Trim(txt_CacheFolderPath.Text)
    userData("DataRejestracji") = Now()
    userData("SetupCompleted") = True
    
    ' 2. Walidacja
    Dim errMsg As String
    errMsg = mod_Validation.ValidateSetupData(userData)
    If errMsg <> "" Then
        MsgBox errMsg, vbExclamation, "Błąd walidacji"
        Exit Sub
    End If
    
    ' 3. Sprawdź czy folder cache istnieje, jeśli nie - utwórz
    mod_Utils.EnsureFolderExists userData("CacheFolderPath")
    
    ' 4. Zapisz do ws_UserCache (z auto-sync do pliku)
    mod_UserCacheSync.SaveUserData userData
    
    ' 5. Zapisz preferencję samouczka
    ' (TODO: dodać pole DontShowTutorial w ws_UserCache?)
    
    ' 6. Zamknij setup, otwórz frm_Main
    Me.Hide
    frm_Main.Show
End Sub

Private Sub btn_Cancel_Click()
    ' Zapytaj czy na pewno
    If MsgBox("Czy na pewno przerwać konfigurację? Aplikacja nie uruchomi się.", _
              vbYesNo + vbQuestion) = vbYes Then
        Me.Hide
        ' Pozwól userowi zamknąć aplikację ręcznie
    End If
End Sub

Private Function GetTutorialText() As String
    ' Wieloliniowy tekst samouczka
    Dim t As String
    t = "Witaj w BNC_Sender." & vbCrLf & vbCrLf
    t = t & "Aplikacja służy do zgłaszania nowych klientów do programu BNC " & _
            "(Bonus New Client)." & vbCrLf & vbCrLf
    t = t & "JAK TO DZIAŁA:" & vbCrLf
    t = t & "1. Wprowadzasz pojedyncze zgłoszenia w formularzu głównym." & vbCrLf
    t = t & "2. Możesz dodać kilka zgłoszeń przed wysyłką - tworzą one batch." & vbCrLf
    t = t & "3. Klikasz 'Wyślij Wniosek BNC' - cały batch leci jednym mailem." & vbCrLf & vbCrLf
    t = t & "ROLA W APLIKACJI:" & vbCrLf
    t = t & "- Jeśli email kierownika jest taki sam jak Twój email służbowy, " & _
            "aplikacja uznaje Cię za kierownika i wysyła wnioski wprost do BNC." & vbCrLf
    t = t & "- W przeciwnym razie wnioski lecą do kierownika z prośbą o " & _
            "weryfikację i przekazanie do BNC." & vbCrLf
    GetTutorialText = t
End Function
```

### Zadania M2.2

- [ ] **2.2.1** Utwórz UserForm `frm_Setup`
- [ ] **2.2.2** Dodaj wszystkie kontrolki zgodnie z layoutem
- [ ] **2.2.3** Ustaw właściwości: nazwy, captions, MaxLength dla pól, MultiLine=True dla samouczka
- [ ] **2.2.4** Zaimplementuj kod formularza
- [ ] **2.2.5** Implementacja `SelectFolder()` - można użyć `Application.FileDialog(msoFileDialogFolderPicker)`
- [ ] **2.2.6** Test manualny:
  - Uruchom `frm_Setup.Show` z Immediate Window
  - Wypełnij dane testowe, kliknij Save
  - Sprawdź czy w `ws_UserCache` są dane i czy `BNC_UserCache.xlsx` jest zaktualizowany
- [ ] **2.2.7** Test scenariuszy negatywnych:
  - Puste pole - dostajesz komunikat błędu
  - Błędny email - dostajesz komunikat błędu
  - Niepoprawny folder - dostajesz komunikat
- [ ] **2.2.8** Eksport `Source/Forms/frm_Setup.frm` + `frm_Setup.frx`
- [ ] **2.2.9** Git commit

### M2.3 - ThisWorkbook.Workbook_Open

**Lokalizacja w VBA**: ThisWorkbook (klikasz na "ThisWorkbook" w Project Explorer)

**Po co istnieje**: Decyduje, co pokazać przy otwarciu pliku - setup czy main form.

```vba
' ==================== ThisWorkbook ====================

Private Sub Workbook_Open()
    ' 1. Auto-recreate plików cache jeśli nie istnieją
    mod_UserCacheSync.EnsureCacheFileExists
    mod_DataCacheSync.EnsureCacheFileExists
    
    ' 2. Decyzja: setup czy main?
    If mod_UserCacheSync.IsSetupCompleted() Then
        frm_Main.Show
    Else
        frm_Setup.Show
    End If
End Sub
```

### Zadania M2.3

- [ ] **2.3.1** Otwórz ThisWorkbook w VBA Editor
- [ ] **2.3.2** Dodaj kod `Workbook_Open`
- [ ] **2.3.3** Zapisz plik xlsm (`ThisWorkbook.Save`)
- [ ] **2.3.4** Test: zamknij plik, otwórz ponownie
  - Jeśli setup nie był zrobiony → pokazuje się `frm_Setup`
  - Jeśli setup był zrobiony → pokazuje się `frm_Main` (jeszcze pusty)
- [ ] **2.3.5** Eksport `Source/ThisWorkbook/ThisWorkbook.cls`
- [ ] **2.3.6** Git commit

### Kryterium akceptacji M2

- [ ] Otwarcie pliku po raz pierwszy pokazuje `frm_Setup`
- [ ] Wypełnienie i zapis danych powoduje przejście do (jeszcze pustego) `frm_Main`
- [ ] Ponowne otwarcie pliku pokazuje `frm_Main` od razu
- [ ] Reset DEV (usunięcie danych z `ws_UserCache`) powoduje powrót do `frm_Setup`
- [ ] Walidacja działa dla każdego pola

### ADR do zapisu

```markdown
## ADR-002: Workbook_Open jako entry point

**Decyzja**: Cała logika decyzji "co pokazać przy starcie" jest w 
ThisWorkbook.Workbook_Open. Frm_Setup i frm_Main nie wiedzą o sobie nawzajem.

**Uzasadnienie**: Single point of entry. Łatwo dodać kolejne ekrany 
(np. frm_Update gdy aplikacja jest stara).

**Konsekwencja**: ThisWorkbook ma minimum kodu, ale jest "centralą".
```

---

## M3 - Main form

> **Czas**: ~3 dni  
> **Wynik**: Działający formularz dodawania zgłoszeń do batcha + lista pending

### M3.1 - frm_Main UI i logika dodawania

**Po co istnieje**: Główny ekran użytkownika. Pozwala wpisywać zgłoszenia i wysyłać batch.

**Layout**:

```
┌──────────────────────────────────────────────────────────────────────┐
│ BNC_Sender - Wniosek BNC                                         [X] │
├──────────────────────────────────────────────────────────────────────┤
│ Zalogowany: Jan Kowalski (CNA: 12345, oddział: W001)                 │
│ Tryb: HANDLOWIEC (wnioski będą wysyłane do kierownika)               │
├──────────────────────────────────────────────────────────────────────┤
│                                                                      │
│ ─── Nowe zgłoszenie ─────────────────────────                        │
│                                                                      │
│ Klient FK:           [____________________]                          │
│ Nazwa klienta:       [____________________]                          │
│ Miesiąc zgłoszenia:  [____________________] (bieżący domyślnie)      │
│ Pole dodatkowe:      [____________________]                          │
│                                                                      │
│                                  [Wyczyść]  [Dodaj do listy]         │
│                                                                      │
│ ─── Lista zgłoszeń do wysłania (3) ──────────                        │
│                                                                      │
│ ┌────────────────────────────────────────────────────────────────┐   │
│ │ ID │ KlientFK │ Nazwa klienta      │ Miesiąc  │ Pole          │   │
│ ├────┼──────────┼────────────────────┼──────────┼───────────────┤   │
│ │ 47 │ 12345    │ Acme Sp z o o     │ 2026-05  │ ...           │   │
│ │ 48 │ 67890    │ Bravo Ltd          │ 2026-05  │ ...           │   │
│ │ 49 │ 11111    │ Charlie SA         │ 2026-05  │ ...           │   │
│ └────────────────────────────────────────────────────────────────┘   │
│                                                                      │
│                              [Pokaż historię]  [Wyślij Wniosek BNC]  │
└──────────────────────────────────────────────────────────────────────┘
```

**Kontrolki**:
- Label: `lbl_UserInfo`, `lbl_RoleInfo`, `lbl_BatchCount`
- TextBox: `txt_KlientFK`, `txt_NazwaKlienta`, `txt_MiesiacZgloszenia`, `txt_Fields`
- ListBox lub ListView: `lst_PendingBatch`
- Button: `btn_Clear`, `btn_AddToList`, `btn_ShowLog`, `btn_SendBatch`

**Kod**:

```vba
' ==================== frm_Main ====================

Private Sub UserForm_Initialize()
    ' Pokaż info o userze
    lbl_UserInfo.Caption = "Zalogowany: " _
        & mod_UserCacheSync.GetUserField("Imie") & " " _
        & mod_UserCacheSync.GetUserField("Nazwisko") _
        & " (CNA: " & mod_UserCacheSync.GetUserField("CNA_HandlowcaID") _
        & ", oddział: " & mod_UserCacheSync.GetUserField("NrOddzialu") & ")"
    
    ' Pokaż info o roli
    If mod_UserCacheSync.IsUserManager() Then
        lbl_RoleInfo.Caption = "Tryb: KIEROWNIK (wnioski wysyłane wprost do BNC)"
    Else
        lbl_RoleInfo.Caption = "Tryb: HANDLOWIEC (wnioski wysyłane do kierownika)"
    End If
    
    ' Domyślny miesiąc zgłoszenia = bieżący
    txt_MiesiacZgloszenia.Text = Format(mod_Utils.GetCurrentMonthYear(), "yyyy-mm")
    
    ' Załaduj listę pending
    RefreshPendingList
End Sub

Private Sub btn_AddToList_Click()
    ' 1. Zbierz dane z pól
    Dim reportData As Object
    Set reportData = CreateObject("Scripting.Dictionary")
    reportData("KlientFK") = Trim(txt_KlientFK.Text)
    reportData("NazwaKlienta") = Trim(txt_NazwaKlienta.Text)
    reportData("MiesiacZgloszenia") = Trim(txt_MiesiacZgloszenia.Text)
    reportData("Fields") = Trim(txt_Fields.Text)
    
    ' 2. Walidacja
    Dim errMsg As String
    errMsg = mod_Validation.ValidateReportData(reportData)
    If errMsg <> "" Then
        MsgBox errMsg, vbExclamation, "Błąd walidacji"
        Exit Sub
    End If
    
    ' 3. Zapisz do ws_DataCache (status = pending)
    Dim newID As Long
    newID = mod_DataCacheSync.AppendRecord(reportData)
    
    mod_Utils.LogInfo "Dodano zgłoszenie ID=" & newID
    
    ' 4. Wyczyść formularz, odśwież listę
    ClearFormFields
    RefreshPendingList
End Sub

Private Sub btn_Clear_Click()
    ClearFormFields
End Sub

Private Sub btn_SendBatch_Click()
    ' Sprawdź czy są pending recordy
    Dim pending As Collection
    Set pending = mod_DataCacheSync.GetPendingRecords()
    
    If pending.Count = 0 Then
        MsgBox "Brak zgłoszeń do wysłania.", vbInformation
        Exit Sub
    End If
    
    ' Potwierdzenie
    Dim msg As String
    If mod_UserCacheSync.IsUserManager() Then
        msg = "Wysłać " & pending.Count & " zgłoszeń wprost do BNC?"
    Else
        msg = "Wysłać " & pending.Count & " zgłoszeń do kierownika " _
            & "(" & mod_UserCacheSync.GetUserField("EmailKierownika") & ")?"
    End If
    
    If MsgBox(msg, vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    
    ' Delegujemy wysyłkę do mod_MailSender (M4)
    Dim success As Boolean
    success = mod_MailSender.SendBatch()
    
    If success Then
        MsgBox "Batch wysłany pomyślnie.", vbInformation
        RefreshPendingList
    Else
        MsgBox "Błąd wysyłki. Sprawdź log.", vbExclamation
    End If
End Sub

Private Sub btn_ShowLog_Click()
    Me.Hide
    frm_Log.Show
End Sub

' ----- Helpers -----

Private Sub ClearFormFields()
    txt_KlientFK.Text = ""
    txt_NazwaKlienta.Text = ""
    txt_Fields.Text = ""
    ' MiesiacZgloszenia zostaje (user może dodać kilka zgłoszeń za ten sam miesiąc)
    txt_KlientFK.SetFocus
End Sub

Private Sub RefreshPendingList()
    Dim pending As Collection
    Set pending = mod_DataCacheSync.GetPendingRecords()
    
    lbl_BatchCount.Caption = "Lista zgłoszeń do wysłania (" & pending.Count & ")"
    
    lst_PendingBatch.Clear
    
    Dim record As Object
    For Each record In pending
        ' ListBox z 5 kolumnami
        lst_PendingBatch.AddItem record("ReportID")
        lst_PendingBatch.List(lst_PendingBatch.ListCount - 1, 1) = record("KlientFK")
        lst_PendingBatch.List(lst_PendingBatch.ListCount - 1, 2) = record("NazwaKlienta")
        lst_PendingBatch.List(lst_PendingBatch.ListCount - 1, 3) = record("MiesiacZgloszenia")
        lst_PendingBatch.List(lst_PendingBatch.ListCount - 1, 4) = record("Fields")
    Next record
End Sub
```

### Zadania M3.1

- [ ] **3.1.1** Utwórz UserForm `frm_Main`
- [ ] **3.1.2** Dodaj wszystkie kontrolki zgodnie z layoutem
- [ ] **3.1.3** Ustaw ListBox: ColumnCount=5, ColumnHeads=True, ColumnWidths
- [ ] **3.1.4** Zaimplementuj cały kod
- [ ] **3.1.5** Test manualny:
  - Otwórz aplikację (już po setupie)
  - Wpisz testowe zgłoszenie, kliknij Dodaj
  - Pojawia się na liście
  - Dodaj kolejne 2-3
  - Sprawdź `ws_DataCache` i `BNC_DataCache.xlsx`
- [ ] **3.1.6** Eksport `Source/Forms/frm_Main.frm` + `.frx`
- [ ] **3.1.7** Git commit

### Kryterium akceptacji M3

- [ ] Możesz dodać minimum 5 zgłoszeń do batcha
- [ ] Lista pending odświeża się po każdym dodaniu
- [ ] `ws_DataCache` ma poprawne wpisy (sprawdzone przez tymczasowe odsłonięcie arkusza)
- [ ] `BNC_DataCache.xlsx` jest zsynchronizowany
- [ ] Walidacja działa dla każdego pola
- [ ] Info o roli (kierownik/handlowiec) wyświetla się poprawnie

---

## M4 - Mail sender

> **Czas**: ~3 dni  
> **Wynik**: Funkcjonalna wysyłka batcha mailem przez Outlook

### M4.1 - mod_MailSender

**Po co istnieje**: Cała logika wysyłki maila, z generowaniem pliku tymczasowego, decision diamond i UPDATE statusu po wysyłce.

**Public API**:

```vba
' ==================== mod_MailSender ====================

' Główna funkcja - wysłanie aktualnego batcha
Public Function SendBatch() As Boolean
    ' Zwraca True jeśli wysłano, False jeśli error
    ' Wewnątrz: cała logika z diagramu data flow Flow B
End Function

' Helpers (private)

Private Function GenerateTempFile(records As Collection) As String
    ' Tworzy plik xlsx w %TEMP% z tabelą records
    ' Zwraca: pełna ścieżka do pliku
End Function

Private Sub SendMailWithAttachment(recipient As String, _
                                    subject As String, _
                                    body As String, _
                                    attachmentPath As String)
    ' Wysyła mail przez Outlook COM
End Sub

Private Sub CleanupTempFile(filePath As String)
    ' Bezpiecznie usuwa plik tymczasowy
End Sub

' Funkcja decyzyjna - kierownik vs handlowiec
Private Function DetermineRecipient() As Object
    ' Zwraca Dictionary z polami: To, Body, Subject
End Function
```

**Pełny pseudokod `SendBatch`**:

```vba
Public Function SendBatch() As Boolean
    Dim tempFilePath As String
    Dim recipientInfo As Object
    Dim pending As Collection
    Dim sentIDs As Collection
    
    On Error GoTo ErrorHandler
    
    ' 1. Pobierz pending
    Set pending = mod_DataCacheSync.GetPendingRecords()
    If pending.Count = 0 Then
        SendBatch = False
        Exit Function
    End If
    
    ' 2. Wygeneruj plik tymczasowy
    tempFilePath = GenerateTempFile(pending)
    mod_Utils.LogInfo "Wygenerowano plik tymczasowy: " & tempFilePath
    
    ' 3. Decision diamond - kierownik vs handlowiec
    Set recipientInfo = DetermineRecipient()
    
    ' 4. Wyślij mail
    SendMailWithAttachment recipientInfo("To"), _
                            recipientInfo("Subject"), _
                            recipientInfo("Body"), _
                            tempFilePath
    
    ' 5. UPDATE Status na 'sent'
    Set sentIDs = New Collection
    Dim record As Object
    For Each record In pending
        sentIDs.Add record("ReportID")
    Next record
    
    mod_DataCacheSync.MarkAsSent sentIDs, recipientInfo("To")
    
    ' 6. Cleanup
    CleanupTempFile tempFilePath
    
    SendBatch = True
    Exit Function

ErrorHandler:
    mod_Utils.LogError "mod_MailSender.SendBatch", Err.Number, Err.Description
    
    ' Cleanup nawet po błędzie
    If tempFilePath <> "" Then CleanupTempFile tempFilePath
    
    SendBatch = False
End Function

Private Function DetermineRecipient() As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    
    Dim emailHandlowca As String, emailKierownika As String, emailBNC As String
    emailHandlowca = mod_UserCacheSync.GetUserField("EmailHandlowca")
    emailKierownika = mod_UserCacheSync.GetUserField("EmailKierownika")
    emailBNC = mod_UserCacheSync.GetUserField("EmailBNC")
    
    If mod_UserCacheSync.IsUserManager() Then
        ' KIEROWNIK - wprost do BNC
        result("To") = emailBNC
        result("Subject") = "Wniosek BNC - " & Format(Now(), "yyyy-mm-dd")
        result("Body") = "Załączam wniosek BNC. Proszę o weryfikację."
    Else
        ' HANDLOWIEC - do kierownika
        result("To") = emailKierownika
        result("Subject") = "Wniosek BNC do akceptacji - " & Format(Now(), "yyyy-mm-dd")
        result("Body") = "Załączam wniosek BNC. Proszę o weryfikację " _
                       & "i przekazanie do " & emailBNC & "."
    End If
    
    Set DetermineRecipient = result
End Function

Private Function GenerateTempFile(records As Collection) As String
    Dim tempPath As String
    Dim fileName As String
    Dim fullPath As String
    
    tempPath = Environ("TEMP")
    fileName = "BNC_Wniosek_" & Format(Now(), "yyyymmdd_hhmmss") & ".xlsx"
    fullPath = tempPath & "\" & fileName
    
    ' Utwórz nowy workbook
    Dim wb As Workbook
    Set wb = Workbooks.Add
    
    Dim ws As Worksheet
    Set ws = wb.Sheets(1)
    
    ' Nagłówki
    ws.Cells(1, 1).Value = "ReportID"
    ws.Cells(1, 2).Value = "KlientFK"
    ws.Cells(1, 3).Value = "NazwaKlienta"
    ws.Cells(1, 4).Value = "CNA_HandlowcaID"
    ws.Cells(1, 5).Value = "NrOddzialu"
    ws.Cells(1, 6).Value = "MiesiacZgloszenia"
    ws.Cells(1, 7).Value = "Fields"
    ws.Cells(1, 8).Value = "CreatedTimestamp"
    
    ' Dane
    Dim row As Long
    row = 2
    Dim record As Object
    For Each record In records
        ws.Cells(row, 1).Value = record("ReportID")
        ws.Cells(row, 2).Value = record("KlientFK")
        ws.Cells(row, 3).Value = record("NazwaKlienta")
        ws.Cells(row, 4).Value = record("CNA_HandlowcaID")
        ws.Cells(row, 5).Value = record("NrOddzialu")
        ws.Cells(row, 6).Value = record("MiesiacZgloszenia")
        ws.Cells(row, 7).Value = record("Fields")
        ws.Cells(row, 8).Value = record("CreatedTimestamp")
        row = row + 1
    Next record
    
    ' Auto-fit columns dla czytelności
    ws.Columns.AutoFit
    
    ' SaveAs xlsx
    Application.DisplayAlerts = False
    wb.SaveAs Filename:=fullPath, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    wb.Close SaveChanges:=False
    
    GenerateTempFile = fullPath
End Function

Private Sub SendMailWithAttachment(recipient As String, subject As String, _
                                    body As String, attachmentPath As String)
    Dim outlookApp As Object
    Dim mailItem As Object
    
    Set outlookApp = CreateObject("Outlook.Application")
    Set mailItem = outlookApp.CreateItem(0)  ' 0 = olMailItem
    
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

Private Sub CleanupTempFile(filePath As String)
    On Error Resume Next  ' Ignore errors w cleanup
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FileExists(filePath) Then
        fso.DeleteFile filePath, True  ' True = force
    End If
    
    On Error GoTo 0
End Sub
```

### Zadania M4.1

- [ ] **4.1.1** Utwórz moduł `mod_MailSender`
- [ ] **4.1.2** Zaimplementuj wszystkie funkcje
- [ ] **4.1.3** Test scenariusza handlowca:
  - Setup: EmailKierownika ≠ EmailHandlowca
  - Dodaj 2-3 zgłoszenia
  - Kliknij Wyślij Wniosek BNC
  - Sprawdź że mail przyszedł do kierownika z body "do przekazania"
  - Sprawdź że status zmienił się na sent z EmailRecipient = email kierownika
- [ ] **4.1.4** Test scenariusza kierownika:
  - Setup: EmailKierownika == EmailHandlowca
  - Dodaj 2-3 zgłoszenia
  - Kliknij Wyślij Wniosek BNC
  - Sprawdź że mail przyszedł do BNC z body "do weryfikacji"
  - Sprawdź EmailRecipient = email BNC
- [ ] **4.1.5** Test scenariusza błędu:
  - Wyłącz Outlook, kliknij Wyślij
  - Sprawdź że dostajesz komunikat błędu, plik tymczasowy nie zostaje na dysku, status pending nie został zmieniony
- [ ] **4.1.6** Eksport + commit

### Kryterium akceptacji M4

- [ ] Wysyłka działa w obu trybach (kierownik i handlowiec)
- [ ] Plik tymczasowy w `%TEMP%` powstaje, jest wysyłany jako załącznik, jest usuwany po wysyłce
- [ ] Status w `ws_DataCache` zmienia się na sent z poprawnym EmailRecipient
- [ ] Po błędzie aplikacja loguje, ale nie crashuje
- [ ] Po wysyłce można dodać kolejne zgłoszenia (zaczynają nowy batch)

---

## M5 - Log + Export

> **Czas**: ~2 dni  
> **Wynik**: Ekran historii + funkcja eksportu BNC_DataCache.xlsx

### M5.1 - mod_Export

**Po co istnieje**: Literal copy `BNC_DataCache.xlsx` do wybranej lokalizacji.

```vba
' ==================== mod_Export ====================

Public Function ExportDataCache(targetPath As String) As Boolean
    ' Kopiuje BNC_DataCache.xlsx do targetPath
    ' targetPath: pełna ścieżka docelowa (włącznie z nazwą pliku)
    ' Zwraca: True jeśli sukces, False jeśli błąd
End Function

Public Function GetSuggestedExportFileName() As String
    ' Generuje nazwę typu BNC_Eksport_Kowalski_2026-05-05.xlsx
End Function
```

### Zadania M5.1

- [ ] **5.1.1** Utwórz moduł `mod_Export`
- [ ] **5.1.2** Implementacja:
  ```vba
  Public Function ExportDataCache(targetPath As String) As Boolean
      Dim sourcePath As String
      sourcePath = mod_UserCacheSync.GetUserField("CacheFolderPath") & "\BNC_DataCache.xlsx"
      
      If Not mod_Utils.FileExists(sourcePath) Then
          mod_Utils.LogError "mod_Export.ExportDataCache", 0, _
              "Plik źródłowy nie istnieje: " & sourcePath
          ExportDataCache = False
          Exit Function
      End If
      
      On Error GoTo ErrorHandler
      
      Dim fso As Object
      Set fso = CreateObject("Scripting.FileSystemObject")
      fso.CopyFile sourcePath, targetPath, True  ' True = overwrite
      
      ExportDataCache = True
      Exit Function
  
  ErrorHandler:
      mod_Utils.LogError "mod_Export.ExportDataCache", Err.Number, Err.Description
      ExportDataCache = False
  End Function
  ```
- [ ] **5.1.3** Test
- [ ] **5.1.4** Eksport + commit

### M5.2 - frm_Log

**Layout**:

```
┌──────────────────────────────────────────────────────────────────────┐
│ BNC_Sender - Historia zgłoszeń                                   [X] │
├──────────────────────────────────────────────────────────────────────┤
│                                                                      │
│ Wszystkich zgłoszeń: 47   |   Pending: 3   |   Sent: 44              │
│                                                                      │
│ ┌────────────────────────────────────────────────────────────────┐   │
│ │ ID │ KlientFK │ Nazwa     │ Status  │ Wysłany do      │ Data  │   │
│ ├────┼──────────┼───────────┼─────────┼─────────────────┼───────┤   │
│ │ 47 │ 12345    │ Acme      │ pending │ -               │ -     │   │
│ │ 46 │ 67890    │ Bravo     │ sent    │ kier@firma.pl   │ 05-05 │   │
│ │ 45 │ 11111    │ Charlie   │ sent    │ kier@firma.pl   │ 05-05 │   │
│ │... │          │           │         │                 │       │   │
│ └────────────────────────────────────────────────────────────────┘   │
│                                                                      │
│       [Eksportuj do pliku]  [Powrót do formularza]                   │
└──────────────────────────────────────────────────────────────────────┘
```

**Kontrolki**:
- Label: `lbl_Stats`
- ListBox: `lst_AllRecords`
- Button: `btn_Export`, `btn_Back`

**Kod**:

```vba
' ==================== frm_Log ====================

Private Sub UserForm_Initialize()
    LoadRecords
End Sub

Private Sub btn_Export_Click()
    ' 1. Sugeruj nazwę pliku
    Dim suggestedName As String
    suggestedName = mod_Export.GetSuggestedExportFileName()
    
    ' 2. Pokaż Save As dialog
    Dim targetPath As String
    targetPath = ShowSaveDialog(suggestedName)
    
    If targetPath = "" Then Exit Sub  ' User anulował
    
    ' 3. Eksportuj
    If mod_Export.ExportDataCache(targetPath) Then
        MsgBox "Wyeksportowano do: " & targetPath, vbInformation
    Else
        MsgBox "Błąd eksportu. Sprawdź log.", vbExclamation
    End If
End Sub

Private Sub btn_Back_Click()
    Me.Hide
    frm_Main.Show
End Sub

Private Sub LoadRecords()
    Dim allRecords As Collection
    Set allRecords = mod_DataCacheSync.GetAllRecords()
    
    Dim pendingCount As Long, sentCount As Long
    pendingCount = 0
    sentCount = 0
    
    lst_AllRecords.Clear
    
    Dim record As Object
    For Each record In allRecords
        ' Liczenie statystyk
        If record("Status") = "pending" Then
            pendingCount = pendingCount + 1
        Else
            sentCount = sentCount + 1
        End If
        
        ' Dodaj do listy (najnowsze na górze)
        lst_AllRecords.AddItem record("ReportID")
        ' ... reszta kolumn ...
    Next record
    
    lbl_Stats.Caption = "Wszystkich: " & allRecords.Count _
        & "   |   Pending: " & pendingCount _
        & "   |   Sent: " & sentCount
End Sub

Private Function ShowSaveDialog(defaultName As String) As String
    ' Application.GetSaveAsFilename(...)
    Dim result As Variant
    result = Application.GetSaveAsFilename( _
        InitialFileName:=defaultName, _
        FileFilter:="Excel Files (*.xlsx), *.xlsx")
    
    If result = False Then
        ShowSaveDialog = ""
    Else
        ShowSaveDialog = CStr(result)
    End If
End Function
```

### Zadania M5.2

- [ ] **5.2.1** Utwórz UserForm `frm_Log`
- [ ] **5.2.2** Dodaj kontrolki
- [ ] **5.2.3** Zaimplementuj kod
- [ ] **5.2.4** Test:
  - Otwórz frm_Log
  - Widać statystyki i listę
  - Kliknij Eksportuj, wybierz lokalizację, sprawdź czy plik powstał
  - Powrót do frm_Main
- [ ] **5.2.5** Eksport + commit

### Kryterium akceptacji M5

- [ ] Wszystkie zgłoszenia (pending + sent) wyświetlają się w log
- [ ] Statystyki są poprawne
- [ ] Eksport tworzy poprawny plik xlsx z tymi samymi danymi co BNC_DataCache.xlsx
- [ ] Powrót do frm_Main działa

---

## M6 - Polish + UAT

> **Czas**: ~3 dni  
> **Wynik**: Wersja v0.9.0 RC (Release Candidate) gotowa do testów z prawdziwymi userami

### M6.1 - Error handling i UX polish

- [ ] **6.1.1** Każda procedura `Public` ma `On Error GoTo ErrorHandler`
- [ ] **6.1.2** Komunikaty błędów dla usera są **zrozumiałe** (nie `Run-time error 91`, tylko `Brak skonfigurowanego adresu kierownika - przejdź do setupu`)
- [ ] **6.1.3** Loading indicator przy długich operacjach (`Application.Cursor = xlWait`)
- [ ] **6.1.4** Disable przycisków podczas operacji (żeby user nie kliknął 3x Wyślij)
- [ ] **6.1.5** Default values w formularzach (np. dzisiejsza data)
- [ ] **6.1.6** Tab order na formularzach (kolejność wciśnięcia Tab)
- [ ] **6.1.7** Akcelaratory klawiszowe (Alt+S = Save, Alt+A = Anuluj)

### M6.2 - User Acceptance Testing

- [ ] **6.2.1** Wybierz 3-5 testowych userów (znajomi z zespołu, którzy zechcą pomóc)
- [ ] **6.2.2** Przygotuj im pakiet:
  - Plik `BNC_Sender_v0.9.0.xlsm`
  - Krótka instrukcja "co testować" (1 strona)
  - Sposób raportowania bugów (np. mail do Ciebie z opisem)
- [ ] **6.2.3** Daj im **minimum 1 tydzień** na używanie
- [ ] **6.2.4** Zbierz feedback, zrób listę bugów i sugestii
- [ ] **6.2.5** Napraw krytyczne bugi (te, które blokują pracę)
- [ ] **6.2.6** Sugestie kosmetyczne dodaj do `TODO.md` (na fazę B lub patch v1.0.1)

### Kryterium akceptacji M6

- [ ] Każdy testowy user mógł przejść przez setup → dodać zgłoszenie → wysłać batch → zobaczyć w log → wyeksportować
- [ ] Zero krytycznych bugów (które blokują pracę)
- [ ] Maksymalnie 3 mniejsze bugi do naprawy w v1.0.0 (reszta na patch)

---

## M7 - Release v1.0.0

- [ ] **7.1** Bump version: zmień nazwę pliku na `BNC_Sender_v1.0.0.xlsm`
- [ ] **7.2** Dopisz do `CHANGELOG.md`:
  ```markdown
  ## v1.0.0 (2026-XX-XX) - First Production Release
  - Pełna funkcjonalność fazy A
  - 3 UserForms, 6 modułów, 2 ukryte arkusze
  - Workflow kierownik vs handlowiec
  - Hybrid cache (worksheet + xlsx backup)
  - Eksport literal copy
  ```
- [ ] **7.3** Eksport finalny do `Source/`
- [ ] **7.4** Git tag: `git tag -a v1.0.0 -m "First production release"`
- [ ] **7.5** Skopiuj plik do `Releases/BNC_Sender_v1.0.0.xlsm`
- [ ] **7.6** Wgraj plik na OneDrive firmowy
- [ ] **7.7** Wyślij komunikat do userów z linkiem do pobrania i instrukcją

---

## 12. Dziennik decyzji architektonicznych

W trakcie implementacji **na bieżąco** dodawaj wpisy do `Notatki/DECISIONS.md`. Format:

```markdown
## ADR-NNN: <tytuł decyzji>

**Data**: 2026-XX-XX  
**Status**: Accepted / Superseded by ADR-XXX

**Kontekst**: Co spowodowało tę decyzję?

**Decyzja**: Co konkretnie wybrałem?

**Alternatywy**: Co rozważałem i dlaczego odrzuciłem?

**Konsekwencje**: Jakie są skutki tej decyzji?
```

Już teraz masz kilka ADR do zapisania:
- ADR-001: Repository Pattern dla cache (z M1)
- ADR-002: Workbook_Open jako entry point (z M2)
- ADR-003: Convention over configuration dla detekcji roli usera (już omawiane)
- ADR-004: Hybrid cache z worksheet jako primary (już omawiane)
- ADR-005: Plik tymczasowy w %TEMP% z literal copy do mail attachment

---

## 13. Lista kontrolna gotowości v1.0.0

### Funkcjonalność

- [ ] frm_Setup pozwala zarejestrować się jednorazowo
- [ ] Samouczek wyświetla się przy pierwszym setupie, ma checkbox "nie pokazuj"
- [ ] frm_Main wyświetla informację o roli (kierownik/handlowiec)
- [ ] Można dodać dowolną liczbę zgłoszeń do batcha
- [ ] Walidacja działa dla każdego pola
- [ ] Wysyłka batcha działa w trybie kierownika (wprost do BNC)
- [ ] Wysyłka batcha działa w trybie handlowca (do kierownika)
- [ ] Treść body jest poprawna dla każdego trybu
- [ ] Status pending → sent z poprawnym EmailRecipient
- [ ] frm_Log pokazuje wszystkie zgłoszenia z statystykami
- [ ] Eksport literal copy działa

### Architektura

- [ ] 7 modułów VBA (mod_Utils, mod_Validation, mod_UserCacheSync, mod_DataCacheSync, mod_MailSender, mod_Export + ThisWorkbook)
- [ ] 3 UserForms (frm_Setup, frm_Main, frm_Log)
- [ ] 2 ukryte arkusze (ws_UserCache, ws_DataCache - very hidden)
- [ ] 1 widoczny arkusz placeholder (Sheet1)
- [ ] Reguła warstw nie złamana (UserForms → moduły logiki → moduły sync → worksheets)
- [ ] Repository Pattern poprawnie zaimplementowany

### Niezawodność

- [ ] Każda Public procedura ma error handling
- [ ] Komunikaty błędów dla usera są zrozumiałe
- [ ] Plik tymczasowy w %TEMP% zawsze jest usuwany (nawet po błędzie)
- [ ] Auto-recreate plików cache działa
- [ ] Aplikacja nie crashuje przy braku Outlooka
- [ ] Aplikacja nie crashuje przy braku folderu cache

### Repo

- [ ] Wszystkie moduły wyeksportowane do `Source/`
- [ ] Git history jest czysta (sensowne commit messages)
- [ ] Tag v1.0.0 utworzony
- [ ] CHANGELOG.md aktualny
- [ ] DECISIONS.md zawiera minimum 5 ADR
- [ ] README.md opisuje jak uruchomić projekt z zera

### Testy z userami

- [ ] Minimum 3 userów testowało aplikację przez tydzień
- [ ] Zero krytycznych bugów
- [ ] Pozytywny feedback od minimum 80% testowych userów

---

**Powodzenia w implementacji.** Pamiętaj: **build small, prove value, then scale**. 🚀
