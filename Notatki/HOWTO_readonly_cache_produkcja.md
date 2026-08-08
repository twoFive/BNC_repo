# HOWTO — read-only cache xlsx w wersji produkcyjnej

> **Status**: **DEFERRED** — implementacja odłożona do fazy produkcyjnej (M7 rollout do 20 handlowców).
> **Nie implementować w dev phase** (Faza A) — nie ma production userów, ochrona bez wartości.

---

## 🎯 Co to jest

Trzy pliki cache xlsx generowane przez aplikację w `CacheFolderPath` (`C:\BNC_CacheFolder\`):

- `BNC_UsersRegistry.xlsx` — lista wszystkich userów (post-ADR-009 zastąpił BNC_UserCache.xlsx jako "user identity storage")
- `BNC_DataCache.xlsx` — historia wszystkich zgłoszeń handlowca
- `BNC_UsersRegistry.xlsx` — lista wszystkich userów (post-M3.3 symetria)

Wszystkie są **backupami** stanu z ukrytych arkuszy w xlsm. Sync jest **jednostronny** (`ws → xlsx`, ADR-002) — jeśli user edytuje xlsx bezpośrednio, następny sync z aplikacji nadpisze zmiany. Ale:
- User może otworzyć xlsx w Excelu, zobaczyć sensitive data (adresy, CNA, historia zgłoszeń)
- User może **przypadkowo** skasować/uszkodzić plik → następny sync go regeneruje, ale przez chwilę brak backupu
- Explorer preview pane trzyma handle → sync fail w międzyczasie

W wersji produkcyjnej warto **oznaczyć pliki jako read-only** — sygnał "nie edytuj, to backup aplikacji, nie danych do modyfikacji".

## 🛡 Dwie warstwy ochrony (opcjonalne)

### Warstwa 1 — Filesystem read-only attribute (`attrib +R`)

**Co daje**: Windows oznacza plik jako read-only w Explorer i większości edytorów. Excel przy otwarciu pokazuje "Read-Only" w tytule paska. User musi świadomie zdjąć atrybut żeby edytować (`attrib -R` lub Properties → Read-only unchecked).

**Ograniczenia**:
- Zaawansowany user może odznaczyć w Properties → 2 kliki
- Antivirus czasem blokuje `SetFileAttribute` — trzeba obsłużyć błąd
- **Sync z aplikacji się nie uda** dopóki nie zdejmiemy atrybutu przed `SaveAs` → nakłada się kod:
  1. `SetFileAttribute cacheFile, vbNormal` (przed każdym `SaveAs`)
  2. `SaveAs` / `Kill` + recreate
  3. `SetFileAttribute cacheFile, vbReadOnly` (po sukcesie)
  4. Wszystko w `On Error` żeby AV/lock nie zablokował flow

### Warstwa 2 — Excel workbook structure protect (opcjonalnie)

**Co daje**: Password-protect struktury workbook — user nie może dodać/usunąć/rename sheets. Widoczne w tytule "(Protected View)".

**Kiedy stosować**: jeśli chcemy dodatkowo chronić przed userem próbującym otworzyć backup i "sobie coś dopisać".

**Ograniczenia**: hasło jest w plaintext w kodzie VBA (albo hardcoded, albo hash — łamalne). Real security zero. Ale sygnał intencji **strong**.

## 📋 Implementacja (referencyjny szkielet — DO NOT COPY YET)

W `mod_UsersRegistrySync.SyncRegistryToFile`, `mod_DataCacheSync.SyncToFile` (2 sync points post-ADR-009):

```vba
' Przed SaveAs — zdejmij read-only jeśli był
On Error Resume Next
SetAttr fullPath, vbNormal
On Error GoTo Cleanup

' istniejący SaveAs flow...
If mod_Utils.FileExists(fullPath) Then Kill fullPath
wbOut.SaveAs Filename:=fullPath, FileFormat:=xlOpenXMLWorkbook
wbOut.Close SaveChanges:=False

' Po SaveAs — nałóż read-only
On Error Resume Next
SetAttr fullPath, vbReadOnly
On Error GoTo 0
```

Dodać w 3 miejscach (analogicznie w każdym `SyncToFile`).

## ⏭ Kiedy zaimplementować

**Trigger**: pre-M7 UAT (User Acceptance Testing z ~3-5 handlowcami).

**Kroki**:
1. Implementacja w 3 modułach sync (~20 linii w każdym)
2. Regression test — modyfikuj cache xlsx, uruchom SendBatch, verify sync overrode change
3. Doc update: DECISIONS.md ADR-002 rozszerzenie ("Konsekwencja: cache xlsx read-only na dysku, aplikacja zdejmuje atrybut przed sync")
4. Bump wersji xlsm (M7 release)
5. Deploy do userów UAT
6. Feedback loop: czy read-only nie przeszkadza w recovery scenarios?

## 🚫 Dlaczego NIE teraz

- **Zero production data** — nic do ochrony
- **Zero production userów** — nikt nie kliknie w cache przez pomyłkę
- **Development friction** — musisz zdejmować atrybut ręcznie żeby zrobić `Kill xlsx` w Immediate przy debugowaniu (patrz Schema v2 migration — akurat robisz operacje na xlsx z Immediate)
- **YAGNI** — kod martwy do M7, dodałby ~60 linii bez wartości teraz

## 🔗 Related

- [`DECISIONS.md`](DECISIONS.md) — ADR-002 (Sync bez clipboard, jednostronny)
- [`DECISIONS.md`](DECISIONS.md) — ADR-008 (Multi-user Registry — Persistence sekcja)
- [`TODO.md`](TODO.md) — dodaj task **M7.X: cache xlsx read-only lock** żeby nie zapomnieć
