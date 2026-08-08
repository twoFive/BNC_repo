# HOWTO — zmiana adresu email zespołu BNC

> Adres BNC to **hardcoded constant** (polityka ADR-003 — user nie ma możliwości edytować, redukcja odpowiedzialności + spójność audytu).
> Zmiana wymaga edycji kodu + re-import modułu.

---

## 🎯 Miejsce zmiany — 1 constant

**Plik**: [`Source/Forms/frm_Setup.code-behind.txt`](../Source/Forms/frm_Setup.code-behind.txt)
**Linia**: ~7

```vba
Private Const HARDCODED_EMAIL_BNC As String = "NOWY_ADRES@domena.pl"
```

To jedyne miejsce w kodzie **executable**, które trzeba zmienić. `mod_MailSender` **nie hardcode'uje** adresu — czyta go z `ws_UsersRegistry` przez `mod_UsersRegistrySync.GetCurrentUserField("EmailBNC")`, gdzie trafia z powyższej stałej podczas `frm_Setup.btn_Save_Click` → `AddNewUser`.

## 🎯 Opcjonalnie — sync spec kontrolki (doc, bez wpływu na runtime)

**Plik**: [`Source/Forms/frm_Setup.LAYOUT.md`](../Source/Forms/frm_Setup.LAYOUT.md)
**Linia**: ~52 (tabela z properties `txt_EmailBNC`)

Zaktualizuj `Text = "..."` na nowy adres — kopia dla czytelności LAYOUT.md, nie wpływa na kompilację.

---

## 🔁 Kroki wykonania (5 minut)

1. Edytuj `frm_Setup.code-behind.txt` — podmień wartość constant
2. Otwórz VBE (`Alt+F11`)
3. Project Explorer → prawy klik `frm_Setup` → **Remove frm_Setup** → **No** (nie eksportuj)
4. **File → Import File...** → wybierz `frm_Setup.frm` (jeśli masz eksportowany) LUB odbuduj kontrolki + wklej code-behind ręcznie
5. **Debug → Compile VBAProject** (`Alt+D` → `L`) — verify że kompiluje
6. `Ctrl+S` (save xlsm)

### Alternatywa jeśli nie chcesz Remove/Import całego formularza

- VBE → dwuklik `frm_Setup` → View Code → znajdź `HARDCODED_EMAIL_BNC` → zmień wartość w miejscu
- `Ctrl+S`

To najszybsza droga jeśli edytujesz bezpośrednio w VBE.

---

## ⚠ Stary adres pozostaje w:

Zgodnie z Twoim rozumieniem — implikacje pomijam, tylko lista:

1. **`ws_UsersRegistry`** — kolumna 8 `EmailBNC` per user (jedyne źródło prawdy post-ADR-009).
2. **`BNC_UsersRegistry.xlsx`** — write-through backup na dysku.
3. **`ws_DataCache.EmailRecipient`** — kolumna 9 dla rekordów wysłanych **w trybie kierownika** (KIEROWNIK → `To = EmailBNC` → rzeczywisty adresat zapisany w audit trail).
4. **`BNC_DataCache.xlsx`** — write-through backup na dysku.

Sent records są **immutable** (ADR-006). Wpisy z zapisanym starym adresem odzwierciedlają rzeczywistość wysyłki w momencie zdarzenia — dowód historyczny.

---

## 🆕 Po zmianie — verify

```
? mod_UsersRegistrySync.GetCurrentUserField("EmailBNC")
```

Zwraca **stary** adres (bo Registry row current usera ma starą wartość zapisaną wcześniej). Constant zadziała dopiero dla:
- **Nowego setupu** (`frm_Setup.btn_Save_Click` przez `AddNewUser`)
- **Force update** existing usera (patrz sekcja poniżej — jeśli potrzeba)

## 🔨 Force update istniejących userów (opcjonalne)

Jeśli chcesz żeby **wszyscy** existing userzy zaczęli używać nowego adresu — Immediate:

```vba
Dim NEW_BNC As String: NEW_BNC = "NOWY_ADRES@domena.pl"
Dim wsReg As Worksheet: Set wsReg = ThisWorkbook.Worksheets("ws_UsersRegistry")
Dim r As Long, last As Long
last = wsReg.Cells(wsReg.Rows.Count, 1).End(xlUp).Row
For r = 2 To last: wsReg.Cells(r, 8).Value = NEW_BNC: Next
mod_UsersRegistrySync.SetCurrentUserField "EmailBNC", NEW_BNC
ThisWorkbook.Save
```

Aktualizuje:
- Wszystkie wiersze w `ws_UsersRegistry` (kolumna 8) — sole source of truth
- Bieżącego usera bezpośrednio (bo Registry = current user storage post-ADR-009)
- **NIE** aktualizuje sent records w `ws_DataCache` (audit trail immutable)

Bez tego skryptu — dopiero nowi userzy dostaną nowy adres, starzy nadal używają zapisanej wartości aż do manualnego `SetUserField`.
