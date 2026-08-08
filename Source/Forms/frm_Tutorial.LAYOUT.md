# frm_Tutorial - specyfikacja layoutu (v3)

> **Cel**: interaktywny wizard z 5 stron wprowadzający usera do BNC_Sender.
> **Wywołanie**: **wyłącznie manualne** - user musi kliknąć żeby zobaczyć samouczek:
> - `frm_Setup.btn_ShowTutorial_Click` (zastępuje obecny placeholder MsgBox)
> - Opcjonalnie: nowy `frm_Main.btn_Help` (M7 nice-to-have)
>
> **Brak auto-show** - user zawsze świadomie wywołuje. Bez persistencji preferencji "nie pokazuj więcej".
>
> **Content source of truth**: [`frm_Tutorial.TEXT.md`](frm_Tutorial.TEXT.md) - code-behind czyta stąd (hardcoded string constants). Zmiana treści = update TEXT.md najpierw, potem re-paste code-behind.
>
> **Plan**: M6/M7

---

## Krok 1 - utwórz UserForm |x

W VBE: **Insert → UserForm**. W oknie Properties (`F4`):

| Property | Wartość |
|---|---|
| `(Name)` | `frm_Tutorial` |
| `Caption` | `BNC_Sender - Samouczek` |
| `Width` | `564` |
| `Height` | `444` |
| `StartUpPosition` | `1 - CenterOwner` |
| `BackColor` | `&H8000000F&` (BtnFace - domyślne szare) |

Wysokość 444 pt = ~340 pt użytecznej wysokości (po odjęciu title bar Excela).

---

## Krok 2 - kontrolki (z dokładnym positioning)

Współrzędne w punktach (pt). Form area: 558 × 420 (bez title bar).

### Sekcja 1: Header (y=6 do y=44, height=38pt)

| # | Typ | Name | Left | Top | Width | Height | Properties |
|---|---|---|---:|---:|---:|---:|---|
| 1 | Label | `lbl_TutorialHeader` | 12 | 12 | 380 | 22 | `Caption = "Samouczek BNC_Sender"`, `Font = Segoe UI 14 Bold`, `ForeColor = &H00804000&` (ciemnoniebieski) |x
| 2 | Label | `lbl_PageIndicator` | 420 | 16 | 130 | 18 | `Caption` wypełnia `RenderPage` (np. `"Strona 3 z 5"`), `Font = 9pt italic`, `TextAlign = 3 - fmTextAlignRight`, `ForeColor = &H00666666&` (medium gray) |x

**Separator visualny** (opcjonalny): Label z `Height = 1`, `BackColor = &H00CCCCCC&`, position (12, 46), width=536 - tworzy poziomą linię pod headerem.

### Sekcja 2: Body (y=54 do y=354, height=300pt)

| # | Typ | Name | Left | Top | Width | Height | Properties |
|---|---|---|---:|---:|---:|---:|---|
| 3 | Label | `lbl_PageTitle` | 18 | 58 | 524 | 24 | `Caption` wypełnia `RenderPage`, `Font = Segoe UI 12 Bold`, `ForeColor = &H00404040&` (dark gray) |x
| 4 | TextBox | `txt_PageBody` | 18 | 90 | 524 | 260 | `Text` wypełnia `RenderPage`, `Font = Segoe UI 10`, **`MultiLine = True`**, **`Locked = True`**, **`Enabled = False`**, `BackColor = &H8000000F&` (BtnFace - matchuje formę żeby wyglądał jak text label, nie edit box), `SpecialEffect = 0 - fmSpecialEffectFlat` (bez border), `ScrollBars = 0 - fmScrollBarsNone`, `ForeColor = &H00404040&` |x

> **Dlaczego TextBox a nie Label**: TextBox ma natywny word-wrap i nie obcina długich linii. Label z `WordWrap=True` bywa zawodny (obcinanie w połowie zdania przy różnych DPI). `Locked=True + Enabled=False` = user nie może zaznaczać/edytować, ale nadal widzi normalnie renderowany tekst.
>
> **Dlaczego BackColor = BtnFace (nie Window)**: żeby TextBox wyglądał jak Label - transparency z tłem formy. `&H80000005&` (Window białe) wygląda jak input field.

### Sekcja 3: Footer (y=364 do y=414, height=50pt)

| # | Typ | Name | Left | Top | Width | Height | Properties |
|---|---|---|---:|---:|---:|---:|---|
| 5 | CommandButton | `btn_Prev` | 18 | 380 | 100 | 26 | `Caption = "<- Wstecz"`, `Enabled = False` na 1. stronie (kod dynamicznie), `TakeFocusOnClick = True` |x
| 6 | CommandButton | `btn_Next` | 340 | 380 | 100 | 26 | `Caption = "Dalej ->"`, **`Default = True`** (Enter = Next), kod zmienia na `"Zakoncz"` na ostatniej stronie |x
| 7 | CommandButton | `btn_Skip` | 450 | 380 | 100 | 26 | `Caption = "Pomiń"`, **`Cancel = True`** (Esc = Skip) |x

> **ASCII-only captions** (v3.1): VBA UserForm renderer nie wyświetla poprawnie znaków ▶ ◀ ✓ w standardowych fontach. Substytucje: `▶` → `>`, `◀` → `<`, `✓` → usunięte. Patrz sekcja "Emoji + special characters" na dole.

### Suma kontrolek: 7 (+ opcjonalny separator = 8)

---

## Krok 3 - code-behind

Pełny code w [`frm_Tutorial.code-behind.txt`](frm_Tutorial.code-behind.txt). Struktura:

- **State**: `m_currentPage` (1..5)
- **Handlery**: `UserForm_Initialize`, `btn_Next_Click`, `btn_Prev_Click`, `btn_Skip_Click` (4 public event handlers)
- **Private**: `RenderPage` (dispatcher), `RenderPage1_Welcome` .. `RenderPage5_Help` (5 per-page renderers)
- **Content storage**: hardcoded String constants w Render subs, source of truth = `frm_Tutorial.TEXT.md`

**Właściwości specjalne w RenderPage**:
- `lbl_PageIndicator.Caption` update every page
- `btn_Prev.Enabled` = False on page 1, True elsewhere
- `btn_Next.Caption` = "Zakoncz" on last page, "Dalej ->" elsewhere (ASCII-only v3.1)

---

## Krok 4 - eksport do `Source/Forms/`

Po smoke teście: prawy klik → **Export File...** → `frm_Tutorial.frm`.

Update `mod_Diagnostic`:
- `ExpectedForms` - 4 → 5 forms (dodać `frm_Tutorial`)
- `ExpectedFormHandlers("frm_Tutorial")` - nowa lista (4 handlery): `UserForm_Initialize`, `btn_Next_Click`, `btn_Prev_Click`, `btn_Skip_Click`
- `ExpectedFormControls("frm_Tutorial")` - nowa lista (7 kontrolek): `lbl_TutorialHeader`, `lbl_PageIndicator`, `lbl_PageTitle`, `txt_PageBody`, `btn_Prev`, `btn_Next`, `btn_Skip`

---

## Krok 5 - trigger update w `frm_Setup`

Zastąp obecny placeholder MsgBox w [`frm_Setup.code-behind.txt`](frm_Setup.code-behind.txt) `btn_ShowTutorial_Click`:

```vba
Private Sub btn_ShowTutorial_Click()
    frm_Tutorial.Show vbModal
End Sub
```

`vbModal` - blokuje frm_Setup do zamknięcia tutorial. User nie zgubi kontekstu setupu.

---

## Wizualne mockup (ASCII)

```
┌────────────────────────────────────────────────────────────┐
│  Samouczek BNC_Sender                    Strona 3 z 5      │  ← header (y=6..46)
│  ────────────────────────────────────────────────────────  │  ← separator (opcjonalny)
│                                                             │
│  Dodawanie zgłoszeń i wysyłka                               │  ← lbl_PageTitle (y=58)
│                                                             │
│  DODAJESZ ZGŁOSZENIE:                                       │  ← txt_PageBody (y=90..350)
│  Wypełnij 3 pola: Klient FK, Nazwa klienta, Miesiąc obrotu  │     (MultiLine TextBox
│  (YYYY-MM, domyślnie bieżący). Kliknij "Dodaj do listy".    │      Locked=True
│                                                             │      Enabled=False
│  POMYŁKA:                                                   │      BackColor=BtnFace
│  Kliknij na wiersz w liście pending → "Usuń zaznaczone" →   │      brak border)
│  potwierdź. Wysłane zgłoszenia są NIEUSUWALNE.              │
│                                                             │
│  WYSYŁKA (najlepiej raz w miesiącu):                        │
│  Kliknij "Wyślij Wniosek BNC" → potwierdzenie → mail        │
│  wychodzi z Outlooka.                                       │
│                                                             │
│                                                             │
│  [<- Wstecz]                       [Dalej ->]  [Pomiń]      │  ← footer (y=380)
└────────────────────────────────────────────────────────────┘
```

## Kolorystyka - podsumowanie

| Element | Kolor | HEX-VBA |
|---|---|---|
| Form BackColor | System BtnFace (szare) | `&H8000000F&` |
| lbl_TutorialHeader.ForeColor | Ciemnoniebieski | `&H00804000&` |
| lbl_PageIndicator.ForeColor | Medium gray | `&H00666666&` |
| lbl_PageTitle.ForeColor | Dark gray | `&H00404040&` |
| txt_PageBody.BackColor | System BtnFace (transparency effect) | `&H8000000F&` |
| txt_PageBody.ForeColor | Dark gray | `&H00404040&` |
| Separator (opcjonalny) | Light gray | `&H00CCCCCC&` |

## Emoji + special characters w treści (v3.1 - ASCII-only)

**Empirical finding** (2026-08-09): VBA UserForm renderer w Excel 2019/365 **nie wyświetla znaków non-ANSI** takich jak `▶ ◀ ✓ ⚙ ✉` w domyślnych fontach (Tahoma, Segoe UI). Zmiana fontu na Segoe UI Symbol renderuje symbole, ale **rozwala polskie diakrytyki** w tym samym caption.

Rozwiązania rozważane i odrzucone:
- **Osobne Labele per emoji z Segoe UI Symbol font** - +5-10 kontrolek, positioning nightmare, non-BMP (👋 🛟 📈) nadal nie działa
- **Image controls z PNG icons** - "maski" pod textem, wymaga ship PNG assets, nie skaluje z DPI
- **Wingdings/Webdings** - legacy fonts, glyphs mapped na dziwne kody (`3` = strzałka), nieczytelne w kodzie
- **Extended ANSI** - zawodne, coraz mniej działa w Win11

**Decyzja: ASCII-only** (v3.1 code-behind). Substytucje w wyświetlanym tekście:

| Original (TEXT.md) | Substitute (code-behind) | Kontekst |
|---|---|---|
| `▶` (right arrow) | `->` | Button "Dalej ->" (user decyzja 2026-08-09: `->` a nie `>` - lepsza czytelność jako strzałka) |
| `◀` (left arrow) | `<-` | Button "<- Wstecz" (symetria z "Dalej ->") |
| `✓` (check) | (usunięte) | Button "Zakoncz" |
| `→` (body arrow) | `->` | Body text steps |
| `•` (bullet) | `-` | Body bullet lists |
| `·` (middle dot separator) | ` - ` (space-hyphen-space) | Title separators |
| `👋 ⚙ 🛟 ✉ 📈` | (usunięte) | Pure decoration - no info loss |

**Divergence TEXT.md ↔ code-behind**: TEXT.md pozostaje "spec ideal" z emoji dla czytelności w markdown viewerze. Code-behind reflektuje **actual UI** (ASCII). Jeśli w M8+ VBA UserForms zaczną poprawnie renderować emoji (upgrade Excel / new font ecosystem) - refactor będzie 5-min edit.
