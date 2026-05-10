# Diagramy architektury — stan docelowy v1.0.0 (post-release)

> **Uwaga**: te diagramy reprezentują **planowany** stan po Release (M7), **NIE aktualny** stan kodu. Aktualny (po M5) jest w nadrzędnym katalogu `doc_v2/diagrams/`.

## Kiedy używać których diagramów

| Sytuacja | Folder |
|---|---|
| Dyskusja "jak architektura wygląda **teraz**" | `doc_v2/diagrams/` (po M5) |
| Pytanie "jak będzie wyglądać **finalna wersja**" | `doc_v2/diagrams/post-release/` (ten folder) |
| Po faktycznym release v1.0.0 | Skopiuj te pliki do nadrzędnego katalogu, zastąp wersję "po M5" |

## Co różni te diagramy od wersji aktualnej (po M5)

### `05_module_architecture.jpg`

| Element | Stan po M5 | Stan post-release v1.0.0 |
|---|---|---|
| `frm_Tutorial` | Dashed border, grey, "M6/M7 — odłożone" | **Pełny moduł** (biały, niebieska ramka, jak inne UserForms) |
| Strzałka `frm_Setup` → `frm_Tutorial` | Brak | **Jest** (fioletowa, dashed, etykieta "btn_ShowTutorial") |
| Bottom note | "frm_Tutorial dashed = zaplanowany w M6/M7" | **"v1.0.0 post-release: wszystkie 4 UserForms zaimplementowane · M6 polish · UAT 3-5 zaliczone"** |
| Subtitle | "stan po M5" | **"stan docelowy v1.0.0 (post-release)"** |

### `02_system_architecture.jpg`

| Element | Stan po M5 | Stan post-release v1.0.0 |
|---|---|---|
| Subtitle | "stan po M5" | **"stan docelowy v1.0.0 (post-release)"** |
| Footer | "BNC_Sender_v0.1.0 · stan kodu po milestone M5" | **"BNC_Sender_v1.0.0 · Stan docelowy po Release (M7) · git tag v1.0.0"** |

> Reszta diagramu 02 jest identyczna — system layers, ticket do IT i wzorce architektoniczne nie zmieniają się przy release.

## Co diagramy NIE zawierają (świadomie)

- **Konkretne bug fixes z UAT** — niemożliwe do przewidzenia przed M6
- **Akceleratory klawiszowe (Alt+S, Alt+A)** — to detale UI, nie wpływają na architekturę modułów
- **Tab order na formularzach** — j.w.

Jeśli M6/M7 wprowadzi **nowy moduł** lub zmieni granice warstw — trzeba będzie zaktualizować źródła HTML.

## Strategia użycia w prezentacji z przełożonym

1. **Domyślnie pokazuj** wersję z nadrzędnego katalogu (po M5) — to **prawda** o aktualnym stanie kodu.
2. **Jeśli przełożony zapyta** "a jak to będzie wyglądać po skończeniu" — pokaż wersję z `post-release/`.
3. **Po release v1.0.0** — wykonaj `mv post-release/*.jpg ../` (zastąp wersję po M5), zaktualizuj `doc_v2/diagrams/README.md`, usuń `post-release/` (cele osiągnięty).

## Regeneracja

Identyczny pipeline jak w nadrzędnym `doc_v2/diagrams/README.md` — Chrome headless → PowerShell System.Drawing.

```bash
cd doc_v2/diagrams/post-release

"/c/Program Files/Google/Chrome/Application/chrome.exe" \
  --headless --disable-gpu --hide-scrollbars \
  --window-size=1600,1220 \
  --screenshot="$(pwd -W)/02_system_architecture.png" \
  "file:///$(pwd -W)/02_system_architecture.html"

"/c/Program Files/Google/Chrome/Application/chrome.exe" \
  --headless --disable-gpu --hide-scrollbars \
  --window-size=1600,1100 \
  --screenshot="$(pwd -W)/05_module_architecture.png" \
  "file:///$(pwd -W)/05_module_architecture.html"
```

```powershell
Add-Type -AssemblyName System.Drawing
$base = (Get-Location).Path
foreach ($f in @('02_system_architecture','05_module_architecture')) {
    $png = [System.Drawing.Image]::FromFile("$base\$f.png")
    $enc = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | ? { $_.MimeType -eq 'image/jpeg' }
    $params = New-Object System.Drawing.Imaging.EncoderParameters(1)
    $params.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter([System.Drawing.Imaging.Encoder]::Quality, 92L)
    $png.Save("$base\$f.jpg", $enc, $params)
    $png.Dispose()
}
```
