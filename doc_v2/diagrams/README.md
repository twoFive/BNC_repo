# Diagramy architektury

Diagramy do prezentacji/dyskusji z biznesem. Stan po milestone M5.

| Plik | Cel |
|---|---|
| `02_system_architecture.jpg` | Komponenty systemu i ich rozmieszczenie (4 warstwy: dystrybucja, komputer handlowca, sieciowa, adresaci) |
| `05_module_architecture.jpg` | Component Diagram (C4 model · poziom 3) — 4 warstwy modułów VBA z zależnościami |
| `*.html` | Źródła diagramów (SVG embedded), regenerowalne |
| `post-release/` | **Wariant docelowy v1.0.0** — frm_Tutorial jako pełny moduł, footery zaktualizowane (do prezentacji "jak będzie wyglądać po skończeniu") |

## Regeneracja JPG ze źródła HTML

Gdy zmieni się kod / architektura, zaktualizuj `*.html` i odpal:

```bash
cd doc_v2/diagrams

# 1. Render HTML → PNG przez headless Chrome
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
# 2. PNG → JPG (PowerShell + System.Drawing, quality 92)
Add-Type -AssemblyName System.Drawing
foreach ($f in @('02_system_architecture','05_module_architecture')) {
    $png = [System.Drawing.Image]::FromFile("$pwd\$f.png")
    $enc = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | ? { $_.MimeType -eq 'image/jpeg' }
    $params = New-Object System.Drawing.Imaging.EncoderParameters(1)
    $params.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter([System.Drawing.Imaging.Encoder]::Quality, 92L)
    $png.Save("$pwd\$f.jpg", $enc, $params)
    $png.Dispose()
}
```

```bash
# 3. Cleanup PNG (zachowujemy tylko HTML jako źródło + JPG jako output)
rm *.png
```

## Wymiary

- **02 system architecture**: 1600 × 1220 px
- **05 module architecture**: 1600 × 1100 px

Drukuj na A3 albo prezentuj na monitorze 1080p+ — tekst pozostanie czytelny.
