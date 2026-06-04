# Diagramy architektury

Diagramy do prezentacji/dyskusji z biznesem. Stan po milestone M5.

| Plik | Cel | Audience |
|---|---|---|
| `01_system_context.jpg` | **C4 Level 1 — System Context**: aktorzy (Handlowiec, Kierownik, Zespół BNC, Dział IT) + systemy zewnętrzne (Outlook, Exchange, OneDrive) wokół BNC_Sender. Statyczna mapa "kto styka się z systemem". Audience: biznes + tech |
| `02_system_architecture.jpg` | **C4 Level 2 — Container**: komponenty systemu i ich rozmieszczenie (4 warstwy: dystrybucja, komputer handlowca, sieciowa, adresaci) | tech / IT |
| `05_module_architecture.jpg` | Component Diagram (C4 model · poziom 3) — 4 warstwy modułów VBA z zależnościami | tech / dev |
| `06_business_process.jpg` | **Workflow biznesowy v1** — single-page composite: panele Faza A intro + workflow + panele Faza B na jednym obrazie (1400×1560). Dla drukowania A2/A3. | **biznes / decyzyjna** |
| `06_business_process_v2.jpg` + `.md` | **Workflow biznesowy v2** — split: JPG to czysty workflow (1400×1080), tekstowy opis Fazy A + wartości + Fazy B w markdown. Dla slajdów + executive summary. | **biznes / decyzyjna** |
| `*.html` | Źródła diagramów (SVG embedded), regenerowalne | — |
| `post-release/` | **Wariant docelowy v1.0.0** — frm_Tutorial jako pełny moduł, footery zaktualizowane (do prezentacji "jak będzie wyglądać po skończeniu") | tech |

## Regeneracja JPG ze źródła HTML

Gdy zmieni się kod / architektura, zaktualizuj `*.html` i odpal:

```bash
cd doc_3/diagrams

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
- **06 business process**: 1400 × 1080 px

Drukuj na A3 albo prezentuj na monitorze 1080p+ — tekst pozostanie czytelny.
