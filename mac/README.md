# EasySlides Mac (`mac/`)

A clean, **macOS-first, 100% background-safe** toolkit for creating, filling, reconstructing, and validating editable academic PPTX decks — designed to run reliably both locally and over **SSH** without disturbing the user's foreground desktop work.

## Why `mac/` Exists

The upstream repository contains ~106k lines of Python, ~80% of which is Windows-oriented desktop automation (`pywin32`, PowerShell COM, `C:/Windows/Fonts` hardcoding) and heavy multi-round workflow governance, while ~20% is a remarkably clean, low-coupling core (`svg_to_pptx`, `pptx_to_svg`, `semantic_template_renderer`, `academic_composition`, `template_fill_pptx`, plus the `templates/` asset library).

`mac/` takes the proven architectural ideas of EasySlides and implements a fresh, compact Python package (`esmac`) that:
- Keeps all macOS runtime code self-contained under `mac/` (zero invasive edits to upstream root files, clean `upstream/main` merges).
- **Never disturbs foreground desktop work**: renders silently in the background via isolated per-run headless LibreOffice profiles (`-env:UserInstallation=...`), while reusing installed Microsoft PowerPoint `DFonts` (`msyh.ttc`, `Kaiti.ttf`, `Simsun.ttc`, `Cambria.ttc`, `Calibri.ttf`, `Aptos.ttf`, `times.ttf`) for both true-font QA measurement and high-fidelity LibreOffice rendering.
- Supports all three high-efficiency workflows through one concise CLI (`esmac`) and a lightweight `SKILL.md`:
  1. **Native editable academic PPTX** (MinerU folder zero-reparse fast-path / PDF / MD / DOCX / XLSX → declarative JSON or direct 1280×720 SVG → DrawingML PPTX).
  2. **Existing PPTX template fill & SVG extraction** (OOXML slide cloning + text/table replacement preserving 100% template formatting).
  3. **Slide image / AI mockup reconstruction** (Layer A raster crops + Layer B native vectors + Layer C native editable text + pixel MAE comparison sheets).

## Setup on macOS

Prerequisites: Homebrew, `uv`, `poppler` (`pdftoppm`), and `libreoffice` (`brew install --cask libreoffice`).

```bash
cd ~/project/easyslides/mac
uv sync --python 3.13 --extra dev

# Verify everything is green (works over SSH too — Office DFonts are auto-injected
# into each isolated headless LibreOffice profile, zero ~/Library/Fonts pollution):
uv run esmac doctor --json
```

## CLI Quick Reference

```bash
# Diagnostics & discovery
uv run esmac doctor --json
uv run esmac templates --json
uv run esmac fonts status|link|unlink

# Route 1: Research material -> Native editable PPTX
uv run esmac ingest /path/to/mineru_or_paper.pdf -o work/ingest
uv run esmac build --plan work/deck_plan.json -o work/deck.pptx
uv run esmac build --svg-dir work/svg -o work/deck.pptx

# Route 2: Existing PPTX template fill / SVG extraction
uv run esmac inspect-pptx template.pptx --scaffold work/fill_plan.json --slides 1,2,3
uv run esmac fill work/fill_plan.json -o work/filled.pptx
uv run esmac extract-svg deck.pptx -o work/svg_out --mode both

# Route 3: Slide image -> Layer A/B/C editable PPTX
uv run esmac reconstruct init slide.png -o work/recon
uv run esmac reconstruct build work/recon/reconstruction_plan.json -o work/rebuilt.pptx

# Rendering & QA (100% background headless by default)
uv run esmac render deck.pptx -o work/png --dpi 150
uv run esmac qa deck.pptx -o work/qa
# Opt-in native PowerPoint render (only when foreground desktop is idle):
uv run esmac render deck.pptx -o work/png_ppt --renderer powerpoint
```
