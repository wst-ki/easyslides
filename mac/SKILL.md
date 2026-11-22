---
name: easyslides-mac
description: >
  Background-safe macOS academic PPTX workflow (`esmac`). Use over SSH or locally
  on macOS to create native editable academic PPTX from papers/MinerU folders,
  fill or edit existing PPTX templates, or reconstruct slide images into editable
  PPTX — without interrupting the user's foreground desktop work.
---

# EasySlides Mac (`esmac`) — Background-Safe PPTX Skill

## Operating Contract (Read First)

1. **Never disturb foreground desktop work**:
   - All normal builds, renders, and QA run **100% silently in the background** via isolated headless LibreOffice (`--renderer auto` / `soffice`).
   - **Never** pass `--renderer powerpoint` unless the user explicitly asks to render with desktop Microsoft PowerPoint.
   - All commands work identically over SSH (`ssh mac`) and in a local terminal.
2. **Run from `mac/` via `uv`**:
   ```bash
   cd ~/project/easyslides/mac
   uv run esmac <command> ...
   ```
   Verify readiness once per machine:
   ```bash
   uv run esmac doctor --json
   ```
   *(Microsoft Office `DFonts` + macOS CJK fonts are automatically symlinked into each isolated headless LibreOffice profile at render time, so SSH `Background` sessions render crisp Microsoft YaHei / KaiTi / Cambria with zero setup and zero changes to `~/Library/Fonts/`.)*
3. **Keep interaction lightweight and fast**:
   - Do **not** interrogate the user with 5 rounds of mandatory popups when the intent is already clear.
   - Ask at most **one** short clarifying message only if genuinely ambiguous (e.g. missing source material, or user didn't pick between native template vs image reconstruction when uploading slide images).
   - Default native template is `academic_general`; default theme palette is `academic_blue` (built-in presets: `academic_blue`, `academic_purple`, `academic_green`, `technology_blue`, or any custom `#RRGGBB`).
   - Prefer "垂听" instead of "聆听" in assistant-authored Chinese closing slides.

---

## Choose One of Three Fast Routes

### Route 1: Paper / Research Material → Native Editable Academic PPTX

1. **Ingest source** (if a MinerU folder already exists with `*_no_images.md` and `images/`, pass that folder directly — **zero re-parsing**!):
   ```bash
   uv run esmac ingest <paper_or_mineru_dir> -o <workdir>/ingest --json
   ```
2. **Author slides** (pick the faster/cleaner option for the task):
   - **Option 1A — Declarative Plan JSON** (fastest for standard academic talks using the 8 official templates):
     First inspect the target template's exact `layout_id` and `slot_id` names:
     ```bash
     uv run esmac templates --inspect academic_general --json
     ```
     Then write `<workdir>/deck_plan.json` and compile:
     ```bash
     uv run esmac build --plan <workdir>/deck_plan.json -o <workdir>/out.pptx --json
     ```
     For `academic_general`, content slides can use `layout_id: "scientific_composition"` with `composition.kind: "auto"` + `reading_task` (`merge_data`, `explain_method`, `read_result`, `compare`, `lookup_values`, `show_coverage`) to get native diagrams, multi-figure comparisons, paired/bar/dot plots, and native tables automatically.
   - **Option 1B — Direct 1280×720 SVG Slides** (full visual freedom; every SVG shape/text/table becomes native editable DrawingML):
     Write `01_cover.svg`, `02_intro.svg`, ... into `<workdir>/svg/` and compile:
     ```bash
     uv run esmac build --svg-dir <workdir>/svg -o <workdir>/out.pptx --json
     ```
     SVG rules: `viewBox="0 0 1280 720"`, white `<rect>` background, `<tspan>` for line wrapping, no `<foreignObject>`, no `rgba()`, no `<mask>`, no `<script>`.

### Route 2: Existing PPTX Template → Fill / Edit / Extract

1. **Inspect & scaffold**:
   ```bash
   uv run esmac inspect-pptx template.pptx --scaffold <workdir>/fill_plan.json --slides 1,2,3,5 --json
   ```
2. **Edit `fill_plan.json`** (update text/table replacements per cloned slide) and **apply via OOXML** (preserves 100% of original formatting, masters, themes, and animations):
   ```bash
   uv run esmac fill <workdir>/fill_plan.json -o <workdir>/filled.pptx --json
   ```
3. **Need to inspect or reuse slides as SVG?**
   ```bash
   uv run esmac extract-svg source.pptx -o <workdir>/extracted --mode both --json
   ```

### Route 3: Slide Image / AI Mockup → Layer A/B/C Editable PPTX

1. **Scaffold from image(s)**:
   ```bash
   uv run esmac reconstruct init <slide_img_or_dir> -o <workdir>/recon --json
   ```
2. **Fill `reconstruction_plan.json`** cleanly separating:
   - `raster_crops` (Layer A): complex illustrations/figures only — never bake readable text into crops; supports pixel or normalized `[0..1]` `src_bbox` and `mask: "none" | "rounded" | "circle"`.
   - `shapes` (Layer B): native vector rectangles, panels, dividers, lines, circles, polygons in `1280x720`.
   - `texts` (Layer C): one logical line/paragraph per native editable text box.
3. **Build + auto-QA + side-by-side MAE diff**:
   ```bash
   uv run esmac reconstruct build <workdir>/recon/reconstruction_plan.json -o <workdir>/rebuilt.pptx --json
   ```

---

## Mandatory Visual Self-Check Before Delivery

Every `build`, `fill`, and `reconstruct build` automatically runs `qa`:
1. Checks OOXML integrity (ZIP validity, no negative `a:ext` extents, valid `notesMaster`).
2. Measures text overflow / off-slide / overlap / minimum font size using **real macOS / Office TrueType & TTC font metrics**.
3. Silently renders all slides to `slides_png/slide_NNN.png` + a labeled `slides_png/contact_sheet.png` in the background.

**Before handing the `.pptx` to the user**:
- Read `qa_report.json` — resolve any `blocking` issue (`TEXT_OVERFLOW`, `TEXT_OFF_SLIDE`, `OOXML_*`).
- Visually inspect `contact_sheet.png` (and full-res `slide_NNN.png` for dense pages or `compare_NNN.png` for reconstruction).
- Fix layout/copy issues by shortening text, adjusting geometry, or splitting slides — never shrink body text below readable projection size (`>= 14–16 pt` for body, `>= 12 pt` for captions).
