"""Route 1: Build native editable DrawingML PPTX from declarative plans or SVG slides."""

from __future__ import annotations

import copy
import json
import re
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from .paths import ensure_repo_scripts_on_path, resolve_template_dir
from .qa import run_qa

_SVG_NS = "http://www.w3.org/2000/svg"
_XLINK_NS = "http://www.w3.org/1999/xlink"
_NUMERIC_CELL_RE = re.compile(
    r"^[+\-−±<>≤≥≈~]?\s*\d[\d,]*(?:\.\d+)?(?:\s*(?:%|‰|±\s*\d+(?:\.\d+)?|×\s*10\^?[\-−]?\d+))?$"
)


def _text_content(el: ET.Element) -> str:
    return "".join(el.itertext()).strip()


def _is_numeric_cell(text: str) -> bool:
    s = text.strip()
    if not s:
        return False
    if _NUMERIC_CELL_RE.match(s):
        return True
    # Accept common academic stats cells like "p < 0.001" or "0.89 ± 0.02"
    if re.match(r"^(?:p\s*[<≤=]\s*0?\.\d+|n\s*=\s*\d[\d,]*|R[²2]\s*=\s*0?\.\d+)$", s, re.IGNORECASE):
        return True
    return False


def _style_three_line_table(
    table_el: ET.Element,
    *,
    primary: str,
    rule_color: str,
    body_box: dict[str, Any] | None = None,
    no_autofit: bool = False,
) -> None:
    """Convert an upstream full-grid native table into a journal-grade three-line table."""
    cells: list[tuple[int, int, ET.Element]] = []
    for child in list(table_el):
        if child.attrib.get("data-pptx-table-cell") == "true":
            try:
                r = int(child.attrib.get("data-pptx-table-row", "0"))
                c = int(child.attrib.get("data-pptx-table-col", "0"))
            except ValueError:
                continue
            if r > 0 and c > 0:
                cells.append((r, c, child))
    if not cells:
        return

    max_row = max(r for r, _, _ in cells)
    max_col = max(c for _, c, _ in cells)

    # Compute target frame (flush with standard academic body box)
    tx = float(body_box.get("x", 52)) if isinstance(body_box, dict) else 52.0
    ty = max(float(body_box.get("y", 196)) + 20.0, 192.0) if isinstance(body_box, dict) else 196.0
    tw = float(body_box.get("w", 1176)) if isinstance(body_box, dict) else 1176.0
    max_h = float(body_box.get("h", 440)) - 60.0 if isinstance(body_box, dict) else 380.0

    # Measure column content length to assign clean proportional widths
    col_weights = [6.0] * max_col
    col_texts: dict[int, list[str]] = {c: [] for c in range(1, max_col + 1)}
    for r, c, cell_el in cells:
        txt = " ".join(_text_content(t) for t in cell_el.iter() if t.tag.rsplit("}", 1)[-1] == "text").strip()
        if r > 1:
            col_texts[c].append(txt)
        eff_len = sum(1.7 if ord(ch) > 127 else 0.9 for ch in txt)
        col_weights[c - 1] = max(col_weights[c - 1], min(eff_len + 4.0, 34.0))

    total_w = sum(col_weights) or 1.0
    raw_widths = [max(140.0, tw * (w / total_w)) for w in col_weights]
    scale_w = tw / sum(raw_widths)
    col_widths = [round(w * scale_w, 2) for w in raw_widths]
    col_widths[-1] = round(tw - sum(col_widths[:-1]), 2)

    header_h = 64.0
    data_rows = max(1, max_row - 1)
    data_h = min(72.0, max(54.0, (max_h - header_h) / data_rows))
    row_heights = [header_h] + [round(data_h, 2)] * data_rows
    th = round(sum(row_heights), 2)

    # Identify numeric columns (right-align header + data)
    numeric_cols = {
        c
        for c, vals in col_texts.items()
        if vals and all(_is_numeric_cell(v) for v in vals)
    }

    table_el.set("data-pptx-table-x", f"{tx:.2f}")
    table_el.set("data-pptx-table-y", f"{ty:.2f}")
    table_el.set("data-pptx-table-w", f"{tw:.2f}")
    table_el.set("data-pptx-table-h", f"{th:.2f}")
    table_el.set("data-pptx-table-col-widths", ",".join(f"{w:.2f}" for w in col_widths))
    table_el.set("data-pptx-table-row-heights", ",".join(f"{h:.2f}" for h in row_heights))

    col_xs = [tx]
    for w in col_widths[:-1]:
        col_xs.append(col_xs[-1] + w)
    row_ys = [ty]
    for h in row_heights[:-1]:
        row_ys.append(row_ys[-1] + h)

    for r, c, cell_el in cells:
        cx = col_xs[c - 1]
        cy = row_ys[r - 1]
        cw = col_widths[c - 1]
        ch = row_heights[r - 1]

        cell_el.set("data-pptx-cell-marL", "114300")
        cell_el.set("data-pptx-cell-marR", "114300")
        cell_el.set("data-pptx-cell-marT", "57150")
        cell_el.set("data-pptx-cell-marB", "57150")
        cell_el.set("data-pptx-table-valign", "middle")
        align = "right" if c in numeric_cols else "left"
        cell_el.set("data-pptx-table-align", align)

        cell_el.set("data-pptx-border-left", "none")
        cell_el.set("data-pptx-border-right", "none")
        cell_el.set("data-pptx-border-top", f"1.5 {primary}" if r == 1 else "none")
        if r == 1:
            cell_el.set("data-pptx-border-bottom", f"1 {rule_color}")
        elif r == max_row:
            cell_el.set("data-pptx-border-bottom", f"1.5 {primary}")
        else:
            cell_el.set("data-pptx-border-bottom", "none")

        for child in cell_el:
            tag = child.tag.rsplit("}", 1)[-1]
            if tag == "rect":
                child.set("x", f"{cx:.2f}")
                child.set("y", f"{cy:.2f}")
                child.set("width", f"{cw:.2f}")
                child.set("height", f"{ch:.2f}")
                child.set("fill", "#FFFFFF")
                child.set("stroke", "none")
                child.set("stroke-width", "0")
            elif tag == "text":
                child.set("data-pptx-box-x", f"{cx + 12:.2f}")
                child.set("data-pptx-box-y", f"{cy + 6:.2f}")
                child.set("data-pptx-box-w", f"{max(20.0, cw - 24):.2f}")
                child.set("data-pptx-box-h", f"{max(20.0, ch - 12):.2f}")
                child.set("data-pptx-valign", "middle")
                child.set("text-anchor", "end" if align == "right" else "start")
                child.set("x", f"{cx + cw - 12:.2f}" if align == "right" else f"{cx + 12:.2f}")
                child.set("y", f"{cy + ch / 2 + 6:.2f}")
                if no_autofit:
                    child.set("data-pptx-no-autofit", "true")
                if r == 1:
                    child.set("fill", primary)


def _inline_vector_svg_images(svg_files: list[Path]) -> None:
    """Inline local ``.svg`` ``<image>`` references as native DrawingML vector groups."""
    ET.register_namespace("", _SVG_NS)
    ET.register_namespace("xlink", _XLINK_NS)

    for svg_path in svg_files:
        raw = svg_path.read_text(encoding="utf-8")
        if ".svg" not in raw.lower():
            continue
        tree = ET.parse(svg_path)
        root = tree.getroot()

        parent_map = {child: parent for parent in root.iter() for child in parent}
        images = [el for el in root.iter() if el.tag.rsplit("}", 1)[-1] == "image"]
        changed = False
        root_defs = next((c for c in root if c.tag.rsplit("}", 1)[-1] == "defs"), None)

        for img in images:
            href = img.attrib.get("href") or img.attrib.get(f"{{{_XLINK_NS}}}href") or ""
            href_clean = href.strip()
            if (
                not href_clean
                or href_clean.startswith(("data:", "http://", "https://"))
                or not href_clean.lower().endswith(".svg")
            ):
                continue
            ref_path = Path(href_clean)
            if not ref_path.is_absolute():
                ref_path = (svg_path.parent / ref_path).resolve()
            if not ref_path.is_file():
                continue

            try:
                sub_root = ET.parse(ref_path).getroot()
            except ET.ParseError:
                continue

            img_x = float(img.attrib.get("x", "0"))
            img_y = float(img.attrib.get("y", "0"))
            img_w = float(img.attrib.get("width", "0"))
            img_h = float(img.attrib.get("height", "0"))
            if img_w <= 0 or img_h <= 0:
                continue

            vb_raw = (sub_root.attrib.get("viewBox") or "").replace(",", " ").split()
            if len(vb_raw) == 4:
                vb_x, vb_y, vb_w, vb_h = (float(v) for v in vb_raw)
            else:
                vb_x, vb_y = 0.0, 0.0
                vb_w = float(re.sub(r"[^0-9.]+", "", sub_root.attrib.get("width", str(img_w))) or img_w)
                vb_h = float(re.sub(r"[^0-9.]+", "", sub_root.attrib.get("height", str(img_h))) or img_h)
            if vb_w <= 0 or vb_h <= 0:
                continue

            scale = min(img_w / vb_w, img_h / vb_h)
            fit_w = vb_w * scale
            fit_h = vb_h * scale
            ox = img_x + (img_w - fit_w) / 2.0 - vb_x * scale
            oy = img_y + (img_h - fit_h) / 2.0 - vb_y * scale

            transform = (
                f"translate({ox:.4f} {oy:.4f})"
                if abs(scale - 1.0) < 1e-4
                else f"translate({ox:.4f} {oy:.4f}) scale({scale:.6f} {scale:.6f})"
            )
            group = ET.Element(f"{{{_SVG_NS}}}g", {"transform": transform})
            if img.attrib.get("id"):
                group.set("id", img.attrib["id"])

            for sub_child in list(sub_root):
                tag = sub_child.tag.rsplit("}", 1)[-1]
                if tag == "defs":
                    if root_defs is None:
                        root_defs = ET.Element(f"{{{_SVG_NS}}}defs")
                        root.insert(0, root_defs)
                    for def_item in list(sub_child):
                        root_defs.append(copy.deepcopy(def_item))
                elif tag not in ("title", "desc", "metadata"):
                    group.append(copy.deepcopy(sub_child))

            parent = parent_map.get(img)
            if parent is not None:
                idx = list(parent).index(img)
                parent.remove(img)
                parent.insert(idx, group)
                changed = True

        if changed:
            tree.write(svg_path, encoding="utf-8", xml_declaration=True)


def _apply_template_scientific_body_style(svg_files: list[Path], template_dir: Path) -> None:
    """Optionally align upstream `scientific-body` colors & dual-font typography with a custom template.

    Activated only when the template's `layouts.json` explicitly declares
    `typography.normalize_scientific_body = true` (leaves built-in templates untouched).
    """
    layouts_path = template_dir / "layouts.json"
    if not layouts_path.is_file():
        return
    try:
        meta = json.loads(layouts_path.read_text(encoding="utf-8-sig"))
    except (OSError, json.JSONDecodeError):
        return
    typo = meta.get("typography")
    if not isinstance(typo, dict) or not typo.get("normalize_scientific_body"):
        return

    colors = meta.get("colors") if isinstance(meta.get("colors"), dict) else {}
    primary = str(colors.get("primary") or "#4472C4")
    accent = str(colors.get("banner") or colors.get("accent") or primary)
    ink = str(colors.get("body") or colors.get("ink") or "#2B3A46")
    border = str(colors.get("rule") or colors.get("border") or "#CBD5E1")
    latin = str(typo.get("latin_font") or "Times New Roman")
    ea = str(typo.get("east_asian_font") or "Microsoft YaHei")
    dual_font = f"'{latin}', '{ea}'"
    body_pt = float(typo.get("body_pt") or 17.0)
    body_px = f"{body_pt * 4.0 / 3.0:.4f}"
    caption_pt = float(typo.get("caption_pt") or 14.0)
    caption_px = f"{caption_pt * 4.0 / 3.0:.4f}"
    table_style = str(typo.get("table_style") or "").strip().lower()
    no_autofit = bool(typo.get("no_autofit"))

    sci_body_box = None
    for layout in meta.get("layouts", []) if isinstance(meta.get("layouts"), list) else []:
        if isinstance(layout, dict) and layout.get("id") == "scientific_composition":
            sci_body_box = layout.get("body_box")
            break

    color_map = {
        "#102F4A": primary,
        "#194A37": primary,
        "#2F759B": accent,
        "#54786B": accent,
        "#17251F": ink,
        "#D3E4DA": border,
        "#D8E1D9": border,
    }
    ET.register_namespace("", _SVG_NS)
    ET.register_namespace("xlink", _XLINK_NS)

    for svg_path in svg_files:
        text = svg_path.read_text(encoding="utf-8")
        if 'id="scientific-body"' not in text:
            continue
        tree = ET.parse(svg_path)
        root = tree.getroot()
        body_group = None
        for el in root.iter():
            if el.attrib.get("id") == "scientific-body":
                body_group = el
                break
        if body_group is None:
            continue
        for el in body_group.iter():
            for attr in ("fill", "stroke"):
                val = el.attrib.get(attr)
                if val and val.upper() in color_map:
                    el.set(attr, color_map[val.upper()])
            local_tag = el.tag.rsplit("}", 1)[-1]
            if local_tag == "tspan" and "font-family" in el.attrib:
                del el.attrib["font-family"]
            elif local_tag == "text":
                el.set("font-family", dual_font)
                if no_autofit and el.attrib.get("data-pptx-no-wrap") != "true":
                    el.set("data-pptx-no-autofit", "true")
                role = el.attrib.get("data-text-role", "")
                if role == "caption":
                    el.set("font-size", caption_px)
                    el.set("font-weight", "400")
                elif role in ("body", "secondary", ""):
                    try:
                        cur = float(el.attrib.get("font-size", "0"))
                    except ValueError:
                        cur = 0.0
                    if 18.0 <= cur <= 26.0:
                        el.set("font-size", body_px)
                    if el.attrib.get("font-weight") in ("600", "700", "bold") and role != "body_heading":
                        el.set("font-weight", "400")
        if table_style == "three_line":
            for tbl in body_group.iter():
                if tbl.attrib.get("data-pptx-table") == "true":
                    _style_three_line_table(
                        tbl,
                        primary=primary,
                        rule_color=border,
                        body_box=sci_body_box,
                        no_autofit=no_autofit,
                    )
        tree.write(svg_path, encoding="utf-8", xml_declaration=True)


def _natural_sort_key(path: Path) -> tuple[int, str]:
    match = re.match(r"^(\d+)", path.stem)
    return (int(match.group(1)) if match else 10**9, path.name.lower())


def export_svgs_to_pptx(
    svg_files: list[Path],
    output_pptx: str | Path,
    *,
    notes: dict[str, str] | None = None,
    canvas_format: str = "ppt169",
    transition: str | None = None,
    verbose: bool = False,
) -> Path:
    """Narrow adapter over upstream DrawingML exporter (`scripts/svg_to_pptx/`).

    Always uses `use_native_shapes=True` so every text box, arrow, table, shape,
    and cropped image becomes natively editable inside PowerPoint/Keynote/LibreOffice.
    """
    if not svg_files:
        raise ValueError("No SVG files provided for PPTX export.")
    out = Path(output_pptx).expanduser().resolve()
    out.parent.mkdir(parents=True, exist_ok=True)

    ensure_repo_scripts_on_path()
    from svg_to_pptx.pptx_builder import create_pptx_with_native_svg

    ok = create_pptx_with_native_svg(
        svg_files=svg_files,
        output_path=out,
        canvas_format=canvas_format,
        verbose=verbose,
        transition=transition,
        use_compat_mode=False,
        notes=notes,
        enable_notes=bool(notes),
        use_native_shapes=True,
    )
    if not ok or not out.is_file():
        raise RuntimeError(f"SVG-to-DrawingML export failed for {out}")
    return out


def _extract_plan_notes(plan: dict[str, Any], svg_files: list[Path]) -> dict[str, str]:
    """Map slide speaker_note fields to rendered SVG stems."""
    slides = [s for s in plan.get("slides", []) if isinstance(s, dict)]
    notes: dict[str, str] = {}
    for slide, svg_path in zip(slides, svg_files):
        note = str(slide.get("speaker_note") or slide.get("notes") or "").strip()
        if note:
            notes[svg_path.stem] = note
    return notes


def build_from_plan(
    plan_path: str | Path,
    output_pptx: str | Path,
    *,
    template: str | None = None,
    work_dir: str | Path | None = None,
    run_qa_gate: bool = True,
    render_backend: str = "auto",
    dpi: int = 150,
) -> dict[str, Any]:
    """Compile a declarative JSON deck plan -> template SVGs -> native editable PPTX."""
    plan_file = Path(plan_path).expanduser().resolve()
    if not plan_file.is_file():
        raise FileNotFoundError(f"Deck plan not found: {plan_file}")
    plan = json.loads(plan_file.read_text(encoding="utf-8-sig"))
    if not isinstance(plan, dict):
        raise ValueError("Deck plan must be a JSON object.")

    template_ref = template or plan.get("template_id") or plan.get("template") or "academic_general"
    template_dir = resolve_template_dir(str(template_ref))

    out_pptx = Path(output_pptx).expanduser().resolve()
    stage_dir = (
        Path(work_dir).expanduser().resolve()
        if work_dir
        else out_pptx.parent / f"{out_pptx.stem}_build"
    )
    svg_dir = stage_dir / "svg"
    svg_dir.mkdir(parents=True, exist_ok=True)

    ensure_repo_scripts_on_path()
    from semantic_template_renderer import render_deck

    render_manifest = render_deck(template_dir, plan_file, svg_dir)
    svg_files = sorted(
        (p for p in svg_dir.glob("*.svg") if p.is_file()),
        key=_natural_sort_key,
    )
    if not svg_files:
        raise RuntimeError(f"Template renderer produced zero SVG slides in {svg_dir}")

    _apply_template_scientific_body_style(svg_files, template_dir)
    _inline_vector_svg_images(svg_files)

    notes = _extract_plan_notes(plan, svg_files)
    export_svgs_to_pptx(svg_files, out_pptx, notes=notes or None)

    qa_report = None
    if run_qa_gate:
        qa_report = run_qa(
            out_pptx,
            report_dir=stage_dir / "qa",
            render=True,
            backend=render_backend,
            dpi=dpi,
        )

    return {
        "schema_version": "easyslides_mac.build_result.v1",
        "mode": "plan",
        "status": qa_report["status"] if qa_report else "pass",
        "template_id": template_dir.name,
        "output_pptx": str(out_pptx),
        "svg_dir": str(svg_dir),
        "slide_count": len(svg_files),
        "render_manifest": render_manifest,
        "qa": qa_report,
    }


def build_from_svg_dir(
    svg_dir: str | Path,
    output_pptx: str | Path,
    *,
    notes_json: str | Path | None = None,
    canvas_format: str = "ppt169",
    run_qa_gate: bool = True,
    render_backend: str = "auto",
    dpi: int = 150,
) -> dict[str, Any]:
    """Compile a directory of 1280x720 SVG slides directly into a native editable PPTX."""
    src_dir = Path(svg_dir).expanduser().resolve()
    if not src_dir.is_dir():
        raise FileNotFoundError(f"SVG directory not found: {src_dir}")

    svg_files = sorted(
        (p for p in src_dir.glob("*.svg") if p.is_file()),
        key=_natural_sort_key,
    )
    if not svg_files:
        raise ValueError(f"No .svg files found in {src_dir}")

    notes: dict[str, str] | None = None
    if notes_json:
        raw_notes = json.loads(Path(notes_json).expanduser().read_text(encoding="utf-8-sig"))
        if isinstance(raw_notes, dict):
            notes = {str(k): str(v) for k, v in raw_notes.items() if str(v).strip()}

    out_pptx = Path(output_pptx).expanduser().resolve()
    export_svgs_to_pptx(
        svg_files,
        out_pptx,
        notes=notes,
        canvas_format=canvas_format,
    )

    qa_report = None
    if run_qa_gate:
        qa_report = run_qa(
            out_pptx,
            report_dir=out_pptx.parent / f"{out_pptx.stem}_qa",
            render=True,
            backend=render_backend,
            dpi=dpi,
        )

    return {
        "schema_version": "easyslides_mac.build_result.v1",
        "mode": "svg_dir",
        "status": qa_report["status"] if qa_report else "pass",
        "svg_dir": str(src_dir),
        "output_pptx": str(out_pptx),
        "slide_count": len(svg_files),
        "qa": qa_report,
    }
