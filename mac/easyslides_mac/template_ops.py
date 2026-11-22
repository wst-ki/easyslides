"""Route 2: Inspect, fill (OOXML clone+patch), and SVG-extract existing PPTX files."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from .paths import ensure_repo_scripts_on_path
from .qa import run_qa


def inspect_pptx(pptx_path: str | Path) -> dict[str, Any]:
    """Analyze an existing PPTX as a native slide library (slots, tables, charts)."""
    src = Path(pptx_path).expanduser().resolve()
    if not src.is_file():
        raise FileNotFoundError(f"PPTX not found: {src}")
    ensure_repo_scripts_on_path()
    from template_fill_pptx import analyze_pptx

    return analyze_pptx(src)


def scaffold_fill_plan(
    pptx_path: str | Path,
    plan_output: str | Path,
    *,
    slides: str | None = None,
) -> dict[str, Any]:
    """Generate an editable JSON fill plan for an existing PPTX template."""
    library = inspect_pptx(pptx_path)
    ensure_repo_scripts_on_path()
    from template_fill_pptx import scaffold_plan

    plan = scaffold_plan(library, slides=slides)
    # Mark ready-to-apply once edited so the user does not have to hunt for a hidden status key.
    plan["status"] = "approved"
    out = Path(plan_output).expanduser().resolve()
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(plan, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return {
        "status": "ok",
        "plan_path": str(out),
        "source_pptx": library["source_pptx"],
        "slide_count": len(plan.get("slides", [])),
    }


def apply_fill_plan(
    plan_path: str | Path,
    output_pptx: str | Path,
    *,
    run_qa_gate: bool = True,
    render_backend: str = "auto",
    dpi: int = 150,
) -> dict[str, Any]:
    """Validate and apply a native OOXML PPTX fill plan."""
    pf = Path(plan_path).expanduser().resolve()
    if not pf.is_file():
        raise FileNotFoundError(f"Fill plan not found: {pf}")
    plan = json.loads(pf.read_text(encoding="utf-8-sig"))
    if not isinstance(plan, dict):
        raise ValueError("Fill plan must be a JSON object.")

    ensure_repo_scripts_on_path()
    from template_fill_pptx import analyze_pptx, apply_plan, validate_plan

    source_pptx = Path(str(plan.get("source_pptx", ""))).expanduser().resolve()
    if not source_pptx.is_file():
        raise FileNotFoundError(f"source_pptx not found: {source_pptx}")
    plan["source_pptx"] = str(source_pptx)

    library = analyze_pptx(source_pptx)
    errors, warnings = validate_plan(plan, library)
    if errors:
        raise ValueError("Invalid fill plan: " + "; ".join(errors))

    out = Path(output_pptx).expanduser().resolve()
    out.parent.mkdir(parents=True, exist_ok=True)
    apply_plan(plan, out)

    qa_report = None
    if run_qa_gate:
        qa_report = run_qa(
            out,
            report_dir=out.parent / f"{out.stem}_qa",
            render=True,
            backend=render_backend,
            dpi=dpi,
        )

    return {
        "schema_version": "easyslides_mac.fill_result.v1",
        "status": qa_report["status"] if qa_report else "pass",
        "source_pptx": str(source_pptx),
        "output_pptx": str(out),
        "slide_count": len(plan.get("slides", [])),
        "plan_warnings": warnings,
        "qa": qa_report,
    }


def extract_pptx_to_svg(
    pptx_path: str | Path,
    output_dir: str | Path,
    *,
    mode: str = "both",
) -> dict[str, Any]:
    """Extract an existing PPTX into layered/flat SVG slides and media assets."""
    src = Path(pptx_path).expanduser().resolve()
    if not src.is_file():
        raise FileNotFoundError(f"PPTX not found: {src}")
    out = Path(output_dir).expanduser().resolve()
    out.mkdir(parents=True, exist_ok=True)

    ensure_repo_scripts_on_path()
    from pptx_to_svg.converter import ConvertOptions, convert_pptx_to_svg

    options = ConvertOptions(inheritance_mode=mode)
    result = convert_pptx_to_svg(src, out, options=options)
    return {
        "schema_version": "easyslides_mac.extract_svg_result.v1",
        "status": "pass",
        "source_pptx": str(src),
        "output_dir": str(out),
        "mode": mode,
        "canvas_px": list(result.canvas_px),
        "slide_count": len(result.slides),
        "master_count": len(result.masters),
        "layout_count": len(result.layouts),
        "theme_colors": result.theme_colors,
        "theme_fonts": result.theme_fonts,
    }
