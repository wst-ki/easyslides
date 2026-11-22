"""Route 1: Build native editable DrawingML PPTX from declarative plans or SVG slides."""

from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any

from .paths import ensure_repo_scripts_on_path, resolve_template_dir
from .qa import run_qa


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
