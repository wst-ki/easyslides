"""Unified, concise CLI for easyslides-mac (`esmac`)."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

from . import __version__
from .builder import build_from_plan, build_from_svg_dir
from .doctor import run_doctor
from .fonts import font_health_summary, sync_office_fonts_to_user_fonts
from .ingest import ingest_source
from .paths import inspect_template, list_official_templates
from .qa import run_qa
from .reconstruct import build_reconstruction, compare_slide_pairs, init_reconstruction
from .renderer import render_pptx
from .template_ops import apply_fill_plan, extract_pptx_to_svg, inspect_pptx, scaffold_fill_plan


def _emit(payload: Any, *, as_json: bool) -> None:
    if as_json:
        print(json.dumps(payload, ensure_ascii=False, indent=2))
        return
    if isinstance(payload, dict):
        status = payload.get("status")
        if status:
            print(f"[{status.upper()}]")
        print(json.dumps(payload, ensure_ascii=False, indent=2))
    else:
        print(json.dumps(payload, ensure_ascii=False, indent=2))


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="esmac",
        description="macOS-first, background-safe academic PPTX toolkit (zero foreground disruption).",
    )
    parser.add_argument("--version", action="version", version=f"esmac {__version__}")
    sub = parser.add_subparsers(dest="command", required=True)

    # doctor
    p_doc = sub.add_parser("doctor", help="Check Python deps, background renderers, fonts, and templates.")
    p_doc.add_argument("--json", action="store_true", dest="as_json", help="Emit JSON output.")

    # fonts
    p_fonts = sub.add_parser("fonts", help="Inspect or symlink Microsoft Office DFonts for headless LibreOffice.")
    p_fonts.add_argument("action", choices=("status", "link", "unlink"), default="status", nargs="?")
    p_fonts.add_argument("--json", action="store_true", dest="as_json")

    # templates
    p_tpl = sub.add_parser("templates", help="List the 8 official native templates or inspect one template's slots.")
    p_tpl.add_argument("--inspect", metavar="TEMPLATE_ID", help="Inspect exact layout_id and slot_id schema of one template.")
    p_tpl.add_argument("--json", action="store_true", dest="as_json")

    # ingest
    p_ing = sub.add_parser(
        "ingest",
        help="Preprocess a MinerU folder, PDF, Markdown, DOCX, XLSX, or PPTX into ingest_manifest.json.",
    )
    p_ing.add_argument("source", help="Path to source file or existing MinerU output directory.")
    p_ing.add_argument("-o", "--out", required=True, help="Output working directory.")
    p_ing.add_argument("--json", action="store_true", dest="as_json")

    # build (Route 1: Native editable PPTX)
    p_bld = sub.add_parser(
        "build",
        help="Build a native editable PPTX from a declarative deck JSON plan or a directory of 1280x720 SVGs.",
    )
    src_group = p_bld.add_mutually_exclusive_group(required=True)
    src_group.add_argument("--plan", help="Path to declarative deck_plan.json.")
    src_group.add_argument("--svg-dir", help="Path to directory of 1280x720 SVG slide files.")
    p_bld.add_argument("-o", "--out", required=True, help="Output .pptx path.")
    p_bld.add_argument("--template", help="Official template ID (default: academic_general) or template dir.")
    p_bld.add_argument("--notes-json", help="Optional JSON mapping SVG stem -> speaker note (for --svg-dir).")
    p_bld.add_argument("--work-dir", help="Optional staging directory for SVGs and QA artifacts.")
    p_bld.add_argument("--no-qa", action="store_true", help="Skip automatic QA and background rendering.")
    p_bld.add_argument("--renderer", choices=("auto", "soffice", "powerpoint"), default="auto")
    p_bld.add_argument("--dpi", type=int, default=150)
    p_bld.add_argument("--json", action="store_true", dest="as_json")

    # inspect-pptx (Route 2a)
    p_ins = sub.add_parser(
        "inspect-pptx",
        help="Inspect an existing PPTX template's slides/slots and optionally scaffold a fill plan.",
    )
    p_ins.add_argument("pptx", help="Path to source .pptx.")
    p_ins.add_argument("--scaffold", help="Write an editable fill_plan.json scaffold to this path.")
    p_ins.add_argument("--slides", help="Optional slide selection for --scaffold, e.g. '1,2,4-6'.")
    p_ins.add_argument("--json", action="store_true", dest="as_json")

    # fill (Route 2b)
    p_fill = sub.add_parser(
        "fill",
        help="Apply a fill_plan.json to clone/reorder/patch an existing PPTX natively via OOXML.",
    )
    p_fill.add_argument("plan", help="Path to fill_plan.json.")
    p_fill.add_argument("-o", "--out", required=True, help="Output .pptx path.")
    p_fill.add_argument("--no-qa", action="store_true")
    p_fill.add_argument("--renderer", choices=("auto", "soffice", "powerpoint"), default="auto")
    p_fill.add_argument("--dpi", type=int, default=150)
    p_fill.add_argument("--json", action="store_true", dest="as_json")

    # extract-svg (Route 2c)
    p_ext = sub.add_parser(
        "extract-svg",
        help="Convert an existing PPTX into editable layered/flat SVG slides and media assets.",
    )
    p_ext.add_argument("pptx", help="Path to source .pptx.")
    p_ext.add_argument("-o", "--out", required=True, help="Output directory.")
    p_ext.add_argument("--mode", choices=("flat", "layered", "both"), default="both")
    p_ext.add_argument("--json", action="store_true", dest="as_json")

    # reconstruct (Route 3: Image -> Editable PPTX)
    p_rec = sub.add_parser(
        "reconstruct",
        help="Reconstruct slide images / AI mockups into Layer A/B/C editable PPTX.",
    )
    rec_sub = p_rec.add_subparsers(dest="rec_action", required=True)

    p_ri = rec_sub.add_parser("init", help="Create a Layer A/B/C reconstruction_plan.json scaffold.")
    p_ri.add_argument("source", help="Slide image file or folder of slide images.")
    p_ri.add_argument("-o", "--out", required=True, help="Output working directory.")
    p_ri.add_argument("--json", action="store_true", dest="as_json")

    p_rb = rec_sub.add_parser("build", help="Compile reconstruction_plan.json -> editable PPTX + visual diff.")
    p_rb.add_argument("plan", help="Path to reconstruction_plan.json.")
    p_rb.add_argument("-o", "--out", required=True, help="Output .pptx path.")
    p_rb.add_argument("--no-qa", action="store_true")
    p_rb.add_argument("--renderer", choices=("auto", "soffice", "powerpoint"), default="auto")
    p_rb.add_argument("--dpi", type=int, default=150)
    p_rb.add_argument("--json", action="store_true", dest="as_json")

    p_rd = rec_sub.add_parser("diff", help="Compare source slide images against rendered PPTX PNGs.")
    p_rd.add_argument("source_dir", help="Directory containing source slide images.")
    p_rd.add_argument("rendered_dir", help="Directory containing rendered slide_NNN.png images.")
    p_rd.add_argument("-o", "--out", required=True, help="Output diff directory.")
    p_rd.add_argument("--json", action="store_true", dest="as_json")

    # render
    p_ren = sub.add_parser(
        "render",
        help="Render a PPTX to per-slide PNGs and a labeled contact_sheet.png in the background.",
    )
    p_ren.add_argument("pptx", help="Path to .pptx.")
    p_ren.add_argument("-o", "--out", required=True, help="Output PNG directory.")
    p_ren.add_argument("--renderer", choices=("auto", "soffice", "powerpoint"), default="auto")
    p_ren.add_argument("--dpi", type=int, default=150)
    p_ren.add_argument("--no-contact-sheet", action="store_true")
    p_ren.add_argument("--json", action="store_true", dest="as_json")

    # qa
    p_qa = sub.add_parser(
        "qa",
        help="Run OOXML validation, true-font text overflow/overlap QA, and silent background render.",
    )
    p_qa.add_argument("pptx", help="Path to .pptx.")
    p_qa.add_argument("-o", "--out", help="Optional QA report directory.")
    p_qa.add_argument("--no-render", action="store_true")
    p_qa.add_argument("--renderer", choices=("auto", "soffice", "powerpoint"), default="auto")
    p_qa.add_argument("--dpi", type=int, default=150)
    p_qa.add_argument("--json", action="store_true", dest="as_json")

    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    as_json = bool(getattr(args, "as_json", False))

    try:
        if args.command == "doctor":
            rep = run_doctor()
            _emit(rep, as_json=as_json)
            return 0 if rep["status"] != "fail" else 1

        if args.command == "fonts":
            if args.action == "status":
                rep = font_health_summary()
            elif args.action == "link":
                rep = sync_office_fonts_to_user_fonts(unlink=False)
            else:
                rep = sync_office_fonts_to_user_fonts(unlink=True)
            _emit(rep, as_json=as_json)
            return 0

        if args.command == "templates":
            if args.inspect:
                _emit(inspect_template(args.inspect), as_json=as_json)
            else:
                _emit({"templates": list_official_templates()}, as_json=as_json)
            return 0

        if args.command == "ingest":
            rep = ingest_source(args.source, args.out)
            _emit(rep, as_json=as_json)
            return 0

        if args.command == "build":
            if args.plan:
                rep = build_from_plan(
                    args.plan,
                    args.out,
                    template=args.template,
                    work_dir=args.work_dir,
                    run_qa_gate=not args.no_qa,
                    render_backend=args.renderer,
                    dpi=args.dpi,
                )
            else:
                rep = build_from_svg_dir(
                    args.svg_dir,
                    args.out,
                    notes_json=args.notes_json,
                    run_qa_gate=not args.no_qa,
                    render_backend=args.renderer,
                    dpi=args.dpi,
                )
            _emit(rep, as_json=as_json)
            return 0 if rep.get("status") != "fail" else 1

        if args.command == "inspect-pptx":
            rep = inspect_pptx(args.pptx)
            if args.scaffold:
                scaffold_rep = scaffold_fill_plan(
                    args.pptx, args.scaffold, slides=args.slides
                )
                rep = {"library": rep, "scaffold": scaffold_rep}
            _emit(rep, as_json=as_json)
            return 0

        if args.command == "fill":
            rep = apply_fill_plan(
                args.plan,
                args.out,
                run_qa_gate=not args.no_qa,
                render_backend=args.renderer,
                dpi=args.dpi,
            )
            _emit(rep, as_json=as_json)
            return 0 if rep.get("status") != "fail" else 1

        if args.command == "extract-svg":
            rep = extract_pptx_to_svg(args.pptx, args.out, mode=args.mode)
            _emit(rep, as_json=as_json)
            return 0

        if args.command == "reconstruct":
            if args.rec_action == "init":
                rep = init_reconstruction(args.source, args.out)
                _emit(rep, as_json=as_json)
                return 0
            if args.rec_action == "build":
                rep = build_reconstruction(
                    args.plan,
                    args.out,
                    run_qa_gate=not args.no_qa,
                    render_backend=args.renderer,
                    dpi=args.dpi,
                )
                _emit(rep, as_json=as_json)
                return 0 if rep.get("status") != "fail" else 1
            if args.rec_action == "diff":
                src_dir = Path(args.source_dir).expanduser().resolve()
                ren_dir = Path(args.rendered_dir).expanduser().resolve()
                src_imgs = sorted(
                    p for p in src_dir.iterdir() if p.is_file() and p.suffix.lower() in {".png", ".jpg", ".jpeg", ".webp"}
                )
                ren_imgs = sorted(
                    p for p in ren_dir.iterdir() if p.is_file() and p.suffix.lower() == ".png" and p.name != "contact_sheet.png"
                )
                rep = compare_slide_pairs(src_imgs, ren_imgs, args.out)
                _emit(rep, as_json=as_json)
                return 0

        if args.command == "render":
            res = render_pptx(
                args.pptx,
                args.out,
                backend=args.renderer,
                dpi=args.dpi,
                contact_sheet=not args.no_contact_sheet,
            )
            _emit(res.to_dict(), as_json=as_json)
            return 0

        if args.command == "qa":
            rep = run_qa(
                args.pptx,
                report_dir=args.out,
                render=not args.no_render,
                backend=args.renderer,
                dpi=args.dpi,
            )
            _emit(rep, as_json=as_json)
            return 0 if rep.get("status") != "fail" else 1

    except Exception as exc:
        err = {"status": "error", "error": f"{type(exc).__name__}: {exc}"}
        print(json.dumps(err, ensure_ascii=False, indent=2), file=sys.stderr)
        return 2

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
