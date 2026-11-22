"""Practical, true-font PPTX quality assurance + background visual rendering."""

from __future__ import annotations

import json
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from pptx import Presentation

from .fonts import estimate_wrapped_lines, resolve_font_path
from .renderer import render_pptx

EMU_PER_INCH = 914400.0
PT_PER_INCH = 72.0

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}


@dataclass(frozen=True)
class BoxInfo:
    slide: int
    shape_idx: int
    name: str
    text: str
    x_in: float
    y_in: float
    w_in: float
    h_in: float
    usable_w_pt: float
    usable_h_pt: float
    font_name: str
    font_size_pt: float
    bold: bool
    line_spacing: float

    @property
    def right(self) -> float:
        return self.x_in + self.w_in

    @property
    def bottom(self) -> float:
        return self.y_in + self.h_in

    @property
    def area(self) -> float:
        return max(0.0, self.w_in) * max(0.0, self.h_in)


def _check_ooxml_package(pptx_path: Path) -> list[dict[str, Any]]:
    issues: list[dict[str, Any]] = []
    try:
        with zipfile.ZipFile(pptx_path, "r") as zf:
            bad = zf.testzip()
            if bad is not None:
                issues.append(
                    {
                        "severity": "blocking",
                        "code": "OOXML_CORRUPT_ZIP",
                        "message": f"Corrupt ZIP member: {bad}",
                    }
                )
                return issues

            names = set(zf.namelist())
            for required in ("[Content_Types].xml", "ppt/presentation.xml"):
                if required not in names:
                    issues.append(
                        {
                            "severity": "blocking",
                            "code": "OOXML_MISSING_PART",
                            "message": f"Missing required OOXML part: {required}",
                        }
                    )

            # Check negative extents that trigger PowerPoint's "Repair" popup.
            slide_parts = sorted(
                n for n in names if n.startswith("ppt/slides/slide") and n.endswith(".xml")
            )
            for idx, part in enumerate(slide_parts, start=1):
                root = ET.fromstring(zf.read(part))
                for ext in root.findall(".//a:ext", NS):
                    cx = int(ext.attrib.get("cx", "0"))
                    cy = int(ext.attrib.get("cy", "0"))
                    if cx < 0 or cy < 0:
                        issues.append(
                            {
                                "severity": "blocking",
                                "code": "OOXML_NEGATIVE_EXTENT",
                                "slide": idx,
                                "message": f"Negative DrawingML extent (cx={cx}, cy={cy}) in {part}",
                            }
                        )

            # If notes slides exist, ensure notesMaster1.xml exists as well.
            has_notes_slides = any(
                n.startswith("ppt/notesSlides/notesSlide") and n.endswith(".xml") for n in names
            )
            if has_notes_slides and "ppt/notesMasters/notesMaster1.xml" not in names:
                issues.append(
                    {
                        "severity": "blocking",
                        "code": "OOXML_MISSING_NOTES_MASTER",
                        "message": "notesSlide parts exist without ppt/notesMasters/notesMaster1.xml",
                    }
                )
    except zipfile.BadZipFile as exc:
        issues.append(
            {
                "severity": "blocking",
                "code": "OOXML_BAD_ZIP",
                "message": f"Not a valid PPTX/ZIP archive: {exc}",
            }
        )
    return issues


def _extract_shape_font(shape: Any) -> tuple[str, float, bool, float]:
    font_name = "Microsoft YaHei"
    size_pt = 16.0
    bold = False
    line_spacing = 1.2
    tf = shape.text_frame
    for para in tf.paragraphs:
        if para.line_spacing:
            try:
                val = float(para.line_spacing)
                if 0.8 <= val <= 3.0:
                    line_spacing = val
            except (TypeError, ValueError):
                pass
        if para.font and para.font.name:
            font_name = para.font.name
        if para.font and para.font.size:
            size_pt = para.font.size.pt
        if para.font and para.font.bold is not None:
            bold = bool(para.font.bold)
        for run in para.runs:
            if not run.text.strip():
                continue
            if run.font.name:
                font_name = run.font.name
            if run.font.size:
                size_pt = run.font.size.pt
            if run.font.bold is not None:
                bold = bool(run.font.bold)
            return font_name, size_pt, bold, line_spacing
    return font_name, size_pt, bold, line_spacing


def _collect_text_boxes(prs: Presentation) -> tuple[list[BoxInfo], float, float]:
    slide_w_in = prs.slide_width / EMU_PER_INCH
    slide_h_in = prs.slide_height / EMU_PER_INCH
    boxes: list[BoxInfo] = []

    for s_idx, slide in enumerate(prs.slides, start=1):
        for sh_idx, shape in enumerate(slide.shapes, start=1):
            if not getattr(shape, "has_text_frame", False):
                continue
            text = (shape.text_frame.text or "").strip()
            if not text:
                continue
            x_in = float(shape.left) / EMU_PER_INCH
            y_in = float(shape.top) / EMU_PER_INCH
            w_in = float(shape.width) / EMU_PER_INCH
            h_in = float(shape.height) / EMU_PER_INCH
            tf = shape.text_frame
            m_left = float(tf.margin_left or 0) / EMU_PER_INCH
            m_right = float(tf.margin_right or 0) / EMU_PER_INCH
            m_top = float(tf.margin_top or 0) / EMU_PER_INCH
            m_bottom = float(tf.margin_bottom or 0) / EMU_PER_INCH
            usable_w_pt = max(1.0, (w_in - m_left - m_right) * PT_PER_INCH)
            usable_h_pt = max(1.0, (h_in - m_top - m_bottom) * PT_PER_INCH)
            font_name, font_size_pt, bold, line_spacing = _extract_shape_font(shape)
            boxes.append(
                BoxInfo(
                    slide=s_idx,
                    shape_idx=sh_idx,
                    name=str(getattr(shape, "name", f"shape_{sh_idx}")),
                    text=text,
                    x_in=x_in,
                    y_in=y_in,
                    w_in=w_in,
                    h_in=h_in,
                    usable_w_pt=usable_w_pt,
                    usable_h_pt=usable_h_pt,
                    font_name=font_name,
                    font_size_pt=font_size_pt,
                    bold=bold,
                    line_spacing=line_spacing,
                )
            )
    return boxes, slide_w_in, slide_h_in


def _intersection_area(a: BoxInfo, b: BoxInfo) -> float:
    x_overlap = max(0.0, min(a.right, b.right) - max(a.x_in, b.x_in))
    y_overlap = max(0.0, min(a.bottom, b.bottom) - max(a.y_in, b.y_in))
    return x_overlap * y_overlap


def _check_text_layout(pptx_path: Path) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    prs = Presentation(str(pptx_path))
    boxes, slide_w_in, slide_h_in = _collect_text_boxes(prs)
    issues: list[dict[str, Any]] = []
    fonts_used: dict[str, str | None] = {}

    by_slide: dict[int, list[BoxInfo]] = {}
    for box in boxes:
        by_slide.setdefault(box.slide, []).append(box)
        if box.font_name not in fonts_used:
            resolved = resolve_font_path(box.font_name, bold=box.bold)
            fonts_used[box.font_name] = str(resolved) if resolved else None

        # 1. Off-slide check
        tol = 0.08
        if (
            box.x_in < -tol
            or box.y_in < -tol
            or box.right > slide_w_in + tol
            or box.bottom > slide_h_in + tol
        ):
            issues.append(
                {
                    "severity": "blocking",
                    "code": "TEXT_OFF_SLIDE",
                    "slide": box.slide,
                    "shape": box.name,
                    "text_preview": box.text[:60],
                    "message": (
                        f"Text box extends outside slide bounds ({slide_w_in:.2f}x{slide_h_in:.2f} in): "
                        f"x={box.x_in:.2f}, y={box.y_in:.2f}, r={box.right:.2f}, b={box.bottom:.2f}"
                    ),
                }
            )

        # 2. Font size readability floor
        if box.font_size_pt < 7.5:
            issues.append(
                {
                    "severity": "blocking",
                    "code": "TEXT_FONT_TOO_SMALL",
                    "slide": box.slide,
                    "shape": box.name,
                    "font_size_pt": round(box.font_size_pt, 1),
                    "text_preview": box.text[:60],
                    "message": f"Font size {box.font_size_pt:.1f} pt is below readable floor (7.5 pt).",
                }
            )
        elif box.font_size_pt < 9.0:
            issues.append(
                {
                    "severity": "warning",
                    "code": "TEXT_FONT_SMALL",
                    "slide": box.slide,
                    "shape": box.name,
                    "font_size_pt": round(box.font_size_pt, 1),
                    "text_preview": box.text[:60],
                    "message": f"Font size {box.font_size_pt:.1f} pt is small for projection (< 9 pt).",
                }
            )

        # 3. Real-font vertical overflow estimation
        lines = estimate_wrapped_lines(
            box.text,
            box.font_name,
            box.font_size_pt,
            box.usable_w_pt,
            bold=box.bold,
        )
        needed_h_pt = lines * box.font_size_pt * max(1.08, box.line_spacing)
        ratio = needed_h_pt / max(1.0, box.usable_h_pt)
        if ratio > 1.20 and (needed_h_pt - box.usable_h_pt) > 6.0:
            issues.append(
                {
                    "severity": "blocking",
                    "code": "TEXT_OVERFLOW",
                    "slide": box.slide,
                    "shape": box.name,
                    "text_preview": box.text[:60],
                    "estimated_lines": lines,
                    "overflow_ratio": round(ratio, 2),
                    "message": (
                        f"Estimated text height {needed_h_pt:.1f} pt exceeds box usable height "
                        f"{box.usable_h_pt:.1f} pt ({ratio:.2f}x, ~{lines} lines)."
                    ),
                }
            )
        elif ratio > 1.05 and (needed_h_pt - box.usable_h_pt) > 4.0:
            issues.append(
                {
                    "severity": "warning",
                    "code": "TEXT_TIGHT_FIT",
                    "slide": box.slide,
                    "shape": box.name,
                    "text_preview": box.text[:60],
                    "estimated_lines": lines,
                    "overflow_ratio": round(ratio, 2),
                    "message": f"Text fit is tight ({ratio:.2f}x usable height); inspect rendered slide.",
                }
            )

    # 4. Pairwise overlap check per slide
    for slide_num, slide_boxes in by_slide.items():
        for i in range(len(slide_boxes)):
            a = slide_boxes[i]
            for j in range(i + 1, len(slide_boxes)):
                b = slide_boxes[j]
                inter = _intersection_area(a, b)
                min_area = min(a.area, b.area)
                if min_area <= 0 or inter < 0.08:
                    continue
                overlap_ratio = inter / min_area
                if overlap_ratio >= 0.60:
                    issues.append(
                        {
                            "severity": "warning",
                            "code": "TEXT_OVERLAP",
                            "slide": slide_num,
                            "shapes": [a.name, b.name],
                            "overlap_ratio": round(overlap_ratio, 2),
                            "message": (
                                f"Text boxes '{a.name}' and '{b.name}' overlap "
                                f"({overlap_ratio:.0%} of smaller box)."
                            ),
                        }
                    )

    stats = {
        "slide_count": len(prs.slides),
        "slide_size_in": [round(slide_w_in, 2), round(slide_h_in, 2)],
        "text_box_count": len(boxes),
        "fonts_resolved": fonts_used,
    }
    return issues, stats


def run_qa(
    pptx_path: str | Path,
    *,
    report_dir: str | Path | None = None,
    render: bool = True,
    backend: str = "auto",
    dpi: int = 150,
) -> dict[str, Any]:
    """Run OOXML + true-font text QA and optional silent background rendering."""
    src = Path(pptx_path).expanduser().resolve()
    if not src.is_file():
        raise FileNotFoundError(f"PPTX not found: {src}")

    out_dir = Path(report_dir).expanduser().resolve() if report_dir else src.parent / f"{src.stem}_qa"
    out_dir.mkdir(parents=True, exist_ok=True)

    issues = _check_ooxml_package(src)
    stats: dict[str, Any] = {}
    if not any(i["code"].startswith("OOXML_BAD") or i["code"] == "OOXML_CORRUPT_ZIP" for i in issues):
        layout_issues, stats = _check_text_layout(src)
        issues.extend(layout_issues)

    render_info: dict[str, Any] | None = None
    if render:
        try:
            rendered = render_pptx(
                src,
                out_dir / "slides_png",
                backend=backend,
                dpi=dpi,
                contact_sheet=True,
            )
            render_info = rendered.to_dict()
        except Exception as exc:
            issues.append(
                {
                    "severity": "blocking",
                    "code": "RENDER_FAILED",
                    "message": str(exc),
                }
            )
            render_info = {"status": "fail", "error": str(exc)}

    blocking = [i for i in issues if i.get("severity") == "blocking"]
    warnings = [i for i in issues if i.get("severity") == "warning"]

    report = {
        "schema_version": "easyslides_mac.qa_report.v1",
        "status": "fail" if blocking else ("warn" if warnings else "pass"),
        "pptx_path": str(src),
        "report_dir": str(out_dir),
        "stats": stats,
        "blocking_count": len(blocking),
        "warning_count": len(warnings),
        "issues": issues,
        "render": render_info,
        "visual_review_note": (
            "Inspect contact_sheet.png and individual slide_NNN.png with a vision-capable model "
            "or human reviewer before final delivery."
        ),
    }
    (out_dir / "qa_report.json").write_text(
        json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    return report
