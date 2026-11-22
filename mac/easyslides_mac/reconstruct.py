"""Route 3: Slide image / AI mockup -> native editable PPTX reconstruction.

Clean Layer A / B / C implementation written from scratch:
- Layer A (raster_crops): complex figures/photos cropped cleanly from the source
  slide image (optional rounded/circle alpha mask). Never bake readable slide text
  into Layer A crops.
- Layer B (shapes): native editable vector shapes (panels, cards, dividers,
  badges, lines/arrows) in 1280x720 SVG coordinates.
- Layer C (texts): native editable text boxes (one logical line/paragraph per
  declared box, never fragmented OCR scraps).

Also provides pixel-level source-vs-render MAE diffing and side-by-side
comparison PNG generation.
"""

from __future__ import annotations

import json
import re
import shutil
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

import numpy as np
from PIL import Image, ImageDraw

from .builder import export_svgs_to_pptx
from .fonts import load_truetype_font
from .qa import run_qa

SVG_NS = "http://www.w3.org/2000/svg"
ET.register_namespace("", SVG_NS)

IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".webp"}


def _natural_key(path: Path) -> tuple[int, str]:
    m = re.search(r"(\d+)", path.stem)
    return (int(m.group(1)) if m else 10**9, path.name.lower())


def init_reconstruction(
    source: str | Path,
    work_dir: str | Path,
    *,
    canvas_w: int = 1280,
    canvas_h: int = 720,
) -> dict[str, Any]:
    """Scaffold a clean Layer A/B/C reconstruction workspace from image(s)."""
    src = Path(source).expanduser().resolve()
    if not src.exists():
        raise FileNotFoundError(f"Source image or directory not found: {src}")

    if src.is_dir():
        images = sorted(
            (p for p in src.iterdir() if p.is_file() and p.suffix.lower() in IMAGE_EXTS),
            key=_natural_key,
        )
    else:
        images = [src]
    if not images:
        raise ValueError(f"No supported slide images found in {src}")

    work = Path(work_dir).expanduser().resolve()
    src_copy_dir = work / "sources"
    src_copy_dir.mkdir(parents=True, exist_ok=True)

    slides_spec: list[dict[str, Any]] = []
    for idx, img_path in enumerate(images, start=1):
        dst = src_copy_dir / f"slide_{idx:03d}{img_path.suffix.lower()}"
        shutil.copy2(img_path, dst)
        with Image.open(dst) as im:
            w, h = im.size
        slides_spec.append(
            {
                "page": idx,
                "source_image": str(dst),
                "source_size": [w, h],
                "background_color": "#FFFFFF",
                "raster_crops": [
                    # Example Layer A entry (delete or replace with real regions):
                    # {
                    #   "id": "fig_1",
                    #   "src_bbox": [x1, y1, x2, y2],   # in source pixel coords (or normalized 0..1)
                    #   "dst_box": [x, y, w, h],        # in 1280x720 canvas coords
                    #   "mask": "none"                  # "none" | "rounded" | "circle"
                    # }
                ],
                "shapes": [
                    # Example Layer B entry:
                    # {"kind": "rect", "x": 40, "y": 40, "w": 1200, "h": 72, "fill": "#102F4A", "rx": 6}
                    # {"kind": "line", "x1": 40, "y1": 120, "x2": 1240, "y2": 120, "stroke": "#CBD5E1", "stroke_width": 2}
                ],
                "texts": [
                    # Example Layer C entry:
                    # {
                    #   "x": 60, "y": 54, "w": 1160, "h": 44,
                    #   "text": "Slide Title",
                    #   "font_family": "Microsoft YaHei",
                    #   "font_size": 28,
                    #   "font_weight": "bold",
                    #   "fill": "#FFFFFF",
                    #   "align": "left"
                    # }
                ],
                "speaker_note": "",
            }
        )

    plan = {
        "schema_version": "easyslides_mac.reconstruction_plan.v1",
        "canvas": {"width": canvas_w, "height": canvas_h},
        "slides": slides_spec,
    }
    plan_path = work / "reconstruction_plan.json"
    plan_path.write_text(json.dumps(plan, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return {
        "status": "ok",
        "work_dir": str(work),
        "plan_path": str(plan_path),
        "slide_count": len(slides_spec),
    }


def _crop_layer_a(
    source_img: Path,
    crop_spec: dict[str, Any],
    out_path: Path,
) -> Path:
    """Crop a clean Layer A visual asset from the source slide image."""
    with Image.open(source_img) as raw:
        im = raw.convert("RGBA")
        w, h = im.size
        bbox = crop_spec.get("src_bbox")
        if not isinstance(bbox, list) or len(bbox) != 4:
            raise ValueError(f"raster_crop requires src_bbox=[x1,y1,x2,y2]: {crop_spec}")
        x1, y1, x2, y2 = [float(v) for v in bbox]
        # Support normalized [0..1] coordinates automatically.
        if max(abs(x1), abs(y1), abs(x2), abs(y2)) <= 1.0:
            x1, y1, x2, y2 = x1 * w, y1 * h, x2 * w, y2 * h
        box = (
            max(0, int(round(x1))),
            max(0, int(round(y1))),
            min(w, int(round(x2))),
            min(h, int(round(y2))),
        )
        if box[2] <= box[0] or box[3] <= box[1]:
            raise ValueError(f"Invalid crop bbox {box} for image {source_img} ({w}x{h})")
        cropped = im.crop(box)

        mask_kind = str(crop_spec.get("mask") or "none").lower()
        if mask_kind in {"circle", "ellipse", "rounded"}:
            cw, ch = cropped.size
            mask = Image.new("L", (cw, ch), 0)
            draw = ImageDraw.Draw(mask)
            if mask_kind in {"circle", "ellipse"}:
                draw.ellipse((0, 0, cw - 1, ch - 1), fill=255)
            else:
                radius = int(crop_spec.get("radius", max(4, min(cw, ch) // 16)))
                draw.rounded_rectangle((0, 0, cw - 1, ch - 1), radius=radius, fill=255)
            alpha = cropped.getchannel("A")
            combined = Image.fromarray(np.minimum(np.array(alpha), np.array(mask)))
            cropped.putalpha(combined)

        out_path.parent.mkdir(parents=True, exist_ok=True)
        cropped.save(out_path, format="PNG")
    return out_path


def _el(parent: ET.Element, tag: str, **attrs: Any) -> ET.Element:
    clean = {k.replace("_", "-"): str(v) for k, v in attrs.items() if v is not None}
    node = ET.SubElement(parent, f"{{{SVG_NS}}}{tag}", clean)
    return node


def _assemble_slide_svg(
    slide_spec: dict[str, Any],
    *,
    canvas_w: int,
    canvas_h: int,
    assets_dir: Path,
    svg_path: Path,
) -> Path:
    root = ET.Element(
        f"{{{SVG_NS}}}svg",
        {
            "viewBox": f"0 0 {canvas_w} {canvas_h}",
            "width": str(canvas_w),
            "height": str(canvas_h),
        },
    )
    bg = str(slide_spec.get("background_color") or "#FFFFFF")
    _el(root, "rect", x=0, y=0, width=canvas_w, height=canvas_h, fill=bg)

    # Layer B (background shapes rendered beneath or alongside assets)
    for shape in slide_spec.get("shapes", []):
        if not isinstance(shape, dict):
            continue
        kind = str(shape.get("kind") or "rect").lower()
        if kind == "rect":
            _el(
                root,
                "rect",
                x=shape["x"],
                y=shape["y"],
                width=shape["w"],
                height=shape["h"],
                rx=shape.get("rx"),
                ry=shape.get("ry", shape.get("rx")),
                fill=shape.get("fill", "none"),
                stroke=shape.get("stroke"),
                stroke_width=shape.get("stroke_width"),
                opacity=shape.get("opacity"),
            )
        elif kind == "line":
            _el(
                root,
                "line",
                x1=shape["x1"],
                y1=shape["y1"],
                x2=shape["x2"],
                y2=shape["y2"],
                stroke=shape.get("stroke", "#1E293B"),
                stroke_width=shape.get("stroke_width", 2),
                stroke_dasharray=shape.get("stroke_dasharray"),
            )
        elif kind == "circle":
            _el(
                root,
                "circle",
                cx=shape["cx"],
                cy=shape["cy"],
                r=shape["r"],
                fill=shape.get("fill", "none"),
                stroke=shape.get("stroke"),
                stroke_width=shape.get("stroke_width"),
            )
        elif kind == "polygon":
            _el(
                root,
                "polygon",
                points=shape["points"],
                fill=shape.get("fill", "none"),
                stroke=shape.get("stroke"),
                stroke_width=shape.get("stroke_width"),
            )

    # Layer A (clean raster crops)
    source_img_raw = slide_spec.get("source_image")
    page_num = int(slide_spec.get("page") or 1)
    for idx, crop in enumerate(slide_spec.get("raster_crops", []), start=1):
        if not isinstance(crop, dict):
            continue
        crop_id = str(crop.get("id") or f"crop_{idx:02d}")
        if crop.get("image_path"):
            asset_file = Path(str(crop["image_path"])).expanduser().resolve()
        else:
            if not source_img_raw:
                raise ValueError(f"Slide {page_num} raster_crop {crop_id} needs source_image or image_path")
            asset_file = assets_dir / f"p{page_num:02d}_{crop_id}.png"
            _crop_layer_a(Path(str(source_img_raw)).expanduser().resolve(), crop, asset_file)
        dx, dy, dw, dh = [float(v) for v in crop["dst_box"]]
        rel_href = Path("assets") / asset_file.name
        if asset_file.parent != assets_dir:
            copied = assets_dir / asset_file.name
            shutil.copy2(asset_file, copied)
            rel_href = Path("assets") / copied.name
        _el(
            root,
            "image",
            href=rel_href.as_posix(),
            x=dx,
            y=dy,
            width=dw,
            height=dh,
            preserveAspectRatio=crop.get("preserve_aspect_ratio", "xMidYMid meet"),
        )

    # Layer C (native editable text boxes)
    for txt in slide_spec.get("texts", []):
        if not isinstance(txt, dict):
            continue
        content = str(txt.get("text") or "")
        if not content:
            continue
        x = float(txt["x"])
        y = float(txt["y"])
        w = float(txt.get("w", 400))
        font_size = float(txt.get("font_size", 18))
        align = str(txt.get("align") or "left").lower()
        anchor_map = {"left": "start", "center": "middle", "right": "end"}
        text_anchor = anchor_map.get(align, "start")
        anchor_x = x if align == "left" else (x + w / 2.0 if align == "center" else x + w)
        line_height = float(txt.get("line_height", font_size * 1.32))
        baseline_y = y + font_size * 0.88

        text_el = _el(
            root,
            "text",
            x=round(anchor_x, 2),
            y=round(baseline_y, 2),
            fill=txt.get("fill", "#1E293B"),
            font_family=txt.get("font_family", "Microsoft YaHei"),
            font_size=font_size,
            font_weight=txt.get("font_weight", "normal"),
            text_anchor=text_anchor,
        )
        lines = content.splitlines()
        if len(lines) == 1:
            text_el.text = lines[0]
        else:
            for li, line in enumerate(lines):
                tspan = _el(
                    text_el,
                    "tspan",
                    x=round(anchor_x, 2),
                    dy=0 if li == 0 else round(line_height, 2),
                )
                tspan.text = line

    svg_path.parent.mkdir(parents=True, exist_ok=True)
    ET.ElementTree(root).write(svg_path, encoding="utf-8", xml_declaration=True)
    return svg_path


def build_reconstruction(
    plan_path: str | Path,
    output_pptx: str | Path,
    *,
    run_qa_gate: bool = True,
    render_backend: str = "auto",
    dpi: int = 150,
) -> dict[str, Any]:
    """Build native editable PPTX from a Layer A/B/C reconstruction_plan.json."""
    pf = Path(plan_path).expanduser().resolve()
    plan = json.loads(pf.read_text(encoding="utf-8-sig"))
    canvas = plan.get("canvas") or {}
    canvas_w = int(canvas.get("width") or 1280)
    canvas_h = int(canvas.get("height") or 720)
    slides = [s for s in plan.get("slides", []) if isinstance(s, dict)]
    if not slides:
        raise ValueError("reconstruction_plan.json contains no slides.")

    out_pptx = Path(output_pptx).expanduser().resolve()
    stage_dir = out_pptx.parent / f"{out_pptx.stem}_reconstruct"
    svg_dir = stage_dir / "svg"
    assets_dir = svg_dir / "assets"
    assets_dir.mkdir(parents=True, exist_ok=True)

    svg_files: list[Path] = []
    notes: dict[str, str] = {}
    for idx, slide_spec in enumerate(slides, start=1):
        svg_path = svg_dir / f"{idx:02d}_slide.svg"
        _assemble_slide_svg(
            slide_spec,
            canvas_w=canvas_w,
            canvas_h=canvas_h,
            assets_dir=assets_dir,
            svg_path=svg_path,
        )
        svg_files.append(svg_path)
        note = str(slide_spec.get("speaker_note") or "").strip()
        if note:
            notes[svg_path.stem] = note

    export_svgs_to_pptx(svg_files, out_pptx, notes=notes or None)

    qa_report = None
    diff_report = None
    if run_qa_gate:
        qa_report = run_qa(
            out_pptx,
            report_dir=stage_dir / "qa",
            render=True,
            backend=render_backend,
            dpi=dpi,
        )
        rendered_pngs = (qa_report.get("render") or {}).get("slide_pngs") or []
        source_imgs = [
            Path(str(s["source_image"])) for s in slides if s.get("source_image") and Path(str(s["source_image"])).is_file()
        ]
        if len(source_imgs) == len(rendered_pngs) and source_imgs:
            diff_report = compare_slide_pairs(
                source_imgs,
                [Path(p) for p in rendered_pngs],
                stage_dir / "diff",
            )

    return {
        "schema_version": "easyslides_mac.reconstruction_result.v1",
        "status": qa_report["status"] if qa_report else "pass",
        "output_pptx": str(out_pptx),
        "svg_dir": str(svg_dir),
        "slide_count": len(svg_files),
        "qa": qa_report,
        "diff": diff_report,
    }


def compare_slide_pairs(
    source_images: list[Path],
    rendered_images: list[Path],
    output_dir: str | Path,
) -> dict[str, Any]:
    """Compute per-slide pixel MAE and generate side-by-side Source | Rendered | Diff PNGs."""
    out = Path(output_dir).expanduser().resolve()
    out.mkdir(parents=True, exist_ok=True)
    font = load_truetype_font("Arial", 18, bold=True)

    rows: list[dict[str, Any]] = []
    for idx, (src_p, ren_p) in enumerate(zip(source_images, rendered_images), start=1):
        with Image.open(src_p) as s_im, Image.open(ren_p) as r_im:
            s_rgb = s_im.convert("RGB")
            r_rgb = r_im.convert("RGB")
            if r_rgb.size != s_rgb.size:
                r_rgb = r_rgb.resize(s_rgb.size, Image.Resampling.LANCZOS)

            s_arr = np.asarray(s_rgb, dtype=np.float32)
            r_arr = np.asarray(r_rgb, dtype=np.float32)
            abs_diff = np.abs(s_arr - r_arr)
            mae = float(abs_diff.mean())
            changed_pct = float((abs_diff.mean(axis=2) > 15.0).mean() * 100.0)

            # Build 3-panel visual comparison (Source | Rendered | 3x Amplified Diff)
            panel_w = 560
            panel_h = max(1, int(round(s_rgb.height * (panel_w / max(1, s_rgb.width)))))
            diff_vis = Image.fromarray(np.clip(abs_diff * 3.0, 0, 255).astype(np.uint8))

            s_thumb = s_rgb.resize((panel_w, panel_h), Image.Resampling.LANCZOS)
            r_thumb = r_rgb.resize((panel_w, panel_h), Image.Resampling.LANCZOS)
            d_thumb = diff_vis.resize((panel_w, panel_h), Image.Resampling.LANCZOS)

            pad = 16
            hdr = 34
            canvas = Image.new("RGB", (panel_w * 3 + pad * 4, panel_h + hdr + pad * 2), (245, 247, 250))
            draw = ImageDraw.Draw(canvas)
            labels = [
                f"Slide {idx:02d} Source",
                f"Rendered (MAE={mae:.1f})",
                f"Diff x3 ({changed_pct:.1f}% changed)",
            ]
            for pi, (lbl, thumb) in enumerate(zip(labels, (s_thumb, r_thumb, d_thumb))):
                x0 = pad + pi * (panel_w + pad)
                draw.text((x0 + 4, pad + 4), lbl, fill=(30, 41, 59), font=font)
                canvas.paste(thumb, (x0, pad + hdr))

            comp_path = out / f"compare_{idx:03d}.png"
            canvas.save(comp_path)
            rows.append(
                {
                    "slide": idx,
                    "source": str(src_p),
                    "rendered": str(ren_p),
                    "mae": round(mae, 2),
                    "changed_pct": round(changed_pct, 2),
                    "comparison_png": str(comp_path),
                }
            )

    avg_mae = round(sum(r["mae"] for r in rows) / max(1, len(rows)), 2)
    report = {
        "schema_version": "easyslides_mac.diff_report.v1",
        "slide_count": len(rows),
        "avg_mae": avg_mae,
        "slides": rows,
    }
    (out / "diff_report.json").write_text(
        json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    return report
