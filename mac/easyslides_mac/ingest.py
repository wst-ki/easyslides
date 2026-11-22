"""Fast source preprocessing for easyslides-mac.

Key efficiency feature:
- Directly recognizes existing MinerU output directories (`*_no_images.md`,
  `*_content_list.json`, `layout.json`, `images/`) so papers already parsed in
  `project-hub` / `docs/参考文献/` are indexed in milliseconds with ZERO re-parsing.
- Also handles PDF (PyMuPDF), Markdown/TXT, DOCX (mammoth), XLSX (openpyxl),
  and existing PPTX (python-pptx) into one uniform `ingest_manifest.json`.
"""

from __future__ import annotations

import json
import re
import shutil
from pathlib import Path
from typing import Any

from PIL import Image


def _image_meta(path: Path) -> dict[str, Any]:
    with Image.open(path) as img:
        w, h = img.size
    return {
        "path": str(path.resolve()),
        "name": path.name,
        "width": w,
        "height": h,
        "aspect_ratio": round(w / max(1, h), 3),
    }


def _extract_headings(markdown: str) -> list[dict[str, Any]]:
    headings: list[dict[str, Any]] = []
    for line_no, line in enumerate(markdown.splitlines(), start=1):
        match = re.match(r"^(#{1,4})\s+(.+?)\s*$", line)
        if match:
            headings.append(
                {"level": len(match.group(1)), "title": match.group(2), "line": line_no}
            )
    return headings


def _ingest_mineru_dir(source_dir: Path, work_dir: Path) -> dict[str, Any]:
    """Zero-reparse ingestion of an existing MinerU folder."""
    md_candidates = sorted(source_dir.glob("*_no_images.md")) or sorted(source_dir.glob("*.md"))
    md_path = md_candidates[0] if md_candidates else None
    markdown_text = md_path.read_text(encoding="utf-8", errors="replace") if md_path else ""

    content_list_candidates = sorted(source_dir.glob("*_content_list.json"))
    captions_by_img: dict[str, str] = {}
    if content_list_candidates:
        try:
            items = json.loads(content_list_candidates[0].read_text(encoding="utf-8"))
            if isinstance(items, list):
                for item in items:
                    if not isinstance(item, dict):
                        continue
                    img_rel = item.get("img_path")
                    if img_rel:
                        cap = item.get("image_caption") or item.get("table_caption") or []
                        if isinstance(cap, list):
                            cap = " ".join(str(c).strip() for c in cap if str(c).strip())
                        captions_by_img[Path(str(img_rel)).name] = str(cap).strip()
        except (OSError, json.JSONDecodeError):
            pass

    figures: list[dict[str, Any]] = []
    img_dir = source_dir / "images"
    if img_dir.is_dir():
        for img_path in sorted(img_dir.iterdir()):
            if img_path.suffix.lower() not in {".png", ".jpg", ".jpeg", ".webp"}:
                continue
            try:
                meta = _image_meta(img_path)
            except Exception:
                continue
            if img_path.name in captions_by_img:
                meta["caption"] = captions_by_img[img_path.name]
            figures.append(meta)

    out_md = work_dir / "source.md"
    out_md.write_text(markdown_text, encoding="utf-8")
    return {
        "source_kind": "mineru_dir",
        "source_path": str(source_dir.resolve()),
        "markdown_path": str(out_md),
        "char_count": len(markdown_text),
        "headings": _extract_headings(markdown_text),
        "figure_count": len(figures),
        "figures": figures,
        "reused_existing_parse": True,
    }


def _ingest_pdf(pdf_path: Path, work_dir: Path) -> dict[str, Any]:
    """Extract structured page text and embedded images from a PDF using PyMuPDF."""
    import pymupdf

    assets_dir = work_dir / "extracted_images"
    assets_dir.mkdir(parents=True, exist_ok=True)
    doc = pymupdf.open(str(pdf_path))
    pages_md: list[str] = []
    figures: list[dict[str, Any]] = []
    seen_xrefs: set[int] = set()

    try:
        for page_idx, page in enumerate(doc, start=1):
            text = (page.get_text("text") or "").strip()
            pages_md.append(f"## Page {page_idx}\n\n{text}\n")
            for img_idx, img_info in enumerate(page.get_images(full=True), start=1):
                xref = img_info[0]
                if xref in seen_xrefs:
                    continue
                seen_xrefs.add(xref)
                try:
                    extracted = doc.extract_image(xref)
                    width = int(extracted.get("width", 0))
                    height = int(extracted.get("height", 0))
                    if width < 120 or height < 120:
                        continue  # skip tiny icons/artifacts
                    ext = extracted.get("ext", "png")
                    out_file = assets_dir / f"p{page_idx:02d}_img{img_idx:02d}.{ext}"
                    out_file.write_bytes(extracted["image"])
                    meta = _image_meta(out_file)
                    meta["page"] = page_idx
                    figures.append(meta)
                except Exception:
                    continue
        page_count = len(doc)
    finally:
        doc.close()

    full_md = "\n".join(pages_md)
    out_md = work_dir / "source.md"
    out_md.write_text(full_md, encoding="utf-8")
    return {
        "source_kind": "pdf",
        "source_path": str(pdf_path.resolve()),
        "page_count": page_count,
        "markdown_path": str(out_md),
        "char_count": len(full_md),
        "headings": _extract_headings(full_md),
        "figure_count": len(figures),
        "figures": figures,
        "note": "Extracted via PyMuPDF. For complex vector multi-panel scientific figures, pass a MinerU folder or render pages at high DPI and crop cleanly.",
    }


def _ingest_markdown_or_text(path: Path, work_dir: Path) -> dict[str, Any]:
    text = path.read_text(encoding="utf-8", errors="replace")
    out_md = work_dir / "source.md"
    out_md.write_text(text, encoding="utf-8")
    # Discover local relative images referenced in Markdown
    figures: list[dict[str, Any]] = []
    for match in re.finditer(r"!\[([^\]]*)\]\(([^)]+)\)", text):
        caption, rel = match.group(1).strip(), match.group(2).strip()
        cand = (path.parent / rel).resolve()
        if cand.is_file() and cand.suffix.lower() in {".png", ".jpg", ".jpeg", ".webp"}:
            try:
                meta = _image_meta(cand)
                if caption:
                    meta["caption"] = caption
                figures.append(meta)
            except Exception:
                pass
    return {
        "source_kind": "markdown",
        "source_path": str(path.resolve()),
        "markdown_path": str(out_md),
        "char_count": len(text),
        "headings": _extract_headings(text),
        "figure_count": len(figures),
        "figures": figures,
    }


def _ingest_docx(path: Path, work_dir: Path) -> dict[str, Any]:
    import mammoth

    with path.open("rb") as fh:
        res = mammoth.convert_to_markdown(fh)
    text = res.value
    out_md = work_dir / "source.md"
    out_md.write_text(text, encoding="utf-8")
    return {
        "source_kind": "docx",
        "source_path": str(path.resolve()),
        "markdown_path": str(out_md),
        "char_count": len(text),
        "headings": _extract_headings(text),
        "figure_count": 0,
        "figures": [],
    }


def _ingest_xlsx(path: Path, work_dir: Path) -> dict[str, Any]:
    import openpyxl

    wb = openpyxl.load_workbook(path, data_only=True)
    sections: list[str] = []
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            continue
        sections.append(f"## Sheet: {sheet_name}\n")
        header = [str(c if c is not None else "") for c in rows[0]]
        sections.append("| " + " | ".join(header) + " |")
        sections.append("| " + " | ".join("---" for _ in header) + " |")
        for row in rows[1:201]:
            vals = [str(c if c is not None else "").replace("\n", " ") for c in row]
            sections.append("| " + " | ".join(vals) + " |")
        sections.append("")
    text = "\n".join(sections)
    out_md = work_dir / "source.md"
    out_md.write_text(text, encoding="utf-8")
    return {
        "source_kind": "xlsx",
        "source_path": str(path.resolve()),
        "markdown_path": str(out_md),
        "char_count": len(text),
        "headings": _extract_headings(text),
        "figure_count": 0,
        "figures": [],
    }


def _ingest_pptx(path: Path, work_dir: Path) -> dict[str, Any]:
    from pptx import Presentation

    prs = Presentation(str(path))
    assets_dir = work_dir / "extracted_images"
    assets_dir.mkdir(parents=True, exist_ok=True)
    sections: list[str] = []
    figures: list[dict[str, Any]] = []

    for s_idx, slide in enumerate(prs.slides, start=1):
        sections.append(f"## Slide {s_idx}\n")
        for sh_idx, shape in enumerate(slide.shapes, start=1):
            if getattr(shape, "has_text_frame", False):
                t = (shape.text_frame.text or "").strip()
                if t:
                    sections.append(t)
            if getattr(shape, "shape_type", None) == 13:  # PICTURE
                try:
                    ext = shape.image.ext
                    out_file = assets_dir / f"slide{s_idx:02d}_pic{sh_idx:02d}.{ext}"
                    out_file.write_bytes(shape.image.blob)
                    meta = _image_meta(out_file)
                    meta["slide"] = s_idx
                    figures.append(meta)
                except Exception:
                    pass
        if slide.has_notes_slide and slide.notes_slide.notes_text_frame:
            note = slide.notes_slide.notes_text_frame.text.strip()
            if note:
                sections.append(f"\n> Notes: {note}")
        sections.append("")

    text = "\n".join(sections)
    out_md = work_dir / "source.md"
    out_md.write_text(text, encoding="utf-8")
    return {
        "source_kind": "pptx",
        "source_path": str(path.resolve()),
        "slide_count": len(prs.slides),
        "markdown_path": str(out_md),
        "char_count": len(text),
        "headings": _extract_headings(text),
        "figure_count": len(figures),
        "figures": figures,
    }


def ingest_source(source: str | Path, output_dir: str | Path) -> dict[str, Any]:
    """Ingest a MinerU directory, PDF, Markdown, DOCX, XLSX, or PPTX into a unified manifest."""
    src = Path(source).expanduser().resolve()
    if not src.exists():
        raise FileNotFoundError(f"Source path not found: {src}")
    work = Path(output_dir).expanduser().resolve()
    work.mkdir(parents=True, exist_ok=True)

    if src.is_dir():
        # Check for MinerU outputs or general folder of markdown/images
        if list(src.glob("*_no_images.md")) or (src / "images").is_dir() or list(src.glob("*.md")):
            manifest = _ingest_mineru_dir(src, work)
        else:
            raise ValueError(f"Directory does not contain Markdown or MinerU artifacts: {src}")
    else:
        ext = src.suffix.lower()
        if ext == ".pdf":
            manifest = _ingest_pdf(src, work)
        elif ext in {".md", ".markdown", ".txt"}:
            manifest = _ingest_markdown_or_text(src, work)
        elif ext == ".docx":
            manifest = _ingest_docx(src, work)
        elif ext in {".xlsx", ".xlsm"}:
            manifest = _ingest_xlsx(src, work)
        elif ext == ".pptx":
            manifest = _ingest_pptx(src, work)
        else:
            raise ValueError(f"Unsupported source extension: {ext!r}")

    manifest["schema_version"] = "easyslides_mac.ingest_manifest.v1"
    manifest["work_dir"] = str(work)
    manifest_path = work / "ingest_manifest.json"
    manifest_path.write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    manifest["manifest_path"] = str(manifest_path)
    return manifest
