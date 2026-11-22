"""Background-safe PPTX -> PNG renderer and contact-sheet generator for macOS.

Design contract:
- Default backend is ALWAYS LibreOffice headless (`soffice`) running with an
  isolated per-run `-env:UserInstallation=...` profile. This guarantees:
    1. Zero foreground focus stealing or Dock interruption to the user.
    2. Zero profile lock conflicts even if LibreOffice is open on desktop.
    3. 100% reliability over SSH (`Background` launchd session).
- Native Microsoft PowerPoint AppleScript export is strictly opt-in
  (`backend="powerpoint"`) and stages files inside PowerPoint's sandbox Group
  Container to avoid blocking macOS file-permission dialogs.
"""

from __future__ import annotations

import math
import os
import re
import shutil
import subprocess
import sys
import tempfile
import uuid
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from PIL import Image, ImageDraw

from .fonts import load_truetype_font, populate_libreoffice_user_fonts
from .paths import BRIDGES_DIR

MAC_SOFFICE_CANDIDATES: tuple[Path, ...] = (
    Path("/Applications/LibreOffice.app/Contents/MacOS/soffice"),
    Path("/opt/homebrew/bin/soffice"),
    Path("/usr/local/bin/soffice"),
)

MAC_POWERPOINT_APP = Path("/Applications/Microsoft PowerPoint.app")
OFFICE_GROUP_CONTAINER = (
    Path.home() / "Library" / "Group Containers" / "UBF8T346G9.Office"
)
POWERPOINT_APPLESCRIPT = BRIDGES_DIR / "powerpoint_export.applescript"


@dataclass(frozen=True)
class RenderResult:
    status: str
    backend: str
    pptx_path: Path
    output_dir: Path
    slide_pngs: list[Path]
    dpi: int
    contact_sheet: Path | None = None
    background_safe: bool = True
    notes: list[str] = ()

    def to_dict(self) -> dict[str, Any]:
        return {
            "status": self.status,
            "backend": self.backend,
            "background_safe": self.background_safe,
            "pptx_path": str(self.pptx_path),
            "output_dir": str(self.output_dir),
            "dpi": self.dpi,
            "slide_count": len(self.slide_pngs),
            "slide_pngs": [str(p) for p in self.slide_pngs],
            "contact_sheet": str(self.contact_sheet) if self.contact_sheet else None,
            "notes": list(self.notes),
        }


def find_soffice(explicit: str | Path | None = None) -> str | None:
    """Locate LibreOffice `soffice` executable across macOS / Linux / PATH."""
    if explicit:
        cand = Path(explicit).expanduser()
        if cand.is_file():
            return str(cand)
        return shutil.which(str(explicit))
    for path in MAC_SOFFICE_CANDIDATES:
        if path.is_file():
            return str(path)
    return shutil.which("soffice") or shutil.which("libreoffice")


def find_pdftoppm(explicit: str | Path | None = None) -> str | None:
    if explicit:
        cand = Path(explicit).expanduser()
        if cand.is_file():
            return str(cand)
        return shutil.which(str(explicit))
    for path in (Path("/opt/homebrew/bin/pdftoppm"), Path("/usr/local/bin/pdftoppm")):
        if path.is_file():
            return str(path)
    return shutil.which("pdftoppm")


def has_macos_powerpoint() -> bool:
    return sys.platform == "darwin" and MAC_POWERPOINT_APP.is_dir()


def _slide_number_key(path: Path) -> int:
    match = re.search(r"(\d+)$", path.stem)
    return int(match.group(1)) if match else 10**9


def _normalize_png_names(output_dir: Path) -> list[Path]:
    pngs = sorted(
        (p for p in output_dir.iterdir() if p.is_file() and p.suffix.lower() == ".png" and p.name != "contact_sheet.png"),
        key=_slide_number_key,
    )
    # Two-phase rename prevents collisions (e.g., slide-1.png vs slide_001.png).
    staged: list[tuple[Path, Path]] = []
    for idx, src in enumerate(pngs, start=1):
        tmp = output_dir / f".__stage_{idx:03d}.png"
        src.rename(tmp)
        staged.append((tmp, output_dir / f"slide_{idx:03d}.png"))
    final: list[Path] = []
    for tmp, dst in staged:
        if dst.exists():
            dst.unlink()
        tmp.rename(dst)
        final.append(dst)
    return final


def _clear_old_slides(output_dir: Path) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    for path in output_dir.glob("slide_*.png"):
        path.unlink(missing_ok=True)


def _pdf_to_pngs(pdf_path: Path, output_dir: Path, *, dpi: int) -> list[Path]:
    """Rasterize a PDF to per-slide PNGs using pdftoppm (fast) or PyMuPDF fallback."""
    _clear_old_slides(output_dir)
    pdftoppm = find_pdftoppm()
    if pdftoppm:
        prefix = output_dir / "slide"
        proc = subprocess.run(
            [pdftoppm, "-png", "-r", str(dpi), str(pdf_path), str(prefix)],
            capture_output=True,
            text=True,
            check=False,
        )
        if proc.returncode == 0:
            files = _normalize_png_names(output_dir)
            if files:
                return files

    # PyMuPDF fallback (pure Python, zero external binary needed)
    import pymupdf

    doc = pymupdf.open(str(pdf_path))
    try:
        zoom = dpi / 72.0
        matrix = pymupdf.Matrix(zoom, zoom)
        out: list[Path] = []
        for idx, page in enumerate(doc, start=1):
            pix = page.get_pixmap(matrix=matrix, alpha=False)
            target = output_dir / f"slide_{idx:03d}.png"
            pix.save(str(target))
            out.append(target)
        return out
    finally:
        doc.close()


def _render_via_soffice(
    pptx_path: Path,
    output_dir: Path,
    *,
    dpi: int,
    soffice_path: str | None = None,
    timeout_s: int = 180,
) -> RenderResult:
    """100% background, zero-GUI-interference rendering via isolated LibreOffice."""
    executable = find_soffice(soffice_path)
    if not executable:
        raise FileNotFoundError(
            "LibreOffice (soffice) not found. On macOS install with: brew install --cask libreoffice"
        )

    env = os.environ.copy()
    # IMPORTANT: Only apply Linux svp headless VCL plugin on Linux.
    # On macOS, setting SAL_USE_VCLPLUGIN=svp breaks LibreOffice's Cocoa/CoreText stack.
    if sys.platform.startswith("linux"):
        env["SAL_USE_VCLPLUGIN"] = "svp"
    # Prevent Java/GUI splash or desktop integration overhead.
    env.setdefault("HOME", str(Path.home()))

    with tempfile.TemporaryDirectory(prefix="esmac-lo-") as tmp_str:
        tmp = Path(tmp_str)
        profile_dir = tmp / "lo_profile"
        pdf_out_dir = tmp / "pdf"
        profile_dir.mkdir(parents=True, exist_ok=True)
        pdf_out_dir.mkdir(parents=True, exist_ok=True)
        injected_fonts = 0
        if sys.platform == "darwin":
            injected_fonts = populate_libreoffice_user_fonts(profile_dir / "user" / "fonts")
        # Copy input PPTX to an ASCII-safe temp name to avoid LibreOffice locale quirks.
        staged_pptx = tmp / "deck.pptx"
        shutil.copy2(pptx_path, staged_pptx)

        cmd = [
            executable,
            f"-env:UserInstallation={profile_dir.as_uri()}",
            "--headless",
            "--invisible",
            "--nologo",
            "--nodefault",
            "--nofirststartwizard",
            "--norestore",
            "--convert-to",
            "pdf",
            "--outdir",
            str(pdf_out_dir),
            str(staged_pptx),
        ]
        proc = subprocess.run(
            cmd,
            env=env,
            capture_output=True,
            text=True,
            timeout=timeout_s,
            check=False,
        )
        pdf_path = pdf_out_dir / "deck.pdf"
        if proc.returncode != 0 or not pdf_path.is_file():
            detail = (proc.stderr or proc.stdout or f"exit code {proc.returncode}").strip()
            raise RuntimeError(f"LibreOffice headless conversion failed: {detail}")

        pngs = _pdf_to_pngs(pdf_path, output_dir, dpi=dpi)
        if not pngs:
            raise RuntimeError("PDF rasterization produced zero slide PNGs.")

    return RenderResult(
        status="pass",
        backend="soffice",
        pptx_path=pptx_path,
        output_dir=output_dir,
        slide_pngs=pngs,
        dpi=dpi,
        background_safe=True,
        notes=[
            f"Rendered completely in background using isolated LibreOffice profile "
            f"(injected {injected_fonts} profile font links for SSH CJK fidelity)."
            if injected_fonts
            else "Rendered completely in background using isolated LibreOffice profile."
        ],
    )


def _render_via_macos_powerpoint(
    pptx_path: Path,
    output_dir: Path,
    *,
    dpi: int,
    timeout_s: int = 180,
) -> RenderResult:
    """Opt-in native macOS Microsoft PowerPoint PDF export via sandbox Group Container."""
    if not has_macos_powerpoint():
        raise FileNotFoundError("Microsoft PowerPoint.app is not installed at /Applications/Microsoft PowerPoint.app")
    if not POWERPOINT_APPLESCRIPT.is_file():
        raise FileNotFoundError(f"AppleScript bridge missing: {POWERPOINT_APPLESCRIPT}")

    osascript = shutil.which("osascript") or "/usr/bin/osascript"
    staging_parent = (
        OFFICE_GROUP_CONTAINER / "TemporaryItems"
        if OFFICE_GROUP_CONTAINER.is_dir()
        else Path(tempfile.gettempdir())
    )
    staging_parent.mkdir(parents=True, exist_ok=True)
    run_dir = staging_parent / f"easyslides_mac_{uuid.uuid4().hex[:10]}"
    run_dir.mkdir(parents=True, exist_ok=True)

    try:
        staged_pptx = run_dir / "deck.pptx"
        staged_pdf = run_dir / "deck.pdf"
        shutil.copy2(pptx_path, staged_pptx)

        proc = subprocess.run(
            [osascript, str(POWERPOINT_APPLESCRIPT), str(staged_pptx), str(staged_pdf)],
            capture_output=True,
            text=True,
            timeout=timeout_s,
            check=False,
        )
        if proc.returncode != 0 or not staged_pdf.is_file():
            detail = (proc.stderr or proc.stdout or f"exit code {proc.returncode}").strip()
            hint = ""
            if "-1743" in detail or "not authorized" in detail.lower():
                hint = (
                    " (macOS Automation TCC permission required: grant terminal/sshd permission "
                    "to control Microsoft PowerPoint in System Settings -> Privacy & Security -> Automation, "
                    "or use default background --renderer soffice.)"
                )
            raise RuntimeError(f"macOS PowerPoint export failed: {detail}{hint}")

        pngs = _pdf_to_pngs(staged_pdf, output_dir, dpi=dpi)
        if not pngs:
            raise RuntimeError("PowerPoint PDF export produced zero slide PNGs.")
    finally:
        shutil.rmtree(run_dir, ignore_errors=True)

    return RenderResult(
        status="pass",
        backend="powerpoint",
        pptx_path=pptx_path,
        output_dir=output_dir,
        slide_pngs=pngs,
        dpi=dpi,
        background_safe=False,
        notes=["Rendered via native macOS Microsoft PowerPoint (sandbox container staging)."],
    )


def make_contact_sheet(
    slide_pngs: list[Path],
    output_path: Path,
    *,
    cols: int = 3,
    thumb_width: int = 600,
    padding: int = 20,
    header_height: int = 34,
) -> Path:
    """Create a labeled grid contact sheet so an Agent can visually inspect all slides at once."""
    if not slide_pngs:
        raise ValueError("Cannot build contact sheet from empty slide_pngs list.")
    cols = max(1, min(cols, len(slide_pngs)))
    rows = math.ceil(len(slide_pngs) / cols)

    thumbs: list[tuple[int, Image.Image]] = []
    max_thumb_h = 0
    for idx, png in enumerate(slide_pngs, start=1):
        with Image.open(png) as img:
            rgb = img.convert("RGB")
            ratio = thumb_width / max(1, rgb.width)
            thumb_h = max(1, int(round(rgb.height * ratio)))
            resized = rgb.resize((thumb_width, thumb_h), Image.Resampling.LANCZOS)
            max_thumb_h = max(max_thumb_h, thumb_h)
            thumbs.append((idx, resized))

    cell_w = thumb_width + padding
    cell_h = max_thumb_h + header_height + padding
    canvas_w = cols * cell_w + padding
    canvas_h = rows * cell_h + padding

    sheet = Image.new("RGB", (canvas_w, canvas_h), (242, 244, 247))
    draw = ImageDraw.Draw(sheet)
    label_font = load_truetype_font("Arial", 18, bold=True)

    for pos, (slide_num, thumb) in enumerate(thumbs):
        r = pos // cols
        c = pos % cols
        x0 = padding + c * cell_w
        y0 = padding + r * cell_h
        draw.rectangle(
            [x0 - 1, y0 - 1, x0 + thumb_width + 1, y0 + header_height + max_thumb_h + 1],
            fill=(255, 255, 255),
            outline=(205, 212, 222),
            width=1,
        )
        draw.text(
            (x0 + 10, y0 + 7),
            f"Slide {slide_num:02d}",
            fill=(30, 41, 59),
            font=label_font,
        )
        sheet.paste(thumb, (x0, y0 + header_height))

    output_path.parent.mkdir(parents=True, exist_ok=True)
    sheet.save(output_path)
    return output_path


def render_pptx(
    pptx_path: str | Path,
    output_dir: str | Path,
    *,
    backend: str = "auto",
    dpi: int = 150,
    contact_sheet: bool = True,
    soffice_path: str | Path | None = None,
) -> RenderResult:
    """Render PPTX slides to PNG.

    Backend selection policy (non-intrusive by design):
    - `auto` / `soffice`: ALWAYS prefers silent headless LibreOffice so foreground
      work on Mac is never disturbed.
    - `powerpoint`: explicitly opts into native macOS Microsoft PowerPoint export.
    """
    src = Path(pptx_path).expanduser().resolve()
    if not src.is_file():
        raise FileNotFoundError(f"PPTX not found: {src}")
    out = Path(output_dir).expanduser().resolve()
    out.mkdir(parents=True, exist_ok=True)

    if backend not in {"auto", "soffice", "powerpoint"}:
        raise ValueError(f"Unsupported backend: {backend!r} (choose auto, soffice, or powerpoint)")

    if backend in {"auto", "soffice"}:
        res = _render_via_soffice(src, out, dpi=dpi, soffice_path=str(soffice_path) if soffice_path else None)
    else:
        res = _render_via_macos_powerpoint(src, out, dpi=dpi)

    sheet_path: Path | None = None
    if contact_sheet and res.slide_pngs:
        sheet_path = make_contact_sheet(res.slide_pngs, out / "contact_sheet.png")

    return RenderResult(
        status=res.status,
        backend=res.backend,
        pptx_path=res.pptx_path,
        output_dir=res.output_dir,
        slide_pngs=res.slide_pngs,
        dpi=res.dpi,
        contact_sheet=sheet_path,
        background_safe=res.background_safe,
        notes=res.notes,
    )
