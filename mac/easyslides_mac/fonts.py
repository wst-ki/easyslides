"""Cross-platform macOS-first font resolver, Office DFonts linker, and text metrics."""

from __future__ import annotations

import math
import os
import sys
from functools import lru_cache
from pathlib import Path
from typing import Any

from PIL import ImageFont

OFFICE_DFONTS_DIRS: tuple[Path, ...] = (
    Path("/Applications/Microsoft PowerPoint.app/Contents/Resources/DFonts"),
    Path("/Applications/Microsoft Word.app/Contents/Resources/DFonts"),
    Path("/Applications/Microsoft Excel.app/Contents/Resources/DFonts"),
)

USER_OFFICE_FONT_LINK_DIR = Path.home() / "Library" / "Fonts" / "EasySlides-Office"

KEY_OFFICE_FONTS: tuple[str, ...] = (
    "msyh.ttc",
    "msyhbd.ttc",
    "msyhl.ttc",
    "Kaiti.ttf",
    "Simsun.ttc",
    "simsunb.ttf",
    "Cambria.ttc",
    "Cambriab.ttf",
    "Cambriai.ttf",
    "Cambriaz.ttf",
    "Calibri.ttf",
    "Calibrib.ttf",
    "Calibrii.ttf",
    "Calibriz.ttf",
    "Aptos.ttf",
    "Aptos-Bold.ttf",
    "Aptos-Italic.ttf",
    "Aptos-Bold-Italic.ttf",
    "Aptos-SemiBold.ttf",
    "times.ttf",
    "timesbd.ttf",
    "timesi.ttf",
    "timesbi.ttf",
    "Arial.ttf",
    "Arialbd.ttf",
    "Ariali.ttf",
    "Arialbi.ttf",
    "Courier New.ttf",
)

MAC_SYSTEM_EXTRA_FONTS: tuple[Path, ...] = (
    Path("/System/Library/Fonts/Hiragino Sans GB.ttc"),
    Path("/System/Library/Fonts/STHeiti Medium.ttc"),
    Path("/System/Library/Fonts/STHeiti Light.ttc"),
    Path("/System/Library/Fonts/Menlo.ttc"),
    Path("/System/Library/Fonts/Supplemental/Songti.ttc"),
    Path("/System/Library/Fonts/Supplemental/Arial.ttf"),
    Path("/System/Library/Fonts/Supplemental/Arial Bold.ttf"),
    Path("/System/Library/Fonts/Supplemental/Times New Roman.ttf"),
    Path("/System/Library/Fonts/Supplemental/Times New Roman Bold.ttf"),
    Path("/System/Library/Fonts/Supplemental/Georgia.ttf"),
    Path("/System/Library/Fonts/Supplemental/Courier New.ttf"),
)

# Canonical lowercase family -> ordered list of (regular_files, bold_files)
FONT_FILE_MAP: dict[str, tuple[tuple[str, ...], tuple[str, ...]]] = {
    "microsoft yahei": (
        ("msyh.ttc", "msyh.ttf", "Hiragino Sans GB.ttc", "STHeiti Medium.ttc", "NotoSansCJK-Regular.ttc"),
        ("msyhbd.ttc", "msyhbd.ttf", "msyh.ttc", "Hiragino Sans GB.ttc", "NotoSansCJK-Bold.ttc"),
    ),
    "微软雅黑": (
        ("msyh.ttc", "msyh.ttf", "Hiragino Sans GB.ttc", "STHeiti Medium.ttc", "NotoSansCJK-Regular.ttc"),
        ("msyhbd.ttc", "msyhbd.ttf", "msyh.ttc", "Hiragino Sans GB.ttc", "NotoSansCJK-Bold.ttc"),
    ),
    "kaiti": (
        ("Kaiti.ttf", "simkai.ttf", "Songti.ttc", "NotoSerifCJK-Regular.ttc"),
        ("Kaiti.ttf", "simkai.ttf", "Songti.ttc", "NotoSerifCJK-Bold.ttc"),
    ),
    "楷体": (
        ("Kaiti.ttf", "simkai.ttf", "Songti.ttc", "NotoSerifCJK-Regular.ttc"),
        ("Kaiti.ttf", "simkai.ttf", "Songti.ttc", "NotoSerifCJK-Bold.ttc"),
    ),
    "stkaiti": (
        ("Kaiti.ttf", "simkai.ttf", "Songti.ttc"),
        ("Kaiti.ttf", "simkai.ttf", "Songti.ttc"),
    ),
    "simsun": (
        ("Simsun.ttc", "simsun.ttc", "Songti.ttc", "NotoSerifCJK-Regular.ttc"),
        ("Simsun.ttc", "simsun.ttc", "Songti.ttc", "NotoSerifCJK-Bold.ttc"),
    ),
    "宋体": (
        ("Simsun.ttc", "simsun.ttc", "Songti.ttc", "NotoSerifCJK-Regular.ttc"),
        ("Simsun.ttc", "simsun.ttc", "Songti.ttc", "NotoSerifCJK-Bold.ttc"),
    ),
    "simhei": (
        ("STHeiti Medium.ttc", "simhei.ttf", "msyhbd.ttc", "Hiragino Sans GB.ttc"),
        ("STHeiti Medium.ttc", "simhei.ttf", "msyhbd.ttc", "Hiragino Sans GB.ttc"),
    ),
    "pingfang sc": (
        ("PingFang.ttc", "msyh.ttc", "Hiragino Sans GB.ttc", "NotoSansCJK-Regular.ttc"),
        ("PingFang.ttc", "msyhbd.ttc", "Hiragino Sans GB.ttc", "NotoSansCJK-Bold.ttc"),
    ),
    "noto sans cjk sc": (
        ("NotoSansCJK-Regular.ttc", "msyh.ttc", "Hiragino Sans GB.ttc"),
        ("NotoSansCJK-Bold.ttc", "msyhbd.ttc", "Hiragino Sans GB.ttc"),
    ),
    "cambria": (
        ("Cambria.ttc", "cambria.ttc", "Georgia.ttf", "Times New Roman.ttf"),
        ("Cambriab.ttf", "cambriab.ttf", "Cambria.ttc", "Georgia Bold.ttf"),
    ),
    "calibri": (
        ("Calibri.ttf", "calibri.ttf", "Arial.ttf", "DejaVuSans.ttf"),
        ("Calibrib.ttf", "calibrib.ttf", "Arial Bold.ttf", "DejaVuSans-Bold.ttf"),
    ),
    "aptos": (
        ("Aptos.ttf", "Arial.ttf", "DejaVuSans.ttf"),
        ("Aptos-Bold.ttf", "Aptos-SemiBold.ttf", "Arial Bold.ttf", "DejaVuSans-Bold.ttf"),
    ),
    "times new roman": (
        ("times.ttf", "Times New Roman.ttf", "Times.ttc", "NotoSerif-Regular.ttf"),
        ("timesbd.ttf", "Times New Roman Bold.ttf", "Times.ttc", "NotoSerif-Bold.ttf"),
    ),
    "arial": (
        ("Arial.ttf", "arial.ttf", "Helvetica.ttc", "DejaVuSans.ttf"),
        ("Arial Bold.ttf", "arialbd.ttf", "Helvetica.ttc", "DejaVuSans-Bold.ttf"),
    ),
    "helvetica": (
        ("Helvetica.ttc", "Arial.ttf", "DejaVuSans.ttf"),
        ("Helvetica.ttc", "Arial Bold.ttf", "DejaVuSans-Bold.ttf"),
    ),
    "georgia": (
        ("Georgia.ttf", "georgia.ttf", "Times New Roman.ttf"),
        ("Georgia Bold.ttf", "georgiab.ttf", "Times New Roman Bold.ttf"),
    ),
    "segoe ui": (
        ("segoeui.ttf", "Aptos.ttf", "Arial.ttf", "DejaVuSans.ttf"),
        ("segoeuib.ttf", "Aptos-Bold.ttf", "Arial Bold.ttf", "DejaVuSans-Bold.ttf"),
    ),
    "consolas": (
        ("consola.ttf", "Menlo.ttc", "DejaVuSansMono.ttf"),
        ("consolab.ttf", "Menlo.ttc", "DejaVuSansMono-Bold.ttf"),
    ),
}

DEFAULT_FALLBACK_FILES: tuple[str, ...] = (
    "msyh.ttc",
    "Hiragino Sans GB.ttc",
    "STHeiti Medium.ttc",
    "NotoSansCJK-Regular.ttc",
    "Arial.ttf",
    "DejaVuSans.ttf",
)


def font_search_dirs() -> list[Path]:
    """Return existing font directories in priority order for the current OS."""
    candidates: list[Path] = []
    if sys.platform == "darwin":
        candidates.extend(
            [
                USER_OFFICE_FONT_LINK_DIR,
                *OFFICE_DFONTS_DIRS,
                Path.home() / "Library" / "Fonts",
                Path("/Library/Fonts"),
                Path("/System/Library/Fonts/Supplemental"),
                Path("/System/Library/Fonts"),
            ]
        )
    elif os.name == "nt":
        windir = Path(os.environ.get("WINDIR", r"C:\Windows"))
        candidates.append(windir / "Fonts")
    else:
        candidates.extend(
            [
                Path.home() / ".local" / "share" / "fonts",
                Path("/usr/share/fonts/opentype/noto"),
                Path("/usr/share/fonts/truetype/dejavu"),
                Path("/usr/share/fonts/truetype"),
                Path("/usr/share/fonts/opentype"),
                Path("/usr/share/fonts"),
            ]
        )
    seen: set[Path] = set()
    existing: list[Path] = []
    for directory in candidates:
        if directory.is_dir() and directory not in seen:
            seen.add(directory)
            existing.append(directory)
    return existing


def find_office_dfonts_dir() -> Path | None:
    """Return the first installed Microsoft Office DFonts directory on macOS."""
    for directory in OFFICE_DFONTS_DIRS:
        if directory.is_dir():
            return directory
    return None


@lru_cache(maxsize=256)
def resolve_font_path(family: str, *, bold: bool = False) -> Path | None:
    """Resolve a logical font family name to a real local .ttf/.ttc/.otf file."""
    key = family.strip().lower()
    dirs = font_search_dirs()
    filenames: list[str] = []

    if key in FONT_FILE_MAP:
        regular_list, bold_list = FONT_FILE_MAP[key]
        filenames.extend(bold_list if bold else regular_list)

    compact = key.replace(" ", "")
    if compact:
        for ext in (".ttf", ".ttc", ".otf"):
            filenames.append(f"{compact}{ext}")
            filenames.append(f"{family.strip()}{ext}")

    filenames.extend(DEFAULT_FALLBACK_FILES)

    # First pass: exact filename match across all dirs in filename-priority order
    # (so e.g. /System/Library/Fonts/Supplemental/Arial.ttf beats a fallback in DFonts).
    for name in filenames:
        for directory in dirs:
            candidate = directory / name
            if candidate.is_file():
                return candidate

    # Second pass: case-insensitive match inside search dirs.
    lower_targets = [name.lower() for name in filenames]
    dir_indexes: list[dict[str, Path]] = []
    for directory in dirs:
        try:
            index = {p.name.lower(): p for p in directory.iterdir() if p.is_file()}
        except OSError:
            continue
        dir_indexes.append(index)
    for target in lower_targets:
        for index in dir_indexes:
            if target in index:
                return index[target]

    # Third pass: shallow recursive scan on Linux /usr/share/fonts where Noto lives in subdirs.
    if sys.platform.startswith("linux"):
        for directory in dirs:
            for target in lower_targets[:6]:
                for hit in directory.rglob(target):
                    if hit.is_file():
                        return hit
    return None


@lru_cache(maxsize=512)
def load_truetype_font(family: str, size_pt: float, *, bold: bool = False) -> ImageFont.FreeTypeFont | None:
    """Load a Pillow FreeTypeFont at pixel size ~= pt (at 72 dpi, 1 pt = 1 px in measurement space)."""
    path = resolve_font_path(family, bold=bold)
    if path is None:
        return None
    pixel_size = max(1, int(round(size_pt)))
    try:
        return ImageFont.truetype(str(path), pixel_size)
    except OSError:
        return None


def _fallback_char_width_pt(ch: str, size_pt: float) -> float:
    code = ord(ch)
    if code > 0x2E7F:  # CJK / fullwidth punctuation
        return size_pt * 1.0
    if ch.isspace():
        return size_pt * 0.28
    if ch in "ilIjtf.,;:|!'\"":
        return size_pt * 0.32
    if ch in "mwMW@%":
        return size_pt * 0.85
    if ch.isupper():
        return size_pt * 0.66
    return size_pt * 0.54


def measure_text_width_pt(text: str, family: str, size_pt: float, *, bold: bool = False) -> float:
    """Measure single-line text width in points using real font metrics when available."""
    if not text:
        return 0.0
    font = load_truetype_font(family, size_pt, bold=bold)
    if font is not None:
        try:
            return float(font.getlength(text))
        except Exception:
            pass
    return sum(_fallback_char_width_pt(ch, size_pt) for ch in text)


def estimate_wrapped_lines(
    text: str,
    family: str,
    size_pt: float,
    max_width_pt: float,
    *,
    bold: bool = False,
) -> int:
    """Estimate wrapped line count inside a text frame of width `max_width_pt`."""
    if max_width_pt <= 0:
        return 999
    paragraphs = text.splitlines() or [""]
    total_lines = 0
    for para in paragraphs:
        if not para.strip():
            total_lines += 1
            continue
        width = measure_text_width_pt(para, family, size_pt, bold=bold)
        if width <= max_width_pt:
            total_lines += 1
            continue
        # Greedy word/CJK wrap for closer fidelity to Office wrapping.
        current = ""
        para_lines = 1
        token = ""
        tokens: list[str] = []
        for ch in para:
            if ord(ch) > 0x2E7F or ch.isspace():
                if token:
                    tokens.append(token)
                    token = ""
                tokens.append(ch)
            else:
                token += ch
        if token:
            tokens.append(token)
        for part in tokens:
            candidate = current + part
            if current and measure_text_width_pt(candidate, family, size_pt, bold=bold) > max_width_pt:
                para_lines += 1
                current = part.lstrip()
            else:
                current = candidate
        total_lines += max(1, para_lines)
    return total_lines


def sync_office_fonts_to_user_fonts(
    dest_dir: Path | None = None,
    *,
    unlink: bool = False,
) -> dict[str, Any]:
    """Symlink installed Microsoft Office DFonts into ~/Library/Fonts/EasySlides-Office.

    LibreOffice on macOS discovers fonts through CoreText (~/Library/Fonts).
    Read-only symlinks take zero disk space, require no sudo, never modify Office,
    and make silent headless LibreOffice render Microsoft YaHei / KaiTi / Cambria
    with high fidelity.
    """
    target_dir = dest_dir or USER_OFFICE_FONT_LINK_DIR
    if unlink:
        removed: list[str] = []
        if target_dir.is_dir():
            for entry in sorted(target_dir.iterdir()):
                if entry.is_symlink():
                    entry.unlink()
                    removed.append(entry.name)
            try:
                target_dir.rmdir()
            except OSError:
                pass
        resolve_font_path.cache_clear()
        load_truetype_font.cache_clear()
        return {
            "action": "unlink",
            "target_dir": str(target_dir),
            "removed_count": len(removed),
            "removed": removed,
        }

    source_dir = find_office_dfonts_dir()
    if source_dir is None:
        return {
            "action": "link",
            "status": "skipped",
            "reason": "Microsoft Office DFonts directory not found",
            "target_dir": str(target_dir),
            "linked_count": 0,
            "linked": [],
        }

    target_dir.mkdir(parents=True, exist_ok=True)
    linked: list[str] = []
    already_linked: list[str] = []
    missing: list[str] = []

    for filename in KEY_OFFICE_FONTS:
        src = source_dir / filename
        if not src.is_file():
            missing.append(filename)
            continue
        dst = target_dir / filename
        if dst.is_symlink() and dst.resolve() == src.resolve():
            already_linked.append(filename)
            continue
        if dst.exists() or dst.is_symlink():
            dst.unlink()
        dst.symlink_to(src)
        linked.append(filename)

    resolve_font_path.cache_clear()
    load_truetype_font.cache_clear()
    return {
        "action": "link",
        "status": "ok",
        "source_dir": str(source_dir),
        "target_dir": str(target_dir),
        "linked_count": len(linked),
        "already_linked_count": len(already_linked),
        "total_available": len(linked) + len(already_linked),
        "linked": linked,
        "missing": missing,
    }


def populate_libreoffice_user_fonts(user_fonts_dir: Path) -> int:
    """Symlink Office DFonts + key macOS CJK/Latin fonts into an isolated LibreOffice profile.

    Crucial for macOS SSH (`Background` launchd session): CoreText inside a
    background SSH session cannot query Aqua's `fontd`, so LibreOffice headless
    otherwise falls back to Latin-only LinuxLibertineG and turns Chinese glyphs
    into tofu boxes. Placing symlinks inside `<UserInstallation>/user/fonts/`
    fixes CJK + Microsoft font embedding completely in <1 ms with zero changes
    to `~/Library/Fonts/`.
    """
    user_fonts_dir.mkdir(parents=True, exist_ok=True)
    count = 0
    seen_lower: set[str] = set()

    def _link(src: Path) -> None:
        nonlocal count
        if not src.is_file():
            return
        key = src.name.lower()
        if key in seen_lower:
            return
        seen_lower.add(key)
        dst = user_fonts_dir / src.name
        if not dst.exists() and not dst.is_symlink():
            dst.symlink_to(src)
        count += 1

    # 1. All fonts in installed Microsoft Office DFonts (msyh.ttc, Kaiti.ttf, Cambria.ttc, etc.)
    for dfonts_dir in OFFICE_DFONTS_DIRS:
        if dfonts_dir.is_dir():
            for name in KEY_OFFICE_FONTS:
                _link(dfonts_dir / name)
            # Also catch case-variant filenames inside DFonts
            try:
                for entry in dfonts_dir.iterdir():
                    if entry.is_file() and entry.suffix.lower() in {".ttf", ".ttc", ".otf"}:
                        if entry.name.lower() in {n.lower() for n in KEY_OFFICE_FONTS}:
                            _link(entry)
            except OSError:
                pass

    # 2. Core macOS system CJK / Latin fallback fonts
    for sys_font in MAC_SYSTEM_EXTRA_FONTS:
        _link(sys_font)

    return count


def font_health_summary() -> dict[str, Any]:
    """Report resolved font paths for key academic presentation families."""
    key_families = (
        "Microsoft YaHei",
        "KaiTi",
        "SimSun",
        "Cambria",
        "Calibri",
        "Aptos",
        "Times New Roman",
        "Arial",
    )
    resolved = {
        fam: (str(p) if (p := resolve_font_path(fam)) is not None else None)
        for fam in key_families
    }
    dfonts_dir = find_office_dfonts_dir()
    linked_files = (
        sorted(p.name for p in USER_OFFICE_FONT_LINK_DIR.iterdir() if p.is_symlink())
        if USER_OFFICE_FONT_LINK_DIR.is_dir()
        else []
    )
    return {
        "office_dfonts_dir": str(dfonts_dir) if dfonts_dir else None,
        "user_font_link_dir": str(USER_OFFICE_FONT_LINK_DIR),
        "user_font_links_count": len(linked_files),
        "resolved": resolved,
        "all_key_fonts_resolved": all(resolved.values()),
    }
