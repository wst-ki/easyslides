"""Environment diagnostics for easyslides-mac (macOS + SSH aware)."""

from __future__ import annotations

import importlib.util
import os
import platform
import shutil
import subprocess
import sys
from pathlib import Path
from typing import Any

from .fonts import font_health_summary
from .paths import MAC_ROOT, REPO_ROOT, list_official_templates
from .renderer import (
    MAC_POWERPOINT_APP,
    OFFICE_GROUP_CONTAINER,
    find_pdftoppm,
    find_soffice,
    has_macos_powerpoint,
)


def _cmd_info(executable: str | None, args: list[str] | None = None) -> dict[str, Any]:
    if not executable:
        return {"available": False}
    try:
        proc = subprocess.run(
            [executable, *(args or ["--version"])],
            capture_output=True,
            text=True,
            timeout=8,
            check=False,
        )
        lines = [ln.strip() for ln in (proc.stdout or proc.stderr).splitlines() if ln.strip()]
        return {
            "available": proc.returncode == 0,
            "path": executable,
            "version": lines[0] if lines else "",
        }
    except (OSError, subprocess.SubprocessError) as exc:
        return {"available": False, "path": executable, "error": str(exc)}


def _mod_available(name: str) -> bool:
    return importlib.util.find_spec(name) is not None


def _session_type() -> str:
    if sys.platform != "darwin":
        return "non-darwin"
    try:
        proc = subprocess.run(
            ["launchctl", "managername"],
            capture_output=True,
            text=True,
            timeout=5,
            check=False,
        )
        return (proc.stdout or "").strip() or "unknown"
    except (OSError, subprocess.SubprocessError):
        return "unknown"


def _powerpoint_version() -> str | None:
    if not has_macos_powerpoint():
        return None
    plist = MAC_POWERPOINT_APP / "Contents" / "Info.plist"
    try:
        proc = subprocess.run(
            ["defaults", "read", str(plist), "CFBundleShortVersionString"],
            capture_output=True,
            text=True,
            timeout=5,
            check=False,
        )
        return proc.stdout.strip() if proc.returncode == 0 else None
    except (OSError, subprocess.SubprocessError):
        return None


def run_doctor() -> dict[str, Any]:
    """Run comprehensive background-readiness diagnostics."""
    required_modules = ("pptx", "PIL", "lxml", "defusedxml", "pymupdf", "numpy")
    optional_modules = ("openpyxl", "mammoth", "bs4", "requests", "svglib", "reportlab")

    mod_status = {m: _mod_available(m) for m in required_modules}
    opt_mod_status = {m: _mod_available(m) for m in optional_modules}

    soffice_bin = find_soffice()
    pdftoppm_bin = find_pdftoppm()
    ffmpeg_bin = shutil.which("ffmpeg") or (
        "/opt/homebrew/bin/ffmpeg" if Path("/opt/homebrew/bin/ffmpeg").is_file() else None
    )

    soffice_info = _cmd_info(soffice_bin, ["--version"])
    pdftoppm_info = _cmd_info(pdftoppm_bin, ["-v"])
    ffmpeg_info = _cmd_info(ffmpeg_bin, ["-version"])

    fonts = font_health_summary()
    templates = list_official_templates()
    session = _session_type()

    blockers: list[str] = []
    warnings: list[str] = []

    for mod, ok in mod_status.items():
        if not ok:
            blockers.append(f"Missing required Python package: {mod}")
    if not soffice_info.get("available"):
        blockers.append(
            "LibreOffice (soffice) unavailable for silent background rendering. "
            "Install on macOS via: brew install --cask libreoffice"
        )
    if not fonts.get("all_key_fonts_resolved"):
        missing_fams = [k for k, v in fonts["resolved"].items() if not v]
        warnings.append(f"Some key academic fonts fell back/missing: {', '.join(missing_fams)}")
    if not pdftoppm_info.get("available"):
        warnings.append("pdftoppm not found; PDF->PNG will use PyMuPDF fallback (still works).")

    return {
        "schema_version": "easyslides_mac.doctor.v1",
        "status": "fail" if blockers else ("warn" if warnings else "pass"),
        "platform": {
            "system": platform.system(),
            "release": platform.release(),
            "machine": platform.machine(),
            "macos_session": session,
            "ssh_background_safe": bool(soffice_info.get("available")),
        },
        "paths": {
            "repo_root": str(REPO_ROOT),
            "mac_root": str(MAC_ROOT),
        },
        "python": {
            "version": sys.version.split()[0],
            "executable": sys.executable,
            "required_modules": mod_status,
            "optional_modules": opt_mod_status,
        },
        "renderers": {
            "default_backend": "soffice",
            "soffice": soffice_info,
            "pdftoppm": pdftoppm_info,
            "ffmpeg": ffmpeg_info,
            "macos_powerpoint": {
                "installed": has_macos_powerpoint(),
                "version": _powerpoint_version(),
                "sandbox_container_ready": OFFICE_GROUP_CONTAINER.is_dir(),
                "policy": "opt-in only (--renderer powerpoint) to avoid foreground interruption",
            },
        },
        "fonts": fonts,
        "templates": {
            "official_count": len(templates),
            "all_present": all(t["exists"] for t in templates),
            "ids": [t["template_id"] for t in templates],
        },
        "env": {
            "mineru_token_configured": bool(os.environ.get("MINERU_API_TOKEN")),
            "image_backend": os.environ.get("IMAGE_BACKEND") or None,
        },
        "blockers": blockers,
        "warnings": warnings,
    }
