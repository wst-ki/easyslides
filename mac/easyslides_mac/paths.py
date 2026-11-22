"""Repository and asset path resolution for easyslides-mac."""

from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any

PACKAGE_DIR = Path(__file__).resolve().parent
MAC_ROOT = PACKAGE_DIR.parent
REPO_ROOT = MAC_ROOT.parent
TEMPLATES_DIR = REPO_ROOT / "templates"
LAYOUTS_DIR = TEMPLATES_DIR / "layouts"
SCRIPTS_DIR = REPO_ROOT / "scripts"
BRIDGES_DIR = MAC_ROOT / "bridges"


def ensure_repo_scripts_on_path() -> None:
    """Expose upstream zero-coupling core kernels on sys.path only inside adapters."""
    for target in (str(REPO_ROOT), str(SCRIPTS_DIR)):
        if target not in sys.path:
            sys.path.insert(0, target)


def resolve_template_dir(template_id_or_path: str | Path) -> Path:
    """Resolve an official template ID or explicit directory path."""
    raw = Path(template_id_or_path).expanduser()
    if raw.is_dir():
        return raw.resolve()
    candidate = LAYOUTS_DIR / str(template_id_or_path)
    if candidate.is_dir():
        return candidate.resolve()
    raise FileNotFoundError(
        f"Template not found: {template_id_or_path!r} (checked path and {LAYOUTS_DIR})"
    )


def list_official_templates() -> list[dict[str, Any]]:
    """Return metadata for the official native templates in templates/template_policy.json."""
    policy_path = TEMPLATES_DIR / "template_policy.json"
    registry_path = TEMPLATES_DIR / "template_registry.json"
    policy = json.loads(policy_path.read_text(encoding="utf-8-sig"))
    registry = (
        json.loads(registry_path.read_text(encoding="utf-8-sig"))
        if registry_path.is_file()
        else {}
    )
    by_id = {
        str(row.get("template_id")): row
        for row in registry.get("templates", [])
        if isinstance(row, dict)
    }
    default_id = str(policy.get("default_native_template_id") or "academic_general")
    rows: list[dict[str, Any]] = []
    for tid in policy.get("official_template_ids", []):
        meta = by_id.get(tid, {})
        tdir = LAYOUTS_DIR / tid
        palettes_file = tdir / "theme_palettes.json"
        palette_ids: list[str] = []
        if palettes_file.is_file():
            try:
                pdata = json.loads(palettes_file.read_text(encoding="utf-8-sig"))
                if isinstance(pdata.get("palettes"), dict):
                    palette_ids = sorted(pdata["palettes"].keys())
            except (OSError, json.JSONDecodeError):
                pass
        rows.append(
            {
                "template_id": tid,
                "is_default": tid == default_id,
                "display_name": meta.get("display_name") or tid,
                "description": meta.get("description") or "",
                "path": str(tdir),
                "exists": tdir.is_dir(),
                "palettes": palette_ids,
            }
        )
    return rows


def inspect_template(template_id_or_path: str | Path) -> dict[str, Any]:
    """Return a concise Agent-friendly schema of a template's layouts, slots, and palettes."""
    tdir = resolve_template_dir(template_id_or_path)
    layouts_file = tdir / "layouts.json"
    if not layouts_file.is_file():
        raise FileNotFoundError(f"Template lacks layouts.json: {tdir}")
    data = json.loads(layouts_file.read_text(encoding="utf-8-sig"))

    palettes: dict[str, str] = {}
    palettes_file = tdir / "theme_palettes.json"
    if palettes_file.is_file():
        try:
            pdata = json.loads(palettes_file.read_text(encoding="utf-8-sig"))
            for pid, pval in (pdata.get("palettes") or {}).items():
                if isinstance(pval, dict):
                    palettes[pid] = str(pval.get("name") or pid)
        except (OSError, json.JSONDecodeError):
            pass

    layouts_summary: list[dict[str, Any]] = []
    for row in data.get("layouts", []):
        if not isinstance(row, dict):
            continue
        slots = [
            {
                "slot_id": s.get("slot_id"),
                "kind": s.get("kind", "text"),
                "required": bool(s.get("required", False)),
                **({"max_lines": s["max_lines"]} if s.get("max_lines") else {}),
                **({"max_chars_per_line": s["max_chars_per_line"]} if s.get("max_chars_per_line") else {}),
            }
            for s in row.get("slots", [])
            if isinstance(s, dict)
        ]
        layouts_summary.append(
            {
                "layout_id": row.get("layout_id"),
                "role": row.get("role"),
                "svg": row.get("svg"),
                "slots": slots,
            }
        )

    return {
        "template_id": data.get("template_id") or tdir.name,
        "path": str(tdir),
        "canvas": data.get("canvas"),
        "palettes": palettes,
        "layout_count": len(layouts_summary),
        "layouts": layouts_summary,
    }
