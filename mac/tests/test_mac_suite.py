"""Automated test suite for easyslides-mac (runs on macOS, Linux, and over SSH)."""

from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest
from PIL import Image, ImageDraw

from easyslides_mac.builder import build_from_plan, build_from_svg_dir
from easyslides_mac.doctor import run_doctor
from easyslides_mac.fonts import (
    estimate_wrapped_lines,
    font_health_summary,
    measure_text_width_pt,
    resolve_font_path,
)
from easyslides_mac.ingest import ingest_source
from easyslides_mac.paths import inspect_template, list_official_templates
from easyslides_mac.qa import run_qa
from easyslides_mac.reconstruct import build_reconstruction, init_reconstruction
from easyslides_mac.renderer import find_soffice, render_pptx
from easyslides_mac.template_ops import (
    apply_fill_plan,
    extract_pptx_to_svg,
    inspect_pptx,
    scaffold_fill_plan,
)


def _sample_svg(title: str, subtitle: str) -> str:
    return f"""<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1280 720" width="1280" height="720">
  <rect x="0" y="0" width="1280" height="720" fill="#FFFFFF"/>
  <rect x="0" y="0" width="1280" height="110" fill="#102F4A"/>
  <text x="64" y="68" fill="#FFFFFF" font-family="Microsoft YaHei" font-size="30" font-weight="bold">{title}</text>
  <text x="64" y="180" fill="#1E293B" font-family="Microsoft YaHei" font-size="20">{subtitle}</text>
  <rect x="64" y="230" width="520" height="360" rx="12" fill="#F1F5F9" stroke="#CBD5E1" stroke-width="2"/>
  <text x="90" y="280" fill="#0F172A" font-family="Microsoft YaHei" font-size="18" font-weight="bold">核心方法与定量对比</text>
  <text x="90" y="324" fill="#334155" font-family="Microsoft YaHei" font-size="16">1. 后台无头渲染，零前台干扰</text>
  <text x="90" y="362" fill="#334155" font-family="Microsoft YaHei" font-size="16">2. 原生 DrawingML 可编辑图形</text>
  <text x="90" y="400" fill="#334155" font-family="Microsoft YaHei" font-size="16">3. 真实字体度量与拼图自检</text>
</svg>
"""


def test_official_templates_discovered():
    templates = list_official_templates()
    ids = [t["template_id"] for t in templates]
    assert len(templates) == 8
    assert "academic_general" in ids
    assert any(t["is_default"] and t["template_id"] == "academic_general" for t in templates)
    assert all(t["exists"] for t in templates)

    info = inspect_template("academic_general")
    assert info["template_id"] == "academic_general"
    layout_ids = {row["layout_id"] for row in info["layouts"]}
    assert {"cover", "toc", "ending", "scientific_composition"}.issubset(layout_ids)


def test_fonts_resolution_and_metrics():
    summary = font_health_summary()
    assert isinstance(summary["resolved"], dict)
    # On macOS with PowerPoint installed, Microsoft YaHei & Cambria must resolve to real files.
    if sys.platform == "darwin":
        assert resolve_font_path("Microsoft YaHei") is not None
        assert resolve_font_path("Cambria") is not None
        assert resolve_font_path("Arial") is not None

    w_en = measure_text_width_pt("Hello World", "Arial", 18)
    w_zh = measure_text_width_pt("流域水沙耦合演变特征分析", "Microsoft YaHei", 18)
    assert w_en > 40
    assert w_zh > 150
    assert estimate_wrapped_lines("短句", "Microsoft YaHei", 18, 300) == 1
    assert estimate_wrapped_lines("这是一段明显超过单行宽度限制的中文学术汇报正文内容测试", "Microsoft YaHei", 18, 120) >= 3


def test_doctor_schema():
    rep = run_doctor()
    assert rep["schema_version"] == "easyslides_mac.doctor.v1"
    assert rep["templates"]["official_count"] == 8
    assert rep["renderers"]["default_backend"] == "soffice"
    if sys.platform == "darwin" and find_soffice():
        assert rep["platform"]["ssh_background_safe"] is True
        assert rep["status"] in {"pass", "warn"}


def test_ingest_mineru_folder_zero_reparse(tmp_path: Path):
    mineru_dir = tmp_path / "paper_mineru"
    img_dir = mineru_dir / "images"
    img_dir.mkdir(parents=True)
    (mineru_dir / "paper_no_images.md").write_text(
        "# Test Paper Title\n\n## 1 Introduction\n\nSome evidence.\n\n## 2 Methods\n\nDetails.\n",
        encoding="utf-8",
    )
    fig_path = img_dir / "fig1.png"
    Image.new("RGB", (800, 400), (20, 80, 140)).save(fig_path)
    (mineru_dir / "paper_content_list.json").write_text(
        json.dumps([{"img_path": "images/fig1.png", "image_caption": ["Figure 1: Basin dynamics."]}]),
        encoding="utf-8",
    )

    out_dir = tmp_path / "ingest_out"
    manifest = ingest_source(mineru_dir, out_dir)
    assert manifest["source_kind"] == "mineru_dir"
    assert manifest["reused_existing_parse"] is True
    assert manifest["figure_count"] == 1
    assert manifest["figures"][0]["aspect_ratio"] == 2.0
    assert "Basin dynamics" in manifest["figures"][0]["caption"]
    assert Path(manifest["manifest_path"]).is_file()


def test_route1_svg_build_qa_and_background_render(tmp_path: Path):
    svg_dir = tmp_path / "svg"
    svg_dir.mkdir()
    (svg_dir / "01_cover.svg").write_text(
        _sample_svg("EasySlides Mac 学术汇报测试", "基于后台无头渲染与真实字体度量"),
        encoding="utf-8",
    )
    (svg_dir / "02_results.svg").write_text(
        _sample_svg("研究结果与跨模型对比", "所有文字与形状均为 PowerPoint 原生可编辑对象"),
        encoding="utf-8",
    )
    notes_file = tmp_path / "notes.json"
    notes_file.write_text(
        json.dumps({"01_cover": "这是封面页备注。", "02_results": "这是结果页备注。"}, ensure_ascii=False),
        encoding="utf-8",
    )

    out_pptx = tmp_path / "deck.pptx"
    has_lo = find_soffice() is not None
    res = build_from_svg_dir(
        svg_dir,
        out_pptx,
        notes_json=notes_file,
        run_qa_gate=has_lo,
    )
    assert out_pptx.is_file()
    assert res["slide_count"] == 2

    if has_lo:
        assert res["qa"] is not None
        assert res["qa"]["blocking_count"] == 0
        render_info = res["qa"]["render"]
        assert render_info["backend"] == "soffice"
        assert render_info["background_safe"] is True
        assert render_info["slide_count"] == 2
        assert Path(render_info["contact_sheet"]).is_file()
    else:
        qa_rep = run_qa(out_pptx, report_dir=tmp_path / "qa_norender", render=False)
        assert qa_rep["blocking_count"] == 0


def test_route1_plan_academic_general_build(tmp_path: Path):
    # Create a small real scientific figure image to exercise scientific_composition.
    fig = tmp_path / "result_fig.png"
    im = Image.new("RGB", (900, 560), (245, 248, 252))
    draw = ImageDraw.Draw(im)
    draw.rectangle([40, 40, 860, 520], outline=(16, 47, 74), width=4)
    draw.rectangle([120, 320, 260, 480], fill=(16, 47, 74))
    draw.rectangle([340, 220, 480, 480], fill=(37, 99, 235))
    draw.rectangle([560, 140, 700, 480], fill=(14, 165, 233))
    im.save(fig)

    plan = {
        "template_id": "academic_general",
        "theme_palette": "academic_blue",
        "discipline": "geography",
        "presenter": "测试汇报人",
        "presenter_institution": "",
        "date": "2026-11-22",
        "slides": [
            {
                "page": "P01",
                "role": "cover",
                "layout_id": "cover_blue_band",
                "slot_payload": {
                    "RESEARCH_FIELD": "地理学 · 流域人地系统",
                    "TITLE": "流域人地系统耦合研究进展",
                    "SUBTITLE": "macOS 后台原生可编辑测试",
                    "REPORTER": "测试汇报人",
                    "AFFILIATION": "示例单位",
                    "DATE": "2026-11-22",
                },
                "speaker_note": "欢迎各位老师同学。",
            },
            {
                "page": "P02",
                "role": "content",
                "layout_id": "scientific_composition",
                "slot_payload": {
                    "CHAPTER_LABEL": "主要结果",
                    "PAGE_TITLE": "多尺度水沙通量在退耕后呈现显著阶段性收敛",
                    "KEY_MESSAGE": "基于 1980–2024 年水文站网观测，上游拦沙与植被恢复共同降低了极端输沙波动。",
                    "SOURCE": "Test et al., 2026",
                },
                "composition": {
                    "kind": "auto",
                    "reading_task": "read_result",
                    "figure": {
                        "path": str(fig),
                        "aspect_ratio": round(900 / 560, 3),
                        "caption": "不同阶段输沙量对比 (Test et al., 2026)",
                    },
                    "facts": [
                        {
                            "heading": "阶段一（1980–1999）",
                            "text": "高强度人类扰动导致输沙变异系数维持高位，极端汛期贡献超过 65%。",
                        },
                        {
                            "heading": "阶段二（2000–2024）",
                            "text": "梯田与植被协同作用下，年均输沙量下降 58%，但局部暴雨仍有脉冲响应。",
                        },
                    ],
                },
                "speaker_note": "本页展示核心结果图与阶段对比。",
            },
        ],
    }
    plan_file = tmp_path / "deck_plan.json"
    plan_file.write_text(json.dumps(plan, ensure_ascii=False, indent=2), encoding="utf-8")
    out_pptx = tmp_path / "academic_plan.pptx"
    has_lo = find_soffice() is not None
    res = build_from_plan(
        plan_file,
        out_pptx,
        work_dir=tmp_path / "build_stage",
        run_qa_gate=has_lo,
    )
    assert out_pptx.is_file()
    assert res["slide_count"] == 2
    if has_lo:
        assert res["qa"]["render"]["slide_count"] == 2
        assert Path(res["qa"]["render"]["contact_sheet"]).is_file()


def test_route2_inspect_fill_and_extract_svg(tmp_path: Path):
    # First build a simple 2-slide PPTX as a template fixture.
    svg_dir = tmp_path / "src_svg"
    svg_dir.mkdir()
    (svg_dir / "01_slide.svg").write_text(_sample_svg("原始模板标题", "原始正文占位"), encoding="utf-8")
    (svg_dir / "02_slide.svg").write_text(_sample_svg("第二页标题", "第二页占位"), encoding="utf-8")
    tpl_pptx = tmp_path / "template.pptx"
    build_from_svg_dir(svg_dir, tpl_pptx, run_qa_gate=False)

    lib = inspect_pptx(tpl_pptx)
    assert lib["slide_count"] == 2

    plan_path = tmp_path / "fill_plan.json"
    scaff = scaffold_fill_plan(tpl_pptx, plan_path, slides="1,2")
    assert plan_path.is_file()
    assert scaff["slide_count"] == 2

    # Patch the first replacement text on slide 1 in fill_plan.json
    plan = json.loads(plan_path.read_text(encoding="utf-8"))
    if plan["slides"][0].get("replacements"):
        plan["slides"][0]["replacements"][0]["text"] = "替换后的全新标题"
    plan_path.write_text(json.dumps(plan, ensure_ascii=False, indent=2), encoding="utf-8")

    filled_pptx = tmp_path / "filled.pptx"
    fill_res = apply_fill_plan(plan_path, filled_pptx, run_qa_gate=False)
    assert filled_pptx.is_file()
    assert fill_res["status"] == "pass"

    # Extract SVGs from filled deck
    svg_extract_dir = tmp_path / "extracted_svg"
    ext_res = extract_pptx_to_svg(filled_pptx, svg_extract_dir, mode="flat")
    assert ext_res["slide_count"] == 2
    assert (svg_extract_dir / "svg" / "slide_01.svg").is_file()


def test_route3_reconstruct_layer_abc_and_diff(tmp_path: Path):
    # Create a synthetic 1280x720 slide image mockup.
    src_img = tmp_path / "mockup_01.png"
    im = Image.new("RGB", (1280, 720), (255, 255, 255))
    draw = ImageDraw.Draw(im)
    draw.rectangle([0, 0, 1280, 96], fill=(16, 47, 74))
    draw.rectangle([60, 150, 560, 560], fill=(219, 234, 254), outline=(37, 99, 235), width=3)
    draw.ellipse([160, 230, 460, 480], fill=(59, 130, 246))
    im.save(src_img)

    work_dir = tmp_path / "recon"
    init_res = init_reconstruction(src_img, work_dir)
    plan_path = Path(init_res["plan_path"])
    plan = json.loads(plan_path.read_text(encoding="utf-8"))

    plan["slides"][0]["shapes"] = [
        {"kind": "rect", "x": 0, "y": 0, "w": 1280, "h": 96, "fill": "#102F4A"},
        {"kind": "rect", "x": 620, "y": 150, "w": 600, "h": 410, "rx": 10, "fill": "#F8FAFC", "stroke": "#CBD5E1", "stroke_width": 2},
    ]
    plan["slides"][0]["raster_crops"] = [
        {
            "id": "chart_region",
            "src_bbox": [60, 150, 560, 560],
            "dst_box": [60, 150, 500, 410],
            "mask": "rounded",
        }
    ]
    plan["slides"][0]["texts"] = [
        {
            "x": 60,
            "y": 28,
            "w": 1160,
            "h": 44,
            "text": "图片重建示例：原生标题与可编辑正文",
            "font_family": "Microsoft YaHei",
            "font_size": 28,
            "font_weight": "bold",
            "fill": "#FFFFFF",
        },
        {
            "x": 650,
            "y": 190,
            "w": 540,
            "h": 180,
            "text": "左侧保留复杂原图裁切（Layer A）\n右侧全部转为原生矢量框与可编辑文字（Layer B/C）",
            "font_family": "Microsoft YaHei",
            "font_size": 18,
            "fill": "#1E293B",
        },
    ]
    plan_path.write_text(json.dumps(plan, ensure_ascii=False, indent=2), encoding="utf-8")

    out_pptx = tmp_path / "reconstructed.pptx"
    has_lo = find_soffice() is not None
    res = build_reconstruction(plan_path, out_pptx, run_qa_gate=has_lo)
    assert out_pptx.is_file()
    assert res["slide_count"] == 1
    if has_lo:
        assert res["diff"] is not None
        assert res["diff"]["slide_count"] == 1
        assert Path(res["diff"]["slides"][0]["comparison_png"]).is_file()


def test_scientific_body_style_normalization_opt_in(tmp_path: Path):
    from easyslides_mac.builder import _apply_template_scientific_body_style

    tpl_dir = tmp_path / "custom_tpl"
    tpl_dir.mkdir()
    (tpl_dir / "layouts.json").write_text(
        json.dumps(
            {
                "colors": {
                    "primary": "#4472C4",
                    "banner": "#4970B5",
                    "body": "#2B3A46",
                    "border": "#CED5DA",
                },
                "typography": {
                    "latin_font": "Times New Roman",
                    "east_asian_font": "Microsoft YaHei",
                    "body_pt": 17,
                    "caption_pt": 14,
                    "normalize_scientific_body": True,
                },
            }
        ),
        encoding="utf-8",
    )

    svg_path = tmp_path / "07_sci.svg"
    svg_path.write_text(
        """<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1280 720" width="1280" height="720">
  <g id="scientific-body">
    <rect x="64" y="200" width="400" height="40" fill="#194A37" stroke="#D3E4DA"/>
    <text x="80" y="228" font-family="Microsoft YaHei" font-size="24" font-weight="700" fill="#17251F" data-text-role="body">表头文本</text>
    <text x="80" y="280" font-family="KaiTi, Times New Roman" font-size="18.6667" font-weight="400" fill="#54786B" data-text-role="caption"><tspan x="80" y="280"><tspan font-family="KaiTi">图题</tspan><tspan font-family="Times New Roman"> Fig. 1</tspan></tspan></text>
  </g>
</svg>
""",
        encoding="utf-8",
    )

    _apply_template_scientific_body_style([svg_path], tpl_dir)
    out = svg_path.read_text(encoding="utf-8")
    assert "#4472C4" in out
    assert "#CED5DA" in out
    assert "#2B3A46" in out
    assert "#4970B5" in out
    assert "#194A37" not in out
    assert "#D3E4DA" not in out
    assert "#17251F" not in out
    assert "KaiTi" not in out
    assert "'Times New Roman', 'Microsoft YaHei'" in out
    assert 'font-size="22.6667"' in out
    assert 'font-size="18.6667"' in out
    assert 'font-weight="700"' not in out

    # Opt-out template must leave SVG untouched.
    tpl_default = tmp_path / "default_tpl"
    tpl_default.mkdir()
    (tpl_default / "layouts.json").write_text(json.dumps({"typography": {}}), encoding="utf-8")
    untouched = tmp_path / "untouched.svg"
    raw = '<svg xmlns="http://www.w3.org/2000/svg"><g id="scientific-body"><rect fill="#194A37"/></g></svg>'
    untouched.write_text(raw, encoding="utf-8")
    _apply_template_scientific_body_style([untouched], tpl_default)
    assert untouched.read_text(encoding="utf-8") == raw

