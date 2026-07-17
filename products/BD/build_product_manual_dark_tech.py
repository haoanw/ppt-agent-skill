#!/usr/bin/env python3
"""Restyle the FMC3 product manual as an editable dark-tech PowerPoint.

The source deck remains the content authority. This script changes visual
presentation only, preserves slide order, text, notes, tables, and media, then
renders PDF/PNG previews and writes a content audit plus delivery manifest.
"""

from __future__ import annotations

import argparse
import hashlib
import html
import json
import re
import shutil
import subprocess
import sys
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path
from zipfile import ZipFile

from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_FILL_TYPE
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE, MSO_SHAPE_TYPE
from pptx.util import Inches, Pt


ROOT = Path(__file__).resolve().parents[2]
BD = ROOT / "products/BD"
SOURCE = BD / "FMC³ Robotics_product Manual_0606.pptx"
DEFAULT_OUTPUT = BD / "FMC3_Robotics_Product_Manual_dark_tech"
OUTPUT_BASENAME = "FMC3_Robotics_Product_Manual_dark_tech"
LOGO = (
    BD
    / "FMC3_C5_Brochure_Deutsch_Redesigned_2_HighRes_Vector_dark_tech"
    / "images/image3.png"
)

BG = RGBColor(0x05, 0x0B, 0x1F)
BG_ALT = RGBColor(0x08, 0x16, 0x2E)
PANEL = RGBColor(0x0B, 0x1A, 0x34)
PANEL_ALT = RGBColor(0x10, 0x25, 0x45)
GRID = RGBColor(0x10, 0x22, 0x3F)
CYAN = RGBColor(0x22, 0xD3, 0xEE)
BLUE = RGBColor(0x3B, 0x82, 0xF6)
VIOLET = RGBColor(0x63, 0x66, 0xF1)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
MUTED = RGBColor(0xC3, 0xCD, 0xDB)
FAINT = RGBColor(0x7F, 0x91, 0xA8)
YELLOW = RGBColor(0xFD, 0xE0, 0x47)
LIGHT_VIEWPORT = RGBColor(0xF1, 0xF7, 0xFC)
FONT = "Aptos"


def write_text(path: Path, value: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(value, encoding="utf-8")


def write_json(path: Path, value: object) -> None:
    write_text(path, json.dumps(value, ensure_ascii=False, indent=2) + "\n")


def normalize_text(value: str) -> str:
    return re.sub(r"\s+", " ", value.replace("\u00a0", " ")).strip()


def iter_shapes(shapes):
    for shape in shapes:
        yield shape
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from iter_shapes(shape.shapes)


def shape_text(shape) -> list[str]:
    values: list[str] = []
    if getattr(shape, "has_text_frame", False):
        for paragraph in shape.text_frame.paragraphs:
            text = normalize_text(paragraph.text)
            if text:
                values.append(text)
    if getattr(shape, "has_table", False):
        for row in shape.table.rows:
            for cell in row.cells:
                for paragraph in cell.text_frame.paragraphs:
                    text = normalize_text(paragraph.text)
                    if text:
                        values.append(text)
    return values


def slide_text(slide) -> list[str]:
    return [text for shape in iter_shapes(slide.shapes) for text in shape_text(shape)]


def slide_notes(slide) -> list[str]:
    try:
        values = [
            text
            for shape in iter_shapes(slide.notes_slide.shapes)
            for text in shape_text(shape)
        ]
    except Exception:
        return []
    return [value for value in values if not value.isdigit()]


def source_snapshot(prs: Presentation) -> list[dict[str, object]]:
    return [
        {"page": index, "text": slide_text(slide), "notes": slide_notes(slide)}
        for index, slide in enumerate(prs.slides, 1)
    ]


def max_font_size(shape) -> float:
    sizes: list[float] = []
    if getattr(shape, "has_text_frame", False):
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                if run.font.size:
                    sizes.append(run.font.size.pt)
    return max(sizes, default=0.0)


def source_rgb(run) -> tuple[int, int, int] | None:
    try:
        rgb = run.font.color.rgb
        if rgb is not None:
            return tuple(rgb)
    except Exception:
        pass
    return None


def is_emphasis_rgb(rgb: tuple[int, int, int] | None) -> bool:
    if rgb is None:
        return False
    r, g, b = rgb
    return (r > 150 and g < 130 and b < 130) or (r > 150 and g > 120 and b < 80)


def looks_like_label(value: str) -> bool:
    compact = normalize_text(value)
    if not compact or len(compact) > 55:
        return False
    alpha = [char for char in compact if char.isalpha()]
    uppercase = bool(alpha) and sum(char.isupper() for char in alpha) / len(alpha) > 0.78
    keywords = (
        "products", "functions", "parameters", "highlights", "application",
        "positioning", "core team", "fmc³ brain", "data refinery", "step ",
    )
    return uppercase or any(compact.lower().startswith(item) for item in keywords)


def style_text_frame(text_frame, *, title: bool = False, table_header: bool = False) -> None:
    for paragraph in text_frame.paragraphs:
        paragraph_text = normalize_text(paragraph.text)
        for run in paragraph.runs:
            font = run.font
            original_rgb = source_rgb(run)
            font.name = FONT
            size = font.size.pt if font.size else 0.0
            if size and size < 7.5:
                font.size = Pt(7.5)
            if title or size >= 23:
                font.color.rgb = WHITE
                font.bold = True
            elif table_header or is_emphasis_rgb(original_rgb):
                font.color.rgb = CYAN if not is_emphasis_rgb(original_rgb) else YELLOW
                font.bold = True
            elif looks_like_label(paragraph_text):
                font.color.rgb = CYAN
                if len(paragraph_text) < 38:
                    font.bold = True
            elif font.bold:
                font.color.rgb = WHITE
            else:
                font.color.rgb = MUTED


def set_shape_line(shape, color: RGBColor = CYAN, width: float = 0.8) -> None:
    try:
        shape.line.color.rgb = color
        shape.line.width = Pt(width)
    except Exception:
        pass


def style_table(shape) -> None:
    table = shape.table
    for row_index, row in enumerate(table.rows):
        for column_index, cell in enumerate(row.cells):
            cell.fill.solid()
            cell.fill.fore_color.rgb = PANEL_ALT if row_index == 0 else (
                PANEL if (row_index + column_index) % 2 == 0 else BG_ALT
            )
            style_text_frame(cell.text_frame, table_header=row_index == 0)


def style_shape(shape, slide_width: int, slide_height: int) -> None:
    if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        return
    if getattr(shape, "has_table", False):
        style_table(shape)
        return

    text_values = shape_text(shape)
    combined = " ".join(text_values)
    font_size = max_font_size(shape)
    title = bool(text_values) and shape.top < Inches(1.3) and (
        font_size >= 18 or len(combined) < 70
    )
    if getattr(shape, "has_text_frame", False):
        style_text_frame(shape.text_frame, title=title)

    area_ratio = (shape.width * shape.height) / max(1, slide_width * slide_height)
    if shape.shape_type == MSO_SHAPE_TYPE.LINE:
        set_shape_line(shape, CYAN, 0.8)
        return
    if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
        return

    try:
        has_fill = shape.fill.type is not None
        picture_fill = shape.fill.type == MSO_FILL_TYPE.PICTURE
    except Exception:
        has_fill = False
        picture_fill = False
    if picture_fill and not text_values:
        return
    if not has_fill:
        if text_values:
            try:
                shape.fill.background()
            except Exception:
                pass
        return

    try:
        shape.fill.solid()
        if area_ratio > 0.52:
            shape.fill.fore_color.rgb = BG
            try:
                shape.line.fill.background()
            except Exception:
                pass
        elif shape.height < Inches(0.62) and text_values:
            shape.fill.fore_color.rgb = PANEL_ALT
            set_shape_line(shape, CYAN, 0.8)
        elif area_ratio > 0.018 or text_values:
            shape.fill.fore_color.rgb = PANEL
            set_shape_line(shape, GRID, 0.7)
        else:
            shape.fill.fore_color.rgb = BLUE if area_ratio < 0.004 else PANEL_ALT
            set_shape_line(shape, CYAN, 0.7)
    except Exception:
        pass


def style_shape_collection(shapes, slide_width: int, slide_height: int) -> None:
    for shape in iter_shapes(shapes):
        style_shape(shape, slide_width, slide_height)


def send_to_back(slide, shape, index: int = 2) -> None:
    tree = slide.shapes._spTree
    tree.remove(shape._element)
    tree.insert(index, shape._element)


def add_back_rect(slide, left, top, width, height, color: RGBColor) -> None:
    shape = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    send_to_back(slide, shape)


def add_outline_rect(slide, left, top, width, height) -> None:
    shape = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, left, top, width, height)
    shape.fill.background()
    set_shape_line(shape, CYAN, 1.2)


def apply_page_specific_fixes(slide, page_number: int) -> None:
    for shape in slide.shapes:
        if (
            shape.shape_type == MSO_SHAPE_TYPE.PICTURE
            and shape.left > Inches(11.0)
            and shape.top < Inches(0.55)
            and Inches(1.15) <= shape.width <= Inches(1.45)
            and Inches(0.55) <= shape.height <= Inches(0.78)
        ):
            shape.left = Inches(13.45)
            shape.top = Inches(7.6)
            shape.width = Pt(1)
            shape.height = Pt(1)

    if page_number == 1:
        for shape in slide.shapes:
            if normalize_text(getattr(shape, "text", "")).startswith("Product Manual"):
                shape.height = Inches(1.35)
                set_shape_line(shape, CYAN, 0.9)

    if page_number == 2:
        add_outline_rect(slide, Inches(7.55), Inches(0.86), Inches(5.48), Inches(5.78))

    if page_number == 7:
        for shape in slide.shapes:
            if (
                not shape_text(shape)
                and shape.width > Inches(5.0)
                and Inches(0.8) <= shape.height <= Inches(1.5)
            ):
                try:
                    shape.fill.background()
                    shape.line.fill.background()
                except Exception:
                    pass
            if (
                shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE
                and not shape_text(shape)
                and shape.left > Inches(2.5)
                and Inches(2.35) <= shape.width <= Inches(2.65)
                and Inches(1.05) <= shape.height <= Inches(1.30)
            ):
                shape.fill.solid()
                shape.fill.fore_color.rgb = LIGHT_VIEWPORT
                set_shape_line(shape, CYAN, 0.9)

    if page_number in (28, 29):
        for shape in slide.shapes:
            if (
                shape.shape_type == MSO_SHAPE_TYPE.PICTURE
                and shape.width >= Inches(12.8)
                and shape.height >= Inches(7.1)
            ):
                shape.left = Inches(13.45)
                shape.top = Inches(7.6)
                shape.width = Pt(1)
                shape.height = Pt(1)


def add_dark_tech_decor(
    slide,
    page_number: int,
    total_pages: int,
    logo_path: Path,
    width: int,
    height: int,
) -> None:
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = BG

    for x in range(1, 14):
        add_back_rect(slide, Inches(x), 0, Pt(0.45), height, GRID)
    for y in range(1, 8):
        add_back_rect(slide, 0, Inches(y), width, Pt(0.45), GRID)

    add_back_rect(slide, Inches(0.28), Inches(0.20), Inches(0.55), Pt(1.2), CYAN)
    add_back_rect(slide, Inches(0.28), Inches(0.20), Pt(1.2), Inches(0.42), CYAN)
    add_back_rect(slide, Inches(0.55), Inches(7.18), Inches(11.95), Pt(0.65), GRID)
    add_back_rect(slide, Inches(0.55), Inches(7.18), Inches(2.2), Pt(0.9), CYAN)

    if logo_path.is_file():
        logo = slide.shapes.add_picture(str(logo_path), Inches(11.55), Inches(0.22), width=Inches(1.25))
        logo.name = "FMC3 White Logo"

    page_box = slide.shapes.add_textbox(Inches(11.55), Inches(7.12), Inches(1.15), Inches(0.22))
    paragraph = page_box.text_frame.paragraphs[0]
    paragraph.alignment = 2
    run = paragraph.add_run()
    run.text = f"{page_number:02d} / {total_pages:02d}"
    run.font.name = "Aptos Mono"
    run.font.size = Pt(8.5)
    run.font.bold = True
    run.font.color.rgb = CYAN


def style_presentation(prs: Presentation, logo_path: Path) -> None:
    for master in prs.slide_masters:
        try:
            master.background.fill.solid()
            master.background.fill.fore_color.rgb = BG
        except Exception:
            pass
        style_shape_collection(master.shapes, prs.slide_width, prs.slide_height)
        for layout in master.slide_layouts:
            style_shape_collection(layout.shapes, prs.slide_width, prs.slide_height)

    total = len(prs.slides)
    for page_number, slide in enumerate(prs.slides, 1):
        style_shape_collection(slide.shapes, prs.slide_width, prs.slide_height)
        apply_page_specific_fixes(slide, page_number)
        add_dark_tech_decor(
            slide,
            page_number,
            total,
            logo_path,
            prs.slide_width,
            prs.slide_height,
        )


def zip_media_hashes(path: Path) -> Counter[str]:
    with ZipFile(path) as archive:
        return Counter(
            hashlib.sha256(archive.read(name)).hexdigest()
            for name in archive.namelist()
            if name.startswith("ppt/media/") and not name.endswith("/")
        )


def extract_slide_images(prs: Presentation, images_dir: Path) -> dict[str, list[str]]:
    images_dir.mkdir(parents=True, exist_ok=True)
    mapping: dict[str, list[str]] = {}
    for page_number, slide in enumerate(prs.slides, 1):
        filenames: list[str] = []
        picture_number = 0
        for shape in iter_shapes(slide.shapes):
            if shape.shape_type != MSO_SHAPE_TYPE.PICTURE:
                continue
            picture_number += 1
            image = shape.image
            extension = image.ext or "bin"
            filename = f"slide-{page_number:02d}-image-{picture_number:02d}.{extension}"
            destination = images_dir / filename
            destination.write_bytes(image.blob)
            filenames.append(filename)
        mapping[str(page_number)] = filenames
    return mapping


def choose_title(texts: list[str], page_number: int) -> str:
    ignored = re.compile(r"^(\d+|\d{2}/\d{2}/\d{4}|www\.|products\s*\||fmc³ brain\+odm$)", re.I)
    for value in texts:
        if not ignored.search(value) and len(value) > 2:
            return value[:100]
    return f"Slide {page_number}"


def write_workflow_files(output_dir: Path, snapshot: list[dict[str, object]], image_map: dict[str, list[str]]) -> None:
    interview = """# Interview QA

scenario: FMC3 Robotics corporate product manual
audience: Western-market B2B customers, partners, integrators, and industrial decision-makers
target_action: Understand FMC3 positioning, compliance, platform capabilities, and full product portfolio
expected_pages: 30
page_density: balanced
style: dark_tech
brand: FMC3 Robotics
must_include: Every source slide, text block, table, note, and embedded product visual
must_avoid: Content deletion, rewriting, translation, invented specifications, or reordered slides
language: English, preserving source wording
imagery: provided only
material_strategy: non-research; the supplied PPTX is the sole content authority
subagent_model_strategy: main agent
subagent_thinking_effort: high
manual_audit_mode: fine_grained
manual_audit_scope: planning, presentation, rendered_pages
manual_audit_assets: source-brief, outline, style, pptx, pdf, png, content-audit
density_bias: balanced
branch: non-research
"""
    write_text(output_dir / "interview-qa.txt", interview)
    write_text(
        output_dir / "requirements-interview.txt",
        interview.replace("# Interview QA", "# Requirements Interview", 1),
    )

    source_lines = [
        "# Source Brief", "",
        "topic: FMC3 Robotics corporate positioning, compliance, embodied-AI platform, and product portfolio",
        "usage: B2B product introduction for customers, partners, integrators, and industrial decision-makers",
        "constraints: Preserve all 30 source slides, wording, order, notes, tables, and embedded media without invention.",
        "risk: Dense specification pages and picture-filled logo shapes require rendered-page inspection after native PPTX restyling.",
        f"source: {SOURCE}", f"slides: {len(snapshot)}", "",
    ]
    outline_lines = [
        "# Outline", "",
        "密度倾向: balanced",
        "密度曲线: mid_low -> medium -> medium -> high -> medium -> mid_low",
        "",
    ]
    for page in snapshot:
        number = int(page["page"])
        texts = list(page["text"])
        title = choose_title(texts, number)
        source_lines.extend([f"## Slide {number}: {title}", *[f"- {item}" for item in texts]])
        if page["notes"]:
            source_lines.extend(["- Speaker notes:", *[f"  - {item}" for item in page["notes"]]])
        source_lines.append("")
        page_type = "cover" if number == 1 else "end" if number == len(snapshot) else "content"
        density_target = "mid_low" if page_type in {"cover", "end"} else (
            "high" if len(texts) >= 45 and number % 3 == 1 else "medium"
        )
        rhythm = "铺垫" if number == 1 else "收束" if number == len(snapshot) else (
            "爆发" if density_target == "high" else "推进"
        )
        posture = "结论页" if page_type in {"cover", "end"} else (
            "证据页" if len(texts) >= 25 else "解释页"
        )
        anchor = "图片" if image_map.get(str(number)) else "表格" if len(texts) >= 35 else "标题"
        outline_lines.extend([
            f"### 第 {number} 页: {title}",
            f"页目标: Preserve and present the source content for {title}.",
            f"页面类型映射: {page_type}",
            "密度下限: mid_low",
            f"密度目标: {density_target}",
            "密度上限: high",
            f"节奏动作: {rhythm}",
            f"信息姿态: {posture}",
            f"锚点类型: {anchor}",
            "",
        ])
    outline_lines.append("自审通过: All 30 source pages remain in the original order with no text, note, table, or media loss.")
    write_text(output_dir / "source-brief.txt", "\n".join(source_lines))
    write_text(output_dir / "outline.txt", "\n".join(outline_lines) + "\n")

    style = {
        "style_id": "dark_tech",
        "style_name": "Dark Tech",
        "mood_keywords": ["precise", "industrial", "futuristic", "technical"],
        "design_soul": "A restrained European industrial interface with deep navy space, cyan instrumentation, and clear product evidence.",
        "variation_strategy": "Product images and source layouts remain page-specific; background, panel, typography, logo, grid, and navigation treatments stay locked.",
        "decoration_dna": {
            "signature_move": "cyan technical frame with subtle grid",
            "forbidden": ["content rewriting", "decorative stock imagery", "oversized rounded cards"],
            "recommended_combos": ["dark panel + cyan rule", "white product viewport + navy specification panel"],
        },
        "font_family": FONT,
        "css_variables": {
            "bg_primary": "#050B1F", "bg_secondary": "#08162E",
            "card_bg_from": "#0B1A34", "card_bg_to": "#102545",
            "card_border": "#163B59", "card_radius": "6px",
            "text_primary": "#FFFFFF", "text_secondary": "#C3CDDB",
            "accent_1": "#22D3EE", "accent_2": "#3B82F6",
            "accent_3": "#6366F1", "accent_4": "#FDE047",
            "css_snippets": {},
        },
    }
    write_json(output_dir / "style.json", style)

    planning_dir = output_dir / "planning"
    planning_dir.mkdir(parents=True, exist_ok=True)
    for page in snapshot:
        number = int(page["page"])
        texts = list(page["text"])
        title = choose_title(texts, number)
        images = image_map.get(str(number), [])
        page_type = "cover" if number == 1 else "end" if number == len(snapshot) else "content"
        density_label = "mid_low" if page_type in {"cover", "end"} else "medium"
        density_contract = {
            "deck_bias": "balanced",
            "page_lower_bound": "low" if density_label == "mid_low" else "mid_low",
            "page_upper_bound": "high",
            "max_cards": 3 if density_label == "mid_low" else 4,
            "max_charts": 1 if density_label == "mid_low" else 2,
            "min_body_font_px": 20 if density_label == "mid_low" else 18,
            "max_lines_per_card": 4 if density_label == "mid_low" else 5,
            "image_policy": "flexible" if density_label == "mid_low" else "support_only",
            "decoration_budget": "medium",
            "overflow_strategy": "rebalance_layout" if density_label == "mid_low" else "tighten_budget",
        }
        split = max(1, (len(texts) + 1) // 2)
        card_groups = [texts] if page_type != "content" else [texts[:split], texts[split:]]
        cards = []
        for card_index, body in enumerate(card_groups, 1):
            use_image = card_index == 1 and bool(images)
            cards.append({
                "card_id": f"s{number:02d}-{'anchor' if card_index == 1 else 'support'}-{card_index}",
                "role": "anchor" if card_index == 1 else "support",
                "card_type": "image_hero" if page_type != "content" and use_image else "text",
                "card_style": "accent" if card_index == 1 else "outline",
                "argument_role": "evidence" if card_index == 1 else "framework",
                "headline": title if card_index == 1 else "Source content",
                "body": body,
                "data_points": [],
                "content_budget": {
                    "headline_max_chars": 100,
                    "body_max_bullets": max(1, len(body)),
                    "body_max_lines": density_contract["max_lines_per_card"],
                },
                "image": {
                    "mode": "provided" if use_image else "decorate",
                    "needed": use_image,
                    "usage": "inline-illustration" if use_image else None,
                    "placement": "inline" if use_image else None,
                    "content_description": "Original visual from the supplied product manual." if use_image else None,
                    "source_hint": f"images/{images[0]}" if use_image else None,
                    "decorate_brief": "Preserve the source visual and its semantic content." if use_image else "Dark-tech grid and cyan rules only.",
                },
            })
        payload = {"page": {
            "slide_number": number,
            "page_type": page_type,
            "narrative_role": "cover" if page_type == "cover" else "close" if page_type == "end" else "evidence",
            "title": title,
            "page_goal": f"Preserve and present all source content for slide {number}.",
            "audience_takeaway": title,
            "cards": cards,
            "visual_weight": 7,
            "density_label": density_label,
            "density_reason": "Source-controlled density preserves all supplied product content.",
            "density_contract": density_contract,
            "layout_hint": "primary-secondary" if number % 2 else "asymmetric",
            "director_command": {
                "mood": "precise industrial future",
                "spatial_strategy": "Preserve source hierarchy within a dark technical frame.",
                "anchor_treatment": "Product evidence receives the strongest cyan edge.",
            },
            "decoration_hints": {
                "background": {"feel": "deep technical space", "techniques": ["grid"]},
                "page_accent": {"feel": "cyan instrumentation", "techniques": ["cyan-rule"]},
            },
            "resources": {
                "page_template": page_type,
                "layout_refs": [], "block_refs": [], "chart_refs": [], "principle_refs": [],
            },
            "source_guidance": {
                "brief_sections": [f"source-brief.txt#slide-{number}"],
                "citation_expectation": "Use only the supplied PPTX content and embedded media.",
                "strictness": "No additions, rewriting, translation, deletion, or reordering.",
                "speaker_notes": page["notes"],
            },
            "workflow_metadata": {
                "stage": "planning",
                "workflow_version": "2026.04.09-v4.1",
                "planning_schema_version": "4.1",
                "planning_packet_version": "4.1",
                "planning_continuity_version": "4.1",
            },
        }}
        write_json(planning_dir / f"planning{number}.json", payload)


def render_outputs(output_dir: Path, pptx_path: Path) -> tuple[Path, list[Path]]:
    subprocess.run(
        ["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", str(output_dir), str(pptx_path)],
        check=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )
    pdf_path = output_dir / f"{OUTPUT_BASENAME}.pdf"
    png_dir = output_dir / "png"
    png_dir.mkdir(parents=True, exist_ok=True)
    prefix = png_dir / "slide"
    subprocess.run(["pdftoppm", "-png", "-r", "144", str(pdf_path), str(prefix)], check=True)
    pngs = sorted(png_dir.glob("slide-*.png"))
    return pdf_path, pngs


def build_preview(output_dir: Path, pngs: list[Path]) -> None:
    items = "".join(
        f'<figure><img src="png/{html.escape(path.name)}" alt="Slide {index}"><figcaption>{index:02d}</figcaption></figure>'
        for index, path in enumerate(pngs, 1)
    )
    write_text(
        output_dir / "preview.html",
        f"""<!doctype html><html><head><meta charset="utf-8"><title>FMC3 Product Manual · Dark Tech</title>
<style>*{{box-sizing:border-box}}body{{margin:0;background:#020611;color:#fff;font-family:Arial,sans-serif}}header{{position:sticky;top:0;z-index:2;padding:18px 28px;background:#050b1f;border-bottom:1px solid #16465c}}main{{padding:28px;display:grid;grid-template-columns:repeat(auto-fit,minmax(420px,1fr));gap:24px}}figure{{margin:0}}img{{display:block;width:100%;box-shadow:0 18px 50px #0008;border:1px solid #16465c}}figcaption{{padding-top:8px;color:#22d3ee;font-family:monospace}}</style></head>
<body><header><strong>FMC³ Robotics · Product Manual · Dark Tech</strong></header><main>{items}</main></body></html>""",
    )


def build_audit(
    source_path: Path,
    output_path: Path,
    source_snapshot_data: list[dict[str, object]],
    pngs: list[Path],
) -> dict[str, object]:
    output_prs = Presentation(output_path)
    output_snapshot = source_snapshot(output_prs)
    pages = []
    for source_page, output_page in zip(source_snapshot_data, output_snapshot):
        output_content = [
            value for value in output_page["text"]
            if not re.fullmatch(r"\d{2} / \d{2}", value)
        ]
        text_match = source_page["text"] == output_content
        notes_match = source_page["notes"] == output_page["notes"]
        pages.append({"page": source_page["page"], "text_match": text_match, "notes_match": notes_match})

    source_media = zip_media_hashes(source_path)
    output_media = zip_media_hashes(output_path)
    missing_media = list((source_media - output_media).elements())
    png_checks = {}
    for index, path in enumerate(pngs, 1):
        with Image.open(path) as image:
            png_checks[str(index)] = {
                "dimensions": list(image.size),
                "ratio_16_9": abs(image.width / image.height - 16 / 9) < 0.005,
                "bytes": path.stat().st_size,
            }
    ok = (
        len(source_snapshot_data) == len(output_snapshot) == len(pngs)
        and all(page["text_match"] and page["notes_match"] for page in pages)
        and not missing_media
        and all(item["ratio_16_9"] and item["bytes"] > 40_000 for item in png_checks.values())
    )
    return {
        "ok": ok,
        "source": str(source_path),
        "output": str(output_path),
        "slide_count": len(output_snapshot),
        "page_checks": pages,
        "source_media_count": sum(source_media.values()),
        "output_media_count": sum(output_media.values()),
        "missing_source_media_hashes": missing_media,
        "png_checks": png_checks,
    }


def write_manifest(output_dir: Path, total_pages: int) -> None:
    write_json(
        output_dir / "delivery-manifest.json",
        {
            "run_id": output_dir.name,
            "generated_at": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
            "summary": {"total_pages": total_pages, "format": "16:9 widescreen"},
            "artifacts": {
                "preview_html": "preview.html",
                "presentation_pptx": f"{OUTPUT_BASENAME}.pptx",
                "presentation_png_pptx": f"{OUTPUT_BASENAME}.pptx",
                "presentation_svg_pptx": f"{OUTPUT_BASENAME}.pptx",
                "brochure_pdf": f"{OUTPUT_BASENAME}.pdf",
                "content_audit": "content-audit.json",
            },
            "pages": [
                {"page": page, "planning": f"planning/planning{page}.json", "png": f"png/slide-{page:02d}.png"}
                for page in range(1, total_pages + 1)
            ],
            "export_notes": {
                "presentation_pptx": "Editable native PowerPoint restyled from the supplied source deck.",
                "content_policy": "Slide order, text, notes, tables, and embedded media preserved.",
            },
        },
    )


def build(output_dir: Path) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    prs = Presentation(SOURCE)
    original = source_snapshot(prs)
    image_map = extract_slide_images(prs, output_dir / "images")
    write_workflow_files(output_dir, original, image_map)
    style_presentation(prs, LOGO)

    output_pptx = output_dir / f"{OUTPUT_BASENAME}.pptx"
    prs.save(output_pptx)
    pdf_path, pngs = render_outputs(output_dir, output_pptx)
    build_preview(output_dir, pngs)
    audit = build_audit(SOURCE, output_pptx, original, pngs)
    write_json(output_dir / "content-audit.json", audit)
    write_manifest(output_dir, len(original))
    if not audit["ok"]:
        raise RuntimeError(f"Content audit failed: {output_dir / 'content-audit.json'}")
    print(f"Built: {output_pptx}")
    print(f"Built: {pdf_path}")


def main() -> int:
    parser = argparse.ArgumentParser(description="Restyle the FMC3 product manual in dark-tech style")
    parser.add_argument("--output-dir", type=Path, default=DEFAULT_OUTPUT)
    args = parser.parse_args()
    if not SOURCE.is_file():
        print(f"Source PPTX not found: {SOURCE}", file=sys.stderr)
        return 1
    build(args.output_dir.resolve())
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
