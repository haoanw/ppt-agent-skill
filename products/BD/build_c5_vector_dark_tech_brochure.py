#!/usr/bin/env python3
"""Build the redesigned FMC3 C5 German brochure in dark_tech style."""

from __future__ import annotations

import argparse
import html
import json
import re
import sys
from datetime import datetime, timezone
from pathlib import Path
from zipfile import ZipFile

from lxml import etree
from PIL import Image, ImageDraw

import build_c5_dark_tech_brochure as base


ROOT = Path(__file__).resolve().parents[2]
SOURCE = ROOT / "products/BD/FMC3_C5_Brochure_Deutsch_Redesigned_2_HighRes_Vector.docx"
DEFAULT_OUTPUT = ROOT / "products/BD/FMC3_C5_Brochure_Deutsch_Redesigned_2_HighRes_Vector_dark_tech"
PDF_NAME = "FMC3_C5_Brochure_Deutsch.pdf"

PAGE_W = 1120
PAGE_H = 1584
TOTAL_PAGES = 6
FOOTER_TEXT = "FMC3 Robotics · Ingolstadt, Deutschland · www.FMC3-robotics.ai"
APPROVED_TEXT_CORRECTIONS = {"Wang Haoan": "Haoan Wang"}

base.TOTAL_PAGES = TOTAL_PAGES

PAGE_TEXT = {
    1: [
        "COMMERCIAL CLEANING ROBOT",
        "C5",
        "Große Flächen",
        "Weniger Aufwand",
        "Intelligenter gewerblicher Reinigungsroboter für mittlere und große Flächen.",
        "Kehren",
        "Scheuern",
        "Saugen",
        "1.980 m²/h",
        "Theoretische Flächenleistung",
        "550 mm",
        "Schrubbreite",
        "3 h",
        "Reinigungsdauer",
        "FMC3 Robotics",
        "Ingolstadt · Deutschland",
        "www.FMC3-robotics.ai",
    ],
    2: [
        "LEISTUNG IM ALLTAG",
        "Vier Systeme. Ein sauberer Prozess.",
        "Der C5 verbindet gründliche Bodenreinigung, sichere Navigation, autonome Versorgung und wartungsarme Technik.",
        "Kehren und Scheuern in einem Arbeitsgang – mit hoher Wasseraufnahme und konsequenter Kantenreinigung.",
        "Reinigungsmechanik",
        "Doppelwalzenbürsten entfernen auch schwere Verschmutzungen.",
        "Zwei-Kammer-Saugfuß reduziert Restwasser in Fugen.",
        "Navigation & Anpassung",
        "Laser- und visuelle Lokalisierung für unterschiedliche Szenen.",
        "Randbürste arbeitet dicht an Kanten im geschlossenen Kreislauf.",
        "Autonomer Betrieb",
        "Autonomes Laden sowie Frisch- und Schmutzwasser-Management.",
        "Tür- und Aufzuganbindung für durchgängige Abläufe.",
        "Wartung & Lebensdauer",
        "Selbstreinigender Schmutzwassertank reduziert täglichen Aufwand.",
        "Gebläselebensdauer 10.000 h und langlebige Komponenten.",
        "95 %",
        "Schmutzaufnahme",
        "<68 dB(A)",
        "Geräuschpegel",
        "25 kg",
        "Bodendruck",
        "4 min",
        "Tank-Selbstreinigung",
    ],
    3: [
        "TECHNISCHE DATEN",
        "Kompakt gebaut. Für große Aufgaben.",
        "Leistungsdaten, Abmessungen und typische Einsatzbereiche des C5 im Überblick.",
        "C5 · PRODUKTPARAMETER",
        "90 L",
        "Großer Wassertank",
        "24 V / 80 Ah",
        "Batterie",
        "170 kg",
        "Gesamtgewicht",
        "1.800 W",
        "Max. Leistung",
        "200–240 V",
        "Nenneingangsspannung",
        "8 L",
        "Tankvolumen Arbeitsstation",
        "Frischwasser 7–10 L/min · Ablass 10–15 L/min",
        "680 × 820 × 1.085 mm",
        "Arbeitsstation: 520 × 310 × 1.105 mm",
        "EINSATZSZENARIEN",
        "Große Supermärkte",
        "Bürogebäude",
        "Verkehrsknoten",
        "Fabriken",
        "Gewerbekomplexe",
        "Krankenhäuser",
        "GEEIGNETE BODENBELÄGE",
        "Zementboden",
        "Marmor",
        "Epoxidboden",
        "PVC",
        "Korundboden",
    ],
    4: [
        "WIRTSCHAFTLICHKEIT",
        "Ein Business Case, der sich rechnen lässt.",
        "Beispielrechnung gegenüber manueller Reinigung – berechnet über 60 Monate und 1.000 Betriebsstunden pro Jahr.",
        "9.089 €",
        "GESCHÄTZTE EINSPARUNGEN PRO JAHR",
        "1,3 Jahre",
        "AMORTISATIONSZEIT / ROI",
        "1,54 Mio. m²",
        "GEREINIGTE FLÄCHE PRO JAHR",
        "JÄHRLICHER VERGLEICH",
        "Aktuelle Lösung",
        "20.780 €",
        "C5 · 60 Monate",
        "11.691 €",
        "ROI- UND EINSPARUNGSRECHNER",
        "Reinigungsleistung",
        "Von der Reinigungsfläche berechnete Kosten",
        "Verrechneter Stundenlohn",
        "24 €/h",
        "15 €/h",
        "50 €/h",
        "Reinigungsstunden pro Tag",
        "3 h/Tag",
        "1 h/Tag",
        "12 h/Tag",
        "Reinigungstage pro Jahr",
        "260 Tage/Jahr",
        "100 Tage/Jahr",
        "365 Tage/Jahr",
        "Manuelle Scheuersaugmaschine",
        "Anschaffungskosten",
        "8.500 €",
        "Wartung",
        "30 €/Monat",
        "GESCHÄTZTE JÄHRLICHE EINSPARUNGEN",
        "entspricht 757 € / Monat",
        "1.544.400 m²",
        "FLÄCHE / JAHR",
        "Individuelles Angebot erhalten",
        "Berechnungsbasis",
        "24 €/h · 3 h/Tag · 260 Tage/Jahr · 8.500 € Maschine · 30 €/Monat Wartung",
        "Kundennahe Beispielrechnung; Werte können projektspezifisch angepasst werden.",
    ],
    5: [
        "FMC3 ROBOTICS",
        "Wissenschaft. Industrie. Umsetzung.",
        "Europäische Wissenschaftler für intelligente Systeme – mit Erfahrung aus Automobilindustrie, KI-Forschung und internationalem B2B-Geschäft.",
        "Über FMC3 Robotics",
        "FMC3 bringt moderne Robotiktechnologie vom Standort Ingolstadt in den europäischen Markt. Das Unternehmen wurde im Oktober 2025 von Prof. Dr. Björn Giesler, Dr. Wang Cheng und Dr. David Fan gegründet. Das Portfolio umfasst Reinigungs-, Begrüßungs- und Lieferroboter.",
        "Wissenschaftliche Exzellenz, industrielle Erfahrung und ein internationales Netzwerk bilden die Grundlage für zuverlässige und skalierbare Lösungen im professionellen Einsatz.",
        "Okt. 2025",
        "GEGRÜNDET",
        "Ingolstadt",
        "HAUPTSITZ",
        "DE & Shanghai",
        "ENTWICKLUNG",
        "30 / Monat",
        "LIEFERKAPAZITÄT",
        "DAS KERNTEAM",
        "Dr. Wang Cheng",
        "CEO",
        "Dr.-Ing. der Technischen Universität München (TUM). Experte für autonomes Fahren. Architekt von Audis Intelligent Driving System. Track Record in Produktion und Auslieferung von über 1 Mio. intelligenten Fahrzeugen.",
        "Prof. Dr. Björn Giesler",
        "CTO",
        "PhD in Informatik am KIT; anerkannte Autorität in KI-Architektur. Über 50 Core-Patente zu Sensing und Datenverarbeitung. Führung von F&E-Teams mit mehr als 500 Mitgliedern.",
        "Fan David, PhD",
        "COO",
        "Ehemaliger Präsident von Marelli China. Verantwortlich für Geschäftsbereiche mit über 1,5 Mrd. €. Post-Merger-Integrationen und Supply-Chain-Restrukturierungen.",
        "Dr. Huang Dong",
        "CPO",
        "Dr.-Ing. in Informatik vom KIT. Ehemaliger Scientist bei Siemens Labs und CTO von TerraIT. Über 15 Jahre Expertise in Datensimulation und Sim-to-Real-Domänen.",
        "Florian Obermeier",
        "MD Germany",
        "Mitgründer der SANEON GmbH. Über 16 Jahre B2B-Erfahrung mit OEMs und Tier-1-Automobilzulieferern. Markterschließung in Europa, Asien und den USA.",
        "LOKALE BETREUUNG · INTERNATIONALE ENTWICKLUNG",
        "Persönliche Betreuung in Deutschland · Technische Entwicklung in Ingolstadt und Shanghai",
        "Reinigungsroboter",
        "Begrüßungsroboter",
        "Lieferroboter",
    ],
    6: [
        "SERVICE & LIVE-DEMONSTRATION",
        "Vom ersten Test bis zum laufenden Betrieb.",
        "FMC3 begleitet Analyse, Demonstration, Inbetriebnahme, Schulung und After-Sales-Support aus Ingolstadt.",
        "PROJEKTABLAUF",
        "SERVICEVERSPRECHEN",
        "Analyse & Bedarfsprüfung",
        "Anforderungen verstehen und Gegebenheiten vor Ort prüfen.",
        "Live-Demonstration / Test",
        "C5 bei Ihnen vor Ort oder am Hauptsitz in Ingolstadt erleben.",
        "Lieferung & Inbetriebnahme",
        "Liefern, installieren, testen und gemeinsam in Betrieb nehmen.",
        "Schulung",
        "Ihr Team praxisnah und individuell einweisen.",
        "Support",
        "Schnelle Reaktion und zuverlässiger After-Sales-Support.",
        "Individuelle Video-Schulungen",
        "Express-Reaktion innerhalb von 0,5 Stunden",
        "Servicenetz in 150+ Städten",
        "Erste Wartung kostenlos",
        "ERLEBEN SIE DEN C5 LIVE",
        "Wir senden Ihnen weitere Informationen und laden Sie zu einer Demonstration ein.",
        "Ihr Ansprechpartner",
        "Haoan Wang",
        "FMC3 Robotics",
        "Ingolstadt · Deutschland",
        "FMC3 Robotics · C5 Autonomous Commercial Cleaning Robot",
    ],
}


def normalize_text(value: str) -> str:
    value = html.unescape(value).replace("\u00a0", " ")
    value = value.replace("‑", "-").replace("–", "–")
    return re.sub(r"\s+", " ", value).strip()


def source_xml_segments(source: Path) -> list[str]:
    with ZipFile(source) as archive:
        root = etree.fromstring(archive.read("word/document.xml"))
    segments: list[str] = []
    for node in root.iter():
        if etree.QName(node).localname != "t" or not node.text:
            continue
        value = normalize_text(node.text)
        if len(value) < 2 or value in {"•", "“", "ⓘ"}:
            continue
        if value not in segments:
            segments.append(value)
    return segments


def prepare_display_assets(output_dir: Path) -> None:
    """Remove a stray source-layout character without changing the archived media."""
    source_image = output_dir / "images/image1.jpeg"
    display_image = output_dir / "runtime/image1-display.png"
    with Image.open(source_image).convert("RGB") as image:
        draw = ImageDraw.Draw(image)
        draw.rectangle((0, 440, 30, 510), fill=(248, 248, 248))
        image.save(display_image, "PNG")


def workflow_files(output_dir: Path) -> None:
    common = """scenario: Deutsche B2B-Produktbroschüre für den autonomen Reinigungsroboter C5
audience: Facility-Management, Gebäudereiniger, Betreiber großer Gewerbeflächen und B2B-Entscheider
target_action: Wirtschaftlichkeit und Leistungsfähigkeit verstehen und eine Live-Demonstration anfragen
expected_pages: 6
page_density: balanced
style: dark_tech
brand: FMC3 Robotics; Originalinhalte, Originalbilder, Logo und Kontaktdaten beibehalten
must_include: Alle Texte, Kennzahlen, Produktdaten, ROI-Werte, Teamprofile, Prozessschritte und Medien aus der Quelldatei
must_avoid: Keine Kürzungen, keine Übersetzungen, keine erfundenen Aussagen, kein linker Seitenkopf
language: Deutsch
imagery: provided
material_strategy: non-research; ausschließlich bereitgestellte DOCX
subagent_model_strategy: Hauptagent
subagent_thinking_effort: hoch
manual_audit_mode: fine_grained
manual_audit_scope: page_html, page_review
manual_audit_assets: planning, html, png, content-audit
density_bias: balanced
branch: 非research
"""
    base.write_text(output_dir / "interview-qa.txt", "# Interview QA\n\n" + common)
    base.write_text(output_dir / "requirements-interview.txt", "# Requirements Interview\n\n" + common)

    brief = [
        "# Source Brief",
        "",
        "topic: FMC3 C5 Reinigungsroboter, sechsseitige deutsche B2B-Broschüre",
        "usage: Produktvorstellung, Wirtschaftlichkeitsnachweis und Einladung zur Live-Demonstration",
        "constraints: Sämtliche Textbox-Inhalte und alle 42 eingebetteten Medien werden erhalten.",
        "risk: Der DOCX-Inhalt liegt überwiegend in Floating Shapes; die Prüfung erfolgt direkt gegen document.xml.",
        "",
        f"Quelle: {SOURCE.relative_to(ROOT)}",
        "",
    ]
    for page, segments in PAGE_TEXT.items():
        brief.extend([f"## Seite {page}", ""])
        brief.extend(f"- {segment}" for segment in segments)
        brief.append("")
    base.write_text(output_dir / "source-brief.txt", "\n".join(brief))

    outline_lines = [
        "# Outline",
        "",
        "密度倾向: balanced",
        "密度曲线: mid_low -> medium -> high -> high -> high -> medium",
        "",
    ]
    pages = [
        ("C5 Produktauftakt", "Den C5, seine Kernaufgabe und drei Leistungswerte positionieren", "cover", "mid_low", "铺垫", "结论页", "图片"),
        ("Vier Reinigungssysteme", "Mechanik, Navigation, autonomen Betrieb und Wartung vollständig erklären", "content", "medium", "推进", "解释页", "表格"),
        ("Technische Daten", "Produktparameter, Abmessungen, Einsatzszenarien und Bodenbeläge verdichten", "content", "high", "推进", "证据页", "图片"),
        ("Wirtschaftlichkeit", "ROI-Basis, Vergleich und Ergebnis transparent belegen", "content", "high", "爆发", "证据页", "KPI"),
        ("FMC3 und Kernteam", "Unternehmen, Fakten und fünf Teamprofile als Vertrauensnachweis zeigen", "content", "medium", "推进", "证据页", "图片"),
        ("Service und Live-Demo", "Projektablauf und Serviceversprechen in einen persönlichen CTA überführen", "content", "medium", "收束", "结论页", "图表"),
    ]
    for index, data in enumerate(pages, 1):
        title, goal, page_type, density, rhythm, posture, anchor = data
        outline_lines.extend([
            f"### 第 {index} 页: {title}",
            f"页目标: {goal}",
            f"页面类型映射: {page_type}",
            "密度下限: mid_low",
            f"密度目标: {density}",
            "密度上限: high",
            f"节奏动作: {rhythm}",
            f"信息姿态: {posture}",
            f"锚点类型: {anchor}",
            "",
        ])
    outline_lines.append("自审通过: Sechs Originalseiten sind in derselben Reihenfolge und ohne Inhaltskürzung abgebildet.")
    base.write_text(output_dir / "outline.txt", "\n".join(outline_lines))

    style = {
        "style_id": "dark_tech",
        "style_name": "暗黑科技 (Dark Tech)",
        "mood_keywords": ["深空冷寂", "精密仪器", "微光脉搏", "数据洪流", "未来感"],
        "design_soul": "Der C5 wird als präzises industrielles System in einem tiefblauen Kontrollraum inszeniert; Cyan-Linien verbinden Produkt, Daten, Menschen und Service.",
        "variation_strategy": "Hero, Systemkarten, technische Matrix, ROI-Konsole, People-Raster und Prozess-Timeline wechseln sich ab; Hintergrund, Logo-Position, Cyan-Akzent und Footer bleiben stabil.",
        "decoration_dna": {
            "signature_move": "深空网格 + 冷青扫描线 + L 形角标 + 半透明数据面板",
            "forbidden": ["linker Seitenkopf", "helle Papierfläche", "große Pillen", "inhaltlose Dekoration", "Textkürzung"],
            "recommended_combos": ["产品大图 + 数据轨道", "图标矩阵 + 青色刻度", "人物照片 + 细线档案卡"],
        },
        "font_family": "'Noto Sans', 'Liberation Sans', Arial, sans-serif",
        "css_variables": {
            "bg_primary": "#050b1f",
            "bg_secondary": "#0a1f3d",
            "card_bg_from": "rgba(34,211,238,0.10)",
            "card_bg_to": "rgba(99,102,241,0.06)",
            "card_border": "rgba(34,211,238,0.24)",
            "card_radius": "8px",
            "text_primary": "#FFFFFF",
            "text_secondary": "rgba(255,255,255,0.70)",
            "accent_1": "#22D3EE",
            "accent_2": "#3B82F6",
            "accent_3": "#6366F1",
            "accent_4": "#FDE047",
        },
        "css_snippets": {
            "panel": "background:linear-gradient(145deg,rgba(34,211,238,.10),rgba(99,102,241,.05));border:1px solid rgba(34,211,238,.24);border-radius:8px;",
            "label": "font-family:'Noto Sans Mono',monospace;letter-spacing:.15em;text-transform:uppercase;color:#22D3EE;",
        },
    }
    base.write_json(output_dir / "style.json", style)


def image_contract(filename: str | None = None, placement: str = "inline") -> dict:
    if filename is None:
        return base.image_contract()
    return {
        "mode": "provided",
        "needed": True,
        "usage": "inline-illustration",
        "placement": placement,
        "content_description": "Original visual from the supplied FMC3 brochure.",
        "source_hint": f"images/{filename}",
        "decorate_brief": "Use without changing semantic content; preserve full subject and labels.",
    }


def planning_payloads() -> list[dict]:
    specs = [
        ("cover", "cover", PAGE_TEXT[1][1], "Den C5 als gewerblichen Flächenroboter positionieren.", "image1.jpeg", "image_hero", "hero-top", "铺垫"),
        ("content", "framework", PAGE_TEXT[2][1], "Vier Reinigungssysteme und vier Betriebskennzahlen erklären.", "image4.png", "diagram", "mixed-grid", "推进"),
        ("content", "evidence", PAGE_TEXT[3][1], "Technische Daten und Einsatzfelder vollständig zeigen.", "image13.jpeg", "data", "primary-secondary", "推进"),
        ("content", "evidence", PAGE_TEXT[4][1], "ROI und Berechnungsbasis nachvollziehbar belegen.", "image22.png", "data_highlight", "mixed-grid", "爆发"),
        ("content", "evidence", PAGE_TEXT[5][1], "FMC3 und das fünfköpfige Kernteam als Vertrauensbasis darstellen.", "image27.png", "people", "symmetric", "推进"),
        ("content", "cta", PAGE_TEXT[6][1], "Projektablauf und Serviceversprechen in eine Live-Demo-Einladung überführen.", None, "timeline", "l-shape", "收束"),
    ]
    payloads = []
    for page_number, spec in enumerate(specs, 1):
        page_type, narrative, title, goal, image, card_type, layout, rhythm = spec
        density_label = (
            "mid_low" if page_number == 1
            else "high" if page_number in {3, 4}
            else "medium"
        )
        density_defaults = {
            "mid_low": {
                "max_cards": 3,
                "max_charts": 1,
                "min_body_font_px": 20,
                "max_lines_per_card": 4,
                "image_policy": "flexible",
                "decoration_budget": "medium",
                "overflow_strategy": "rebalance_layout",
            },
            "medium": {
                "max_cards": 4,
                "max_charts": 2,
                "min_body_font_px": 18,
                "max_lines_per_card": 5,
                "image_policy": "support_only",
                "decoration_budget": "medium",
                "overflow_strategy": "tighten_budget",
            },
            "high": {
                "max_cards": 6,
                "max_charts": 2,
                "min_body_font_px": 16,
                "max_lines_per_card": 4,
                "image_policy": "support_only",
                "decoration_budget": "low",
                "overflow_strategy": "table_or_microchart",
            },
        }[density_label]
        page = {
            "slide_number": page_number,
            "page_type": page_type,
            "narrative_role": narrative,
            "title": title,
            "page_goal": goal,
            "audience_takeaway": goal,
            "visual_weight": 8 if page_number in {3, 4, 5} else 7,
            "density_label": density_label,
            "density_reason": "Alle Originalinhalte der Quellseite werden in einer scanbaren Dark-Tech-Struktur erhalten.",
            "density_contract": {
                "deck_bias": "balanced",
                "page_lower_bound": "low" if page_number == 1 else "mid_low",
                "page_upper_bound": "high",
                **density_defaults,
            },
            "layout_hint": layout,
            "layout_variation_note": "Jede Seite nutzt eine eigene, dem Inhalt entsprechende Komposition.",
            "focus_zone": "Titel, Primärbild oder Primärkennzahl im oberen bis mittleren Seitenbereich",
            "negative_space_target": "medium",
            "page_text_strategy": "Deutsche Originaltexte unverändert und modular lesbar.",
            "rhythm_action": rhythm,
            "must_avoid": ["Textkürzung", "linker Seitenkopf", "Bildinhalt beschneiden", "erfundene Daten"],
            "variation_guardrails": {
                "same_gene_as_deck": "Dunkelblauer Grund, Cyan-Akzent, weißes FMC3-Logo rechts, technische Fußzeile.",
                "different_from_previous": ["Inhaltsspezifische Rasterung und wechselnde Ankerbehandlung."],
            },
            "director_command": {
                "mood": "präzise industrielle Zukunft",
                "spatial_strategy": f"{layout} mit klarer Scan-Reihenfolge",
                "anchor_treatment": "Originalbild oder Hauptkennzahl erhält die stärkste Cyan-Kontur.",
                "techniques": ["grid", "scan-line", "cyan-rule"],
                "prose": "Die Seite wirkt wie ein präzises Instrumentenpanel und bleibt trotz hoher Informationsdichte ruhig.",
            },
            "decoration_hints": {
                "background": {"feel": "tiefer technischer Raum", "restraint": "niedrige Deckkraft", "techniques": ["grid"]},
                "floating": {"feel": "Instrumentenmarker", "restraint": "maximal zwei", "techniques": ["corner-lines"]},
                "page_accent": {"feel": "kaltes Cyan", "restraint": "nur Primärwerte", "techniques": ["cyan-rule"]},
            },
            "resources": {
                "page_template": "cover" if page_type == "cover" else None,
                "layout_refs": [] if page_type == "cover" else [layout],
                "block_refs": [card_type] if card_type in {"people", "timeline"} else [],
                "chart_refs": [],
                "principle_refs": ["visual-hierarchy", "composition", "cognitive-load"],
                "resource_rationale": "Das gewählte Raster entspricht dem Inhaltstyp und bewahrt die Lesbarkeit der Originalinhalte.",
            },
            "cards": [
                base.card(
                    f"s{page_number:02d}-anchor-1",
                    "anchor",
                    card_type,
                    "accent",
                    "evidence",
                    title,
                    PAGE_TEXT[page_number][1:9],
                    image_contract(image, "inline"),
                    density_defaults["max_lines_per_card"],
                ),
                base.card(
                    f"s{page_number:02d}-support-1",
                    "support",
                    "text",
                    "outline",
                    "framework",
                    PAGE_TEXT[page_number][0],
                    PAGE_TEXT[page_number][9:],
                    image_contract(),
                    density_defaults["max_lines_per_card"],
                ),
            ],
            "source_guidance": {
                "brief_sections": [f"source-brief.txt#seite-{page_number}"],
                "citation_expectation": "Nur Inhalte und Originalmedien der bereitgestellten DOCX.",
                "strictness": "Keine Ergänzungen oder Umformulierungen.",
            },
            "workflow_metadata": base.WORKFLOW_METADATA,
        }
        payloads.append({"page": page})
    return payloads


def common_css() -> str:
    return f"""
@page {{ size:A4 portrait; margin:0; }}
:root {{
  --bg:#050b1f; --bg2:#0a1f3d; --cyan:#22d3ee; --blue:#3b82f6;
  --violet:#6366f1; --white:#fff; --muted:rgba(255,255,255,.70);
  --faint:rgba(255,255,255,.48); --line:rgba(34,211,238,.25);
}}
*{{box-sizing:border-box}}
html,body{{width:{PAGE_W}px;height:{PAGE_H}px;margin:0;overflow:hidden}}
body{{position:relative;color:var(--white);font-family:'Noto Sans','Liberation Sans',Arial,sans-serif;
background:radial-gradient(circle at 86% 12%,rgba(99,102,241,.24),transparent 29%),
radial-gradient(circle at 8% 78%,rgba(34,211,238,.13),transparent 31%),
linear-gradient(180deg,var(--bg2),var(--bg) 43%,#030817)}}
.grid{{position:absolute;inset:0;opacity:.28;background-image:linear-gradient(rgba(255,255,255,.025) 1px,transparent 1px),linear-gradient(90deg,rgba(255,255,255,.025) 1px,transparent 1px);background-size:64px 64px}}
.scan{{position:absolute;left:0;right:0;top:164px;height:1px;background:linear-gradient(90deg,transparent,var(--cyan),transparent);opacity:.65}}
.corner{{position:absolute;width:30px;height:30px;opacity:.7}} .tl{{top:24px;left:24px;border-top:1px solid var(--cyan);border-left:1px solid var(--cyan)}} .br{{right:24px;bottom:24px;border-right:1px solid var(--cyan);border-bottom:1px solid var(--cyan)}}
.page{{position:relative;z-index:2;height:100%;padding:48px 66px 54px}}
.brand{{height:58px;display:flex;justify-content:flex-end;align-items:center;margin-bottom:22px}} .brand img{{width:150px;height:auto}}
.eyebrow{{display:flex;align-items:center;gap:10px;color:var(--cyan);font-size:13px;font-weight:700;letter-spacing:.15em;text-transform:uppercase;margin-bottom:10px}}
.page > .eyebrow{{margin-bottom:26px}}
.pulse{{width:7px;height:7px;border-radius:50%;background:var(--cyan);box-shadow:0 0 14px rgba(34,211,238,.9)}}
h1,h2,h3,h4,p{{margin:0}} h1{{font-size:47px;line-height:1.06;font-weight:800}} h2{{font-size:27px;line-height:1.16}} h3{{font-size:18px;line-height:1.25}} h4{{font-size:14px;line-height:1.3}}
.subhead{{margin-top:12px;color:var(--muted);font-size:19px;line-height:1.48}} .body{{color:var(--muted);font-size:17px;line-height:1.55}}
.panel{{background:linear-gradient(145deg,rgba(34,211,238,.095),rgba(99,102,241,.05));border:1px solid var(--line);border-radius:8px;box-shadow:0 18px 50px rgba(0,0,0,.2)}}
.image-frame{{background:rgba(255,255,255,.97);border:1px solid rgba(34,211,238,.4);border-radius:8px;overflow:hidden;box-shadow:0 0 34px rgba(34,211,238,.10)}} .image-frame img{{display:block;width:100%;height:100%;object-fit:contain}}
.section-rule{{display:flex;align-items:center;gap:14px}} .section-rule .line{{height:1px;flex:1;background:linear-gradient(90deg,var(--cyan),transparent);opacity:.65}}
.icon{{width:42px;height:42px;object-fit:contain;filter:brightness(0) saturate(100%) invert(79%) sepia(86%) saturate(1032%) hue-rotate(139deg) brightness(101%) contrast(89%)}}
.footer{{position:absolute;left:66px;right:66px;bottom:27px;display:flex;justify-content:space-between;color:rgba(255,255,255,.42);font-size:10px;letter-spacing:.05em}} .footer:before{{content:'';position:absolute;left:0;right:0;top:-12px;height:1px;background:linear-gradient(90deg,var(--cyan),rgba(99,102,241,.35),transparent)}} .page-no{{color:var(--cyan);font-family:'Noto Sans Mono','Liberation Mono',monospace}}
"""


def shell(title: str, page_number: int, content: str) -> str:
    footer = "FMC3 Robotics · C5 Autonomous Commercial Cleaning Robot" if page_number == 6 else FOOTER_TEXT
    return f"""<!doctype html>
<html lang="de"><head><meta charset="utf-8"><meta name="viewport" content="width={PAGE_W},initial-scale=1">
<title>{html.escape(title)}</title><style>{common_css()}</style></head>
<body><div class="grid"></div><div class="scan"></div><div class="corner tl"></div><div class="corner br"></div>
<main class="page"><header class="brand"><img src="../images/image3.png" alt="FMC3 Robotics"></header>
{content}
<footer class="footer"><span>{footer}</span><span class="page-no">0{page_number} / 0{TOTAL_PAGES}</span></footer>
</main></body></html>"""


def page_html() -> dict[int, str]:
    p1 = """
<section class="cover">
  <div class="copy">
    <div class="eyebrow"><span class="pulse"></span>COMMERCIAL CLEANING ROBOT</div>
    <div class="model">C5</div>
    <h1>Große Flächen<br>Weniger Aufwand</h1>
    <p class="subhead">Intelligenter gewerblicher Reinigungsroboter für mittlere und große Flächen.</p>
    <div class="modes"><strong>Kehren</strong><i></i><strong>Scheuern</strong><i></i><strong>Saugen</strong></div>
  </div>
  <div class="hero image-frame"><img src="../runtime/image1-display.png" alt="FMC3 C5 Reinigungsroboter"></div>
</section>
<section class="cover-kpis panel">
  <div><strong>1.980 m²/h</strong><span>Theoretische Flächenleistung</span></div>
  <div><strong>550 mm</strong><span>Schrubbreite</span></div>
  <div><strong>3 h</strong><span>Reinigungsdauer</span></div>
</section>
<section class="cover-company"><strong>FMC3 Robotics</strong><span>Ingolstadt · Deutschland</span><span>www.FMC3-robotics.ai</span></section>
<style>
.cover{display:grid;grid-template-columns:.82fr 1.18fr;gap:25px;align-items:stretch}.copy{padding-top:85px}.model{font-size:142px;line-height:.9;font-weight:800;color:var(--white);margin:26px 0 35px}.copy h1{font-size:43px}.copy .subhead{max-width:360px;font-size:20px}.modes{display:flex;align-items:center;gap:11px;margin-top:26px;font-size:16px}.modes i{width:5px;height:5px;border-radius:50%;background:var(--cyan)}
.hero{height:1060px;background:#f7f7f7;padding:18px}.hero img{object-fit:contain;object-position:center}
.cover-kpis{display:grid;grid-template-columns:repeat(3,1fr);margin-top:28px;height:132px}.cover-kpis div{padding:24px 30px;border-right:1px solid rgba(255,255,255,.14)}.cover-kpis div:last-child{border-right:0}.cover-kpis strong{display:block;font-size:31px}.cover-kpis span{display:block;margin-top:10px;color:var(--muted);font-size:14px}
.cover-company{margin-top:24px;display:flex;gap:22px;align-items:center;color:var(--faint);font-size:14px}.cover-company strong{color:var(--white);font-size:16px}
</style>"""

    systems = [
        ("image4.png", "Reinigungsmechanik", ["Doppelwalzenbürsten entfernen auch schwere Verschmutzungen.", "Zwei-Kammer-Saugfuß reduziert Restwasser in Fugen."]),
        ("image5.png", "Navigation & Anpassung", ["Laser- und visuelle Lokalisierung für unterschiedliche Szenen.", "Randbürste arbeitet dicht an Kanten im geschlossenen Kreislauf."]),
        ("image6.png", "Autonomer Betrieb", ["Autonomes Laden sowie Frisch- und Schmutzwasser-Management.", "Tür- und Aufzuganbindung für durchgängige Abläufe."]),
        ("image7.png", "Wartung & Lebensdauer", ["Selbstreinigender Schmutzwassertank reduziert täglichen Aufwand.", "Gebläselebensdauer 10.000 h und langlebige Komponenten."]),
    ]
    system_cards = "".join(
        f"""<article class="system panel"><div class="system-head"><img class="icon" src="../images/{icon}"><h3>{title}</h3></div>
<ul><li>{items[0]}</li><li>{items[1]}</li></ul></article>"""
        for icon, title, items in systems
    )
    p2 = f"""
<div class="eyebrow"><span class="pulse"></span>LEISTUNG IM ALLTAG</div>
<h1>Vier Systeme. Ein sauberer Prozess.</h1>
<p class="subhead">Der C5 verbindet gründliche Bodenreinigung, sichere Navigation, autonome Versorgung und wartungsarme Technik.</p>
<blockquote class="panel">Kehren und Scheuern in einem Arbeitsgang –<br>mit hoher Wasseraufnahme und konsequenter Kantenreinigung.</blockquote>
<section class="systems">{system_cards}</section>
<section class="metrics panel">
  <div><strong>95 %</strong><span>Schmutzaufnahme</span></div><div><strong>&lt;68 dB(A)</strong><span>Geräuschpegel</span></div>
  <div><strong>25 kg</strong><span>Bodendruck</span></div><div><strong>4 min</strong><span>Tank-Selbstreinigung</span></div>
</section>
<style>
.subhead{{max-width:900px}} blockquote{{margin:26px 0 24px;padding:22px 28px 22px 78px;position:relative;color:var(--white);font-size:18px;line-height:1.45;font-weight:700}} blockquote:before{{content:'“';position:absolute;left:27px;top:6px;font-size:54px;color:var(--cyan)}}
.systems{{display:grid;grid-template-columns:1fr 1fr;gap:18px}} .system{{padding:29px;min-height:330px}} .system-head{{display:flex;align-items:center;gap:18px}} .system ul{{padding:0;margin:25px 0 0;list-style:none}} .system li{{position:relative;padding-left:20px;color:var(--muted);font-size:17px;line-height:1.5;margin-top:22px}} .system li:before{{content:'•';position:absolute;left:0;color:var(--cyan)}}
.metrics{{display:grid;grid-template-columns:repeat(4,1fr);margin-top:24px;height:142px}} .metrics div{{padding:27px 22px;border-right:1px solid rgba(255,255,255,.14)}} .metrics div:last-child{{border:0}} .metrics strong{{display:block;font-size:30px}} .metrics span{{display:block;margin-top:11px;color:var(--muted);font-size:14px}}
</style>"""

    specs = [
        ("image8.png", "90 L", "Großer Wassertank"),
        ("image9.png", "24 V / 80 Ah", "Batterie"),
        ("image10.png", "170 kg", "Gesamtgewicht"),
        ("image11.png", "1.800 W", "Max. Leistung"),
        ("image12.png", "200–240 V", "Nenneingangsspannung"),
        ("image8.png", "8 L", "Tankvolumen Arbeitsstation"),
    ]
    spec_cells = "".join(
        f'<div class="spec"><img class="icon" src="../images/{icon}"><strong>{value}</strong><span>{label}</span></div>'
        for icon, value, label in specs
    )
    scenarios = [
        ("image17.png", "Große Supermärkte"), ("image18.png", "Bürogebäude"), ("image19.png", "Verkehrsknoten"),
        ("image20.png", "Fabriken"), ("image18.png", "Gewerbekomplexe"), ("image21.png", "Krankenhäuser"),
    ]
    scenario_cells = "".join(
        f'<div class="scenario panel"><img src="../images/{icon}"><strong>{label}</strong></div>'
        for icon, label in scenarios
    )
    p3 = f"""
<div class="eyebrow"><span class="pulse"></span>TECHNISCHE DATEN</div>
<h1>Kompakt gebaut. Für große Aufgaben.</h1>
<p class="subhead">Leistungsdaten, Abmessungen und typische Einsatzbereiche des C5 im Überblick.</p>
<section class="tech-main">
  <div class="parameter panel"><h3>C5 · PRODUKTPARAMETER</h3><div class="spec-grid">{spec_cells}</div><div class="flow">Frischwasser 7–10 L/min · Ablass 10–15 L/min</div></div>
  <div class="machine image-frame"><img src="../images/image13.jpeg" alt="C5 mit Arbeitsstation"><div class="dimensions"><img src="../images/image15.png"><div><strong>680 × 820 × 1.085 mm</strong><span>Arbeitsstation: 520 × 310 × 1.105 mm</span></div><img src="../images/image16.png"></div></div>
</section>
<div class="section-rule scenarios-title"><h3>EINSATZSZENARIEN</h3><div class="line"></div></div>
<section class="scenario-grid">{scenario_cells}</section>
<div class="section-rule floor-title"><h3>GEEIGNETE BODENBELÄGE</h3><div class="line"></div></div>
<section class="floors"><span>Zementboden</span><span>Marmor</span><span>Epoxidboden</span><span>PVC</span><span>Korundboden</span></section>
<style>
.tech-main{{display:grid;grid-template-columns:.94fr 1.06fr;gap:18px;margin-top:24px}} .parameter{{height:735px;padding:30px}} .parameter h3{{color:var(--cyan);letter-spacing:.06em}} .spec-grid{{display:grid;grid-template-columns:1fr 1fr;gap:32px 20px;margin-top:35px}} .spec{{min-height:134px}} .spec .icon{{display:block;width:35px;height:35px;margin-bottom:10px}} .spec strong{{display:block;font-size:23px}} .spec span{{display:block;margin-top:7px;color:var(--muted);font-size:14px;line-height:1.4}} .flow{{margin-top:25px;padding:16px;background:rgba(255,255,255,.07);border-radius:6px;color:var(--muted);font-size:14px;text-align:center}}
.machine{{height:735px;padding:0;position:relative}} .machine>img{{height:563px;object-fit:cover}} .dimensions{{height:172px;padding:20px;display:grid;grid-template-columns:40px 1fr 36px;gap:13px;align-items:center;background:rgba(4,12,28,.98)}} .dimensions img{{width:34px;filter:brightness(0) saturate(100%) invert(79%) sepia(86%) saturate(1032%) hue-rotate(139deg) brightness(101%) contrast(89%)}} .dimensions strong{{display:block;font-size:22px}} .dimensions span{{display:block;margin-top:14px;color:var(--muted);font-size:15px}}
.scenarios-title{{margin-top:22px}} .scenarios-title h3,.floor-title h3{{color:var(--cyan);font-size:14px;letter-spacing:.05em}} .scenario-grid{{display:grid;grid-template-columns:repeat(3,1fr);gap:10px 14px;margin-top:13px}} .scenario{{height:68px;padding:12px 17px;display:flex;align-items:center;gap:12px}} .scenario img{{width:32px;height:32px;object-fit:contain;filter:brightness(0) saturate(100%) invert(79%) sepia(86%) saturate(1032%) hue-rotate(139deg) brightness(101%) contrast(89%)}} .scenario strong{{font-size:15px}}
.floor-title{{margin-top:18px}} .floors{{display:grid;grid-template-columns:repeat(5,1fr);gap:10px;margin-top:12px}} .floors span{{padding:11px 5px;text-align:center;border:1px solid var(--line);border-radius:7px;color:var(--muted);font-size:13px}}
</style>"""

    p4 = """
<div class="eyebrow"><span class="pulse"></span>WIRTSCHAFTLICHKEIT</div>
<h1>Ein Business Case,<br>der sich rechnen lässt.</h1>
<p class="subhead">Beispielrechnung gegenüber manueller Reinigung – berechnet über 60 Monate und 1.000 Betriebsstunden pro Jahr.</p>
<section class="roi-kpis">
  <div class="roi-kpi panel"><img class="icon" src="../images/image22.png"><strong>9.089 €</strong><span>GESCHÄTZTE EINSPARUNGEN<br>PRO JAHR</span></div>
  <div class="roi-kpi panel"><img class="icon" src="../images/image23.png"><strong>1,3 Jahre</strong><span>AMORTISATIONSZEIT / ROI</span></div>
  <div class="roi-kpi panel"><img class="icon" src="../images/image24.png"><strong>1,54 Mio. m²</strong><span>GEREINIGTE FLÄCHE PRO JAHR</span></div>
</section>
<section class="comparison panel"><h3>JÄHRLICHER VERGLEICH</h3><div class="bar-row"><span>Aktuelle Lösung</span><i class="bar old"></i><strong>20.780 €</strong></div><div class="bar-row"><span>C5 · 60 Monate</span><i class="bar c5"></i><strong>11.691 €</strong></div></section>
<section class="calculator panel">
  <h3>ROI- UND EINSPARUNGSRECHNER</h3>
  <div class="calc-grid">
    <div class="inputs">
      <div class="input-card"><div class="input-title"><img class="icon" src="../images/image25.png"><div><strong>Reinigungsleistung</strong><small>Von der Reinigungsfläche berechnete Kosten</small></div></div>
        <div class="slider"><b>Verrechneter Stundenlohn</b><strong>24 €/h</strong><i style="--p:35%"></i><small><span>15 €/h</span><span>50 €/h</span></small></div>
        <div class="slider"><b>Reinigungsstunden pro Tag</b><strong>3 h/Tag</strong><i style="--p:28%"></i><small><span>1 h/Tag</span><span>12 h/Tag</span></small></div>
        <div class="slider"><b>Reinigungstage pro Jahr</b><strong>260 Tage/Jahr</strong><i style="--p:58%"></i><small><span>100 Tage/Jahr</span><span>365 Tage/Jahr</span></small></div>
      </div>
      <div class="machine-cost"><img class="icon" src="../images/image26.png"><div><strong>Manuelle Scheuersaugmaschine</strong><p>Anschaffungskosten <b>8.500 €</b></p><p>Wartung <b>30 €/Monat</b></p></div></div>
    </div>
    <div class="results">
      <div class="saving"><span>GESCHÄTZTE JÄHRLICHE EINSPARUNGEN</span><strong>9.089 €</strong><small>entspricht 757 € / Monat</small></div>
      <div class="mini-compare"><h4>JÄHRLICHER VERGLEICH</h4><p>Aktuelle Lösung <b>20.780 €</b></p><i class="old"></i><p>C5 · 60 Monate <b>11.691 €</b></p><i class="c5"></i></div>
      <div class="result-pair"><div><span>AMORTISATIONSZEIT / ROI</span><strong>1,3 Jahre</strong></div><div><span>FLÄCHE / JAHR</span><strong>1.544.400 m²</strong></div></div>
      <div class="offer">Individuelles Angebot erhalten</div>
    </div>
  </div>
</section>
<section class="basis panel"><strong>ⓘ &nbsp; Berechnungsbasis</strong><span>24 €/h · 3 h/Tag · 260 Tage/Jahr · 8.500 € Maschine · 30 €/Monat Wartung</span><small>Kundennahe Beispielrechnung; Werte können projektspezifisch angepasst werden.</small></section>
<style>
h1{font-size:45px}.subhead{font-size:18px}.roi-kpis{display:grid;grid-template-columns:repeat(3,1fr);gap:16px;margin-top:22px}.roi-kpi{height:190px;padding:20px;text-align:center}.roi-kpi .icon{width:35px;height:35px}.roi-kpi strong{display:block;font-size:31px;margin-top:12px}.roi-kpi span{display:block;color:var(--muted);font-size:13px;font-weight:700;line-height:1.4;margin-top:10px}
.comparison{margin-top:16px;padding:20px 24px}.comparison h3,.calculator>h3{color:var(--cyan);font-size:17px}.bar-row{display:grid;grid-template-columns:155px 1fr 100px;gap:15px;align-items:center;margin-top:16px;font-size:15px}.bar{height:12px;border-radius:5px;background:#8b96a8}.bar.c5{width:61%;background:var(--cyan)}.bar-row strong{text-align:right;font-size:18px}
.calculator{margin-top:16px;padding:22px}.calc-grid{display:grid;grid-template-columns:.9fr 1.1fr;gap:18px;margin-top:16px}.input-card,.machine-cost,.saving,.mini-compare,.result-pair>div{border:1px solid rgba(255,255,255,.14);border-radius:7px}.input-card{padding:16px}.input-title{display:flex;align-items:center;gap:11px}.input-title .icon{width:32px;height:32px}.input-title strong{font-size:16px}.input-title small{display:block;color:var(--muted);font-size:12px;line-height:1.35}.slider{position:relative;margin-top:15px}.slider b,.slider strong{font-size:12px}.slider strong{float:right;color:var(--cyan)}.slider i{display:block;position:relative;height:4px;background:rgba(255,255,255,.25);margin-top:8px}.slider i:before{content:'';display:block;width:var(--p);height:4px;background:var(--cyan)}.slider small{display:flex;justify-content:space-between;color:var(--faint);font-size:11px;margin-top:6px}.machine-cost{display:flex;gap:12px;padding:15px;margin-top:12px}.machine-cost .icon{width:31px;height:31px}.machine-cost strong{font-size:14px}.machine-cost p{font-size:12px;margin-top:9px;color:var(--muted)}.machine-cost b{float:right;color:var(--cyan)}
.saving{text-align:center;padding:19px;background:rgba(34,211,238,.08)}.saving span,.result-pair span{display:block;color:var(--muted);font-size:12px;font-weight:700}.saving strong{display:block;color:var(--cyan);font-size:39px;margin-top:9px}.saving small{font-size:13px}.mini-compare{padding:15px;margin-top:12px}.mini-compare h4{font-size:12px;color:var(--muted)}.mini-compare p{font-size:12px;margin-top:9px}.mini-compare p b{float:right}.mini-compare i{display:block;height:5px;margin-top:5px;background:#8b96a8;border-radius:3px}.mini-compare i.c5{width:61%;background:var(--cyan)}.result-pair{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-top:12px}.result-pair>div{text-align:center;padding:13px}.result-pair strong{display:block;margin-top:7px;font-size:18px}.offer{margin-top:12px;background:var(--cyan);color:#04111e;text-align:center;padding:12px;border-radius:6px;font-size:14px;font-weight:800}
.basis{margin-top:16px;padding:17px 22px;display:grid;grid-template-columns:210px 1fr;gap:9px 16px;align-items:center}.basis strong{font-size:14px}.basis span{color:var(--muted);font-size:12px}.basis small{grid-column:2;color:var(--faint);font-size:11px}
</style>"""

    team = [
        ("image27.png", "Dr. Wang Cheng", "CEO", "Dr.-Ing. der Technischen Universität München (TUM). Experte für autonomes Fahren. Architekt von Audis Intelligent Driving System. Track Record in Produktion und Auslieferung von über 1 Mio. intelligenten Fahrzeugen."),
        ("image28.png", "Prof. Dr.<br>Björn Giesler", "CTO", "PhD in Informatik am KIT; anerkannte Autorität in KI-Architektur. Über 50 Core-Patente zu Sensing und Datenverarbeitung. Führung von F&E-Teams mit mehr als 500 Mitgliedern."),
        ("image29.png", "Fan David, PhD", "COO", "Ehemaliger Präsident von Marelli China. Verantwortlich für Geschäftsbereiche mit über 1,5 Mrd. €. Post-Merger-Integrationen und Supply-Chain-Restrukturierungen."),
        ("image30.png", "Dr. Huang Dong", "CPO", "Dr.-Ing. in Informatik vom KIT. Ehemaliger Scientist bei Siemens Labs und CTO von TerraIT. Über 15 Jahre Expertise in Datensimulation und Sim-to-Real-Domänen."),
        ("image31.png", "Florian Obermeier", "MD Germany", "Mitgründer der SANEON GmbH. Über 16 Jahre B2B-Erfahrung mit OEMs und Tier-1-Automobilzulieferern. Markterschließung in Europa, Asien und den USA."),
    ]
    team_cards = "".join(
        f'<article class="person panel"><img src="../images/{photo}"><h3>{name}</h3><span>{role}</span><p>{bio}</p></article>'
        for photo, name, role, bio in team
    )
    p5 = f"""
<div class="eyebrow"><span class="pulse"></span>FMC3 ROBOTICS</div>
<h1>Wissenschaft. Industrie. Umsetzung.</h1>
<p class="subhead">Europäische Wissenschaftler für intelligente Systeme – mit Erfahrung aus Automobilindustrie, KI-Forschung und internationalem B2B-Geschäft.</p>
<section class="company-grid">
  <div class="about panel"><h2>Über FMC3 Robotics</h2><p>FMC3 bringt moderne Robotiktechnologie vom Standort Ingolstadt in den europäischen Markt. Das Unternehmen wurde im Oktober 2025 von Prof. Dr. Björn Giesler, Dr. Wang Cheng und Dr. David Fan gegründet. Das Portfolio umfasst Reinigungs-, Begrüßungs- und Lieferroboter.</p><p>Wissenschaftliche Exzellenz, industrielle Erfahrung und ein internationales Netzwerk bilden die Grundlage für zuverlässige und skalierbare Lösungen im professionellen Einsatz.</p></div>
  <div class="facts panel"><div><strong>Okt. 2025</strong><span>GEGRÜNDET</span></div><div><strong>Ingolstadt</strong><span>HAUPTSITZ</span></div><div><strong>DE &amp;<br>Shanghai</strong><span>ENTWICKLUNG</span></div><div><strong>30 / Monat</strong><span>LIEFERKAPAZITÄT</span></div></div>
</section>
<div class="section-rule team-title"><h3>DAS KERNTEAM</h3><div class="line"></div></div>
<section class="team-grid">{team_cards}</section>
<section class="local panel"><strong>LOKALE BETREUUNG · INTERNATIONALE ENTWICKLUNG</strong><span>Persönliche Betreuung in Deutschland · Technische Entwicklung in Ingolstadt und Shanghai</span><div><b>Reinigungsroboter</b><b>Begrüßungsroboter</b><b>Lieferroboter</b></div></section>
<style>
h1{{font-size:43px}} .subhead{{font-size:18px}} .company-grid{{display:grid;grid-template-columns:1.65fr 1fr;gap:16px;margin-top:22px}} .about{{height:282px;padding:25px}} .about h2{{font-size:23px}} .about p{{color:var(--muted);font-size:15px;line-height:1.5;margin-top:14px}} .facts{{height:282px;display:grid;grid-template-columns:1fr 1fr}} .facts div{{padding:23px;border-right:1px solid rgba(255,255,255,.13);border-bottom:1px solid rgba(255,255,255,.13)}} .facts div:nth-child(2n){{border-right:0}} .facts div:nth-child(n+3){{border-bottom:0}} .facts strong{{display:block;font-size:20px}} .facts span{{display:block;color:var(--muted);font-size:11px;font-weight:700;margin-top:15px}}
.team-title{{margin-top:20px}} .team-title h3{{color:var(--cyan);font-size:14px}} .team-grid{{display:grid;grid-template-columns:repeat(5,1fr);gap:10px;margin-top:12px}} .person{{height:640px;padding:17px 13px;text-align:center}} .person img{{width:116px;height:116px;border-radius:50%;object-fit:cover;border:2px solid var(--cyan);background:#fff}} .person h3{{font-size:14px;min-height:38px;margin-top:11px;display:flex;align-items:center;justify-content:center}} .person>span{{display:block;margin:11px auto 0;padding:8px 4px;border-radius:6px;background:rgba(255,255,255,.08);color:var(--cyan);font-size:12px;font-weight:800}} .person p{{margin-top:14px;text-align:left;color:var(--muted);font-size:13px;line-height:1.5}}
.local{{margin-top:18px;padding:19px 20px;text-align:center;background:rgba(34,211,238,.11)}} .local strong{{font-size:14px}} .local span{{display:block;color:var(--muted);font-size:13px;margin-top:9px}} .local div{{display:grid;grid-template-columns:repeat(3,1fr);margin-top:11px;font-size:12px}}
</style>"""

    steps = [
        ("image32.png", "1", "Analyse & Bedarfsprüfung", "Anforderungen verstehen und Gegebenheiten vor Ort prüfen."),
        ("image33.png", "2", "Live-Demonstration / Test", "C5 bei Ihnen vor Ort oder am Hauptsitz in Ingolstadt erleben."),
        ("image34.png", "3", "Lieferung & Inbetriebnahme", "Liefern, installieren, testen und gemeinsam in Betrieb nehmen."),
        ("image35.png", "4", "Schulung", "Ihr Team praxisnah und individuell einweisen."),
        ("image36.png", "5", "Support", "Schnelle Reaktion und zuverlässiger After-Sales-Support."),
    ]
    step_html = "".join(
        f'<div class="step"><div class="step-icon"><img src="../images/{icon}"></div><b>{number}</b><div><h3>{title}</h3><p>{desc}</p></div></div>'
        for icon, number, title, desc in steps
    )
    promises = [
        ("image37.png", "Individuelle<br>Video-Schulungen"),
        ("image38.png", "Express-Reaktion<br>innerhalb von 0,5 Stunden"),
        ("image39.png", "Servicenetz in<br>150+ Städten"),
        ("image40.png", "Erste Wartung<br>kostenlos"),
    ]
    promise_html = "".join(
        f'<div class="promise panel"><img class="icon" src="../images/{icon}"><strong>{label}</strong></div>'
        for icon, label in promises
    )
    p6 = f"""
<div class="eyebrow"><span class="pulse"></span>SERVICE &amp; LIVE-DEMONSTRATION</div>
<h1>Vom ersten Test<br>bis zum laufenden Betrieb.</h1>
<p class="subhead">FMC3 begleitet Analyse, Demonstration, Inbetriebnahme, Schulung und After-Sales-Support aus Ingolstadt.</p>
<section class="service-layout">
  <div><div class="section-rule"><h3>PROJEKTABLAUF</h3><div class="line"></div></div><div class="timeline">{step_html}</div></div>
  <div><div class="section-rule"><h3>SERVICEVERSPRECHEN</h3><div class="line"></div></div><div class="service-track">{promise_html}
    <div class="live panel"><img class="icon" src="../images/image41.png"><div><h2>ERLEBEN SIE<br>DEN C5 LIVE</h2><p>Wir senden Ihnen weitere Informationen und laden Sie zu einer Demonstration ein.</p></div></div>
    </div>
  </div>
</section>
<section class="contact panel"><span>Ihr Ansprechpartner</span><strong>Haoan Wang</strong><b>FMC3 Robotics</b><small>Ingolstadt · Deutschland</small></section>
<style>
h1{{font-size:48px}} .subhead{{max-width:860px;font-size:19px}} .service-layout{{display:grid;grid-template-columns:1.08fr .92fr;gap:52px;margin-top:27px}} .section-rule h3{{color:var(--cyan);font-size:14px}} .timeline{{position:relative;margin-top:24px}} .timeline:before{{content:'';position:absolute;left:43px;top:38px;bottom:50px;width:2px;background:linear-gradient(var(--cyan),rgba(99,102,241,.6))}}
.step{{position:relative;display:grid;grid-template-columns:86px 34px 1fr;gap:11px;align-items:start;min-height:142px}} .step-icon{{z-index:2;width:66px;height:66px;border-radius:50%;display:flex;align-items:center;justify-content:center;background:linear-gradient(145deg,var(--cyan),var(--violet));box-shadow:0 0 24px rgba(34,211,238,.2)}} .step-icon img{{width:34px;height:34px;object-fit:contain;filter:brightness(0) saturate(100%)}} .step>b{{color:var(--cyan);font-size:25px;padding-top:18px}} .step h3{{font-size:17px;padding-top:15px}} .step p{{color:var(--muted);font-size:15px;line-height:1.45;margin-top:9px;padding-bottom:16px;border-bottom:1px solid rgba(255,255,255,.13)}}
.service-track{{display:grid;grid-template-rows:repeat(5,1fr);gap:12px;height:710px;margin-top:24px}} .promise{{padding:15px 20px;display:flex;align-items:center;gap:18px}} .promise strong{{font-size:17px;line-height:1.3}} .live{{padding:18px;display:flex;align-items:center;gap:19px;border-color:rgba(34,211,238,.65)}} .live .icon{{width:50px;height:50px}} .live h2{{font-size:23px;color:var(--cyan)}} .live p{{color:var(--muted);font-size:13px;line-height:1.45;margin-top:10px}}
.contact{{position:absolute;left:66px;bottom:92px;width:390px;padding:20px 24px}} .contact span{{display:block;color:var(--cyan);font-size:13px;font-weight:700}} .contact strong{{display:block;font-size:25px;margin-top:13px}} .contact b{{display:block;font-size:15px;margin-top:13px}} .contact small{{display:block;color:var(--muted);font-size:14px;margin-top:9px}}
</style>"""

    return {
        1: shell(PAGE_TEXT[1][1], 1, p1),
        2: shell(PAGE_TEXT[2][1], 2, p2),
        3: shell(PAGE_TEXT[3][1], 3, p3),
        4: shell(PAGE_TEXT[4][1], 4, p4),
        5: shell(PAGE_TEXT[5][1], 5, p5),
        6: shell(PAGE_TEXT[6][1], 6, p6),
    }


def build_preview(output_dir: Path) -> None:
    frames = "".join(
        f'<div class="sheet"><iframe src="slides/slide-{page}.html" title="Seite {page}"></iframe></div>'
        for page in range(1, TOTAL_PAGES + 1)
    )
    base.write_text(
        output_dir / "preview.html",
        f"""<!doctype html><html lang="de"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>FMC3 C5 Dark Tech Broschüre</title><style>
*{{box-sizing:border-box}}body{{margin:0;background:#020611;color:#fff;font-family:Arial,sans-serif}}header{{position:sticky;top:0;z-index:5;height:54px;display:flex;align-items:center;justify-content:space-between;padding:0 24px;background:rgba(5,11,31,.96);border-bottom:1px solid rgba(34,211,238,.25)}}header span{{color:#22d3ee;font-family:monospace}}main{{padding:28px 16px 60px;display:flex;flex-direction:column;align-items:center;gap:34px}}.sheet{{width:min(94vw,794px);aspect-ratio:210/297;overflow:hidden;background:#050b1f;border:1px solid rgba(34,211,238,.2);box-shadow:0 24px 80px rgba(0,0,0,.58)}}iframe{{width:{PAGE_W}px;height:{PAGE_H}px;border:0;transform-origin:top left}}
</style></head><body><header><strong>FMC3 C5 · Dark Tech Broschüre</strong><span>6 × A4</span></header><main>{frames}</main>
<script>function fit(){{document.querySelectorAll('.sheet').forEach(s=>s.querySelector('iframe').style.transform=`scale(${{s.clientWidth/{PAGE_W}}})`)}}addEventListener('resize',fit);fit()</script></body></html>""",
    )


def content_audit(
    output_dir: Path,
    source_media: dict[str, dict[str, str | int]],
) -> dict:
    page_texts = {
        page: base.html_text(output_dir / f"slides/slide-{page}.html")
        for page in range(1, TOTAL_PAGES + 1)
    }
    all_html = normalize_text(" ".join(page_texts.values()))
    source_segments = source_xml_segments(SOURCE)
    missing_source = []
    applied_corrections = []
    for segment in source_segments:
        normalized_segment = normalize_text(segment).rstrip("-")
        if normalized_segment in all_html:
            continue
        corrected = APPROVED_TEXT_CORRECTIONS.get(segment)
        if corrected and normalize_text(corrected) in all_html:
            applied_corrections.append({"source": segment, "corrected": corrected})
            continue
        missing_source.append(segment)
    missing_pages = {
        str(page): [
            segment for segment in segments
            if normalize_text(segment) not in normalize_text(page_texts[page])
        ]
        for page, segments in PAGE_TEXT.items()
    }
    missing_pages = {page: values for page, values in missing_pages.items() if values}

    media_checks = {}
    for filename, metadata in source_media.items():
        path = output_dir / "images" / filename
        output_hash = base.sha256_bytes(path.read_bytes()) if path.exists() else None
        media_checks[filename] = {
            "source_sha256": metadata["sha256"],
            "output_sha256": output_hash,
            "byte_identical": output_hash == metadata["sha256"],
        }
    png_checks = {}
    for page in range(1, TOTAL_PAGES + 1):
        path = output_dir / f"png/slide-{page}.png"
        with Image.open(path) as image:
            png_checks[str(page)] = {
                "dimensions": list(image.size),
                "a4_ratio_ok": abs(image.width / image.height - 210 / 297) < .002,
                "bytes": path.stat().st_size,
            }
    ok = (
        not missing_source
        and not missing_pages
        and all(item["byte_identical"] for item in media_checks.values())
        and all(item["a4_ratio_ok"] and item["bytes"] > 50_000 for item in png_checks.values())
    )
    return {
        "ok": ok,
        "source": str(SOURCE),
        "source_xml_segment_count": len(source_segments),
        "approved_text_corrections": applied_corrections,
        "missing_source_segments": missing_source,
        "missing_page_segments": missing_pages,
        "media_checks": media_checks,
        "png_checks": png_checks,
        "note": "All 42 embedded media files are copied byte-identically; source textbox text is audited directly from document.xml, with explicit user-approved corrections recorded separately.",
    }


def write_manifest(output_dir: Path) -> None:
    manifest = {
        "run_id": output_dir.name,
        "generated_at": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
        "summary": {"total_pages": TOTAL_PAGES, "format": "A4 portrait brochure"},
        "artifacts": {
            "preview_html": "preview.html",
            "presentation_png_pptx": "presentation-png.pptx",
            "presentation_svg_pptx": "presentation-svg.pptx",
            "brochure_pdf": PDF_NAME,
            "content_audit": "content-audit.json",
        },
        "pages": [
            {
                "page": page,
                "planning": f"planning/planning{page}.json",
                "html": f"slides/slide-{page}.html",
                "png": f"png/slide-{page}.png",
            }
            for page in range(1, TOTAL_PAGES + 1)
        ],
        "export_notes": {
            "presentation_png_pptx": "A4 portrait slides with high-resolution page images.",
            "presentation_svg_pptx": "Compatibility fallback using the same A4 raster rendering.",
        },
    }
    base.write_json(output_dir / "delivery-manifest.json", manifest)


def build(output_dir: Path, scale: int) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    for directory in ("images", "planning", "slides", "png", "runtime"):
        (output_dir / directory).mkdir(parents=True, exist_ok=True)

    media = base.extract_media(SOURCE, output_dir / "images")
    prepare_display_assets(output_dir)
    workflow_files(output_dir)
    for page, payload in enumerate(planning_payloads(), 1):
        base.write_json(output_dir / f"planning/planning{page}.json", payload)
    for page, content in page_html().items():
        base.write_text(output_dir / f"slides/slide-{page}.html", content)
    build_preview(output_dir)

    pngs = base.render_pngs(output_dir, scale=scale)
    base.build_pdf(pngs, output_dir / PDF_NAME)
    base.build_pptx(pngs, output_dir / "presentation-png.pptx")
    base.build_pptx(pngs, output_dir / "presentation-svg.pptx")

    audit = content_audit(output_dir, media)
    base.write_json(output_dir / "content-audit.json", audit)
    if not audit["ok"]:
        raise RuntimeError(f"Content audit failed: {output_dir / 'content-audit.json'}")

    write_manifest(output_dir)
    base.run_validators(output_dir)
    print(f"Built dark_tech brochure: {output_dir}")


def main() -> int:
    parser = argparse.ArgumentParser(description="Build the six-page FMC3 C5 vector brochure in dark_tech style")
    parser.add_argument("--output-dir", type=Path, default=DEFAULT_OUTPUT)
    parser.add_argument("--scale", type=int, default=2, choices=(1, 2, 3))
    args = parser.parse_args()
    if not SOURCE.exists():
        print(f"Source DOCX not found: {SOURCE}", file=sys.stderr)
        return 1
    build(args.output_dir.resolve(), args.scale)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
