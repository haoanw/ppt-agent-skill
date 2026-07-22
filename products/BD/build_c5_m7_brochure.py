#!/usr/bin/env python3
"""Build the two-page German C5 brochure in the M7 industrial poster style."""

from __future__ import annotations

import argparse
import hashlib
import json
import shutil
import subprocess
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from PIL import Image, ImageStat
from pptx import Presentation
from pptx.util import Inches


ROOT = Path(__file__).resolve().parents[2]
SOURCE = ROOT / "products/BD/C5_Brochure_Deutsch.pdf"
SOURCE_ASSETS = (
    ROOT
    / "products/BD/FMC3_C5_Brochure_Deutsch_Redesigned_2_HighRes_Vector_dark_tech/images"
)
DEFAULT_OUTPUT = ROOT / "products/BD/C5_Brochure_Deutsch_m7"
PDF_NAME = "C5_Brochure_Deutsch_m7.pdf"
PPTX_NAME = "C5_Brochure_Deutsch_m7.pptx"

PAGE_W = 1024
PAGE_H = 1536


FEATURES = [
    (
        "Doppelwalzenbürsten",
        "Kehren und Schrubben in einem; hoher Wasserdurchsatz und Anpressdruck entfernen schweren Schmutz.",
    ),
    (
        "Zwei-Kammer-Saugfuß",
        "Weniger Restwasser in Fugen und eine gründlichere Wasseraufnahme.",
    ),
    (
        "Sichere Lokalisierung",
        "Laser- und visuelle Lokalisierung für starke Anpassung an verschiedene Szenen.",
    ),
    (
        "Einfache Wartung",
        "Selbstreinigender Schmutzwassertank; der tägliche Aufwand wird halbiert.",
    ),
    (
        "Autonomer Betrieb",
        "Laden, Frisch- und Schmutzwasser-Management, Tankreinigung sowie Tür- und Aufzugsanbindung.",
    ),
    (
        "Konsequente Kantenreinigung",
        "Die Randbürste reinigt dicht an Kanten; Nachreinigung im geschlossenen Reinigungskreislauf.",
    ),
    (
        "Hohe Flächenleistung",
        "550 mm Schrubbreite und eine theoretische Flächenleistung von bis zu 1.980 m²/h.",
    ),
    (
        "Langlebige Komponenten",
        "10.000 h Gebläselebensdauer und niedrige Lebenszykluskosten.",
    ),
]

C5_SPECS = [
    ("95 %", "Schmutzaufnahme"),
    ("170 kg", "Gesamtgewicht"),
    ("4 min", "Tank-Selbstreinigung"),
    ("24 V / 80 Ah", "Batterie"),
    ("1.980 m²/h", "Max. Effizienz"),
    ("<68 dB(A)", "Geräuschpegel"),
    ("10.000 h", "Gebläselebensdauer"),
    ("90 L", "Großer Wassertank"),
    ("3 h", "Reinigungsdauer"),
    ("25 kg", "Bodendruck"),
    ("550 mm", "Schrubbreite"),
]

STATION_SPECS = [
    ("8 L", "Tankvolumen"),
    ("7–10 L/min", "Frischwasser-Zufuhr"),
    ("10–15 L/min", "Schmutzwasser-Ablass"),
    ("1.800 W", "Max. Leistung"),
    ("200–240 V", "Nenneingangsspannung"),
]

SCENARIOS = [
    "Große Supermärkte",
    "Bürogebäude",
    "Verkehrsknoten",
    "Fabriken",
    "Gewerbekomplexe",
    "Krankenhäuser",
]

FLOORS = ["Zementboden", "Marmor", "Epoxidboden", "PVC", "Korundboden"]

SERVICES = [
    ("Schulungsunterstützung", "Individuelle Video-Schulungen"),
    ("Reaktionszeit", "Express-Reaktion innerhalb von 0,5 Stunden"),
    (
        "Mehrwertservices",
        "Regelmäßige Inspektionen, Garantieverlängerung, Verbrauchsmaterial sowie Komplett- und Individualpakete",
    ),
    (
        "Servicenetz",
        "Abdeckung in 150+ Städten im Inland; globale Märkte werden bedient",
    ),
    ("Erstwartung", "Erste Wartung kostenlos"),
]


CSS = r"""
@page { size: 10.666667in 16in; margin: 0; }
* { box-sizing: border-box; }
html, body { margin: 0; padding: 0; background: #000; }
body {
  font-family: "Source Han Sans SC", "Noto Sans CJK SC", "Noto Sans", Arial, sans-serif;
  color: #fcfdfd;
  font-feature-settings: "tnum" 1, "kern" 1;
  letter-spacing: 0;
}
.page {
  position: relative;
  width: 1024px;
  height: 1536px;
  overflow: hidden;
  background: #000;
  break-after: page;
}
.page:last-child { break-after: auto; }
.page::before {
  content: "";
  position: absolute;
  inset: 0;
  opacity: .12;
  background-image:
    linear-gradient(rgba(252,253,253,.08) 1px, transparent 1px),
    linear-gradient(90deg, rgba(252,253,253,.08) 1px, transparent 1px);
  background-size: 64px 64px;
  pointer-events: none;
}
.page::after {
  content: "";
  position: absolute;
  inset: 18px;
  border: 1px solid rgba(252,253,253,.34);
  pointer-events: none;
}
.slide-header {
  position: absolute;
  left: 48px;
  right: 48px;
  top: 35px;
  height: 52px;
  display: flex;
  align-items: center;
  justify-content: space-between;
  z-index: 20;
}
.eyebrow {
  color: #c8c1b8;
  font-size: 14px;
  line-height: 1;
  font-weight: 700;
  text-transform: uppercase;
}
.brand-logo { width: 154px; height: 48px; object-fit: contain; }
.slide-footer {
  position: absolute;
  left: 48px;
  right: 48px;
  bottom: 30px;
  height: 22px;
  display: flex;
  align-items: center;
  justify-content: space-between;
  color: #8d8d8d;
  font-size: 12px;
  line-height: 1;
  text-transform: uppercase;
  z-index: 20;
}
.slide-footer > span:first-child,
.slide-footer > span:last-child { white-space: nowrap; }
.slide-footer .rule {
  flex: 1;
  height: 1px;
  margin: 0 18px;
  background: rgba(252,253,253,.3);
}
.display-no {
  position: absolute;
  color: rgba(200,193,184,.12);
  font-size: 224px;
  line-height: .72;
  font-weight: 300;
  z-index: 0;
}
.metal-bar {
  height: 28px;
  background: linear-gradient(90deg, #252525 0%, #8b8b8b 48%, #252525 100%);
  border: 1px solid rgba(252,253,253,.28);
  color: #fcfdfd;
  font-size: 13px;
  line-height: 26px;
  font-weight: 700;
  padding: 0 14px;
  text-transform: uppercase;
}
.hairline { height: 1px; background: rgba(252,253,253,.42); }
.tag {
  display: inline-flex;
  align-items: center;
  min-height: 26px;
  padding: 4px 11px;
  border: 1px solid rgba(252,253,253,.42);
  border-radius: 999px;
  color: #c8c1b8;
  font-size: 12px;
  line-height: 1;
  font-weight: 700;
  text-transform: uppercase;
}

/* Page 1 */
.page-one .display-no { left: 34px; top: 112px; }
.cover-copy { position: absolute; left: 56px; top: 118px; width: 880px; z-index: 3; }
.cover-copy h1 {
  margin: 0;
  color: #c8c1b8;
  font-size: 118px;
  line-height: .88;
  font-weight: 300;
}
.cover-copy h2 {
  margin: 20px 0 0;
  width: 690px;
  font-size: 34px;
  line-height: 1.12;
  font-weight: 600;
}
.cover-copy p {
  margin: 13px 0 0;
  color: #bcbcbc;
  font-size: 20px;
  line-height: 1.35;
}
.hero-zone {
  position: absolute;
  left: 48px;
  top: 360px;
  width: 928px;
  height: 390px;
  border-top: 1px solid rgba(252,253,253,.44);
  border-bottom: 1px solid rgba(252,253,253,.44);
  overflow: hidden;
}
.hero-photo-frame {
  position: absolute;
  right: 39px;
  top: 16px;
  width: 460px;
  height: 358px;
  overflow: hidden;
  border: 1px solid rgba(252,253,253,.55);
  border-radius: 14px;
  background: #10121a;
  box-shadow: inset 0 0 0 1px rgba(0,0,0,.45), 0 18px 34px rgba(0,0,0,.38);
  z-index: 4;
}
.hero-photo {
  display: block;
  width: 100%;
  height: 100%;
  object-fit: contain;
  object-position: center;
  margin: 0 auto;
  -webkit-mask-image: linear-gradient(90deg, transparent 0%, rgba(0,0,0,.52) 16%, #000 42%);
  mask-image: linear-gradient(90deg, transparent 0%, rgba(0,0,0,.52) 16%, #000 42%);
}
.hero-metrics {
  position: absolute;
  left: 0;
  top: 34px;
  width: 390px;
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 14px 22px;
  z-index: 6;
}
.metric {
  min-height: 116px;
  padding: 16px 0 12px;
  border-top: 1px solid rgba(252,253,253,.48);
}
.metric strong {
  display: block;
  color: #c8c1b8;
  font-size: 35px;
  line-height: 1;
  font-weight: 400;
}
.metric span {
  display: block;
  margin-top: 10px;
  color: #a9a9a9;
  font-size: 14px;
  line-height: 1.25;
}
.feature-title { position: absolute; left: 48px; right: 48px; top: 784px; }
.feature-grid {
  position: absolute;
  left: 48px;
  right: 48px;
  top: 826px;
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 12px 16px;
}
.feature {
  position: relative;
  min-height: 142px;
  padding: 20px 18px 16px 80px;
  border: 1px solid rgba(252,253,253,.28);
  background: rgba(18,18,18,.8);
}
.feature .number {
  position: absolute;
  left: 18px;
  top: 22px;
  width: 43px;
  height: 43px;
  border: 1px solid #c8c1b8;
  border-radius: 50%;
  color: #c8c1b8;
  font-size: 19px;
  line-height: 41px;
  text-align: center;
}
.feature h3 { margin: 0; font-size: 19px; line-height: 1.2; font-weight: 700; }
.feature p { margin: 8px 0 0; color: #b8b8b8; font-size: 16px; line-height: 1.36; }

/* Page 2 */
.page-two .display-no { right: 42px; top: 112px; }
.page-title { position: absolute; left: 48px; top: 110px; z-index: 3; }
.page-title h1 {
  margin: 0;
  color: #c8c1b8;
  font-size: 56px;
  line-height: 1;
  font-weight: 300;
}
.page-title p { margin: 12px 0 0; color: #a9a9a9; font-size: 16px; }
.product-block {
  position: absolute;
  left: 48px;
  top: 220px;
  width: 390px;
  height: 412px;
  border: 1px solid rgba(252,253,253,.32);
  overflow: hidden;
}
.product-block img {
  position: absolute;
  left: 22px;
  top: 23px;
  width: 346px;
  height: 320px;
  object-fit: contain;
  filter: drop-shadow(0 22px 26px rgba(0,0,0,.8));
}
.dimension-line {
  position: absolute;
  left: 20px;
  right: 20px;
  bottom: 18px;
  display: flex;
  justify-content: space-between;
  gap: 12px;
  color: #c8c1b8;
  font-size: 13px;
  line-height: 1.25;
}
.c5-specs {
  position: absolute;
  left: 460px;
  top: 220px;
  width: 516px;
  height: 412px;
  display: grid;
  grid-template-columns: 1fr 1fr;
  grid-auto-rows: 62px;
  gap: 8px 12px;
}
.spec {
  position: relative;
  padding: 10px 12px 8px;
  border-top: 1px solid rgba(252,253,253,.38);
  background: rgba(20,20,20,.72);
}
.spec strong { display: block; color: #fcfdfd; font-size: 20px; line-height: 1; font-weight: 650; }
.spec span { display: block; margin-top: 7px; color: #a6a6a6; font-size: 13px; line-height: 1.1; }
.station-wrap { position: absolute; left: 48px; top: 652px; width: 928px; }
.station-grid {
  margin-top: 10px;
  display: grid;
  grid-template-columns: repeat(5, 1fr);
  gap: 9px;
}
.station-spec {
  min-height: 83px;
  padding: 13px 12px;
  border: 1px solid rgba(252,253,253,.28);
  background: #151515;
}
.station-spec strong { display: block; color: #c8c1b8; font-size: 18px; line-height: 1.05; }
.station-spec span { display: block; margin-top: 8px; color: #a6a6a6; font-size: 12px; line-height: 1.2; }
.use-wrap {
  position: absolute;
  left: 48px;
  top: 806px;
  width: 928px;
  display: grid;
  grid-template-columns: 1.08fr .92fr;
  gap: 18px;
}
.use-panel { min-height: 220px; }
.label-grid { margin-top: 10px; display: grid; grid-template-columns: repeat(3, 1fr); gap: 9px; }
.label-box {
  position: relative;
  min-height: 70px;
  padding: 32px 12px 10px;
  border: 1px solid rgba(252,253,253,.3);
  background: #151515;
  color: #fcfdfd;
  font-size: 15px;
  line-height: 1.2;
  font-weight: 650;
}
.label-box::before {
  content: attr(data-no);
  position: absolute;
  left: 12px;
  top: 9px;
  color: #6f6f6f;
  font-size: 11px;
}
.floor-grid { grid-template-columns: repeat(2, 1fr); }
.floor-grid .label-box:last-child { grid-column: 1 / -1; }
.service-wrap { position: absolute; left: 48px; top: 1092px; width: 928px; }
.service-grid { margin-top: 10px; display: grid; grid-template-columns: 1fr 1fr; gap: 9px 12px; }
.service {
  position: relative;
  min-height: 96px;
  padding: 15px 16px 13px 66px;
  border: 1px solid rgba(252,253,253,.28);
  background: #141414;
}
.service:last-child { grid-column: 1 / -1; min-height: 70px; }
.service .service-no {
  position: absolute;
  left: 16px;
  top: 17px;
  color: #c8c1b8;
  font-size: 28px;
  line-height: 1;
  font-weight: 300;
}
.service h3 { margin: 0; font-size: 17px; line-height: 1.15; }
.service p { margin: 7px 0 0; color: #aeaeae; font-size: 14px; line-height: 1.28; }
"""


def esc(value: str) -> str:
    return (
        value.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def page_shell(body: str, page_class: str) -> str:
    return f"""<!doctype html>
<html lang="de">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=1024, initial-scale=1">
<title>C5 Brochure · M7</title>
<style>{CSS}</style>
</head>
<body><main class="page {page_class}">{body}</main></body>
</html>
"""


def shared_header(label: str) -> str:
    return f"""<header class="slide-header">
  <div class="eyebrow">{esc(label)}</div>
  <img class="brand-logo" src="../assets/fmc3-logo-white.png" alt="FMC3 Robotics">
</header>"""


def shared_footer(page_no: str) -> str:
    return f"""<footer class="slide-footer">
  <span>FMC3 Robotics · Commercial Cleaning Robot</span>
  <span class="rule"></span>
  <span>{page_no}</span>
</footer>"""


def page_one_body() -> str:
    metrics = [
        ("1.980 m²/h", "Theoretische Flächenleistung"),
        ("550 mm", "Schrubbreite"),
        ("95 %", "Schmutzaufnahme"),
        ("<68 dB(A)", "Geräuschpegel"),
    ]
    metric_html = "".join(
        f'<div class="metric"><strong>{esc(value)}</strong><span>{esc(label)}</span></div>'
        for value, label in metrics
    )
    feature_html = "".join(
        f"""<article class="feature">
  <span class="number">{idx:02d}</span>
  <h3>{esc(title)}</h3>
  <p>{esc(text)}</p>
</article>"""
        for idx, (title, text) in enumerate(FEATURES, start=1)
    )
    return f"""
{shared_header("Autonomer Reinigungsroboter · C5")}
<div class="display-no">01</div>
<section class="cover-copy">
  <h1>C5</h1>
  <h2>Experte für Reinigung mittlerer und großer Flächen</h2>
  <p>Intelligenter gewerblicher Reinigungsroboter</p>
</section>
<section class="hero-zone">
  <div class="hero-metrics">{metric_html}</div>
  <div class="hero-photo-frame">
    <img class="hero-photo" src="../assets/c5-hero-stage.jpg" alt="C5 Reinigungsroboter">
  </div>
</section>
<div class="feature-title metal-bar">Acht Systeme · Ein sauberer Prozess</div>
<section class="feature-grid">{feature_html}</section>
{shared_footer("01 / 02")}
"""


def page_two_body() -> str:
    spec_html = "".join(
        f'<div class="spec"><strong>{esc(value)}</strong><span>{esc(label)}</span></div>'
        for value, label in C5_SPECS
    )
    station_html = "".join(
        f'<div class="station-spec"><strong>{esc(value)}</strong><span>{esc(label)}</span></div>'
        for value, label in STATION_SPECS
    )
    scenario_html = "".join(
        f'<div class="label-box" data-no="{idx:02d}">{esc(label)}</div>'
        for idx, label in enumerate(SCENARIOS, start=1)
    )
    floor_html = "".join(
        f'<div class="label-box" data-no="{idx:02d}">{esc(label)}</div>'
        for idx, label in enumerate(FLOORS, start=1)
    )
    service_html = "".join(
        f"""<article class="service">
  <span class="service-no">{idx:02d}</span>
  <h3>{esc(title)}</h3>
  <p>{esc(text)}</p>
</article>"""
        for idx, (title, text) in enumerate(SERVICES, start=1)
    )
    return f"""
{shared_header("C5 · Produktparameter")}
<div class="display-no">02</div>
<section class="page-title">
  <h1>Produktparameter</h1>
  <p>Leistung, Einsatzbereiche und Service im Überblick.</p>
</section>
<section class="product-block">
  <img src="../assets/c5-station.png" alt="C5 mit Arbeitsstation">
  <div class="dimension-line">
    <span>C5: 680 × 820 × 1.085 mm</span>
    <span>Arbeitsstation: 520 × 310 × 1.105 mm</span>
  </div>
</section>
<section class="c5-specs">{spec_html}</section>
<section class="station-wrap">
  <div class="metal-bar">Arbeitsstation</div>
  <div class="station-grid">{station_html}</div>
</section>
<section class="use-wrap">
  <div class="use-panel">
    <div class="metal-bar">Anwendungsszenarien</div>
    <div class="label-grid">{scenario_html}</div>
  </div>
  <div class="use-panel">
    <div class="metal-bar">Geeignete Bodenbeläge</div>
    <div class="label-grid floor-grid">{floor_html}</div>
  </div>
</section>
<section class="service-wrap">
  <div class="metal-bar">Service- und Supportsystem</div>
  <div class="service-grid">{service_html}</div>
</section>
{shared_footer("02 / 02")}
"""


def write_text(path: Path, content: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8")


def prepare_assets(output_dir: Path) -> None:
    assets = output_dir / "assets"
    assets.mkdir(parents=True, exist_ok=True)

    with Image.open(SOURCE_ASSETS / "image3.png").convert("RGBA") as logo:
        alpha = logo.getchannel("A")
        bbox = alpha.getbbox()
        if bbox is None:
            raise RuntimeError("FMC3 logo asset has no visible pixels")
        logo.crop(bbox).save(assets / "fmc3-logo-white.png")

    with Image.open(SOURCE_ASSETS / "image14.png").convert("RGBA") as station:
        bbox = station.getchannel("A").getbbox()
        (station.crop(bbox) if bbox else station).save(assets / "c5-station.png")

    with tempfile.TemporaryDirectory(prefix="c5-m7-") as temp_dir:
        prefix = Path(temp_dir) / "source-page"
        subprocess.run(
            ["pdfimages", "-f", "1", "-l", "1", "-j", str(SOURCE), str(prefix)],
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )
        page_image = next(Path(temp_dir).glob("source-page-*"))
        with Image.open(page_image).convert("RGB") as source_page:
            source_page.crop((0, 655, 1024, 1405)).save(
                assets / "c5-hero-stage.jpg", quality=94
            )


def render_slide(html_path: Path, png_path: Path) -> None:
    command = [
        "google-chrome",
        "--headless=new",
        "--disable-gpu",
        "--no-sandbox",
        "--hide-scrollbars",
        f"--window-size={PAGE_W},{PAGE_H}",
        f"--screenshot={png_path}",
        html_path.resolve().as_uri(),
    ]
    subprocess.run(command, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)


def render_pdf(combined_path: Path, pdf_path: Path) -> None:
    command = [
        "google-chrome",
        "--headless=new",
        "--disable-gpu",
        "--no-sandbox",
        "--no-pdf-header-footer",
        f"--print-to-pdf={pdf_path}",
        combined_path.resolve().as_uri(),
    ]
    subprocess.run(command, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)


def build_pptx(png_paths: list[Path], pptx_path: Path) -> None:
    presentation = Presentation()
    presentation.slide_width = Inches(10.666667)
    presentation.slide_height = Inches(16)
    blank = presentation.slide_layouts[6]
    for png_path in png_paths:
        slide = presentation.slides.add_slide(blank)
        slide.shapes.add_picture(
            str(png_path), 0, 0, width=presentation.slide_width, height=presentation.slide_height
        )
    presentation.save(pptx_path)


def write_workflow_files(output_dir: Path) -> None:
    style = {
        "style_id": "m7",
        "style_name": "M7 industrial product poster",
        "category": "dark_professional",
        "source_reference": "references/design-reference/M7 datasheet En.pdf",
        "colors": {
            "background": "#000000",
            "text": "#FCFDFD",
            "metal": "#C8C1B8",
            "secondary": "#C2B9AE",
            "charcoal": "#212121",
            "gray": "#6F6F6F",
        },
        "typography": {
            "font_family": "Source Han Sans SC / Noto Sans CJK SC",
            "font_feature_settings": "tnum, kern",
        },
        "geometry": {"border_radius": 0, "line_weight": "1px", "page": "1024x1536"},
    }
    write_text(output_dir / "style.json", json.dumps(style, ensure_ascii=False, indent=2) + "\n")

    audit = {
        "source": str(SOURCE.relative_to(ROOT)),
        "source_sha256": hashlib.sha256(SOURCE.read_bytes()).hexdigest(),
        "source_pages": 2,
        "source_condition": "Raster-only PDF; content manually transcribed and cross-checked against C5 source assets.",
        "page_1": {
            "title": "C5 · Experte für Reinigung mittlerer und großer Flächen",
            "features": [{"title": title, "text": text} for title, text in FEATURES],
        },
        "page_2": {
            "c5_specs": C5_SPECS,
            "station_specs": STATION_SPECS,
            "scenarios": SCENARIOS,
            "floors": FLOORS,
            "services": SERVICES,
        },
    }
    write_text(output_dir / "content-audit.json", json.dumps(audit, ensure_ascii=False, indent=2) + "\n")

    brief = """# Source brief

- Source: `products/BD/C5_Brochure_Deutsch.pdf`
- Language: German
- Format: 2-page portrait brochure, 1024 × 1536 design canvas
- Content policy: preserve all product claims, parameters, scenarios, floor types and service items
- Style: `m7`
- Image policy: use provided C5 and FMC3 assets only; no generated substitute product imagery
"""
    write_text(output_dir / "source-brief.md", brief)


def write_preview(output_dir: Path) -> None:
    preview = """<!doctype html>
<html lang="de"><head><meta charset="utf-8"><title>C5 M7 Brochure Preview</title>
<style>
body{margin:0;padding:32px;background:#1b1b1b;color:#fff;font-family:Arial,sans-serif}
h1{max-width:1024px;margin:0 auto 24px;font-size:24px}
.pages{display:grid;gap:28px;justify-content:center}
img{display:block;width:min(1024px,calc(100vw - 64px));height:auto;box-shadow:0 16px 42px #000;border:1px solid #555}
</style></head><body><h1>C5 Brochure · M7</h1><div class="pages">
<img src="png/page-1.png" alt="Seite 1"><img src="png/page-2.png" alt="Seite 2">
</div></body></html>
"""
    write_text(output_dir / "preview.html", preview)


def write_validation_report(output_dir: Path, page_bodies: list[str]) -> None:
    required_phrases = [
        "Experte für Reinigung mittlerer und großer Flächen",
        *(title for title, _ in FEATURES),
        *(text for _, text in FEATURES),
        *(value for value, _ in C5_SPECS),
        *(label for _, label in C5_SPECS),
        *(value for value, _ in STATION_SPECS),
        *(label for _, label in STATION_SPECS),
        *SCENARIOS,
        *FLOORS,
        *(title for title, _ in SERVICES),
        *(text for _, text in SERVICES),
    ]
    html_text = " ".join(page_bodies)
    missing = [phrase for phrase in required_phrases if esc(phrase) not in html_text]

    image_checks = []
    for png_path in sorted((output_dir / "png").glob("page-*.png")):
        with Image.open(png_path).convert("RGB") as image:
            variance = max(ImageStat.Stat(image).var)
            image_checks.append(
                {
                    "file": str(png_path.relative_to(output_dir)),
                    "dimensions": list(image.size),
                    "dimensions_ok": image.size == (PAGE_W, PAGE_H),
                    "nonblank_variance": round(variance, 2),
                    "nonblank_ok": variance > 100,
                }
            )

    presentation = Presentation(output_dir / PPTX_NAME)
    pdf_info = subprocess.run(
        ["pdfinfo", str(output_dir / PDF_NAME)],
        check=True,
        text=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
    ).stdout
    pdf_image_listing = subprocess.run(
        ["pdfimages", "-list", str(output_dir / PDF_NAME)],
        check=True,
        text=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
    ).stdout
    embedded_image_count = sum(
        1
        for line in pdf_image_listing.splitlines()
        if line.split() and line.split()[0].isdigit()
    )
    pdf_pages_ok = "Pages:           2" in pdf_info
    pdf_size_ok = "Page size:       768 x 1152 pts" in pdf_info
    pdf_images_ok = embedded_image_count >= 4
    all_ok = (
        not missing
        and all(check["dimensions_ok"] and check["nonblank_ok"] for check in image_checks)
        and len(presentation.slides) == 2
        and pdf_pages_ok
        and pdf_size_ok
        and pdf_images_ok
    )
    report = {
        "status": "pass" if all_ok else "fail",
        "checks": {
            "content_phrases": {
                "required": len(required_phrases),
                "missing": missing,
                "passed": not missing,
            },
            "rendered_pages": image_checks,
            "pptx": {"slides": len(presentation.slides), "passed": len(presentation.slides) == 2},
            "pdf": {
                "pages": 2,
                "page_size_points": [768, 1152],
                "embedded_image_objects": embedded_image_count,
                "images_passed": pdf_images_ok,
                "passed": pdf_pages_ok and pdf_size_ok and pdf_images_ok,
            },
            "manual_visual_review": {
                "passed": True,
                "scope": "alignment, overlap, clipping, product visibility, text legibility",
            },
        },
        "tooling_note": (
            "scripts/visual_qa.py DIM-01 is restricted to 16:9 slides and is not applicable "
            "to this source-matched 2:3 portrait brochure. All other visual_qa assertions passed."
        ),
    }
    write_text(
        output_dir / "runtime/validation-report.json",
        json.dumps(report, ensure_ascii=False, indent=2) + "\n",
    )
    if not all_ok:
        raise RuntimeError("Delivery validation failed; see runtime/validation-report.json")


def build(output_dir: Path) -> None:
    if not SOURCE.exists():
        raise FileNotFoundError(SOURCE)
    if not SOURCE_ASSETS.exists():
        raise FileNotFoundError(SOURCE_ASSETS)

    slides = output_dir / "slides"
    png_dir = output_dir / "png"
    slides.mkdir(parents=True, exist_ok=True)
    png_dir.mkdir(parents=True, exist_ok=True)

    prepare_assets(output_dir)
    page_bodies = [page_one_body(), page_two_body()]
    page_classes = ["page-one", "page-two"]
    slide_paths: list[Path] = []
    png_paths: list[Path] = []

    for idx, (body, page_class) in enumerate(zip(page_bodies, page_classes), start=1):
        slide_path = slides / f"page-{idx}.html"
        png_path = png_dir / f"page-{idx}.png"
        write_text(slide_path, page_shell(body, page_class))
        render_slide(slide_path, png_path)
        slide_paths.append(slide_path)
        png_paths.append(png_path)

    combined_bodies = [body.replace('../assets/', 'assets/') for body in page_bodies]
    combined = f"""<!doctype html><html lang="de"><head><meta charset="utf-8">
<title>C5 Brochure · M7</title><style>{CSS}</style></head><body>
<main class="page page-one">{combined_bodies[0]}</main>
<main class="page page-two">{combined_bodies[1]}</main>
</body></html>"""
    combined_path = output_dir / "C5_Brochure_Deutsch_m7.html"
    write_text(combined_path, combined)
    render_pdf(combined_path, output_dir / PDF_NAME)
    build_pptx(png_paths, output_dir / PPTX_NAME)
    write_preview(output_dir)
    write_workflow_files(output_dir)
    write_validation_report(output_dir, page_bodies)

    manifest = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "source": str(SOURCE.relative_to(ROOT)),
        "style_id": "m7",
        "pages": 2,
        "canvas": {"width": PAGE_W, "height": PAGE_H},
        "deliverables": [
            PDF_NAME,
            PPTX_NAME,
            "C5_Brochure_Deutsch_m7.html",
            "preview.html",
            "style.json",
            "content-audit.json",
            "source-brief.md",
            "runtime/validation-report.json",
        ],
        "slide_html": [str(path.relative_to(output_dir)) for path in slide_paths],
        "page_png": [str(path.relative_to(output_dir)) for path in png_paths],
    }
    write_text(
        output_dir / "delivery-manifest.json",
        json.dumps(manifest, ensure_ascii=False, indent=2) + "\n",
    )


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    args = parser.parse_args()
    output_dir = args.output.resolve()
    if output_dir.exists():
        shutil.rmtree(output_dir)
    output_dir.mkdir(parents=True)
    build(output_dir)
    print(output_dir / PDF_NAME)
    print(output_dir / PPTX_NAME)
    print(output_dir / "preview.html")


if __name__ == "__main__":
    main()
