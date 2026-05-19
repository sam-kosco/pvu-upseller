"""
PVU Upsell PDF Generator
========================
Takes the structured dict from parse_payload.py and produces a
polished sales PDF matching the Foxtrot Aviation example document style.

Layout:
  - Page 1 header:    Foxtrot Aviation logo (centered, page 1 only)
  - Every page header (small, top-right): "Foxtrot Aviation Services /
                                           Duncan Aviation PVU / MM/DD/YYYY"
  - Every page footer (centered):  Stephen's name | phone | email
  - Body: aircraft header, Stephen's intro, three supersections

Services are organized into three supersections:
  1. Metal Polish & Protection      — Brightwork, Xylon
  2. Paint Correction & Coatings    — Ceramic Coating, Permagard, Polymer
  3. Detail Work                    — Interior Detail, Exterior Detail, Carpet Extraction

Note: Xylon's AI observation and field notes appear BEFORE the Brightwork
example (before/after) photos, because the marketing photos show a plane that
received both services together.

Requirements:
    pip install reportlab pillow requests anthropic numpy

Place in the same repo directory:
    logo.png            — Foxtrot Aviation logo
    service_context.md  — Service knowledge base for the AI rewriter
"""

import base64
import io
import os
import requests
import anthropic
from datetime import date
from PIL import Image
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Image as RLImage,
    Table, TableStyle, HRFlowable, KeepTogether
)
from reportlab.lib.enums import TA_CENTER
from reportlab.pdfgen import canvas as rl_canvas

# ─────────────────────────────────────────────────────────────
# PAGE GEOMETRY
# ─────────────────────────────────────────────────────────────
PAGE_W, PAGE_H = letter
MARGIN         = 0.75 * inch
CONTENT_W      = PAGE_W - 2 * MARGIN

# Logo on page 1 (2816×1270 → 2.2:1 aspect)
LOGO_W         = 3.0 * inch
LOGO_H         = LOGO_W / 2.216          # ~1.35 inch
HEADER_HEIGHT  = LOGO_H + 0.25 * inch   # space reserved at top of page 1

# Footer and small header heights reserved every page
FOOTER_HEIGHT  = 0.45 * inch
SMALL_HDR_H    = 0.35 * inch

# ─────────────────────────────────────────────────────────────
# SUPERSECTION STRUCTURE
# ─────────────────────────────────────────────────────────────
SUPERSECTIONS = [
    {
        "title":    "Metal Polish & Protection",
        "services": ["Brightwork", "Xylon"],
        "boilerplate": (
            "Your aircraft's bare metal surfaces — leading edges, nacelles, wing tips, "
            "and stabilizers — are among the most exposed components on the airframe. "
            "Regular polishing and protection keep these surfaces free of oxidation and "
            "corrosion, preserve their appearance, and can even improve aerodynamic "
            "performance by maintaining a smooth, laminar surface."
        ),
    },
    {
        "title":    "Paint Correction & Protective Coatings",
        "services": ["Ceramic Coating", "Permagard Coating", "Polymer Coating"],
        "boilerplate": (
            "Aircraft paint is constantly under attack from UV radiation, exhaust carbon, "
            "hydraulic fluids, and environmental contaminants. Left unprotected, even "
            "well-maintained paint oxidizes, becomes porous, and loses its gloss — "
            "increasing drag and reducing long-term value. The options below represent "
            "good, better, and best levels of protection, all completable without "
            "affecting your return-to-service date."
        ),
    },
    {
        "title":    "Detail Work",
        "services": ["Interior Detail", "Exterior Detail", "Carpet Extraction"],
        "boilerplate": (
            "A thorough detail during your maintenance event is the most cost-effective "
            "way to protect your aircraft's interior and exterior surfaces between "
            "major service intervals. Our team is already on-site and familiar with "
            "your aircraft, making this the ideal time to address accumulated wear "
            "and restore a like-new appearance inside and out."
        ),
    },
]

# ─────────────────────────────────────────────────────────────
# Services whose example (before/after) photos are shared with
# the PREVIOUS service in the same supersection.
# Xylon's example photos are Brightwork's photos (same job).
# ─────────────────────────────────────────────────────────────
DEFER_EXAMPLE_PHOTOS_TO = {
    "Xylon": "Brightwork",   # Xylon text appears first; photos render under Brightwork
}

# ─────────────────────────────────────────────────────────────
# CONFIG — toggle before/after marketing photos per service
# ─────────────────────────────────────────────────────────────
INCLUDE_EXAMPLE_PHOTOS = {
    "Brightwork":        True,
    "Ceramic Coating":   True,
    "Permagard Coating": True,
    "Polymer Coating":   True,
    "Interior Detail":   False,
    "Exterior Detail":   True,
    "Carpet Extraction": True,
    "Xylon":             False,   # Xylon shares Brightwork's photos — handled via deferral
}

# ─────────────────────────────────────────────────────────────
# SharePoint config
# ─────────────────────────────────────────────────────────────
DRIVE_ID           = "b!_bzXaIx86kOufgJN3ih-BaDIDthKYuxJkJtLi1Bm5irGjCEnK-VHSpBRRm3_SDKU"
EXAMPLE_PHOTO_BASE = "Assets/Service Example Photos"

EXAMPLE_PHOTO_NAMES = {
    "Brightwork":        ("Brightwork Before.jpg",        "Brightwork After.jpg"),
    "Ceramic Coating":   ("Ceramic Coating Before.jpg",   "Ceramic Coating After.jpg"),
    "Permagard Coating": ("Permagard Before.jpg",         "Permagard After.jpg"),
    "Polymer Coating":   ("Polymer Before.jpg",           "Polymer After.jpg"),
    "Interior Detail":   None,
    "Exterior Detail":   None,
    "Carpet Extraction": ("Carpet Extraction Before.jpg", "Carpet Extraction After.jpg"),
    "Xylon":             None,
}

# ─────────────────────────────────────────────────────────────
# Per-service boilerplate — Stephen replaces placeholders
# ─────────────────────────────────────────────────────────────
SERVICE_BOILERPLATE = {
    "Brightwork": (
        "Brightwork Boilerplate Text — Stephen to provide final copy. "
        "This section describes the brightwork polishing service, what it addresses, "
        "and why it is recommended for this aircraft type."
    ),
    "Xylon": (
        "Xylon Boilerplate Text — Stephen to provide final copy. "
        "This section describes the Xylon corrosion-inhibiting treatment "
        "and the long-term protection it provides for leading edges."
    ),
    "Ceramic Coating": (
        "Ceramic Coating Boilerplate Text — Stephen to provide final copy. "
        "This section describes the paint correction and ceramic coating process, "
        "the 3-year warranty, hydrophobic properties, and performance benefits."
    ),
    "Permagard Coating": (
        "Permagard Boilerplate Text — Stephen to provide final copy. "
        "This section describes Permagard as the industry standard in paint protection "
        "and the yearly booster treatment requirement."
    ),
    "Polymer Coating": (
        "Polymer Coating Boilerplate Text — Stephen to provide final copy. "
        "This section describes the polymer coating option, its protection level, "
        "and how long it typically lasts."
    ),
    "Interior Detail": (
        "Interior Detail Boilerplate Text — Stephen to provide final copy. "
        "This section describes the interior detail service including deep cleaning "
        "of all cabin surfaces, leather, and soft goods."
    ),
    "Exterior Detail": (
        "Exterior Detail Boilerplate Text — Stephen to provide final copy. "
        "This section describes the exterior detail service and what it covers."
    ),
    "Carpet Extraction": (
        "Carpet Extraction Boilerplate Text — Stephen to provide final copy. "
        "This section describes carpet extraction, why it is recommended at every "
        "maintenance event, and what contaminants it removes."
    ),
}

# ─────────────────────────────────────────────────────────────
# Stephen's contact info
# ─────────────────────────────────────────────────────────────
STEPHEN_NAME  = "Stephen Chadbourn"
STEPHEN_PHONE = "520.981.8942"
STEPHEN_EMAIL = "stephen.chadbourn@foxtrotaviation.com"

STEPHEN_INTRO = (
    "Hello, I'm Stephen Chadbourn, General Manager of the Foxtrot Aviation team at "
    "Duncan Aviation in Provo, Utah. We provide aircraft detailing during maintenance "
    "events and are committed to returning your aircraft in exceptional condition."
)

# ─────────────────────────────────────────────────────────────
# Load service_context.md once at module startup
# ─────────────────────────────────────────────────────────────
def _load_service_context() -> str:
    here = os.path.dirname(os.path.abspath(__file__))
    path = os.path.join(here, "service_context.md")
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            return f.read()
    print("  Warning: service_context.md not found.")
    return ""

SERVICE_CONTEXT = _load_service_context()


# ─────────────────────────────────────────────────────────────
# Resolve logo path
# ─────────────────────────────────────────────────────────────
def _find_logo() -> str | None:
    here = os.path.dirname(os.path.abspath(__file__))
    for p in [
        os.path.join(here, "logo.png"),
        "/mnt/user-data/uploads/logo.png",
        "/mnt/project/Fox_Logo_Red_background.png",
    ]:
        if os.path.exists(p):
            return p
    return None

LOGO_PATH = _find_logo()


# ─────────────────────────────────────────────────────────────
# Canvas callbacks — logo header (page 1), small info header
# (all pages), and footer (all pages)
# ─────────────────────────────────────────────────────────────
def _draw_small_header(c: rl_canvas.Canvas) -> None:
    """
    Small right-aligned header on every page:
        Foxtrot Aviation Services   Duncan Aviation PVU   MM/DD/YYYY
    """
    today     = date.today().strftime("%m/%d/%Y")
    hdr_text  = f"Foxtrot Aviation Services   |   Duncan Aviation PVU   |   {today}"
    c.saveState()
    c.setFont("Helvetica", 7.5)
    c.setFillColor(colors.HexColor("#777777"))
    y = PAGE_H - MARGIN * 0.55
    c.drawRightString(PAGE_W - MARGIN, y, hdr_text)
    # Thin rule underneath
    c.setStrokeColor(colors.HexColor("#dddddd"))
    c.setLineWidth(0.3)
    c.line(MARGIN, y - 4, PAGE_W - MARGIN, y - 4)
    c.restoreState()


def _draw_footer(c: rl_canvas.Canvas) -> None:
    """Centered footer on every page: Name  |  Phone  |  Email"""
    footer_text = f"{STEPHEN_NAME}   |   {STEPHEN_PHONE}   |   {STEPHEN_EMAIL}"
    c.saveState()
    c.setFont("Helvetica", 8)
    c.setFillColor(colors.HexColor("#555555"))
    rule_y = MARGIN * 0.55
    c.setStrokeColor(colors.HexColor("#cccccc"))
    c.setLineWidth(0.4)
    c.line(MARGIN, rule_y + 10, PAGE_W - MARGIN, rule_y + 10)
    c.drawCentredString(PAGE_W / 2, rule_y - 2, footer_text)
    c.restoreState()


def _draw_logo_header(c: rl_canvas.Canvas) -> None:
    """Centered logo at top of page 1 only."""
    if not LOGO_PATH:
        return
    x = (PAGE_W - LOGO_W) / 2
    y = PAGE_H - MARGIN - LOGO_H
    c.drawImage(
        LOGO_PATH, x, y,
        width=LOGO_W, height=LOGO_H,
        preserveAspectRatio=True,
        mask="auto",
    )


def on_first_page(c: rl_canvas.Canvas, doc) -> None:
    _draw_logo_header(c)
    _draw_small_header(c)
    _draw_footer(c)


def on_later_pages(c: rl_canvas.Canvas, doc) -> None:
    _draw_small_header(c)
    _draw_footer(c)


# ─────────────────────────────────────────────────────────────
# PIL Image -> base64 JPEG (for Claude vision API)
# ─────────────────────────────────────────────────────────────
def _pil_to_b64(pil_img: Image.Image, max_px: int = 1024) -> str:
    img = pil_img.convert("RGB")
    w, h = img.size
    if max(w, h) > max_px:
        scale = max_px / max(w, h)
        img = img.resize((int(w * scale), int(h * scale)), Image.LANCZOS)
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=75)
    return base64.b64encode(buf.getvalue()).decode("utf-8")


# ─────────────────────────────────────────────────────────────
# AI note rewriter
# ─────────────────────────────────────────────────────────────
def rewrite_notes(raw_notes: str, service_name: str, condition_photos: list) -> str:
    if not raw_notes and not condition_photos:
        return ""

    boilerplate = SERVICE_BOILERPLATE.get(service_name, "")
    boilerplate_is_placeholder = boilerplate.startswith(service_name + " Boilerplate Text")

    content = []
    instruction = (
        f"You are writing one paragraph for a professional aircraft detailing proposal "
        f"sent by Foxtrot Aviation to an aircraft owner or operator.\n\n"
        f"SERVICE: {service_name}\n\n"
    )
    if boilerplate and not boilerplate_is_placeholder:
        instruction += (
            f"The proposal already contains this boilerplate paragraph:\n"
            f"\"\"\"\n{boilerplate}\n\"\"\"\n\n"
            f"Do NOT repeat, paraphrase, or restate anything already covered above.\n\n"
        )
    instruction += (
        f"Stephen's raw field note: \"{raw_notes}\"\n\n" if raw_notes
        else "No field note provided. Write based on the photos.\n\n"
    )
    if condition_photos:
        instruction += (
            f"The following {min(len(condition_photos), 3)} photo(s) show the actual "
            f"condition of this aircraft. Use what you see.\n\n"
        )
    instruction += (
        "Write 2-3 sentences: warm, professional, client-facing. Describe the specific "
        "condition observed and why the service is recommended. No heading, no preamble."
    )
    content.append({"type": "text", "text": instruction})
    for photo in condition_photos[:3]:
        try:
            content.append({
                "type": "image",
                "source": {"type": "base64", "media_type": "image/jpeg",
                           "data": _pil_to_b64(photo)},
            })
        except Exception as e:
            print(f"  Warning: could not encode photo for {service_name}: {e}")

    try:
        client = anthropic.Anthropic()
        kwargs = {
            "model":      "claude-sonnet-4-6",
            "max_tokens": 400,
            "messages":   [{"role": "user", "content": content}],
        }
        if SERVICE_CONTEXT:
            kwargs["system"] = (
                "You are an expert aviation detailing specialist writing professional "
                "client-facing proposals for Foxtrot Aviation Services. Use the following "
                "knowledge base but write specifically about the aircraft at hand.\n\n"
                + SERVICE_CONTEXT
            )
        return client.messages.create(**kwargs).content[0].text.strip()
    except Exception as e:
        print(f"  Warning: AI rewrite failed for {service_name}: {e}")
        return raw_notes


# ─────────────────────────────────────────────────────────────
# SharePoint photo fetcher
# ─────────────────────────────────────────────────────────────
def fetch_sharepoint_photo(graph_token: str, filename: str) -> Image.Image | None:
    path = f"{EXAMPLE_PHOTO_BASE}/{filename}"
    url  = f"https://graph.microsoft.com/v1.0/drives/{DRIVE_ID}/root:/{path}:/content"
    try:
        resp = requests.get(url, headers={"Authorization": f"Bearer {graph_token}"},
                            timeout=30)
        if resp.status_code == 200:
            return Image.open(io.BytesIO(resp.content))
        print(f"  Warning: could not fetch {filename} (HTTP {resp.status_code})")
    except Exception as e:
        print(f"  Warning: error fetching {filename}: {e}")
    return None


# ─────────────────────────────────────────────────────────────
# PIL -> ReportLab image helper
# ─────────────────────────────────────────────────────────────
def pil_to_rl(pil_img: Image.Image, max_width: float, max_height: float) -> RLImage:
    buf = io.BytesIO()
    pil_img.convert("RGB").save(buf, format="JPEG", quality=85)
    buf.seek(0)
    w, h  = pil_img.size
    scale = min(max_width / w, max_height / h, 1.0)
    return RLImage(buf, width=w * scale, height=h * scale)


# ─────────────────────────────────────────────────────────────
# Photo block builders
# ─────────────────────────────────────────────────────────────
def build_condition_photo_block(pil_images: list, tail: str = "") -> list:
    if not pil_images:
        return []
    styles    = getSampleStyleSheet()
    caption_s = ParagraphStyle("caption_c", parent=styles["Normal"],
                               fontSize=8, fontName="Helvetica-Bold",
                               alignment=TA_CENTER,
                               textColor=colors.HexColor("#333333"),
                               spaceAfter=3)
    label_text = f"{tail} Current Condition" if tail else "Current Condition"
    n = len(pil_images)
    if n == 1:
        return [
            KeepTogether([
                Paragraph(label_text, caption_s),
                pil_to_rl(pil_images[0], CONTENT_W, 3.5 * inch),
            ]),
            Spacer(1, 8),
        ]
    col_w = (CONTENT_W - 6) / 2
    rows  = []
    for i in range(0, n, 2):
        left  = pil_to_rl(pil_images[i],     col_w, 2.5 * inch)
        right = pil_to_rl(pil_images[i + 1], col_w, 2.5 * inch) if i + 1 < n else ""
        rows.append([left, right])
    tbl = Table(rows, colWidths=[col_w, col_w])
    tbl.setStyle(TableStyle([
        ("VALIGN",(0,0),(-1,-1),"TOP"),
        ("LEFTPADDING",(0,0),(-1,-1),3), ("RIGHTPADDING",(0,0),(-1,-1),3),
        ("TOPPADDING",(0,0),(-1,-1),3),  ("BOTTOMPADDING",(0,0),(-1,-1),3),
    ]))
    return [KeepTogether([Paragraph(label_text, caption_s), tbl]), Spacer(1, 8)]


def build_example_photo_block(before_img, after_img, label: str = "Service Example Photo") -> list:
    if before_img is None and after_img is None:
        return []
    styles    = getSampleStyleSheet()
    caption_s = ParagraphStyle("caption_e", parent=styles["Normal"],
                               fontSize=8, fontName="Helvetica-Bold",
                               alignment=TA_CENTER,
                               textColor=colors.HexColor("#333333"),
                               spaceAfter=3)
    label_s   = ParagraphStyle("label", parent=styles["Normal"],
                               fontSize=9, alignment=TA_CENTER,
                               textColor=colors.HexColor("#444444"))
    col_w = (CONTENT_W - 6) / 2

    def cell(img, before_after):
        if img is None:
            return ""
        return [pil_to_rl(img, col_w, 2.5 * inch), Paragraph(f"<b>{before_after}</b>", label_s)]

    tbl = Table([[cell(before_img, "BEFORE"), cell(after_img, "AFTER")]],
                colWidths=[col_w, col_w])
    tbl.setStyle(TableStyle([
        ("VALIGN",(0,0),(-1,-1),"TOP"), ("ALIGN",(0,0),(-1,-1),"CENTER"),
        ("LEFTPADDING",(0,0),(-1,-1),3), ("RIGHTPADDING",(0,0),(-1,-1),3),
        ("TOPPADDING",(0,0),(-1,-1),3),  ("BOTTOMPADDING",(0,0),(-1,-1),3),
    ]))
    return [
        KeepTogether([
            Paragraph(label, caption_s),
            tbl,
        ]),
        Spacer(1, 8),
    ]


# ─────────────────────────────────────────────────────────────
# Main PDF generator
# ─────────────────────────────────────────────────────────────
def generate_pdf(data: dict, graph_token: str, output_path: str = "upsell.pdf") -> str:
    styles    = getSampleStyleSheet()
    title_s   = ParagraphStyle("title_s", parent=styles["Normal"],
                               fontSize=12, leading=16, spaceAfter=2,
                               fontName="Helvetica-Bold")
    body_s    = ParagraphStyle("body_s", parent=styles["Normal"],
                               fontSize=10, leading=14, spaceAfter=6)
    super_s   = ParagraphStyle("super_s", parent=styles["Normal"],
                               fontSize=13, leading=17, spaceAfter=4,
                               fontName="Helvetica-Bold",
                               textColor=colors.HexColor("#1a1a1a"))
    service_s = ParagraphStyle("service_s", parent=styles["Normal"],
                               fontSize=11, leading=14, spaceAfter=4,
                               fontName="Helvetica-Bold")

    story = []

    # Reserve space for logo on page 1
    story.append(Spacer(1, HEADER_HEIGHT))

    # ── Aircraft header ───────────────────────────────────────
    make_model = " ".join(filter(None, [data.get("make",""), data.get("model","")]))
    # q54 = owner name (new field); q8 = customer/operator (fallback)
    owner = data.get("owner", "").strip() or data.get("customer", "").strip()
    story.append(Paragraph(f"<b>Tail:</b> {data.get('tail','')}", title_s))
    if make_model:
        story.append(Paragraph(f"<b>Model:</b> {make_model}", title_s))
    if owner:
        story.append(Paragraph(f"<b>Owner:</b> {owner}", title_s))
    story.append(Spacer(1, 10))

    # ── Stephen's intro ───────────────────────────────────────
    story.append(Paragraph(STEPHEN_INTRO, body_s))
    story.append(Spacer(1, 8))

    # ── Build upsell lookup ───────────────────────────────────
    upsell_map = {u["service"]: u for u in data.get("upsells", [])}

    # ── Supersections ─────────────────────────────────────────
    for section in SUPERSECTIONS:
        section_upsells = [upsell_map[s] for s in section["services"] if s in upsell_map]
        if not section_upsells:
            continue

        story.append(HRFlowable(width="100%", thickness=1.0,
                                color=colors.HexColor("#888888"), spaceAfter=6))
        story.append(Paragraph(section["title"], super_s))
        if section["boilerplate"]:
            story.append(Paragraph(section["boilerplate"], body_s))
        story.append(Spacer(1, 6))

        # ── Pass 1: render all service text blocks (boilerplate,
        #           AI notes, condition photos) for every service
        #           in this section, in order. No example photos yet.
        # ── Pass 2: after all text is done, append example photos
        #           once per unique photo set (handles Brightwork/Xylon
        #           sharing the same before/after images).
        # ─────────────────────────────────────────────────────

        # Pass 1 — text + condition photos
        for upsell in section_upsells:
            service = upsell["service"]
            price   = upsell.get("price", "")
            notes   = upsell.get("notes", "")
            photos  = upsell.get("photos", [])

            story.append(HRFlowable(width="100%", thickness=0.4,
                                    color=colors.HexColor("#cccccc"), spaceAfter=4))
            price_str = f": ${price}" if price else ""
            story.append(Paragraph(f"{service}{price_str}", service_s))

            svc_boilerplate = SERVICE_BOILERPLATE.get(service, "")
            if svc_boilerplate:
                story.append(Paragraph(svc_boilerplate, body_s))

            if notes or photos:
                print(f"  Rewriting notes for {service} "
                      f"({'with' if photos else 'without'} photos)...")
                polished = rewrite_notes(notes, service, photos)
                if polished:
                    story.append(Paragraph(polished, body_s))

            if photos:
                story.extend(build_condition_photo_block(photos, tail=data.get("tail", "")))

            story.append(Spacer(1, 6))

        # Pass 2 — example (before/after) photos, deduplicated
        # Build a list of unique photo sets to render, resolving
        # DEFER_EXAMPLE_PHOTOS_TO so shared sets appear only once.
        rendered_photo_keys: set[str] = set()
        for upsell in section_upsells:
            service = upsell["service"]

            # Resolve which service's photo set to use
            photo_key = DEFER_EXAMPLE_PHOTOS_TO.get(service, service)

            if photo_key in rendered_photo_keys:
                continue
            if not INCLUDE_EXAMPLE_PHOTOS.get(photo_key, False):
                continue

            photo_names = EXAMPLE_PHOTO_NAMES.get(photo_key)
            if not photo_names:
                continue

            before_name, after_name = photo_names
            print(f"  Fetching example photos for {photo_key}...")
            before_img = fetch_sharepoint_photo(graph_token, before_name)
            after_img  = fetch_sharepoint_photo(graph_token, after_name)
            # Build label: include any services sharing this photo set
            sharers = [photo_key] + [
                svc for svc, donor in DEFER_EXAMPLE_PHOTOS_TO.items()
                if donor == photo_key and svc in upsell_map
            ]
            example_label = " and ".join(sharers) + " Example Photo"
            story.extend(build_example_photo_block(before_img, after_img, label=example_label))
            rendered_photo_keys.add(photo_key)

    # ── Build PDF ─────────────────────────────────────────────
    doc = SimpleDocTemplate(
        output_path,
        pagesize=letter,
        leftMargin=MARGIN,
        rightMargin=MARGIN,
        topMargin=MARGIN + SMALL_HDR_H,         # room for small header
        bottomMargin=MARGIN + FOOTER_HEIGHT,    # room for footer
    )
    doc.build(story, onFirstPage=on_first_page, onLaterPages=on_later_pages)
    print(f"\nPDF saved: {output_path}")
    return os.path.abspath(output_path)


# ─────────────────────────────────────────────────────────────
# Smoke test
# ─────────────────────────────────────────────────────────────
if __name__ == "__main__":
    import numpy as np
    from PIL import Image as PILImage

    def _fake_photo(color=(180, 180, 200)):
        return PILImage.fromarray(
            np.full((300, 400, 3), color, dtype=np.uint8)
        )

    fake_data = {
        "tail": "N253SY", "make": "Cessna", "model": "Citation X",
        "customer": "Duncan Aviation",
        "indoc_date": "2026-04-17", "rts_date": "2026-04-30",
        "included_services": ["Brightwork"],
        "upsells": [
            {"service": "Brightwork",      "price": "1800",  "notes": "moderate oxidation on leading edges",
             "photos": [_fake_photo((210,210,200)), _fake_photo((200,205,210))]},
            {"service": "Xylon",           "price": "1200",  "notes": "no current protection on leading edges",
             "photos": []},
            {"service": "Ceramic Coating", "price": "26760", "notes": "minor checking on crown",
             "photos": [_fake_photo((230,225,220))]},
            {"service": "Interior Detail", "price": "850",   "notes": "seats dirty, carpet stained",
             "photos": [_fake_photo((200,200,220)), _fake_photo((210,190,200))]},
            {"service": "Carpet Extraction","price": "420",  "notes": "heavy traffic wear",
             "photos": [_fake_photo((190,210,190))]},
        ],
    }

    import generate_upsell_pdf as _self
    _self.rewrite_notes          = lambda n, s, p: f"[AI: {n or 'visual observation'}]"
    _self.fetch_sharepoint_photo = lambda t, f: _fake_photo((220, 220, 180))

    print("Running smoke test...")
    out = generate_pdf(fake_data, graph_token="FAKE", output_path="/tmp/upsell_test.pdf")
    print(f"Done -> {out}")
