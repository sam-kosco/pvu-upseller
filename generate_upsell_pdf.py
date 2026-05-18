"""
PVU Upsell PDF Generator
========================
Takes the structured dict from parse_payload.py and produces a
polished sales PDF matching the Foxtrot Aviation example document style.

Pipeline:
  1. AI rewrites Stephen's raw notes into professional client-facing prose
  2. Marketing before/after photos fetched from SharePoint via Graph API
  3. ReportLab assembles single-column PDF with logo header, per-service
     sections (boilerplate + rewritten notes + condition photos + B/A photos),
     and Stephen's contact footer.

Usage:
    from generate_upsell_pdf import generate_pdf
    output_path = generate_pdf(parsed_data, graph_token, output_path="upsell.pdf")

Requirements:
    pip install reportlab pillow requests anthropic
"""

import io
import os
import requests
import anthropic
from PIL import Image
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Image as RLImage,
    Table, TableStyle, HRFlowable
)
from reportlab.lib.enums import TA_LEFT, TA_CENTER

# ─────────────────────────────────────────────────────────────
# CONFIG — toggle before/after marketing photos per service
# ─────────────────────────────────────────────────────────────
INCLUDE_EXAMPLE_PHOTOS = {
    "Brightwork":        True,
    "Ceramic Coating":   True,
    "Permagard Coating": True,
    "Polymer Coating":   True,
    "Interior Detail":   False,   # no example photos yet
    "Exterior Detail":   True,
    "Carpet Extraction": True,
    "Xylon":             False,   # no example photos yet
}

# ─────────────────────────────────────────────────────────────
# SharePoint config
# ─────────────────────────────────────────────────────────────
DRIVE_ID      = "b!_bzXaIx86kOufgJN3ih-BaDIDthKYuxJkJtLi1Bm5irGjCEnK-VHSpBRRm3_SDKU"
EXAMPLE_PHOTO_BASE = "Assets/Service Example Photos"

# Photo filenames match SharePoint exactly (service name + Before/After.jpg)
EXAMPLE_PHOTO_NAMES = {
    "Brightwork":        ("Brightwork Before.jpg",        "Brightwork After.jpg"),
    "Ceramic Coating":   ("Ceramic Coating Before.jpg",   "Ceramic Coating After.jpg"),
    "Permagard Coating": ("Permagard Before.jpg",         "Permagard After.jpg"),
    "Polymer Coating":   ("Polymer Before.jpg",           "Polymer After.jpg"),
    "Interior Detail":   None,
    "Exterior Detail":   None,
    "Carpet Extraction": ("Carpet Extraction Before.jpg", "Carpet Extraction After.jpg"),
    "Xylon":             ("Xylon Before.jpg",             "Xylon After.jpg"),
}

# ─────────────────────────────────────────────────────────────
# Boilerplate marketing text per service
# Replace placeholder strings with Stephen's actual copy later.
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
STEPHEN_NAME    = "Stephen Chadbourn"
STEPHEN_PHONE   = "C. 520.981.8942"
STEPHEN_EMAIL   = "E. stephen.Chadbourn@FoxtrotAviation.com"
STEPHEN_INTRO   = (
    "Hello, I'm Stephen Chadbourn, General Manager of the Foxtrot Aviation team at "
    "Duncan Aviation in Provo, Utah. We provide aircraft detailing during maintenance "
    "events and are committed to returning your aircraft in exceptional condition."
)


# ─────────────────────────────────────────────────────────────
# AI note rewriter
# ─────────────────────────────────────────────────────────────
def rewrite_notes(raw_notes: str, service_name: str) -> str:
    """
    Calls the Claude API to rewrite Stephen's casual field notes into
    professional, client-facing sales prose (2-3 sentences).
    Returns raw_notes unchanged if the API call fails.
    """
    if not raw_notes or not raw_notes.strip():
        return ""

    try:
        client = anthropic.Anthropic()
        response = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=300,
            messages=[{
                "role": "user",
                "content": (
                    f"Rewrite the following technician field note into professional, "
                    f"client-facing sales language for an aircraft detailing proposal. "
                    f"The service is: {service_name}. "
                    f"Keep it concise (2-3 sentences), warm, and specific to the observation. "
                    f"Do not add information not present in the note. "
                    f"Output only the rewritten text, no preamble.\n\n"
                    f"Raw note: {raw_notes}"
                )
            }]
        )
        return response.content[0].text.strip()
    except Exception as e:
        print(f"  Warning: AI rewrite failed for {service_name}: {e}")
        return raw_notes


# ─────────────────────────────────────────────────────────────
# SharePoint photo fetcher
# ─────────────────────────────────────────────────────────────
def fetch_sharepoint_photo(graph_token: str, filename: str) -> Image.Image | None:
    """
    Fetches a single photo from the Service Example Photos folder on SharePoint.
    Returns a PIL Image or None on failure.
    """
    path = f"{EXAMPLE_PHOTO_BASE}/{filename}"
    url  = (
        f"https://graph.microsoft.com/v1.0/drives/{DRIVE_ID}/root:/{path}:/content"
    )
    headers = {"Authorization": f"Bearer {graph_token}"}

    try:
        resp = requests.get(url, headers=headers, timeout=30)
        if resp.status_code == 200:
            return Image.open(io.BytesIO(resp.content))
        else:
            print(f"  Warning: could not fetch {filename} (HTTP {resp.status_code})")
            return None
    except Exception as e:
        print(f"  Warning: error fetching {filename}: {e}")
        return None


# ─────────────────────────────────────────────────────────────
# PIL Image → ReportLab Image helper
# ─────────────────────────────────────────────────────────────
def pil_to_rl(pil_img: Image.Image, max_width: float, max_height: float) -> RLImage:
    """Convert a PIL Image to a ReportLab Image flowable, scaled to fit."""
    buf = io.BytesIO()
    rgb = pil_img.convert("RGB")
    rgb.save(buf, format="JPEG", quality=85)
    buf.seek(0)

    w, h = pil_img.size
    scale = min(max_width / w, max_height / h, 1.0)
    return RLImage(buf, width=w * scale, height=h * scale)


# ─────────────────────────────────────────────────────────────
# Photo block builders
# ─────────────────────────────────────────────────────────────
PAGE_WIDTH   = letter[0]
MARGIN       = 0.75 * inch
CONTENT_W    = PAGE_WIDTH - 2 * MARGIN   # ~7 inches

def build_condition_photo_block(pil_images: list) -> list:
    """
    Builds a single-column or two-column condition photo block
    from a list of PIL Images. Returns a list of ReportLab flowables.
    """
    if not pil_images:
        return []

    flowables = []
    n = len(pil_images)

    if n == 1:
        flowables.append(pil_to_rl(pil_images[0], CONTENT_W, 3.5 * inch))
        flowables.append(Spacer(1, 8))
    else:
        # Two-column grid
        col_w = (CONTENT_W - 6) / 2
        rows  = []
        for i in range(0, n, 2):
            left  = pil_to_rl(pil_images[i],   col_w, 2.5 * inch)
            right = pil_to_rl(pil_images[i+1], col_w, 2.5 * inch) if i+1 < n else ""
            rows.append([left, right])

        tbl = Table(rows, colWidths=[col_w, col_w])
        tbl.setStyle(TableStyle([
            ("VALIGN",    (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING",  (0, 0), (-1, -1), 3),
            ("RIGHTPADDING", (0, 0), (-1, -1), 3),
            ("TOPPADDING",   (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING",(0, 0), (-1, -1), 3),
        ]))
        flowables.append(tbl)
        flowables.append(Spacer(1, 8))

    return flowables


def build_example_photo_block(before_img, after_img) -> list:
    """
    Builds a labelled Before / After two-column block.
    before_img and after_img are PIL Images (may be None).
    """
    if before_img is None and after_img is None:
        return []

    styles  = getSampleStyleSheet()
    label_s = ParagraphStyle("label", parent=styles["Normal"],
                             fontSize=9, alignment=TA_CENTER,
                             textColor=colors.HexColor("#444444"))

    col_w = (CONTENT_W - 6) / 2

    def cell(pil_img, label):
        if pil_img is None:
            return ""
        rl_img = pil_to_rl(pil_img, col_w, 2.5 * inch)
        return [rl_img, Paragraph(f"<b>{label}</b>", label_s)]

    row = [cell(before_img, "BEFORE"), cell(after_img, "AFTER")]
    tbl = Table([row], colWidths=[col_w, col_w])
    tbl.setStyle(TableStyle([
        ("VALIGN",    (0, 0), (-1, -1), "TOP"),
        ("ALIGN",     (0, 0), (-1, -1), "CENTER"),
        ("LEFTPADDING",  (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING",   (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 3),
    ]))
    return [tbl, Spacer(1, 8)]


# ─────────────────────────────────────────────────────────────
# Main PDF generator
# ─────────────────────────────────────────────────────────────
def generate_pdf(data: dict, graph_token: str, output_path: str = "upsell.pdf") -> str:
    """
    Generate the upsell sales PDF.

    Parameters
    ----------
    data        : dict from parse_payload.parse_payload()
    graph_token : valid Microsoft Graph OAuth2 bearer token
    output_path : where to save the PDF

    Returns
    -------
    str : absolute path to the saved PDF
    """

    styles = getSampleStyleSheet()

    # Custom styles
    title_s = ParagraphStyle("title_s", parent=styles["Normal"],
                             fontSize=12, leading=16, spaceAfter=2,
                             fontName="Helvetica-Bold")
    body_s  = ParagraphStyle("body_s",  parent=styles["Normal"],
                             fontSize=10, leading=14, spaceAfter=6)
    head_s  = ParagraphStyle("head_s",  parent=styles["Heading2"],
                             fontSize=11, leading=14, spaceAfter=4,
                             fontName="Helvetica-Bold")
    small_s = ParagraphStyle("small_s", parent=styles["Normal"],
                             fontSize=9,  leading=13)

    story = []

    # ── Logo ──────────────────────────────────────────────────
    logo_path = "/mnt/project/Fox_Logo_Red_background.png"
    if os.path.exists(logo_path):
        logo = RLImage(logo_path, width=1.5*inch, height=1.1*inch)
        story.append(logo)
        story.append(Spacer(1, 6))

    # ── Aircraft header ───────────────────────────────────────
    make_model = " ".join(filter(None, [data.get("make",""), data.get("model","")]))
    story.append(Paragraph(f"<b>Tail:</b> {data.get('tail','')}", title_s))
    if make_model:
        story.append(Paragraph(f"<b>Model:</b> {make_model}", title_s))
    story.append(Paragraph(f"<b>Owner:</b> {data.get('customer','')}", title_s))
    story.append(Spacer(1, 10))

    # ── Stephen's intro ───────────────────────────────────────
    story.append(Paragraph(STEPHEN_INTRO, body_s))
    story.append(Spacer(1, 8))

    # ── Upsell service sections ───────────────────────────────
    for upsell in data.get("upsells", []):
        service = upsell["service"]
        price   = upsell.get("price", "")
        notes   = upsell.get("notes", "")
        photos  = upsell.get("photos", [])   # list of PIL Images

        # Section divider
        story.append(HRFlowable(width="100%", thickness=0.5,
                                color=colors.HexColor("#cccccc"), spaceAfter=6))

        # Section header: "Service Name: $Price"
        price_str = f": ${price}" if price else ""
        story.append(Paragraph(f"{service}{price_str}", head_s))

        # Boilerplate marketing text
        boilerplate = SERVICE_BOILERPLATE.get(service, "")
        if boilerplate:
            story.append(Paragraph(boilerplate, body_s))

        # AI-rewritten condition notes
        if notes:
            print(f"  Rewriting notes for {service}...")
            polished = rewrite_notes(notes, service)
            if polished:
                story.append(Paragraph(polished, body_s))

        # Condition photos from JotForm (after text)
        if photos:
            story.extend(build_condition_photo_block(photos))

        # Marketing before/after photos from SharePoint
        if INCLUDE_EXAMPLE_PHOTOS.get(service, False):
            photo_names = EXAMPLE_PHOTO_NAMES.get(service)
            if photo_names:
                before_name, after_name = photo_names
                print(f"  Fetching example photos for {service}...")
                before_img = fetch_sharepoint_photo(graph_token, before_name)
                after_img  = fetch_sharepoint_photo(graph_token, after_name)
                story.extend(build_example_photo_block(before_img, after_img))

        story.append(Spacer(1, 6))

    # ── Footer: Stephen's contact info ────────────────────────
    story.append(HRFlowable(width="100%", thickness=0.5,
                            color=colors.HexColor("#cccccc"), spaceAfter=8))
    story.append(Paragraph(STEPHEN_NAME,  body_s))
    story.append(Paragraph(STEPHEN_PHONE, body_s))
    story.append(Paragraph(STEPHEN_EMAIL, body_s))

    # ── Build PDF ─────────────────────────────────────────────
    doc = SimpleDocTemplate(
        output_path,
        pagesize=letter,
        leftMargin=MARGIN, rightMargin=MARGIN,
        topMargin=0.75*inch, bottomMargin=0.75*inch,
    )
    doc.build(story)
    print(f"\nPDF saved: {output_path}")
    return os.path.abspath(output_path)


# ─────────────────────────────────────────────────────────────
# Smoke test (no SharePoint, no AI — tests layout only)
# ─────────────────────────────────────────────────────────────
if __name__ == "__main__":
    from PIL import Image as PILImage
    import numpy as np

    def _fake_photo(color=(180, 180, 200)):
        """Creates a tiny solid-color PIL image for layout testing."""
        arr = np.full((300, 400, 3), color, dtype=np.uint8)
        return PILImage.fromarray(arr)

    fake_data = {
        "tail":              "N253SY",
        "make":              "Cessna",
        "model":             "Citation X",
        "customer":          "Duncan Aviation",
        "indoc_date":        "2026-04-17",
        "rts_date":          "2026-04-30",
        "included_services": ["Brightwork", "Ceramic Coating"],
        "upsells": [
            {
                "service": "Interior Detail",
                "price":   "850",
                "notes":   "seats look a bit dirty, some stains on carpet",
                "photos":  [_fake_photo((200,200,220)), _fake_photo((210,190,200))],
            },
            {
                "service": "Carpet Extraction",
                "price":   "1280",
                "notes":   "carpet is really worn, lots of dirt tracked in",
                "photos":  [_fake_photo((190,210,190))],
            },
        ],
    }

    print("Running smoke test (no SharePoint, no AI)...")
    print("  Skipping AI rewrite and SharePoint fetches in smoke test.")

    # Temporarily disable AI + SharePoint for the smoke test
    import generate_upsell_pdf as _self
    _orig_rewrite = _self.rewrite_notes
    _orig_fetch   = _self.fetch_sharepoint_photo
    _self.rewrite_notes         = lambda notes, svc: f"[Rewritten] {notes}"
    _self.fetch_sharepoint_photo = lambda token, fname: _fake_photo((220,220,180))

    out = generate_pdf(fake_data, graph_token="FAKE", output_path="/tmp/upsell_test.pdf")

    _self.rewrite_notes         = _orig_rewrite
    _self.fetch_sharepoint_photo = _orig_fetch

    print(f"Smoke test complete → {out}")
