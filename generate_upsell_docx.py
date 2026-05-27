"""
PVU Upsell Word Document Generator
====================================
Generates a .docx version of the upsell proposal alongside the PDF.
Stephen can open it in Word and make any last-minute edits before
sending it to the customer.

Content mirrors generate_upsell_pdf.py exactly:
  - Logo on page 1
  - Small header (Foxtrot Aviation Services | Duncan Aviation PVU | date) every page
  - Aircraft info block
  - Stephen's intro
  - Three supersections with per-service blocks
  - Condition photos (from JotForm) with "[Tail] Current Condition" caption
  - Example photos (from SharePoint) with "[Service] Example Photo" caption
  - Footer: Stephen's name | phone | email every page

Usage:
    from generate_upsell_docx import generate_docx
    output_path = generate_docx(parsed_data, graph_token, output_path="upsell.docx")

Requirements:
    pip install pillow requests
    npm install -g docx
"""

import io
import json
import os
import subprocess
import tempfile
from datetime import date

import requests
from PIL import Image

# ── Re-use all config from generate_upsell_pdf ─────────────────────────────
from generate_upsell_pdf import (
    SUPERSECTIONS,
    INCLUDE_EXAMPLE_PHOTOS,
    EXAMPLE_PHOTO_NAMES,
    SERVICE_BOILERPLATE,
    DEFER_EXAMPLE_PHOTOS_TO,
    DISPLAY_NAMES,
    SUPERSECTION_EXAMPLE_PHOTOS,
    STEPHEN_NAME,
    STEPHEN_PHONE,
    STEPHEN_EMAIL,
    STEPHEN_INTRO,
    SERVICE_CONTEXT,
    DRIVE_ID,
    EXAMPLE_PHOTO_BASE,
    rewrite_notes,
    fetch_sharepoint_photo,
    _find_logo,
)

LOGO_PATH = _find_logo()


# ─────────────────────────────────────────────────────────────────────────────
# PIL image → temp JPEG path (docx-js needs a file path, not bytes)
# ─────────────────────────────────────────────────────────────────────────────
def _save_temp_image(pil_img: Image.Image, suffix: str = ".jpg") -> str:
    """Save a PIL image to a temp file and return the path."""
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    pil_img.convert("RGB").save(tmp.name, "JPEG", quality=85)
    tmp.close()
    return tmp.name


# ─────────────────────────────────────────────────────────────────────────────
# Build the JS data payload for the docx-js script
# ─────────────────────────────────────────────────────────────────────────────
def _build_doc_data(data: dict, graph_token: str) -> dict:
    """
    Collect all content into a plain Python dict that gets JSON-serialised
    and passed to the Node.js docx-builder script.

    Images are saved to temp files; the dict carries their paths.
    """
    today     = date.today().strftime("%m/%d/%Y")
    tail      = data.get("tail", "")
    owner     = data.get("owner", "").strip() or data.get("customer", "").strip()
    make_model = " ".join(filter(None, [data.get("make", ""), data.get("model", "")]))

    upsell_map = {u["service"]: u for u in data.get("upsells", [])}

    doc_data = {
        "today":       today,
        "tail":        tail,
        "makeModel":   make_model,
        "owner":       owner,
        "intro":       STEPHEN_INTRO,
        "footerName":  STEPHEN_NAME,
        "footerPhone": STEPHEN_PHONE,
        "footerEmail": STEPHEN_EMAIL,
        "logoPath":    LOGO_PATH or "",
        "sections":    [],
        "tempFiles":   [],   # collect for cleanup
    }

    for section in SUPERSECTIONS:
        section_upsells = [upsell_map[s] for s in section["services"] if s in upsell_map]
        if not section_upsells:
            continue

        sec = {
            "title":      section["title"],
            "boilerplate": section["boilerplate"],
            "services":   [],
            "sectionExamplePhotos": [],
            "sectionExampleCaption": "",
        }

        # ── Supersection-level example photos ──────────────────────────
        ss_photos = SUPERSECTION_EXAMPLE_PHOTOS.get(section["title"])
        if ss_photos:
            before_name, after_name = ss_photos
            print(f"  [docx] Fetching supersection example photos for {section['title']}...")
            before_img = fetch_sharepoint_photo(graph_token, before_name)
            after_img  = fetch_sharepoint_photo(graph_token, after_name)
            before_path = _save_temp_image(before_img) if before_img else ""
            after_path  = _save_temp_image(after_img)  if after_img  else ""
            if before_path:
                doc_data["tempFiles"].append(before_path)
            if after_path:
                doc_data["tempFiles"].append(after_path)
            sec["sectionExamplePhotos"] = [before_path, after_path]
            sec["sectionExampleCaption"] = f"{section['title']} Example Photo"

        # ── Pass 1: service text blocks + condition photos ────────────────
        for upsell in section_upsells:
            service = upsell["service"]
            display = DISPLAY_NAMES.get(service, service)
            raw_price = upsell.get("price", "")
            price = f"{int(raw_price):,}" if raw_price and raw_price.isdigit() else raw_price
            notes   = upsell.get("notes", "")
            photos  = upsell.get("photos", [])   # PIL Images

            # AI rewrite
            polished = ""
            if notes or photos:
                print(f"  [docx] Rewriting notes for {display} "
                      f"({'with' if photos else 'without'} photos)...")
                polished = rewrite_notes(notes, service, photos)

            # Save condition photos to temp files
            condition_photo_paths = []
            for img in photos:
                p = _save_temp_image(img)
                condition_photo_paths.append(p)
                doc_data["tempFiles"].append(p)

            svc_data = {
                "name":              display,
                "price":             price,
                "polishedNotes":     polished,
                "conditionPhotos":   condition_photo_paths,
                "conditionCaption":  f"{tail} Current Condition" if tail else "Current Condition",
                "examplePhotos":     [],   # filled in pass 2
                "exampleCaption":    "",
            }
            sec["services"].append(svc_data)

        # ── Pass 2: example photos after all text ─────────────────────────
        rendered_photo_keys: set = set()
        for i, upsell in enumerate(section_upsells):
            service   = upsell["service"]
            photo_key = DEFER_EXAMPLE_PHOTOS_TO.get(service, service)

            if photo_key in rendered_photo_keys:
                continue
            if not INCLUDE_EXAMPLE_PHOTOS.get(photo_key, False):
                continue

            photo_names = EXAMPLE_PHOTO_NAMES.get(photo_key)
            if not photo_names:
                continue

            before_name, after_name = photo_names
            print(f"  [docx] Fetching example photos for {photo_key}...")
            before_img = fetch_sharepoint_photo(graph_token, before_name)
            after_img  = fetch_sharepoint_photo(graph_token, after_name)

            before_path = _save_temp_image(before_img) if before_img else ""
            after_path  = _save_temp_image(after_img)  if after_img  else ""
            if before_path:
                doc_data["tempFiles"].append(before_path)
            if after_path:
                doc_data["tempFiles"].append(after_path)

            # Build label same as PDF
            sharers = [photo_key] + [
                svc for svc, donor in DEFER_EXAMPLE_PHOTOS_TO.items()
                if donor == photo_key and svc in upsell_map
            ]
            example_label = " and ".join(sharers) + " Example Photo"

            # Attach to the LAST service in this section (renders after all text)
            sec["services"][-1]["examplePhotos"] = [before_path, after_path]
            sec["services"][-1]["exampleCaption"] = example_label
            rendered_photo_keys.add(photo_key)

        doc_data["sections"].append(sec)

    return doc_data


# ─────────────────────────────────────────────────────────────────────────────
# Node.js docx builder script
# ─────────────────────────────────────────────────────────────────────────────
DOCX_BUILDER_JS = r"""
const fs   = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  ImageRun, Header, Footer, AlignmentType, HeadingLevel, BorderStyle,
  WidthType, VerticalAlign, PageNumber, ShadingType,
} = require('docx');

const data       = JSON.parse(fs.readFileSync(process.argv[2], 'utf8'));
const outputPath = process.argv[3];

// ── Helpers ────────────────────────────────────────────────────────────────
const GRAY    = "555555";
const DGRAY   = "1a1a1a";
const LGRAY   = "888888";
const CAPGRAY = "333333";

// US Letter with 0.75" margins → content width = 12240 - 2*1080 = 10080 DXA
const MARGIN     = 1080;   // 0.75 inch
const CONTENT_W  = 10080;  // DXA
const HALF_W     = 4980;   // ~half minus gutter
const GUTTER     = 120;

function hRule() {
  return new Paragraph({
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: LGRAY, space: 1 } },
    spacing: { after: 80 },
    children: [],
  });
}

function thinRule() {
  return new Paragraph({
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "cccccc", space: 1 } },
    spacing: { after: 60 },
    children: [],
  });
}

function bodyPara(text, opts = {}) {
  if (!text) return null;
  return new Paragraph({
    spacing: { after: 120 },
    children: [new TextRun({ text, font: "Arial", size: 20,
      color: opts.color || "000000", bold: opts.bold || false })],
  });
}

function captionPara(text) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 60 },
    children: [new TextRun({ text, font: "Arial", size: 16,
      bold: true, color: CAPGRAY })],
  });
}

function imageCell(imgPath, label) {
  const children = [];
  if (imgPath && fs.existsSync(imgPath)) {
    const buf  = fs.readFileSync(imgPath);
    const ext  = path.extname(imgPath).replace('.','').toLowerCase();
    const type = ext === 'jpg' ? 'jpeg' : ext;
    // max ~3.3" wide × 2.3" tall in half-column
    children.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new ImageRun({
        data: buf, type,
        transformation: { width: 310, height: 210 },
      })],
    }));
  }
  if (label) {
    children.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: label, font: "Arial",
        size: 16, bold: true, color: CAPGRAY })],
    }));
  }
  return new TableCell({
    width: { size: HALF_W, type: WidthType.DXA },
    margins: { top: 80, bottom: 80, left: GUTTER, right: GUTTER },
    borders: {
      top:    { style: BorderStyle.NONE },
      bottom: { style: BorderStyle.NONE },
      left:   { style: BorderStyle.NONE },
      right:  { style: BorderStyle.NONE },
    },
    children,
  });
}

function photoTable(leftPath, rightPath, leftLabel, rightLabel) {
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [HALF_W, HALF_W],
    rows: [new TableRow({ children: [
      imageCell(leftPath,  leftLabel),
      imageCell(rightPath, rightLabel),
    ]})],
  });
}

function conditionPhotoBlock(paths, caption) {
  const items = [];
  if (!paths || paths.length === 0) return items;
  items.push(captionPara(caption));
  if (paths.length === 1) {
    const buf = fs.existsSync(paths[0]) ? fs.readFileSync(paths[0]) : null;
    if (buf) {
      const ext  = path.extname(paths[0]).replace('.','').toLowerCase();
      const type = ext === 'jpg' ? 'jpeg' : ext;
      items.push(new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 120 },
        children: [new ImageRun({
          data: buf, type,
          transformation: { width: 580, height: 380 },
        })],
      }));
    }
  } else {
    // Pair up in rows of 2
    for (let i = 0; i < paths.length; i += 2) {
      items.push(photoTable(paths[i], paths[i+1] || '', '', ''));
      items.push(new Paragraph({ spacing: { after: 60 }, children: [] }));
    }
  }
  return items;
}

function examplePhotoBlock(paths, caption) {
  const items = [];
  if (!paths || paths.filter(Boolean).length === 0) return items;
  items.push(captionPara(caption));
  items.push(photoTable(paths[0] || '', paths[1] || '', 'BEFORE', 'AFTER'));
  items.push(new Paragraph({ spacing: { after: 120 }, children: [] }));
  return items;
}

// ── Page header (all pages — logo is in body, not here) ───────────────────
function makeHeader() {
  const hdrText = `Foxtrot Aviation Services   |   Duncan Aviation PVU   |   ${data.today}`;
  return new Header({
    children: [new Paragraph({
      alignment: AlignmentType.RIGHT,
      border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "dddddd", space: 4 } },
      spacing: { after: 120 },
      children: [new TextRun({ text: hdrText, font: "Arial", size: 15, color: "777777" })],
    })],
  });
}

// ── Page footer ────────────────────────────────────────────────────────────
function makeFooter() {
  const footerText = `${data.footerName}   |   ${data.footerPhone}   |   ${data.footerEmail}`;
  return new Footer({
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      border: { top: { style: BorderStyle.SINGLE, size: 4, color: "cccccc", space: 4 } },
      spacing: { before: 80 },
      children: [new TextRun({ text: footerText, font: "Arial", size: 16, color: GRAY })],
    })],
  });
}

// ── Build body ─────────────────────────────────────────────────────────────
const body = [];

// Logo — centered at top of document (page 1 body, not header)
if (data.logoPath && fs.existsSync(data.logoPath)) {
  const logoBuf = fs.readFileSync(data.logoPath);
  body.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 160 },
    children: [new ImageRun({
      data: logoBuf, type: 'png',
      transformation: { width: 252, height: 114 },  // ~1.75" x 0.79"
    })],
  }));
}

// Aircraft header
const infoLines = [
  `Tail: ${data.tail}`,
  data.makeModel ? `Model: ${data.makeModel}` : null,
  data.owner     ? `Owner: ${data.owner}`     : null,
].filter(Boolean);

for (const line of infoLines) {
  body.push(new Paragraph({
    spacing: { after: 60 },
    children: [new TextRun({ text: line, font: "Arial", size: 24, bold: true })],
  }));
}
body.push(new Paragraph({ spacing: { after: 160 }, children: [] }));
body.push(bodyPara(data.intro));
body.push(new Paragraph({ spacing: { after: 120 }, children: [] }));

// Supersections
for (const section of data.sections) {
  body.push(hRule());
  body.push(new Paragraph({
    spacing: { after: 80 },
    children: [new TextRun({ text: section.title, font: "Arial",
      size: 26, bold: true, color: DGRAY })],
  }));
  const bp = bodyPara(section.boilerplate);
  if (bp) body.push(bp);

  // Supersection-level example photos
  for (const el of examplePhotoBlock(section.sectionExamplePhotos, section.sectionExampleCaption)) {
    body.push(el);
  }

  body.push(new Paragraph({ spacing: { after: 100 }, children: [] }));

  for (const svc of section.services) {
    body.push(thinRule());

    // Service heading with price
    const headText = svc.price ? `${svc.name}: $${svc.price}` : svc.name;
    body.push(new Paragraph({
      spacing: { after: 80 },
      children: [new TextRun({ text: headText, font: "Arial", size: 22, bold: true })],
    }));

    // AI-rewritten notes
    const nbp = bodyPara(svc.polishedNotes);
    if (nbp) body.push(nbp);

    // Condition photos (JotForm)
    for (const el of conditionPhotoBlock(svc.conditionPhotos, svc.conditionCaption)) {
      body.push(el);
    }

    // Example photos (SharePoint) — after all text
    for (const el of examplePhotoBlock(svc.examplePhotos, svc.exampleCaption)) {
      body.push(el);
    }

    body.push(new Paragraph({ spacing: { after: 100 }, children: [] }));
  }
}

// ── Assemble document ──────────────────────────────────────────────────────
const doc = new Document({
  styles: {
    default: { document: { run: { font: "Arial", size: 20 } } },
  },
  sections: [{
    properties: {
      page: {
        size:   { width: 12240, height: 15840 },
        margin: { top: MARGIN, right: MARGIN, bottom: MARGIN + 360, left: MARGIN },
      },
    },
    headers: {
      default: makeHeader(),
    },
    footers: {
      default: makeFooter(),
    },
    children: body,
  }],
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(outputPath, buf);
  console.log('docx written:', outputPath);
});
"""


# ─────────────────────────────────────────────────────────────────────────────
# Main entry point
# ─────────────────────────────────────────────────────────────────────────────
def generate_docx(data: dict, graph_token: str,
                  output_path: str = "upsell.docx") -> str:
    """
    Generate the upsell Word document.

    Parameters
    ----------
    data        : dict from parse_payload.parse_payload()
    graph_token : valid Microsoft Graph OAuth2 bearer token
    output_path : where to save the .docx

    Returns
    -------
    str : absolute path to the saved .docx
    """
    # 1. Build data payload (AI rewrites + photo temp files)
    doc_data = _build_doc_data(data, graph_token)

    # 2. Write the JS builder script to a temp file
    js_path   = tempfile.NamedTemporaryFile(delete=False, suffix=".js")
    js_path.write(DOCX_BUILDER_JS.encode())
    js_path.close()
    doc_data["tempFiles"].append(js_path.name)

    # 3. Write the JSON data payload to a temp file
    json_path = tempfile.NamedTemporaryFile(delete=False, suffix=".json",
                                            mode="w", encoding="utf-8")
    # Remove tempFiles list from the JSON (no need to pass it to Node)
    payload = {k: v for k, v in doc_data.items() if k != "tempFiles"}
    json.dump(payload, json_path)
    json_path.close()
    doc_data["tempFiles"].append(json_path.name)

    # 4. Run the Node.js builder
    abs_output = os.path.abspath(output_path)
    try:
        result = subprocess.run(
            ["node", js_path.name, json_path.name, abs_output],
            capture_output=True, text=True, timeout=120,
        )
        if result.returncode != 0:
            raise RuntimeError(
                f"docx builder failed:\nSTDOUT: {result.stdout}\nSTDERR: {result.stderr}"
            )
        print(f"Docx saved: {abs_output}")
    finally:
        # 5. Clean up temp files
        for p in doc_data["tempFiles"]:
            try:
                os.unlink(p)
            except Exception:
                pass

    return abs_output
