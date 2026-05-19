"""
PVU Upsell Payload Parser
=========================
Parses the raw JotForm payload delivered by Power Automate and returns a
clean structured dict ready for PDF generation.

Usage:
    from parse_payload import parse_payload
    data = parse_payload(body)   # body = request.json["body"]

Output structure:
    {
        "tail":             "N253SY",
        "make":             "Cessna",
        "model":            "Citation X",
        "technician":       "Sam Kosco",
        "customer":         "Duncan",
        "indoc_date":       "2026-04-17",
        "rts_date":         "2026-04-30",
        "included_services": ["Brightwork", "Ceramic Coating"],
        "upsells": [
            {
                "service": "Interior Detail",
                "price":   "850",
                "notes":   "They will only want 1 coating",
                "photos":  [<PIL Image>, <PIL Image>, ...]   # only if photos present
            },
            ...
        ]
    }

Only services where Y/N == "Yes" appear in upsells[].
Included services come from q7 (semicolon-separated string).
"""

import base64
import io

from PIL import Image


# ─────────────────────────────────────────────────────────────
# Field map:  service_name -> (yn_q, photos_q, price_q, notes_q)
# ─────────────────────────────────────────────────────────────
SERVICE_MAP = {
    "Brightwork":       (10, 11, 26, 27),
    "Ceramic Coating":  (12, 13, 29, 30),
    "Permagard Coating":(14, 15, 32, 33),
    "Polymer Coating":  (16, 17, 35, 36),
    "Interior Detail":  (18, 19, 38, 39),
    "Exterior Detail":  (20, 21, 41, 42),
    "Carpet Extraction":(22, 23, 44, 45),
    "Xylon":            (49, 53, 50, 51),
}


def _get(body, q):
    """Return body[str(q)], or '' if missing."""
    return body.get(str(q), "")


def _extract_images(body, q):
    """
    Extract PIL Images from a JotForm file-upload field.
    Returns a list of PIL Image objects (may be empty).
    """
    entries = _get(body, q)
    if not isinstance(entries, list):
        return []

    images = []
    for entry in entries:
        b64 = entry.get("file", "")
        if b64:
            try:
                img_bytes = base64.b64decode(b64)
                img = Image.open(io.BytesIO(img_bytes))
                img.load()          # force decode now so we catch bad data early
                images.append(img)
            except Exception as e:
                print(f"  Warning: could not decode image in q{q}: {e}")
    return images


def parse_payload(body: dict) -> dict:
    """
    Parse the Power Automate / JotForm payload body and return structured data.

    Parameters
    ----------
    body : dict
        The 'body' key from the full Power Automate JSON payload.

    Returns
    -------
    dict
        Structured submission data (see module docstring).
    """

    # ── Header fields ──────────────────────────────────────────
    included_raw = _get(body, 7)
    if included_raw:
        included_services = [s.strip() for s in included_raw.split(";") if s.strip()]
    else:
        included_services = []

    result = {
        "technician":        _get(body, 3),
        "tail":              _get(body, 4).upper(),
        "indoc_date":        _get(body, 5),
        "rts_date":          _get(body, 6),
        "included_services": included_services,
        "customer":          _get(body, 8),
        "make":              _get(body, 47),
        "model":             _get(body, 48),
        "owner":             _get(body, 54),
        "upsells":           [],
    }

    # ── Upsell services ────────────────────────────────────────
    for service_name, (yn_q, photos_q, price_q, notes_q) in SERVICE_MAP.items():
        yn_value = _get(body, yn_q)

        # Only include services explicitly marked "Yes"
        if str(yn_value).strip().lower() != "yes":
            continue

        photos = _extract_images(body, photos_q)

        upsell = {
            "service": service_name,
            "price":   str(_get(body, price_q)).strip(),
            "notes":   str(_get(body, notes_q)).strip(),
            "photos":  photos,
        }
        result["upsells"].append(upsell)

    return result


# ─────────────────────────────────────────────────────────────
# Quick smoke test — run this file directly to validate parsing
# against a minimal fake payload.
# ─────────────────────────────────────────────────────────────
if __name__ == "__main__":
    fake_body = {
        "3":  "Sam Kosco",
        "4":  "N253SY",
        "5":  "2026-04-17",
        "6":  "2026-04-30",
        "7":  "Brightwork; Ceramic Coating",
        "8":  "Duncan",
        "47": "Cessna",
        "48": "Citation X",
        # Brightwork → No
        "10": "No",
        "11": [{"url": "", "file": "", "name": ""}],
        "26": "",
        "27": "",
        # Interior Detail → Yes
        "18": "Yes",
        "19": [{"url": "", "file": "", "name": ""}],
        "38": "850",
        "39": "They will only want 1 coating",
        # Carpet Extraction → Yes
        "22": "Yes",
        "23": [{"url": "", "file": "", "name": ""}],
        "44": "1280",
        "45": "Carpet is heavily soiled",
    }

    data = parse_payload(fake_body)

    print("=== HEADER ===")
    for k, v in data.items():
        if k != "upsells":
            print(f"  {k}: {v}")

    print(f"\n=== UPSELLS ({len(data['upsells'])}) ===")
    for u in data["upsells"]:
        print(f"  {u['service']}")
        print(f"    price:  {u['price']}")
        print(f"    notes:  {u['notes']}")
        print(f"    photos: {len(u['photos'])} image(s)")

    print("\nDone.")
