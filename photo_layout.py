"""
photo_layout.py
===============
Takes 1–5 images (JPEG/PNG, any size/orientation) and produces a single
formatted composite image.

Canvas sizing philosophy
------------------------
Each layout has a "natural" size derived from standard cell dimensions:
    Landscape cell : 420 × 236 px  (≈16:9)
    Portrait cell  : 280 × 373 px  (≈3:4)

The grid of cells (plus 8 px gutters) determines the natural canvas size.
If that natural size exceeds MAX_W (1275 px) or MAX_H (1650 px), it is
scaled down proportionally.  Layouts that are naturally compact (e.g. PPPP,
a single P) stay small rather than being stretched to fill the page.

This means the returned Image can be any size up to 1275 × 1650 px.
When embedding in a PDF, scale it to the desired print width and let the
height follow naturally.

Usage
-----
    from photo_layout import build_photo_block

    img = build_photo_block(["a.jpg", "b.png", "c.jpg"])
    img.save("block.png")

    build_photo_block(paths, output_path="block.png")

Orientation classification
--------------------------
    Landscape (L) : width >= height
    Portrait  (P) : width <  height
"""

from PIL import Image, ImageOps

# ── Constants ──────────────────────────────────────────────────────────────────
MAX_W  = 1275   # hard ceiling — never wider than a Letter page @ 150 dpi
MAX_H  = 1650   # hard ceiling — never taller than a full Letter page
GUTTER = 8
G      = GUTTER
BG     = (255, 255, 255)

# Base cell dimensions — the building blocks for natural canvas sizing
L_W, L_H = 420, 236    # landscape cell  ≈ 16:9
P_W, P_H = 280, 373    # portrait cell   ≈ 3:4


# ── Natural canvas sizes per layout ───────────────────────────────────────────
# (natural_w, natural_h) — derived from cell grid + gutters.
# Scaled down if either dimension exceeds MAX_W / MAX_H.

NATURAL = {
    # count 1
    "L":     (L_W,              L_H),
    "P":     (P_W,              P_H),
    # count 2
    "LL":    (2*L_W + G,        L_H),
    "LP":    (L_W + G + P_W,    max(L_H, P_H)),
    "PP":    (2*P_W + G,        P_H),
    # count 3
    "LLL":   (L_W + G + L_W,              L_H + G + L_H//2),
    "LLP":   (L_W + G + P_W,              2*L_H + G),
    "LPP":   (2*P_W + G,                  L_H + G + P_H),
    "PPP":   (3*P_W + 2*G,                P_H),
    # count 4
    "LLLL":  (2*L_W + G,                  2*L_H + G),
    "LLLP":  (L_W + G + P_W,              3*L_H + 2*G),
    "LLPP":  (2*L_W + G,                  L_H + G + P_H),
    "LPPP":  (3*P_W + 2*G,                L_H//2 + G + P_H),
    "PPPP":  (2*P_W + G,                  2*P_H + G),
    # count 5
    "LLLLL": (3*L_W + 2*G,                2*L_H + G),
    "LLLLP": (2*L_W + G + P_W,            2*L_H + G),
    "LLLPP": (3*L_W + 2*G,                L_H + G + P_H),
    "LLPPP": (3*P_W + 2*G,                L_H + G + P_H),
    "LPPPP": (L_W + G + 2*P_W,            2*P_H + G),
    "PPPPP": (3*P_W + 2*G,                2*P_H + G),
}


def _canvas_size(key):
    """Scale natural size down if it exceeds either hard ceiling."""
    nw, nh = NATURAL[key]
    scale  = min(MAX_W / nw, MAX_H / nh, 1.0)   # 1.0 = never scale UP
    return round(nw * scale), round(nh * scale)


# ── Core helpers ───────────────────────────────────────────────────────────────

def _orient(img):
    return "P" if img.height > img.width else "L"


def _fit(img, box_w, box_h):
    """Resize + centre-crop image to fill box exactly."""
    sw, sh = img.size
    scale  = max(box_w / sw, box_h / sh)
    nw     = max(1, round(sw * scale))
    nh     = max(1, round(sh * scale))
    img    = img.resize((nw, nh), Image.LANCZOS)
    left   = (nw - box_w) // 2
    top    = (nh - box_h) // 2
    return img.crop((left, top, left + box_w, top + box_h))


def _place(canvas, img, x, y, w, h):
    canvas.paste(_fit(img, w, h), (x, y))


def _load(paths):
    out = []
    for p in paths:
        im = Image.open(p).convert("RGB")
        try:
            im = ImageOps.exif_transpose(im)
        except Exception:
            pass
        out.append(im)
    return out


def _sort_LP(imgs):
    return ([i for i in imgs if _orient(i) == "L"],
            [i for i in imgs if _orient(i) == "P"])


# ── Grid helpers (take explicit W, H) ─────────────────────────────────────────

def _cols(n, W):
    avail = W - G * (n - 1)
    b     = avail // n
    ws    = [b] * n
    ws[-1] += avail - b * n
    return ws

def _rows(n, H):
    avail = H - G * (n - 1)
    b     = avail // n
    hs    = [b] * n
    hs[-1] += avail - b * n
    return hs

def _xs(widths):
    x = [0]
    for w in widths[:-1]:
        x.append(x[-1] + w + G)
    return x

def _ys(heights):
    y = [0]
    for h in heights[:-1]:
        y.append(y[-1] + h + G)
    return y


# ══════════════════════════════════════════════════════════════════════════════
# Layout functions — each receives (canvas, imgs, W, H)
# ══════════════════════════════════════════════════════════════════════════════

# ── Count 1 ───────────────────────────────────────────────────────────────────

def layout_L(c, imgs, W, H):
    _place(c, imgs[0], 0, 0, W, H)

def layout_P(c, imgs, W, H):
    # Fit to height, centre horizontally with white bars
    img   = imgs[0]
    scale = H / img.height
    fw    = min(round(img.width * scale), W)
    _place(c, img, (W - fw) // 2, 0, fw, H)


# ── Count 2 ───────────────────────────────────────────────────────────────────

def layout_LL(c, imgs, W, H):
    ws = _cols(2, W); xs = _xs(ws)
    for i in range(2):
        _place(c, imgs[i], xs[i], 0, ws[i], H)

def layout_LP(c, imgs, W, H):
    L, P = _sort_LP(imgs)
    # Width split proportional to natural cell widths
    lw = round(W * L_W / (L_W + G + P_W))
    pw = W - lw - G
    _place(c, L[0], 0,      0, lw, H)
    _place(c, P[0], lw + G, 0, pw, H)

def layout_PP(c, imgs, W, H):
    ws = _cols(2, W); xs = _xs(ws)
    for i in range(2):
        _place(c, imgs[i], xs[i], 0, ws[i], H)


# ── Count 3 ───────────────────────────────────────────────────────────────────

def layout_LLL(c, imgs, W, H):
    # Large left, two stacked right
    # Natural: left col = L_W, right col = L_W, top row = L_H, bot row = L_H//2
    nat_w, nat_h = NATURAL["LLL"]
    lw = round(W * L_W / nat_w)
    rw = W - lw - G
    top_h = round(H * L_H / nat_h)
    bot_h = H - top_h - G
    _place(c, imgs[0], 0,      0,         lw, H)
    _place(c, imgs[1], lw + G, 0,         rw, top_h)
    _place(c, imgs[2], lw + G, top_h + G, rw, bot_h)

def layout_LLP(c, imgs, W, H):
    L, P = _sort_LP(imgs)
    nat_w, _ = NATURAL["LLP"]
    lw = round(W * L_W / nat_w)
    pw = W - lw - G
    hs = _rows(2, H); ys = _ys(hs)
    _place(c, L[0], 0,      ys[0], lw, hs[0])
    _place(c, L[1], 0,      ys[1], lw, hs[1])
    _place(c, P[0], lw + G, 0,     pw, H)

def layout_LPP(c, imgs, W, H):
    L, P  = _sort_LP(imgs)
    nat_w, nat_h = NATURAL["LPP"]
    top_h = round(H * L_H / nat_h)
    bot_h = H - top_h - G
    ws = _cols(2, W); xs = _xs(ws)
    _place(c, L[0], 0,     0,         W,     top_h)
    _place(c, P[0], xs[0], top_h + G, ws[0], bot_h)
    _place(c, P[1], xs[1], top_h + G, ws[1], bot_h)

def layout_PPP(c, imgs, W, H):
    ws = _cols(3, W); xs = _xs(ws)
    for i in range(3):
        _place(c, imgs[i], xs[i], 0, ws[i], H)


# ── Count 4 ───────────────────────────────────────────────────────────────────

def layout_LLLL(c, imgs, W, H):
    ws = _cols(2, W); xs = _xs(ws)
    hs = _rows(2, H); ys = _ys(hs)
    idx = 0
    for r in range(2):
        for col in range(2):
            _place(c, imgs[idx], xs[col], ys[r], ws[col], hs[r])
            idx += 1

def layout_LLLP(c, imgs, W, H):
    L, P = _sort_LP(imgs)
    nat_w, _ = NATURAL["LLLP"]
    lw = round(W * L_W / nat_w)
    pw = W - lw - G
    hs = _rows(3, H); ys = _ys(hs)
    for i in range(3):
        _place(c, L[i], 0,      ys[i], lw, hs[i])
    _place(c, P[0], lw + G, 0,     pw, H)

def layout_LLPP(c, imgs, W, H):
    L, P = _sort_LP(imgs)
    nat_w, nat_h = NATURAL["LLPP"]
    top_h = round(H * L_H / nat_h)
    bot_h = H - top_h - G
    ws = _cols(2, W); xs = _xs(ws)
    _place(c, L[0], xs[0], 0,         ws[0], top_h)
    _place(c, L[1], xs[1], 0,         ws[1], top_h)
    _place(c, P[0], xs[0], top_h + G, ws[0], bot_h)
    _place(c, P[1], xs[1], top_h + G, ws[1], bot_h)

def layout_LPPP(c, imgs, W, H):
    L, P  = _sort_LP(imgs)
    nat_w, nat_h = NATURAL["LPPP"]
    top_h = round(H * (L_H // 2) / nat_h)
    bot_h = H - top_h - G
    ws = _cols(3, W); xs = _xs(ws)
    _place(c, L[0], 0, 0, W, top_h)
    for i in range(3):
        _place(c, P[i], xs[i], top_h + G, ws[i], bot_h)

def layout_PPPP(c, imgs, W, H):
    ws = _cols(2, W); xs = _xs(ws)
    hs = _rows(2, H); ys = _ys(hs)
    idx = 0
    for r in range(2):
        for col in range(2):
            _place(c, imgs[idx], xs[col], ys[r], ws[col], hs[r])
            idx += 1


# ── Count 5 ───────────────────────────────────────────────────────────────────

def layout_LLLLL(c, imgs, W, H):
    hs = _rows(2, H); ys = _ys(hs)
    ws2 = _cols(2, W); xs2 = _xs(ws2)
    ws3 = _cols(3, W); xs3 = _xs(ws3)
    for i in range(2):
        _place(c, imgs[i],     xs2[i], ys[0], ws2[i], hs[0])
    for i in range(3):
        _place(c, imgs[2 + i], xs3[i], ys[1], ws3[i], hs[1])

def layout_LLLLP(c, imgs, W, H):
    L, P = _sort_LP(imgs)
    nat_w, _ = NATURAL["LLLLP"]
    lw_total = round(W * (2*L_W + G) / nat_w)
    pw       = W - lw_total - G
    ws2 = _cols(2, lw_total); xs2 = _xs(ws2)
    hs  = _rows(2, H);        ys  = _ys(hs)
    idx = 0
    for r in range(2):
        for col in range(2):
            _place(c, L[idx], xs2[col], ys[r], ws2[col], hs[r])
            idx += 1
    _place(c, P[0], lw_total + G, 0, pw, H)

def layout_LLLPP(c, imgs, W, H):
    L, P  = _sort_LP(imgs)
    nat_w, nat_h = NATURAL["LLLPP"]
    top_h = round(H * L_H / nat_h)
    bot_h = H - top_h - G
    ws3 = _cols(3, W); xs3 = _xs(ws3)
    ws2 = _cols(2, W); xs2 = _xs(ws2)
    for i in range(3):
        _place(c, L[i], xs3[i], 0,         ws3[i], top_h)
    for i in range(2):
        _place(c, P[i], xs2[i], top_h + G, ws2[i], bot_h)

def layout_LLPPP(c, imgs, W, H):
    L, P  = _sort_LP(imgs)
    nat_w, nat_h = NATURAL["LLPPP"]
    top_h = round(H * L_H / nat_h)
    bot_h = H - top_h - G
    ws2 = _cols(2, W); xs2 = _xs(ws2)
    ws3 = _cols(3, W); xs3 = _xs(ws3)
    for i in range(2):
        _place(c, L[i], xs2[i], 0,         ws2[i], top_h)
    for i in range(3):
        _place(c, P[i], xs3[i], top_h + G, ws3[i], bot_h)

def layout_LPPPP(c, imgs, W, H):
    L, P = _sort_LP(imgs)
    nat_w, _ = NATURAL["LPPPP"]
    lw   = round(W * L_W / nat_w)
    pw   = W - lw - G
    ws2  = _cols(2, pw)
    xs2  = [lw + G + x for x in _xs(ws2)]
    hs   = _rows(2, H); ys = _ys(hs)
    _place(c, L[0], 0, 0, lw, H)
    idx = 0
    for r in range(2):
        for col in range(2):
            _place(c, P[idx], xs2[col], ys[r], ws2[col], hs[r])
            idx += 1

def layout_PPPPP(c, imgs, W, H):
    hs = _rows(2, H); ys = _ys(hs)
    ws3 = _cols(3, W); xs3 = _xs(ws3)
    ws2 = _cols(2, W); xs2 = _xs(ws2)
    for i in range(3):
        _place(c, imgs[i],     xs3[i], ys[0], ws3[i], hs[0])
    for i in range(2):
        _place(c, imgs[3 + i], xs2[i], ys[1], ws2[i], hs[1])


# ── Dispatch ───────────────────────────────────────────────────────────────────
LAYOUTS = {
    "L":     layout_L,     "P":     layout_P,
    "LL":    layout_LL,    "LP":    layout_LP,    "PP":    layout_PP,
    "LLL":   layout_LLL,   "LLP":   layout_LLP,   "LPP":   layout_LPP,   "PPP":   layout_PPP,
    "LLLL":  layout_LLLL,  "LLLP":  layout_LLLP,  "LLPP":  layout_LLPP,
    "LPPP":  layout_LPPP,  "PPPP":  layout_PPPP,
    "LLLLL": layout_LLLLL, "LLLLP": layout_LLLLP, "LLLPP": layout_LLLPP,
    "LLPPP": layout_LLPPP, "LPPPP": layout_LPPPP, "PPPPP": layout_PPPPP,
}


# ── Public API ─────────────────────────────────────────────────────────────────

def build_photo_block(image_paths, output_path=None):
    """
    Compose 1-5 images into a naturally-sized photo block.

    Canvas dimensions are derived from the layout's natural grid proportions.
    Portrait-heavy layouts are taller; landscape-heavy layouts are wider but
    shorter.  Neither dimension exceeds MAX_W (1275) or MAX_H (1650).

    Parameters
    ----------
    image_paths : list of str  — 1 to 5 JPEG or PNG file paths
    output_path : str or None  — if provided, saves result as PNG

    Returns
    -------
    PIL Image (RGB, variable size)
    """
    n = len(image_paths)
    if not 1 <= n <= 5:
        raise ValueError(f"Expected 1-5 images, got {n}")

    imgs = _load(image_paths)
    key  = "".join(sorted([_orient(i) for i in imgs]))

    if key not in LAYOUTS:
        raise KeyError(f"No layout for key '{key}'")

    W, H   = _canvas_size(key)
    canvas = Image.new("RGB", (W, H), BG)
    LAYOUTS[key](canvas, imgs, W, H)

    if output_path:
        canvas.save(output_path)
        print(f"Saved {output_path}  [{W}×{H}]  key={key}")

    return canvas


# ── CLI ────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    import sys
    if len(sys.argv) < 3:
        print("Usage: python photo_layout.py output.png img1.jpg [img2 ...]")
        sys.exit(1)
    build_photo_block(sys.argv[2:], output_path=sys.argv[1])
