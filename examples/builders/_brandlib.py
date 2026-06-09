# SPDX-License-Identifier: MIT
"""Shared helpers for the BrandDocs example builders.

Two concerns the mechanical rebrand from the synthetic source fixtures made
fragile:

* ``rgba`` derives a logo pixel tuple straight from a palette hex constant, so
  the in-process logo can never drift from the brand palette again (the original
  hand-built logos kept literal source RGB tuples through the rebrand).
* ``freeze_ooxml`` rewrites a saved OOXML package so two builds are byte-for-byte
  identical: it pins every zip member timestamp and scrubs the docProps/core.xml
  ``dcterms:created`` / ``dcterms:modified`` values, recursing into nested OOXML
  (e.g. the chart's embedded workbook inside a .pptx). Without it, wall-clock
  leaks make the documented ``Regenerate`` step leave a dirty git tree.
"""

from __future__ import annotations

import io
import re
import struct
import zipfile
import zlib

_FIXED_DT = (1980, 1, 1, 0, 0, 0)  # constant DOS time for every zip member
_FIXED_ISO = "2026-06-05T00:00:00Z"  # constant W3CDTF for core.xml timestamps
_TS_RE = re.compile(
    rb"(<dcterms:(?:created|modified)[^>]*>)[^<]*(</dcterms:(?:created|modified)>)"
)


# BrandDocs theme palette as bare 6-digit hexes (no leading FF), one source of
# truth shared by every builder's theme1.xml clrScheme rewrite.
_BRAND_NAVY_HEX = "16213F"  # navy   -> dk1 (text)
_BRAND_TEAL_HEX = "2B7CD3"  # teal   -> accent1 (primary) + hlink
_BRAND_AMBER_HEX = "E0742B"  # amber  -> accent4 (danger) + accent2 + folHlink
_BRAND_LIGHT_HEX = "EAF1FF"  # light  -> lt1 (surface)


def brand_theme_slots() -> dict[str, str]:
    """Frozen ``clrScheme`` slot -> 6-digit hex map for the 4 BrandDocs colors.

    Additive, deterministic, side-effect free. A new helper (no existing
    signature touched) that any builder may import to drive its in-place
    ``theme1.xml`` clrScheme rewrite from a single source of truth, so the four
    BrandDocs hexes can never drift apart across the .docx / .pptx / .xlsx
    themes. Keys are the canonical DrawingML theme slots the extractor reads;
    values are bare ``RRGGBB`` (no ``#``, no leading ``FF``). Returns a fresh
    dict each call so a caller mutating the result can never corrupt the map.
    """
    return {
        "dk1": _BRAND_NAVY_HEX,
        "lt1": _BRAND_LIGHT_HEX,
        "accent1": _BRAND_TEAL_HEX,
        "accent2": _BRAND_AMBER_HEX,
        "accent4": _BRAND_AMBER_HEX,
        "hlink": _BRAND_TEAL_HEX,
        "folHlink": _BRAND_AMBER_HEX,
    }


def rgba(hexstr: str, alpha: int = 255) -> tuple:
    """``'FF16213F'`` or ``'16213F'`` -> ``(0x16, 0x21, 0x3F, alpha)``.

    Takes the last 6 hex chars, so it accepts both the 6-digit (docx) and the
    8-digit ARGB (xlsx) palette constants - the logo colour is now tied to the
    brand palette and can never drift from it again.
    """
    h = hexstr[-6:]
    return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16), alpha)


def _freeze_bytes(data: bytes) -> bytes:
    out = io.BytesIO()
    with (
        zipfile.ZipFile(io.BytesIO(data)) as zin,
        zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout,
    ):
        for info in zin.infolist():
            payload = zin.read(info.filename)
            if payload[:2] == b"PK" and info.filename.lower().endswith(
                (".xlsx", ".docx", ".pptx")
            ):
                payload = _freeze_bytes(
                    payload
                )  # nested OOXML (pptx embedded workbook)
            if info.filename.endswith("core.xml"):
                payload = _TS_RE.sub(
                    rb"\g<1>" + _FIXED_ISO.encode() + rb"\g<2>", payload
                )
            zi = zipfile.ZipInfo(info.filename, date_time=_FIXED_DT)
            zi.compress_type = zipfile.ZIP_DEFLATED
            zi.external_attr = info.external_attr
            zout.writestr(zi, payload)
    return out.getvalue()


def freeze_ooxml(path) -> None:
    """Rewrite the OOXML file at ``path`` in place so two builds are byte-identical."""
    p = str(path)
    with open(p, "rb") as fh:
        data = fh.read()
    frozen = _freeze_bytes(data)
    with open(p, "wb") as fh:
        fh.write(frozen)


# ---------------------------------------------------------------------------
# Brand imagery (deterministic, generated in-process).
#
# ``branddocs_mark_png`` is intentionally a simple wordmark: the brand name
# rendered as text, with no icon, badge, or decorative glyph. Every example
# template embeds the SAME generated PNG so the committed binaries stay synthetic
# and reproducible without shipping a proprietary asset.
# ---------------------------------------------------------------------------
# BrandDocs synthetic palette (navy, blue, amber, light).
_NAVY = (0x16, 0x21, 0x3F, 255)
_BLUE = (0x2B, 0x7C, 0xD3, 255)
_AMBER = (0xE0, 0x74, 0x2B, 255)
_LIGHT = (0xEA, 0xF1, 0xFF, 255)
_GRID = (0x2E, 0x3C, 0x68, 255)
_WHITE = (0xFF, 0xFF, 0xFF, 255)
_SS = 4  # supersample factor


def _png_bytes(rgba: bytes, w: int, h: int) -> bytes:
    """Encode a straight-RGBA buffer (len w*h*4) as a PNG, zlib level 9."""
    raw = bytearray()
    stride = w * 4
    for y in range(h):
        raw.append(0)  # filter type 0 (None) per scanline
        raw.extend(rgba[y * stride : (y + 1) * stride])

    def _chunk(tag: bytes, data: bytes) -> bytes:
        return (
            struct.pack(">I", len(data))
            + tag
            + data
            + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)
        )

    sig = b"\x89PNG\r\n\x1a\n"
    ihdr = struct.pack(">IIBBBBB", w, h, 8, 6, 0, 0, 0)  # 8-bit RGBA
    idat = zlib.compress(bytes(raw), 9)
    return sig + _chunk(b"IHDR", ihdr) + _chunk(b"IDAT", idat) + _chunk(b"IEND", b"")


def _downsample(hi: bytearray, W: int, H: int, s: int) -> tuple[bytes, int, int]:
    """Alpha-weighted box downsample of a hi-res RGBA buffer by factor ``s``."""
    ow, oh = W // s, H // s
    out = bytearray(ow * oh * 4)
    for oy in range(oh):
        for ox in range(ow):
            sr = sg = sb = sa = 0
            for dy in range(s):
                base = ((oy * s + dy) * W + ox * s) * 4
                for dx in range(s):
                    i = base + dx * 4
                    a = hi[i + 3]
                    sr += hi[i] * a
                    sg += hi[i + 1] * a
                    sb += hi[i + 2] * a
                    sa += a
            o = (oy * ow + ox) * 4
            n = s * s
            if sa:
                out[o] = sr // sa
                out[o + 1] = sg // sa
                out[o + 2] = sb // sa
            out[o + 3] = sa // n
    return bytes(out), ow, oh


def _in_round(x: int, y: int, x0: int, y0: int, x1: int, y1: int, r: int) -> bool:
    if x < x0 or x >= x1 or y < y0 or y >= y1:
        return False
    if r <= 0:
        return True
    for cx, cy in (
        (x0 + r, y0 + r),
        (x1 - 1 - r, y0 + r),
        (x0 + r, y1 - 1 - r),
        (x1 - 1 - r, y1 - 1 - r),
    ):
        in_corner = (
            (x < x0 + r and (cx == x0 + r)) or (x > x1 - 1 - r and cx == x1 - 1 - r)
        ) and ((y < y0 + r and cy == y0 + r) or (y > y1 - 1 - r and cy == y1 - 1 - r))
        if in_corner:
            dx, dy = x - cx, y - cy
            return dx * dx + dy * dy <= r * r
    return True


def _fill_round(buf, W, x0, y0, x1, y1, r, color):
    for y in range(max(0, y0), y1):
        row = y * W * 4
        for x in range(max(0, x0), x1):
            if _in_round(x, y, x0, y0, x1, y1, r):
                i = row + x * 4
                buf[i], buf[i + 1], buf[i + 2], buf[i + 3] = color


def _stroke_round(buf, W, x0, y0, x1, y1, r, t, color):
    ix0, iy0, ix1, iy1 = x0 + t, y0 + t, x1 - t, y1 - t
    ir = max(0, r - t)
    for y in range(max(0, y0), y1):
        for x in range(max(0, x0), x1):
            if _in_round(x, y, x0, y0, x1, y1, r) and not _in_round(
                x, y, ix0, iy0, ix1, iy1, ir
            ):
                i = (y * W + x) * 4
                buf[i], buf[i + 1], buf[i + 2], buf[i + 3] = color


def _disc(buf, W, H, cx, cy, rad, color):
    for y in range(max(0, cy - rad), min(H, cy + rad + 1)):
        for x in range(max(0, cx - rad), min(W, cx + rad + 1)):
            dx, dy = x - cx, y - cy
            if dx * dx + dy * dy <= rad * rad:
                i = (y * W + x) * 4
                buf[i], buf[i + 1], buf[i + 2], buf[i + 3] = color


def branddocs_mark_png(width: int = 640, height: int = 160) -> bytes:
    """Return a transparent BrandDocs wordmark PNG.

    The logo is simply the brand name: ``Brand`` in navy and ``Docs`` in blue.
    Pillow's embedded default scalable font keeps this deterministic and avoids
    depending on platform fonts.
    """
    from PIL import Image, ImageDraw, ImageFont

    scale = 4
    W, H = width * scale, height * scale
    img = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    target_w = int(W * 0.9)
    target_h = int(H * 0.66)
    size = max(12, int(H * 0.58))
    stroke = 1
    font = ImageFont.load_default(size=size)
    bbox = draw.textbbox((0, 0), "BrandDocs", font=font, stroke_width=stroke)
    while size > 12 and (bbox[2] - bbox[0] > target_w or bbox[3] - bbox[1] > target_h):
        size -= 4
        stroke = max(1, size // 48)
        font = ImageFont.load_default(size=size)
        bbox = draw.textbbox((0, 0), "BrandDocs", font=font, stroke_width=stroke)

    text_w = bbox[2] - bbox[0]
    text_h = bbox[3] - bbox[1]
    x = (W - text_w) // 2 - bbox[0]
    y = (H - text_h) // 2 - bbox[1]
    brand_advance = int(round(draw.textlength("Brand", font=font)))
    draw.text(
        (x, y), "Brand", font=font, fill=_NAVY, stroke_width=stroke, stroke_fill=_NAVY
    )
    draw.text(
        (x + brand_advance, y),
        "Docs",
        font=font,
        fill=_BLUE,
        stroke_width=stroke,
        stroke_fill=_BLUE,
    )
    if scale != 1:
        img = img.resize((width, height), Image.Resampling.LANCZOS)
    out = io.BytesIO()
    img.save(out, format="PNG", compress_level=9)
    return out.getvalue()


def branddocs_curve_png(width: int = 480, height: int = 200) -> bytes:
    """A small rising 'growth curve' figure on a light brand card.

    A light rounded card, faint navy axes, a thick blue rising polyline and
    amber vertex dots - a real chart-like figure (deterministic, stdlib only).
    """
    W, H = width * _SS, height * _SS
    buf = bytearray(W * H * 4)
    pad = int(8 * _SS)
    # light card background.
    _fill_round(buf, W, 0, 0, W, H, int(10 * _SS), _LIGHT)
    # plot box.
    px0, py0, px1, py1 = pad * 4, pad * 2, W - pad * 2, H - pad * 4
    # axes (thin navy).
    aw = max(_SS, 2 * _SS)
    for y in range(py1, py1 + aw):
        for x in range(px0, px1):
            i = (y * W + x) * 4
            buf[i], buf[i + 1], buf[i + 2], buf[i + 3] = _GRID
    for x in range(px0, px0 + aw):
        for y in range(py0, py1):
            i = (y * W + x) * 4
            buf[i], buf[i + 1], buf[i + 2], buf[i + 3] = _GRID
    # rising series.
    ys = [0.18, 0.34, 0.30, 0.55, 0.74, 0.95]
    n = len(ys)
    pts = []
    for k, v in enumerate(ys):
        x = px0 + int((px1 - px0) * k / (n - 1))
        y = py1 - int((py1 - py0) * v)
        pts.append((x, y))
    lw = max(_SS, 3 * _SS)
    for (x0, y0), (x1, y1) in zip(pts, pts[1:]):
        steps = max(abs(x1 - x0), abs(y1 - y0), 1)
        for s in range(steps + 1):
            x = x0 + (x1 - x0) * s // steps
            y = y0 + (y1 - y0) * s // steps
            _disc(buf, W, H, x, y, lw, _BLUE)
    for x, y in pts:
        _disc(buf, W, H, x, y, lw + _SS, _AMBER)
    rgba, ow, oh = _downsample(buf, W, H, _SS)
    return _png_bytes(rgba, ow, oh)
