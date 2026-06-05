# SPDX-License-Identifier: MIT
"""Shared helpers for the DocuSkills example builders.

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
import zipfile

_FIXED_DT = (1980, 1, 1, 0, 0, 0)        # constant DOS time for every zip member
_FIXED_ISO = "2026-06-05T00:00:00Z"      # constant W3CDTF for core.xml timestamps
_TS_RE = re.compile(
    rb"(<dcterms:(?:created|modified)[^>]*>)[^<]*(</dcterms:(?:created|modified)>)"
)


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
    with zipfile.ZipFile(io.BytesIO(data)) as zin, \
            zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            payload = zin.read(info.filename)
            if payload[:2] == b"PK" and info.filename.lower().endswith(
                (".xlsx", ".docx", ".pptx")
            ):
                payload = _freeze_bytes(payload)  # nested OOXML (pptx embedded workbook)
            if info.filename.endswith("core.xml"):
                payload = _TS_RE.sub(rb"\g<1>" + _FIXED_ISO.encode() + rb"\g<2>", payload)
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
