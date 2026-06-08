# SPDX-License-Identifier: MIT
"""Deterministic builder for the COMPLEX synthetic DOCX example template.

Produces ``examples/templates/branddocs_template.docx``: a 100% synthetic
(``BrandDocs Corp``, never proprietary) Word template that stresses the brand-docx
extractor / generator across as many Word component types as can be authored
without Microsoft Word - python-docx for the bulk, and raw lxml for the many
parts python-docx cannot reach (block-level ``w:sdt`` content controls, real
TOC / SEQ complex fields, ``word/numbering.xml`` abstractNum/num, a custom
``w:type="table"`` style, header logo drawing, ``PAGE`` field, footnotes, and a
landscape ``w:sectPr``).

Components authored (each is a real OOXML structure, not a text approximation):

  COVER (multi-slot front matter, all in the cover region before the TOC):
    * a block-level ``w:sdt`` TITLE content control with ``w:alias='Title'`` and a
      realistic synthetic title - the shape ``cover.discover_cover`` keys on as an
      SDT anchor (python-docx cannot author this, so it is lxml);
    * a SUBTITLE / description placeholder paragraph;
    * a DOCUMENT-ID placeholder paragraph ("Document ID: {{doc_id}}");
    * a DATE placeholder paragraph ("{{date}}").
    * a compact executive scorecard table with three synthetic KPI tiles.

  INDEX FRONT MATTER (three real complex fields, each with cached demo entries):
    * an outline Table of Contents  ``TOC \\o "1-3" \\h \\z \\u``;
    * a Table of Tables             ``TOC \\h \\z \\c "Table"``;
    * a Table of Figures            ``TOC \\h \\z \\c "Figure"``.
    Each carries a cached result (styled entry paragraphs with PAGEREF) so the
    field renders before Word recomputes it, and each ``\\c`` switch is the
    opaque seq id ``structure.inventory_fields`` surfaces.

  NUMBERING (real ``word/numbering.xml``, referenced via ``w:numPr`` from named
    paragraph styles, not direct formatting):
    * a 2-level BULLET list  -> styles "BrandDocs Bullet L1" / "BrandDocs Bullet L2";
    * a 1-level NUMBERED list -> style "BrandDocs Number L1".

  TABLES: a custom ``w:type="table"`` style "BrandDocs Table" (header-row
    shading + row banding via ``w:tblStylePr`` conditional formatting) applied to
    a sample revenue table, an executive scorecard, a risk/readiness matrix, and
    a landscape rollout matrix, with real ``SEQ Table`` captions on body tables.

  FIGURE: two real inline PNG figures, each with a ``SEQ Figure`` caption -
    Figure 1 is the shared text-only BrandDocs wordmark (image1.png) and Figure 2
    is a real rising growth curve (image2.png, a distinct media part).

  CALLOUT: a paragraph style "BrandDocs Callout" with shading + a box border.
  QUOTE: a paragraph style "BrandDocs Quote" with a navy/amber editorial treatment.

  HEADER / FOOTER: the shared BrandDocs wordmark in the default header, and a
    ``PAGE`` field in the default footer.

  SECTIONS: a PORTRAIT first section and a LANDSCAPE second section (distinct
    ``w:sectPr`` page size + orientation).

  FOOTNOTE: one real footnote (``word/footnotes.xml`` + a referencing run).

  DEMO BODY: instruction / lorem-ipsum body content (an "Example heading"
    Heading-1 and following paragraphs) a generation is expected to clear.

The output is byte-reproducible: every id / image byte / part is fixed, no
random or timestamp, so re-running yields an identical file.

Run:
    PYTHONPATH=scripts .venv/bin/python examples/builders/build_branddocs_docx.py
"""

from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Twips
from lxml import etree

from _brandlib import branddocs_curve_png, branddocs_mark_png, freeze_ooxml

OUT = Path(__file__).resolve().parents[1] / "templates" / "branddocs_template.docx"

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"

# Synthetic BrandDocs Corp brand palette (made-up; never proprietary).
BRAND_NAVY = "16213F"
BRAND_TEAL = "2B7CD3"
BRAND_AMBER = "E0742B"
BRAND_LIGHT = "EAF1FF"
BRAND_BAND = "DCE7FF"
WHITE = "FFFFFF"


# ---------------------------------------------------------------------------
# lxml element helpers
# ---------------------------------------------------------------------------
def _w(tag: str) -> str:
    return f"{{{W}}}{tag}"


def _el(tag: str, **attrs) -> etree._Element:
    """Make a ``w:``-namespaced element with ``w:``-namespaced attributes."""
    e = etree.SubElement(etree.Element(_w("_root")), _w(tag))
    e.getparent().remove(e)
    for k, v in attrs.items():
        e.set(_w(k), v)
    return e


def _sub(parent: etree._Element, tag: str, **attrs) -> etree._Element:
    e = etree.SubElement(parent, _w(tag))
    for k, v in attrs.items():
        e.set(_w(k), v)
    return e


def _run(text: str, *, preserve: bool = True, instr: bool = False) -> etree._Element:
    """A ``w:r`` carrying either a ``w:t`` or a ``w:instrText``."""
    r = _el("r")
    leaf = _sub(r, "instrText" if instr else "t")
    leaf.text = text
    if preserve:
        leaf.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    return r


def _fldchar(kind: str, *, dirty: bool = False) -> etree._Element:
    r = _el("r")
    fc = _sub(r, "fldChar", fldCharType=kind)
    if dirty and kind == "begin":
        fc.set(_w("dirty"), "true")
    return r


def _add_external_hyperlink(doc, paragraph, url: str, text: str) -> None:
    """Append a real ``w:hyperlink`` (external relationship) carrying ``text``.

    Builds the ``r:id``-namespaced attribute directly (the ``_sub`` helper only sets
    ``w:`` attributes). No ``Hyperlink`` character style is referenced (the template
    does not register one); the link is realized purely by the relationship.
    """
    rid = doc.part.relate_to(url, f"{R}/hyperlink", is_external=True)
    hl = etree.SubElement(paragraph._p, _w("hyperlink"))
    hl.set(f"{{{R}}}id", rid)
    hl.append(_run(text))


def _set_pstyle(paragraph, style_id: str) -> None:
    """Stamp ``w:pPr/w:pStyle@w:val`` on a paragraph by STYLE ID directly.

    Bypasses python-docx's ``style=`` name lookup (which is deprecated for ids and
    cannot reference a custom style not registered under its display name), and
    matches how the extractor keys on style ids.
    """
    p = paragraph._p
    pPr = p.get_or_add_pPr()
    existing = pPr.find(_w("pStyle"))
    if existing is not None:
        existing.set(_w("val"), style_id)
        return
    pStyle = etree.SubElement(pPr, _w("pStyle"))
    pStyle.set(_w("val"), style_id)
    pPr.insert(0, pStyle)


def _p(doc, text: str = "", style_id: str | None = None):
    """Add a body paragraph, optionally stamping a style id directly."""
    paragraph = doc.add_paragraph(text)
    if style_id:
        _set_pstyle(paragraph, style_id)
    return paragraph


# ---------------------------------------------------------------------------
# Custom STYLES (paragraph styles + the branded table style). Authored straight
# into ``word/styles.xml`` via lxml so we control header shading / banding /
# borders python-docx cannot express.
# ---------------------------------------------------------------------------
def _add_paragraph_style(
    styles,
    style_id,
    name,
    *,
    based_on="Normal",
    color=None,
    bold=False,
    size_pt=None,
    shading=None,
    box_border=None,
    left_accent=None,
):
    st = _sub(styles, "style", type="paragraph", styleId=style_id)
    st.set(_w("customStyle"), "1")
    _sub(st, "name", val=name)
    _sub(st, "basedOn", val=based_on)
    _sub(st, "qFormat")
    pPr = _sub(st, "pPr")
    if shading:
        _sub(pPr, "shd", val="clear", color="auto", fill=shading)
    if box_border:
        pbdr = _sub(pPr, "pBdr")
        for side in ("top", "left", "bottom", "right"):
            sz = "12"
            col = box_border
            # BP-CALLOUT-ACCENT: thick amber accent bar on the left edge.
            if side == "left" and left_accent:
                sz = "24"
                col = left_accent
            _sub(pbdr, side, val="single", sz=sz, space="6", color=col)
        _sub(pPr, "spacing", before="120", after="120")
    rPr = _sub(st, "rPr")
    if bold:
        _sub(rPr, "b")
    if color:
        _sub(rPr, "color", val=color)
    if size_pt:
        _sub(rPr, "sz", val=str(int(size_pt * 2)))
    return st


def _add_list_style(styles, style_id, name, num_id, ilvl=0):
    """A list paragraph style that references a w:num via w:numPr.

    ``ilvl`` pins the list level on the style's ``w:numPr`` (``w:ilvl`` before
    ``w:numId``, OOXML order). A second-level bullet style MUST declare
    ``w:ilvl=1`` so ``structure.style_num_binding`` reads level 1 and the role
    resolves to ``list.bullet.2`` (1-based = ilvl+1) instead of colliding with
    the level-1 bullet role on a missing-ilvl default of 0.
    """
    st = _sub(styles, "style", type="paragraph", styleId=style_id)
    st.set(_w("customStyle"), "1")
    _sub(st, "name", val=name)
    _sub(st, "basedOn", val="ListParagraph")
    _sub(st, "qFormat")
    pPr = _sub(st, "pPr")
    numPr = _sub(pPr, "numPr")
    _sub(numPr, "ilvl", val=str(ilvl))
    _sub(numPr, "numId", val=str(num_id))
    return st


def _add_table_style(styles):
    """A custom ``w:type='table'`` style: header-row shading + row banding."""
    st = _sub(styles, "style", type="table", styleId="BrandDocsTable")
    st.set(_w("customStyle"), "1")
    _sub(st, "name", val="BrandDocs Table")
    _sub(st, "basedOn", val="TableNormal")
    _sub(st, "uiPriority", val="99")
    # Base table: thin navy grid + banded-row size hint.
    tblPr = _sub(st, "tblPr")
    _sub(tblPr, "tblStyleRowBandSize", val="1")
    borders = _sub(tblPr, "tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        _sub(borders, side, val="single", sz="4", space="0", color=BRAND_NAVY)
    # Default cell run color.
    rPr = _sub(st, "rPr")
    _sub(rPr, "color", val=BRAND_NAVY)
    # First-row (header) conditional formatting: navy fill + white bold text.
    fr = _sub(st, "tblStylePr", type="firstRow")
    fr_rpr = _sub(fr, "rPr")
    _sub(fr_rpr, "b")
    _sub(fr_rpr, "color", val=WHITE)
    fr_tcpr = _sub(fr, "tcPr")
    _sub(fr_tcpr, "shd", val="clear", color="auto", fill=BRAND_NAVY)
    # Banded rows: light-blue fill on every other row.
    band = _sub(st, "tblStylePr", type="band1Horz")
    band_tcpr = _sub(band, "tcPr")
    _sub(band_tcpr, "shd", val="clear", color="auto", fill=BRAND_BAND)
    return st


def _ensure_list_paragraph_style(styles):
    """python-docx's default styles.xml has no 'List Paragraph'; add a minimal one."""
    for st in styles.findall(_w("style")):
        if st.get(_w("styleId")) == "ListParagraph":
            return
    st = _sub(styles, "style", type="paragraph", styleId="ListParagraph")
    _sub(st, "name", val="List Paragraph")
    _sub(st, "basedOn", val="Normal")
    _sub(st, "uiPriority", val="34")
    _sub(st, "qFormat")
    pPr = _sub(st, "pPr")
    _sub(pPr, "ind", left="720")


def _ensure_toc_styles(styles):
    """Add TOC entry styles + TOCHeading + Caption + Footnote styles if absent."""
    have = {st.get(_w("styleId")) for st in styles.findall(_w("style"))}

    def add_simple(style_id, name, *, based_on="Normal", ui="39"):
        if style_id in have:
            return
        st = _sub(styles, "style", type="paragraph", styleId=style_id)
        _sub(st, "name", val=name)
        _sub(st, "basedOn", val=based_on)
        _sub(st, "uiPriority", val=ui)
        if style_id.startswith("TOC"):
            pPr = _sub(st, "pPr")
            _sub(pPr, "tabs")  # placeholder; real entries carry their own tabs

    for lvl in (1, 2, 3):
        add_simple(f"TOC{lvl}", f"TOC {lvl}")
    add_simple("TOCHeading", "TOC Heading", based_on="Heading1")
    add_simple("TableofFigures", "Table of Figures")
    if "Caption" not in have:
        st = _sub(styles, "style", type="paragraph", styleId="Caption")
        _sub(st, "name", val="Caption")
        _sub(st, "basedOn", val="Normal")
        _sub(st, "uiPriority", val="35")
        _sub(st, "qFormat")
        rPr = _sub(st, "rPr")
        _sub(rPr, "i")
        _sub(rPr, "color", val=BRAND_TEAL)
        _sub(rPr, "sz", val="18")
    if "FootnoteText" not in have:
        st = _sub(styles, "style", type="paragraph", styleId="FootnoteText")
        _sub(st, "name", val="Footnote Text")
        _sub(st, "basedOn", val="Normal")
        rPr = _sub(st, "rPr")
        _sub(rPr, "sz", val="20")
    if "FootnoteReference" not in have:
        st = _sub(styles, "style", type="character", styleId="FootnoteReference")
        _sub(st, "name", val="Footnote Reference")
        rPr = _sub(st, "rPr")
        _sub(rPr, "vertAlign", val="superscript")


def _build_styles(doc):
    styles = doc.styles.element
    _ensure_list_paragraph_style(styles)
    _ensure_toc_styles(styles)
    # Branded paragraph styles.
    _add_paragraph_style(
        styles,
        "BrandDocsCoverTitle",
        "BrandDocs Cover Title",
        color=BRAND_NAVY,
        bold=True,
        size_pt=28,
    )
    _add_paragraph_style(
        styles,
        "BrandDocsCoverSubtitle",
        "BrandDocs Cover Subtitle",
        color=BRAND_TEAL,
        size_pt=14,
    )
    _add_paragraph_style(
        styles,
        "BrandDocsCallout",
        "BrandDocs Callout",
        color=BRAND_NAVY,
        shading=BRAND_LIGHT,
        box_border=BRAND_TEAL,
        left_accent=BRAND_AMBER,
    )
    _add_paragraph_style(
        styles,
        "BrandDocsQuote",
        "BrandDocs Quote",
        color=BRAND_NAVY,
        shading=BRAND_BAND,
        box_border=BRAND_NAVY,
        left_accent=BRAND_AMBER,
        size_pt=12,
    )
    # List styles -> reference w:num 1 (bullet L1), 2 (bullet L2), 3 (number L1).
    # BUG-LIST-ILVL: the L2 bullet pins w:ilvl=1 so it binds to level 1 of
    # abstractNum 0 (a distinct ``list.bullet.2`` role, not a dedup of L1).
    _add_list_style(
        styles, "BrandDocsBulletL1", "BrandDocs Bullet L1", num_id=1, ilvl=0
    )
    _add_list_style(
        styles, "BrandDocsBulletL2", "BrandDocs Bullet L2", num_id=2, ilvl=1
    )
    _add_list_style(
        styles, "BrandDocsNumberL1", "BrandDocs Number L1", num_id=3, ilvl=0
    )
    # Branded table style.
    _add_table_style(styles)


# ---------------------------------------------------------------------------
# NUMBERING - a real ``word/numbering.xml`` with abstractNum + num.
# python-docx exposes no numbering authoring API for a fresh document, so the
# part is built with lxml and attached to the package.
# ---------------------------------------------------------------------------
def _populate_numbering(root) -> None:
    """Fill an existing ``w:numbering`` element with BrandDocs abstractNum/num defs.

    The python-docx default template already ships a (empty) ``word/numbering.xml``
    part, already related from ``document.xml``. We MUST reuse that part - adding a
    second part of the same name corrupts the zip - so this mutates the existing
    ``CT_Numbering`` element in place rather than authoring a fresh part.
    """
    for child in list(root):
        root.remove(child)

    def abstract(aid, levels):
        an = _sub(root, "abstractNum", abstractNumId=str(aid))
        _sub(an, "multiLevelType", val="hybridMultilevel")
        for lvl, level in enumerate(levels):
            fmt, text, indent = level[0], level[1], level[2]
            font = level[3] if len(level) > 3 else None
            lvl_el = _sub(an, "lvl", ilvl=str(lvl))
            _sub(lvl_el, "start", val="1")
            _sub(lvl_el, "numFmt", val=fmt)
            # BUG-LIST-BULLETGLYPH: an EXPLICIT, real Unicode lvlText glyph on a
            # plain text font. The old defs used Symbol-font private-use codepoints
            # (U+F0B7 / U+F0A7) which Word maps to a bullet/section mark but
            # LibreOffice renders as a stray club/box (no glyph at those PUA
            # points without the Symbol font). Plain Unicode bullets on a standard
            # font render identically in Word and LibreOffice.
            _sub(lvl_el, "lvlText", val=text)
            _sub(lvl_el, "lvlJc", val="left")
            pPr = _sub(lvl_el, "pPr")
            _sub(pPr, "ind", left=str(indent), hanging="360")
            if fmt == "bullet" and font:
                rPr = _sub(lvl_el, "rPr")
                rfonts = _sub(rPr, "rFonts")
                rfonts.set(_w("ascii"), font)
                rfonts.set(_w("hAnsi"), font)
                rfonts.set(_w("cs"), font)
                rfonts.set(_w("hint"), "default")
        return an

    # abstractNum 0: two-level bullet. L1 = filled round bullet (U+2022), L2 =
    # en-dash (U+2013), both on Arial so the glyphs are readable in Word AND
    # LibreOffice (no Symbol-font club/box).
    abstract(
        0,
        [
            ("bullet", "•", 720, "Arial"),
            ("bullet", "–", 1440, "Arial"),
        ],
    )
    # abstractNum 1: decimal numbered list.
    abstract(1, [("decimal", "%1.", 720)])

    # num bindings: 1->bullet(ilvl0 entry), 2->bullet, 3->decimal.
    for num_id, aid in ((1, 0), (2, 0), (3, 1)):
        n = _sub(root, "num", numId=str(num_id))
        _sub(n, "abstractNumId", val=str(aid))


# ---------------------------------------------------------------------------
# FOOTNOTES - a real ``word/footnotes.xml`` with the two reserved separators
# plus one authored footnote (id 2).
# ---------------------------------------------------------------------------
def _build_footnotes_xml() -> bytes:
    nsmap = {"w": W}
    root = etree.Element(_w("footnotes"), nsmap=nsmap)

    def sep(fid, kind):
        fn = _sub(root, "footnote", type=kind, id=str(fid))
        p = _sub(fn, "p")
        r = _sub(p, "r")
        _sub(r, kind if kind == "separator" else "continuationSeparator")

    s = _sub(root, "footnote", type="separator", id="-1")
    p = _sub(s, "p")
    r = _sub(p, "r")
    _sub(r, "separator")
    c = _sub(root, "footnote", type="continuationSeparator", id="0")
    p = _sub(c, "p")
    r = _sub(p, "r")
    _sub(r, "continuationSeparator")

    # The authored footnote (id 2).
    fn = _sub(root, "footnote", id="2")
    p = _sub(fn, "p")
    pPr = _sub(p, "pPr")
    _sub(pPr, "pStyle", val="FootnoteText")
    ref_run = _sub(p, "r")
    ref_rpr = _sub(ref_run, "rPr")
    _sub(ref_rpr, "rStyle", val="FootnoteReference")
    _sub(ref_run, "footnoteRef")
    txt = _sub(p, "r")
    t = _sub(txt, "t")
    t.text = " BrandDocs Corp is a fictional company used only for testing."
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _attach_part(doc, partname, content_type, xml_bytes, rel_type):
    """Register a new package part + content-type override + document rel.

    Returns the relationship id assigned to ``document.xml -> partname``.
    """
    from docx.opc.part import Part
    from docx.opc.packuri import PackURI

    package = doc.part.package
    puri = PackURI("/" + partname)
    part = Part(puri, content_type, xml_bytes, package)
    rid = doc.part.relate_to(part, rel_type)
    return rid


# ---------------------------------------------------------------------------
# COVER - a block-level w:sdt title content control + placeholder paragraphs.
# ---------------------------------------------------------------------------
def _cover_title_sdt():
    """A block-level ``w:sdt`` cover-title content control (lxml; python-docx
    cannot author block-level SDTs).

    CR-COVER-TITLE: the SDT keeps ``w:alias='Title'`` (the strong cover-title
    anchor evidence ``cover._sdt_is_title`` keys on) but carries a REALISTIC
    synthetic title instead of the "Insert title here" prompt - the bare prompt
    read like an unfinished draft. The ``w:showingPlcHdr`` flag is dropped so Word
    treats the text as real content, not a greyed placeholder. Cover-anchor count
    is unaffected: ``discover_cover`` keys the SDT anchor on the container, not on
    the text, and the alias still classifies it as the title slot.
    """
    sdt = _el("sdt")
    sdtPr = _sub(sdt, "sdtPr")
    rpr = _sub(sdtPr, "rPr")
    _sub(rpr, "color", val=BRAND_NAVY)
    _sub(rpr, "sz", val="56")
    _sub(rpr, "b")
    _sub(sdtPr, "alias", val="Title")
    _sub(sdtPr, "tag", val="branddocs_title")
    _sub(sdtPr, "id", val="101")
    placeholder = _sub(sdtPr, "placeholder")
    _sub(placeholder, "docPart", val="DefaultPlaceholder_Title")
    _sub(sdtPr, "text")
    _sub(sdt, "sdtEndPr")
    sdtContent = _sub(sdt, "sdtContent")
    p = _sub(sdtContent, "p")
    pPr = _sub(p, "pPr")
    _sub(pPr, "pStyle", val="BrandDocsCoverTitle")
    r = _sub(p, "r")
    rpr2 = _sub(r, "rPr")
    _sub(rpr2, "color", val=BRAND_NAVY)
    _sub(rpr2, "sz", val="56")
    _sub(rpr2, "b")
    t = _sub(r, "t")
    t.text = "BrandDocs Brand Operations Review"
    return sdt


def _cover_band(doc):
    """BP-COVER-BAND: a navy brand banner paragraph with an amber bottom rule.

    Carries ``w:pPr/w:shd`` navy fill + a thick amber ``w:pBdr`` bottom border.
    Its text is EMPTY so ``_paragraph_is_placeholder_slot`` (which excludes empty
    paragraphs) never promotes it to a 5th cover anchor; placed after the date
    slot it leaves the ``sdt.0/para.1/para.2/para.3`` anchor indices stable.
    """
    p = doc.add_paragraph("")
    pPr = p._p.get_or_add_pPr()
    _sub(pPr, "shd", val="clear", color="auto", fill=BRAND_NAVY)
    pbdr = _sub(pPr, "pBdr")
    _sub(pbdr, "bottom", val="single", sz="24", space="2", color=BRAND_AMBER)
    _sub(pPr, "spacing", before="60", after="180")
    return p


def _build_cover(doc):
    body = doc.element.body
    # 1) Block-level SDT title (inserted as the very first body child).
    sdt = _cover_title_sdt()
    body.insert(0, sdt)

    # 2..n) Placeholder paragraphs for subtitle / doc-id / date, each in the
    # cover region (before the TOC), each a short single-line slot. The slot
    # demo values are realistic synthetic content (CR-COVER-VALUES); the SDT
    # title above now carries a realistic synthetic title (CR-COVER-TITLE) while
    # still classifying as the title slot via its ``w:alias='Title'``.
    sub_p = _p(
        doc,
        "Annual Brand Operations Review - BrandDocs Corp (synthetic)",
        "BrandDocsCoverSubtitle",
    )
    docid_p = doc.add_paragraph("Document ID: DSK-BR-2026-014")
    date_p = doc.add_paragraph("June 5, 2026")
    # A navy brand band sits BELOW the date slot (empty text, not an anchor).
    band_p = _cover_band(doc)
    # Move them right after the SDT (they were appended at the end of the body),
    # preserving order: SDT -> subtitle -> doc-id -> date -> band.
    for p in (band_p._p, date_p._p, docid_p._p, sub_p._p):
        body.remove(p)
        sdt.addnext(p)


def _build_cover_scorecard(doc):
    """A compact executive scorecard in the front matter.

    It is a table, not a paragraph, so it enriches the rendered cover without
    creating extra paragraph cover anchors. This gives the example a realistic
    corporate cover surface while keeping the cover-slot ids stable.
    """
    table = doc.add_table(rows=2, cols=3)
    table.style = "BrandDocs Table"
    _set_col_widths(table, (2160, 2160, 2160))
    labels = ("Brand health", "Template coverage", "Audit posture")
    values = ("92", "3 formats", "Deep QA")
    for cell, label in zip(table.rows[0].cells, labels):
        cell.text = label
    for cell, value in zip(table.rows[1].cells, values):
        cell.text = value
    doc.add_paragraph("")


# ---------------------------------------------------------------------------
# INDEX FRONT MATTER - three real complex fields with cached demo entries.
# ---------------------------------------------------------------------------
def _toc_heading(doc, text):
    return _p(doc, text, "TOCHeading")


def _toc_entry(doc, label, page, *, style):
    """A cached TOC/index entry paragraph: a nested PAGEREF field + tab + page."""
    p = _p(doc, "", style)
    pp = p._p
    # The entry hyperlink text run.
    r = _sub(pp, "r")
    t = _sub(r, "t")
    t.text = label
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    # A right tab + a cached PAGEREF (begin / instr / separate / page / end).
    tab = _sub(pp, "r")
    _sub(tab, "tab")
    pp.append(_fldchar("begin"))
    pp.append(_run(f" PAGEREF _Toc{page:04d} \\h ", instr=True))
    pp.append(_fldchar("separate"))
    pp.append(_run(str(page)))
    pp.append(_fldchar("end"))
    return p


def _complex_field(doc, instr, cached_paragraph_builder):
    """Wrap a complex field around cached result paragraphs.

    Layout (Word's pattern for a multi-paragraph field):
      p0: [begin(dirty)] [instrText]
      p1..pn: cached entry paragraphs (the "result"), each a normal paragraph
      pE: [end]
    A ``separate`` fldChar precedes the cached result; the closing ``end`` sits
    in its own trailing paragraph.
    """
    # Field begin + instruction (its own paragraph).
    begin_p = doc.add_paragraph()
    bp = begin_p._p
    bp.append(_fldchar("begin", dirty=True))
    bp.append(_run(instr, instr=True))
    bp.append(_fldchar("separate"))
    # Cached result entries (real styled paragraphs in between).
    cached_paragraph_builder(doc)
    # Field end (its own paragraph).
    end_p = doc.add_paragraph()
    end_p._p.append(_fldchar("end"))


def _build_index_front_matter(doc):
    # --- Table of Contents (outline) ---
    _toc_heading(doc, "Table of Contents")

    def _toc_entries(d):
        _toc_entry(d, "1  Overview", 3, style="TOC1")
        _toc_entry(d, "1.1  Scope", 3, style="TOC2")
        _toc_entry(d, "1.2  Audience", 4, style="TOC2")
        _toc_entry(d, "2  Methodology", 5, style="TOC1")
        _toc_entry(d, "2.1  Data sources", 5, style="TOC2")
        _toc_entry(d, "3  Results", 6, style="TOC1")

    _complex_field(doc, 'TOC \\o "1-3" \\h \\z \\u ', _toc_entries)

    # --- Table of Tables ---
    _toc_heading(doc, "Table of Tables")

    def _tot_entries(d):
        # Cached entries mirror the two real SEQ Table captions in the body and
        # the landscape appendix (does NOT change the fields count).
        _toc_entry(
            d, "Table 1. BrandDocs FY2026 quarterly revenue", 5, style="TableofFigures"
        )
        _toc_entry(
            d, "Table 2. BrandDocs risk and readiness matrix", 6, style="TableofFigures"
        )
        _toc_entry(
            d, "Table 3. BrandDocs program rollout matrix", 7, style="TableofFigures"
        )

    _complex_field(doc, 'TOC \\h \\z \\c "Table" ', _tot_entries)

    # --- Table of Figures ---
    _toc_heading(doc, "Table of Figures")

    def _tof_entries(d):
        _toc_entry(d, "Figure 1. BrandDocs Corp wordmark", 2, style="TableofFigures")
        _toc_entry(d, "Figure 2. BrandDocs growth curve", 6, style="TableofFigures")

    _complex_field(doc, 'TOC \\h \\z \\c "Figure" ', _tof_entries)


# ---------------------------------------------------------------------------
# DEMO BODY - lists, table+caption, figure+caption, callout, footnote, demo
# heading content. Everything below the index front matter is the freeform body
# a generation would clear.
# ---------------------------------------------------------------------------
def _seq_caption(doc, prefix, seq_name, tail, *, style="Caption"):
    """A real ``SEQ`` caption paragraph: 'Prefix N. tail' with a live SEQ field."""
    p = _p(doc, "", style)
    pp = p._p
    pp.append(_run(f"{prefix} "))
    pp.append(_fldchar("begin"))
    pp.append(_run(f" SEQ {seq_name} \\* ARABIC ", instr=True))
    pp.append(_fldchar("separate"))
    pp.append(_run("1"))
    pp.append(_fldchar("end"))
    pp.append(_run(f". {tail}"))
    return p


def _build_lists(doc):
    # CR-LISTS: realistic two-level bullet list (two L1 each with two L2
    # children) + a numbered rollout sequence. Same style ids and structure as
    # before, now exercising the L2 ilvl=1 binding (BUG-LIST-ILVL) properly.
    _p(doc, "Brand operating principles", "Heading1")
    _p(doc, "Consistency before customization", "BrandDocsBulletL1")
    _p(doc, "One palette: navy, teal, amber", "BrandDocsBulletL2")
    _p(doc, "Type scale fixed across all templates", "BrandDocsBulletL2")
    _p(doc, "Templates are contracts, not suggestions", "BrandDocsBulletL1")
    _p(doc, "Cover, contents, and indices are preserved", "BrandDocsBulletL2")
    _p(doc, "Body content is regenerated per request", "BrandDocsBulletL2")
    _p(doc, "Rollout sequence", "Heading2")
    _p(doc, "Extract the template surface and brand profile", "BrandDocsNumberL1")
    _p(doc, "Review the captured slots and index front matter", "BrandDocsNumberL1")
    _p(doc, "Generate the branded document", "BrandDocsNumberL1")


def _build_table(doc):
    # CR-TABLE: header + four quarters of internally-consistent synthetic data.
    # BUG-TABLE-NUMFMT: one currency style ($X.XXM) and one percent style
    # (+X.X%) across every row, so the columns read as formatted figures.
    _p(doc, "BrandDocs quarterly revenue", "Heading2")
    table = doc.add_table(rows=5, cols=4)
    table.style = "BrandDocs Table"
    # Tell Word which conditional formats to apply (first row + banding).
    tblPr = table._tbl.tblPr
    _sub(
        tblPr,
        "tblLook",
        firstRow="1",
        lastRow="0",
        firstColumn="0",
        lastColumn="0",
        noHBand="0",
        noVBand="1",
    )
    hdr = ("Quarter", "Revenue", "Growth", "Region")
    for c, label in zip(table.rows[0].cells, hdr):
        c.text = label
    data = [
        ("Q1 2026", "$3.20M", "+8.4%", "North"),
        ("Q2 2026", "$3.48M", "+8.8%", "South"),
        ("Q3 2026", "$3.91M", "+12.4%", "EMEA"),
        ("Q4 2026", "$4.25M", "+8.7%", "APAC"),
    ]
    for r, row in enumerate(data, start=1):
        for c, val in zip(table.rows[r].cells, row):
            c.text = val
    _seq_caption(
        doc,
        "Table",
        "Table",
        "BrandDocs FY2026 quarterly revenue by region (synthetic data).",
    )


def _build_figure(doc, logo_rid, logo_cx, logo_cy):
    _p(doc, "BrandDocs wordmark", "Heading2")
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    run._r.append(
        _inline_drawing(logo_rid, logo_cx, logo_cy, "BrandDocsLogoFigure", 100)
    )
    _seq_caption(doc, "Figure", "Figure", "BrandDocs Corp wordmark (synthetic).")


def _build_callout(doc):
    # CR-CALLOUT: an on-brand synthetic note (not a meta tooling instruction).
    _p(
        doc,
        "BrandDocs Corp is a synthetic, fictional company used to demonstrate an "
        "on-brand internal brief. All figures, names, and regions in this template "
        "are illustrative.",
        "BrandDocsCallout",
    )


def _build_editorial_quote(doc):
    _p(
        doc,
        "Design rule: preserve proven structure only while it improves the final "
        "artifact. If visual QA shows blank space, stale content, or broken flow, "
        "repair the composition rather than copying the template blindly.",
        "BrandDocsQuote",
    )


def _build_risk_matrix(doc):
    _p(doc, "Readiness and risk matrix", "Heading2")
    doc.add_paragraph(
        "The matrix below gives the template a realistic governance surface: it "
        "mixes status, owner, signal, and mitigation text in a branded table."
    )
    table = doc.add_table(rows=5, cols=4)
    table.style = "BrandDocs Table"
    tblPr = table._tbl.tblPr
    _sub(
        tblPr,
        "tblLook",
        firstRow="1",
        lastRow="0",
        firstColumn="0",
        lastColumn="0",
        noHBand="0",
        noVBand="1",
    )
    headers = ("Area", "Signal", "Owner", "Mitigation")
    for c, label in zip(table.rows[0].cells, headers):
        c.text = label
    rows = [
        ("Visual QA", "Renderer drift", "Platform", "Smoke-test DOCX/PPTX/XLSX"),
        (
            "Comprehension",
            "Missing inventory",
            "Model Ops",
            "Fail closed on empty refs",
        ),
        ("Brand fidelity", "Style fallback", "Design", "Re-assert captured styles"),
        ("Delivery", "Stale cache", "Engineering", "Refresh visible field results"),
    ]
    for r, row in enumerate(rows, start=1):
        for c, val in zip(table.rows[r].cells, row):
            c.text = val
    _set_col_widths(table, (1980, 1980, 1800, 3600))
    _seq_caption(
        doc, "Table", "Table", "BrandDocs risk and readiness matrix (synthetic)."
    )


def _build_footnote_paragraph(doc, footnote_id):
    _p(doc, "BrandDocs footnote demo", "Heading2")
    body = doc.add_paragraph("BrandDocs Corp")
    body.add_run(" is a registered placeholder brand")
    # Append a footnoteReference run (id -> footnotes.xml).
    fr = _sub(body._p, "r")
    frpr = _sub(fr, "rPr")
    _sub(frpr, "rStyle", val="FootnoteReference")
    _sub(fr, "footnoteReference", id=str(footnote_id))
    body.add_run(" used throughout this template.")


def _build_demo_body(doc):
    # CR-DEMO-BODY: keep the H1 + para + H2 + para demo region, but replace the
    # lorem filler with readable synthetic prose. detect_demo_region keys on the
    # first body-region Heading-1 structurally (style id + captured own text),
    # so demo_region.present stays True with realistic copy.
    _p(doc, "Overview", "Heading1")
    doc.add_paragraph(
        "This brief summarizes how the BrandDocs Corp brand system is applied "
        "across internal documents. It exists to show a complete, on-brand "
        "template; a generation run replaces this body with the requested content."
    )
    _p(doc, "Methodology", "Heading2")
    doc.add_paragraph(
        "Figures are drawn from a synthetic operations dataset maintained by the "
        "BrandDocs brand office. Revenue, growth, and regional splits are "
        "illustrative and should not be read as real performance data."
    )
    # A real external hyperlink (w:hyperlink -> external relationship): a near-
    # universal template element, and one the generator now authors natively too.
    link_p = doc.add_paragraph("Full brand guidelines: ")
    _add_external_hyperlink(
        doc, link_p, "https://example.com/brand", "brand.example.com"
    )


# ---------------------------------------------------------------------------
# DRAWINGS - an inline picture (figure) and a header logo, both referencing the
# same image part by relationship id.
# ---------------------------------------------------------------------------
def _inline_drawing(rid, cx, cy, name, doc_pr_id):
    drawing = _el("drawing")
    inline = etree.SubElement(drawing, f"{{{WP}}}inline")
    inline.set("distT", "0")
    inline.set("distB", "0")
    inline.set("distL", "0")
    inline.set("distR", "0")
    ext = etree.SubElement(inline, f"{{{WP}}}extent")
    ext.set("cx", str(cx))
    ext.set("cy", str(cy))
    eff = etree.SubElement(inline, f"{{{WP}}}effectExtent")
    for k in ("l", "t", "r", "b"):
        eff.set(k, "0")
    docpr = etree.SubElement(inline, f"{{{WP}}}docPr")
    docpr.set("id", str(doc_pr_id))
    docpr.set("name", name)
    cnv = etree.SubElement(inline, f"{{{WP}}}cNvGraphicFramePr")
    locks = etree.SubElement(cnv, f"{{{A}}}graphicFrameLocks")
    locks.set("noChangeAspect", "1")
    graphic = etree.SubElement(inline, f"{{{A}}}graphic")
    gdata = etree.SubElement(graphic, f"{{{A}}}graphicData")
    gdata.set("uri", PIC)
    pic = etree.SubElement(gdata, f"{{{PIC}}}pic")
    nvpic = etree.SubElement(pic, f"{{{PIC}}}nvPicPr")
    cnvpr = etree.SubElement(nvpic, f"{{{PIC}}}cNvPr")
    cnvpr.set("id", "0")
    cnvpr.set("name", name)
    etree.SubElement(nvpic, f"{{{PIC}}}cNvPicPr")
    blipfill = etree.SubElement(pic, f"{{{PIC}}}blipFill")
    blip = etree.SubElement(blipfill, f"{{{A}}}blip")
    blip.set(f"{{{R}}}embed", rid)
    stretch = etree.SubElement(blipfill, f"{{{A}}}stretch")
    etree.SubElement(stretch, f"{{{A}}}fillRect")
    sppr = etree.SubElement(pic, f"{{{PIC}}}spPr")
    xfrm = etree.SubElement(sppr, f"{{{A}}}xfrm")
    off = etree.SubElement(xfrm, f"{{{A}}}off")
    off.set("x", "0")
    off.set("y", "0")
    extb = etree.SubElement(xfrm, f"{{{A}}}ext")
    extb.set("cx", str(cx))
    extb.set("cy", str(cy))
    geom = etree.SubElement(sppr, f"{{{A}}}prstGeom")
    geom.set("prst", "rect")
    etree.SubElement(geom, f"{{{A}}}avLst")
    return drawing


def _build_header_footer(doc, logo_rid, cx, cy):
    section = doc.sections[0]
    header = section.header
    header.is_linked_to_previous = False
    hp = header.paragraphs[0]
    hp.text = ""
    run = hp.add_run()
    # The header logo references the header part's OWN relationship to the image.
    run._r.append(_inline_drawing(logo_rid, cx, cy, "BrandDocsHeaderLogo", 200))

    footer = section.footer
    footer.is_linked_to_previous = False
    fp = footer.paragraphs[0]
    fp.text = "Page "
    fpp = fp._p
    fpp.append(_fldchar("begin"))
    fpp.append(_run(" PAGE ", instr=True))
    fpp.append(_fldchar("separate"))
    fpp.append(_run("1"))
    fpp.append(_fldchar("end"))


def _relate_image_to(part, image_blob, partname="media/image1.png"):
    """Ensure ``part`` (document or header) relates to a shared image part.

    The image part is added once to the package; each consuming part gets its own
    relationship id back.
    """
    from docx.opc.part import Part
    from docx.opc.packuri import PackURI

    package = part.package
    puri = PackURI("/word/" + partname)
    image_part = None
    for p in package.iter_parts():
        if p.partname == puri:
            image_part = p
            break
    if image_part is None:
        image_part = Part(puri, "image/png", image_blob, package)
    rid = part.relate_to(image_part, f"{R}/image")
    return rid


# ---------------------------------------------------------------------------
# SECTIONS - convert the document into a portrait section followed by a
# landscape section.
# ---------------------------------------------------------------------------
def _set_col_widths(table, widths_twips):
    """Pin explicit per-column widths (Twips) on a table so a wide landscape
    table fits the usable page width without column clipping (PL-PRINT-MARGINS).

    Disables autofit and stamps ``w:tcW`` (dxa) on every cell of every column,
    which is how Word honours fixed column widths.
    """
    table.autofit = False
    table.allow_autofit = False
    tblPr = table._tbl.tblPr
    layout = _sub(tblPr, "tblLayout")
    layout.set(_w("type"), "fixed")
    for row in table.rows:
        for cell, w_tw in zip(row.cells, widths_twips):
            cell.width = Twips(w_tw)


def _add_landscape_section(doc, curve_rid=None, fig_cx=0, fig_cy=0):
    new_section = doc.add_section(WD_SECTION.NEW_PAGE)
    new_section.orientation = WD_ORIENT.LANDSCAPE
    # Swap page width/height for landscape (python-docx does not auto-swap).
    new_section.page_width = Twips(15840)  # 11"
    new_section.page_height = Twips(12240)  # 8.5"
    _p(doc, "BrandDocs landscape appendix", "Heading1")
    doc.add_paragraph(
        "This appendix sits in a landscape section. It carries a wide program "
        "rollout matrix and a second brand figure that the portrait body cannot "
        "fit. Demo content the generator may clear."
    )

    # PL-LANDSCAPE-CONTENT: a wide 7-column synthetic "program rollout matrix",
    # reusing the same custom table style + tblLook (banding exercised again).
    _p(doc, "Program rollout matrix", "Heading2")
    wide = doc.add_table(rows=5, cols=7)
    wide.style = "BrandDocs Table"
    wtblPr = wide._tbl.tblPr
    _sub(
        wtblPr,
        "tblLook",
        firstRow="1",
        lastRow="0",
        firstColumn="0",
        lastColumn="0",
        noHBand="0",
        noVBand="1",
    )
    whdr = ("Workstream", "Owner", "Q1", "Q2", "Q3", "Q4", "Status")
    for c, label in zip(wide.rows[0].cells, whdr):
        c.text = label
    wdata = [
        (
            "Template surface",
            "Brand Office",
            "Scope",
            "Build",
            "QA",
            "Ship",
            "On track",
        ),
        (
            "Profile extraction",
            "Platform",
            "Spec",
            "Build",
            "Verify",
            "Tune",
            "On track",
        ),
        (
            "Generation engine",
            "Platform",
            "Design",
            "Build",
            "Build",
            "Verify",
            "At risk",
        ),
        (
            "Rollout & training",
            "Enablement",
            "Plan",
            "Draft",
            "Pilot",
            "Launch",
            "Planned",
        ),
    ]
    for r, row in enumerate(wdata, start=1):
        for c, val in zip(wide.rows[r].cells, row):
            c.text = val
    # PL-PRINT-MARGINS: fixed column widths sum to 13680 twips (15840 - 2*1in
    # margins), so the 7-column table fits the landscape usable width.
    _set_col_widths(wide, (2880, 2400, 1680, 1680, 1680, 1680, 1680))
    _seq_caption(
        doc, "Table", "Table", "BrandDocs FY2026 program rollout matrix (synthetic)."
    )

    # PL-LANDSCAPE-CAPTION-SEQ / CR-FIG2-CURVE: a real SECOND inline figure so the
    # cached "Figure 2" in the Table of Figures corresponds to a rendered figure.
    # This is a DISTINCT media part (word/media/image2.png, the growth-curve PNG),
    # NOT a reuse of the logo - Figure 2 is a real rising growth curve, not the
    # wordmark. The curve PNG is ~12:5, so the inline extent matches that aspect.
    if curve_rid is not None:
        _p(doc, "BrandDocs growth curve", "Heading2")
        fp = doc.add_paragraph()
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        frun = fp.add_run()
        frun._r.append(
            _inline_drawing(curve_rid, fig_cx, fig_cy, "BrandDocsGrowthFigure", 300)
        )
        _seq_caption(
            doc,
            "Figure",
            "Figure",
            "BrandDocs Corp growth curve (synthetic illustration).",
        )
    return new_section


# ---------------------------------------------------------------------------
# settings.xml - footnotePr + a docId-free, deterministic settings part already
# exists; we just request field update on open so cached TOC/SEQ recompute.
# ---------------------------------------------------------------------------
def _request_update_fields(doc):
    settings = doc.settings.element
    if settings.find(_w("updateFields")) is None:
        uf = _el("updateFields", val="true")
        settings.insert(0, uf)


def build(out: Path = OUT) -> Path:
    doc = Document()  # python-docx default template (single portrait section)

    # 1) Styles (paragraph + list + table) authored into styles.xml.
    _build_styles(doc)

    # 2) Numbering: the default template already ships an (empty) numbering part,
    # already related from document.xml. Populate it in place (a duplicate part
    # of the same name would corrupt the zip).
    _populate_numbering(doc.part.numbering_part.element)

    # 3) Footnotes part attached; footnote id 2 is the authored note.
    _attach_part(
        doc,
        "word/footnotes.xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
        _build_footnotes_xml(),
        f"{R}/footnotes",
    )

    # 4) Shared logo image: a generated text-only BrandDocs wordmark. The SAME
    # wordmark every BrandDocs template embeds (CR-LOGO-MARK). Related to the
    # document part (for Figure 1) and to the header part (for the header logo)
    # independently. A SECOND media part carries the real growth-curve figure
    # (CR-FIG2-CURVE).
    logo_blob = branddocs_mark_png(640, 160)  # wide wordmark (image1.png)
    curve_blob = branddocs_curve_png(480, 200)  # ~12:5 growth curve (image2.png)
    doc_logo_rid = _relate_image_to(doc.part, logo_blob)
    doc_curve_rid = _relate_image_to(doc.part, curve_blob, partname="media/image2.png")

    # 5) Cover (block-level SDT title + placeholder slots), then the three index
    # fields, then the demo body. add_* appends in document order, so build the
    # front matter first, then the body.
    _build_cover(doc)
    _build_cover_scorecard(doc)
    _build_index_front_matter(doc)

    # Body content (freeform; a generation clears this region).
    _build_lists(doc)
    _build_table(doc)
    # Figure 1 (the wordmark) uses the document-part image relationship. Preserve
    # its 4:1 aspect ratio so the text never stretches.
    fig_cx, fig_cy = 2743200, 685800
    _build_figure(doc, doc_logo_rid, fig_cx, fig_cy)
    _build_callout(doc)
    _build_editorial_quote(doc)
    _build_risk_matrix(doc)
    _build_footnote_paragraph(doc, footnote_id=2)
    _build_demo_body(doc)

    # 6) Header (logo) + footer (PAGE field). The header part needs its OWN
    # relationship to the shared image part.
    section = doc.sections[0]
    header_logo_rid = _relate_image_to(section.header.part, logo_blob)
    # Header logo is the SAME 4:1 wordmark, kept compact and undistorted.
    _build_header_footer(doc, header_logo_rid, 914400, 228600)

    # 7) A second, landscape section after the portrait body. It carries a wide
    # rollout-matrix table and a real second figure (the cached "Figure 2"),
    # which is the DISTINCT growth-curve media part (word/media/image2.png).
    # The curve PNG is 480x200 (~12:5), so the extent matches: cx=2286000 EMU
    # (2.5in), cy=2286000*5//12=952500 EMU (exact 12:5, no distortion).
    land_fig_cx, land_fig_cy = 2286000, 952500
    _add_landscape_section(
        doc, curve_rid=doc_curve_rid, fig_cx=land_fig_cx, fig_cy=land_fig_cy
    )

    # 8) Ask Word to refresh the cached fields on open.
    _request_update_fields(doc)

    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(out)
    freeze_ooxml(out)
    return out


if __name__ == "__main__":
    path = build()
    print(f"built {path}")
