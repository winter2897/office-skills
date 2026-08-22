#!/usr/bin/env python3
"""Report builder - any organisation, one template.

    python3 report.py template [out.docx]            # blank template with {{PLACEHOLDER}}
    python3 report.py write content.json [out.docx]  # filled report
    python3 report.py selfcheck                      # verify the script still works

Branding lives in assets/brand.json - company name, logo, cover colour. Replace
the logo file, change the name, done; no code edit.

Both modes go through the same renderer, so a written report is structurally
identical to the template. Requires: pip3 install python-docx
"""
import json
import os
import tempfile
import re
import sys

try:
    from docx import Document
except ImportError:                    # the one dependency, and the one fix
    sys.exit("error: python-docx is not installed. Run:\n\n"
             "    pip3 install python-docx\n")
from docx.enum.section import WD_SECTION
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import (WD_ALIGN_VERTICAL, WD_ROW_HEIGHT_RULE,
                             WD_TABLE_ALIGNMENT)
from docx.enum.text import (WD_ALIGN_PARAGRAPH, WD_BREAK, WD_TAB_ALIGNMENT,
                            WD_TAB_LEADER)
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Pt, RGBColor, Twips

# realpath, not abspath: the usual install symlinks this skill into
# ~/.claude/skills, and the shared assets/ sits above the real file.
HERE = os.path.dirname(os.path.realpath(__file__))
EXAMPLE = os.path.join(os.path.dirname(HERE), "example.json")
# The sibling skills that render through this script. Their JSON is
# documentation, and documentation drifts, so selfcheck builds it too.
SIBLINGS = [os.path.join(os.path.dirname(os.path.dirname(HERE)), "test", name)
            for name in ("skeleton.json", "example.json")]


def find_assets(start=HERE, levels=3):
    """The shared assets/ directory, searched upwards from this script.

    Branding belongs to the organisation, not to one skill, so it lives at the
    root of the repository and every skill under it reads the same logo. Walk
    up until an assets/brand.json turns up; the nearest one wins, so a skill can
    still carry its own. Nothing found - the skill was copied out on its own -
    means brand defaults and a company name set in text instead of a logo.
    """
    d = start
    for _ in range(levels):
        d = os.path.dirname(d)
        cand = os.path.join(d, "assets")
        if os.path.isfile(os.path.join(cand, "brand.json")):
            return cand
    return None


ASSETS = find_assets()
BRAND_FILE = os.path.join(ASSETS, "brand.json") if ASSETS else None

# The fonts Office installs on both Windows and macOS, all of them carrying
# Vietnamese diacritics. Anything outside this list is not refused to be
# difficult: .docx stores the font name, not the font, so a reader without it
# gets a silent substitute and a layout nobody proof-read.
FONTS = ["Times New Roman", "Arial", "Calibri", "Cambria", "Georgia", "Tahoma",
         "Trebuchet MS", "Verdana"]
DEFAULT_FONT = FONTS[0]
FONT = DEFAULT_FONT   # per document; "font" in content.json overrides it
MONO = "Consolas"
DEFAULT_BODY_PT = 13  # Vietnamese report standard (TT 01/2011/TT-BNV: TNR 13-14pt)
BODY_PT = DEFAULT_BODY_PT   # per document; "body_pt" in content.json overrides it
DASH = "–"            # en dash everywhere; never the longer em dash

INK = RGBColor(0x1A, 0x1A, 0x1A)      # logo black
MUTED = RGBColor(0x59, 0x59, 0x59)    # secondary text
CHIP = "DDDDDD"                        # cover version chip
RULE = "8C8C8C"                        # cover rule under the title
LINE = "BFBFBF"                        # rules / table borders
BAND = "F2F2F2"                        # table header fill
CODE_BG = "F7F7F7"
WARN = RGBColor(0xC0, 0x00, 0x00)

# Everything that changes when the skill moves to another organisation. The
# shipped assets/brand.json is a working sample - replace the logo files and the
# company name and nothing else has to change. "brand" in content.json overrides
# any of these for a single report.
BRAND_DEFAULTS = {
    "company": "Your Company",
    "logo": "logo.png",            # header mark; relative to assets/, or absolute
    "mark": "logo-mark.png",       # cover mark, bottom right
    "logo_box_cm": [2.9, 1.2],     # the logo is fitted inside this box, not stretched
    "mark_box_cm": [1.6, 1.6],
    "cover_bg": "F4F4F4",          # cover ground, hex - generated, not a shipped file
    "ink": "1A1A1A",               # body ink
    "lang": "en",                  # furniture language when content.json is silent
    "font": DEFAULT_FONT,
    "doc_no_example": "REP-2026-001",
    "email_example": "name@company.com",
}
BRAND = dict(BRAND_DEFAULTS)   # per document; build() reloads it
UPPER_H1 = True                # formal layout shouts level 1; simple does not

# width, height, left, right, top, bottom - centimetres. The generous A4 top
# margin leaves room for the logo header; Letter keeps the 1 in margins its
# documents are written to.
PAGES = {
    "a4":     (21.0, 29.7, 2.0, 2.0, 3.5, 2.5),
    "letter": (21.59, 27.94, 2.5, 2.5, 2.5, 2.5),
}
PAGE = "a4"       # per document; "page" in content.json overrides it
TWIPS_PER_CM = 1440 / 2.54


def content_width(page):
    """Twips between the margins - what a table's widths must add up to."""
    w, _, left, right, _, _ = PAGES[page]
    return int(round((w - left - right) * TWIPS_PER_CM))


CONTENT_W = content_width(PAGE)   # rebuilt by build() from the page
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

# Fixed document furniture. Report prose is Vietnamese by default; set
# "lang": "en" in content.json for a foreign-customer report.
LABELS = {
    "vi": {
        "lang_tag": "vi-VN",
        "document_type": "Báo cáo",
        "object": "Đối tượng",
        "doc_no": "Mã số tài liệu",
        "version": "Phiên bản",
        "date": "Ngày ban hành",
        "prepared_by": "Người lập",
        "tested_by": "Người thực hiện thử nghiệm",
        "reviewed_by": "Người kiểm tra",
        "approved_by": "Người phê duyệt",
        "classification": "Mức bảo mật",
        "classification_default": "Nội bộ",
        "doc_control": "THÔNG TIN TÀI LIỆU",
        "author": "Tác giả",
        "revision_history": "LỊCH SỬ PHIÊN BẢN",
        "revisions": "Lịch sử phiên bản",
        "rev": "Phiên bản",
        "rev_date": "Ngày",
        "rev_desc": "Nội dung thay đổi",
        "rev_author": "Người thực hiện",
        "toc_title": "MỤC LỤC",
        "lof_title": "DANH MỤC HÌNH VẼ",
        "lot_title": "DANH MỤC BẢNG BIỂU",
        "abbr_title": "DANH MỤC KÝ HIỆU VÀ CHỮ VIẾT TẮT",
        "abbr_short": "Từ viết tắt",
        "abbr_full": "Tên đầy đủ",
        "abbr_meaning": "Diễn giải",
        "references_title": "Tài liệu tham khảo",
        "page": "Trang",
        "figure": "Hình",
        "figure_slot": "[ chèn hình ]",
        "table": "Bảng",
        "listing": "Đoạn mã",
        "hdr_doc_no": "Mã số",
        "hdr_rev": "Phiên bản",
        "project": "Dự án",
        "title_prefix": "TIÊU ĐỀ",
    },
    "en": {
        "lang_tag": "en-US",
        "document_type": "Report",
        "object": "Object / Subject",
        "doc_no": "Document No.",
        "version": "Version",
        "date": "Date",
        "prepared_by": "Prepared by",
        "tested_by": "Testing conducted by",
        "reviewed_by": "Reviewed by",
        "approved_by": "Approved by",
        "classification": "Classification",
        "classification_default": "Confidential",
        "doc_control": "DOCUMENT INFORMATION",
        "author": "Author",
        "revision_history": "REVISION HISTORY",
        "revisions": "Revisions",
        "rev": "Rev",
        "rev_date": "Date",
        "rev_desc": "Description",
        "rev_author": "Author",
        "toc_title": "TABLE OF CONTENTS",
        "lof_title": "LIST OF FIGURES",
        "lot_title": "LIST OF TABLES",
        "abbr_title": "ABBREVIATIONS",
        "abbr_short": "Abbreviation",
        "abbr_full": "Full term",
        "abbr_meaning": "Meaning",
        "references_title": "References",
        "page": "Page",
        "figure": "Figure",
        "figure_slot": "[ insert figure ]",
        "table": "Table",
        "listing": "Listing",
        "hdr_doc_no": "Doc. No",
        "hdr_rev": "Rev",
        "project": "Project Reference",
        "title_prefix": "TITLE",
    },
}


def load_brand(spec=None):
    """assets/brand.json over the defaults, then "brand" in content.json."""
    brand = dict(BRAND_DEFAULTS)
    if BRAND_FILE:
        with open(BRAND_FILE, encoding="utf-8") as fh:
            layer = json.load(fh)
        extra = set(layer) - set(BRAND_DEFAULTS)
        if extra:
            raise ValueError("brand.json: unknown key(s) %s - known: %s"
                             % (", ".join(sorted(extra)),
                                ", ".join(sorted(BRAND_DEFAULTS))))
        brand.update(layer)
    if spec:
        brand.update(spec.get("brand") or {})
    return brand


def brand_asset(key):
    """A brand image path, resolved against assets/ unless already absolute."""
    path = BRAND.get(key)
    if not path:
        return None
    if os.path.isabs(path):
        return path
    return os.path.join(ASSETS, path) if ASSETS else None


# --------------------------------------------------------------------------- xml helpers
def el(tag, **attrs):
    e = OxmlElement(tag)
    for k, v in attrs.items():
        e.set(qn("w:" + k), str(v))
    return e


# ECMA-376 child order – appending out of order makes Word reject the file.
PPR_ORDER = ["w:pStyle", "w:keepNext", "w:keepLines", "w:pageBreakBefore", "w:framePr",
             "w:widowControl", "w:numPr", "w:suppressLineNumbers", "w:pBdr", "w:shd",
             "w:tabs", "w:suppressAutoHyphens", "w:kinsoku", "w:wordWrap",
             "w:overflowPunct", "w:topLinePunct", "w:autoSpaceDE", "w:autoSpaceDN",
             "w:bidi", "w:adjustRightInd", "w:snapToGrid", "w:spacing", "w:ind",
             "w:contextualSpacing", "w:mirrorIndents", "w:suppressOverlap", "w:jc",
             "w:textDirection", "w:textAlignment", "w:textboxTightWrap",
             "w:outlineLvl", "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr", "w:pPrChange"]
RPR_ORDER = ["w:rStyle", "w:rFonts", "w:b", "w:bCs", "w:i", "w:iCs", "w:caps",
             "w:smallCaps", "w:strike", "w:dstrike", "w:outline", "w:shadow",
             "w:emboss", "w:imprint", "w:noProof", "w:snapToGrid", "w:vanish",
             "w:webHidden", "w:color", "w:spacing", "w:w", "w:kern", "w:position",
             "w:sz", "w:szCs", "w:highlight", "w:u", "w:effect", "w:bdr", "w:shd",
             "w:fitText", "w:vertAlign", "w:rtl", "w:cs", "w:em", "w:lang",
             "w:eastAsianLayout", "w:specVanish", "w:oMath"]
SECTPR_ORDER = ["w:footnotePr", "w:endnotePr", "w:type", "w:pgSz", "w:pgMar",
                "w:paperSrc", "w:pgBorders", "w:lnNumType", "w:pgNumType", "w:cols",
                "w:formProt", "w:vAlign", "w:noEndnote", "w:titlePg", "w:textDirection",
                "w:bidi", "w:rtlGutter", "w:docGrid", "w:printerSettings"]


def insert_ordered(parent, child, order):
    idx = order.index("w:" + child.tag.split("}")[1])
    for existing in parent:
        etag = "w:" + existing.tag.split("}")[1]
        if etag in order and order.index(etag) > idx:
            existing.addprevious(child)
            return child
    parent.append(child)
    return child


def text_el(tag, text):
    t = OxmlElement(tag)
    t.text = text
    t.set(XML_SPACE, "preserve")
    return t


def field(paragraph, instr, placeholder="", dirty=False):
    """A complex field: begin / instruction / cached result / end."""
    r = paragraph.add_run()
    fc = el("w:fldChar", fldCharType="begin")
    if dirty:
        fc.set(qn("w:dirty"), "true")
    r._r.append(fc)
    paragraph.add_run()._r.append(text_el("w:instrText", instr))
    paragraph.add_run()._r.append(el("w:fldChar", fldCharType="separate"))
    if placeholder:
        paragraph.add_run(placeholder)
    paragraph.add_run()._r.append(el("w:fldChar", fldCharType="end"))


def field_start(instr):
    """The begin/instruction/separate runs of a field, as bare elements."""
    begin = OxmlElement("w:r")
    fc = el("w:fldChar", fldCharType="begin")
    fc.set(qn("w:dirty"), "true")
    begin.append(fc)
    code = OxmlElement("w:r")
    code.append(text_el("w:instrText", instr))
    sep = OxmlElement("w:r")
    sep.append(el("w:fldChar", fldCharType="separate"))
    return [begin, code, sep]


def simple_field(paragraph, instr, cached):
    """A one-shot field (SEQ, STYLEREF) carrying a cached result for viewers."""
    f = el("w:fldSimple", instr=instr)
    r = OxmlElement("w:r")
    r.append(text_el("w:t", cached))
    f.append(r)
    paragraph._p.append(f)
    return f


def borders(pPr, edges, sz=8, color=LINE):
    b = OxmlElement("w:pBdr")
    for edge in edges:
        b.append(el("w:" + edge, val="single", sz=sz, space=1, color=color))
    return insert_ordered(pPr, b, PPR_ORDER)


def shade(tcPr, fill):
    tcPr.append(el("w:shd", val="clear", color="auto", fill=fill))


# A picture pinned to page coordinates instead of flowing with the text – Word
# has no page layout, so the cover art is three of these sitting behind the type.
ANCHOR = ('<wp:anchor %s behindDoc="1" distT="0" distB="0" distL="0" distR="0"'
          ' simplePos="0" relativeHeight="%d" locked="1" layoutInCell="0"'
          ' allowOverlap="1"><wp:simplePos x="0" y="0"/>'
          '<wp:positionH relativeFrom="page"><wp:posOffset>%d</wp:posOffset>'
          '</wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>%d'
          '</wp:posOffset></wp:positionV><wp:extent cx="%d" cy="%d"/>'
          '<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>'
          '<wp:docPr id="%d" name="cover art %d"/><wp:cNvGraphicFramePr/></wp:anchor>')
_anchor_id = [900]


def float_picture(paragraph, path, width, x, y):
    """Place a picture at (x, y) from the top-left page corner, behind the text."""
    _anchor_id[0] += 1
    n = _anchor_id[0]
    run = paragraph.add_run()
    run.add_picture(path, width=width)
    inline = run._r.find(qn("w:drawing"))[0]
    anchor = parse_xml(ANCHOR % (nsdecls("wp"), n, int(x), int(y),
                                 int(inline.extent.cx), int(inline.extent.cy), n, n))
    anchor.append(inline.graphic)
    inline.getparent().replace(inline, anchor)


def solid_png(path, hex_color, w=210, h=297):
    """A flat A4-shaped PNG in one colour - the cover ground.

    Written here rather than shipped so the cover colour is a hex string in
    brand.json instead of a binary somebody has to regenerate with rsvg.
    """
    import struct
    import zlib
    rgb = bytes.fromhex(hex_color.lstrip("#"))
    if len(rgb) != 3:
        raise ValueError("cover_bg must be a 6-digit hex colour, got %r" % hex_color)
    raw = (b"\x00" + rgb * w) * h          # one filter byte per scanline

    def chunk(tag, data):
        body = tag + data
        return struct.pack(">I", len(data)) + body + struct.pack(">I", zlib.crc32(body))

    with open(path, "wb") as fh:
        fh.write(b"\x89PNG\r\n\x1a\n"
                 + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0))
                 + chunk(b"IDAT", zlib.compress(raw, 9))
                 + chunk(b"IEND", b""))
    return path


def place_logo(paragraph, path, box_cm, where):
    """Fit a logo inside box_cm keeping its aspect ratio. False if it cannot.

    add_picture with a width alone lets Word derive the height, so a square
    logo becomes as tall as it is wide and inflates the header on every page
    of the document. Fit to a box instead, and let the narrow dimension bind.
    """
    from docx.image.image import Image
    if not path:
        sys.stderr.write("warning: no %s image - assets/brand.json was not found "
                         "or names none\n" % where)
        return False
    if not os.path.exists(path):
        sys.stderr.write("warning: %s image not found: %s\n" % (where, path))
        return False
    try:
        img = Image.from_file(path)
    except Exception as exc:               # SVG, a truncated file, anything
        sys.stderr.write("warning: cannot read the %s image (%s: %s) - convert it "
                         "to PNG or JPEG\n" % (where, os.path.basename(path), exc))
        return False

    if max(img.px_width, img.px_height) > 2000:
        sys.stderr.write("warning: %s is %dx%d px - downscale it, the .docx "
                         "carries the pixels it is given\n"
                         % (os.path.basename(path), img.px_width, img.px_height))
    ratio = img.width / float(img.height)
    if ratio > 8 or ratio < 1 / 3.0:
        sys.stderr.write("warning: %s has a %.1f:1 aspect ratio and will render "
                         "small inside a %gx%g cm box - widen it with "
                         "logo_box_cm / mark_box_cm in brand.json\n"
                         % (os.path.basename(path), ratio, box_cm[0], box_cm[1]))

    scale = min(Cm(box_cm[0]) / float(img.width), Cm(box_cm[1]) / float(img.height))
    paragraph.add_run().add_picture(path, width=int(img.width * scale),
                                    height=int(img.height * scale))
    return True


def run_shade(run, fill):
    """Shade one run – a coloured chip around the text, without a table."""
    insert_ordered(run._r.get_or_add_rPr(),
                   el("w:shd", val="clear", color="auto", fill=fill), RPR_ORDER)


def set_size(paragraph, points, color=MUTED):
    for run in paragraph.runs:
        run.font.name = FONT
        run.font.size = Pt(points)
        run.font.color.rgb = color


CODE_SPAN = re.compile(r"`([^`\n]+)`")


def rich_text(p, text, bold=False, italic=False, size=None, color=None):
    """`lệnh --cờ` giữa hai backtick becomes a monospace run. That is the only markup."""
    for i, chunk in enumerate(CODE_SPAN.split(str(text))):
        if not chunk:
            continue
        r = p.add_run(chunk)
        if bold:
            r.bold = True
        if italic:
            r.italic = True
        if color:
            r.font.color.rgb = color
        if i % 2:                                  # odd chunks were inside backticks
            r.font.size = Pt((size or BODY_PT) - 1)
            force_font(r._element.get_or_add_rPr(), MONO)
        elif size:
            r.font.size = Pt(size)
    return p


def cell_text(cell, text, bold=False, size=11, color=None, align=None):
    cell.text = ""
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    if align is not None:
        p.alignment = align
    return rich_text(p, text, bold=bold, size=size,
                     color=INK if color is None else color)


def para(doc, text="", style=None, size=None, bold=False, italic=False,
         color=None, align=None, before=None, after=None):
    p = doc.add_paragraph(style=style)
    if align is not None:
        p.alignment = align
    if before is not None:
        p.paragraph_format.space_before = Pt(before)
    if after is not None:
        p.paragraph_format.space_after = Pt(after)
    if text:
        r = p.add_run(text)
        r.bold = bold
        r.italic = italic
        if size:
            r.font.size = Pt(size)
        if color:
            r.font.color.rgb = color
    return p


_bookmark_id = [1000]


def bookmark(paragraph, name):
    """Anchor a paragraph so PAGEREF and hyperlinks can point at it."""
    bid = _bookmark_id[0]
    _bookmark_id[0] += 1
    at = 1 if paragraph._p.find(qn("w:pPr")) is not None else 0
    paragraph._p.insert(at, el("w:bookmarkStart", id=bid, name=name))
    paragraph._p.append(el("w:bookmarkEnd", id=bid))


# --------------------------------------------------------------------------- numbering
HEADING_NUMID, BULLET_NUMID = 90, 91


def install_numbering(doc):
    """Auto-numbered headings (1. / 1.1 / 1.1.1) and bullets."""
    numbering = doc.part.numbering_part.element

    heading = el("w:abstractNum", abstractNumId=900)
    heading.append(el("w:multiLevelType", val="multilevel"))
    for i, txt in enumerate(["%1.", "%1.%2", "%1.%2.%3"]):
        lvl = el("w:lvl", ilvl=i)
        lvl.append(el("w:start", val=1))
        lvl.append(el("w:numFmt", val="decimal"))
        lvl.append(el("w:suff", val="space"))   # "1. TIÊU ĐỀ", not a tab column
        lvl.append(el("w:lvlText", val=txt))
        lvl.append(el("w:lvlJc", val="left"))
        pPr = OxmlElement("w:pPr")
        # Google Docs drops w:suff and always tabs. Without a stop of our own it
        # tabs to the 1.27 cm default and the number floats away from the title,
        # so the hanging indent parks it at 0.6 cm – widening per level, or the
        # longer "1.1.1" would overrun it and fall through to that default.
        pPr.append(el("w:ind", left=340 + 170 * i, hanging=340 + 170 * i))
        lvl.append(pPr)
        heading.append(lvl)

    bullet = el("w:abstractNum", abstractNumId=901)
    bullet.append(el("w:multiLevelType", val="hybridMultilevel"))
    # Escapes, not literals: these bullets live in the private-use area and
    # get eaten by editors and copy-paste, silently leaving the list blank.
    for i, (char, font, left) in enumerate([
        ("\uF0B7", "Symbol", 340),        # •
        ("o", "Courier New", 737),
        ("\uF0A7", "Wingdings", 1134),    # ▪
    ]):
        lvl = el("w:lvl", ilvl=i)
        lvl.append(el("w:start", val=1))
        lvl.append(el("w:numFmt", val="bullet"))
        lvl.append(el("w:lvlText", val=char))
        lvl.append(el("w:lvlJc", val="left"))
        pPr = OxmlElement("w:pPr")
        pPr.append(el("w:ind", left=left, hanging=283))
        lvl.append(pPr)
        rPr = OxmlElement("w:rPr")
        rPr.append(el("w:rFonts", ascii=font, hAnsi=font, hint="default"))
        lvl.append(rPr)
        bullet.append(lvl)

    anchor = numbering.find(qn("w:num"))
    pos = list(numbering).index(anchor) if anchor is not None else len(numbering)
    for a in (bullet, heading):
        numbering.insert(pos, a)

    for num_id, abstract_id in ((HEADING_NUMID, 900), (BULLET_NUMID, 901)):
        n = el("w:num", numId=num_id)
        n.append(el("w:abstractNumId", val=abstract_id))
        numbering.append(n)


# --------------------------------------------------------------------------- styles
def style_numpr(style, ilvl, num_id):
    numPr = OxmlElement("w:numPr")
    numPr.append(el("w:ilvl", val=ilvl))
    numPr.append(el("w:numId", val=num_id))
    insert_ordered(style.element.get_or_add_pPr(), numPr, PPR_ORDER)


def pick_font(name):
    """Resolve the requested font to its canonical name, or refuse it."""
    if not name:
        return DEFAULT_FONT
    for known in FONTS:
        if name.strip().lower() == known.lower():
            return known
    raise ValueError("font %r is not one of: %s" % (name, ", ".join(FONTS)))


def force_font(rPr, name=None):
    """Built-in Word styles reference theme fonts, which beat w:ascii. Strip them."""
    name = name or FONT
    rFonts = rPr.get_or_add_rFonts()
    for attr in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
        if rFonts.get(qn("w:" + attr)) is not None:
            del rFonts.attrib[qn("w:" + attr)]
    for attr in ("ascii", "hAnsi", "eastAsia", "cs"):
        rFonts.set(qn("w:" + attr), name)


def install_styles(doc, lang_tag="vi-VN"):
    styles = doc.styles

    normal = styles["Normal"]
    normal.font.name = FONT
    normal.font.size = Pt(BODY_PT)
    normal.font.color.rgb = INK
    normal.paragraph_format.space_after = Pt(6)
    normal.paragraph_format.line_spacing = 1.15
    nrpr = normal.element.get_or_add_rPr()
    force_font(nrpr)
    insert_ordered(nrpr, el("w:lang", val=lang_tag), RPR_ORDER)

    def base(name, size, bold=False, italic=False, color=INK, before=0, after=6,
             keep_next=False):
        s = styles[name]
        s.font.name = FONT
        s.font.size = Pt(size)
        s.font.bold = bold
        s.font.italic = italic
        s.font.color.rgb = color
        pf = s.paragraph_format
        pf.space_before = Pt(before)
        pf.space_after = Pt(after)
        pf.keep_with_next = keep_next
        pf.line_spacing = 1.15
        force_font(s.element.get_or_add_rPr())
        return s

    t = base("Title", 30, bold=True, after=4)
    t.paragraph_format.keep_with_next = True
    tpr = t.element.get_or_add_pPr()
    for b in tpr.findall(qn("w:pBdr")):      # Word's default Title rule
        tpr.remove(b)

    h1 = base("Heading 1", 16, bold=True, before=20, after=8, keep_next=True)
    if UPPER_H1:
        # The rule separates sections that each own a page. Seven of them down
        # one flowing document is a ladder, not a separation.
        borders(h1.element.get_or_add_pPr(), ["bottom"], sz=6)
    style_numpr(h1, 0, HEADING_NUMID)

    h2 = base("Heading 2", 14, bold=True, before=14, after=6, keep_next=True)
    style_numpr(h2, 1, HEADING_NUMID)

    h3 = base("Heading 3", 13, bold=True, italic=True, color=RGBColor(0x40, 0x40, 0x40),
              before=10, after=4, keep_next=True)
    style_numpr(h3, 2, HEADING_NUMID)

    body = base("Body Text", BODY_PT, after=6)
    body.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    style_numpr(base("List Bullet", BODY_PT, after=4), 0, BULLET_NUMID)
    style_numpr(base("List Bullet 2", BODY_PT, after=4), 1, BULLET_NUMID)

    cap = base("Caption", 11, italic=True, color=MUTED, before=2, after=10)
    cap.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER

    base("Header", 10, color=MUTED, after=0)
    base("Footer", 10, color=MUTED, after=0)
    base("Quote", BODY_PT, italic=True, color=MUTED, before=6, after=6)

    install_unnumbered_headings(doc)
    install_extra_styles(doc)


def install_unnumbered_headings(doc):
    """Kết luận / Phụ lục / Tài liệu tham khảo: no number, still in the TOC."""
    for level in (1, 2, 3):
        s = doc.styles.add_style("Heading %d Unnumbered" % level,
                                 WD_STYLE_TYPE.PARAGRAPH)
        s.base_style = doc.styles["Heading %d" % level]
        s.quick_style = True
        pPr = s.element.get_or_add_pPr()
        numPr = OxmlElement("w:numPr")
        numPr.append(el("w:ilvl", val=0))
        numPr.append(el("w:numId", val=0))       # numId 0 cancels inherited numbering
        insert_ordered(pPr, numPr, PPR_ORDER)
        insert_ordered(pPr, el("w:outlineLvl", val=level - 1), PPR_ORDER)


def install_extra_styles(doc):
    code = doc.styles.add_style("Code", WD_STYLE_TYPE.PARAGRAPH)
    code.base_style = doc.styles["Normal"]
    code.font.name = MONO
    code.font.size = Pt(10)
    code.font.color.rgb = INK
    force_font(code.element.get_or_add_rPr(), MONO)
    pf = code.paragraph_format
    pf.space_before = Pt(1)
    pf.space_after = Pt(1)
    pf.line_spacing = 1.0

    ref = doc.styles.add_style("Reference", WD_STYLE_TYPE.PARAGRAPH)
    ref.base_style = doc.styles["Normal"]
    ref.font.size = Pt(BODY_PT - 1)
    pf = ref.paragraph_format
    pf.left_indent = Twips(680)
    pf.first_line_indent = Twips(-680)   # hanging, so [12] stays in the margin
    pf.space_after = Pt(4)


def install_table_style(doc):
    """Thin grey grid, shaded header row."""
    st = el("w:style", type="table", styleId="ReportTable")
    st.append(el("w:name", val="Report Table"))
    st.append(el("w:basedOn", val="TableNormal"))
    st.append(el("w:uiPriority", val="59"))
    st.append(OxmlElement("w:qFormat"))

    rPr = OxmlElement("w:rPr")
    rPr.append(el("w:rFonts", ascii=FONT, hAnsi=FONT, cs=FONT))
    rPr.append(el("w:sz", val=22))
    rPr.append(el("w:szCs", val=22))
    st.append(rPr)

    tblPr = OxmlElement("w:tblPr")
    tblBorders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        tblBorders.append(el("w:" + edge, val="single", sz=4, space=0, color=LINE))
    tblPr.append(tblBorders)
    mar = OxmlElement("w:tblCellMar")
    for edge, w in (("top", 60), ("left", 108), ("bottom", 60), ("right", 108)):
        mar.append(el("w:" + edge, w=w, type="dxa"))
    tblPr.append(mar)
    st.append(tblPr)

    band = el("w:tblStylePr", type="firstRow")
    bpr = OxmlElement("w:rPr")
    bpr.append(el("w:b", val="1"))
    band.append(bpr)
    btc = OxmlElement("w:tcPr")
    btc.append(el("w:shd", val="clear", color="auto", fill=BAND))
    band.append(btc)
    st.append(band)

    doc.styles.element.append(st)


def set_theme_fonts(doc):
    """Point the theme's major/minor Latin typeface at the document font."""
    from lxml import etree
    ns = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    for part in doc.part.package.iter_parts():
        if str(part.partname).endswith("theme1.xml"):
            tree = etree.fromstring(part.blob)
            for scheme in ("majorFont", "minorFont"):
                for font in tree.iter(ns + scheme):
                    for tf in font.iter(ns + "latin"):
                        tf.set("typeface", FONT)
            part._blob = etree.tostring(tree, xml_declaration=True,
                                        encoding="UTF-8", standalone=True)
            return


def auto_update_fields(doc):
    """Word refreshes the TOC, captions and page numbers on open.

    CT_Settings is a sequence: w:updateFields has to sit just before w:compat.
    Out of order it is silently dropped by strict readers, and every PAGEREF
    then shows its cached placeholder – a contents list of nothing but page 1.
    """
    settings = doc.settings.element
    if settings.find(qn("w:updateFields")) is not None:
        return
    flag = el("w:updateFields", val="true")
    compat = settings.find(qn("w:compat"))
    if compat is None:
        settings.append(flag)
    else:
        compat.addprevious(flag)


# --------------------------------------------------------------------------- page furniture
def setup_page(section):
    w, h, left, right, top, bottom = PAGES[PAGE]
    section.page_width, section.page_height = Cm(w), Cm(h)
    section.left_margin, section.right_margin = Cm(left), Cm(right)
    section.top_margin = Cm(top)
    section.bottom_margin = Cm(bottom)
    section.header_distance = Cm(1.25)
    section.footer_distance = Cm(1.25)
    return section


def page_numbering(section, fmt=None, start=None):
    sectPr = section._sectPr
    pg = sectPr.find(qn("w:pgNumType"))
    if pg is None:
        pg = insert_ordered(sectPr, OxmlElement("w:pgNumType"), SECTPR_ORDER)
    if fmt:
        pg.set(qn("w:fmt"), fmt)
    if start is not None:
        pg.set(qn("w:start"), str(start))


def table_border(tbl, edge, sz=6, color=LINE):
    tblPr = tbl._tbl.tblPr
    b = tblPr.find(qn("w:tblBorders"))
    if b is None:
        b = OxmlElement("w:tblBorders")
        tblPr.append(b)
    b.append(el("w:" + edge, val="single", sz=sz, space=0, color=color))


def cell_margins(tbl, top=0, left=0, bottom=0, right=0):
    mar = OxmlElement("w:tblCellMar")
    for edge, w in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        mar.append(el("w:" + edge, w=w, type="dxa"))
    tbl._tbl.tblPr.append(mar)


def trailing_spacer(part, tbl):
    """Move the part's stub paragraph after the table and shrink it to nothing."""
    p = part.paragraphs[0]
    tbl._tbl.addnext(p._p)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.0
    pPr = p._p.get_or_add_pPr()
    rPr = pPr.find(qn("w:rPr"))
    if rPr is None:
        rPr = insert_ordered(pPr, OxmlElement("w:rPr"), PPR_ORDER)
    insert_ordered(rPr, el("w:sz", val=2), RPR_ORDER)
    insert_ordered(rPr, el("w:szCs", val=2), RPR_ORDER)


def build_header(section, meta):
    """Logo left, document number and version right, rule underneath."""
    hdr = section.header
    hdr.is_linked_to_previous = False
    hdr.paragraphs[0].text = ""
    tbl = hdr.add_table(rows=1, cols=2, width=Twips(CONTENT_W))
    tbl.autofit = False
    c0, c1 = tbl.rows[0].cells
    tbl.columns[0].width = c0.width = Twips(3200)
    tbl.columns[1].width = c1.width = Twips(CONTENT_W - 3200)

    p = c0.paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
    if not place_logo(p, brand_asset("logo"), BRAND["logo_box_cm"], "header logo"):
        p.add_run(BRAND["company"]).bold = True
        set_size(p, 11, INK)

    L = meta["_labels"]
    if meta["_simple"]:
        # a bench document is filed by what it is and what it belongs to, not by
        # its revision - the revision table is three lines further down the page
        lines = [(meta["document_type"].upper(), meta["doc_no"]),
                 (L["project"], meta["project"])]
    else:
        lines = [(L["hdr_doc_no"], meta["doc_no"]),
                 (L["hdr_rev"], meta["version"])]
    lines = [(label, value) for label, value in lines if value]
    for i, (label, value) in enumerate(lines):
        pp = c1.paragraphs[0] if i == 0 else c1.add_paragraph()
        pp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        pp.paragraph_format.space_after = Pt(0)
        pp.paragraph_format.line_spacing = 1.0
        rich_text(pp, label, bold=True, size=9, color=MUTED)
        rich_text(pp, ": %s" % value, size=9, color=MUTED)

    cell_margins(tbl, bottom=140)
    table_border(tbl, "bottom")
    trailing_spacer(hdr, tbl)


def build_footer(section, meta):
    """Company left, page number right, rule above.

    Just PAGE, no "of N" – page numbering restarts between the front matter
    (i, ii, iii) and the body (1, 2, 3), so a document total would not match.
    """
    ftr = section.footer
    ftr.is_linked_to_previous = False
    tbl = ftr.add_table(rows=1, cols=2, width=Twips(CONTENT_W))
    tbl.autofit = False
    cell_margins(tbl, top=120)
    table_border(tbl, "top")
    trailing_spacer(ftr, tbl)

    left, right = tbl.rows[0].cells
    left.width = Twips(CONTENT_W - 2600)
    right.width = Twips(2600)

    L = meta["_labels"]
    p = left.paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
    p.add_run(BRAND["company"])
    set_size(p, 9)

    p = right.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.paragraph_format.space_after = Pt(0)
    p.add_run(L["page"] + " ")
    field(p, " PAGE ", "1")
    set_size(p, 9)


# --------------------------------------------------------------------------- planning pass
class Plan:
    """One walk of the section tree, before anything is rendered.

    Heading numbers, figure/table numbers and bookmark names all have to be
    known up front: the table of contents and the lists of figures/tables sit
    in the front matter, ahead of the content they describe.
    """

    def __init__(self):
        self.headings = []    # (level, number|None, text, bookmark)
        self.figures = []     # (number, caption, bookmark)
        self.tables = []
        self.listings = []    # (number, caption, bookmark)
        self.equations = []   # number


def plan_document(sections, references_title=None):
    """Figures and tables are numbered flat: Hình 1, Hình 2, Bảng 1 ...

    Chapter-scoped captions (Hình 2.5) need `STYLEREF 1 \\s` to read the nearest
    heading's list number. Our headings take their numbering from the *style*,
    not from the paragraph, so Word finds no paragraph number to read and falls
    back to the heading text, printing "Bảng Câu hỏi mở cần chốt.9". Flat SEQ
    numbering has no such dependency and comes out right in every viewer.
    """
    plan = Plan()
    state = {"fig": 0, "tbl": 0, "lst": 0, "eq": 0}

    def nxt(kind):
        state[kind] += 1
        return str(state[kind])

    def scan_blocks(blocks):
        for b in blocks:
            kind = b.get("type", "para")
            if kind == "figure" and b.get("caption"):
                plan.figures.append((nxt("fig"), b["caption"],
                                     "_Ref90%04d" % len(plan.figures)))
            elif kind == "table" and b.get("caption"):
                plan.tables.append((nxt("tbl"), b["caption"],
                                    "_Ref91%04d" % len(plan.tables)))
            elif kind == "code" and b.get("caption"):
                plan.listings.append((nxt("lst"), b["caption"],
                                      "_Ref92%04d" % len(plan.listings)))
            elif kind == "equation":
                plan.equations.append(nxt("eq"))

    def walk(secs, level, prefix, numbered_branch):
        index = 0
        for sec in secs:
            numbered = numbered_branch and sec.get("numbered", True)
            if numbered:
                index += 1
                number = ".".join(str(n) for n in prefix + (index,))
            else:
                number = None
            plan.headings.append((level, number, sec["heading"],
                                  "_Toc90%04d" % len(plan.headings)))
            scan_blocks(sec.get("blocks", []))
            walk(sec.get("sections", []), level + 1,
                 prefix + (index,) if numbered else prefix, numbered)

    walk(sections, 1, (), True)
    if references_title:   # Heading 1 Unnumbered – Word's TOC picks it up
        plan.headings.append((1, None, references_title,
                              "_Toc90%04d" % len(plan.headings)))
    return plan


# --------------------------------------------------------------------------- generated lists
def dotted_line(doc, indent, text, bookmark_name, page="1", bold=False, size=None):
    """One entry of a contents list: text, dot leader, page number, hyperlink."""
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(6 if bold else 2)
    pf.space_after = Pt(0)
    pf.left_indent = Twips(indent)
    pf.tab_stops.add_tab_stop(Twips(CONTENT_W), WD_TAB_ALIGNMENT.RIGHT,
                              WD_TAB_LEADER.DOTS)
    run = p.add_run(text)
    run.bold = bold
    run.font.size = Pt(size or BODY_PT)
    run.font.color.rgb = INK
    p.add_run()._r.append(OxmlElement("w:tab"))
    field(p, " PAGEREF %s \\h " % bookmark_name, page, dirty=True)

    link = el("w:hyperlink", anchor=bookmark_name)   # clicking jumps to the target
    p._p.append(link)
    for r in p._p.findall(qn("w:r")):
        link.append(r)
    return p


def wrap_in_field(lines, instr):
    """Put a generated list inside its Word field, so Word can rebuild it.

    Start and end ride on the first and last entry – a paragraph of their own
    would show up as a blank line.
    """
    first = lines[0]._p
    at = 1 if first.find(qn("w:pPr")) is not None else 0
    for offset, child in enumerate(field_start(instr)):
        first.insert(at + offset, child)
    end = OxmlElement("w:r")
    end.append(el("w:fldChar", fldCharType="end"))
    lines[-1]._p.append(end)


def front_heading(doc, text, page_break=False):
    """Centred banner on its own page - or, in a flat document, a plain head.

    A 16 pt centred line halfway down a bench document reads as a mistake: it
    is the furniture of a page that no longer exists.
    """
    if page_break:
        doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
    if not UPPER_H1:
        return para(doc, text.capitalize() if text.isupper() else text,
                    size=13, bold=True, before=16, after=8)
    return para(doc, text, size=16, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
                after=14)


def build_toc(doc, plan, meta, page_break=False):
    front_heading(doc, meta["_labels"]["toc_title"], page_break=page_break)
    if not plan.headings:
        return
    lines = []
    for level, number, heading, name in plan.headings:
        heading = h1_case(heading, level)
        label = ("%s %s" % (number + ".", heading)) if number and level == 1 else \
                ("%s %s" % (number, heading)) if number else heading
        lines.append(dotted_line(doc, (level - 1) * 340, label, name,
                                 bold=(level == 1),
                                 size=BODY_PT if level == 1 else BODY_PT - 1))
    wrap_in_field(lines, ' TOC \\o "1-3" \\h \\z \\u ')


def build_caption_list(doc, items, title, seq_name, page_break):
    """DANH MỤC HÌNH VẼ / DANH MỤC BẢNG BIỂU – a TOC over caption sequences."""
    front_heading(doc, title, page_break=page_break)
    lines = [dotted_line(doc, 0, "%s %s %s %s" % (seq_name, number, DASH, text), name)
             for number, text, name in items]
    wrap_in_field(lines, ' TOC \\h \\z \\c "%s" ' % seq_name)


def build_revision_history(doc, revisions, meta):
    """Under the document-control table on page i, never on the cover.

    The cover answers what the document is – a fixed job. The revision table
    answers how it changed, and grows with every issue.
    """
    L = meta["_labels"]
    front_heading(doc, L["revision_history"]).paragraph_format.space_before = Pt(22)
    grid_table(doc, [L["rev"], L["rev_date"], L["rev_desc"], L["rev_author"]],
               revisions, [1400, 1800, 4638, 1800])


def build_abbreviations(doc, rows, meta, page_break):
    L = meta["_labels"]
    front_heading(doc, L["abbr_title"], page_break=page_break)
    wide = max(len(r) for r in rows) >= 3
    header = [L["abbr_short"], L["abbr_full"]] + ([L["abbr_meaning"]] if wide else [])
    widths = [1900, 3800, 3938] if wide else [2200, 7438]
    grid_table(doc, header, rows, widths)


def build_references(doc, entries, meta, bookmark_name=None, page_break=True):
    """Tài liệu tham khảo – numbered so prose can cite [1], [2]."""
    heading = doc.add_paragraph(h1_case(meta["_labels"]["references_title"], 1),
                                style="Heading 1 Unnumbered")
    heading.paragraph_format.page_break_before = page_break
    if bookmark_name:
        bookmark(heading, bookmark_name)
    for i, entry in enumerate(entries, start=1):
        doc.add_paragraph("[%d]\t%s" % (i, entry), style="Reference")


# --------------------------------------------------------------------------- content blocks
def kv_table(doc, rows, label_w=2800):
    tbl = doc.add_table(rows=len(rows), cols=2)
    tbl.style = doc.styles["Report Table"]
    tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
    tbl.autofit = False
    for i, (k, v) in enumerate(rows):
        c0, c1 = tbl.rows[i].cells
        c0.width, c1.width = Twips(label_w), Twips(CONTENT_W - label_w)
        shade(c0._tc.get_or_add_tcPr(), BAND)
        cell_text(c0, k, bold=True)
        cell_text(c1, v)
    return tbl


def grid_table(doc, header, rows, widths=None):
    cols = len(header)
    widths = widths or [CONTENT_W // cols] * cols
    tbl = doc.add_table(rows=1 + len(rows), cols=cols)
    tbl.style = doc.styles["Report Table"]
    tbl.autofit = False
    for i, cell in enumerate(tbl.rows[0].cells):
        cell.width = Twips(widths[i])
        cell_text(cell, header[i], bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    tbl.rows[0]._tr.get_or_add_trPr().append(el("w:tblHeader", val="true"))
    for r, values in enumerate(rows, start=1):
        for i, cell in enumerate(tbl.rows[r].cells):
            cell.width = Twips(widths[i])
            cell_text(cell, values[i] if i < len(values) else "")
    return tbl


def caption(doc, seq_name, number, text, bookmark_name):
    """`Hình 3 – mô tả`, where the 3 is a SEQ field rather than typed text.

    Word recomputes it on update, so inserting a figure renumbers everything
    after it, and the list of figures picks the same SEQ up by name.
    """
    p = doc.add_paragraph(style="Caption")
    p.add_run(seq_name + " ")
    simple_field(p, " SEQ %s \\* ARABIC " % seq_name, number)
    p.add_run(" %s %s" % (DASH, text))
    bookmark(p, bookmark_name)
    return p


def code_block(doc, text):
    """A single bordered cell – keeps a listing together and off the body grid."""
    tbl = doc.add_table(rows=1, cols=1)
    tbl.autofit = False
    tbl.columns[0].width = Twips(CONTENT_W)
    cell = tbl.rows[0].cells[0]
    cell.width = Twips(CONTENT_W)
    for edge in ("top", "left", "bottom", "right"):
        table_border(tbl, edge, sz=4)
    cell_margins(tbl, top=100, left=140, bottom=100, right=140)
    shade(cell._tc.get_or_add_tcPr(), CODE_BG)
    cell.text = ""
    lines = text.split("\n")
    for i, line in enumerate(lines):
        p = cell.paragraphs[0] if i == 0 else cell.add_paragraph()
        p.style = doc.styles["Code"]
        p.add_run(line or " ")
    return tbl


def equation_block(doc, text, number=None):
    """Equation centred, number right – the usual three-column, borderless trick."""
    tbl = doc.add_table(rows=1, cols=3)
    tbl.autofit = False
    widths = [1000, 7638, 1000]
    cells = tbl.rows[0].cells
    for i, cell in enumerate(cells):
        cell.width = Twips(widths[i])
        tbl.columns[i].width = Twips(widths[i])
    cell_margins(tbl)
    cell_text(cells[1], text, size=BODY_PT, align=WD_ALIGN_PARAGRAPH.CENTER)
    if number:
        p = cells[2].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        p.add_run("(")
        simple_field(p, " SEQ CT \\* ARABIC ", number)
        p.add_run(")")
        set_size(p, BODY_PT, INK)
    return tbl


def separate_tables(doc):
    """Two <w:tbl> siblings with nothing between them are ONE table to Word.

    A report with a spec table under a data table, or two listings in a row,
    used to come out as a single grid with the columns of whichever table came
    first. Split them once at the end instead of teaching nine call sites to
    trail a spacer.
    """
    prev = None
    for child in list(doc.element.body):
        if child.tag == qn("w:tbl") and prev is not None and prev.tag == qn("w:tbl"):
            p = OxmlElement("w:p")
            pPr = OxmlElement("w:pPr")
            pPr.append(el("w:spacing", before=0, after=0, line=20, lineRule="exact"))
            rPr = OxmlElement("w:rPr")
            rPr.append(el("w:sz", val=2))
            rPr.append(el("w:szCs", val=2))
            pPr.append(rPr)
            p.append(pPr)
            child.addprevious(p)
        prev = child


def render_blocks(doc, blocks, ctx):
    L = ctx["labels"]
    for block in blocks:
        kind = block.get("type", "para")

        if kind == "para":
            rich_text(doc.add_paragraph(style="Body Text"), block["text"])

        elif kind == "quote":
            rich_text(doc.add_paragraph(style="Quote"), block["text"])

        elif kind == "bullets":
            for item in block["items"]:
                if isinstance(item, (list, tuple)):      # nested level
                    for sub in item:
                        rich_text(doc.add_paragraph(style="List Bullet 2"), sub)
                else:
                    rich_text(doc.add_paragraph(style="List Bullet"), item)

        elif kind == "kv":
            kv_table(doc, [tuple(r) for r in block["rows"]],
                     label_w=block.get("label_width", 2800))

        elif kind == "table":
            # Table captions go ABOVE the table, figure captions below it.
            if block.get("caption"):
                number, text, name = next(ctx["tables"])
                caption(doc, L["table"], number, text, name) \
                    .paragraph_format.keep_with_next = True
            grid_table(doc, block["header"], block["rows"], block.get("widths"))

        elif kind == "figure":
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(8)
            path = block.get("path")
            if path and os.path.exists(path):
                p.add_run().add_picture(path, width=Cm(block.get("width_cm", 14)))
            else:
                if path:   # asked for an image and did not get one – say so
                    sys.stderr.write("warning: figure not found, placeholder used"
                                     " instead: %s\n" % os.path.abspath(path))
                p.add_run(block.get("placeholder", L["figure_slot"])).font.color.rgb = MUTED
            if block.get("caption"):
                number, text, name = next(ctx["figures"])
                caption(doc, L["figure"], number, text, name)

        elif kind == "code":
            # Listing captions go ABOVE the listing, like table captions.
            if block.get("caption"):
                number, text, name = next(ctx["listings"])
                if block.get("language"):
                    text = "%s (%s)" % (text, block["language"])
                caption(doc, L["listing"], number, text, name).paragraph_format \
                    .keep_with_next = True
            code_block(doc, block["text"])

        elif kind == "equation":
            equation_block(doc, block["text"], next(ctx["equations"], None))

        elif kind == "pagebreak":
            doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

        else:
            raise ValueError("unknown block type: %r" % kind)


def h1_case(text, level):
    """Top-level headings print upper case, except in the simple layout.

    As text, not as the w:caps effect: Google Docs has no all-caps run format
    and turns w:caps into small caps on import. Word's rebuilt contents list
    copies the heading text, so build_toc has to shout the same entries back -
    which is why this is one function and not an upper() at each call site.

    A flat document is read straight through, and shouting every section title
    at a reader who never left the page is noise.
    """
    return text.upper() if level == 1 and UPPER_H1 else text


def render_sections(doc, sections, ctx, level=1, numbered_branch=True):
    for section in sections:
        numbered = numbered_branch and section.get("numbered", True)
        style = "Heading %d" % min(level, 3)
        if not numbered:
            style += " Unnumbered"
        p = doc.add_paragraph(h1_case(section["heading"], level), style=style)
        if level == 1:
            # Every top-level section opens a page. Not the first one: it already
            # sits at the top of the body section, and a break there would leave
            # a blank page behind.
            if ctx["opened_body"] and ctx["page_per_section"]:
                p.paragraph_format.page_break_before = True
            ctx["opened_body"] = True
        bookmark(p, next(ctx["headings"]))
        render_blocks(doc, section.get("blocks", []), ctx)
        render_sections(doc, section.get("sections", []), ctx, level + 1, numbered)


# --------------------------------------------------------------------------- document parts
def letterspace(paragraph, twentieths=60):
    """Track out a short label so it reads as a rubric, not a sentence."""
    for run in paragraph.runs:
        insert_ordered(run._r.get_or_add_rPr(), el("w:spacing", val=twentieths), RPR_ORDER)


def month_year(date):
    """17/08/2026 -> 08/2026. Anything else is passed through untouched."""
    parts = date.split("/")
    return "%s/%s" % (parts[1], parts[2]) if len(parts) == 3 else date


def cover_cell(cell, valign=None):
    """A cover cell starts empty – python-docx hands out a stub paragraph."""
    if valign is not None:
        cell.vertical_alignment = valign
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
    return p


def build_cover(doc, meta):
    """Identity row, title band, author and logo along the bottom.

    Every element lives in a cell of one borderless table: a cover made of
    floating boxes falls apart the first time somebody edits it in Word, and
    this template is meant to be edited. Only the grey ground is a picture,
    anchored to the page and locked so it cannot be dragged.

    The eight identity fields moved to page i – a cover carrying a table of
    eight rows reads as a form, not as a cover.
    """
    L = meta["_labels"]

    art = doc.add_paragraph()          # host for the page-anchored ground
    art.paragraph_format.space_after = Pt(0)
    art.paragraph_format.line_spacing = Pt(1)
    fd, bg = tempfile.mkstemp(suffix=".png")
    os.close(fd)
    try:
        float_picture(art, solid_png(bg, BRAND["cover_bg"]), Cm(21), 0, 0)
    finally:
        os.remove(bg)                  # add_picture already copied it into the part

    tbl = doc.add_table(rows=3, cols=2)
    tbl.autofit = False
    cell_margins(tbl)
    # One column split for every row – Word keeps a single grid per table, and
    # rows that disagree about it push the table off the page.
    for row, height in zip(tbl.rows, (Cm(1.8), Cm(17.8), Cm(3.2))):
        row.height = height
        row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
        row.cells[0].width = Twips(CONTENT_W - 5600)
        row.cells[1].width = Twips(5600)

    p = cover_cell(tbl.cell(0, 0))
    name = p.add_run(BRAND["company"])
    name.bold = True
    name.font.size = Pt(20)
    name.font.color.rgb = INK
    insert_ordered(name._r.get_or_add_rPr(), el("w:spacing", val=40), RPR_ORDER)
    date = tbl.cell(0, 0).add_paragraph()
    date.paragraph_format.space_after = Pt(0)
    date.add_run(month_year(meta["date"])).bold = True
    set_size(date, 11)

    p = cover_cell(tbl.cell(0, 1))
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    chip = p.add_run("  %s %s  " % (L["version"], meta["version"]))
    chip.bold = True
    chip.font.size = Pt(10)
    chip.font.color.rgb = INK
    run_shade(chip, CHIP)

    band = tbl.cell(1, 0).merge(tbl.cell(1, 1))
    rubric = cover_cell(band)
    rubric.paragraph_format.space_before = Pt(54)
    rubric.paragraph_format.space_after = Pt(8)
    rubric.add_run(meta["document_type"]).bold = True
    set_size(rubric, 10)
    letterspace(rubric, 30)

    title = band.add_paragraph(meta["title"], style="Title")
    title.paragraph_format.space_after = Pt(12)
    title.paragraph_format.right_indent = Cm(4.5)   # keep the title off the edge

    rule = band.add_paragraph()
    rule.paragraph_format.space_after = Pt(0)
    rule.paragraph_format.line_spacing = Pt(1)
    rule.paragraph_format.right_indent = Cm(11)
    borders(rule._p.get_or_add_pPr(), ["bottom"], sz=24, color=RULE)

    author = tbl.cell(2, 0)
    p = cover_cell(author, WD_ALIGN_VERTICAL.BOTTOM)
    p.add_run(L["author"]).bold = True
    set_size(p, 10, INK)
    for text in (meta["prepared_by"], meta["email"]):
        if text:
            line = author.add_paragraph(text)
            line.paragraph_format.space_after = Pt(0)
            set_size(line, 10)

    p = cover_cell(tbl.cell(2, 1), WD_ALIGN_VERTICAL.BOTTOM)
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    # the mark alone: the cover already carries the company name in the top left
    place_logo(p, brand_asset("mark"), BRAND["mark_box_cm"], "cover mark")


def build_simple_head(doc, meta, revisions):
    """Title, who wrote it, who ran it, revisions - and straight into section 1.

    A test record is read at a bench, not filed: no cover, no control page, no
    roman front matter. The styles, the branding and the numbering are the same
    ones the formal layout uses; only the furniture around them is gone.
    """
    L = meta["_labels"]
    # Heading 1 weight, not the cover-sized Title style: the flat document opens
    # on a line, not on a title page.
    title = para(doc, "", before=0, after=10)
    rich_text(title, "%s: %s" % (L["title_prefix"], meta["title"]),
              bold=True, size=16)
    for label, value in ((L["prepared_by"], meta["prepared_by"]),
                         (L["tested_by"], meta["tested_by"]),
                         (L["date"], meta["date"]),
                         (L["doc_no"], meta["doc_no"])):
        if value:
            p = para(doc, after=0)
            rich_text(p, "%s: " % label, bold=True, size=11)
            rich_text(p, value, size=11)
    if revisions:
        para(doc, L["revisions"], size=11, bold=True, before=12, after=6)
        grid_table(doc, [L["rev_date"], L["rev_author"], L["rev_desc"], L["rev"]],
                   [[r[1], r[3], r[2], r[0]] for r in revisions],
                   [1800, 1800, 4238, 1800])


def build_doc_control(doc, meta):
    """The identity fields the cover no longer carries – first thing on page 2."""
    L = meta["_labels"]
    front_heading(doc, L["doc_control"])
    kv_table(doc, [
        (L["object"], meta["object"]),
        (L["doc_no"], meta["doc_no"]),
        (L["version"], meta["version"]),
        (L["date"], meta["date"]),
        (L["prepared_by"], meta["prepared_by"]),
    ] + ([(L["tested_by"], meta["tested_by"])] if meta["tested_by"] else []) + [
        (L["reviewed_by"], meta["reviewed_by"]),
        (L["approved_by"], meta["approved_by"]),
        (L["classification"], meta["classification"]),
    ], label_w=3000)



def build_instructions(doc):
    """The red page the template carries and a real report never does."""
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
    p = para(doc, "HOW TO USE THIS TEMPLATE %s DELETE THIS PAGE BEFORE SENDING" % DASH,
             size=14, bold=True, color=WARN, after=10)
    borders(p._p.get_or_add_pPr(), ["bottom"], sz=6, color="C00000")

    for rule in [
        "Replace every {{PLACEHOLDER}} with real content. Nothing of the form "
        "{{...}} may survive into the copy you send.",
        "Do NOT type section numbers. Heading 1 / 2 / 3 number themselves "
        "(1. / 1.1 / 1.1.1) and renumber when a section is added or removed.",
        "A section that carries no number (Conclusion, Appendix, References) uses "
        "the Heading 1 Unnumbered style %s it still appears in the table of "
        "contents, without a number." % DASH,
        "Do NOT type figure or table numbers either. Captions are SEQ fields, so "
        "they number themselves (Figure 1, Table 2) and renumber on insert.",
        "TABLE and LISTING captions sit ABOVE; FIGURE captions sit BELOW the figure.",
        "A command, flag, path or file name inside a sentence goes between "
        "backticks to come out in Consolas.",
        "Every top-level section (Heading 1) starts on a new page.",
        "Use only the styles that ship with the template: Title, Heading 1-3 (and "
        "their Unnumbered variants), Body Text, List Bullet, List Bullet 2, "
        "Caption, Code, Reference, Report Table. Never hand-format (bold / size / "
        "colour) in place of a style.",
        "Prose must cite figures and tables by number (\"see Figure 2\"), never "
        "\"the figure below\".",
        "Use the en dash %s never the longer em dash." % DASH,
        "The contents list and the lists of figures and tables are generated with "
        "page numbers already in them; Word refreshes them on open, or select all "
        "and press F9.",
        "The cover carries no page number; the front matter (document information, "
        "revision history, contents, lists) is numbered i, ii, iii; the body "
        "restarts at 1.",
        "Branding %s company name, logo, cover colour %s lives in "
        "assets/brand.json, not in this document." % (DASH, DASH),
    ]:
        doc.add_paragraph(rule, style="List Bullet")

    para(doc, "Placeholders", size=11, bold=True, before=14, after=6)
    grid_table(doc, ["Placeholder", "Meaning", "Example"], [
        ["{{DOCUMENT_TYPE}}", "Kind of document", "Report"],
        ["{{TITLE}}", "Report title", "Payload integration study"],
        ["{{SUBJECT}}", "What the report is about", "Phase One P3"],
        ["{{DOC_NO}}", "Internal document number", BRAND["doc_no_example"]],
        ["{{VERSION}}", "Version", "1.0"],
        ["{{DATE}}", "Issue date", "17/08/2026"],
        ["{{PREPARED_BY}}", "Author", "Name %s Role" % DASH],
        ["{{AUTHOR_EMAIL}}", "Author email, shown on the cover",
         BRAND["email_example"]],
        ["{{CLASSIFICATION}}", "Classification", "Confidential"],
    ], [3000, 3838, 2800])


# --------------------------------------------------------------------------- build
META_DEFAULTS = {
    "document_type": "",          # falls back to the language's default rubric
    "title": "",           # falls back to the language's default rubric
    "object": "",
    "project": "",             # the programme the document belongs to
    "doc_no": "",
    "version": "1.0",
    "date": "",
    "prepared_by": "",
    "tested_by": "",           # a test report names who ran it, not only who wrote it
    "reviewed_by": "",
    "approved_by": "",
    "email": "",
    "classification": "",         # falls back to the language's default
    "revisions": [],
}


# Every key the schema knows. A typo used to be dropped without a word and
# print as an empty field halfway down page i.
SPEC_KEYS = set(META_DEFAULTS) | {"lang", "font", "body_pt", "page", "brand",
                                  "layout", "toc", "list_of_figures",
                                  "abbreviations_list",
                                  "list_of_tables", "abbreviations", "references",
                                  "sections"}
SECTION_KEYS = {"heading", "numbered", "blocks", "sections"}
# type -> (required fields, optional fields)
BLOCK_SPEC = {
    "para": (["text"], []),
    "quote": (["text"], []),
    "bullets": (["items"], []),
    "kv": (["rows"], ["label_width"]),
    "table": (["header", "rows"], ["widths", "caption"]),
    "figure": ([], ["path", "width_cm", "caption", "placeholder"]),
    "code": (["text"], ["language", "caption"]),
    "equation": (["text"], []),
    "pagebreak": ([], []),
}


def validate(spec):
    """Reject the JSON here, with a line the writer can act on.

    Everything below this point assumes the spec is shaped right, and fails
    deep in the renderer with a KeyError naming an XML helper if it is not.
    """
    def bad(msg, *args):
        raise ValueError(msg % args if args else msg)

    def unknown(got, allowed, what):
        extra = set(got) - allowed
        if extra:
            bad("%s: unknown key(s) %s – known: %s", what,
                ", ".join(sorted(extra)), ", ".join(sorted(allowed)))

    unknown(spec, SPEC_KEYS, "content.json")
    unknown(spec.get("brand") or {}, set(BRAND_DEFAULTS), "brand")
    if (spec.get("lang") or BRAND["lang"]) not in LABELS:
        bad("lang must be one of: %s", ", ".join(sorted(LABELS)))
    if spec.get("layout", "formal") not in ("formal", "simple"):
        bad("layout must be 'formal' or 'simple', got %r", spec["layout"])
    if spec.get("page", "a4") not in PAGES:
        bad("page must be one of: %s", ", ".join(sorted(PAGES)))
    if not 8 <= spec.get("body_pt", BODY_PT) <= 14:
        bad("body_pt must be between 8 and 14, got %r", spec["body_pt"])
    pick_font(spec.get("font") or BRAND["font"])   # raises on a font Word may lack

    for row in spec.get("revisions") or []:
        if len(row) != 4:
            bad("revisions row needs 4 cells "
                "(version, date, description, author): %r", row)
    for row in spec.get("abbreviations") or []:
        if len(row) not in (2, 3):
            bad("abbreviations row needs 2 or 3 cells: %r", row)

    def check_blocks(blocks, where):
        for i, block in enumerate(blocks, start=1):
            kind = block.get("type", "para")
            at = "%s, block %d (%s)" % (where, i, kind)
            if kind not in BLOCK_SPEC:
                bad("%s: unknown type – one of: %s", at,
                    ", ".join(sorted(BLOCK_SPEC)))
            required, optional = BLOCK_SPEC[kind]
            unknown(block, set(required) | set(optional) | {"type"}, at)
            for f in required:
                if not block.get(f):
                    bad("%s: missing %r", at, f)
            if kind == "kv":
                for n, row in enumerate(block["rows"], start=1):
                    if len(row) != 2:
                        bad("%s: row %d needs 2 cells (label, value): %r", at, n, row)
            if kind == "table":
                cols = len(block["header"])
                for n, row in enumerate(block["rows"], start=1):
                    if len(row) > cols:
                        bad("%s: row %d has %d cells, the header has %d",
                            at, n, len(row), cols)
                widths = block.get("widths")
                if widths and len(widths) != cols:
                    bad("%s: %d widths for %d columns", at, len(widths), cols)
                if widths and sum(widths) != CONTENT_W:
                    bad("%s: widths sum to %d, the text column is %d twips wide",
                        at, sum(widths), CONTENT_W)

    def walk(sections, path):
        for i, section in enumerate(sections, start=1):
            where = "section %s%d" % (path, i)
            unknown(section, SECTION_KEYS, where)
            if not section.get("heading"):
                bad("%s: missing 'heading'", where)
            check_blocks(section.get("blocks", []), where)
            walk(section.get("sections", []), "%s%d." % (path, i))

    walk(spec.get("sections", []), "")
    return spec


def build(spec, out_path, instructions=False):
    # ponytail: module globals, not parameters threaded through forty calls.
    # One document per process in every real use; selfcheck sets them per build.
    global FONT, BRAND, INK, UPPER_H1, PAGE, CONTENT_W, BODY_PT
    BRAND = load_brand(spec)
    PAGE = spec.get("page", "a4")
    CONTENT_W = content_width(PAGE)
    BODY_PT = spec.get("body_pt", DEFAULT_BODY_PT)
    # A simple document is one section that starts at page 1: no cover, no
    # control page, no roman front matter, and generated lists off unless asked.
    simple = spec.get("layout", "formal") == "simple"
    lists_on = not simple
    UPPER_H1 = not simple
    validate(spec)
    labels = LABELS[spec.get("lang") or BRAND["lang"]]
    meta = dict(META_DEFAULTS)
    meta.update({k: v for k, v in spec.items() if k in META_DEFAULTS})
    meta["_labels"] = labels
    meta["_simple"] = simple
    meta["title"] = meta["title"] or labels["document_type"]
    meta["document_type"] = meta["document_type"] or labels["document_type"]
    meta["classification"] = meta["classification"] or labels["classification_default"]

    FONT = pick_font(spec.get("font") or BRAND["font"])
    INK = RGBColor.from_string(BRAND["ink"].lstrip("#").upper())

    doc = Document()
    doc.element.body.clear_content()
    install_numbering(doc)
    install_styles(doc, labels["lang_tag"])
    install_table_style(doc)
    set_theme_fonts(doc)
    auto_update_fields(doc)

    plan = plan_document(spec.get("sections", []),
                         labels["references_title"] if spec.get("references") else None)

    abbreviations = spec.get("abbreviations") or []
    revisions = meta.get("revisions") or []

    if simple:
        body = setup_page(doc.sections[0])
        build_header(body, meta)
        build_footer(body, meta)
        build_simple_head(doc, meta, revisions)
    else:
        # Section 1 – cover. Its header/footer are defined here and inherited by
        # the later sections; the cover itself is exempt via a blank first page.
        cover = setup_page(doc.sections[0])
        cover.different_first_page_header_footer = True
        build_header(cover, meta)
        build_footer(cover, meta)
        build_cover(doc, meta)

        front = setup_page(doc.add_section(WD_SECTION.NEW_PAGE))
        front.different_first_page_header_footer = False
        page_numbering(front, fmt="lowerRoman", start=1)

        # document control and revision history share page i – both answer "which
        # document is this", and neither fills a page on its own
        build_doc_control(doc, meta)
        if revisions:
            build_revision_history(doc, revisions, meta)

    if spec.get("toc", lists_on):
        build_toc(doc, plan, meta, page_break=not simple)
    if plan.figures and spec.get("list_of_figures", lists_on):
        build_caption_list(doc, plan.figures, labels["lof_title"],
                           labels["figure"], page_break=not simple)
    if plan.tables and spec.get("list_of_tables", lists_on):
        build_caption_list(doc, plan.tables, labels["lot_title"],
                           labels["table"], page_break=not simple)
    if abbreviations and spec.get("abbreviations_list", lists_on):
        build_abbreviations(doc, abbreviations, meta, page_break=not simple)

    if not simple:
        body = setup_page(doc.add_section(WD_SECTION.NEW_PAGE))
        body.different_first_page_header_footer = False
        page_numbering(body, fmt="decimal", start=1)

    ctx = {
        "labels": labels,
        "opened_body": False,
        "page_per_section": not simple,
        "headings": iter([h[3] for h in plan.headings]),
        "figures": iter(plan.figures),
        "tables": iter(plan.tables),
        "listings": iter(plan.listings),
        "equations": iter(plan.equations),
    }
    render_sections(doc, spec.get("sections", []), ctx)

    if spec.get("references"):
        build_references(doc, spec["references"], meta, next(ctx["headings"], None),
                         page_break=not simple)
    if instructions:
        build_instructions(doc)

    separate_tables(doc)

    core = doc.core_properties
    core.title = meta["title"]
    core.author = meta["prepared_by"] or BRAND["company"]
    core.category = "%s report" % BRAND["company"]

    doc.save(out_path)
    return out_path


# --------------------------------------------------------------------------- placeholder spec
def placeholder_spec():
    """The blank template is just a report whose content is placeholders.

    The scaffold is deliberately domain-neutral: whoever fills it in renames
    the sections. The furniture follows the brand language, the placeholder
    tokens stay English because they are variable names, not prose.
    """
    return {
        "lang": BRAND["lang"],
        "document_type": "{{DOCUMENT_TYPE}}",
        "title": "{{TITLE}}",
        "object": "{{SUBJECT}}",
        "doc_no": "{{DOC_NO}}",
        "version": "{{VERSION}}",
        "date": "{{DATE}}",
        "prepared_by": "{{PREPARED_BY}}",
        "reviewed_by": "{{REVIEWED_BY}}",
        "approved_by": "{{APPROVED_BY}}",
        "email": "{{AUTHOR_EMAIL}}",
        "classification": "{{CLASSIFICATION}}",
        "revisions": [["{{VERSION}}", "{{DATE}}", "{{CHANGE_DESCRIPTION}}",
                       "{{CHANGE_AUTHOR}}"]],
        "abbreviations": [["{{ABBR}}", "{{FULL_TERM}}", "{{MEANING}}"]],
        "references": ["{{AUTHOR}}, {{DOCUMENT_TITLE}}, {{SOURCE}}, {{YEAR}}"],
        "sections": [
            {"heading": "Overview", "blocks": [
                {"type": "para", "text": "{{OVERVIEW %s 3-5 sentences: why this "
                                         "document exists, what it covers, and the "
                                         "conclusion in one line.}}" % DASH}]},
            {"heading": "Background", "blocks": [
                {"type": "bullets", "items": ["{{BACKGROUND_1}}", "{{BACKGROUND_2}}",
                                              "{{BACKGROUND_3}}"]},
                {"type": "figure", "placeholder": "[ {{INSERT_FIGURE}} ]",
                 "caption": "{{FIGURE_CAPTION}}"}]},
            {"heading": "Specification", "blocks": [
                {"type": "para", "text": "{{SPECIFICATION_INTRO}}"},
                {"type": "kv", "label_width": 3200, "rows": [
                    ["{{PARAMETER_1}}", "{{VALUE_1}}"],
                    ["{{PARAMETER_2}}", "{{VALUE_2}}"],
                    ["{{PARAMETER_3}}", "{{VALUE_3}}"],
                    ["{{PARAMETER_4}}", "{{VALUE_4}}"]]}]},
            {"heading": "Method", "sections": [
                {"heading": "{{STEP_NAME_1}}", "blocks": [
                    {"type": "bullets", "items": ["{{STEP_DETAIL_1}}",
                                                  "{{SUPPORTING_LINK}}"]},
                    {"type": "code", "language": "bash",
                     "caption": "{{LISTING_CAPTION}}",
                     "text": "{{COMMAND_OR_CONFIGURATION}}"}]},
                {"heading": "{{STEP_NAME_2}}", "blocks": [
                    {"type": "bullets", "items": ["{{STEP_DETAIL_2}}"]}]},
                {"heading": "{{STEP_NAME_3}}", "blocks": [
                    {"type": "bullets", "items": ["{{STEP_DETAIL_3}}"]}]}]},
            {"heading": "Results", "blocks": [
                {"type": "table",
                 "caption": "{{TABLE_CAPTION}}",
                 "header": ["{{COL_1}}", "{{COL_2}}", "{{COL_3}}", "{{COL_4}}",
                            "{{COL_5}}", "{{COL_6}}"],
                 "widths": [900, 2000, 1700, 2038, 1400, 1600],
                 "rows": [["{{CELL}}"] * 6] + [[""] * 6] * 4}]},
            {"heading": "Conclusion", "numbered": False, "blocks": [
                {"type": "para", "text": "{{CONCLUSION %s what is confirmed, what "
                                         "the reader has to decide, and the next "
                                         "step you recommend.}}" % DASH}]},
        ],
    }


def main(argv):
    if len(argv) < 2 or argv[1] not in ("template", "write"):
        print(__doc__)
        return 1

    global BRAND
    BRAND = load_brand()

    if argv[1] == "template":
        out = argv[2] if len(argv) > 2 else "Report_Template.docx"
        print(build(placeholder_spec(), out, instructions=True))
        return 0

    if len(argv) < 3:
        print("usage: report.py write content.json [out.docx]")
        return 1
    with open(argv[2], encoding="utf-8") as fh:
        spec = json.load(fh)
    out = argv[3] if len(argv) > 3 else (spec.get("doc_no") or "report") + ".docx"
    try:
        print(build(spec, out))
    except ValueError as exc:   # a spec the writer can fix – no traceback needed
        sys.stderr.write("error: %s\n" % exc)
        return 1
    return 0


def _selfcheck():
    """Smallest runnable check: both modes produce a valid, styled docx."""
    import contextlib
    import io as _io
    import re
    import zipfile
    from docx import Document as _D

    spec = {
        "lang": "vi",
        "title": "Check", "doc_no": "X-1", "date": "today",
        "revisions": [["1.0", "today", "first", "Q"]],
        "abbreviations": [["GCS", "Ground Control Station", "Trạm điều khiển"]],
        "references": ["Tác giả A, Tài liệu B, 2026"],
        "sections": [
            {"heading": "One", "blocks": [
                {"type": "para", "text": "body with `--inline --code` in it"},
                {"type": "bullets", "items": ["a", ["nested"]]},
                {"type": "kv", "rows": [["k", "v"]]},
                {"type": "table", "header": ["A", "B"], "rows": [["1", "2"]],
                 "caption": "cap"},
                {"type": "figure", "caption": "fig"},
                {"type": "code", "text": "line1\nline2", "language": "bash",
                 "caption": "listing cap"},
                {"type": "equation", "text": "E = mc^2"}],
             "sections": [{"heading": "Two", "blocks": [{"type": "para", "text": "x"}]}]},
            {"heading": "Two chapter", "blocks": [{"type": "figure", "caption": "f2"}]},
            {"heading": "Kết luận", "numbered": False,
             "blocks": [{"type": "para", "text": "done"}]},
        ],
    }

    def adjacent_tables(path):
        body = _D(path).element.body
        tags = [c.tag for c in body]
        return sum(1 for a, b in zip(tags, tags[1:])
                   if a == b == qn("w:tbl"))

    with tempfile.TemporaryDirectory() as tmp:
        t = build(placeholder_spec(), os.path.join(tmp, "t.docx"), instructions=True)
        r = build(spec, os.path.join(tmp, "r.docx"))

        # the shipped example is documentation, and documentation drifts
        with open(EXAMPLE, encoding="utf-8") as fh:
            build(json.load(fh), os.path.join(tmp, "x.docx"))

        for path in (t, r):
            names = {s.name for s in _D(path).styles}
            for want in ("Title", "Heading 1", "Heading 2", "Heading 1 Unnumbered",
                         "Body Text", "List Bullet", "Caption", "Code", "Reference",
                         "Report Table"):
                assert want in names, (path, want)

        d = _D(r)
        texts = [p.text for p in d.paragraphs]
        styles = [p.style.name for p in d.paragraphs]
        # level 1 shouts as text, so Word's rebuilt contents list shouts too
        assert "ONE" in texts and "Two" in texts, texts
        assert "1. ONE\t1" in texts and "1.1 Two\t1" in texts, texts
        assert "Heading 1 Unnumbered" in styles, "unnumbered heading not used"
        # document control and revision history sit on page i, not on the cover
        rev_at = texts.index("LỊCH SỬ PHIÊN BẢN")
        assert rev_at < texts.index("MỤC LỤC"), "revision history should precede the TOC"
        assert texts.index("THÔNG TIN TÀI LIỆU") < rev_at, "control table should lead"
        cover = d.tables[0]          # the cover table comes before the control table
        assert cover.cell(0, 0).text.startswith(BRAND["company"]), cover.cell(0, 0).text
        assert "Check" in cover.cell(1, 0).text, "title should sit in the cover table"
        assert "Tác giả" in cover.cell(2, 0).text, "author block missing from the cover"
        assert cover.cell(1, 0).text.startswith("Báo cáo"), cover.cell(1, 0).text
        assert month_year("17/08/2026") == "08/2026", "cover date should read month/year"
        assert month_year("{{NGAY}}") == "{{NGAY}}", "unparsable date must pass through"

        # captions carry SEQ fields, not typed numbers
        caps = [t for t in texts if t.startswith(("Hình", "Bảng"))]
        assert any(c.startswith("Hình") for c in caps), caps
        assert any(c.startswith("Bảng") for c in caps), caps
        assert all(DASH in c for c in caps), caps

        with zipfile.ZipFile(r) as z:
            xml = z.read("word/document.xml").decode()
            styles_xml = z.read("word/styles.xml").decode()
            assert FONT in styles_xml, "font not applied"
            assert "vi-VN" in styles_xml, "language not set"
            assert 'SEQ Hình' in xml and 'SEQ Bảng' in xml, "caption SEQ fields missing"
            assert 'SEQ Đoạn mã' in xml, "listing caption SEQ field missing"
            # inline `code` must reach the run, not stay as literal backticks
            assert '<w:t>--inline --code</w:t>' in xml, "inline code run missing"
            assert "`" not in xml, "a backtick survived into the document"
            # the cover is one borderless table over a single anchored ground
            assert xml.count("<wp:anchor") == 1, "cover ground not anchored to the page"
            assert "w:framePr" not in xml, "cover should not float anything but the ground"
            # STYLEREF cannot read style-level heading numbers; it must stay out
            assert "STYLEREF" not in xml, "STYLEREF would print the heading text"
            # every top-level section opens a page, except the first in the body
            breaks = xml.count("<w:pageBreakBefore/>")
            assert breaks == 3, "expected 3 page breaks, got %d" % breaks
            assert 'TOC \\h \\z \\c "Hình"' in xml, "list of figures missing"
            assert 'TOC \\h \\z \\c "Bảng"' in xml, "list of tables missing"
            assert xml.count("lowerRoman") == 1, "front matter numbering missing"
            # out of schema order this flag is dropped and every TOC page reads 1
            settings = z.read("word/settings.xml").decode()
            assert 0 < settings.index("w:updateFields") < settings.index("w:compat"), \
                "updateFields must sit just before w:compat"
            assert "\u2014" not in xml, "em dash leaked into the document"
            # bullet glyphs are private-use chars that editors like to eat
            num = z.read("word/numbering.xml").decode()
            assert '<w:lvlText w:val=""/>' in num, "level-1 bullet glyph lost"
            assert '<w:lvlText w:val=""/>' in num, "level-3 bullet glyph lost"

        # two <w:tbl> in a row are one table to Word – nothing may sit adjacent
        for path in (t, r, os.path.join(tmp, "x.docx")):
            assert adjacent_tables(path) == 0, ("tables merge in", path)
        stacked = build(dict(spec, sections=[{"heading": "S", "blocks": [
            {"type": "kv", "rows": [["k", "v"]]},
            {"type": "kv", "rows": [["k", "v"]]},
            {"type": "code", "text": "a"},
            {"type": "equation", "text": "e"}]}]),
            os.path.join(tmp, "s.docx"))
        assert adjacent_tables(stacked) == 0, "back-to-back tables were left merged"

        # the written-out contents must match what Word rebuilds: references
        # carry an outline level, so they belong in both
        toc = texts[texts.index("MỤC LỤC") + 1:]
        toc = toc[:toc.index("")]
        assert any(t2.startswith("TÀI LIỆU THAM KHẢO") for t2 in toc), toc

        # the spec is checked before anything is rendered
        for broken, why in (
                ({"author": "x"}, "unknown top-level key"),
                ({"lang": "vn"}, "unknown language"),
                ({"sections": [{"blocks": []}]}, "section without a heading"),
                ({"sections": [{"heading": "H", "blocks": [{"type": "para"}]}]},
                 "para without text"),
                ({"sections": [{"heading": "H", "blocks": [
                    {"type": "table", "header": ["A"], "rows": [["1"]],
                     "widths": [100]}]}]}, "widths that do not fill the column"),
                ({"sections": [{"heading": "H", "blocks": [
                    {"type": "figure", "captions": "typo"}]}]}, "misspelt block key")):
            try:
                validate(dict(spec, **broken))
                raise AssertionError("validate accepted a spec with a " + why)
            except ValueError:
                pass

        # flat caption numbering, continuous across sections
        plan = plan_document(spec["sections"])
        assert [n for n, _, _ in plan.figures] == ["1", "2"], plan.figures
        assert [n for n, _, _ in plan.tables] == ["1"], plan.tables
        assert [n for n, _, _ in plan.listings] == ["1"], plan.listings
        assert plan.headings[-1][1] is None, "Kết luận should carry no number"

        # a different font reaches the styles, the theme and the table style
        f = build(dict(spec, font="calibri"), os.path.join(tmp, "f.docx"))
        with zipfile.ZipFile(f) as z:
            assert "Calibri" in z.read("word/styles.xml").decode(), "font not applied"
            assert "Calibri" in z.read("word/theme/theme1.xml").decode(), "theme not set"
        try:
            build(dict(spec, font="Inter"), os.path.join(tmp, "g.docx"))
            raise AssertionError("a font Word cannot be trusted to have was accepted")
        except ValueError:
            pass

        # no "lang" in content.json falls back to the brand's - shipped as en
        nolang = dict(spec)
        del nolang["lang"]
        e = _D(build(nolang, os.path.join(tmp, "e.docx")))
        etexts = [p.text for p in e.paragraphs]
        assert "TABLE OF CONTENTS" in etexts, etexts
        assert any(t.startswith("Figure") for t in etexts), etexts

        # the simple layout drops the furniture and keeps the styles
        flat = build(dict(spec, layout="simple"), os.path.join(tmp, "flat.docx"))
        fd = _D(flat)
        ftexts = [p.text for p in fd.paragraphs]
        assert len(fd.sections) == 1, "simple layout should be a single section"
        assert ftexts[0] == "TIÊU ĐỀ: Check", ftexts[:3]   # title line, no cover
        assert "THÔNG TIN TÀI LIỆU" not in ftexts, "control page survived"
        assert "MỤC LỤC" not in ftexts, "contents should default off when flat"
        assert "One" in ftexts and "ONE" not in ftexts, "flat headings should not shout"
        # the abbreviations banner used to land in the middle of a flat document
        assert not any(t.isupper() and len(t) > 12 for t in ftexts), \
            [t for t in ftexts if t.isupper() and len(t) > 12]
        assert fd.tables[0].cell(0, 0).text == "Ngày", fd.tables[0].cell(0, 0).text
        # the header names the document and its programme, not its revision
        fhdr = fd.sections[0].header.tables[0].rows[0].cells[1].text
        assert "BÁO CÁO: X-1" in fhdr, fhdr
        assert "Phiên bản" not in fhdr, fhdr
        with zipfile.ZipFile(flat) as z:
            fxml = z.read("word/document.xml").decode()
            fstyles = z.read("word/styles.xml").decode()
            assert "<w:pageBreakBefore/>" not in fxml, "flat sections must not break"
            assert "<wp:anchor" not in fxml, "cover ground survived into a flat document"
            assert "lowerRoman" not in fxml, "flat documents number from 1"
            assert "SEQ Bảng" in fxml, "captions still number themselves when flat"
            h1 = re.search(r'w:styleId="Heading1".*?</w:style>', fstyles, re.S).group(0)
            assert "w:pBdr" not in h1, "the heading rule belongs to paged sections"
        # the programme reaches the header when it is given
        pj = _D(build(dict(spec, layout="simple", project="ALTA X"),
                      os.path.join(tmp, "pj.docx")))
        assert "Dự án: ALTA X" in pj.sections[0].header.tables[0].rows[0].cells[1].text

        # Letter narrows the text column, and the widths must follow the page
        lt = build(dict(spec, page="letter", sections=[{"heading": "S", "blocks": [
            {"type": "table", "header": ["A", "B"], "rows": [["1", "2"]],
             "widths": [4405, 5000]}]}]), os.path.join(tmp, "lt.docx"))
        assert abs(_D(lt).sections[0].page_width - Cm(21.59)) < 2000, "not Letter"
        assert content_width("letter") == 9405, content_width("letter")
        assert content_width("a4") == 9638, content_width("a4")
        try:
            build(dict(spec, page="letter", sections=[{"heading": "S", "blocks": [
                {"type": "table", "header": ["A"], "rows": [["1"]],
                 "widths": [9638]}]}]), os.path.join(tmp, "lt2.docx"))
            raise AssertionError("A4 widths were accepted on a Letter page")
        except ValueError:
            pass

        # body_pt reaches the styles, and is refused outside a sane range
        with zipfile.ZipFile(build(dict(spec, body_pt=11),
                                   os.path.join(tmp, "bp.docx"))) as z:
            nrm = re.search(r'w:styleId="Normal".*?</w:style>',
                            z.read("word/styles.xml").decode(), re.S).group(0)
            assert '<w:sz w:val="22"/>' in nrm, nrm
        for bad_pt in (4, 30):
            try:
                validate(dict(spec, body_pt=bad_pt))
                raise AssertionError("validate accepted body_pt=%r" % bad_pt)
            except ValueError:
                pass
        # the lists are off by default, not unavailable
        # ... and when one is asked for it gets a section head, not a banner
        assert "Mục lục" in [p.text for p in
                             _D(build(dict(spec, layout="simple", toc=True),
                                      os.path.join(tmp, "flat2.docx"))).paragraphs]
        try:
            validate(dict(spec, layout="fancy"))
            raise AssertionError("validate accepted an unknown layout")
        except ValueError:
            pass

        # a brand reaches the cover, the footer and the document properties
        b = build(dict(spec, brand={"company": "ACME Corp"}),
                  os.path.join(tmp, "b.docx"))
        bd = _D(b)
        assert bd.tables[0].cell(0, 0).text.startswith("ACME Corp"), "brand not on cover"
        foot = bd.sections[0].footer.tables[0].rows[0].cells[0].text
        assert "ACME Corp" in foot, foot
        assert bd.core_properties.category.startswith("ACME Corp")
        assert load_brand({"brand": {"company": "X"}})["logo"] == BRAND_DEFAULTS["logo"], \
            "an override must not drop the rest of the brand"

        # a logo of any shape is fitted into its box, never stretched by width
        # alone - a square logo scaled to 2.9 cm wide would be 2.9 cm tall and
        # push the body down on every page
        def fitted(png, box):
            d0 = _D()
            with contextlib.redirect_stderr(_io.StringIO()):   # the shapes warn on purpose
                assert place_logo(d0.add_paragraph(), png, box, "test"), png
            ext = d0.paragraphs[0].runs[0]._r.find(qn("w:drawing"))[0].extent
            return ext.cx, ext.cy

        box = [2.9, 1.2]
        wide = solid_png(os.path.join(tmp, "wide.png"), "112233", 1000, 100)
        tall = solid_png(os.path.join(tmp, "tall.png"), "112233", 100, 500)
        square = solid_png(os.path.join(tmp, "sq.png"), "112233", 400, 400)
        for png in (wide, tall, square):
            cx, cy = fitted(png, box)
            assert cx <= Cm(box[0]) + 1 and cy <= Cm(box[1]) + 1, (png, cx, cy)
            assert cx > 0 and cy > 0, png
        assert fitted(wide, box)[0] > fitted(tall, box)[0], \
            "a wide logo should bind on width, a tall one on height"

        # an unreadable or missing logo warns and falls back to the company name
        for bad_logo in ("", os.path.join(tmp, "nope.png"),
                         os.path.join(ASSETS or tmp, "logo-mark.svg")):
            err = _io.StringIO()
            with contextlib.redirect_stderr(err):
                nd = _D(build(dict(spec, brand={"company": "NoLogo", "logo": bad_logo,
                                                "mark": bad_logo}),
                              os.path.join(tmp, "n.docx")))
            assert "warning:" in err.getvalue(), (bad_logo, err.getvalue())
            hdr = nd.sections[0].header.tables[0].rows[0].cells[0].text
            assert hdr == "NoLogo", (bad_logo, hdr)

        # the cover ground is generated from a hex colour, not shipped as a file
        import zlib
        blob = open(solid_png(os.path.join(tmp, "bg.png"), "#102030", 4, 3), "rb").read()
        raw = zlib.decompressobj().decompress(blob[blob.index(b"IDAT") + 4:])
        assert raw == (b"\x00" + b"\x10\x20\x30" * 4) * 3, raw
        build(dict(spec, brand={"cover_bg": "102030"}), os.path.join(tmp, "c.docx"))
        try:
            solid_png(os.path.join(tmp, "x.png"), "nothex")
            raise AssertionError("a malformed cover_bg was accepted")
        except ValueError:
            pass
        try:
            validate(dict(spec, brand={"colour": "red"}))
            raise AssertionError("validate accepted an unknown brand key")
        except ValueError:
            pass

        # Sibling skills render through this script, and their JSON is
        # documentation too. Built last: they set the globals for their own
        # page, font and body size, and everything above reads the defaults.
        for i, sib in enumerate(SIBLINGS):
            assert os.path.exists(sib), sib
            with open(sib, encoding="utf-8") as fh:
                sspec = json.load(fh)
            sdoc = _D(build(sspec, os.path.join(tmp, "sib%d.docx" % i)))
            assert sdoc.paragraphs[0].text.startswith("TITLE: "), sib
            hdr = sdoc.sections[0].header.tables[0].rows[0].cells[1].text
            assert hdr.startswith("TEST DOCUMENT: "), (sib, hdr)
            assert abs(sdoc.sections[0].page_width - Cm(21.59)) < 2000, sib
            with zipfile.ZipFile(os.path.join(tmp, "sib%d.docx" % i)) as z:
                assert "Arial" in z.read("word/styles.xml").decode(), sib

    print("selfcheck ok")


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "selfcheck":
        _selfcheck()
    else:
        sys.exit(main(sys.argv))
