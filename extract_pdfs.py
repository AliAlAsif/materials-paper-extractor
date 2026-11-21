"""
Paper data extractor v2  (runs on your computer, uses no Claude usage)

Libraries:  PyMuPDF (main engine)  +  pdfplumber (fallback)  +  pandas/openpyxl (Excel output)
Install:    pip install pymupdf pdfplumber pandas openpyxl

For every PDF in PDF_DIR it creates, inside OUT_DIR ("extracted"):
  extraction_results.xlsx   <- OPEN THIS FIRST. Sheets: Summary, Values, Property tables
  summary.csv, values_found.csv, property_table_rows.csv   (same data as CSV)
  tables/                   <- TABLES ONLY: every table as a CSV file; the first two lines are
                               "# Table N | full title" and "# page | source | engine"
  table_images/             <- a picture of each table (to check the CSV against)
  figures/                  <- FIGURES ONLY: every figure/graph cut out as a PNG, named by figure number
  figures_index.csv         <- each figure with its caption and page (use with WebPlotDigitizer)
  table_captions.csv        <- every "Table N" with its full title and page number
Tables and figures are kept completely separate: figure areas never contain tables, and
table files never contain figures. Each property-table row carries table_no + table_title.

How tables are found (most reliable first):
  1. PyMuPDF, ruled-line tables
  2. PyMuPDF, text-position mode   (journal tables without lines; pages with a "Table N" caption only)
  3. pdfplumber, text-position mode (only if 1 and 2 found nothing on that caption page)
Tables found in modes 2-3 are marked "textmode" and are less reliable.

IMPORTANT: this is a helper. Always check every value, its units and which sample
it belongs to against the PDF. Scanned PDFs have no text: they are flagged for manual work.

HOW TO RUN IN PYCHARM
  1. Put your PDFs in a folder called "pdfs" next to this file (sub-folders are fine).
  2. Right-click this file -> Run.
"""

import argparse
import csv
import hashlib
import re
import traceback
from html import escape
from pathlib import Path
from urllib.parse import quote

try:
    import pymupdf as fitz          # new import name
except ImportError:                  # older PyMuPDF
    import fitz

try:
    import pdfplumber
except ImportError:
    pdfplumber = None

try:
    import pandas as pd
except ImportError:
    pd = None

# ============================== SETTINGS ==============================
PDF_DIR = "pdfs"
OUT_DIR = "extracted"
MAX_COLS_SAVED = 14          # columns kept in the "Property tables" sheet
SAVE_FIGURES = True          # cut every figure out of the PDF as a PNG
SAVE_TABLE_IMAGES = True     # also save a picture of each table
ADD_CROP_HEADINGS = True     # put captions and paper references above PNG crops
CSV_COMMENT_HEADER = True    # put the table number + title at the top of each tables/*.csv
FIG_DPI = 200                # figure picture quality (200 is good for WebPlotDigitizer)
TABLE_DPI = 150
# ======================================================================

PROPS = {
    "yield strength": r"yield\s+(strength|stress)|0\.2\s*%\s*(proof|offset)|proof\s+stress|\bYS\b|\u03c3\s*0\.2|\u03c3y",
    "tensile strength (UTS)": r"ultimate\s+tensile|tensile\s+strength|\bUTS\b|\u03c3\s*(UTS|b)\b",
    "elongation": r"elongation|ductility|fracture\s+strain|strain\s+to\s+failure",
    "hardness": r"hardness|\bHV\d*\b|\bHRC\b|vickers",
    "Young's modulus": r"young'?s?\s+modul|elastic\s+modul|modulus\s+of\s+elasticity",
    "thermal conductivity": r"thermal\s+conductivity",
    "thermal diffusivity": r"thermal\s+diffusivity",
    "specific heat": r"specific\s+heat|heat\s+capacity",
    "thermal expansion (CTE)": r"thermal\s+expansion|\bCTE\b",
    "beta transus": r"transus",
    "density": r"relative\s+density|\bdensity\b",
    "fatigue": r"fatigue\s+(strength|limit|life)|endurance\s+limit",
}
PROPS = {k: re.compile(v, re.I) for k, v in PROPS.items()}

UNIT = (r"(MPa|GPa|kPa|HV\d*(\.\d+)?|HRC|HB|W\s*[/\u00b7.]?\s*\(?\s*m\s*[\u00b7.\s-]*K\)?|W\s*m-?1\s*K-?1|"
        r"J\s*[/\u00b7.]?\s*\(?\s*kg\s*[\u00b7.\s-]*K\)?|J\s*kg-?1\s*K-?1|mm2\s*/\s*s|mm\u00b2\s*/\s*s|m2\s*/\s*s|"
        r"10\s*-\s*6\s*(/|\s)?\s*K|\u00d7\s*10\s*-\s*6|\u00b0C|\s%|%|\sK\b|g\s*/\s*cm3|g\s*/\s*cm\u00b3)")
NUM = r"\d+(?:[.,]\d+)?"
VALUE = re.compile(rf"{NUM}(?:\s*(?:\u00b1|\+/-)\s*{NUM})?(?:\s*[-\u2013]\s*{NUM})?\s*{UNIT}", re.I)
ALLOY = re.compile(
    r"Ti[\s-]*6[\s-]*Al[\s-]*4[\s-]*V|Ti6Al4V|\bTC4\b|\bTC11\b|\bTA15\b|\bTi60\b|Ti[\s-]*5553|Ti[\s-]*1023|"
    r"Ti[\s-]*\d{1,2}(?:\.\d)?[\s-]*(?:Al|Nb|Zr|Mo|Ta|V|Sn|Cr|Fe)[\w.\-]*|\bTiAl\b|\bTi2AlNb\b|\bCP[\s-]*Ti\b|"
    r"Ti[\s-]*\d{2,5}[A-Z]?\b|titanium\s+alloys?", re.I)
DOI_RX = re.compile(r"10\.\d{4,9}/[^\s\"<>\]\)]+", re.I)
CAPTION_RX = re.compile(r"^\s*Table\s+[A-Z]?\d+", re.I)
FIG_CAPTION_RX = re.compile(r"^\s*(Fig(?:ure)?s?\.?\s*[A-Z]?\d+[a-z]?)", re.I)


def norm(s):
    s = str(s or "")
    # PDFs sometimes contain control bytes (illegal in Excel) and odd spaces
    s = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\ufffe\uffff]", " ", s)
    s = s.replace("\u2212", "-").replace("\u2010", "-").replace("\u2011", "-").replace("\u00a0", " ")
    return s


def clean_cell(c):
    return re.sub(r"\s+", " ", norm(c)).strip()


def props_in(text):
    return [k for k, rx in PROPS.items() if rx.search(text)]


def block_sentences(text):
    text = norm(text)
    text = re.sub(r"(\w)-\n([a-z])", r"\1\2", text)          # join hyphenated line breaks
    text = re.sub(r"\s+", " ", text)
    # protect abbreviations so they do not split sentences, then restore them
    text = re.sub(r"\b(Fig|Figs|Eq|Eqs|Ref|Refs|et al|No|vs|approx)\.", lambda m: m.group(1) + "\u0001", text)
    parts = re.split(r"(?<=[.;])\s+(?=[A-Z0-9\[])", text)
    return [p.replace("\u0001", ".").strip() for p in parts if len(p.strip()) > 20]


def clean_rows(rows):
    rows = [[clean_cell(c) for c in r] for r in (rows or [])]
    rows = [r for r in rows if any(r)]
    if not rows:
        return []
    width = max(len(r) for r in rows)
    rows = [r + [""] * (width - len(r)) for r in rows]
    keep = [j for j in range(width) if any(r[j] for r in rows)]
    return [[r[j] for j in keep] for r in rows]


def good_table(rows, textmode):
    if not rows:
        return False
    n_rows, n_cols = len(rows), len(rows[0])
    if n_rows < (3 if textmode else 2) or n_cols < (3 if textmode else 2):
        return False
    cells = [c for r in rows for c in r]
    filled = [c for c in cells if c]
    if len(filled) < 0.3 * len(cells):
        return False
    numeric_rows = sum(any(re.search(r"\d", c) for c in r[1:]) for r in rows[1:])
    if numeric_rows < (2 if textmode else 1):
        return False
    if textmode and (sum(len(c) for c in filled) / len(filled)) > 45:    # paragraphs, not a table
        return False
    return True


def overlap(a, b):
    x0, y0, x1, y1 = max(a[0], b[0]), max(a[1], b[1]), min(a[2], b[2]), min(a[3], b[3])
    if x1 <= x0 or y1 <= y0:
        return 0.0
    inter = (x1 - x0) * (y1 - y0)
    return inter / max(1.0, min((a[2] - a[0]) * (a[3] - a[1]), (b[2] - b[0]) * (b[3] - b[1])))


def cell_text(words, rect):
    """Text of one table cell: words whose centre lies inside the cell, left-to-right, top-to-bottom."""
    x0, y0, x1, y1 = rect
    inside = [w for w in words if x0 <= (w[0] + w[2]) / 2 <= x1 and y0 <= (w[1] + w[3]) / 2 <= y1]
    inside.sort(key=lambda w: (round(w[1] / 3), w[0]))
    return " ".join(w[4] for w in inside)


def mupdf_tables(page, strategy):
    """Table LAYOUT from PyMuPDF, cell TEXT from page words.
    (PyMuPDF's own cell reader can garble decimals such as 8.2, so it is not used.)"""
    out = []
    try:
        finder = page.find_tables(strategy=strategy)
    except Exception:
        return out
    words = page.get_text("words")
    for t in getattr(finder, "tables", []):
        try:
            rows = []
            for r in t.rows:
                rows.append(["" if c is None else cell_text(words, c) for c in r.cells])
            if not rows:
                rows = t.extract()
            out.append((tuple(t.bbox), rows))
        except Exception:
            continue
    return out


def plumber_tables(pl_pdf, pno):
    out = []
    if pl_pdf is None:
        return out
    try:
        page = pl_pdf.pages[pno]
        settings = {"vertical_strategy": "text", "horizontal_strategy": "text"}
        for t in page.find_tables(settings):
            out.append((tuple(t.bbox), t.extract()))
    except Exception:
        pass
    return out


CAPTION_START_RX = re.compile(r"^\s*(Table|TABLE|Tab\.)\s*([A-Z]?\d+)\b[.:]?\s*", re.I)
CAPTION_ANY_RX = re.compile(r"\b(Table|TABLE|Tab\.)\s*([A-Z]?\d+)\b[.:]?\s*", re.I)
# "Table 2 summarizes the ..." / "Table 4 shows ..." are in-text references, not titles
BODY_CAPTION_RX = re.compile(
    r"^\s*(?:Table|TABLE|Tab\.)\s*[A-Z]?\d+\s+(?:"
    r"summariz(?:e|es|ed)|shows?|showed|lists?|listed|presents?|presented|gives?|"
    r"reports?|reported|displays?|provides?|indicates?|compares?|compared|"
    r"contains?|includes?|outlines?|details?|also)\b", re.I)
HEADING_RX = re.compile(
    r"^\s*(?:\d+(?:\.\d+)*\.?\s+[A-Z]|Fig(?:ure)?s?\.?\s*[A-Z]?\d+|Table\s+[A-Z]?\d+)", re.I)
# prose that just happens to begin with "Table N" (a reference inside a paragraph)
PROSE_START_RX = re.compile(
    r"^(?:The|This|These|Those|It|As|In|For|At|However|Moreover|Therefore|Thus|Also|We|Our|Such|"
    r"There|Furthermore|Additionally|Besides|Finally|First|Second)\b", re.I)
PROSE_VERB_RX = re.compile(
    r"\b(?:is|are|was|were|has|have|had|can|could|may|might|will|would|should|shows|show|"
    r"presents|present|means|indicates)\b", re.I)
REF_TAIL_RX = re.compile(
    r"(?:shown|presented|listed|given|reported|summari[sz]ed|provided|displayed|detailed|described|"
    r"seen|found|included|tabulated|illustrated)\s+(?:in|as|to)\s*$", re.I)


def line_text(ln):
    return clean_cell("".join(s.get("text", "") for s in ln.get("spans", [])))


def looks_like_table_line(t):
    """A line that is really a table row/header (mostly numbers or one-word cells)."""
    toks = t.split()
    if len(toks) < 3:
        return False
    nums = sum(1 for w in toks if re.fullmatch(r"[-+]?\d+(?:[.,]\d+)?%?", w))
    return nums / len(toks) >= 0.5


def clean_title(title):
    """Keep the title only: cut a body sentence that was glued to it in the PDF text."""
    title = clean_cell(title).strip(" .;:-")
    m = re.search(r"(?<!e\.g)(?<!i\.e)(?<!al)\.\s+(?=[A-Z])", title)
    if m and len(title[:m.start()].split()) >= 4:
        title = title[:m.start()]
    return title.strip(" .;:-")


def _next_block_title(blocks, bi, block, label):
    """'Table 6' alone in a block: its title is usually the next block (sometimes rotated)."""
    brect = fitz.Rect(block["bbox"])
    for nb in blocks[bi + 1: bi + 3]:
        if nb.get("type") != 0:
            continue
        nrect = fitz.Rect(nb["bbox"])
        gap = nrect.y0 - brect.y1
        y_overlap = min(brect.y1, nrect.y1) - max(brect.y0, nrect.y0)
        x_overlap = min(brect.x1, nrect.x1) - max(brect.x0, nrect.x0)
        if gap > 90 and y_overlap < 20 and x_overlap <= 0:
            continue
        parts = []
        for nl in nb.get("lines", [])[:5]:
            nt = line_text(nl)
            if not nt:
                continue
            if HEADING_RX.match(nt) or looks_like_table_line(nt):
                break
            parts.append(nt)
            if nt.endswith(".") and len(nt) > 35:
                break
        if not parts:
            continue
        title = clean_cell(" ".join(parts))
        if len(title) < 3:
            continue
        return clean_cell(f"{label} {title[:400]}"), nrect
    return None, None


def page_captions(page):
    """Find 'Table N' captions line-by-line and rebuild the FULL title.

    Block-level matching missed captions that share a block with body text;
    line-level matching also lets us keep only the title (not the whole paragraph).
    Returns [{'bbox', 'label', 'number', 'title', 'full'}]."""
    out = []
    try:
        data = page.get_text("dict")
    except Exception:
        return out
    blocks = [b for b in data.get("blocks", []) if b.get("type") == 0]
    for bi, b in enumerate(blocks):
        lines = b.get("lines", [])
        for li, ln in enumerate(lines):
            txt = line_text(ln)
            m = CAPTION_START_RX.match(txt)
            if not m:
                continue
            parts = [txt]
            rect = fitz.Rect(ln["bbox"])
            j = li + 1
            while j < len(lines) and len(parts) < 7 and len(" ".join(parts)) < 480:
                nxt = line_text(lines[j])
                if not nxt:
                    j += 1
                    continue
                if HEADING_RX.match(nxt) or looks_like_table_line(nxt):
                    break
                parts.append(nxt)
                rect |= fitz.Rect(lines[j]["bbox"])
                j += 1
                if nxt.endswith(".") and len(nxt) > 35:
                    break
            full = clean_cell(" ".join(parts))
            # 'Table 6' alone: merge with the title in the following block
            if re.fullmatch(r"(?i)table\s+[A-Z]?\d+[.:]?", full):
                merged, extra = _next_block_title(blocks, bi, b, full)
                if merged:
                    full, rect = merged, (rect | extra if extra is not None else rect)
            mm = CAPTION_ANY_RX.search(full)
            if not mm:
                continue
            num = mm.group(2)
            label = f"Table {num}"
            title = clean_title(full[mm.end():])
            full = clean_cell(f"{label} {title}")
            prev_txt = line_text(lines[li - 1]) if li > 0 else ""
            if BODY_CAPTION_RX.match(full):
                continue                       # in-text reference such as "Table 2 shows ..."
            if title and not re.search(r"[A-Za-z\u00c0-\u024f]", title):
                continue                       # punctuation only, e.g. "Table 10)."
            if not title and li > 0:
                continue                       # reference in the middle of a paragraph
            if li > 0 and title and re.match(r"^[,;:.)]", title):
                continue                       # line-break reference such as "Table 4, Ti-31011 ..."
            if li > 0 and title and PROSE_START_RX.match(title) and PROSE_VERB_RX.search(title[:140]):
                continue                       # "Table 1. The equipment used in this study is ..."
            if li > 0 and REF_TAIL_RX.search(prev_txt):
                continue                       # " ... are shown in" + "Table 1. The ..."
            out.append({"bbox": tuple(rect), "label": label, "number": num,
                        "title": title, "full": full})
    seen, uniq = set(), []
    for c in out:
        key = (c["number"].upper(), c["title"][:50].lower())
        if key in seen:
            continue
        seen.add(key)
        uniq.append(c)
    return uniq


def _caption_gap(cap, tab, page_w):
    cap, tab = fitz.Rect(cap), fitz.Rect(tab)
    x_overlap = min(cap.x1, tab.x1) - max(cap.x0, tab.x0)
    denom = max(1.0, min(cap.width, tab.width))
    if x_overlap / denom < 0.15 and not (cap.width > 0.55 * page_w or tab.width > 0.55 * page_w):
        return None
    if cap.y1 <= tab.y0 + 5:                   # caption above the table
        gap = tab.y0 - cap.y1
    elif cap.y0 >= tab.y1 - 5:                 # caption below the table
        gap = (cap.y0 - tab.y1) + 22
    else:                                      # overlapping (e.g. rotated text)
        gap = 25.0
    if gap < -5 or gap > 170:
        return None
    return max(0.0, gap)


def _set_caption(t, c):
    t["caption_no"] = c["label"]
    t["caption_title"] = c["title"]
    t["caption_full"] = c["full"]


def assign_captions(captions, tables, page_w):
    """Attach each caption to at most one table (nearest first, then leftover pairing)."""
    for t in tables:
        t["caption_no"] = t["caption_title"] = t["caption_full"] = ""
    if not captions or not tables:
        return
    pairs = []
    for ci, c in enumerate(captions):
        for ti, t in enumerate(tables):
            g = _caption_gap(c["bbox"], t["bbox"], page_w)
            if g is not None:
                pairs.append((g, ci, ti))
    pairs.sort()
    used_c, used_t = set(), set()
    for _g, ci, ti in pairs:
        if ci in used_c or ti in used_t:
            continue
        _set_caption(tables[ti], captions[ci])
        used_c.add(ci)
        used_t.add(ti)
    rest_c = [i for i in range(len(captions)) if i not in used_c]
    rest_t = [i for i in range(len(tables)) if i not in used_t]
    if not rest_t:
        return
    if len(rest_c) == len(rest_t):
        cap_order = sorted(rest_c, key=lambda i: (captions[i]["bbox"][1], captions[i]["bbox"][0]))
        tab_order = sorted(rest_t, key=lambda i: (tables[i]["bbox"][1], tables[i]["bbox"][0]))
        for ci, ti in zip(cap_order, tab_order):
            _set_caption(tables[ti], captions[ci])
    elif len(rest_t) == 1 and rest_c:
        ti = rest_t[0]
        ci = min(rest_c, key=lambda i: abs(captions[i]["bbox"][1] - tables[ti]["bbox"][1]))
        _set_caption(tables[ti], captions[ci])


def is_metadata_table(rows):
    """Journal 'ARTICLE INFO / ABSTRACT', nomenclature or reference-list boxes are not data tables."""
    if not rows:
        return False
    head = re.sub(r"[^a-z]", "", " ".join(rows[0]).lower())
    if any(w in head for w in ("articleinfo", "abstract", "keywords", "articlehistory")):
        return True
    joined = " ".join(rows[0]).strip().lower()
    if re.match(r"^(article history|nomenclature|keywords|abbreviations|symbols|contents)\b", joined):
        return True
    if len(rows) >= 6:
        refs = sum(1 for r in rows if re.match(r"^\s*(?:\[\d+\]|\d+[.)]\s)", " ".join(r)))
        if refs >= 0.30 * len(rows):
            return True
    return False


def graphics_rects(page, table_boxes):
    """Rectangles of images and vector drawings on the page, excluding anything that belongs to a table."""
    page_area = page.rect.width * page.rect.height
    rects = []
    try:
        for info in page.get_image_info():
            rects.append(fitz.Rect(info["bbox"]))
    except Exception:
        pass
    try:
        for d in page.get_drawings():
            rects.append(fitz.Rect(d["rect"]))
    except Exception:
        pass
    keep = []
    for r in rects:
        if r.is_empty and r.width < 0.5 and r.height < 0.5:
            continue
        if r.width * r.height > 0.7 * page_area:           # page background
            continue
        if any(overlap(tuple(r), tb) > 0.5 for tb in table_boxes):  # belongs to a table
            continue
        keep.append(r)
    return keep


def find_figures(page, blocks, fig_caps, table_boxes):
    """For each 'Fig. N' caption find the graphic area just above (or below) it. Returns [(label, caption, rect)]."""
    gfx = graphics_rects(page, table_boxes)
    found, used = [], []
    text_blocks = [(fitz.Rect(b[:4]), re.sub(r"\s+", " ", norm(b[4])).strip()) for b in blocks]
    for (cb, label, cap) in fig_caps:
        def xover(r):
            w = min(cb.x1, r.x1) - max(cb.x0, r.x0)
            return w / max(1.0, min(cb.width, r.width))
        # nearest body text / caption above this caption in the same column = upper limit
        limit = 0.0
        for r, t in text_blocks:
            if r.y1 <= cb.y0 - 1 and xover(r) > 0.3 and (len(t) > 150 or FIG_CAPTION_RX.match(t) or CAPTION_RX.match(t)):
                limit = max(limit, r.y1)
        cluster = [g for g in gfx if g.y0 >= limit - 2 and g.y1 <= cb.y0 + 3 and xover(g) > 0.2]
        above = True
        if not cluster:      # some journals put the caption above the figure
            above = False
            nxt = min([r.y0 for r, t in text_blocks if r.y0 >= cb.y1 + 1 and xover(r) > 0.3 and len(t) > 150] + [page.rect.y1])
            cluster = [g for g in gfx if g.y0 >= cb.y1 - 3 and g.y1 <= nxt + 2 and xover(g) > 0.2]
        if not cluster:
            continue
        area = fitz.Rect(cluster[0])
        for g in cluster[1:]:
            area |= g
        # add short text (axis labels, legends) that sits just outside the graphics
        grown = fitz.Rect(area.x0 - 30, area.y0 - 28, area.x1 + 30, min(cb.y0, area.y1 + 40))
        for r, t in text_blocks:
            if len(t) < 70 and grown.intersects(r) and r.y1 <= cb.y0 + 1 and \
                    not (CAPTION_RX.match(t) or FIG_CAPTION_RX.match(t)):
                area |= r
        area = fitz.Rect(area.x0 - 4, area.y0 - 4, area.x1 + 4, area.y1 + 4) & page.rect
        if above:
            area.y1 = min(area.y1, cb.y0 - 0.5)      # never include the caption itself
        else:
            area.y0 = max(area.y0, cb.y1 + 0.5)
        if area.width < 40 or area.height < 40:
            continue
        if any(overlap(tuple(area), tb) > 0.5 for tb in table_boxes):
            continue
        found.append((label, cap, area))
    return found


def save_crop(page, clip, dpi, path, caption, source, doi):
    crop = page.get_pixmap(clip=clip, dpi=dpi)
    if not ADD_CROP_HEADINGS:
        crop.save(str(path))
        return

    paper = f"Source: {source} | PDF page {page.number + 1}"
    if doi:
        paper += f" | DOI: {doi}"
    heading = f"<p><b>{escape(caption)}</b></p><p>{escape(paper)}</p>"
    css = "body {font-family: sans-serif; font-size: 18pt} p {margin: 0 0 6pt}"
    width = max(crop.width, 520)
    padding = 14
    with fitz.open() as doc:
        height = 128
        while True:
            header = doc.new_page(width=width, height=height)
            spare, _ = header.insert_htmlbox(
                fitz.Rect(padding, padding, width - padding, height - padding),
                heading, css=css, scale_low=1 if height < 2048 else 0,
            )
            if spare >= 0:
                break
            doc.delete_page(0)
            height *= 2
        header_pixels = header.get_pixmap(clip=(0, 0, width, height - int(spare)))

    output = fitz.Pixmap(fitz.csRGB, fitz.IRect(0, 0, width, header_pixels.height + crop.height), False)
    output.clear_with(255)
    output.copy(header_pixels, header_pixels.irect)
    crop.set_origin(0, header_pixels.height)
    output.copy(crop, crop.irect)
    output.set_dpi(dpi, dpi)
    output.save(str(path))


def process_pdf(path, rel, out_dirs, results):
    out_tables, out_figs, out_tabimgs = out_dirs["tables"], out_dirs["figures"], out_dirs["table_images"]
    info = {"file": rel, "doi": "", "pages": 0, "status": "ok", "scanned_pages": 0, "figures_saved": 0,
            "tables_lines": 0, "tables_textmode": 0, "property_tables": 0, "captions": 0,
            "properties_mentioned": "", "value_sentences": 0}
    mentioned = set()
    doc = fitz.open(str(path))
    paper_title = clean_cell((doc.metadata or {}).get("title"))
    source = f"{paper_title} ({rel})" if paper_title else rel
    info["pages"] = len(doc)
    pl_pdf = None
    total_chars = 0
    prev_open = None          # caption of a table that continues onto the next page
    known_titles = {}         # table number -> full title (for "(continued)" captions)
    source_id = hashlib.sha256(Path(rel).as_posix().encode("utf-8")).hexdigest()[:12]
    clean_stem = re.sub(r"[^\w-]+", "_", Path(rel).stem)[:55]
    stem = f"{clean_stem}_{source_id}"
    try:
        for pno in range(len(doc)):
            page = doc[pno]
            blocks = [b for b in page.get_text("blocks", sort=True) if b[6] == 0]
            page_chars = sum(len(b[4]) for b in blocks)
            total_chars += page_chars
            if page_chars < 80:
                try:
                    covered = sum(fitz.Rect(i["bbox"]).width * fitz.Rect(i["bbox"]).height
                                  for i in page.get_image_info()) / max(1.0, page.rect.width * page.rect.height)
                except Exception:
                    covered = 0
                if covered > 0.5:
                    info["scanned_pages"] += 1
            if pno == 0 and not info["doi"]:
                m = DOI_RX.search(" ".join(norm(b[4]) for b in blocks))
                if m:
                    info["doi"] = m.group(0).rstrip(".,;")

            # ---- captions on this page (line-level, full title, no body references)
            caption_items = page_captions(page)
            for c in caption_items:                    # remember real titles first
                if c["title"] and not c["title"].lower().startswith("(continued)"):
                    known_titles[c["number"]] = c["title"]
            for c in caption_items:                    # fill "(continued)" from the real title
                if c["title"].lower().startswith("(continued)"):
                    prev = known_titles.get(c["number"])
                    if prev:
                        c["title"] = prev
                        c["full"] = clean_cell(f"Table {c['number']} (continued) {prev}")
            for c in caption_items:
                results["captions"].append({"file": rel, "page": pno + 1, "table_no": c["label"],
                                            "table_title": c["title"], "caption": c["full"]})
            info["captions"] += len(caption_items)

            # ---- value sentences
            for b in blocks:
                for s in block_sentences(b[4]):
                    hits = [k for k, rx in PROPS.items() if rx.search(s)]
                    if not hits:
                        continue
                    mentioned.update(hits)
                    vals = [m.group(0).strip() for m in VALUE.finditer(s)]
                    if vals:
                        alloy = "; ".join(sorted({m.group(0).strip() for m in ALLOY.finditer(s)}))[:80]
                        for k in hits:
                            results["values"].append({"file": rel, "doi": info["doi"], "page": pno + 1, "property": k,
                                                      "alloy_in_sentence": alloy, "values_found": "; ".join(vals[:6]),
                                                      "sentence": s[:420]})
                            info["value_sentences"] += 1

            # ---- tables: lines -> text mode (PyMuPDF) -> text mode (pdfplumber)
            accepted = []   # each: {bbox, rows, engine, caption_no, caption_title, caption_full}
            for bbox, raw in mupdf_tables(page, "lines"):
                rows = clean_rows(raw)
                if is_metadata_table(rows):
                    continue
                if good_table(rows, False):
                    accepted.append({"bbox": bbox, "rows": rows, "engine": "pymupdf-lines"})
            if caption_items:
                for bbox, raw in mupdf_tables(page, "text"):
                    if any(overlap(bbox, a["bbox"]) > 0.5 for a in accepted):
                        continue
                    rows = clean_rows(raw)
                    if is_metadata_table(rows):
                        continue
                    if good_table(rows, True):
                        accepted.append({"bbox": bbox, "rows": rows, "engine": "pymupdf-textmode"})
                if not accepted:
                    if pl_pdf is None and pdfplumber is not None:
                        pl_pdf = pdfplumber.open(str(path))
                    for bbox, raw in plumber_tables(pl_pdf, pno):
                        rows = clean_rows(raw)
                        if is_metadata_table(rows):
                            continue
                        if good_table(rows, True):
                            accepted.append({"bbox": bbox, "rows": rows, "engine": "pdfplumber-textmode"})

            # keep visual order, then attach the titles
            accepted.sort(key=lambda a: (round(a["bbox"][1] / 10), a["bbox"][0]))
            assign_captions(caption_items, accepted, page.rect.width)
            if prev_open:                       # table continued from the previous page
                for t in accepted:
                    if not t["caption_no"] and t["bbox"][1] < 0.28 * page.rect.height:
                        t["caption_no"] = prev_open[0]
                        t["caption_title"] = prev_open[1]
                        t["caption_full"] = clean_cell(f"{prev_open[0]} (continued) {prev_open[1]}")
                        break
            page_open = None
            if accepted:
                last = max(accepted, key=lambda a: a["bbox"][3])
                if last["caption_no"] and last["bbox"][3] > 0.72 * page.rect.height:
                    page_open = (last["caption_no"], last["caption_title"])

            if SAVE_FIGURES:
                fig_caps = []
                for b in blocks:
                    t = re.sub(r"\s+", " ", norm(b[4])).strip()
                    m = FIG_CAPTION_RX.match(t)
                    if m:
                        fig_caps.append((fitz.Rect(b[:4]), m.group(1), t))
                figure_names = set()
                for i, (label, cap, area) in enumerate(
                    find_figures(page, blocks, fig_caps, [a["bbox"] for a in accepted]), 1
                ):
                    safe = re.sub(r"[^\w]+", "_", label).strip("_")
                    if safe in figure_names:
                        safe = f"{safe}_{i}"
                    figure_names.add(safe)
                    fpath = out_figs / f"{stem}_p{pno + 1}_{safe}.png"
                    try:
                        save_crop(page, area, FIG_DPI, fpath, cap, source, info["doi"])
                        info["figures_saved"] += 1
                        results["figures"].append({"file": rel, "doi": info["doi"], "page": pno + 1, "figure": label,
                                                   "caption": cap, "image": str(fpath.name),
                                                   "figure_image": f"figures/{fpath.name}"})
                    except Exception as exc:
                        print(f"Could not save {fpath.name} from {rel}: {exc}")

            for k, t in enumerate(accepted, 1):
                bbox, rows, engine = t["bbox"], t["rows"], t["engine"]
                table_no, table_title, caption = t["caption_no"], t["caption_title"], t["caption_full"]
                image_link = ""
                if SAVE_TABLE_IMAGES:
                    image_path = out_tabimgs / f"{stem}_p{pno + 1}_t{k}.png"
                    try:
                        clip = fitz.Rect(bbox[0] - 4, bbox[1] - 4, bbox[2] + 4, bbox[3] + 4) & page.rect
                        for c in caption_items:                 # keep the caption out of the picture
                            cx0, cy0, cx1, cy1 = c["bbox"]
                            if cx1 <= clip.x0 or cx0 >= clip.x1:
                                continue
                            if clip.y0 < cy1 <= bbox[1] + 2:
                                clip.y0 = max(clip.y0, cy1 + 0.5)
                            elif bbox[3] - 2 <= cy0 < clip.y1:
                                clip.y1 = min(clip.y1, cy0 - 0.5)
                        heading = caption or f"{table_no or f'Detected table {k}'} (caption unavailable)"
                        save_crop(page, clip, TABLE_DPI, image_path, heading, source, info["doi"])
                        image_link = f"table_images/{image_path.name}"
                    except Exception as exc:
                        print(f"Could not save {image_path.name} from {rel}: {exc}")
                tag = "lines" if engine.endswith("lines") else "textmode"
                info["tables_lines" if tag == "lines" else "tables_textmode"] += 1
                table_path = out_tables / f"{stem}_p{pno + 1}_t{k}_{tag}.csv"
                with table_path.open("w", newline="", encoding="utf-8-sig") as f:
                    if CSV_COMMENT_HEADER:                      # title travels with the table data
                        f.write(f"# {table_no or 'table number not found'} | {table_title}\n")
                        f.write(f"# page {pno + 1} | {rel} | {engine}\n")
                    csv.writer(f).writerows(rows)
                blob = caption + " " + " ".join(c for r in rows for c in r)
                tprops = props_in(blob)
                if not tprops and re.search(r"MPa|GPa|\bHV\b|W\s*/\s*m|J\s*/\s*kg|10-6|mm2\s*/\s*s",
                                            " ".join(rows[0]), re.I):
                    tprops = ["property units in header"]
                if tprops:
                    info["property_tables"] += 1
                    header = " | ".join(c for c in rows[0] if c)
                    review = []
                    if tag == "textmode":
                        review.append("Check reconstructed columns")
                    if not table_title:
                        review.append("Check caption")
                    for ri, r in enumerate(rows[1:], 1):
                        rec = {"file": rel, "doi": info["doi"], "page": pno + 1, "table": k,
                               "table_no": table_no, "table_title": table_title, "engine": engine,
                               "caption": caption, "table_image": image_link,
                               "table_csv": f"tables/{table_path.name}", "review_notes": "; ".join(review),
                               "properties_in_table": "; ".join(tprops),
                               "header": header[:200], "row": ri}
                        for j in range(min(len(r), MAX_COLS_SAVED)):
                            rec[f"c{j + 1}"] = r[j]
                        results["table_rows"].append(rec)
            prev_open = page_open
    finally:
        n_pages_for_status = len(doc)
        doc.close()
        if pl_pdf is not None:
            pl_pdf.close()
    if total_chars < 60 * max(1, n_pages_for_status):
        info["status"] = "NO TEXT (scanned PDF: extract by hand)"
    elif info["scanned_pages"]:
        info["status"] = f"ok, but {info['scanned_pages']} scanned page(s)"
    info["properties_mentioned"] = "; ".join(sorted(mentioned))
    return info


def write_csv(path, rows):
    if not rows:
        path.write_text("", encoding="utf-8")
        return
    cols = []
    for r in rows:
        for k in r:
            if k not in cols:
                cols.append(k)
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=cols)
        w.writeheader()
        w.writerows(rows)


def write_excel(path, sheets):
    if pd is None:
        print("pandas/openpyxl not installed: skipping the Excel file (CSV files are still written).")
        return
    try:
        with pd.ExcelWriter(path, engine="openpyxl") as xw:
            for name, rows in sheets:
                df = pd.DataFrame(rows if rows else [{"info": "nothing found"}])
                df.to_excel(xw, sheet_name=name, index=False)
                ws = xw.sheets[name]
                ws.freeze_panes = "A2"
                ws.auto_filter.ref = ws.dimensions
                for cells in ws.iter_rows(min_row=2):
                    for cell in cells:
                        if cell.data_type == "f":
                            cell.data_type = "s"
                for col in ws.columns:
                    width = min(60, max(len(str(c.value)) if c.value is not None else 0 for c in col[:200]) + 2)
                    if col[0].value in {"table_image", "table_csv", "figure_image"}:
                        for cell in col[1:]:
                            if cell.value:
                                cell.hyperlink = quote(cell.value, safe="/")
                                cell.style = "Hyperlink"
                                cell.value = "Open CSV" if col[0].value == "table_csv" else "Open PNG"
                        width = 18
                    ws.column_dimensions[col[0].column_letter].width = max(8, width)
    except Exception as e:
        print(f"Could not write the Excel file ({type(e).__name__}). CSV files are still written.")


def main():
    parser = argparse.ArgumentParser(description="Extract candidate material properties from research PDFs.")
    parser.add_argument("--pdf-dir", type=Path, default=Path(PDF_DIR), help="directory containing input PDFs")
    parser.add_argument("--out-dir", type=Path, default=Path(OUT_DIR), help="directory for extraction results")
    args = parser.parse_args()
    root = args.pdf_dir
    pdfs = sorted(root.rglob("*.pdf"))
    if not pdfs:
        parser.error(f"No PDFs found in {root}")
    out_dir = args.out_dir
    out_dirs = {"tables": out_dir / "tables", "figures": out_dir / "figures", "table_images": out_dir / "table_images"}
    for d in out_dirs.values():
        d.mkdir(parents=True, exist_ok=True)
    print(f"{len(pdfs)} PDFs found in '{root}'")
    results = {"values": [], "table_rows": [], "captions": [], "figures": []}
    summary = []
    for n, p in enumerate(pdfs, 1):
        rel = str(p.relative_to(root))
        try:
            info = process_pdf(p, rel, out_dirs, results)
        except Exception as e:
            traceback.print_exc()
            info = {"file": rel, "status": f"ERROR {type(e).__name__}: {e}"}
        summary.append(info)
        print(f"[{n}/{len(pdfs)}] {rel[:52]:52} {info.get('status', '')[:28]:28} "
              f"tables={info.get('tables_lines', 0) + info.get('tables_textmode', 0)} "
              f"(property tables={info.get('property_tables', 0)}) figures={info.get('figures_saved', 0)} values={info.get('value_sentences', 0)}")

    write_csv(out_dir / "summary.csv", summary)
    write_csv(out_dir / "values_found.csv", results["values"])
    write_csv(out_dir / "property_table_rows.csv", results["table_rows"])
    write_csv(out_dir / "table_captions.csv", results["captions"])
    write_csv(out_dir / "figures_index.csv", results["figures"])
    write_excel(out_dir / "extraction_results.xlsx",
                [("Summary", summary), ("Values", results["values"]),
                 ("Property tables", results["table_rows"]), ("Table captions", results["captions"]),
                 ("Figures", results["figures"])])
    print(f"\nDone. Open '{out_dir / 'extraction_results.xlsx'}' first.")
    print("Papers with NO TEXT or ERROR need to be done by hand:")
    for s in summary:
        if not str(s.get("status", "")).startswith("ok"):
            print("  -", s["file"], "->", s["status"])


if __name__ == "__main__":
    main()
