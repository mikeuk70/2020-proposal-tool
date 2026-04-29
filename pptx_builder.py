"""
20.20 Proposal PPTX Builder — python-pptx version
Builds a clean branded deck matching real 20.20 proposal layouts.
Uses python-pptx throughout — no XML hacking.
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
import re, os

# ── DIMENSIONS ────────────────────────────────────────────────────────────────
W = Inches(13.33)   # widescreen 16:9
H = Inches(7.5)

# ── PALETTE ───────────────────────────────────────────────────────────────────
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
BLACK  = RGBColor(0x1A, 0x1A, 0x1A)
DARK   = RGBColor(0x11, 0x14, 0x18)   # near-black cover bg
MID    = RGBColor(0x2D, 0x2D, 0x2D)   # body text
GREY   = RGBColor(0x88, 0x88, 0x88)   # secondary / footer
LGREY  = RGBColor(0xF2, 0xF1, 0xEE)   # light bg
RULE   = RGBColor(0xDD, 0xDB, 0xD5)   # divider lines
DEFAULT_ACCENT = RGBColor(0xE9, 0x71, 0x32)  # 20.20 orange

# ── FONTS ─────────────────────────────────────────────────────────────────────
# Filson Pro is 20.20's licensed font — falls back to system fonts on machines
# that don't have it. We embed the name so PowerPoint uses it when available.
F_HEAD  = 'Filson Pro Heavy'    # headings
F_BODY  = 'Filson Pro'          # body text
F_FALL  = 'Arial'               # fallback

def _accent(client_name):
    """Return accent RGBColor for client, falling back to 20.20 orange."""
    COLOURS = {
        'aston villa': (0x5C, 0x1A, 0x2E), 'villa': (0x5C, 0x1A, 0x2E),
        'newcastle': (0xC9, 0xA8, 0x4C), 'nufc': (0xC9, 0xA8, 0x4C),
        'brighton': (0x00, 0x57, 0xB8), 'bhafc': (0x00, 0x57, 0xB8),
        'arsenal': (0xEF, 0x01, 0x07),
        'liverpool': (0xC8, 0x10, 0x2E),
        'chelsea': (0x03, 0x46, 0x94),
        'crystal palace': (0x1B, 0x45, 0x8F), 'cpfc': (0x1B, 0x45, 0x8F),
        'leeds': (0xFF, 0xCD, 0x00), 'lufc': (0xFF, 0xCD, 0x00),
        'sunderland': (0xEB, 0x17, 0x2B), 'safc': (0xEB, 0x17, 0x2B),
        'west ham': (0x7A, 0x26, 0x3A),
        'manchester city': (0x6C, 0xAB, 0xDD), 'man city': (0x6C, 0xAB, 0xDD),
        'manchester united': (0xDA, 0x29, 0x1C), 'man utd': (0xDA, 0x29, 0x1C),
        'tottenham': (0x13, 0x22, 0x57), 'spurs': (0x13, 0x22, 0x57),
        'luton': (0xF7, 0x83, 0x1A), 'luton town': (0xF7, 0x83, 0x1A),
        'sheffield': (0xEE, 0x27, 0x37),
        'nottingham': (0xE5, 0x32, 0x33),
        'wolverhampton': (0xFD, 0xB9, 0x13), 'wolves': (0xFD, 0xB9, 0x13),
    }
    cl = (client_name or '').lower()
    for key, rgb in COLOURS.items():
        if key in cl:
            return RGBColor(*rgb)
    return DEFAULT_ACCENT


# ── TEXT HELPERS ──────────────────────────────────────────────────────────────
def _clean(txt):
    if not txt: return ''
    txt = re.sub(r'\*\*([^*]+)\*\*', r'\1', txt)
    txt = re.sub(r'\*([^*]+)\*', r'\1', txt)
    txt = re.sub(r'^#{1,3}\s*', '', txt, flags=re.MULTILINE)
    txt = re.sub(r'\n{3,}', '\n\n', txt)
    return txt.strip()

def _paragraphs(txt):
    """Split text into (type, content) tuples: 'bullet' or 'prose'."""
    result = []
    for line in _clean(txt).split('\n'):
        s = line.strip()
        if not s:
            continue
        # Skip internal headings Claude adds
        if re.match(r'^(stage \d|riba|your brief|cover letter|our approach|next steps|fees)', s, re.I) and '.' not in s:
            continue
        if s.startswith(('-', '\u2022', '*')):
            result.append(('bullet', re.sub(r'^[-\u2022*]\s*', '', s)))
        elif re.match(r'^\d+[.):]', s):
            result.append(('bullet', re.sub(r'^\d+[.):]\s*', '', s)))
        else:
            result.append(('prose', s))
    return result

def _get_prose(txt, max_sentences=3):
    paras = [c for t, c in _paragraphs(txt) if t == 'prose']
    text = ' '.join(paras)
    sentences = re.split(r'(?<=[.!?])\s+', text)
    return ' '.join(sentences[:max_sentences]).strip()

def _get_bullets(txt, max_n=8):
    bullets = [c for t, c in _paragraphs(txt) if t == 'bullet']
    if not bullets:
        bullets = [c for t, c in _paragraphs(txt) if t == 'prose' and len(c) > 20]
    return bullets[:max_n]


# ── DRAWING HELPERS ───────────────────────────────────────────────────────────
def _bg(slide, colour):
    """Fill slide background with solid colour."""
    from pptx.oxml.ns import qn
    from lxml import etree
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = colour

def _rect(slide, x, y, w, h, colour, alpha=None):
    """Add a filled rectangle."""
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE
        x, y, w, h
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = colour
    shape.line.fill.background()
    return shape

def _line(slide, x, y, w, colour=None):
    """Add a thin horizontal rule."""
    c = colour or RULE
    shape = slide.shapes.add_shape(1, x, y, w, Pt(1))
    shape.fill.solid()
    shape.fill.fore_color.rgb = c
    shape.line.fill.background()

def _logo(slide, dark=False):
    """Add 20.20 logo text top right."""
    tb = slide.shapes.add_textbox(W - Inches(0.9), Inches(0.15), Inches(0.75), Inches(0.6))
    tf = tb.text_frame
    tf.word_wrap = False
    p1 = tf.paragraphs[0]
    p1.alignment = PP_ALIGN.CENTER
    run = p1.add_run()
    run.text = '20'
    run.font.name = F_HEAD
    run.font.size = Pt(13)
    run.font.bold = True
    run.font.color.rgb = WHITE if dark else BLACK

    p2 = tf.add_paragraph()
    p2.alignment = PP_ALIGN.CENTER
    r2 = p2.add_run()
    r2.text = '20'
    r2.font.name = F_HEAD
    r2.font.size = Pt(13)
    r2.font.bold = True
    r2.font.color.rgb = WHITE if dark else BLACK

    # Red dot
    dot = slide.shapes.add_shape(1,
        W - Inches(0.59), Inches(0.44), Inches(0.09), Inches(0.09))
    dot.fill.solid()
    dot.fill.fore_color.rgb = RGBColor(0xE8, 0x25, 0x1A)
    dot.line.fill.background()

def _section_label(slide, label, accent):
    """Small uppercase section label top left."""
    tb = slide.shapes.add_textbox(Inches(0.5), Inches(0.12), Inches(5), Inches(0.22))
    tf = tb.text_frame
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = label.upper()
    run.font.name = F_BODY
    run.font.size = Pt(8)
    run.font.bold = True
    run.font.color.rgb = accent

def _footer(slide):
    """Confidential footer bottom left."""
    tb = slide.shapes.add_textbox(Inches(0.5), H - Inches(0.3), Inches(9), Inches(0.2))
    tf = tb.text_frame
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = 'CONFIDENTIAL  \u00a9  20.20 Limited 2026'
    run.font.name = F_BODY
    run.font.size = Pt(7.5)
    run.font.color.rgb = GREY

def _heading(slide, text, x=None, y=None, w=None, size=26, colour=None, bold=True):
    x = x or Inches(0.5)
    y = y or Inches(0.28)
    w = w or (W - Inches(1.2))
    tb = slide.shapes.add_textbox(x, y, w, Inches(1.0))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    run = p.add_run()
    run.text = text.upper() if bold else text
    run.font.name = F_HEAD
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = colour or BLACK
    return tb

def _body_text(slide, paragraphs_data, x, y, w, h, size=12.5):
    """
    Add a text box with mixed prose and bullets.
    paragraphs_data: list of (type, text) — type is 'prose', 'bullet', or 'spacer'
    """
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    first = True
    for typ, text in paragraphs_data:
        if first:
            p = tf.paragraphs[0]
            first = False
        else:
            p = tf.add_paragraph()
        if typ == 'spacer':
            p.space_after = Pt(4)
            continue
        if typ == 'bullet':
            from pptx.oxml.ns import qn
            from lxml import etree
            pPr = p._pPr if p._pPr is not None else p._p.get_or_add_pPr()
            buNone = pPr.find(qn('a:buNone'))
            if buNone is not None:
                pPr.remove(buNone)
            buChar = etree.SubElement(pPr, qn('a:buChar'))
            buChar.set('char', '\u2022')
            p.space_before = Pt(2)
        run = p.add_run()
        run.text = text
        run.font.name = F_BODY
        run.font.size = Pt(size)
        run.font.color.rgb = MID
        if typ == 'prose':
            p.space_after = Pt(6)
    return tb


# ── SLIDE BUILDERS ────────────────────────────────────────────────────────────

def slide_cover(prs, venue, client, contact, role, date_s, accent):
    """Slide 1: Dark cover with image placeholder."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank
    _bg(slide, DARK)

    # Image placeholder directive (right half)
    tb_img = slide.shapes.add_textbox(Inches(6.5), Inches(0), Inches(6.83), H)
    tf = tb_img.text_frame
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = '[IMAGE: Full-bleed stadium or hospitality photography — atmospheric, premium, on-brand]'
    run.font.name = F_BODY
    run.font.size = Pt(9)
    run.font.color.rgb = RGBColor(0x44, 0x44, 0x44)

    # Left panel content
    # 20.20 logo
    logo_tb = slide.shapes.add_textbox(Inches(0.5), Inches(0.4), Inches(1.2), Inches(0.8))
    tf2 = logo_tb.text_frame
    p2 = tf2.paragraphs[0]
    r2 = p2.add_run()
    r2.text = '20\n20'
    r2.font.name = F_HEAD
    r2.font.size = Pt(18)
    r2.font.bold = True
    r2.font.color.rgb = WHITE

    # Accent rule
    _rect(slide, Inches(0.5), Inches(1.4), Inches(0.6), Inches(0.04), accent)

    # Project name
    tb_v = slide.shapes.add_textbox(Inches(0.5), Inches(1.55), Inches(5.5), Inches(1.2))
    tf_v = tb_v.text_frame
    tf_v.word_wrap = True
    p_v = tf_v.paragraphs[0]
    r_v = p_v.add_run()
    r_v.text = venue
    r_v.font.name = F_HEAD
    r_v.font.size = Pt(36)
    r_v.font.bold = True
    r_v.font.color.rgb = WHITE

    # Subtitle
    tb_sub = slide.shapes.add_textbox(Inches(0.5), Inches(2.9), Inches(5.5), Inches(0.4))
    tf_sub = tb_sub.text_frame
    p_sub = tf_sub.paragraphs[0]
    r_sub = p_sub.add_run()
    r_sub.text = 'Hospitality design proposal'
    r_sub.font.name = F_BODY
    r_sub.font.size = Pt(14)
    r_sub.font.color.rgb = RGBColor(0xCC, 0xCC, 0xCC)

    # Prepared for
    tb_pr = slide.shapes.add_textbox(Inches(0.5), H - Inches(1.2), Inches(5.5), Inches(0.8))
    tf_pr = tb_pr.text_frame
    p_pr = tf_pr.paragraphs[0]
    r_pr = p_pr.add_run()
    r_pr.text = client
    r_pr.font.name = F_HEAD
    r_pr.font.size = Pt(13)
    r_pr.font.bold = True
    r_pr.font.color.rgb = WHITE

    p_pr2 = tf_pr.add_paragraph()
    r_pr2 = p_pr2.add_run()
    r_pr2.text = f'Prepared for {contact}' + (f', {role}' if role else '') + (f'  |  {date_s}' if date_s else '')
    r_pr2.font.name = F_BODY
    r_pr2.font.size = Pt(10)
    r_pr2.font.color.rgb = GREY


def slide_hello(prs, body_text, accent):
    """Slide 2: INTRODUCTION / HELLO — full text letter."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(slide, WHITE)
    _logo(slide)
    _section_label(slide, 'Introduction', accent)
    _line(slide, Inches(0.5), Inches(0.9), W - Inches(1.0))

    # HELLO heading
    _heading(slide, 'Hello', y=Inches(0.95), size=28, bold=True)
    _line(slide, Inches(0.5), Inches(1.62), W - Inches(1.0))

    # Full letter text
    paras = _paragraphs(body_text)
    # Filter out heading-like first lines
    paras = [(t, c) for t, c in paras if not (t == 'prose' and len(c) < 25 and '.' not in c)]

    _body_text(slide, paras, Inches(0.5), Inches(1.7), W - Inches(1.0), Inches(5.0), size=12.5)
    _footer(slide)


def slide_brief(prs, body_text, accent):
    """Slide 3: YOUR BRIEF — full text, left column."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(slide, WHITE)
    _logo(slide)
    _section_label(slide, 'Your brief', accent)
    _line(slide, Inches(0.5), Inches(0.9), W - Inches(1.0))
    _heading(slide, 'Our understanding', y=Inches(0.95), size=26)
    _line(slide, Inches(0.5), Inches(1.7), W - Inches(1.0))

    paras = _paragraphs(body_text)
    paras = [(t, c) for t, c in paras if not (t == 'prose' and len(c) < 30 and '.' not in c)]

    # Left column prose, right column bullets if both exist
    prose_items = [(t, c) for t, c in paras if t == 'prose']
    bullet_items = [(t, c) for t, c in paras if t == 'bullet']

    if bullet_items:
        _body_text(slide, prose_items, Inches(0.5), Inches(1.8), Inches(5.8), Inches(5.0))
        # Vertical rule
        _rect(slide, Inches(6.5), Inches(1.8), Inches(0.01), Inches(4.8), RULE)
        _body_text(slide, bullet_items, Inches(6.65), Inches(1.8), Inches(6.0), Inches(5.0))
    else:
        _body_text(slide, prose_items, Inches(0.5), Inches(1.8), W - Inches(1.0), Inches(5.0))

    _footer(slide)


def slide_stage(prs, section_label, stage_title, body_text, accent):
    """Stage slide: section label, bold title, Scope left / Deliverables right, fee bottom right."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(slide, WHITE)
    _logo(slide)
    _section_label(slide, section_label, accent)
    _line(slide, Inches(0.5), Inches(0.9), W - Inches(1.0))

    # Title
    _heading(slide, stage_title, y=Inches(0.95), size=26)
    _line(slide, Inches(0.5), Inches(1.7), W - Inches(1.0))

    # Scope / Deliverables headings
    tb_sh = slide.shapes.add_textbox(Inches(0.5), Inches(1.78), Inches(5.8), Inches(0.3))
    r_sh = tb_sh.text_frame.paragraphs[0].add_run()
    r_sh.text = 'Scope'
    r_sh.font.name = F_HEAD
    r_sh.font.size = Pt(13)
    r_sh.font.bold = True
    r_sh.font.color.rgb = accent

    tb_dh = slide.shapes.add_textbox(Inches(6.65), Inches(1.78), Inches(6.0), Inches(0.3))
    r_dh = tb_dh.text_frame.paragraphs[0].add_run()
    r_dh.text = 'Deliverables'
    r_dh.font.name = F_HEAD
    r_dh.font.size = Pt(13)
    r_dh.font.bold = True
    r_dh.font.color.rgb = accent

    # Vertical rule
    _rect(slide, Inches(6.5), Inches(1.78), Inches(0.01), Inches(4.6), RULE)

    # Scope prose (left)
    prose = _get_prose(body_text, 4)
    if prose:
        _body_text(slide, [('prose', prose)], Inches(0.5), Inches(2.15), Inches(5.8), Inches(4.2), size=12)

    # Deliverables bullets (right)
    bullets = _get_bullets(body_text, 8)
    if bullets:
        _body_text(slide, [('bullet', b) for b in bullets],
                   Inches(6.65), Inches(2.15), Inches(6.0), Inches(4.2), size=11.5)

    # Fee bottom right
    _line(slide, Inches(6.65), H - Inches(0.85), Inches(6.0), RULE)
    tb_fee = slide.shapes.add_textbox(Inches(6.65), H - Inches(0.82), Inches(6.0), Inches(0.35))
    tf_fee = tb_fee.text_frame
    p_tot = tf_fee.paragraphs[0]
    r_tot = p_tot.add_run()
    r_tot.text = 'Total'
    r_tot.font.name = F_BODY
    r_tot.font.size = Pt(8)
    r_tot.font.color.rgb = GREY

    p_fee = tf_fee.add_paragraph()
    r_fee = p_fee.add_run()
    r_fee.text = '[FEE: TBC — rate card required]'
    r_fee.font.name = F_HEAD
    r_fee.font.size = Pt(16)
    r_fee.font.bold = True
    r_fee.font.color.rgb = BLACK

    _footer(slide)


def slide_fees(prs, stages_data, accent):
    """Fees and timings slide matching real 20.20 format."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(slide, WHITE)
    _logo(slide)
    _section_label(slide, 'Our methodology', accent)
    _line(slide, Inches(0.5), Inches(0.9), W - Inches(1.0))
    _heading(slide, 'Summary fees and timings', y=Inches(0.95), size=24)
    _line(slide, Inches(0.5), Inches(1.68), W - Inches(1.0))

    # Column headers
    col_heads = ['', 'Fees', 'Timings', 'Invoicing schedule']
    col_x = [Inches(0.5), Inches(4.8), Inches(6.2), Inches(7.5)]
    for i, (h_txt, hx) in enumerate(zip(col_heads, col_x)):
        if not h_txt: continue
        tb = slide.shapes.add_textbox(hx, Inches(1.75), Inches(1.8), Inches(0.25))
        r = tb.text_frame.paragraphs[0].add_run()
        r.text = h_txt
        r.font.name = F_BODY
        r.font.size = Pt(9)
        r.font.color.rgb = GREY

    row_h = Inches(0.9)
    row_y_start = Inches(2.05)

    stages = stages_data[:4]
    for i, stage in enumerate(stages):
        y = row_y_start + i * row_h
        bg_col = LGREY if i % 2 == 0 else WHITE
        _rect(slide, Inches(0.5), y, W - Inches(1.0), row_h - Inches(0.04), bg_col)

        # Stage number
        tb_n = slide.shapes.add_textbox(Inches(0.5), y + Inches(0.1), Inches(0.45), Inches(0.7))
        r_n = tb_n.text_frame.paragraphs[0].add_run()
        r_n.text = str(i + 1)
        r_n.font.name = F_HEAD
        r_n.font.size = Pt(28)
        r_n.font.bold = True
        r_n.font.color.rgb = BLACK

        # Stage title
        tb_t = slide.shapes.add_textbox(Inches(1.0), y + Inches(0.1), Inches(3.6), Inches(0.65))
        tf_t = tb_t.text_frame
        tf_t.word_wrap = True
        p_t = tf_t.paragraphs[0]
        r_t = p_t.add_run()
        r_t.text = stage.get('title', '').upper()
        r_t.font.name = F_BODY
        r_t.font.size = Pt(10)
        r_t.font.bold = True
        r_t.font.color.rgb = BLACK
        if stage.get('sub'):
            p_s = tf_t.add_paragraph()
            r_s = p_s.add_run()
            r_s.text = stage['sub']
            r_s.font.name = F_BODY
            r_s.font.size = Pt(9)
            r_s.font.color.rgb = MID

        # Fee
        tb_f = slide.shapes.add_textbox(Inches(4.7), y + Inches(0.2), Inches(1.4), Inches(0.45))
        r_f = tb_f.text_frame.paragraphs[0].add_run()
        r_f.text = stage.get('fee', '[FEE: TBC]')
        r_f.font.name = F_HEAD
        r_f.font.size = Pt(11)
        r_f.font.bold = True
        r_f.font.color.rgb = BLACK

        # Timing
        tb_ti = slide.shapes.add_textbox(Inches(6.1), y + Inches(0.2), Inches(1.2), Inches(0.45))
        r_ti = tb_ti.text_frame.paragraphs[0].add_run()
        r_ti.text = stage.get('timing', '')
        r_ti.font.name = F_BODY
        r_ti.font.size = Pt(10)
        r_ti.font.color.rgb = MID

        # Invoicing
        tb_inv = slide.shapes.add_textbox(Inches(7.4), y + Inches(0.1), Inches(5.0), Inches(0.65))
        tf_inv = tb_inv.text_frame
        tf_inv.word_wrap = True
        r_inv = tf_inv.paragraphs[0].add_run()
        r_inv.text = stage.get('invoicing', '50% at start, 50% at completion')
        r_inv.font.name = F_BODY
        r_inv.font.size = Pt(9.5)
        r_inv.font.color.rgb = MID

    # Total row
    total_y = row_y_start + len(stages) * row_h + Inches(0.1)
    _line(slide, Inches(0.5), total_y, W - Inches(1.0), BLACK)
    tb_tot = slide.shapes.add_textbox(Inches(1.0), total_y + Inches(0.08), Inches(3.5), Inches(0.5))
    tf_tot = tb_tot.text_frame
    p_tot = tf_tot.paragraphs[0]
    r_tot = p_tot.add_run()
    r_tot.text = 'TOTAL 20.20 FEES'
    r_tot.font.name = F_HEAD
    r_tot.font.size = Pt(10)
    r_tot.font.bold = True
    r_tot.font.color.rgb = BLACK

    tb_totf = slide.shapes.add_textbox(Inches(4.7), total_y + Inches(0.08), Inches(1.4), Inches(0.35))
    r_totf = tb_totf.text_frame.paragraphs[0].add_run()
    r_totf.text = '[FEE: TBC]'
    r_totf.font.name = F_HEAD
    r_totf.font.size = Pt(11)
    r_totf.font.bold = True
    r_totf.font.color.rgb = BLACK

    # Small print
    tb_sm = slide.shapes.add_textbox(Inches(0.5), H - Inches(0.5), Inches(9), Inches(0.28))
    r_sm = tb_sm.text_frame.paragraphs[0].add_run()
    r_sm.text = 'Fees are exclusive of VAT, 3rd party costs, general expenses and travel. Subject to contract.'
    r_sm.font.name = F_BODY
    r_sm.font.size = Pt(8)
    r_sm.font.italic = True
    r_sm.font.color.rgb = GREY

    _footer(slide)


def slide_divider(prs, word, accent):
    """Dark section divider slide."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(slide, DARK)
    _logo(slide, dark=True)

    tb = slide.shapes.add_textbox(Inches(0.5), H - Inches(2.2), W - Inches(1.0), Inches(1.8))
    tf = tb.text_frame
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = word
    run.font.name = F_HEAD
    run.font.size = Pt(72)
    run.font.bold = True
    run.font.color.rgb = WHITE


def slide_next_steps(prs, body_text, accent):
    """Next steps — 4 numbered cards."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(slide, WHITE)
    _logo(slide)
    _section_label(slide, 'Next steps', accent)
    _line(slide, Inches(0.5), Inches(0.9), W - Inches(1.0))
    _heading(slide, 'Next steps', y=Inches(0.95), size=26)
    _line(slide, Inches(0.5), Inches(1.68), W - Inches(1.0))

    bullets = _get_bullets(body_text, 4)
    default_steps = [
        ('Review this proposal', 'Share with your team. Note any questions, changes or additions.'),
        ('Return your feedback', 'Send us your comments and we will revise and reissue.'),
        ('Site visit and kick-off', 'If you are minded to appoint us, we propose visiting the venue together.'),
        ('Appointment', 'We can move quickly. We are ready to mobilise from instruction.'),
    ]

    card_w = Inches(2.9)
    card_h = Inches(4.2)
    card_y = Inches(1.85)
    gap = Inches(0.22)
    start_x = Inches(0.5)

    for i in range(4):
        x = start_x + i * (card_w + gap)
        _rect(slide, x, card_y, card_w, card_h, LGREY)

        # Number
        tb_n = slide.shapes.add_textbox(x + Inches(0.18), card_y + Inches(0.15), Inches(0.6), Inches(0.55))
        r_n = tb_n.text_frame.paragraphs[0].add_run()
        r_n.text = f'0{i+1}'
        r_n.font.name = F_HEAD
        r_n.font.size = Pt(22)
        r_n.font.bold = True
        r_n.font.color.rgb = accent

        _rect(slide, x + Inches(0.18), card_y + Inches(0.75), Inches(1.5), Inches(0.03), accent)

        # Title
        if i < len(bullets):
            # Parse from generated content
            title = bullets[i]
            desc = ''
        else:
            title, desc = default_steps[i]

        tb_t = slide.shapes.add_textbox(x + Inches(0.18), card_y + Inches(0.85), card_w - Inches(0.3), Inches(0.55))
        tf_t = tb_t.text_frame
        tf_t.word_wrap = True
        r_t = tf_t.paragraphs[0].add_run()
        r_t.text = title
        r_t.font.name = F_HEAD
        r_t.font.size = Pt(11)
        r_t.font.bold = True
        r_t.font.color.rgb = BLACK

        if desc:
            tb_d = slide.shapes.add_textbox(x + Inches(0.18), card_y + Inches(1.5), card_w - Inches(0.3), Inches(2.3))
            tf_d = tb_d.text_frame
            tf_d.word_wrap = True
            r_d = tf_d.paragraphs[0].add_run()
            r_d.text = desc
            r_d.font.name = F_BODY
            r_d.font.size = Pt(10)
            r_d.font.color.rgb = MID

    _footer(slide)


# ── MAIN BUILD FUNCTION ───────────────────────────────────────────────────────

def find_section(sections, *keys):
    for key in keys:
        kl = key.lower()
        for sec in sections:
            h = sec.get('heading', '').lower()
            if kl in h or h in kl:
                return _clean(sec.get('body', ''))
    return ''

def build_pptx_clean(sections, meta, output_path):
    """Build a complete branded proposal PPTX using python-pptx."""
    from pptx import Presentation
    from pptx.util import Inches, Pt, Emu

    client  = meta.get('client', '')
    venue   = meta.get('venue', 'Project')
    contact = meta.get('contact', '')
    role    = meta.get('role', '')
    date_s  = meta.get('date', '')
    accent  = _accent(client)

    # Start from a blank presentation — no template XML to fight with
    prs = Presentation()
    prs.slide_width  = W
    prs.slide_height = H

    # ── Slide 1: Cover ────────────────────────────────────────────────────────
    slide_cover(prs, venue, client, contact, role, date_s, accent)

    # ── Slide 2: Hello / cover letter ────────────────────────────────────────
    cover_txt = find_section(sections, 'cover letter', 'cover', 'hello')
    slide_hello(prs, cover_txt, accent)

    # ── Slide 3: Your brief ───────────────────────────────────────────────────
    brief_txt = find_section(sections, 'your brief', 'brief reflection', 'understanding')
    slide_brief(prs, brief_txt, accent)

    # ── Slide 4: Our methodology divider ─────────────────────────────────────
    slide_divider(prs, 'Our methodology', accent)

    # ── Slides 5-8: Stages ───────────────────────────────────────────────────
    stage_map = [
        ('stage 1', 'strategic framework', 'Stage 1 — Strategic framework'),
        ('stage 2', 'concept design',      'Stage 2 — Concept design'),
        ('stage 3', 'design development',  'Stage 3 — Design development'),
        ('stages 4', 'design intent',       'Stages 4, 5 and 6'),
    ]
    for key1, key2, label in stage_map:
        txt = find_section(sections, key1, key2)
        slide_stage(prs, 'Our methodology', label, txt, accent)

    # ── Slide 9: Fees ─────────────────────────────────────────────────────────
    fees_stages = [
        {'title': 'Strategic framework', 'sub': 'Workshop, site visit and proposition',
         'fee': '[FEE: TBC]', 'timing': '1 week', 'invoicing': '100% at start of stage'},
        {'title': 'Concept design',      'sub': 'Layouts, materials, CGI visuals',
         'fee': '[FEE: TBC]', 'timing': '2 weeks', 'invoicing': '50% at start, 50% at completion'},
        {'title': 'Design development',  'sub': 'Sample boards and concept freeze',
         'fee': '[FEE: TBC]', 'timing': '2 weeks', 'invoicing': '50% at start, 50% at completion'},
        {'title': 'Design intent and artwork', 'sub': 'Drawing pack and graphic artwork',
         'fee': '[FEE: TBC]', 'timing': '4 weeks', 'invoicing': '50% at start, 50% at completion'},
    ]
    slide_fees(prs, fees_stages, accent)

    # ── Slide 10: Next steps ──────────────────────────────────────────────────
    next_txt = find_section(sections, 'next steps', 'programme', 'fees')
    slide_next_steps(prs, next_txt, accent)

    prs.save(output_path)
    return output_path
