"""
20.20 Design Agency — Proposal Generator
Hosted Flask app for LawLiss / 20.20
"""

import os, json, uuid, threading, queue, time, base64, re, copy, zipfile, tempfile, shutil
import anthropic
from flask import Flask, request, jsonify, send_file, Response, render_template

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 20 * 1024 * 1024  # 20MB max upload

# ── CONFIG ────────────────────────────────────────────────────────────────────
ANTHROPIC_KEY = os.environ.get('ANTHROPIC_API_KEY', '')
TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), '2020_template_slim_b64.txt')

# In-memory job store (fine for this volume — a few per day)
jobs = {}  # job_id -> {'status', 'progress', 'sections', 'meta', 'pptx_path', 'error'}


# ── NAMESPACES ────────────────────────────────────────────────────────────────
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'

import xml.etree.ElementTree as ET
for prefix, uri in [
    ('p', P), ('a', A),
    ('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'),
    ('a16', 'http://schemas.microsoft.com/office/drawing/2014/main'),
]:
    ET.register_namespace(prefix, uri)


# ── CLIENT COLOURS ────────────────────────────────────────────────────────────
CLUB_COLOURS = {
    'aston villa': '5C1A2E', 'villa': '5C1A2E',
    'newcastle': 'C9A84C', 'nufc': 'C9A84C',
    'brighton': '0057B8', 'bhafc': '0057B8', 'amex': '0057B8',
    'arsenal': 'EF0107',
    'liverpool': 'C8102E',
    'chelsea': '034694',
    'crystal palace': '1B458F', 'cpfc': '1B458F',
    'leeds': 'FFCD00', 'lufc': 'FFCD00',
    'sunderland': 'EB172B', 'safc': 'EB172B',
    'west ham': '7A263A',
    'manchester city': '6CABDD', 'man city': '6CABDD',
    'manchester united': 'DA291C', 'man utd': 'DA291C',
    'tottenham': '132257', 'spurs': '132257',
    'everton': '003399',
    'sheffield': 'EE2737',
    'nottingham forest': 'E53233',
    'leicester': '003090',
    'wolves': 'FDB913',
    'celtic': '00843D',
    'rangers': '0033A0',
}
DEFAULT_COLOUR = 'E97132'  # 20.20 orange

def detect_colour(client_name):
    if not client_name:
        return DEFAULT_COLOUR
    cl = client_name.lower()
    for key, colour in CLUB_COLOURS.items():
        if key in cl:
            return colour
    return DEFAULT_COLOUR


# ── TEXT HELPERS ──────────────────────────────────────────────────────────────
def clean(txt):
    if not txt:
        return ''
    txt = re.sub(r'\*\*([^*]+)\*\*', r'\1', txt)
    txt = re.sub(r'\*([^*]+)\*', r'\1', txt)
    txt = re.sub(r'^#{1,3}\s*', '', txt, flags=re.MULTILINE)
    txt = re.sub(r'\n{3,}', '\n\n', txt)
    return txt.strip()

def first_sentences(txt, n=2):
    s = clean(txt)
    return ' '.join(re.split(r'(?<=[.!?])\s+', s)[:n]).strip()

def explicit_bullets(txt, max_n=8):
    s = clean(txt)
    bullets = []
    for line in s.split('\n'):
        stripped = line.strip()
        if stripped.startswith(('-', '\u2022', '*')):
            item = re.sub(r'^[-\u2022*]\s*', '', stripped).strip()
            if len(item) > 8 and item not in bullets:
                bullets.append(item)
        elif re.match(r'^\d+[.):]', stripped):
            item = re.sub(r'^\d+[.):]+\s*', '', stripped).strip()
            if len(item) > 8 and item not in bullets:
                bullets.append(item)
    if not bullets:
        for line in s.split('\n'):
            line = re.sub(r'^[-\u2022*\d+.]:?\s*', '', line).strip()
            if len(line) > 12 and line not in bullets:
                bullets.append(line)
    return bullets[:max_n]

def prose_only(txt, n=3):
    s = clean(txt)
    lines = s.split('\n')
    prose = []
    for line in lines:
        stripped = line.strip()
        if stripped.startswith(('-', '\u2022', '*')) or re.match(r'^\d+[.):]', stripped):
            break
        if stripped.lower().rstrip(':') in ('deliverables', 'outputs', 'scope', 'process', 'objective'):
            break
        prose.append(stripped)
    text = ' '.join(l for l in prose if l)
    return ' '.join(re.split(r'(?<=[.!?])\s+', text)[:n]).strip()

def find_section(sections, *keys):
    for key in keys:
        kl = key.lower()
        for sec in sections:
            h = sec.get('heading', '').lower()
            if kl in h or h in kl:
                return clean(sec.get('body', ''))
    return ''


# ── XML HELPERS ───────────────────────────────────────────────────────────────
def get_txbodies(root):
    return [e for e in root.iter() if e.tag == f'{{{P}}}txBody']

def full_text(tb):
    return ''.join(e.text for e in tb.iter() if e.tag == f'{{{A}}}t' and e.text)

def get_first_rPr(tb):
    for r in tb.iter(f'{{{A}}}r'):
        rPr = r.find(f'{{{A}}}rPr')
        if rPr is not None:
            return copy.deepcopy(rPr)
    return None

def make_rPr(tmpl=None, bold=False, colour=None):
    rPr = ET.Element(f'{{{A}}}rPr')
    rPr.set('lang', 'en-GB')
    rPr.set('dirty', '0')
    if tmpl is not None:
        for attr in ['sz', 'lang']:
            if tmpl.get(attr):
                rPr.set(attr, tmpl.get(attr))
        if not colour:
            for child in tmpl:
                if any(k in child.tag for k in ('Fill', 'latin', 'ea', 'cs')):
                    rPr.append(copy.deepcopy(child))
        else:
            for child in tmpl:
                if any(k in child.tag for k in ('latin', 'ea', 'cs')) and 'Fill' not in child.tag:
                    rPr.append(copy.deepcopy(child))
    if colour:
        sf = ET.SubElement(rPr, f'{{{A}}}solidFill')
        sc = ET.SubElement(sf, f'{{{A}}}srgbClr')
        sc.set('val', colour.upper().lstrip('#'))
    if bold:
        rPr.set('b', '1')
    return rPr

def set_text(tb, text, tmpl_rPr=None, bold=False, colour=None):
    for p in [e for e in tb if e.tag == f'{{{A}}}p']:
        tb.remove(p)
    p = ET.SubElement(tb, f'{{{A}}}p')
    r = ET.SubElement(p, f'{{{A}}}r')
    r.append(make_rPr(tmpl_rPr, bold=bold, colour=colour))
    ET.SubElement(r, f'{{{A}}}t').text = text

def set_paragraphs(tb, items, tmpl_rPr=None):
    for p in [e for e in tb if e.tag == f'{{{A}}}p']:
        tb.remove(p)
    for text, opts in items:
        p = ET.SubElement(tb, f'{{{A}}}p')
        pPr = ET.SubElement(p, f'{{{A}}}pPr')
        pPr.set('lvl', '0')
        if opts.get('bullet'):
            ET.SubElement(pPr, f'{{{A}}}buFont').set('typeface', 'Arial')
            ET.SubElement(pPr, f'{{{A}}}buChar').set('char', '\u2022')
        else:
            ET.SubElement(pPr, f'{{{A}}}buNone')
        if text:
            r = ET.SubElement(p, f'{{{A}}}r')
            r.append(make_rPr(tmpl_rPr, bold=opts.get('bold', False),
                               colour=opts.get('colour')))
            ET.SubElement(r, f'{{{A}}}t').text = text
        else:
            ET.SubElement(p, f'{{{A}}}endParaRPr').set('lang', 'en-GB')

def replace_colour(xml_str, old, new):
    old, new = old.upper(), new.upper()
    for v in [old, old.lower()]:
        xml_str = xml_str.replace(f'val="{v}"', f'val="{new}"')
    return xml_str


# ── SLIDE BUILDERS ────────────────────────────────────────────────────────────
def build_cover(root, venue, contact, role, date_s):
    r = copy.deepcopy(root)
    for tb in get_txbodies(r):
        ft = full_text(tb)
        rPr = get_first_rPr(tb)
        if 'Concept Presentation' in ft:
            set_text(tb, venue, rPr)
        elif 'Stage 2' in ft and len(ft) < 20:
            set_text(tb, 'Hospitality design proposal', rPr)
        elif any(x in ft for x in ["June", "25th", "'25"]) and date_s:
            set_text(tb, date_s, rPr)
    return r

def build_hello(root):
    r = copy.deepcopy(root)
    for tb in get_txbodies(r):
        if full_text(tb).strip() == 'PowerPoint Template':
            set_text(tb, '', get_first_rPr(tb))
    return r

def build_dark_divider(root, word):
    r = copy.deepcopy(root)
    for tb in get_txbodies(r):
        ft = full_text(tb).strip()
        rPr = get_first_rPr(tb)
        if ft == 'Hello':
            set_text(tb, word, rPr)
        elif ft == 'PowerPoint Template':
            set_text(tb, '', rPr)
    return r

def build_content_slide(root, section_label, title, intro, bullets):
    r = copy.deepcopy(root)
    for tb in get_txbodies(r):
        ft = full_text(tb)
        rPr = get_first_rPr(tb)
        if 'PowerPoint Template' in ft and len(ft) < 40:
            set_text(tb, section_label, rPr)
        elif 'Example header' in ft:
            set_text(tb, title, rPr)
        elif 'Lorem ipsum' in ft or ('lorem' in ft.lower() and len(ft) > 30):
            items = []
            if intro:
                items.append((intro, {}))
            if bullets:
                items.append(('', {}))
                for b in bullets:
                    items.append((b, {'bullet': True}))
            if items:
                set_paragraphs(tb, items, rPr)
    return r

def build_stage_slide(slide14_raw, section_label, stage_title, body, deliverables, accent):
    root = ET.fromstring(slide14_raw)
    tbs = get_txbodies(root)
    for i, tb in enumerate(tbs):
        rPr = get_first_rPr(tb)
        if i == 0:
            set_text(tb, section_label, rPr)
        elif i == 2:
            set_text(tb, stage_title, rPr)
        elif i == 3:
            prose = prose_only(body, 3)
            if prose:
                set_paragraphs(tb, [(prose, {})], rPr)
        elif i == 4:
            set_text(tb, 'Scope', rPr, colour=accent)
        elif i == 5:
            dl = [d for d in deliverables if d] if deliverables else []
            if not dl:
                dl = explicit_bullets(body, 8)
            items = [(d, {'bullet': True}) for d in dl[:8]]
            if items:
                set_paragraphs(tb, items, rPr)
        elif i == 7:
            set_text(tb, '[FEE: TBC — rate card required]', rPr, bold=True)
        elif i == 8:
            set_text(tb, 'Deliverables', rPr, colour=accent)
    return ET.tostring(root, encoding='unicode', xml_declaration=True)

def build_fees_slide(root, stages, accent):
    r = copy.deepcopy(root)
    stage_keys = ['WORKSHOP & DEFINITION', 'CONCEPT DESIGN', 'DESIGN DEVELOPMENT', 'DESIGN INTENT']
    si = 0
    for tb in get_txbodies(r):
        ft = full_text(tb)
        rPr = get_first_rPr(tb)
        if 'PowerPoint Template' in ft and len(ft) < 40:
            set_text(tb, 'Our methodology', rPr)
        elif 'Summary fees and timings' in ft:
            set_text(tb, 'Summary fees and timings', rPr)
        elif any(k in ft for k in stage_keys) and si < len(stages):
            sd = stages[si]
            items = [(sd['title'], {'bold': True})]
            if sd.get('sub'):
                items.append((sd['sub'], {}))
            set_paragraphs(tb, items, rPr)
            si += 1
        elif 'TOTAL 20.20' in ft or 'AFEES' in ft:
            set_paragraphs(tb, [
                ('TOTAL 20.20 FEES', {'bold': True}),
                ('Planning and Interiors', {}),
                ('Identity and Graphics', {}),
                ('Strategy and Management', {}),
            ], rPr)
        elif ft.startswith('£') and 'Bobby' not in ft and len(ft) < 30:
            set_text(tb, '[FEE: TBC]', rPr, bold=True)
    return r


# ── PPTX BUILDER ─────────────────────────────────────────────────────────────
def build_pptx(sections, meta):
    """Build a PPTX from proposal sections and metadata. Returns path to temp file."""
    accent = detect_colour(meta.get('client', ''))
    venue = meta.get('venue', 'Project')
    contact = meta.get('contact', '')
    role = meta.get('role', '')
    date_s = meta.get('date', '')

    if not os.path.exists(TEMPLATE_PATH):
        raise FileNotFoundError('Template not found — ensure 2020_template_slim_b64.txt is deployed.')

    with open(TEMPLATE_PATH, 'rb') as f:
        template_bytes = base64.b64decode(f.read())

    tmpdir = tempfile.mkdtemp(prefix='2020_')
    tpl_pptx = os.path.join(tmpdir, 'template.pptx')
    unpacked = os.path.join(tmpdir, 'unpacked')
    os.makedirs(unpacked)

    with open(tpl_pptx, 'wb') as f:
        f.write(template_bytes)

    with zipfile.ZipFile(tpl_pptx, 'r') as z:
        z.extractall(unpacked)
        slide14_raw = z.read('ppt/slides/slide14.xml').decode('utf-8')
        slide26_raw = z.read('ppt/slides/slide26.xml').decode('utf-8')

    # Replace accent2 in all themes
    themes_dir = os.path.join(unpacked, 'ppt', 'theme')
    if os.path.exists(themes_dir):
        for tf in os.listdir(themes_dir):
            if tf.endswith('.xml'):
                fpath = os.path.join(themes_dir, tf)
                with open(fpath) as f:
                    tx = f.read()
                tx = re.sub(
                    r'<a:accent2><a:srgbClr val="[0-9A-Fa-f]{6}"/></a:accent2>',
                    f'<a:accent2><a:srgbClr val="{accent}"/></a:accent2>',
                    tx
                )
                tx = replace_colour(tx, 'E97132', accent)
                with open(fpath, 'w') as f:
                    f.write(tx)

    def load_slide(n):
        with open(os.path.join(unpacked, 'ppt', 'slides', f'slide{n}.xml')) as f:
            return ET.fromstring(f.read())

    def save_slide(n, root_elem):
        xml = ET.tostring(root_elem, encoding='unicode', xml_declaration=True)
        xml = replace_colour(xml, 'E8251A', accent)
        with open(os.path.join(unpacked, 'ppt', 'slides', f'slide{n}.xml'), 'w', encoding='utf-8') as f:
            f.write(xml)

    def save_slide_str(n, xml_str):
        xml_str = replace_colour(xml_str, 'E8251A', accent)
        with open(os.path.join(unpacked, 'ppt', 'slides', f'slide{n}.xml'), 'w', encoding='utf-8') as f:
            f.write(xml_str)

    # Build slides
    save_slide(1, build_cover(load_slide(1), venue, contact, role, date_s))
    save_slide(27, build_hello(load_slide(27)))

    brief = find_section(sections, 'your brief', 'brief reflection', 'understanding')
    save_slide(8, build_content_slide(
        load_slide(8), 'Your brief', 'Our understanding',
        first_sentences(brief, 2), explicit_bullets(brief, 5) or []
    ))

    save_slide(28, build_dark_divider(load_slide(28), 'Our methodology'))

    for slide_n, key1, key2, label in [
        (14, 'stage 1', 'strategic framework', 'Stage 1 — Strategic framework'),
        (9,  'stage 2', 'concept design',      'Stage 2 — Concept design'),
        (10, 'stage 3', 'design development',  'Stage 3 — Design development'),
        (11, 'stages 4', 'design intent',       'Stages 4, 5 and 6'),
    ]:
        txt = find_section(sections, key1, key2)
        save_slide_str(slide_n, build_stage_slide(
            slide14_raw, 'Our methodology', label, txt, explicit_bullets(txt, 8), accent
        ))

    fees_stages = [
        {'title': 'STRATEGIC FRAMEWORK', 'sub': 'Workshop, site visit and proposition'},
        {'title': 'CONCEPT DESIGN',      'sub': 'Layouts, materials, CGI visuals'},
        {'title': 'DESIGN DEVELOPMENT',  'sub': 'Sample boards and concept freeze'},
        {'title': 'DESIGN INTENT & ARTWORK', 'sub': 'Drawing pack and graphic artwork'},
    ]
    save_slide(16, build_fees_slide(load_slide(16), fees_stages, accent))
    save_slide_str(26, replace_colour(slide26_raw, 'E8251A', accent))

    # Pack into PPTX
    output_path = os.path.join(tmpdir, 'output.pptx')
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for root_dir, dirs, files in os.walk(unpacked):
            for file in files:
                fp = os.path.join(root_dir, file)
                zout.write(fp, os.path.relpath(fp, unpacked))

    return output_path, tmpdir


# ── AI PIPELINE ───────────────────────────────────────────────────────────────
SYSTEM_PROMPT = """You are a proposal writer for 20.20 Design Agency, a specialist hospitality and stadium interior design consultancy. You write first-draft proposals that the account team will review and refine.

VOICE: Confident, direct, commercially aware, personal. Short sentences. Active voice. No em dashes. No AI phrases (leveraging, seamless, holistic, transformative). The client name appears only in the cover letter — all other sections say "the club", "the venue", "the project".

DESIGN PRINCIPLES: Hospitality Pyramid (tier each space), Narrative Before Design (names and stories before materials), Guest Journey Mapping, Brand Integration Without Decoration, Non-Matchday Flexibility, Graphic Identity as Interior Design, Commercial Consciousness, CGI from Stage 2, Collaborative Design Team, Concept Freeze.

PLACEHOLDERS: [FEE: TBC — rate card required] for fees, [IMAGE REQUIRED: description] for visuals, [CONFIRM WITH CLIENT: note] for assumptions."""

SECTIONS = [
    ('cover',   'Cover letter',                  'Write the cover letter. Address {contact} by first name. Client is {client}. Personal, direct, 3-4 short paragraphs. This is the ONLY section where the client name appears.\n\n{ctx}'),
    ('brief',   'Your brief',                    'Write the brief reflection titled "Your brief". Show understanding of the commercial context. Use "the club" or "the venue" — not the client name. 2-3 paragraphs.\n\n{ctx}'),
    ('approach','Our approach',                  'Write a short "Our approach" intro (4-6 sentences). Methodology overview — RIBA-staged, commercially conscious, narrative-led. No client name.\n\n{ctx}'),
    ('stage1',  'Stage 1 — Strategic framework', 'Write Stage 1 (Strategic framework / RIBA 2 / 1 week). Objective paragraph, process paragraph, then bullet deliverables. No client name.\n\n{ctx}'),
    ('stage2',  'Stage 2 — Concept design',      'Write Stage 2 (Concept design / RIBA 2 / 2 weeks). Objective, process, bullet deliverables including CGI commitment (min 2 per space). No client name.\n\n{ctx}'),
    ('stage3',  'Stage 3 — Design development',  'Write Stage 3 (Design development / RIBA 2 / 2 weeks). Objective, process, bullet deliverables including concept freeze. No client name.\n\n{ctx}'),
    ('stage456','Stages 4, 5 and 6',             'Write combined section for Stage 4 (Design intent/RIBA 3/4 wks), Stage 5 (Coordination/RIBA 4-5), Stage 6 (Handover/RIBA 6). One paragraph + bullets per stage. No client name.\n\n{ctx}'),
    ('fees',    'Fees and timings',               'Write the fees section. Six stages, [FEE: TBC] for all figures. Three disciplines per stage. Note VAT and expenses exclusions.\n\n{ctx}'),
    ('programme','Estimated programme',           'Write a programme narrative (1 paragraph) covering the season window and key milestones. No client name.\n\n{ctx}'),
    ('nextsteps','Next steps',                    'Write next steps. Four actions: review, feedback, site visit, appointment. Direct and confident.\n\n{ctx}'),
]

def build_context(meta, spaces_text=''):
    return (
        f"PROJECT: {meta.get('venue','')}\n"
        f"CONTACT: {meta.get('contact','')}{', '+meta.get('role','') if meta.get('role') else ''}\n"
        f"BRIEF TYPE: {meta.get('brief_type','')}\n"
        f"RIBA STAGES: {meta.get('riba_stages','TBC')}\n"
        f"BUDGET: {meta.get('budget','Not stated')}\n"
        f"SPACES: {spaces_text or 'See scope summary'}\n"
        f"SCOPE: {meta.get('scope','')}"
    )

def run_pipeline(job_id, pdf_b64=None, brief_text=None):
    """Background thread: extract → research → generate → build PPTX."""
    job = jobs[job_id]
    client = anthropic.Anthropic(api_key=ANTHROPIC_KEY)

    def progress(msg, pct=None):
        job['progress'].append({'msg': msg, 'pct': pct})

    try:
        # ── STEP 1: EXTRACT ─────────────────────────────────────────────────
        progress('Reading the brief...', 5)
        extract_prompt = (
            'Read this client brief. Extract all structured data and return ONLY valid JSON:\n'
            '{"brief_type":"newbuild|refurb|single|sponsor|arena",'
            '"client":"","venue":"","primary_contact":"","contact_role":"",'
            '"proposal_deadline":"","completion_target":"","budget_stated":"",'
            '"riba_stages":"","spaces":[{"name":"","tier":"","size":"","budget":""}],'
            '"payback_target":"","brief_source":"Direct approach from client|Via lead architect or PM|Formal open tender (ITT)|Referral|Repeat client / existing relationship|Unknown",'
            '"scope_summary":"2-3 sentences"}'
        )

        if pdf_b64:
            msg_content = [
                {'type': 'document', 'source': {'type': 'base64', 'media_type': 'application/pdf', 'data': pdf_b64}},
                {'type': 'text', 'text': extract_prompt},
            ]
        else:
            msg_content = extract_prompt + '\n\nBrief:\n' + (brief_text or '')[:4000]

        resp = client.messages.create(
            model='claude-sonnet-4-20250514',
            max_tokens=1400,
            messages=[{'role': 'user', 'content': msg_content}]
        )
        raw = resp.content[0].text.replace('```json', '').replace('```', '').strip()
        m = re.search(r'\{[\s\S]*\}', raw)
        if not m:
            raise ValueError('Could not extract brief data from the document.')
        ex = json.loads(m.group(0))
        job['extracted'] = ex

        meta = {
            'client':      ex.get('client', ''),
            'venue':       ex.get('venue', ''),
            'contact':     ex.get('primary_contact', ''),
            'role':        ex.get('contact_role', ''),
            'brief_type':  ex.get('brief_type', ''),
            'riba_stages': ex.get('riba_stages', ''),
            'budget':      ex.get('budget_stated', ''),
            'scope':       ex.get('scope_summary', ''),
            'date':        time.strftime('%-d %B %Y'),
        }
        job['meta'] = meta
        spaces_text = '\n'.join(
            f"- {s['name']}" + (f" ({s['tier']})" if s.get('tier') else '') +
            (f", {s['size']}" if s.get('size') else '') +
            (f", {s['budget']}" if s.get('budget') else '')
            for s in ex.get('spaces', [])
        ) or 'Not listed'

        progress(f'Brief read — {meta["client"] or "client"} / {meta["venue"] or "project"}', 10)

        # ── STEP 2: RESEARCH ─────────────────────────────────────────────────
        progress('Researching the client...', 15)
        time.sleep(12)  # Let rate limit recover after extraction

        research_prompt = (
            f'Research {meta["contact"]} at {meta["client"] or meta["venue"]} for a design agency pitch. '
            'Return ONLY valid JSON:\n'
            '{"contact_profile":"2-3 sentences","org_context":"current position","'
            'why_now":"why this brief exists","ambitions":"strategic goals","confidence":"high|medium|low"}'
        )
        try:
            resp2 = client.messages.create(
                model='claude-sonnet-4-20250514',
                max_tokens=800,
                tools=[{'type': 'web_search_20250305', 'name': 'web_search'}],
                messages=[{'role': 'user', 'content': research_prompt}]
            )
            txt2 = ' '.join(b.text for b in resp2.content if hasattr(b, 'text'))
            m2 = re.search(r'\{[\s\S]*\}', txt2)
            job['intel'] = json.loads(m2.group(0)) if m2 else {}
        except Exception:
            job['intel'] = {}

        progress('Client research complete', 20)

        # ── STEP 3: GENERATE SECTIONS ────────────────────────────────────────
        ctx = build_context(meta, spaces_text)
        sections = []
        total = len(SECTIONS)
        GAP = 7  # seconds between API calls

        for i, (sid, label, prompt_tpl) in enumerate(SECTIONS):
            pct = 20 + int((i / total) * 65)
            progress(f'Writing: {label} ({i+1} of {total})', pct)

            prompt = prompt_tpl.format(
                contact=meta.get('contact', 'the contact'),
                client=meta.get('client', 'the client'),
                ctx=ctx
            )

            for attempt in range(3):
                try:
                    if attempt > 0:
                        wait = 35 if attempt == 1 else 55
                        progress(f'Rate limit — retrying {label} in {wait}s...', pct)
                        time.sleep(wait)

                    resp3 = client.messages.create(
                        model='claude-sonnet-4-20250514',
                        max_tokens=800,
                        system=SYSTEM_PROMPT,
                        messages=[{'role': 'user', 'content': prompt}]
                    )
                    sections.append({
                        'id': sid,
                        'heading': label,
                        'body': resp3.content[0].text.strip()
                    })
                    break
                except anthropic.RateLimitError:
                    if attempt == 2:
                        sections.append({'id': sid, 'heading': label, 'body': f'[Could not generate — add manually]'})
                except Exception as e:
                    sections.append({'id': sid, 'heading': label, 'body': f'[Error: {str(e)[:80]}]'})
                    break

            if i < total - 1:
                time.sleep(GAP)

        job['sections'] = sections
        progress('All sections written', 85)

        # ── STEP 4: BUILD PPTX ───────────────────────────────────────────────
        progress('Building PowerPoint from template...', 88)
        pptx_path, tmpdir = build_pptx(sections, meta)
        job['pptx_path'] = pptx_path
        job['tmpdir'] = tmpdir
        job['status'] = 'done'
        progress('Done', 100)

    except Exception as e:
        job['status'] = 'error'
        job['error'] = str(e)
        progress(f'Error: {e}', None)


# ── ROUTES ────────────────────────────────────────────────────────────────────
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/submit', methods=['POST'])
def submit():
    if not ANTHROPIC_KEY:
        return jsonify({'error': 'API key not configured on server.'}), 500

    job_id = str(uuid.uuid4())[:8]
    jobs[job_id] = {
        'status': 'running',
        'progress': [],
        'sections': [],
        'meta': {},
        'intel': {},
        'extracted': {},
        'pptx_path': None,
        'tmpdir': None,
        'error': None,
    }

    pdf_b64 = None
    brief_text = None

    if 'brief_pdf' in request.files and request.files['brief_pdf'].filename:
        f = request.files['brief_pdf']
        pdf_b64 = base64.b64encode(f.read()).decode('ascii')
    elif request.form.get('brief_text'):
        brief_text = request.form.get('brief_text')
    else:
        return jsonify({'error': 'Please upload a PDF or paste the brief text.'}), 400

    t = threading.Thread(target=run_pipeline, args=(job_id, pdf_b64, brief_text), daemon=True)
    t.start()

    return jsonify({'job_id': job_id})

@app.route('/status/<job_id>')
def status(job_id):
    job = jobs.get(job_id)
    if not job:
        return jsonify({'error': 'Job not found'}), 404
    return jsonify({
        'status':   job['status'],
        'progress': job['progress'],
        'sections': job['sections'],
        'meta':     job['meta'],
        'intel':    job['intel'],
        'error':    job['error'],
    })

@app.route('/rebuild', methods=['POST'])
def rebuild():
    """Rebuild PPTX from edited sections (user reviewed and changed text)."""
    data = request.get_json()
    job_id = data.get('job_id')
    sections = data.get('sections', [])
    meta = data.get('meta', {})

    if not sections:
        return jsonify({'error': 'No sections provided'}), 400

    try:
        pptx_path, tmpdir = build_pptx(sections, meta)
        # Store under same job_id
        if job_id and job_id in jobs:
            # Clean up old tmpdir
            old_tmp = jobs[job_id].get('tmpdir')
            if old_tmp and old_tmp != tmpdir:
                shutil.rmtree(old_tmp, ignore_errors=True)
            jobs[job_id]['pptx_path'] = pptx_path
            jobs[job_id]['tmpdir'] = tmpdir
        return jsonify({'status': 'ok', 'job_id': job_id})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/download/<job_id>')
def download(job_id):
    job = jobs.get(job_id)
    if not job or not job.get('pptx_path'):
        return 'Not found', 404

    venue = job.get('meta', {}).get('venue', 'Proposal').replace(' ', '_').replace("'", '')
    filename = f'{venue}_20.20_Proposal.pptx'

    return send_file(
        job['pptx_path'],
        as_attachment=True,
        download_name=filename,
        mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
    )

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
