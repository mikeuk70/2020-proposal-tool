"""
20.20 Design Agency — Proposal Generator
Hosted Flask app for LawLiss / 20.20
"""

import os, json, uuid, threading, queue, time, base64, re, copy, zipfile, tempfile, shutil
import anthropic
from flask import Flask, request, jsonify, send_file, Response

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 20 * 1024 * 1024  # 20MB max upload

# ── CONFIG ────────────────────────────────────────────────────────────────────
ANTHROPIC_KEY = os.environ.get('ANTHROPIC_API_KEY', '')
# Find template file - check several locations
_here = os.path.dirname(os.path.abspath(__file__))
_candidates = [
    os.path.join(_here, '2020_template_slim_b64.txt'),
    '/app/2020_template_slim_b64.txt',
    os.path.join(os.getcwd(), '2020_template_slim_b64.txt'),
]
TEMPLATE_PATH = next((p for p in _candidates if os.path.exists(p)), _candidates[0])

# File-based job store — survives restarts and works across gunicorn workers
JOBS_DIR = os.path.join(tempfile.gettempdir(), '2020_jobs')
os.makedirs(JOBS_DIR, exist_ok=True)

def job_path(job_id):
    return os.path.join(JOBS_DIR, f'{job_id}.json')

def pptx_path_for(job_id):
    return os.path.join(JOBS_DIR, f'{job_id}.pptx')

def load_job(job_id):
    p = job_path(job_id)
    if not os.path.exists(p):
        return None
    try:
        with open(p, 'r') as f:
            return json.load(f)
    except Exception:
        return None

def save_job(job_id, job):
    p = job_path(job_id)
    try:
        with open(p, 'w') as f:
            json.dump(job, f)
    except Exception:
        pass

def update_job(job_id, **kwargs):
    job = load_job(job_id) or {}
    job.update(kwargs)
    save_job(job_id, job)

def append_progress(job_id, msg, pct=None):
    job = load_job(job_id) or {}
    prog = job.get('progress', [])
    prog.append({'msg': msg, 'pct': pct})
    job['progress'] = prog
    save_job(job_id, job)

def append_section(job_id, section):
    job = load_job(job_id) or {}
    secs = job.get('sections', [])
    secs.append(section)
    job['sections'] = secs
    save_job(job_id, job)


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

    # Pack into PPTX — [Content_Types].xml must be first, then _rels/.rels
    # Use the original template as base and patch modified slides in
    output_path = os.path.join(tmpdir, 'output.pptx')

    # Build a map of modified files
    modified = {}
    for root_dir, dirs, files in os.walk(unpacked):
        for file in files:
            fp = os.path.join(root_dir, file)
            arc = os.path.relpath(fp, unpacked).replace(os.sep, '/')
            with open(fp, 'rb') as f:
                modified[arc] = f.read()

    # Write zip with correct ordering: content types first, then rels, then everything else
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        # 1. Content types must be first
        if '[Content_Types].xml' in modified:
            zout.writestr('[Content_Types].xml', modified['[Content_Types].xml'])
        # 2. Root rels
        if '_rels/.rels' in modified:
            zout.writestr('_rels/.rels', modified['_rels/.rels'])
        # 3. Everything else
        for arc, data in modified.items():
            if arc not in ('[Content_Types].xml', '_rels/.rels'):
                zout.writestr(arc, data)

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
    client = anthropic.Anthropic(api_key=ANTHROPIC_KEY)

    def progress(msg, pct=None):
        append_progress(job_id, msg, pct)

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
        update_job(job_id, extracted=ex)

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
        update_job(job_id, meta=meta)
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
            update_job(job_id, intel=json.loads(m2.group(0)) if m2 else {})
        except Exception:
            update_job(job_id, intel={})

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
                    sec = {'id': sid, 'heading': label, 'body': resp3.content[0].text.strip()}
                    sections.append(sec)
                    append_section(job_id, sec)
                    break
                except anthropic.RateLimitError:
                    if attempt == 2:
                        sec = {'id': sid, 'heading': label, 'body': '[Could not generate — add manually]'}
                    sections.append(sec)
                    append_section(job_id, sec)
                except Exception as e:
                    sec = {'id': sid, 'heading': label, 'body': f'[Error: {str(e)[:80]}]'}
                    sections.append(sec)
                    append_section(job_id, sec)
                    break

            if i < total - 1:
                time.sleep(GAP)

        job['sections'] = sections
        progress('All sections written', 85)

        # ── STEP 4: BUILD PPTX ───────────────────────────────────────────────
        progress('Building PowerPoint from template...', 88)
        try:
            pptx_path, tmpdir = build_pptx(sections, meta)
            if not os.path.exists(pptx_path):
                raise FileNotFoundError('Output PPTX was not created')
            update_job(job_id, pptx_path=pptx_path, status='done')
            progress('Done — click Download PowerPoint', 100)
        except Exception as pptx_err:
            import traceback
            err_detail = traceback.format_exc()
            update_job(job_id, status='done', pptx_error=str(pptx_err), pptx_traceback=err_detail)
            progress(f'Sections complete. PowerPoint failed: {pptx_err}', 100)

    except Exception as e:
        import traceback
        # Even on pipeline error, mark done if we have sections
        job = load_job(job_id) or {}
        if job.get('sections'):
            update_job(job_id, status='done', error=str(e))
        else:
            update_job(job_id, status='error', error=str(e))
        progress(f'Error: {e}', None)


# ── ROUTES ────────────────────────────────────────────────────────────────────
INDEX_HTML = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>20.20 Proposal Generator</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
:root{
  --nv:#1B2340;--gd:#C9A84C;--rd:#E97132;
  --bg:#F5F4F1;--white:#fff;--bd:#E0DED8;
  --tx:#1A1A1A;--tx2:#666;--r:8px;--rl:14px
}
body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Arial,sans-serif;
  background:var(--bg);color:var(--tx);min-height:100vh}

/* NAV */
nav{background:var(--nv);padding:0 2rem;display:flex;align-items:center;
  justify-content:space-between;height:54px;position:sticky;top:0;z-index:100}
.logo{display:flex;align-items:center;gap:10px}
.logo-mark{font-size:13px;font-weight:700;color:var(--white);line-height:1.1;letter-spacing:-0.5px}
.logo-mark span{color:var(--gd)}
.logo-name{font-size:13px;color:rgba(255,255,255,.55);border-left:1px solid rgba(255,255,255,.2);padding-left:10px}
.nav-status{font-size:12px;color:rgba(255,255,255,.45)}

/* LAYOUT */
.page{max-width:860px;margin:0 auto;padding:2rem 1.5rem 4rem}

/* PANELS */
.panel{background:var(--white);border-radius:var(--rl);margin-bottom:1.25rem;overflow:hidden;
  box-shadow:0 1px 4px rgba(0,0,0,.06)}
.panel-head{padding:1rem 1.25rem;border-bottom:1px solid var(--bd);display:flex;
  align-items:center;justify-content:space-between}
.panel-head h2{font-size:15px;font-weight:600;color:var(--nv)}
.panel-head .step-badge{font-size:10px;font-weight:700;text-transform:uppercase;
  letter-spacing:.08em;color:var(--tx2);background:var(--bg);
  padding:3px 10px;border-radius:20px}
.panel-body{padding:1.25rem}

/* FORM */
.tab-row{display:flex;border:1px solid var(--bd);border-radius:var(--r);overflow:hidden;margin-bottom:1rem}
.tab-btn{flex:1;padding:8px;font-size:13px;font-weight:600;border:none;cursor:pointer;
  font-family:inherit;transition:all .15s}
.tab-btn.active{background:var(--nv);color:var(--white)}
.tab-btn.inactive{background:var(--bg);color:var(--tx2);border-left:1px solid var(--bd)}
.field-label{display:block;font-size:12px;font-weight:600;margin-bottom:5px;color:var(--tx2)}
input[type=file]{display:block;width:100%;font-size:13px;padding:8px;
  border:1px solid var(--bd);border-radius:var(--r);background:var(--bg);cursor:pointer;font-family:inherit}
textarea{width:100%;font-size:13px;padding:10px;border:1px solid var(--bd);
  border-radius:var(--r);font-family:inherit;resize:vertical;min-height:160px;
  line-height:1.5;outline:none;transition:border-color .15s}
textarea:focus{border-color:var(--nv)}

/* BUTTONS */
.btn{display:inline-flex;align-items:center;gap:6px;padding:9px 20px;border:none;
  border-radius:var(--r);font-size:13px;font-weight:600;cursor:pointer;
  font-family:inherit;transition:opacity .15s}
.btn:hover{opacity:.88}
.btn-primary{background:var(--nv);color:var(--white)}
.btn-gold{background:var(--gd);color:var(--nv)}
.btn-outline{background:transparent;border:1px solid var(--bd);color:var(--tx)}
.btn:disabled{opacity:.4;cursor:not-allowed}

/* PROGRESS */
.progress-wrap{margin:1rem 0}
.progress-bar-bg{height:5px;background:var(--bd);border-radius:3px;overflow:hidden;margin-bottom:.5rem}
.progress-bar-fill{height:100%;background:var(--nv);border-radius:3px;transition:width .5s}
.progress-msg{font-size:12px;color:var(--tx2);text-align:center;min-height:1.2em}

/* SECTIONS */
.section-card{border:1px solid var(--bd);border-radius:var(--r);overflow:hidden;margin-bottom:10px}
.section-head{display:flex;align-items:center;justify-content:space-between;
  padding:.6rem 1rem;background:var(--nv);cursor:pointer;user-select:none}
.section-head-title{font-size:13px;font-weight:600;color:var(--white)}
.section-copy-btn{font-size:11px;color:rgba(255,255,255,.55);padding:2px 8px;
  background:rgba(255,255,255,.1);border:none;border-radius:4px;cursor:pointer;font-family:inherit}
.section-copy-btn:hover{background:rgba(255,255,255,.2);color:var(--white)}
.section-body{padding:.75rem 1rem;background:var(--white)}
.section-body textarea{min-height:80px;border:none;padding:0;background:transparent;
  font-size:13px;line-height:1.7;resize:vertical;outline:none;color:var(--tx)}

/* META */
.meta-grid{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:1rem}
.meta-card{background:var(--bg);border:1px solid var(--bd);border-radius:var(--r);padding:.75rem 1rem}
.meta-label{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;
  color:var(--tx2);margin-bottom:3px}
.meta-value{font-size:13px;color:var(--tx)}
.intel-card{background:var(--bg);border:1px solid var(--bd);border-radius:var(--r);padding:1rem}
.intel-row{padding:.5rem 0;border-bottom:1px solid var(--bd);font-size:13px;line-height:1.5}
.intel-row:last-child{border:none}
.intel-lbl{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;
  color:var(--gd);margin-bottom:2px}

/* ACTIONS BAR */
.actions-bar{display:flex;gap:8px;flex-wrap:wrap;padding-top:1rem;
  border-top:1px solid var(--bd);margin-top:1rem}

/* HIDDEN */
.hidden{display:none}

/* PILL */
.pill{display:inline-block;font-size:10px;font-weight:700;text-transform:uppercase;
  letter-spacing:.06em;padding:2px 10px;border-radius:20px}
.pill-running{background:#EAF3DE;color:#3B6D11}
.pill-done{background:#E6F1FB;color:#185FA5}
.pill-error{background:#FCEBEB;color:#A32D2D}

/* ERROR */
.error-box{background:#FCEBEB;border:1px solid #F09595;border-radius:var(--r);
  padding:.75rem 1rem;font-size:13px;color:#A32D2D;margin-top:.75rem}

@media(max-width:600px){
  .meta-grid{grid-template-columns:1fr}
  .page{padding:1rem .75rem 3rem}
}
</style>
</head>
<body>

<nav>
  <div class="logo">
    <div class="logo-mark">20<br><span>20</span></div>
    <div class="logo-name">Proposal Generator</div>
  </div>
  <div class="nav-status" id="nav-status"></div>
</nav>

<div class="page">

  <!-- STEP 1: INPUT -->
  <div class="panel" id="panel-input">
    <div class="panel-head">
      <h2>Add the brief</h2>
      <span class="step-badge">Step 1</span>
    </div>
    <div class="panel-body">
      <p style="font-size:13px;color:var(--tx2);margin-bottom:1rem;line-height:1.5">
        Upload a PDF brief or paste the text. The tool reads it, researches the client,
        writes the proposal sections, and builds a branded PowerPoint — ready to review and download.
      </p>

      <div class="tab-row">
        <button class="tab-btn active" id="tab-pdf" onclick="switchTab('pdf')">↑ Upload PDF</button>
        <button class="tab-btn inactive" id="tab-text" onclick="switchTab('text')">Paste text</button>
      </div>

      <div id="panel-pdf">
        <label class="field-label">Select PDF brief</label>
        <input type="file" id="brief-pdf" accept=".pdf">
        <p style="font-size:11px;color:var(--tx2);margin-top:.4rem">
          Presentations, ITTs, emails saved as PDF — anything works
        </p>
      </div>

      <div id="panel-text" class="hidden">
        <label class="field-label">Paste brief text</label>
        <textarea id="brief-text" placeholder="Paste the brief here — email, copied PDF text, ITT, meeting notes..."></textarea>
      </div>

      <div id="submit-error" class="error-box hidden"></div>

      <div style="margin-top:1rem">
        <button class="btn btn-primary" id="submit-btn" onclick="submitBrief()">
          Generate proposal →
        </button>
      </div>
    </div>
  </div>

  <!-- STEP 2: PROGRESS -->
  <div class="panel hidden" id="panel-progress">
    <div class="panel-head">
      <h2>Generating</h2>
      <span class="pill pill-running" id="status-pill">Running</span>
    </div>
    <div class="panel-body">
      <div class="progress-wrap">
        <div class="progress-bar-bg"><div class="progress-bar-fill" id="prog-bar" style="width:0%"></div></div>
        <div class="progress-msg" id="prog-msg">Starting...</div>
      </div>
      <p style="font-size:12px;color:var(--tx2);line-height:1.5">
        Writing each section with a short pause between calls to stay within API rate limits.
        This takes around 90 seconds. Sections appear below as they complete.
      </p>
    </div>
  </div>

  <!-- STEP 3: SECTIONS (appear during generation) -->
  <div class="panel hidden" id="panel-sections">
    <div class="panel-head">
      <h2>Proposal sections</h2>
      <span class="step-badge">Review and edit</span>
    </div>
    <div class="panel-body">
      <p style="font-size:12px;color:var(--tx2);margin-bottom:1rem">
        Each section is editable. Make any changes before downloading the PowerPoint.
      </p>
      <div id="sections-list"></div>
    </div>
  </div>

  <!-- STEP 4: CLIENT INTEL -->
  <div class="panel hidden" id="panel-intel">
    <div class="panel-head">
      <h2>Client intelligence</h2>
      <span class="step-badge">Verify before pitch</span>
    </div>
    <div class="panel-body" id="intel-body"></div>
  </div>

  <!-- STEP 5: ACTIONS -->
  <div class="panel hidden" id="panel-actions">
    <div class="panel-head">
      <h2>Download</h2>
      <span class="step-badge">Step 2</span>
    </div>
    <div class="panel-body">
      <p style="font-size:13px;color:var(--tx2);margin-bottom:1rem;line-height:1.5">
        The PowerPoint uses the 20.20 branded template with Filson Pro fonts, correct layouts and your client's colour.
        Image placeholders include specific direction for the creative team.
        Fees show [FEE: TBC] — apply the rate card before sending.
      </p>
      <div class="actions-bar">
        <button class="btn btn-gold" id="download-btn" onclick="downloadPPTX()">
          ↓ Download PowerPoint
        </button>
        <button class="btn btn-outline" onclick="rebuildAndDownload()">
          ↓ Rebuild from edited sections
        </button>
        <button class="btn btn-outline" onclick="resetAll()" style="margin-left:auto">
          New brief
        </button>
      </div>
      <div id="rebuild-status" style="font-size:12px;color:var(--tx2);margin-top:.5rem;display:none"></div>
    </div>
  </div>

</div><!-- .page -->

<script>
let activeTab = 'pdf';
let currentJobId = null;
let pollInterval = null;
let lastProgressLen = 0;
let currentMeta = {};
let jobDone = false;

function switchTab(t) {
  activeTab = t;
  document.getElementById('panel-pdf').classList.toggle('hidden', t !== 'pdf');
  document.getElementById('panel-text').classList.toggle('hidden', t !== 'text');
  document.getElementById('tab-pdf').className = 'tab-btn ' + (t === 'pdf' ? 'active' : 'inactive');
  document.getElementById('tab-text').className = 'tab-btn ' + (t === 'text' ? 'active' : 'inactive');
}

async function submitBrief() {
  var errEl = document.getElementById('submit-error');
  errEl.classList.add('hidden');

  var fd = new FormData();
  if (activeTab === 'pdf') {
    var f = document.getElementById('brief-pdf').files[0];
    if (!f) { errEl.textContent = 'Please select a PDF file first.'; errEl.classList.remove('hidden'); return; }
    fd.append('brief_pdf', f);
  } else {
    var txt = document.getElementById('brief-text').value.trim();
    if (!txt) { errEl.textContent = 'Please paste the brief text.'; errEl.classList.remove('hidden'); return; }
    fd.append('brief_text', txt);
  }

  document.getElementById('submit-btn').disabled = true;
  document.getElementById('submit-btn').textContent = 'Submitting...';

  try {
    var resp = await fetch('/submit', { method: 'POST', body: fd });
    var data = await resp.json();
    if (data.error) throw new Error(data.error);
    currentJobId = data.job_id;
    // Put job_id in URL so user can bookmark/debug
    window.history.replaceState(null, '', '/?job=' + data.job_id);
    showProgress();
    pollInterval = setInterval(pollStatus, 2000);
  } catch(e) {
    errEl.textContent = e.message;
    errEl.classList.remove('hidden');
    document.getElementById('submit-btn').disabled = false;
    document.getElementById('submit-btn').textContent = 'Generate proposal →';
  }
}

function showProgress() {
  document.getElementById('panel-progress').classList.remove('hidden');
  document.getElementById('nav-status').textContent = 'Generating… (job: ' + currentJobId + ')';
}

async function pollStatus() {
  if (!currentJobId) return;
  try {
    var resp = await fetch('/status/' + currentJobId);
    var data = await resp.json();

    // Update progress
    if (data.progress && data.progress.length > lastProgressLen) {
      var latest = data.progress[data.progress.length - 1];
      document.getElementById('prog-msg').textContent = latest.msg;
      if (latest.pct != null) {
        document.getElementById('prog-bar').style.width = latest.pct + '%';
      }
      lastProgressLen = data.progress.length;
    }

    // Show sections as they arrive
    if (data.sections && data.sections.length > 0) {
      document.getElementById('panel-sections').classList.remove('hidden');
      renderSections(data.sections);
    }

    // Store meta
    if (data.meta && data.meta.client) {
      currentMeta = data.meta;
    }

    // Done
    if (data.status === 'done' || data.status === 'error') {
      clearInterval(pollInterval);
      jobDone = true;
      document.getElementById('prog-bar').style.width = '100%';

      if (data.status === 'done') {
        document.getElementById('status-pill').textContent = 'Complete';
        document.getElementById('status-pill').className = 'pill pill-done';
        document.getElementById('nav-status').textContent = data.pptx_ready ? 'Ready to download' : 'Sections ready';
        document.getElementById('prog-msg').textContent = data.pptx_ready ? 'Complete' : (data.pptx_error || 'Sections complete');
      } else {
        document.getElementById('status-pill').textContent = 'Error';
        document.getElementById('status-pill').className = 'pill pill-error';
        document.getElementById('nav-status').textContent = 'Error';
        document.getElementById('prog-msg').textContent = data.error || 'Unknown error';
      }

      // Show intel if available
      if (data.intel && Object.keys(data.intel).length) {
        renderIntel(data.intel, data.meta);
      }
      // Always show actions panel if we have sections
      if (data.sections && data.sections.length > 0) {
        var actionsPanel = document.getElementById('panel-actions');
        actionsPanel.classList.remove('hidden');
        // Update download button based on pptx status
        var dlBtn = document.getElementById('download-btn');
        if (data.pptx_ready) {
          dlBtn.disabled = false;
          dlBtn.textContent = '↓ Download PowerPoint';
        } else {
          dlBtn.disabled = true;
          dlBtn.textContent = 'PowerPoint unavailable — use Rebuild';
          // Show error detail
          var rs = document.getElementById('rebuild-status');
          rs.style.display = 'block';
          rs.style.color = '#A32D2D';
          rs.textContent = data.pptx_error
            ? 'PowerPoint build failed: ' + data.pptx_error + '. Try "Rebuild from edited sections".'
            : 'PowerPoint not built. Try "Rebuild from edited sections".';
        }
        actionsPanel.scrollIntoView({ behavior: 'smooth', block: 'start' });
      }
    }

  } catch(e) {
    console.error('Poll error:', e);
  }
}

function renderSections(sections) {
  var list = document.getElementById('sections-list');
  sections.forEach(function(sec, i) {
    var existing = document.getElementById('sec-card-' + sec.id);
    if (existing) {
      // Update textarea if user hasn't edited it
      var ta = existing.querySelector('textarea');
      if (ta && ta.dataset.pristine !== 'false') {
        ta.value = sec.body;
        ta.style.height = 'auto';
        ta.style.height = ta.scrollHeight + 'px';
      }
      return;
    }
    var card = document.createElement('div');
    card.className = 'section-card';
    card.id = 'sec-card-' + sec.id;
    card.innerHTML =
      '<div class="section-head" onclick="toggleSection(this)">' +
        '<span class="section-head-title">' + sec.heading + '</span>' +
        '<button class="section-copy-btn" onclick="copySec(event,\\'' + sec.id + '\\')">Copy</button>' +
      '</div>' +
      '<div class="section-body">' +
        '<textarea id="sec-ta-' + sec.id + '" onchange="this.dataset.pristine=\\'false\\'" ' +
          'oninput="this.style.height=\\'auto\\';this.style.height=this.scrollHeight+\\'px\\'">' +
          escHtml(sec.body) + '</textarea>' +
      '</div>';
    list.appendChild(card);
    var ta = card.querySelector('textarea');
    ta.dataset.pristine = 'true';
    setTimeout(function(){ ta.style.height='auto'; ta.style.height=ta.scrollHeight+'px'; }, 50);
    card.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  });
}

function toggleSection(head) {
  var body = head.nextElementSibling;
  body.style.display = body.style.display === 'none' ? 'block' : 'none';
}

function copySec(e, sid) {
  e.stopPropagation();
  var ta = document.getElementById('sec-ta-' + sid);
  if (!ta) return;
  var btn = e.target;
  navigator.clipboard.writeText(ta.value).then(function() {
    btn.textContent = 'Copied ✓';
    setTimeout(function(){ btn.textContent = 'Copy'; }, 2000);
  });
}

function renderIntel(intel, meta) {
  document.getElementById('panel-intel').classList.remove('hidden');
  var rows = [];
  if (intel.contact_profile) rows.push(['Contact', intel.contact_profile]);
  if (intel.org_context)     rows.push(['Organisation right now', intel.org_context]);
  if (intel.why_now)         rows.push(['Why this brief exists', intel.why_now]);
  if (intel.ambitions)       rows.push(['Strategic ambitions', intel.ambitions]);

  document.getElementById('intel-body').innerHTML =
    '<div class="intel-card">' +
    rows.map(function(r) {
      return '<div class="intel-row"><div class="intel-lbl">' + r[0] + '</div>' + escHtml(r[1]) + '</div>';
    }).join('') +
    '<p style="font-size:11px;color:var(--tx2);margin-top:.75rem">Verify key facts before the pitch meeting.</p>' +
    '</div>';
}

function collectSections() {
  var secs = [];
  document.querySelectorAll('[id^="sec-ta-"]').forEach(function(ta) {
    var sid = ta.id.replace('sec-ta-', '');
    var card = document.getElementById('sec-card-' + sid);
    var heading = card ? card.querySelector('.section-head-title').textContent : sid;
    secs.push({ id: sid, heading: heading, body: ta.value });
  });
  return secs;
}

function downloadPPTX() {
  if (!currentJobId) return;
  window.location.href = '/download/' + currentJobId;
}

async function rebuildAndDownload() {
  var sections = collectSections();
  var st = document.getElementById('rebuild-status');
  st.style.display = 'block';
  st.textContent = 'Rebuilding PowerPoint from your edited sections...';

  try {
    var resp = await fetch('/rebuild', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ job_id: currentJobId, sections: sections, meta: currentMeta })
    });
    var data = await resp.json();
    if (data.error) throw new Error(data.error);
    st.textContent = 'Done — downloading...';
    setTimeout(function() { window.location.href = '/download/' + currentJobId; }, 500);
  } catch(e) {
    st.textContent = 'Error: ' + e.message;
    st.style.color = '#A32D2D';
  }
}

function resetAll() {
  clearInterval(pollInterval);
  currentJobId = null;
  lastProgressLen = 0;
  currentMeta = {};
  jobDone = false;
  document.getElementById('sections-list').innerHTML = '';
  document.getElementById('intel-body').innerHTML = '';
  document.getElementById('brief-pdf').value = '';
  document.getElementById('brief-text').value = '';
  document.getElementById('prog-bar').style.width = '0%';
  document.getElementById('prog-msg').textContent = 'Starting...';
  document.getElementById('submit-btn').disabled = false;
  document.getElementById('submit-btn').textContent = 'Generate proposal →';
  document.getElementById('nav-status').textContent = '';
  ['panel-progress','panel-sections','panel-intel','panel-actions'].forEach(function(id) {
    document.getElementById(id).classList.add('hidden');
  });
  document.getElementById('submit-error').classList.add('hidden');
  document.getElementById('rebuild-status').style.display = 'none';
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

function escHtml(s) {
  return (s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}
</script>
</body>
</html>
"""

@app.route('/')
def index():
    return INDEX_HTML

@app.route('/submit', methods=['POST'])
def submit():
    if not ANTHROPIC_KEY:
        return jsonify({'error': 'API key not configured on server.'}), 500

    job_id = str(uuid.uuid4())[:8]
    save_job(job_id, {
        'status': 'running',
        'progress': [],
        'sections': [],
        'meta': {},
        'intel': {},
        'extracted': {},
        'pptx_path': None,
        'error': None,
    })

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

@app.route('/debug/<job_id>')
def debug(job_id):
    """Shows full job state for troubleshooting."""
    job = load_job(job_id)
    if not job:
        return jsonify({'error': 'Job not found', 'jobs_dir': JOBS_DIR}), 404
    return jsonify({
        'status':     job.get('status'),
        'error':      job.get('error'),
        'pptx_path':  job.get('pptx_path'),
        'pptx_exists': os.path.exists(job['pptx_path']) if job.get('pptx_path') else False,
        'template_exists': os.path.exists(TEMPLATE_PATH),
        'template_path': TEMPLATE_PATH,
        'jobs_dir': JOBS_DIR,
        'sections_count': len(job.get('sections', [])),
        'progress_last': job.get('progress', [{}])[-1] if job.get('progress') else None,
    })

@app.route('/status/<job_id>')
def status(job_id):
    job = load_job(job_id)
    if not job:
        return jsonify({'error': 'Job not found'}), 404
    return jsonify({
        'status':     job.get('status'),
        'progress':   job.get('progress', []),
        'sections':   job.get('sections', []),
        'meta':       job.get('meta', {}),
        'intel':      job.get('intel', {}),
        'error':      job.get('error'),
        'pptx_error': job.get('pptx_error'),
        'pptx_ready': bool(job.get('pptx_path') and os.path.exists(job.get('pptx_path',''))),
    })

@app.route('/rebuild', methods=['POST'])
def rebuild():
    """Rebuild PPTX from edited sections."""
    data = request.get_json()
    job_id = data.get('job_id')
    sections = data.get('sections', [])
    meta = data.get('meta', {})

    if not sections:
        return jsonify({'error': 'No sections provided'}), 400

    try:
        pptx_path, _tmpdir = build_pptx(sections, meta)
        if job_id:
            update_job(job_id, pptx_path=pptx_path, status='done', pptx_error=None)
        return jsonify({'status': 'ok', 'job_id': job_id})
    except Exception as e:
        import traceback
        return jsonify({'error': str(e), 'detail': traceback.format_exc()[-500:]}), 500

@app.route('/download/<job_id>')
def download(job_id):
    job = load_job(job_id)
    if not job:
        return 'Job not found — jobs are cleared when the server restarts. Please generate again.', 404
    if job.get('error'):
        return f'Generation failed: {job["error"]}', 500
    if not job.get('pptx_path'):
        return f'PowerPoint not ready yet — status is {job.get("status","unknown")}. Try again in a moment.', 404
    if not os.path.exists(job['pptx_path']):
        return 'PowerPoint file missing — server may have restarted. Please generate again.', 404

    venue = job.get('meta', {}).get('venue', 'Proposal').replace(' ', '_').replace("'", '').replace('&','and').replace('(','').replace(')','')
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
