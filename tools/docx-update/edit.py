"""Apply the resume.md edits to a COPY of the hand-maintained Word resume.

Only text changes: every new paragraph or run is cloned from one already in
the document, so fonts, sizes, colours, bullets and spacing come from the
original. Usage: python3 edit.py in.docx out.docx [--move N:M ...]
"""
import os, re, sys, zipfile
from xml.sax.saxutils import escape

src, dst = sys.argv[1], sys.argv[2]
zin = zipfile.ZipFile(src)
doc = zin.read('word/document.xml').decode('utf-8')
rels = zin.read('word/_rels/document.xml.rels').decode('utf-8')

spans = [(m.start(), m.end()) for m in re.finditer(r'<w:p[ >].*?</w:p>', doc, re.S)]

# ---- Pre-split two hand-made paragraphs so roles can move between columns ----
# (a) The 2016-2018 role's last bullet is split across 80 (bullet) and 81 (a
#     tab-indented continuation that also holds the column break and the
#     EARLY CAREER heading). Join the continuation onto bullet 80 and leave 81
#     as just the break + heading. 80 shared 81's bottom border only so Word
#     would draw one rule under the heading; it no longer needs it.
# (b) 114 holds the sidebar's last line, a column break, and the 2014-2016
#     role title. Split it after the break (B gets the same pPr).
p80, p81, p114 = (doc[a:b] for a, b in (spans[80], spans[81], spans[114]))
assert 'Oversaw event-driven Microservices' in p80 and 'EARLY CAREER EXPERIENCE' in p81
assert 'Top proposals for' in p114 and 'IT ARCHITECT, API MANAGEMENT' in p114
cb = '<w:br w:type="column"/></w:r>'
pre81, post81 = p81.split(cb)
pre81_runs = pre81.split('</w:pPr>', 1)[1]
def last_run_start(x):
    return [m.start() for m in re.finditer(r'<w:r[ >]', x)][-1]   # runs carry w:rsid* attributes
pre81_runs = pre81_runs[:last_run_start(pre81_runs)]             # drop the run that opened the break
pre81_runs, n_tab = re.subn(r'<w:r(?: [^>]*)?>(?:<w:rPr>(?:(?!</w:rPr>).)*</w:rPr>)?<w:tab/></w:r>', '', pre81_runs, count=1, flags=re.S)
assert n_tab == 1 and 'Java SpringBoot' in pre81_runs and 'EARLY CAREER' not in pre81_runs
pre81_runs = re.sub(r'(<w:t)(?: xml:space="preserve")?>', r'\1 xml:space="preserve">', pre81_runs, count=1)
pre81_runs = pre81_runs.replace('<w:t xml:space="preserve">', '<w:t xml:space="preserve"> ', 1)
# Now one multi-line bullet, so give it the paragraph settings of the bullet
# before it in the same role (79) instead of 80's one-line hand indent.
p79 = doc[spans[79][0]:spans[79][1]]; assert 'Improved the Lead Time' in re.sub(r'<[^>]+>', '', p79)
ppr79 = re.search(r'<w:pPr>.*?</w:pPr>', p79, re.S).group(0)
new80 = re.sub(r'<w:pPr>.*?</w:pPr>', lambda m: ppr79, p80, count=1, flags=re.S).replace('</w:p>', pre81_runs + '</w:p>')
head81 = p81.split('</w:pPr>', 1)[0] + '</w:pPr>' + pre81[last_run_start(pre81):] + cb + post81
assert 'Java SpringBoot' not in head81 and 'EARLY CAREER EXPERIENCE' in head81
assert 'Java SpringBoot' in new80 and 'Oversaw event-driven' in new80
# Like the original, the column break belongs at the START of the paragraph
# that opens the next column (a break at the end of a paragraph leaves a blank
# line at the top of the new column). --break-before=N says which paragraph
# gets it; the default is the 2014-2016 title, as in the original.
a114, b114 = p114.split(cb)
BREAK_RUN = a114[last_run_start(a114):] + cb
a114 = a114[:last_run_start(a114)] + '</w:p>'
b114_open = re.match(r'<w:p[^>]*>', p114).group(0)
b114 = re.sub(r' w14:(?:paraId|textId)="[^"]*"', '', b114_open) + re.search(r'<w:pPr>.*?</w:pPr>', p114, re.S).group(0) + b114
doc = (doc[:spans[80][0]] + new80 + doc[spans[80][1]:spans[81][0]] + head81 + doc[spans[81][1]:spans[114][0]]
       + a114 + b114 + doc[spans[114][1]:])
spans = [(m.start(), m.end()) for m in re.finditer(r'<w:p[ >].*?</w:p>', doc, re.S)]
P = [doc[a:b] for a, b in spans]
ORIG = list(P)
# Now: 80 = whole bullet, 81 = column break + EARLY CAREER, 114 = sidebar end + break, 115 = 2014-2016 title.

def strip_ids(p):
    return re.sub(r' w14:(?:paraId|textId)="[^"]*"', '', p, count=2)

def first_rpr(p):
    body = p.split('</w:pPr>', 1)[1]
    return re.search(r'<w:rPr>.*?</w:rPr>', body, re.S).group(0)

def ppr(p):
    m = re.search(r'<w:pPr>.*?</w:pPr>', p, re.S)
    return m.group(0) if m else ''

def run(rpr, text, italic=False, underline=False):
    if italic:
        rpr = re.sub(r'(<w:rFonts[^>]*/>(?:<w:b/>)?(?:<w:bCs/>)?)', r'\1<w:i/><w:iCs/>', rpr, count=1)
    if underline:
        rpr = rpr.replace('</w:rPr>', '<w:u w:val="single"/></w:rPr>')
    return f'<w:r>{rpr}<w:t xml:space="preserve">{escape(text)}</w:t></w:r>'

def para(template, segments, pPr=None, keep_bookmarks=True, clone=False):
    """Rebuild a paragraph: template's pPr (or pPr), template's first-run rPr,
    new text. segments: str or list of (text, kind) with kind in {'', 'i', 'link:rId'}."""
    if isinstance(segments, str):
        segments = [(segments, '')]
    rpr = first_rpr(template)
    open_tag = re.match(r'<w:p[^>]*>', template).group(0)
    marks = ''.join(re.findall(r'<w:bookmark(?:Start|End)[^>]*/>', template)) if keep_bookmarks else ''
    out = []
    for text, kind in segments:
        if kind.startswith('link:'):
            out.append(f'<w:hyperlink r:id="{kind[5:]}">{run(rpr, text, underline=True)}</w:hyperlink>')
        else:
            out.append(run(rpr, text, italic=(kind == 'i')))
    p = open_tag + (pPr if pPr is not None else ppr(template)) + marks + ''.join(out) + '</w:p>'
    return strip_ids(p) if clone else p

def with_spacing_after(pPr, twips):
    # w:spacing sits after w:widowControl / numPr / pBdr / shd / tabs in pPr order.
    return pPr.replace('<w:widowControl/>', f'<w:widowControl/><w:spacing w:after="{twips}"/>', 1)

def must_replace(p, old, new):
    assert p.count(old) == 1, (old, p.count(old))
    return p.replace(old, new)

# Hyperlink for "Architect".
ARCH_RID = 'rIdArchitectPage'
if ARCH_RID not in rels:
    rels = rels.replace('</Relationships>',
        f'<Relationship Id="{ARCH_RID}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" '
        'Target="https://resume.indrayam.com/architect.html" TargetMode="External"/></Relationships>')

def check(i, start):
    t = ''.join(re.findall(r'<w:t(?: [^>]*)?>([^<]*)</w:t>', ORIG[i]))
    assert t.startswith(start), (i, t[:60])

# ---- Summary (para 8) -> three paragraphs -------------------------------
check(8, 'Results-driven Senior Software Engineer')
s8 = ORIG[8]
P[8] = (
    para(s8, [('Senior Staff Software Engineer and Solutions ', ''), ('Architect', 'link:' + ARCH_RID),
              (' who builds by the Unix philosophy: do one thing well, and make things composable. '
               'It has stood the test of time for a reason, and I strive to bring that same elegance '
               'and power into the products I build.', '')],
         pPr=with_spacing_after(ppr(s8), 100))
  + para(s8, "For the past decade, I have built developer platforms for Cisco’s 30,000+ engineers: "
             "opinionated CI/CD pipelines, release orchestration, and, most recently, multi-agent AI systems. "
             "I typically own an initiative end to end, the way a product GM would. I define the problem and "
             "the value to the user, architect the solution, align the engineers who build it, bake adoption "
             "metrics in from day one, and champion it with senior executives and future customers until it "
             "lands. I write code, craft the story behind it, and pitch it, caring as much about the details "
             "as about zooming out so others can see the value.",
         pPr=with_spacing_after(ppr(s8), 100), keep_bookmarks=False, clone=True)
  + para(s8, "Simply put, I like to build things that matter, evangelize the ideas behind them, and inspire "
             "the people around me to do their best work to date.", keep_bookmarks=False, clone=True)
)

# ---- Key Achievements ------------------------------------------------------
check(11, 'Invited to present'); check(12, 'Led Cisco IT'); check(16, 'Oversaw development')
item_ppr, last_ppr = ppr(ORIG[12]), ppr(ORIG[16])
P[11] = ''  # Cisco Live moves to the end
P[12] = (
    para(ORIG[12], "Helped shape the enterprise rollout of AI coding tools (GitHub Copilot, Cursor, Windsurf, "
                   "Codex, Claude Code) to 30,000+ engineers, and designed CodeTokens, a centralized in-house "
                   "system for tracking and governing token usage and spend across all five tools", clone=True)
  + para(ORIG[12], [("Led Cisco IT’s first Software Developer Experience Portal (", ''), ('Code', 'i'),
                    (") as a Thought Leader, Evangelist, and Solutions Architect. Today it serves ", ''),
                    ('all', 'i'), (" of Cisco’s software teams", '')])
)
P[16] = (
    ORIG[16]                       # keeps its own space-before, which makes the gap above it
  + para(ORIG[11], ''.join(re.findall(r'<w:t(?: [^>]*)?>([^<]*)</w:t>', ORIG[11])), pPr=last_ppr, clone=True)
)

# ---- Role 1: now AI-Native Engineering, 2025 - Present --------------------
check(19, 'SR. STAFF SOFTWARE ENGINEER'); check(20, 'Cisco Systems'); check(25, 'STAFF SOFTWARE ENGINEER')
P[19] = must_replace(ORIG[19], '<w:t>(Enterprise DevOps-as-a-Service)</w:t>', '<w:t>(AI-Native Engineering)</w:t>')
P[20] = must_replace(ORIG[20], '<w:t>3</w:t>', '<w:t>5</w:t>')
new_bullets = [
    "Led the design and rollout of the DevTools AI Agent, a supervisor-agent system that routes developer "
    "requests to specialized agents for onboarding and operational support across 12+ software delivery tools.",
    "Architected and delivered 12+ DevTools MCP (Model Context Protocol) services exposing 100+ tools, giving "
    "AI assistants and agents standardized access to the software delivery platform.",
    "Partnered on the company-wide rollout of AI coding tools (GitHub Copilot, Cursor, ChatGPT/Codex, Claude "
    "Code), and conceived and designed CodeTokens to give leadership per-developer visibility into token and "
    "model usage.",
    "Led solution architecture to modernize a 30+-year-old defect tracking system (40,000+ users, 1M+ "
    "edits/year), driving alignment across multiple business units.",
]
check(21, 'Lead efforts'); check(22, 'Drive widespread'); check(23, 'Investigate solutions')
P[21] = para(ORIG[21], new_bullets[0])
P[22] = para(ORIG[22], new_bullets[1])
P[23] = para(ORIG[23], new_bullets[2]) + para(ORIG[23], new_bullets[3], clone=True)

# ---- Role 2: Sr. Staff (Enterprise DevOps-as-a-Service), 2022 - 2025 ------
P[25] = strip_ids(ORIG[19])          # same two-line bold title as the original role 1
check(26, 'Cisco Systems')
P[26] = must_replace(ORIG[26], '<w:t>2023</w:t>', '<w:t>2025</w:t>')
check(60, 'Reduced costs')
P[60] = ORIG[60] + para(ORIG[60], "Led efforts to standardize an enterprise-scale PPM tool for Operations and "
                                  "Engineering, including conducting a thorough analysis of Aha!, Atlassian "
                                  "Jira Align, and ServiceNow.", clone=True)

# ---- Sidebar: Strengths, Experience heading, Languages ---------------------
# "Agile / Iterative Methodologies" becomes "AI Tools Wrangling" with a new
# sparkles icon (Office-icon style: 96x96 SVG, fill #47B9DF, 192px PNG
# fallback) and moves to sit after System Design.
check(31, 'System Design'); check(33, 'Agile / Iterative Methodologies')
ai = must_replace(ORIG[33], '<w:t>Agile / Iterative Methodologies</w:t>', '<w:t>Wrangling AI Agents/Bots</w:t>')
ai = must_replace(ai, 'r:embed="rId28"', 'r:embed="rIdSparklesPng"')
ai = must_replace(ai, 'r:embed="rId29"', 'r:embed="rIdSparklesSvg"')
ai = ai.replace('descr="Arrow circle with solid fill"', 'descr="Sparkles with solid fill"')
assert ai.count('Sparkles with solid fill') == 2
P[33] = ''
# Technical Architecture folds into System Design (keeps the Venn icon);
# Solutions Delivery & Deployment is dropped (DevOps and SDLC cover it).
check(30, ' Technical Architecture'); check(32, ' Solutions Delivery')
P[30] = ''
P[31] = must_replace(ORIG[31], '<w:t>System Design</w:t>', '<w:t>Architecture &amp; System Design</w:t>') + ai
P[32] = ''
ICON_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'icon')
NEW_MEDIA = {'word/media/image-sparkles.png': open(os.path.join(ICON_DIR, 'sparkles.png'), 'rb').read(),
             'word/media/image-sparkles.svg': open(os.path.join(ICON_DIR, 'sparkles.svg'), 'rb').read()}
for rid, target in (('rIdSparklesPng', 'media/image-sparkles.png'), ('rIdSparklesSvg', 'media/image-sparkles.svg')):
    if rid not in rels:
        rels = rels.replace('</Relationships>',
            f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
            f'Target="{target}"/></Relationships>')
check(40, 'EXPERIENCE / FLUENCY')
P[40] = must_replace(must_replace(ORIG[40], '<w:t>EXPERIENCE / FLUENC</w:t>', '<w:t>EXPERIENCE</w:t>'),
                     '<w:t>Y</w:t></w:r>', '</w:r>')
check(55, 'Languages:'); check(56, 'Rust, Python, Java')
P[55] = must_replace(ORIG[55], '<w:t>Languages:</w:t>', '<w:t>Languages of Interest:</w:t>')
P[56] = must_replace(ORIG[56], '<w:t xml:space="preserve">Rust, Python, Java </w:t>', '<w:t>Rust, Go</w:t>')

# ---- Early Career: only the job titles bold; company names regular -------
for i, company in ((83, 'Cisco Systems'), (86, 'Cisco Systems'), (88, 'Lucent Technologies')):
    check(i, company)
    # bold appears on the text run and on the paragraph mark; clear both
    assert '<w:b/>' in ORIG[i].split('</w:pPr>', 1)[1]
    P[i] = ORIG[i].replace('<w:b/>', '').replace('<w:bCs/>', '')
check(84, 'IT TECHNICAL LEADER, WEB 2.0')
bold_rpr = first_rpr(ORIG[84]); plain_rpr = bold_rpr.replace('<w:b/>', '')
P[84] = must_replace(ORIG[84],
    '<w:t>IT TECHNICAL LEADER, WEB 2.0 (Customer Advocacy IT Team) Cisco Systems</w:t></w:r>',
    '<w:t>IT TECHNICAL LEADER, WEB 2.0 (Customer Advocacy IT Team)</w:t></w:r>'
    f'<w:r>{bold_rpr}<w:br/></w:r><w:r>{plain_rpr}<w:t>Cisco Systems</w:t></w:r>')

# ---- Optional re-flow: --move A-B:C moves paragraphs A..B to just before C --
# "+spacer" appends a copy of the blank spacer paragraph used between roles (61).
moves = [a for a in sys.argv[3:] if a.startswith('--move=')]
SPACER = strip_ids(ORIG[61]); assert not re.search(r'<w:t[ >]', SPACER)
order = list(range(len(P)))
extra = {}
for mv in moves:
    spec = mv.split('=', 1)[1]
    add_spacer = spec.endswith('+spacer'); spec = spec.replace('+spacer', '')
    rng, dest = spec.split(':')
    a, b = map(int, rng.split('-'))
    block = [i for i in order if a <= i <= b]
    order = [i for i in order if not a <= i <= b]
    k = len(order) if dest == 'end' else order.index(int(dest))
    order[k:k] = block
    if add_spacer:
        extra[b] = SPACER
bb = [int(a.split('=')[1]) for a in sys.argv[3:] if a.startswith('--break-before=')] or [115]
for n in bb:
    assert '<w:br w:type="column"/>' not in P[n], f'paragraph {n} already has a column break'
    P[n] = re.sub(r'(</w:pPr>)', lambda m: m.group(1) + BREAK_RUN, P[n], count=1)
# --drop=N removes an empty spacer paragraph (refuses to drop one with text).
for d in [a for a in sys.argv[3:] if a.startswith('--drop=')]:
    n = int(d.split('=')[1]); assert not re.search(r'<w:t[ >]', P[n]), f'paragraph {n} has text'
    order.remove(n)

# Body-level markup between paragraphs (only the Summary's bookmarkEnds) stays
# attached to the paragraph before it.
gap = {i: doc[spans[i][1]:spans[i + 1][0]] for i in range(len(spans) - 1)}
for i, g in gap.items():
    assert not g.strip() or re.fullmatch(r'(?:<w:bookmarkEnd [^>]*/>)+', g.strip()), (i, g[:80])
out = doc[:spans[0][0]] + ''.join(P[i] + gap.get(i, '') + extra.get(i, '') for i in order) + doc[spans[-1][1]:]

# --photo=FILE swaps the header portrait. The circle, size, position and ring
# belong to the shape, so only the picture bytes change. image1.jfif is the
# modern picture; image2.jpeg is the same photo for the older-Word fallback.
photo = [a.split('=', 1)[1] for a in sys.argv[3:] if a.startswith('--photo=')]
PHOTO_FILES = ('word/media/image1.jfif', 'word/media/image2.jpeg')
if photo:
    assert zin.read(PHOTO_FILES[0]) == zin.read(PHOTO_FILES[1])
    photo_bytes = open(photo[0], 'rb').read()
    assert photo_bytes[:3] == b'\xff\xd8\xff', 'photo must be a JPEG'

zout = zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED)
for item in zin.infolist():
    data = zin.read(item.filename)
    if photo and item.filename in PHOTO_FILES:
        data = photo_bytes
    if item.filename == 'word/document.xml':
        data = out.encode('utf-8')
    elif item.filename == 'word/_rels/document.xml.rels':
        data = rels.encode('utf-8')
    zout.writestr(item, data)
ct = zin.read('[Content_Types].xml').decode()
assert 'Extension="png"' in ct and 'Extension="svg"' in ct, 'content types for png/svg missing'
for name, data in NEW_MEDIA.items():
    zout.writestr(name, data)
zout.close()
print('wrote', dst)
