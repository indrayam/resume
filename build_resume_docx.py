from pathlib import Path

from PIL import Image, ImageDraw
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_ROW_HEIGHT_RULE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor


ROOT = Path(__file__).resolve().parent
OUT_DIR = ROOT / "output" / "docx"
OUT_DIR.mkdir(parents=True, exist_ok=True)
PHOTO_SRC = ROOT / "tmp" / "pdfs" / "resume-photo.jpg"
PHOTO_CIRCLE = ROOT / "tmp" / "qa-docx" / "resume-photo-circle.png"
OUTPUT = OUT_DIR / "anand-sharma-resume.docx"

ACCENT = "35B4DE"
TEXT = "25282A"
MUTED = "6D7377"
WHITE = "FFFFFF"
FONT = "Courier New"


def set_run_font(run, size=10.5, bold=False, color=TEXT, italic=False):
    run.font.name = FONT
    run._element.get_or_add_rPr().rFonts.set(qn("w:ascii"), FONT)
    run._element.get_or_add_rPr().rFonts.set(qn("w:hAnsi"), FONT)
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = RGBColor.from_string(color)


def set_paragraph_border(paragraph, *, bottom=None):
    ppr = paragraph._p.get_or_add_pPr()
    borders = ppr.find(qn("w:pBdr"))
    if borders is None:
        borders = OxmlElement("w:pBdr")
        ppr.append(borders)
    if bottom:
        tag = OxmlElement("w:bottom")
        tag.set(qn("w:val"), "single")
        tag.set(qn("w:sz"), str(bottom.get("sz", 6)))
        tag.set(qn("w:space"), str(bottom.get("space", 1)))
        tag.set(qn("w:color"), bottom.get("color", TEXT))
        borders.append(tag)


def remove_borders(element):
    ppr = element._element.pPr if hasattr(element._element, "pPr") else element._p.get_or_add_pPr()
    borders = ppr.find(qn("w:pBdr")) if ppr is not None else None
    if borders is not None:
        ppr.remove(borders)


def set_paragraph_shading(paragraph, fill):
    ppr = paragraph._p.get_or_add_pPr()
    shd = ppr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        ppr.append(shd)
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)


def set_cell_borders(cell, color=WHITE, size=0):
    tcpr = cell._tc.get_or_add_tcPr()
    borders = tcpr.first_child_found_in("w:tcBorders")
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tcpr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        tag = borders.find(qn(f"w:{edge}"))
        if tag is None:
            tag = OxmlElement(f"w:{edge}")
            borders.append(tag)
        tag.set(qn("w:val"), "nil" if size == 0 else "single")
        tag.set(qn("w:sz"), str(size))
        tag.set(qn("w:color"), color)


def set_cell_width(cell, width_inches):
    cell.width = Inches(width_inches)
    tcpr = cell._tc.get_or_add_tcPr()
    tcw = tcpr.first_child_found_in("w:tcW")
    if tcw is None:
        tcw = OxmlElement("w:tcW")
        tcpr.append(tcw)
    tcw.set(qn("w:w"), str(round(width_inches * 1440)))
    tcw.set(qn("w:type"), "dxa")


def set_cell_margins(cell, top=0, start=0, bottom=0, end=0):
    tcpr = cell._tc.get_or_add_tcPr()
    tc_mar = tcpr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tcpr.append(tc_mar)
    for side, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tc_mar.find(qn(f"w:{side}"))
        if node is None:
            node = OxmlElement(f"w:{side}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def add_hyperlink(paragraph, text, url, color=TEXT, underline=False, size=10.5):
    part = paragraph.part
    rid = part.relate_to(url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", is_external=True)
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), rid)
    run = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    rfonts = OxmlElement("w:rFonts")
    rfonts.set(qn("w:ascii"), FONT)
    rfonts.set(qn("w:hAnsi"), FONT)
    rpr.append(rfonts)
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), str(round(size * 2)))
    rpr.append(sz)
    color_el = OxmlElement("w:color")
    color_el.set(qn("w:val"), color)
    rpr.append(color_el)
    if underline:
        u = OxmlElement("w:u")
        u.set(qn("w:val"), "single")
        rpr.append(u)
    run.append(rpr)
    text_el = OxmlElement("w:t")
    text_el.text = text
    run.append(text_el)
    hyperlink.append(run)
    paragraph._p.append(hyperlink)
    return hyperlink


def add_page_field(paragraph):
    run = paragraph.add_run()
    set_run_font(run, size=8, color=MUTED)
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    text_run = OxmlElement("w:t")
    text_run.text = "1"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.extend([begin, instr, separate, text_run, end])


def make_circle_photo():
    image = Image.open(PHOTO_SRC).convert("RGB")
    side = min(image.size)
    left = (image.width - side) // 2
    top = max(0, (image.height - side) // 2 - 35)
    top = min(top, image.height - side)
    image = image.crop((left, top, left + side, top + side)).resize((600, 600), Image.Resampling.LANCZOS)
    rgba = image.convert("RGBA")
    mask = Image.new("L", rgba.size, 0)
    ImageDraw.Draw(mask).ellipse((8, 8, 592, 592), fill=255)
    rgba.putalpha(mask)
    canvas = Image.new("RGBA", rgba.size, (255, 255, 255, 0))
    canvas.alpha_composite(rgba)
    draw = ImageDraw.Draw(canvas)
    draw.ellipse((8, 8, 592, 592), outline=(37, 40, 42, 255), width=5)
    PHOTO_CIRCLE.parent.mkdir(parents=True, exist_ok=True)
    canvas.save(PHOTO_CIRCLE)


def configure_styles(doc):
    normal = doc.styles["Normal"]
    normal.font.name = FONT
    normal._element.rPr.rFonts.set(qn("w:ascii"), FONT)
    normal._element.rPr.rFonts.set(qn("w:hAnsi"), FONT)
    normal.font.size = Pt(10.5)
    normal.font.color.rgb = RGBColor.from_string(TEXT)
    normal.paragraph_format.space_after = Pt(4)
    normal.paragraph_format.line_spacing = 1.03

    title = doc.styles["Title"]
    title.font.name = FONT
    title._element.rPr.rFonts.set(qn("w:ascii"), FONT)
    title._element.rPr.rFonts.set(qn("w:hAnsi"), FONT)
    title.font.size = Pt(22)
    title.font.bold = True
    title.font.color.rgb = RGBColor.from_string(TEXT)
    title.paragraph_format.space_before = Pt(0)
    title.paragraph_format.space_after = Pt(2)
    title.paragraph_format.line_spacing = 1.0
    remove_borders(title)

    for style_name in ("Heading 1", "Heading 2", "Heading 3"):
        style = doc.styles[style_name]
        style.font.name = FONT
        style._element.rPr.rFonts.set(qn("w:ascii"), FONT)
        style._element.rPr.rFonts.set(qn("w:hAnsi"), FONT)
        style.font.size = Pt(10.5)
        style.font.bold = True
        style.font.color.rgb = RGBColor.from_string(ACCENT if style_name == "Heading 1" else TEXT)
        style.paragraph_format.keep_with_next = True
        style.paragraph_format.line_spacing = 1.0


def add_plain_paragraph(doc, text, *, bold=False, color=TEXT, size=10.5, before=0, after=4, keep=False):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.line_spacing = 1.03
    p.paragraph_format.keep_with_next = keep
    run = p.add_run(text)
    set_run_font(run, size=size, bold=bold, color=color)
    return p


def add_section_heading(doc, text):
    p = doc.add_paragraph(style="Heading 1")
    p.paragraph_format.space_before = Pt(11)
    p.paragraph_format.space_after = Pt(5)
    p.paragraph_format.keep_with_next = True
    p.paragraph_format.line_spacing = 1.0
    run = p.add_run(text.upper())
    set_run_font(run, size=10.5, bold=True, color=ACCENT)
    set_paragraph_border(p, bottom={"color": TEXT, "sz": 6, "space": 2})
    return p


def add_bullet(doc, text, *, after=1.2):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.17)
    p.paragraph_format.first_line_indent = Inches(-0.11)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.line_spacing = 1.03
    p.paragraph_format.widow_control = True
    bullet = p.add_run("•  ")
    set_run_font(bullet, size=10.5, color=TEXT)
    run = p.add_run(text)
    set_run_font(run, size=10.5, color=TEXT)
    return p


def add_role(doc, title, *, job_title=None, dates=None, location=None, bullets=None, description=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(5)
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.keep_with_next = True
    p.paragraph_format.line_spacing = 1.0
    set_run_font(p.add_run(title), size=10.5, bold=True, color=TEXT)

    lines = [("Job Title:", job_title), ("Dates:", dates), ("Location:", location)]
    lines = [(label, value) for label, value in lines if value]
    if lines:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.keep_with_next = True if bullets or description else False
        p.paragraph_format.line_spacing = 1.0
        for index, (label, value) in enumerate(lines):
            label_run = p.add_run(label + " ")
            set_run_font(label_run, size=10.5, bold=True, color=TEXT)
            value_run = p.add_run(value)
            set_run_font(value_run, size=10.5, color=MUTED)
            if index < len(lines) - 1:
                p.add_run().add_break()

    if description:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.03
        set_run_font(p.add_run(description), size=10.5, color=TEXT)
    if bullets:
        for bullet in bullets:
            add_bullet(doc, bullet)


def add_company(doc, name):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(5)
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.keep_with_next = True
    p.paragraph_format.line_spacing = 1.0
    set_run_font(p.add_run(name), size=10.5, bold=True, color=TEXT)
    return p


def build_document():
    make_circle_photo()
    doc = Document()
    configure_styles(doc)
    section = doc.sections[0]
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)
    section.left_margin = Inches(0.62)
    section.right_margin = Inches(0.62)
    section.top_margin = Inches(0.52)
    section.bottom_margin = Inches(0.55)
    section.header_distance = Inches(0.0)
    section.footer_distance = Inches(0.23)
    section.gutter = Inches(0)

    # Full-width cyan rule at the top of every page.
    header = section.header
    hp = header.paragraphs[0]
    hp.paragraph_format.space_before = Pt(0)
    hp.paragraph_format.space_after = Pt(0)
    hp.paragraph_format.line_spacing = 1.0
    hp.paragraph_format.left_indent = Inches(-0.62)
    hp.paragraph_format.right_indent = Inches(-0.62)
    hp.paragraph_format.line_spacing = Pt(3)
    set_paragraph_shading(hp, ACCENT)
    hp.add_run(" ")

    footer = section.footer
    fp = footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fp.paragraph_format.space_before = Pt(0)
    fp.paragraph_format.space_after = Pt(0)
    add_page_field(fp)

    # Resume masthead with an editable text block and a circular image.
    table = doc.add_table(rows=1, cols=2)
    table.autofit = False
    table.allow_autofit = False
    row = table.rows[0]
    row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
    left, right = row.cells
    set_cell_width(left, 6.05)
    set_cell_width(right, 1.21)
    for cell in (left, right):
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
        set_cell_borders(cell)
        set_cell_margins(cell, top=0, start=0, bottom=0, end=0)

    p = left.paragraphs[0]
    p.style = doc.styles["Title"]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.0
    remove_borders(p)
    set_run_font(p.add_run("Anand Sharma"), size=22, bold=True, color=TEXT)

    p = left.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.0
    set_run_font(p.add_run("Architect / Senior Staff Engineer"), size=10.5, bold=True, color=TEXT)

    p = left.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.0
    set_run_font(p.add_run("Staff Software Engineer, Solutions Architect, DevOps, AI-Pilled"), size=10.5, color=TEXT)

    for contact in ("Austin, TX", "(408) 597-7614", "anand.sharma@gmail.com"):
        p = left.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.13)
        p.paragraph_format.first_line_indent = Inches(-0.09)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.0
        set_run_font(p.add_run("•  "), size=10.5, color=TEXT)
        set_run_font(p.add_run(contact), size=10.5, color=TEXT)
    p = left.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.13)
    p.paragraph_format.first_line_indent = Inches(-0.09)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.0
    set_run_font(p.add_run("•  "), size=10.5, color=TEXT)
    add_hyperlink(p, "LinkedIn", "https://www.linkedin.com/in/indrayam/", color=TEXT, underline=False)

    p = right.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    picture_run = p.add_run()
    picture = picture_run.add_picture(str(PHOTO_CIRCLE), width=Inches(1.18))
    picture._inline.docPr.set("title", "Anand Sharma portrait")
    picture._inline.docPr.set("descr", "Portrait of Anand Sharma")

    spacer = doc.add_paragraph()
    spacer.paragraph_format.space_before = Pt(0)
    spacer.paragraph_format.space_after = Pt(0)
    spacer.paragraph_format.line_spacing = 1.0

    add_section_heading(doc, "Summary")
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.03
    add_hyperlink(p, "Software architect", "https://resume.indrayam.com/architect.html", color=TEXT, underline=False)
    set_run_font(p.add_run(" who still believes the Unix terminal is the most elegant interface ever invented. I started in graduate school with an SGI workstation running IRIX; more than two decades later, I bring that same instinct to my work at Cisco, where I build enterprise DevOps and developer-platform systems used by 30,000+ engineers. My focus is making complex software delivery feel simple - from CI/CD, release orchestration, and cloud-native architecture to AI-assisted tools and multi-agent systems. I care deeply about the details, but I also understand the importance of zooming out - distilling the essence of the problem, its context, and its possibilities so others can see what matters."), size=10.5, color=TEXT)
    add_plain_paragraph(doc, "I like to build things, evangelize the ideas behind them, and inspire the people around me to make things that matter.", after=4)

    add_section_heading(doc, "Key Achievements")
    achievements = [
        "Invited to present at the prestigious Cisco Live 2017 and Cisco Live 2016 events on advancements in DevOps and Enterprise Class Continuous Delivery.",
        "Led Cisco IT's first Software Developer Experience as a Thought Leader, Evangelist, and Solutions Architect.",
        "Innovated development of Cisco's first centrally managed Continuous Delivery Platform.",
        "Architected the first customer-facing Enterprise App Store for Cisco's Cius product.",
        "Created the first Social Computing Platform for Cisco.",
        "Oversaw development of the first Cisco API Management Platform, featuring a customer-facing API Console.",
    ]
    for item in achievements:
        add_bullet(doc, item)

    add_section_heading(doc, "Experience")
    add_company(doc, "Cisco Systems")
    add_role(doc, "Senior Staff Software Engineer", job_title="Sr. Staff Software Engineer (Enterprise DevOps-as-a-Service)", dates="February 2023 - Present", location="Austin, Texas Metropolitan Area - Remote", bullets=[
        "Lead efforts to standardize an enterprise-scale PPM tool for Operations and Engineering, including conducting a thorough analysis of Aha!, Atlassian Jira Align, and ServiceNow.",
        "Drive widespread adoption of Jira Align across multiple teams following a presentation to the COO on key findings.",
        "Investigate solutions to modernize the stack, reduce technical debt, and offer a frictionless experience for the 25K+ Cisco employees across Engineering, Customer Support, and Operations, replacing a homegrown tool for defect lifecycle management.",
    ])
    add_role(doc, "Staff Software Engineer", dates="January 2022 - February 2023", location="Austin, Texas - Remote", bullets=[
        "Spearhead the definition, design, documentation, and delivery of cutting-edge software delivery (Enterprise DevOps) tools and experiences to a team of over 30K software engineers.",
        "Shape enterprise-wide decision-making by presenting frequent, impactful strategic and operational updates related to services and service offerings to Senior IT Leadership.",
        "Provide insightful thought leadership and technical solutions while evangelizing solutions to Cisco IT customers.",
        "Reduced costs by launching the Cisco FCS Software Distribution solution using Dev Hub Download to support shipping FCS binaries as container images instead of paying Cloud Providers to host FCS container images.",
    ])
    add_role(doc, "Senior DevOps Architect", job_title="Sr. Solutions Architect, DevOps", dates="February 2019 - January 2022", location="Raleigh, North Carolina - On-site", bullets=[
        "Introduced first-of-its-kind Code Analytics tool to measure software delivery performance (quality, security, and release metrics) for over approximately 17K unique software, enabling 1K+ engineers to track progress on DevSecOps metrics.",
        "Increased adoption of Cisco IT CI/CD practices for software delivery to 90% across the organization.",
        "Boosted the percentage of IT applications deployed into production with unit tests from 7% to 66% and improved unit test coverage for applications by 600%.",
        "Led event-driven microservices architecture build using a Java SpringBoot, Kafka, Redis, and Angular UX stack.",
    ])
    add_role(doc, "DevOps Architect", job_title="IT Architect, DevOps", dates="January 2018 - February 2019", location="Raleigh, North Carolina - On-site", bullets=[
        "Drove delivery of approximately 50K software promotions to higher lifecycle environments by creating Code Release as a single software release orchestrator to deliver software using any deployment tool anywhere.",
        "Grew total raw Production Deployments number by 90% and increased median daily Production Deployments by approximately 300% to approximately 100 per day.",
        "Reduced software promotion times to approximately 10 minutes, which improved Deployment Frequency DORA metrics for Cisco IT.",
        "Led event-driven microservices architecture build using a Java SpringBoot, Kafka, Redis, Argo Workflow, and React-based UX stack.",
    ])
    add_role(doc, "DevOps Architect", job_title="IT Architect, DevOps", dates="March 2016 - January 2018", location="Raleigh, North Carolina - On-site", bullets=[
        "Led Cisco IT's first-ever Software Developer Experience Portal as a Thought Leader, Evangelist, and Solutions Architect. The Software Developer Experience removes friction and contextual work for the developer, allowing them to continue to focus on building software.",
        "Improved the Lead Time for Changes DORA metric by reducing self-service requests to approximately 5 minutes, resulting in the self-service of approximately 22K CRUD requests across 4K+ unique Software Engineers using Code Console in 2018.",
        "Oversaw event-driven Microservices architecture built using Java SpringBoot, MongoDB, Redis, and React-based UX stack.",
    ])
    add_role(doc, "DevOps Architect", job_title="IT Architect, API Management", dates="February 2014 - March 2016", location="Raleigh, North Carolina - On-site", bullets=[
        "Launched Cisco IT's first Continuous Delivery Platform with a mix of open-source and commercial tools, resulting in approximately 90% adoption of CI/CD tools and mature software engineering practices to deliver software.",
        "Partnered with Cisco IT Delivery Organizations to successfully deliver an enterprise-wide Continuous Delivery solution that scaled across approximately 100 IT Services and approximately 2,000 IT Applications.",
        "Enhanced deployment of IT Apps into production with Unit Tests from less than 10% to 60%+ by building a homegrown solution to handle continuous delivery of Oracle software, addressing a gap in Oracle's existing offerings. This solution is currently used by 90% of Cisco teams for 100+ deployments a day.",
        "Delivered a series of core DevOps tools including GitHub SaaS, GitHub Enterprise, CloudBees Jenkins, Artifactory, SonarQube, Spinnaker, and CloudBees CDRO.",
    ])
    add_role(doc, "Information Technology System Architect", dates="July 2010 - February 2014", location="Raleigh, North Carolina - On-site", description="No detailed description was surfaced in the current LinkedIn profile.")
    add_role(doc, "Information Technology Technical Lead", job_title="IT Technical Leader, Web 2.0 (Customer Advocacy IT Team)", dates="September 2005 - July 2010", location="San Jose, California - On-site", description="No detailed description was surfaced in the current LinkedIn profile.")
    add_role(doc, "Information Technology Engineer", job_title="IT Software Engineer (Cisco Advanced Services IT Team)", dates="August 2001 - September 2005", location="San Jose, California - On-site", bullets=[
        "Architected and designed web-based applications for the Advanced Services business unit.",
        "Developed web-based application services on the J2EE platform, primarily using open-source MVC frameworks such as Struts and WebWork.",
        "Continued to act in the capacity of System Administrator of the Professional Services Automation application. Member of the core team responsible for both the US and Global launch of this application.",
        "Provided 24x7 systems and application administrative support to all AS-IT applications, including ASA (Advanced Services Automation) and ASAP (Advanced Services Account Portal).",
    ])
    add_role(doc, "Information Technology Analyst", dates="February 2001 - August 2001", bullets=[
        "Responsible for understanding business needs and assisted in authoring business requirements documents (BRDs).",
        "Worked with the Technical Lead in authoring Systems Requirement Documents (SRDs).",
        "Assisted in the formulation of testing and quality assurance scripts and plans for applications.",
    ])

    add_company(doc, "Avaya")
    add_role(doc, "Web Architect", dates="March 2000 - February 2001", description="The New Technology/Architecture group architected the various components of the e-business infrastructure. It was also responsible for researching new technologies with the aim of building world-class e-business services for customers.", bullets=[
        "Researched ways to make Avaya's e-business applications WAP-enabled and researched new publishing frameworks based on Java and XSLT technologies.",
    ])

    add_company(doc, "Lucent Technologies")
    add_role(doc, "IT Software Engineer", dates="January 1999 - February 2001", bullets=[
        "Played a lead role in designing and developing Lucent's first Single Sign-on solution for external web applications, leveraging the Netegrity SiteMinder solution.",
    ])
    add_role(doc, "Web Technical Leader", dates="February 1999 - March 2000", description="Customer Self Support (support.lucent.com) was an extremely busy site. As the lead developer of the first SSO pilot, took on the added responsibility of leading the effort to migrate the entire application into an NES presentation layer, NAS (Netscape Application Server) business-logic layer, and Oracle 7.3.x data layer framework.", bullets=[
        "Led a team of three developers and rewrote all server-side programs into Java applogics.",
        "Administered the development NES/NAS/Oracle environment for the CSS project.",
    ])

    add_company(doc, "Arizona State University")
    add_role(doc, "IS Lab Lead", dates="May 1996 - December 1998", description="The IS Lab catered to the technology needs of Arizona State instructors and students. It focused on helping the community embrace new and upcoming technologies in education.", bullets=[
        "Worked on projects involving web applications, proofs of concept for new technologies and products, and helping IS Lab customers get introduced to them.",
    ])

    add_section_heading(doc, "Education")
    add_role(doc, "Master of Computer Science, Software Engineering", description="Arizona State University - Tempe, AZ")
    add_role(doc, "Master of Science, Industrial Engineering", description="Arizona State University - Tempe, AZ")

    add_section_heading(doc, "Awards and Presentations")
    add_role(doc, "2018 Gold Stevie Winner", description="Enterprise Software Release Digitization, Secure DevOps")
    add_role(doc, "Cisco Live 2016", description='Panel discussion: "Addressing the What\'s, Why\'s and How\'s in the New World of DevOps"')
    add_role(doc, "Cisco Live 2017", description='Presentation: "Enterprise Class Continuous Delivery"')
    add_role(doc, "Cloud Identity Summit 2012 - Vail, CO", description='Presentation: "Who Says Elephant Can\'t Dance? Securely Externalizing APIs @ Cisco" hosted by Ping Identity')

    # Keep the document free of Word's default editing metadata.
    core = doc.core_properties
    core.title = "Anand Sharma Resume"
    core.subject = "Editable resume"
    core.author = "Anand Sharma"
    core.keywords = "resume, software architect, DevOps"
    doc.save(OUTPUT)
    print(OUTPUT)


if __name__ == "__main__":
    build_document()
