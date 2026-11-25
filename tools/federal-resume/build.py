"""Build the 2-page federal resume (VA / IRS applications) as a new DOCX.

Usage:  uv run --no-project --with python-docx python tools/federal-resume/build.py OUT.docx
Then:   tools/docx-update/export.sh OUT.docx OUT.pdf

This does not touch the published resume in output/. Titles, locations, and education follow LinkedIn (read 2026-11-24); Cisco
dates for 2018 onward follow references/resume.md, as Anand confirmed; bullet
text follows references/resume.md.
"""
import sys

from docx import Document
from docx.enum.text import WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

FONT = "Calibri"
BODY = 10
ACCENT = RGBColor(0x1F, 0x3A, 0x5F)

CONTACT = ["Austin, TX", "(408) 597-7614", "anand.sharma@gmail.com",
           "linkedin.com/in/indrayam", "U.S. Citizen"]

SUPERVISOR = "Mohan Raghupathy"

SUMMARY = (
    "Senior Staff Software Engineer and Solutions Architect with 25+ years at Cisco "
    "Systems building developer platforms for 30,000+ engineers: CI/CD pipelines, "
    "release orchestration, DevSecOps metrics, and multi-agent AI systems. Owns "
    "initiatives end to end: defines the problem and user value, architects the "
    "solution, aligns the engineers who build it, builds adoption metrics in from day "
    "one, and briefs senior executives until it lands."
)

# (title, employer, location, start, end, [bullets])
JOBS = [
    ("Senior Staff Software Engineer (AI-Native Engineering)", "Cisco Systems",
     "Austin, TX (Remote)", "08/2025", "Present", [
        "Led the design and rollout of the DevTools AI Agent, a supervisor-agent system that routes developer requests to specialized agents for onboarding and operational support across 12+ software delivery tools.",
        "Architected and delivered 12+ DevTools MCP (Model Context Protocol) services exposing 100+ tools, giving AI assistants and agents standardized access to the software delivery platform.",
        "Partnered on the company-wide rollout of AI coding tools (GitHub Copilot, Cursor, ChatGPT/Codex, Claude Code), and conceived and designed CodeTokens to give leadership per-developer visibility into token and model usage.",
        "Led solution architecture to modernize a 30+-year-old defect tracking system (40,000+ users, 1M+ edits/year), driving alignment across multiple business units.",
    ]),
    ("Senior Staff Software Engineer (Enterprise DevOps-as-a-Service)", "Cisco Systems",
     "Austin, TX (Remote)", "01/2022", "07/2025", [
        "Defined, designed, documented, and delivered enterprise software delivery (DevOps) tools and experiences for 30,000+ software engineers.",
        "Shaped enterprise-wide decisions through frequent strategic and operational briefings on services and service offerings to senior IT leadership.",
        "Reduced costs by launching the Cisco FCS Software Distribution solution, using Dev Hub Download to ship FCS binaries as container images instead of paying cloud providers to host them.",
        "Led the effort to standardize an enterprise-scale PPM tool for Operations and Engineering, including an analysis of Aha!, Atlassian Jira Align, and ServiceNow.",
    ]),
    ("Senior Solution Architect, DevOps", "Cisco Systems",
     "Raleigh, NC / Austin, TX (Remote from 06/2021)", "02/2019", "01/2022", [
        "Introduced the first-of-its-kind Code Analytics tool to measure software delivery performance (quality, security, and release metrics) for ~17,000 unique software components, enabling 1,000+ engineers to track DevSecOps metrics.",
        "Increased adoption of Cisco IT CI/CD practices for software delivery to 90% across the organization.",
        "Raised the share of IT applications deployed to production with unit tests from 7% to 66% and improved unit test coverage by 600%.",
        "Led an event-driven microservices build on Java Spring Boot, Kafka, Redis, and Angular.",
    ]),
    ("IT Architect, DevOps", "Cisco Systems",
     "Raleigh, NC (On-site)", "01/2018", "02/2019", [
        "Created Code Release, a single release orchestrator that delivers software with any deployment tool anywhere, driving ~50,000 software promotions to higher lifecycle environments.",
        "Grew total production deployments by 90% and median daily production deployments by ~300% to ~100 per day; cut promotion times to ~10 minutes, improving DORA Deployment Frequency.",
    ]),
    ("IT Architect, DevOps", "Cisco Systems",
     "Raleigh, NC (On-site)", "03/2016", "01/2018", [
        "Led Cisco IT's first Software Developer Experience Portal as thought leader, evangelist, and solutions architect, removing friction so developers can focus on building software.",
        "Improved DORA Lead Time for Changes by cutting self-service requests to ~5 minutes; ~22,000 CRUD requests self-served by 4,000+ engineers through Code Console in 2018.",
    ]),
    ("IT Architect, API Management", "Cisco Systems",
     "Raleigh, NC (On-site)", "02/2014", "03/2016", [
        "Launched Cisco IT's first Continuous Delivery Platform, reaching ~90% adoption of CI/CD tools across ~100 IT services and ~2,000 IT applications.",
        "Built a homegrown continuous delivery solution for Oracle software, raising production deployments with unit tests from under 10% to 60%+; used by 90% of Cisco teams for 100+ deployments a day.",
        "Delivered core DevOps tools: GitHub SaaS and Enterprise, CloudBees Jenkins, Artifactory, SonarQube, Spinnaker, and CloudBees CDRO.",
    ]),
    ("Information Technology System Architect", "Cisco Systems",
     "Raleigh, NC (On-site)", "07/2010", "02/2014", []),
    ("Information Technology Technical Lead", "Cisco Systems",
     "San Jose, CA (On-site)", "09/2005", "07/2010", []),
    ("Information Technology Engineer", "Cisco Systems",
     "San Jose, CA (On-site)", "08/2001", "09/2005", [
        "Architected and built web applications for the Advanced Services business unit on J2EE with Struts and WebWork.",
        "System administrator for the Professional Services Automation application; core team member for its US and global launch; provided 24x7 support for all AS-IT applications.",
    ]),
    ("Information Technology Analyst", "Cisco Systems",
     "San Jose, CA", "02/2001", "08/2001", [
        "Gathered business needs and co-authored business and system requirements documents and QA test plans.",
    ]),
    ("Web Architect", "Avaya (Lucent Octel division, spun off as Avaya)", None, "03/2000", "02/2001", [
        "Researched WAP enablement of Avaya's e-business applications and new Java/XSLT publishing frameworks.",
    ]),
    ("IT Software Engineer", "Lucent Technologies", None, "01/1999", "02/2001", [
        "Played a lead role designing and developing Lucent's first single sign-on solution for external web applications (Netegrity SiteMinder).",
    ]),
    ("Web Technical Leader", "Lucent Technologies", None, "02/1999", "03/2000", [
        "Led three developers migrating the Customer Self Support site (support.lucent.com) to a Netscape NES/NAS and Oracle stack.",
    ]),
]

EDUCATION = [
    ("Master of Science, Computer Science", "Arizona State University, Tempe, AZ", "1997 – 1998"),
    ("Master of Science in Engineering, Information Systems", "Arizona State University, Tempe, AZ", "1995 – 1997"),
    ("Bachelor of Technology, Mechanical Engineering", "National Institute of Technology (NIT), Kurukshetra, India", "1991 – 1995"),
]

SKILLS = [
    ("AI", "Multi-agent systems, MCP services, Claude Code, Codex, Cursor, GitHub Copilot"),
    ("Cloud & compute", "AWS (EC2, EKS, S3, Lambda), containers, Kubernetes, OpenStack"),
    ("Architecture & data", "Event-driven microservices; MongoDB, Postgres, Oracle, Kafka, Redis"),
    ("DevSecOps", "GitHub, GitHub Actions, Jenkins, Artifactory, SonarQube, Spinnaker, SAST/DAST"),
    ("Integration & SaaS", "Mashery, Layer7, Boomi, Salesforce Sales Cloud, Workday"),
]

AWARDS = [
    "2018 Gold Stevie Award: Enterprise Software Release Digitization, Secure DevOps",
    "Cisco Live 2017 presenter: “Enterprise Class Continuous Delivery”",
    "Cisco Live 2016 panelist: “Addressing the What’s, Why’s and How’s in the New World of DevOps”",
    "Cloud Identity Summit 2012 presenter: “Who Says Elephant Can’t Dance? Securely Externalizing APIs @ Cisco”",
]


def fmt(p, before=0, after=0, line=1.0):
    pf = p.paragraph_format
    pf.space_before, pf.space_after, pf.line_spacing = Pt(before), Pt(after), line


def run(p, text, bold=False, italic=False, size=BODY, color=None):
    r = p.add_run(text)
    r.bold, r.italic = bold, italic
    r.font.size, r.font.name = Pt(size), FONT
    if color is not None:
        r.font.color.rgb = color
    return r


def bottom_border(p):
    pPr = p._p.get_or_add_pPr()
    bdr = OxmlElement("w:pBdr")
    b = OxmlElement("w:bottom")
    for k, v in (("val", "single"), ("sz", "6"), ("space", "1"), ("color", "1F3A5F")):
        b.set(qn(f"w:{k}"), v)
    bdr.append(b)
    pPr.append(bdr)


def heading(doc, text):
    p = doc.add_paragraph()
    fmt(p, before=7, after=3)
    run(p, text.upper(), bold=True, size=11, color=ACCENT)
    bottom_border(p)
    p.paragraph_format.keep_with_next = True


def bullet(doc, text):
    p = doc.add_paragraph(style="List Bullet")
    fmt(p, after=1)
    p.paragraph_format.left_indent = Inches(0.2)
    p.paragraph_format.first_line_indent = Inches(-0.14)
    run(p, text)


def right_tab(p, width):
    p.paragraph_format.tab_stops.add_tab_stop(width, WD_TAB_ALIGNMENT.RIGHT)


def build(out):
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Inches(8.5), Inches(11)
    for side in ("left_margin", "right_margin"):
        setattr(sec, side, Inches(0.6))
    sec.top_margin = sec.bottom_margin = Inches(0.5)
    width = sec.page_width - sec.left_margin - sec.right_margin

    normal = doc.styles["Normal"]
    normal.font.name, normal.font.size = FONT, Pt(BODY)
    normal.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)

    p = doc.add_paragraph()
    fmt(p)
    p.alignment = 1
    run(p, "ANAND SHARMA", bold=True, size=18, color=ACCENT)
    p = doc.add_paragraph()
    fmt(p, after=2)
    p.alignment = 1
    run(p, "  |  ".join(CONTACT), size=BODY)

    heading(doc, "Summary")
    p = doc.add_paragraph()
    fmt(p)
    run(p, SUMMARY)

    heading(doc, "Professional Experience")
    for i, (title, employer, loc, start, end, bullets) in enumerate(JOBS):
        p = doc.add_paragraph()
        fmt(p, before=0 if i == 0 else 5)
        right_tab(p, width)
        p.paragraph_format.keep_with_next = True
        run(p, title, bold=True)
        run(p, f"\t{start} – {end}", bold=True)
        p = doc.add_paragraph()
        fmt(p, after=1)
        p.paragraph_format.keep_with_next = bool(bullets)
        right_tab(p, width)
        run(p, employer + (f", {loc}" if loc else ""), italic=True)
        run(p, "\tFull-time, 40 hours per week", italic=True)
        if i == 0:
            p = doc.add_paragraph()
            fmt(p, after=1)
            p.paragraph_format.keep_with_next = True
            run(p, f"Supervisor: {SUPERVISOR}", italic=True)
        for b in bullets:
            bullet(doc, b)

    heading(doc, "Education")
    for degree, school, years in EDUCATION:
        p = doc.add_paragraph()
        fmt(p, after=1)
        right_tab(p, width)
        run(p, degree, bold=True)
        if school:
            run(p, f", {school}")
        run(p, f"\t{years}")

    heading(doc, "Technical Skills")
    for label, items in SKILLS:
        p = doc.add_paragraph()
        fmt(p, after=1)
        run(p, f"{label}: ", bold=True)
        run(p, items)

    heading(doc, "Awards and Presentations")
    for a in AWARDS:
        bullet(doc, a)

    doc.core_properties.title = "Anand Sharma – Federal Resume"
    doc.core_properties.author = "Anand Sharma"
    doc.save(out)


if __name__ == "__main__":
    build(sys.argv[1])
