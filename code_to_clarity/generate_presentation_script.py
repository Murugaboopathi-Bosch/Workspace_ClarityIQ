"""
Generates the "Onboard Yourself AI" presentation speaking script as a Word document.
Run:  python generate_presentation_script.py
"""

from docx import Document
from docx.shared import Pt, RGBColor, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from datetime import datetime
import os

# ── Colours ──────────────────────────────────────────────────────
BOSCH_RED    = RGBColor(0xE2, 0x00, 0x1A)   # Bosch Red
NAVY         = RGBColor(0x0A, 0x16, 0x28)   # Dark navy
TEAL         = RGBColor(0x00, 0x7A, 0xC2)   # Bosch Blue
LIGHT_GRAY   = RGBColor(0xF4, 0xF6, 0xF8)
WHITE        = RGBColor(0xFF, 0xFF, 0xFF)
DARK_GRAY    = RGBColor(0x44, 0x44, 0x44)


def set_cell_bg(cell, hex_str):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  hex_str)
    tcPr.append(shd)


def add_colored_heading(doc, text, level=1, color=NAVY):
    p    = doc.add_heading(text, level=level)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    for run in p.runs:
        run.font.color.rgb = color
        run.font.bold      = True
    return p


def add_banner(doc, title, subtitle=""):
    """Full-width banner cell."""
    tbl  = doc.add_table(rows=1, cols=1)
    cell = tbl.cell(0, 0)
    set_cell_bg(cell, "0A1628")
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(title)
    r.font.size      = Pt(22)
    r.font.bold      = True
    r.font.color.rgb = WHITE
    if subtitle:
        p2 = cell.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r2 = p2.add_run(subtitle)
        r2.font.size      = Pt(12)
        r2.font.color.rgb = RGBColor(0xA0, 0xC4, 0xFF)
    doc.add_paragraph()
    return tbl


def add_section_banner(doc, number, title):
    tbl  = doc.add_table(rows=1, cols=1)
    cell = tbl.cell(0, 0)
    set_cell_bg(cell, "E2001A")
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r = p.add_run(f"  {number}   {title.upper()}")
    r.font.size      = Pt(13)
    r.font.bold      = True
    r.font.color.rgb = WHITE
    doc.add_paragraph()
    return tbl


def add_speaker_note(doc, text):
    """Light-blue speaker note box."""
    tbl  = doc.add_table(rows=1, cols=1)
    cell = tbl.cell(0, 0)
    set_cell_bg(cell, "E8F4FD")
    p = cell.paragraphs[0]
    r = p.add_run("🎤  SPEAKER NOTE:  ")
    r.font.bold      = True
    r.font.color.rgb = TEAL
    r.font.size      = Pt(9)
    r2 = p.add_run(text)
    r2.font.color.rgb = DARK_GRAY
    r2.font.size      = Pt(9)
    r2.font.italic    = True
    doc.add_paragraph()


def add_speaking_block(doc, label, text):
    """Numbered speaking point with bold label."""
    p = doc.add_paragraph(style="Normal")
    p.paragraph_format.left_indent  = Cm(0.5)
    p.paragraph_format.space_after  = Pt(6)
    r1 = p.add_run(f"{label}  ")
    r1.font.bold      = True
    r1.font.color.rgb = BOSCH_RED
    r1.font.size      = Pt(11)
    r2 = p.add_run(text)
    r2.font.color.rgb = DARK_GRAY
    r2.font.size      = Pt(11)


def add_bullet(doc, text, indent=1):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.left_indent = Cm(indent * 0.8)
    r = p.add_run(text)
    r.font.size      = Pt(11)
    r.font.color.rgb = DARK_GRAY


def add_divider(doc):
    p = doc.add_paragraph("─" * 90)
    p.runs[0].font.color.rgb = RGBColor(0xCC, 0xCC, 0xCC)
    p.runs[0].font.size      = Pt(8)


def add_timing_box(doc, timing):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = p.add_run(f"⏱  {timing}")
    r.font.size      = Pt(9)
    r.font.bold      = True
    r.font.color.rgb = RGBColor(0x88, 0x88, 0x88)


# ═════════════════════════════════════════════════════════════════
#  BUILD THE DOCUMENT
# ═════════════════════════════════════════════════════════════════
doc = Document()

# Page margins
for section in doc.sections:
    section.top_margin    = Cm(1.8)
    section.bottom_margin = Cm(1.8)
    section.left_margin   = Cm(2.0)
    section.right_margin  = Cm(2.0)

# ── COVER ─────────────────────────────────────────────────────────
add_banner(
    doc,
    "🚀  ONBOARD YOURSELF AI",
    "Code to Clarity  ·  Bosch Digital Innovation League 2026"
)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("Presentation Speaking Script")
r.font.size      = Pt(14)
r.font.bold      = True
r.font.color.rgb = NAVY

p2 = doc.add_paragraph()
p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
r2 = p2.add_run(f"Audience: Investors · Managers · Department Heads\n{datetime.today().strftime('%B %d, %Y')}")
r2.font.size      = Pt(10)
r2.font.color.rgb = DARK_GRAY
r2.font.italic    = True

doc.add_paragraph()
add_divider(doc)
doc.add_paragraph()

# ── HOW TO USE THIS DOCUMENT ──────────────────────────────────────
add_colored_heading(doc, "📖  How to Use This Document", level=2, color=TEAL)
add_bullet(doc, "Read the text in each section out loud — it is written in simple, natural spoken English.")
add_bullet(doc, "Red labels (SAY:, PAUSE:, etc.) are instructions for you — do NOT read them out loud.")
add_bullet(doc, "Blue speaker notes give you tips on what to do at that moment (show slide, demo, etc.).")
add_bullet(doc, "Total estimated speaking time: 12 – 15 minutes.")
doc.add_paragraph()
add_divider(doc)
doc.add_page_break()


# ══════════════════════════════════════════════════════════════════
#  SECTION 1 — OPENING
# ══════════════════════════════════════════════════════════════════
add_section_banner(doc, "SECTION 1", "Opening — Welcome & Introduction")
add_timing_box(doc, "1 – 2 minutes")

add_speaker_note(doc, "Walk to the front. Smile. Take a breath. Make eye contact before you start speaking.")

add_speaking_block(doc, "SAY:", "Good morning everyone. Thank you so much for being here today.")

add_speaking_block(doc, "SAY:", "My name is Boopathi, and I am part of the BD COB AI Synergy team here at Bosch. Today I am going to share something that I built as part of the Bosch Digital Innovation League 2026 — an initiative we are calling 'Onboard Yourself AI'.")

add_speaking_block(doc, "SAY:", "This is not just a tool. It is a solution to a real pain point that our teams face every single day. And in the next 12 to 15 minutes, I am going to show you exactly what that problem is, how we solved it, and what it means for Bosch going forward.")

add_speaking_block(doc, "SAY:", "So let us get started.")

add_speaker_note(doc, "Advance to your first slide — the title slide with 'Code to Clarity'.")
doc.add_paragraph()
add_divider(doc)
doc.add_page_break()


# ══════════════════════════════════════════════════════════════════
#  SECTION 2 — THE PROBLEM
# ══════════════════════════════════════════════════════════════════
add_section_banner(doc, "SECTION 2", "The Problem — Why This Matters")
add_timing_box(doc, "2 – 3 minutes")

add_speaker_note(doc, "Speak slowly here. Let the audience feel the pain. This is the emotional hook.")

add_speaking_block(doc, "SAY:", "Let me start with a question. Raise your hand if your team has ever received a Java project — or any software project — and had absolutely no documentation for it.")

add_speaker_note(doc, "Pause. Look around the room. Give people a moment to respond.")

add_speaking_block(doc, "SAY:", "I thought so. This is one of the most common problems in software delivery — and it is one that costs us a huge amount of time and money.")

add_speaking_block(doc, "SAY:", "When a new developer joins a project, or when a project is handed over from one team to another, the first question is always: 'What does this system actually do?' And most of the time, the answer is — nobody really knows, unless they read through thousands of lines of code.")

add_speaking_block(doc, "SAY:", "Now, reading code is not the same as understanding business requirements. A developer can read the code. But a business manager cannot. A product owner cannot. And even a new developer needs days — sometimes weeks — just to understand what the system was built for.")

add_speaking_block(doc, "SAY:", "And here is the worst part. Even after all that effort, what you end up with is still just a person's interpretation — which may or may not be correct. There is no structured Business Requirements Document. There is no Functional Requirements Document. There is no single source of truth.")

add_speaking_block(doc, "SAY:", "This is the problem we set out to solve.")

add_speaker_note(doc, "Advance to the 'Problem Statement' slide if you have one.")
doc.add_paragraph()
add_divider(doc)
doc.add_page_break()


# ══════════════════════════════════════════════════════════════════
#  SECTION 3 — THE IDEA
# ══════════════════════════════════════════════════════════════════
add_section_banner(doc, "SECTION 3", "The Idea — What If AI Could Do This For Us?")
add_timing_box(doc, "1 – 2 minutes")

add_speaking_block(doc, "SAY:", "So the idea was simple. What if we could give an AI the source code of any Java project, and it would automatically generate a proper Business Requirements Document and a Functional Requirements Document — in minutes?")

add_speaking_block(doc, "SAY:", "Not a summary. Not bullet points. A real, professional, formatted Word document — the kind that a Business Analyst would produce after weeks of interviews and analysis.")

add_speaking_block(doc, "SAY:", "That is exactly what we built. We called it 'Code to Clarity'.")

add_speaker_note(doc, "Advance to the 'Solution Overview' slide.")
doc.add_paragraph()
add_divider(doc)
doc.add_page_break()


# ══════════════════════════════════════════════════════════════════
#  SECTION 4 — THE SOLUTION
# ══════════════════════════════════════════════════════════════════
add_section_banner(doc, "SECTION 4", "The Solution — Code to Clarity")
add_timing_box(doc, "2 – 3 minutes")

add_speaking_block(doc, "SAY:", "Code to Clarity is an AI-powered web application. It is built with Python, and it uses the same Large Language Model that powers tools like ChatGPT — but connected through the Bosch internal AI gateway, so everything stays within our network and our security boundaries.")

add_speaking_block(doc, "SAY:", "Here is how it works in three simple steps:")

add_bullet(doc, "Step 1 — You open the web application in your browser. You give it your project name, and you upload your Java source code folder.")
add_bullet(doc, "Step 2 — The AI reads every single Java file. It identifies the REST APIs, the business rules, the data entities, the service operations — all automatically.")
add_bullet(doc, "Step 3 — It generates two professional Word documents: a BRD — Business Requirements Document — and an FRD — Functional Requirements Document. You download them as a ZIP file. Done.")

add_speaking_block(doc, "SAY:", "The whole process takes less than two minutes for a typical project. Something that would normally take a Business Analyst two to three weeks — done in under two minutes.")

add_speaking_block(doc, "SAY:", "And the output is not generic. It uses the actual names from your code — your endpoint paths, your entity names, your service method names — and translates them into proper business language that anyone can understand.")

add_speaker_note(doc, "Advance to the 'How It Works' or architecture slide.")
doc.add_paragraph()
add_divider(doc)
doc.add_page_break()


# ══════════════════════════════════════════════════════════════════
#  SECTION 5 — LIVE DEMO WALKTHROUGH
# ══════════════════════════════════════════════════════════════════
add_section_banner(doc, "SECTION 5", "Live Demo Walkthrough")
add_timing_box(doc, "2 – 3 minutes")

add_speaker_note(doc, "Switch to your browser. Have the app running in your browser. Have the sample Java project ready to upload.")

add_speaking_block(doc, "SAY:", "Let me show you this live. I have the application running right here.")

add_speaking_block(doc, "SAY:", "This is the web interface. Very simple. You type in a project name — I will call it 'User Management System' — and then you select your Java source folder and click Generate.")

add_speaker_note(doc, "Upload the Java project folder and click Generate. Let the progress show on screen.")

add_speaking_block(doc, "SAY:", "You can see it is working now. In the background, the AI is reading the code, identifying the structure, and writing the documentation. Watch the progress bar.")

add_speaker_note(doc, "Wait for the download to complete. Then open the ZIP file and show the two Word documents.")

add_speaking_block(doc, "SAY:", "And just like that — we have our documents. A BRD and an FRD, both in professional Word format, ready to share with any stakeholder.")

add_speaking_block(doc, "SAY:", "Let me open the BRD quickly. You can see it has an Executive Summary, Business Objectives, Scope, Stakeholder roles, Business Rules — all properly numbered and formatted. This is production-quality documentation.")

add_speaker_note(doc, "Scroll slowly through the document. Don't rush. Let people read a few lines.")

add_speaking_block(doc, "SAY:", "And this was all generated automatically, just from the source code. No interviews. No manual writing. No weeks of effort.")

doc.add_paragraph()
add_divider(doc)
doc.add_page_break()


# ══════════════════════════════════════════════════════════════════
#  SECTION 6 — BUSINESS VALUE
# ══════════════════════════════════════════════════════════════════
add_section_banner(doc, "SECTION 6", "Business Value — Why This Matters for Bosch")
add_timing_box(doc, "2 minutes")

add_speaker_note(doc, "Return to your slides. This is the business case section — speak confidently.")

add_speaking_block(doc, "SAY:", "Now let us talk about what this actually means in business terms.")

add_speaking_block(doc, "SAY:", "The average time for a Business Analyst to produce a BRD and FRD for a medium-sized project is between 10 and 15 working days. With Code to Clarity, that same output is produced in under two minutes.")

add_speaking_block(doc, "SAY:", "Think about what that means for onboarding. When a new team member joins a project, instead of spending their first two weeks trying to understand the codebase, they can read a clear, structured document on day one. They are productive from the start.")

add_speaking_block(doc, "SAY:", "Think about project handovers. When a project moves from one team to another — or from a vendor to Bosch — there is now an automatic way to generate the documentation, with no dependency on the original developers being available.")

add_speaking_block(doc, "SAY:", "Think about compliance and audits. We always struggle to have up-to-date documentation. With this tool, documentation can be regenerated at any time, from the latest version of the code, in minutes.")

add_speaking_block(doc, "SAY:", "And everything runs through the Bosch AI gateway. There is no data leaving Bosch. It is secure, it is compliant, and it is ready to scale.")

doc.add_paragraph()
add_divider(doc)
doc.add_page_break()


# ══════════════════════════════════════════════════════════════════
#  SECTION 7 — THE BIGGER PICTURE
# ══════════════════════════════════════════════════════════════════
add_section_banner(doc, "SECTION 7", "The Bigger Picture — Onboard Yourself AI")
add_timing_box(doc, "1 – 2 minutes")

add_speaking_block(doc, "SAY:", "Now, Code to Clarity is the proof of concept — the first working version. But the bigger vision is what we are calling 'Onboard Yourself AI'.")

add_speaking_block(doc, "SAY:", "The idea is this: What if any employee at Bosch — whether they are a developer, a manager, or a new joiner — could onboard themselves to any project, any system, any codebase, just by using AI? Without needing anyone to explain it to them. Without reading hundreds of pages of outdated documentation.")

add_speaking_block(doc, "SAY:", "Code to Clarity is step one. It takes code and produces documentation. But the next steps could include:")

add_bullet(doc, "Automatically generating test cases from the same code.")
add_bullet(doc, "Creating project summaries in multiple languages for global teams.")
add_bullet(doc, "Building a chatbot that lets you ask questions about any project in plain English.")
add_bullet(doc, "Extending support beyond Java to Python, C++, and other languages used at Bosch.")

add_speaking_block(doc, "SAY:", "The foundation is already built. What you see today is just the beginning.")

doc.add_paragraph()
add_divider(doc)
doc.add_page_break()


# ══════════════════════════════════════════════════════════════════
#  SECTION 8 — WHAT WE NEED
# ══════════════════════════════════════════════════════════════════
add_section_banner(doc, "SECTION 8", "What We Need — The Ask")
add_timing_box(doc, "1 minute")

add_speaker_note(doc, "This is your call to action. Be direct and confident.")

add_speaking_block(doc, "SAY:", "We have proven the concept. The tool works. The output quality is good. And the business case is clear.")

add_speaking_block(doc, "SAY:", "What we need now is support to take this from a proof of concept to a production-ready internal tool that any Bosch team can use.")

add_speaking_block(doc, "SAY:", "Specifically, we are looking for:")

add_bullet(doc, "A dedicated internal hosting environment — so any Bosch team can access it through the intranet.")
add_bullet(doc, "Access to a stable Bosch AI API subscription — so we have guaranteed capacity for multiple users.")
add_bullet(doc, "A small cross-functional team of two to three people — to extend the features and integrate with existing Bosch tools like JIRA, Confluence, or SharePoint.")

add_speaking_block(doc, "SAY:", "With that support, we can have a fully production-ready version within three months.")

doc.add_paragraph()
add_divider(doc)
doc.add_page_break()


# ══════════════════════════════════════════════════════════════════
#  SECTION 9 — CLOSING
# ══════════════════════════════════════════════════════════════════
add_section_banner(doc, "SECTION 9", "Closing — Thank You")
add_timing_box(doc, "1 minute")

add_speaker_note(doc, "Slow down. Smile. Make this feel personal and sincere.")

add_speaking_block(doc, "SAY:", "I want to leave you with one thought.")

add_speaking_block(doc, "SAY:", "We talk a lot about AI transformation. About using AI to improve our processes. But transformation does not happen in big announcements. It happens in small, practical tools that solve real problems for real people, every single day.")

add_speaking_block(doc, "SAY:", "Code to Clarity is that kind of tool. It is practical. It solves a real problem. It saves real time. And it is built right here, by the Bosch team, for the Bosch team.")

add_speaking_block(doc, "SAY:", "Thank you very much for your time and attention today. I am happy to take any questions.")

add_speaker_note(doc, "Step back from the podium slightly. Smile. Wait for questions or applause.")

doc.add_paragraph()
add_divider(doc)
doc.add_paragraph()


# ══════════════════════════════════════════════════════════════════
#  Q&A PREP
# ══════════════════════════════════════════════════════════════════
add_colored_heading(doc, "❓  Q&A Preparation — Likely Questions & Answers", level=2, color=TEAL)
doc.add_paragraph()

qa_pairs = [
    (
        "Q: Is the data from our code safe? Does it go outside Bosch?",
        "A: No. Everything runs through the Bosch internal AI gateway — aoai-farm.bosch-temp.com. The code never leaves the Bosch network. It is the same security model used by all Bosch AI tools."
    ),
    (
        "Q: What programming languages does it support right now?",
        "A: The current version supports Java — specifically Spring Boot projects, which is the most common framework in our software teams. Extending to other languages like Python or C++ is on the roadmap and is technically straightforward."
    ),
    (
        "Q: How accurate is the generated documentation?",
        "A: Very accurate for structure — the endpoints, entities, and business rules are extracted directly from the code, so they are factual. The AI adds the business context and language on top. For a production tool, we would recommend a quick human review before sharing externally — which takes maybe 30 minutes instead of two weeks."
    ),
    (
        "Q: How much does it cost to run?",
        "A: The main cost is the AI API usage. For typical project sizes, each BRD+FRD generation costs less than one euro in API calls. For a team of 20 developers generating documentation once per sprint, the monthly cost would be negligible compared to the time saved."
    ),
    (
        "Q: Can it work with our existing documentation tools like Confluence or SharePoint?",
        "A: Not yet — but this is exactly one of the next features on the roadmap. The documents are currently generated as Word files, which can be manually uploaded. Automated integration with Confluence or SharePoint is very achievable with a few weeks of development."
    ),
    (
        "Q: Who built this? Is there a team behind it?",
        "A: This proof of concept was built individually as part of the Bosch Digital Innovation League 2026. With proper team support, it can be scaled into a full product very quickly."
    ),
]

for q, a in qa_pairs:
    p = doc.add_paragraph()
    r1 = p.add_run(q + "\n")
    r1.font.bold      = True
    r1.font.color.rgb = NAVY
    r1.font.size      = Pt(11)
    r2 = p.add_run(a)
    r2.font.color.rgb = DARK_GRAY
    r2.font.size      = Pt(11)
    doc.add_paragraph()


# ── Footer note ───────────────────────────────────────────────────
add_divider(doc)
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("Code to Clarity  ·  Onboard Yourself AI  ·  Bosch Digital Innovation League 2026")
r.font.size      = Pt(8)
r.font.color.rgb = RGBColor(0xAA, 0xAA, 0xAA)
r.font.italic    = True


# ── Save ──────────────────────────────────────────────────────────
output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")
os.makedirs(output_dir, exist_ok=True)

filename = f"Presentation_Script_OnboardYourselfAI_{datetime.today().strftime('%Y%m%d')}.docx"
path     = os.path.join(output_dir, filename)
doc.save(path)
print(f"\n✅  Presentation script saved to:\n    {path}\n")
