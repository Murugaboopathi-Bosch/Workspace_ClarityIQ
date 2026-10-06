"""
╔══════════════════════════════════════════════════════════════════╗
║         CODE TO CLARITY — AI-Powered Requirement Discovery      ║
║         Bosch Digital Innovation League 2026                    ║
║                                                                  ║
║  Reads your Java project → Sends to OpenAI →                    ║
║  Generates BRD + FRD Word documents automatically               ║
╚══════════════════════════════════════════════════════════════════╝
"""

import os, re, sys, time
from pathlib import Path
from datetime import datetime
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

try:
    from openai import OpenAI
except ImportError:
    print("❌  Missing package. Run:  pip install openai")
    sys.exit(1)

try:
    import javalang
    HAS_JAVALANG = True
except ImportError:
    HAS_JAVALANG = False

# ══════════════════════════════════════════════════════════════════
#  ⚙️  CONFIG — Edit these 3 lines before running
# ══════════════════════════════════════════════════════════════════
CONFIG = {
    "project_path" : "./sample-java-project",   # ← path to your Java project
    "project_name" : "User Management",         # ← name shown in documents 
    # "genaiplatform-farm-subscription-key": "e9be0439278b47c5a83ef3f5cb9e00cc",  # ← your OpenAI API key (or set OPENAI_API_KEY env var)
    # "genaiplatform-farm-subscription-key": "0e82b61cb34d4d2591afc26a73c2a90c",                        # Thamarai ← your OpenAI API key (or set OPENAI_API_KEY env var)
    "genaiplatform-farm-subscription-key": "00e8f8a7482e49949a8f7c6de20d59d7",                        # ← your OpenAI API key (or set OPENAI_API_KEY env var)
    "openai_base_url": "https://aoai-farm.bosch-temp.com/api/openai/deployments/askbosch-prod-farm-openai-gpt-4o-mini-2024-07-18/chat/completions?api-version=2024-08-01-preview",                       # ← optional: Bosch gateway URL (leave blank for default)
    "openai_model"  : "gpt-4o",                  # ← model to use
    "output_dir"    : "./output",
}
# ══════════════════════════════════════════════════════════════════

NAVY   = RGBColor(0x0A, 0x16, 0x28)
TEAL   = RGBColor(0x00, 0xB4, 0xD8)
ACCENT = RGBColor(0xFF, 0x6B, 0x35)
GRAY   = RGBColor(0x64, 0x74, 0x8B)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)


# ──────────────────────────────────────────────────────────────────
# STEP 1 — LOAD JAVA FILES
# ──────────────────────────────────────────────────────────────────
def classify(name, content):
    n = name.lower(); c = content.lower()
    if "controller" in n:              return "Controller"
    if "service" in n:                 return "Service"
    if "repository" in n or "dao" in n:return "Repository"
    if "@entity" in c:                 return "Entity"
    if any(x in n for x in ["dto","request","response"]): return "DTO"
    if any(x in n for x in ["exception","notfound"]):     return "Exception"
    if any(x in n for x in ["config","configuration"]):   return "Config"
    if "@springbootapplication" in c:  return "Main"
    return "Other"

def load_project(path):
    print(f"\n{'─'*55}")
    print(f"  📂 STEP 1 — Loading Java files from:")
    print(f"  {path}")
    print(f"{'─'*55}")

    root = Path(path).resolve()
    if not root.exists():
        print(f"\n  ❌ Folder not found: {root}")
        print(f"  👉 Update CONFIG['project_path'] and try again.\n")
        sys.exit(1)

    files = []
    for p in sorted(root.rglob("*.java")):
        if "test" in str(p).lower(): continue
        content = p.read_text(encoding="utf-8", errors="ignore")
        files.append({
            "name"   : p.name,
            "path"   : str(p.relative_to(root)),
            "content": content,
            "lines"  : len(content.splitlines()),
            "type"   : classify(p.name, content),
        })

    from collections import Counter
    counts = Counter(f["type"] for f in files)
    print(f"\n  ✅ {len(files)} Java files found:\n")
    for t, c in sorted(counts.items(), key=lambda x: -x[1]):
        print(f"     {'█'*c:<12} {c}x  {t}")
    return files


# ──────────────────────────────────────────────────────────────────
# STEP 2 — EXTRACT CODE STRUCTURE
# ──────────────────────────────────────────────────────────────────
def extract_structure(files):
    print(f"\n{'─'*55}")
    print(f"  🔍 STEP 2 — Extracting code structure")
    print(f"{'─'*55}")

    endpoints, rules, entities, services = [], [], [], []

    http_map = {"GetMapping":"GET","PostMapping":"POST","PutMapping":"PUT",
                "DeleteMapping":"DELETE","PatchMapping":"PATCH"}

    for f in files:
        c = f["content"]

        # REST endpoints
        if f["type"] == "Controller":
            base_m = re.search(r'@RequestMapping\s*\(\s*["\']([^"\']+)["\']', c)
            base   = base_m.group(1) if base_m else ""
            for ann, method in http_map.items():
                for m in re.finditer(
                    rf'@{ann}\s*\(?["\']?([^"\')\n]*)["\']?\)?\s*\n\s*'
                    rf'(?:public|protected)\s+\S+\s+(\w+)',c):
                    path = (base + "/" + m.group(1).strip()).replace("//","/")
                    endpoints.append({"method":method,"path":path or base,
                                      "handler":m.group(2),"file":f["name"]})

        # Business rules from comments
        for m in re.finditer(r"//\s*Business Rule:\s*(.+)", c):
            rules.append(m.group(1).strip())

        # Validation annotations = business rules
        for ann in ["@NotBlank","@NotNull","@Email","@Size","@Min","@Max","@DecimalMin"]:
            if ann.lower() in c.lower():
                rules.append(f"Input validation enforced via {ann} in {f['name']}")

        # Entities
        if f["type"] == "Entity":
            fields = re.findall(r"private\s+(\w+)\s+(\w+)\s*;", c)
            entities.append({"name":f["name"].replace(".java",""), "fields":fields})

        # Services
        if f["type"] == "Service":
            methods = re.findall(r"public\s+\w+\s+(\w+)\s*\(", c)
            services.append({"name":f["name"].replace(".java",""), "methods":methods})

    print(f"\n  ✅ {len(endpoints)} REST endpoints detected")
    print(f"  ✅ {len(rules)} business rules / validations found")
    print(f"  ✅ {len(entities)} domain entities found")
    print(f"\n  REST Endpoints:")
    for ep in endpoints:
        print(f"     [{ep['method']:<6}] {ep['path']:<40} → {ep['handler']}()")

    return {"endpoints":endpoints, "rules":list(set(rules)),
            "entities":entities, "services":services}


# ──────────────────────────────────────────────────────────────────
# STEP 3 — BUILD PROMPTS
# ──────────────────────────────────────────────────────────────────
PERSONA = """
You are a Senior Business Analyst with 15+ years of enterprise experience.
Your task is to analyse Java Spring Boot source code and produce professional
documentation that business stakeholders and new developers can understand.

STRICT RULES:
1. Use BUSINESS language — never say "class", "method", "annotation", "JPA"
2. Focus on WHAT the system does, not HOW it is coded
3. Use actual names found in the code (endpoint paths, entity names, field names)
4. Be specific and detailed — generic vague descriptions are not acceptable
5. Include examples or scenarios where applicable to enhance understanding.
6. Write in complete sentences, professional tone
"""

def make_brd_prompt(name, structure, code):
    eps  = "\n".join(f"  [{e['method']}] {e['path']} → {e['handler']}()" for e in structure["endpoints"])
    ents = "\n".join(f"  {e['name']}: {[f[1] for f in e['fields'][:8]]}" for e in structure["entities"])
    svcs = "\n".join(f"  {s['name']}: {s['methods'][:6]}" for s in structure["services"])
    brs  = "\n".join(f"  - {r}" for r in structure["rules"][:20])
    return f"""
{PERSONA}

PROJECT: {name}

DETECTED STRUCTURE:
REST Endpoints:
{eps or '  (none detected - infer from code)'}

Domain Entities:
{ents or '  (none detected)'}

Service Operations:
{svcs or '  (none detected)'}

Business Rules Found:
{brs or '  (none detected)'}

FULL SOURCE CODE:
{code}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Generate a complete BUSINESS REQUIREMENTS DOCUMENT (BRD):

## 1. EXECUTIVE SUMMARY
3-4 sentences. What does this system do? What business problem does it solve?

## 2. BUSINESS OBJECTIVES
List 5-6 numbered objectives (BO-01 through BO-06). What business goals does this achieve?

## 3. SCOPE
### 3.1 In Scope
List all business capabilities this system includes.
### 3.2 Out of Scope
List what this system does NOT do (infer from what's missing).

## 4. STAKEHOLDERS & USER ROLES
For each user role detected: role name, who they are, what they can do.

## 5. BUSINESS RULES
Number each rule BR-01, BR-02 etc. Convert ALL validation annotations and
service-layer checks into plain English business rules. Include at least 8.

## 6. BUSINESS WORKFLOWS
For each major process (detected from service methods and endpoints):
- Workflow Name
- Trigger
- Step-by-step numbered process
- Outcome

## 7. NON-FUNCTIONAL REQUIREMENTS
Security, data retention, integrations — based on code evidence.

## 8. ASSUMPTIONS & CONSTRAINTS
What did you assume? What constraints exist?
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Write detailed, professional content. Minimum 1500 words.
"""

def make_frd_prompt(name, structure, code):
    eps  = "\n".join(f"  [{e['method']}] {e['path']} → {e['handler']}()" for e in structure["endpoints"])
    ents = "\n".join(f"  {e['name']}: {[f[1] for f in e['fields'][:8]]}" for e in structure["entities"])
    return f"""
{PERSONA}

PROJECT: {name}

DETECTED REST ENDPOINTS:
{eps or '  (infer from source code)'}

DETECTED ENTITIES:
{ents or '  (infer from source code)'}

FULL SOURCE CODE:
{code}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Generate a complete FUNCTIONAL REQUIREMENTS DOCUMENT (FRD):

## 1. SYSTEM OVERVIEW
Architecture summary, main components, technology stack in business terms.

## 2. FUNCTIONAL MODULES
Table listing all modules, their purpose, and which components implement them.

## 3. DETAILED FUNCTIONAL REQUIREMENTS
For EVERY endpoint/feature write one FR entry:

### FR-[nn]: [Feature Name]
**Description:** What this feature does
**Actors:** Who uses it (which role)
**Inputs:** Required data / parameters
**Processing:** What the system does step by step
**Outputs:** What is returned or what changes
**Business Rules Applied:** BR numbers from BRD
**Validation Rules:** What is checked before processing
**Error Scenarios:** What can fail and what happens
**Priority:** High / Medium / Low

## 4. API CATALOGUE
Table with columns: Endpoint | HTTP Method | Purpose | Access Role | Response

## 5. DATA REQUIREMENTS
### 5.1 Data Entities
For each entity: name, purpose, key attributes, relationships.
### 5.2 Data Validation Rules
Table of all field-level validation rules.

## 6. SECURITY & ACCESS CONTROL
Table showing which roles can perform which operations (✅ / ❌).

## 7. ERROR HANDLING
List all error scenarios, their HTTP codes, and business meaning.
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Be thorough. Cover EVERY endpoint. Minimum 2000 words.
"""


# ──────────────────────────────────────────────────────────────────
# STEP 4 — CALL OPENAI
# ──────────────────────────────────────────────────────────────────
def call_openai(prompt, client, model_name, label, direct_url=None, subscription_key=None):
    """Call the OpenAI-compatible API.

    When *direct_url* is given (e.g. Bosch APIM gateway) a plain
    requests.post() is used so the full URL is hit exactly as supplied.
    Otherwise the standard OpenAI SDK client is used.
    """
    import requests as _requests

    print(f"\n  📡 Sending {label} request to OpenAI ({model_name})...")
    print(f"     Prompt size: ~{len(prompt)//4:,} tokens")

    # Azure OpenAI (Bosch gateway) does not accept "model" in the body —
    # the model is already encoded in the deployment URL.
    base_payload = {
        "messages": [
            {"role": "system", "content": "You are a Senior Business Analyst."},
            {"role": "user",   "content": prompt},
        ],
        "max_tokens": 8192,
        "temperature": 0.1,
    }

    for attempt in range(3):
        try:
            if direct_url and subscription_key:
                # ── Bosch APIM / Azure gateway — direct HTTP call ──────────
                payload = base_payload  # no "model" key for Azure deployments
                headers = {
                    "Content-Type": "application/json",
                    "api-key": subscription_key,
                    "Ocp-Apim-Subscription-Key": subscription_key,
                }
                resp = _requests.post(
                    direct_url, headers=headers, json=payload, timeout=300
                )
                if not resp.ok:
                    print(f"  ❌ Azure response {resp.status_code}: {resp.text}")
                resp.raise_for_status()
                text = resp.json()["choices"][0]["message"]["content"]
            else:
                # ── Standard OpenAI SDK — include model name in body ───────
                payload = {**base_payload, "model": model_name}
                response = client.chat.completions.create(**payload)
                text = response.choices[0].message.content

            print(f"  ✅ Received {len(text):,} characters of content")
            return text

        except Exception as e:
            err = str(e).lower()
            if "rate" in err or "quota" in err or "429" in err:
                wait = (attempt + 1) * 20
                print(f"  ⏳ Rate limit. Waiting {wait}s...")
                time.sleep(wait)
            elif attempt < 2:
                print(f"  ⚠️  Error (attempt {attempt+1}): {e}. Retrying...")
                time.sleep(5)
            else:
                print(f"  ❌ OpenAI call failed: {e}")
                return f"[Generation failed: {e}]"
    return "[Max retries exceeded]"


# ──────────────────────────────────────────────────────────────────
# STEP 5 — WRITE WORD DOCUMENT
# ──────────────────────────────────────────────────────────────────
def shade_cell(cell, hex_color):
    tc = cell._tc; tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), hex_color)
    shd.set(qn("w:val"),  "clear")
    shd.set(qn("w:color"),"auto")
    tcPr.append(shd)

def add_runs(para, text):
    for i, part in enumerate(re.split(r"\*\*(.+?)\*\*", text)):
        if not part: continue
        r = para.add_run(part)
        r.font.name = "Calibri"; r.font.size = Pt(11)
        if i % 2 == 1:
            r.bold = True; r.font.color.rgb = NAVY

def write_table(doc, lines):
    rows = [l for l in lines if not re.match(r"^\|[\s\-|]+\|$", l.strip())]
    parsed = [[c.strip() for c in r.strip("|").split("|")] for r in rows]
    if not parsed: return
    cols = max(len(r) for r in parsed)
    tbl  = doc.add_table(rows=len(parsed), cols=cols)
    tbl.style = "Table Grid"
    for ri, row in enumerate(parsed):
        for ci in range(cols):
            cell = tbl.rows[ri].cells[ci]
            txt  = row[ci] if ci < len(row) else ""
            cell.text = txt
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(3)
            p.paragraph_format.space_after  = Pt(3)
            if ri == 0:
                if p.runs:
                    p.runs[0].bold=True; p.runs[0].font.color.rgb=WHITE
                    p.runs[0].font.size=Pt(9); p.runs[0].font.name="Calibri"
                shade_cell(cell, "0A1628")
            else:
                if p.runs:
                    p.runs[0].font.size=Pt(9); p.runs[0].font.name="Calibri"
                if ri % 2 == 0:
                    shade_cell(cell, "F0F4F8")
    doc.add_paragraph()

def write_doc(content, doc_type, project_name, output_dir):
    os.makedirs(output_dir, exist_ok=True)
    ts   = datetime.now().strftime("%Y%m%d_%H%M")
    safe = project_name.replace(" ","_").replace("/","_")
    path = os.path.join(output_dir, f"{doc_type}_{safe}_{ts}.docx")

    doc = Document()
    for sec in doc.sections:
        sec.page_width=Cm(21); sec.page_height=Cm(29.7)
        sec.left_margin=sec.right_margin=Cm(2.5)
        sec.top_margin=sec.bottom_margin=Cm(2.0)

    doc.styles["Normal"].font.name = "Calibri"
    doc.styles["Normal"].font.size = Pt(11)
    for hname, sz, col in [("Heading 1",16,NAVY),("Heading 2",13,NAVY),("Heading 3",11,TEAL)]:
        s=doc.styles[hname]; s.font.name="Calibri"
        s.font.size=Pt(sz); s.font.bold=True; s.font.color.rgb=col

    # ── Cover page ─────────────────────────────────────────────────
    for _ in range(7): doc.add_paragraph()
    for txt, sz, col, italic in [
        (doc_type,       30, NAVY, False),
        (project_name,   18, TEAL, False),
        ("","",None,False),
        (f"Auto-generated by Code to Clarity  |  {datetime.now().strftime('%B %d, %Y')}", 10, GRAY, True),
        ("Bosch Digital Innovation League 2026  |  CONFIDENTIAL", 9, GRAY, False),
    ]:
        if not txt: doc.add_paragraph(); continue
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(txt)
        r.font.size=Pt(sz); r.font.color.rgb=col
        r.font.name="Calibri"; r.italic=italic
        if sz == 30: r.bold = True

    doc.add_page_break()

    # ── Content ────────────────────────────────────────────────────
    lines = content.split("\n"); i = 0
    while i < len(lines):
        line = lines[i].rstrip()
        if line.startswith("## "):
            doc.add_heading(line[3:].strip(), level=1)
            sp = doc.add_paragraph()
            sp.paragraph_format.space_after = Pt(3)
            sr = sp.add_run("─"*68)
            sr.font.color.rgb=TEAL; sr.font.size=Pt(7)
        elif line.startswith("### "):
            doc.add_heading(line[4:].strip(), level=2)
        elif line.startswith("#### "):
            doc.add_heading(line[5:].strip(), level=3)
        elif line.startswith("|") and "|" in line[1:]:
            tbl = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                tbl.append(lines[i].strip()); i += 1
            write_table(doc, tbl); continue
        elif line.startswith("- ") or line.startswith("* "):
            p = doc.add_paragraph(style="List Bullet")
            add_runs(p, line[2:].strip())
        elif re.match(r"^\d+\.\s", line):
            p = doc.add_paragraph(style="List Number")
            add_runs(p, re.sub(r"^\d+\.\s","",line).strip())
        elif "**" in line:
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(2)
            add_runs(p, line)
        elif not line.strip():
            sp = doc.add_paragraph(); sp.paragraph_format.space_after=Pt(1)
        else:
            p = doc.add_paragraph(); add_runs(p, line)
        i += 1

    # ── Footer ─────────────────────────────────────────────────────
    for sec in doc.sections:
        fp = sec.footer.paragraphs[0]
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        fr = fp.add_run(f"Code to Clarity  |  {project_name}  |  {doc_type}  |  BDIL 2026  |  CONFIDENTIAL")
        fr.font.size=Pt(8); fr.font.color.rgb=GRAY

    doc.save(path)
    return path


# ──────────────────────────────────────────────────────────────────
# MAIN — Orchestrates all steps
# ──────────────────────────────────────────────────────────────────
def main():
    print("""
╔══════════════════════════════════════════════╗
║   CODE TO CLARITY — Starting Agent          ║
║   Bosch Digital Innovation League 2026      ║
╚══════════════════════════════════════════════╝""")

    # Read config
    cfg          = CONFIG
    project_path = sys.argv[1] if len(sys.argv) > 1 else cfg["project_path"]
    project_name = sys.argv[2] if len(sys.argv) > 2 else cfg["project_name"]
    api_key      = cfg["genaiplatform-farm-subscription-key"]
    model_name   = cfg["openai_model"]
    base_url     = cfg["openai_base_url"] or None

    if not api_key:
        print("\n❌  No OpenAI API key found!")
        print("    Option A: Edit CONFIG['openai_api_key'] in this file")
        print("    Option B: set env variable:  set OPENAI_API_KEY=your_key\n")
        sys.exit(1)

    # When a full gateway URL is supplied (Bosch APIM / Azure), we use
    # requests.post() directly and skip the OpenAI SDK's URL logic.
    direct_url  = base_url if base_url and "/chat/completions" in base_url else None
    sub_key     = api_key if direct_url else None

    # Fall-back: standard OpenAI SDK client (used when no gateway URL)
    client = OpenAI(api_key=api_key)
    print(f"\n  ✅ OpenAI client ready ({model_name})")
    if direct_url:
        print(f"     Using Bosch gateway → direct HTTP mode")

    # STEP 1 — Load
    files = load_project(project_path)

    # STEP 2 — Analyse
    print(f"\n{'─'*55}")
    print(f"  🔍 STEP 2 — Analysing code structure")
    print(f"{'─'*55}")
    structure = extract_structure(files)

    # STEP 3 — Prepare code text (group by type, controllers + services first)
    print(f"\n{'─'*55}")
    print(f"  📦 STEP 3 — Preparing code for AI analysis")
    print(f"{'─'*55}")
    priority = ["Controller","Service","Entity","Repository","DTO","Exception","Config","Main","Other"]
    grouped  = {t:[] for t in priority}
    for f in files:
        grouped.setdefault(f["type"],[]).append(f)

    # ~128k token model limit minus 8k response minus ~5k for prompt scaffolding
    # leaves ~115k tokens for code.  At ~4 chars/token → 460,000 chars budget.
    MAX_CODE_CHARS = 460_000
    code_text = ""
    budget    = MAX_CODE_CHARS
    truncated = False

    for t in priority:
        batch = grouped.get(t, [])
        if not batch: continue
        header = f"\n\n{'#'*50}\n# {t.upper()} FILES\n{'#'*50}\n"
        if budget <= 0:
            truncated = True
            break
        code_text += header
        budget -= len(header)
        for f in batch:
            entry = f"\n// FILE: {f['path']}\n{'─'*40}\n{f['content']}\n"
            if len(entry) > budget:
                code_text += entry[:budget] + "\n... [FILE TRUNCATED — token limit reached]\n"
                budget    = 0
                truncated = True
                break
            code_text += entry
            budget -= len(entry)
        if truncated:
            break

    if truncated:
        code_text += "\n\n[NOTE: Some files were omitted to stay within the model token limit. The structure summary above captures all endpoints, entities, and rules.]"

    print(f"  ✅ Code prepared (~{len(code_text)//4:,} tokens){' — truncated to fit token limit' if truncated else ''}")

    # STEP 4 — Generate BRD
    print(f"\n{'─'*55}")
    print(f"  🤖 STEP 4 — Generating BRD via OpenAI")
    print(f"{'─'*55}")
    brd_prompt  = make_brd_prompt(project_name, structure, code_text)
    brd_content = call_openai(brd_prompt, client, model_name, "BRD",
                               direct_url=direct_url, subscription_key=sub_key)

    print(f"\n  ⏳ Waiting 5 seconds before next API call...")
    time.sleep(5)

    # STEP 4b — Generate FRD
    print(f"\n{'─'*55}")
    print(f"  🤖 STEP 4 — Generating FRD via OpenAI")
    print(f"{'─'*55}")
    frd_prompt  = make_frd_prompt(project_name, structure, code_text)
    frd_content = call_openai(frd_prompt, client, model_name, "FRD",
                               direct_url=direct_url, subscription_key=sub_key)

    # STEP 5 — Write documents
    print(f"\n{'─'*55}")
    print(f"  📄 STEP 5 — Writing Word documents")
    print(f"{'─'*55}")
    brd_path = write_doc(brd_content, "BRD", project_name, cfg["output_dir"])
    frd_path = write_doc(frd_content, "FRD", project_name, cfg["output_dir"])

    print(f"""
╔══════════════════════════════════════════════╗
║   ✅  CODE TO CLARITY COMPLETE!             ║
╚══════════════════════════════════════════════╝

  📄 BRD → {brd_path}
  📋 FRD → {frd_path}

  Open the files in Microsoft Word to view them.
""")

if __name__ == "__main__":
    main()
