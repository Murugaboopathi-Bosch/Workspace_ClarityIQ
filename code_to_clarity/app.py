"""
╔══════════════════════════════════════════════════════════════════╗
║         CODE TO CLARITY — Flask Web UI                         ║
║         Upload a Java project folder → Download BRD + FRD      ║
╚══════════════════════════════════════════════════════════════════╝
"""

import os, sys, io, time, zipfile, shutil, tempfile
from flask import Flask, render_template, request, send_file, jsonify

# ── Import core logic from code_to_clarity.py ─────────────────────
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from code_to_clarity import (
    load_project,
    extract_structure,
    make_brd_prompt,
    make_frd_prompt,
    call_openai,
    write_doc,
    CONFIG,
)
from openai import OpenAI

# ══════════════════════════════════════════════════════════════════
app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 2 * 1024 * 1024 * 1024  # 2 GB (only .java files reach the server)


# ── Helpers ────────────────────────────────────────────────────────
# ~128k token model limit minus 8k response minus ~5k for prompt scaffolding
# leaves ~115k tokens for code.  At ~4 chars/token → 460,000 chars budget.
_MAX_CODE_CHARS = 460_000

def _build_code_text(files):
    priority = [
        "Controller", "Service", "Entity", "Repository",
        "DTO", "Exception", "Config", "Main", "Other",
    ]
    grouped = {t: [] for t in priority}
    for f in files:
        grouped.setdefault(f["type"], []).append(f)

    code_text = ""
    budget = _MAX_CODE_CHARS
    truncated = False

    for t in priority:
        batch = grouped.get(t, [])
        if not batch:
            continue
        header = f"\n\n{'#'*50}\n# {t.upper()} FILES\n{'#'*50}\n"
        if budget <= 0:
            truncated = True
            break
        code_text += header
        budget -= len(header)
        for f in batch:
            entry = f"\n// FILE: {f['path']}\n{'─'*40}\n{f['content']}\n"
            if len(entry) > budget:
                # Include a truncated version of this file
                snippet = entry[:budget]
                code_text += snippet + "\n... [FILE TRUNCATED — token limit reached]\n"
                budget = 0
                truncated = True
                break
            code_text += entry
            budget -= len(entry)
        if truncated:
            break

    if truncated:
        code_text += "\n\n[NOTE: Some files were omitted to stay within the model token limit. The structure summary above captures all endpoints, entities, and rules.]"

    return code_text


def _init_openai():
    """Build the OpenAI client + resolve direct gateway URL from CONFIG."""
    cfg        = CONFIG
    api_key    = (cfg.get("genaiplatform-farm-subscription-key", "")
                  or os.environ.get("OPENAI_API_KEY", ""))
    model_name = cfg.get("openai_model", "gpt-4o")
    base_url   = cfg.get("openai_base_url", "") or None
    direct_url = base_url if base_url and "/chat/completions" in base_url else None
    sub_key    = api_key if direct_url else None
    client     = OpenAI(api_key=api_key)
    return client, model_name, direct_url, sub_key


# ── Error Handlers ───────────────────────────────────────────────
@app.errorhandler(413)
def too_large(e):
    return jsonify(error="Upload too large (max 2 GB). Your browser may be sending non-Java files. Please try again — the page now filters only .java files before uploading."), 413


# ── Routes ─────────────────────────────────────────────────────────
@app.route("/")
def index():
    return render_template("index.html")


@app.route("/generate", methods=["POST"])
def generate():
    project_name   = (request.form.get("project_name") or "My Project").strip()
    uploaded_files = request.files.getlist("project_folder")

    if not uploaded_files or uploaded_files[0].filename == "":
        return jsonify(error="No files uploaded. Please select a project folder."), 400

    tmp_project = tempfile.mkdtemp(prefix="ctc_proj_")
    tmp_output  = tempfile.mkdtemp(prefix="ctc_out_")

    try:
        # ── Save only .java files, preserving relative path ───────
        saved = 0
        for file in uploaded_files:
            fname = file.filename.replace("\\", "/")
            if not fname.lower().endswith(".java"):
                continue
            parts = fname.split("/")
            # Browser prepends the top-level folder name — strip it
            rel  = "/".join(parts[1:]) if len(parts) > 1 else parts[0]
            dest = os.path.join(tmp_project, rel)
            os.makedirs(os.path.dirname(dest), exist_ok=True)
            file.save(dest)
            saved += 1

        if saved == 0:
            return jsonify(error="No .java files found in the uploaded folder."), 400

        # ── Pipeline ──────────────────────────────────────────────
        files     = load_project(tmp_project)
        if not files:
            return jsonify(error="No Java source files could be read."), 400

        structure = extract_structure(files)
        code_text = _build_code_text(files)

        client, model_name, direct_url, sub_key = _init_openai()

        brd_content = call_openai(
            make_brd_prompt(project_name, structure, code_text),
            client, model_name, "BRD",
            direct_url=direct_url, subscription_key=sub_key,
        )
        time.sleep(3)
        frd_content = call_openai(
            make_frd_prompt(project_name, structure, code_text),
            client, model_name, "FRD",
            direct_url=direct_url, subscription_key=sub_key,
        )

        brd_path = write_doc(brd_content, "BRD", project_name, tmp_output)
        frd_path = write_doc(frd_content, "FRD", project_name, tmp_output)

        # ── Bundle into in-memory ZIP ──────────────────────────────
        safe    = project_name.replace(" ", "_").replace("/", "_")
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.write(brd_path, os.path.basename(brd_path))
            zf.write(frd_path, os.path.basename(frd_path))
        zip_buf.seek(0)

        return send_file(
            zip_buf,
            as_attachment=True,
            download_name=f"{safe}_BRD_FRD.zip",
            mimetype="application/zip",
        )

    except SystemExit:
        return jsonify(error="Project folder could not be read. Ensure it contains .java files."), 400
    except Exception as e:
        return jsonify(error=str(e)), 500
    finally:
        shutil.rmtree(tmp_project, ignore_errors=True)
        shutil.rmtree(tmp_output,  ignore_errors=True)


# ══════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    print("""
╔══════════════════════════════════════════════╗
║   CODE TO CLARITY — Web UI Starting         ║
╚══════════════════════════════════════════════╝
  Open http://localhost:8080 in your browser
""")
    app.run(debug=False, host="0.0.0.0", port=8080)
