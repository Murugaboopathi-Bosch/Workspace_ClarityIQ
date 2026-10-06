# 🚀 CODE TO CLARITY — Run It Locally (Step-by-Step Guide)
### Bosch Digital Innovation League 2026

---

## 📦 WHAT'S INSIDE THIS FOLDER

```
code_to_clarity_local/
│
├── code_to_clarity.py       ← THE MAIN AGENT (this is the brain)
├── requirements.txt         ← Python packages needed
├── run_windows.bat          ← Double-click to run on Windows
├── run_mac_linux.sh         ← Run this on Mac or Linux
├── README.md                ← This guide
│
└── sample-java-project/     ← SAMPLE INPUT (a real Spring Boot project)
    └── src/main/java/com/bosch/usermanagement/
        ├── UserManagementApplication.java
        ├── controller/
        │   ├── UserController.java
        │   └── EmployeeController.java
        ├── service/
        │   ├── UserService.java
        │   └── EmployeeService.java
        ├── model/
        │   ├── User.java
        │   ├── Employee.java
        │   └── Department.java
        ├── repository/
        ├── dto/
        └── exception/

output/                      ← YOUR DOCUMENTS APPEAR HERE (created when you run)
    BRD_MyApp_20260222.docx
    FRD_MyApp_20260222.docx
```

---

## ✅ PREREQUISITES — What You Need Before Starting

| Requirement | Check | Download Link |
|---|---|---|
| Python 3.9+ | Open terminal → type `python --version` | https://www.python.org/downloads/ |
| pip (Python package manager) | Comes with Python | (included) |
| Microsoft Word | To open .docx output files | Office 365 / MS Office |
| Gemini API Key | From Google AI Studio (free) | https://aistudio.google.com/app/apikey |
| Internet connection | For Gemini API calls | — |

---

## 🔑 STEP 1 — GET YOUR GEMINI API KEY (Free, 2 minutes)

1. Open your browser → go to: **https://aistudio.google.com/app/apikey**
2. Sign in with your Google account
3. Click **"Create API Key"**
4. Copy the key (looks like: `AIzaSyXXXXXXXXXXXXXXXXXXXXXXXXX`)
5. Keep it safe — you'll need it in Step 3

> **For Bosch internal Gemini gateway:** Contact your AI Platform team
> for the internal API key and endpoint URL.

---

## 💻 STEP 2 — INSTALL PYTHON (Skip if already installed)

### Windows:
1. Go to https://www.python.org/downloads/
2. Download Python 3.11 or higher
3. Run the installer
4. ✅ **IMPORTANT:** Tick the box **"Add Python to PATH"** before clicking Install
5. Click Install Now → wait for completion
6. Open Command Prompt → type `python --version` → should show `Python 3.11.x`

### Mac:
```bash
# If you have Homebrew:
brew install python3

# Verify:
python3 --version
```

### Linux (Ubuntu/Debian):
```bash
sudo apt update
sudo apt install python3 python3-pip
python3 --version
```

---

## 📁 STEP 3 — CONFIGURE THE AGENT

Open `code_to_clarity.py` in any text editor (Notepad, VS Code, etc.)

Find this section near the top of the file (around line 35):

```python
CONFIG = {
    "project_path" : "./sample-java-project",   # ← path to your Java project
    "project_name" : "My Java Application",     # ← name shown in documents
    "gemini_api_key": None,                      # ← PASTE YOUR KEY HERE
    "gemini_model" : "gemini-1.5-pro",
    "output_dir"   : "./output",
}
```

Make these changes:

```python
CONFIG = {
    "project_path" : "./sample-java-project",        # ← keep this to test with sample
    "project_name" : "Bosch User Management System", # ← change to your project name
    "gemini_api_key": "AIzaSyXXXXXXXXXXXXXXXXXX",   # ← PASTE YOUR KEY BETWEEN THE QUOTES
    "gemini_model" : "gemini-1.5-pro",
    "output_dir"   : "./output",
}
```

Save the file.

---

## ▶️  STEP 4 — INSTALL PACKAGES & RUN

### Option A — Windows (Easiest):
1. Open the `code_to_clarity_local` folder in File Explorer
2. Double-click **`run_windows.bat`**
3. A black Command Prompt window opens — watch the progress
4. Done! Output files appear in the `output/` folder

### Option B — Mac or Linux:
```bash
# Open Terminal, navigate to this folder
cd /path/to/code_to_clarity_local

# Make the script runnable (first time only)
chmod +x run_mac_linux.sh

# Run it
./run_mac_linux.sh
```

### Option C — Manual (any OS):
```bash
# 1. Open terminal in this folder

# 2. Install packages (only needed once)
pip install -r requirements.txt

# 3. Run the agent on the sample project
python code_to_clarity.py

# 4. Or with arguments (no need to edit config):
python code_to_clarity.py ./sample-java-project "My Project Name"
```

---

## 👀 STEP 5 — WATCH THE TERMINAL OUTPUT

While running, you will see this in your terminal:

```
╔══════════════════════════════════════════════╗
║   CODE TO CLARITY — Starting Agent          ║
╚══════════════════════════════════════════════╝

  ✅ Gemini ready (gemini-1.5-pro)

───────────────────────────────────────────────
  📂 STEP 1 — Loading Java files
───────────────────────────────────────────────
  ✅ 18 Java files found:

     ████ 4x  DTO
     ███  3x  Entity
     ██   2x  Controller
     ██   2x  Service

───────────────────────────────────────────────
  🔍 STEP 2 — Analysing code structure
───────────────────────────────────────────────
  ✅ 15 REST endpoints detected
  ✅ 30 business rules / validations found

  REST Endpoints:
     [POST  ] /api/users/register         → registerUser()
     [GET   ] /api/users/{id}             → getUserById()
     [PUT   ] /api/employees/{id}/salary  → updateSalary()
     ...

───────────────────────────────────────────────
  📦 STEP 3 — Preparing code for AI
───────────────────────────────────────────────
  ✅ Code prepared (~6,200 tokens)

───────────────────────────────────────────────
  🤖 STEP 4 — Generating BRD via Gemini
───────────────────────────────────────────────
  📡 Sending BRD request to Gemini...
     Prompt size: ~8,400 tokens
  ✅ Received 11,230 characters of content

  ⏳ Waiting 5 seconds before next API call...

───────────────────────────────────────────────
  🤖 STEP 4 — Generating FRD via Gemini
───────────────────────────────────────────────
  📡 Sending FRD request to Gemini...
     Prompt size: ~8,600 tokens
  ✅ Received 14,850 characters of content

───────────────────────────────────────────────
  📄 STEP 5 — Writing Word documents
───────────────────────────────────────────────

╔══════════════════════════════════════════════╗
║   ✅  CODE TO CLARITY COMPLETE!             ║
╚══════════════════════════════════════════════╝

  📄 BRD → ./output/BRD_My_Project_20260222_1430.docx
  📋 FRD → ./output/FRD_My_Project_20260222_1432.docx
```

---

## 📄 STEP 6 — OPEN YOUR DOCUMENTS

1. Go to the `output/` folder
2. You will see two files:
   - `BRD_YourProject_[timestamp].docx` — Business Requirements Document
   - `FRD_YourProject_[timestamp].docx` — Functional Requirements Document
3. Double-click either file to open in Microsoft Word

---

## 🔁 STEP 7 — RUN ON YOUR OWN JAVA PROJECT

Once the sample works, point the agent at your real Java project:

```python
# In code_to_clarity.py, change CONFIG:
CONFIG = {
    "project_path" : "C:/Users/YourName/IdeaProjects/my-real-project",  # Windows
    # OR
    "project_path" : "/home/yourname/projects/my-real-project",          # Mac/Linux
    "project_name" : "My Real Project Name",
    "gemini_api_key": "your-key-here",
    ...
}
```

Then run again:
```bash
python code_to_clarity.py
```

---

## ⚡ PRO TIP — Get Better Output

Add `// Business Rule:` comments in your Java Service classes:

```java
// Business Rule: Salary cannot be reduced below current value
if (newSalary < emp.getSalary()) {
    throw new IllegalArgumentException("...");
}

// Business Rule: Only active employees can be transferred
if (emp.getStatus() != EmploymentStatus.ACTIVE) {
    throw new IllegalStateException("...");
}
```

The agent picks up these comments automatically and converts them
into numbered business rules (BR-01, BR-02...) in your BRD.

---

## 🛠️ TROUBLESHOOTING

| Problem | Solution |
|---|---|
| `python: command not found` | Use `python3` instead of `python` on Mac/Linux |
| `No module named google.generativeai` | Run: `pip install -r requirements.txt` |
| `❌ No Gemini API key found` | Paste your key in CONFIG or run: `export GEMINI_API_KEY="your_key"` |
| `❌ Folder not found` | Use full absolute path in project_path e.g. `C:/Users/name/myproject` |
| `Rate limit error` | Free Gemini tier has limits. Wait 1 minute and run again |
| `Output file is empty` | Check Gemini API key is valid at aistudio.google.com |
| `pip not found` | Run: `python -m pip install -r requirements.txt` |

---

## 📞 HOW LONG DOES IT TAKE?

| Project Size | Files | Time |
|---|---|---|
| Small (sample project) | ~18 files | ~30–60 seconds |
| Small (your project) | < 50 files | ~1–2 minutes |
| Medium | 50–200 files | ~3–5 minutes |

---

*Code to Clarity — Bosch Digital Innovation League 2026*
