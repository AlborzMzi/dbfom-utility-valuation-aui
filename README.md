# DBFOM Utility Valuation — Aurora Utilities Inc. (AUI)

Independent DBFOM utility valuation: wastewater reclamation facility — **Excel model**, **two Word briefs**, and a **reproducible Python build script**.

---

## How to run

You can rebuild the Excel model and Word briefs either in **GitHub Codespaces** (browser-only) or **locally**.

### Option A — GitHub Codespaces (no local installs)
1. Click the green **Code** button → **Codespaces** tab → **Create codespace on main**.
2. In the Codespaces terminal, run:
   ```bash
   pip install -r requirements.txt
   python src/generate_model.py
   ```
3. The script writes:
   - `models/AUI_DBFO_Wastewater_Model.xlsx`
   - `docs/Project_Description.docx`
   - `docs/Executive_Summary.docx`
4. Use the **Source Control** panel to **Commit** and **Push** back to GitHub.

### Option B — Run locally
1. Ensure **Python 3.10+** is installed.
2. From the repo folder, run:
   ```bash
   python -m venv .venv
   # macOS/Linux
   source .venv/bin/activate
   # Windows PowerShell
   .venv\Scripts\Activate.ps1

   pip install -r requirements.txt
   python src/generate_model.py
   ```
3. Outputs will appear under `models/` and `docs/` as listed above.

---

## Download the deliverables

- **Excel model:** [models/AUI_DBFO_Wastewater_Model.xlsx](models/AUI_DBFO_Wastewater_Model.xlsx)  
- **Project Description (DOCX):** [docs/Project_Description.docx](docs/Project_Description.docx)  
- **Executive Summary (DOCX):** [docs/Executive_Summary.docx](docs/Executive_Summary.docx)

---

## Repo structure
```text
.
├─ models/                      # Excel model (generated)
├─ docs/                        # Project Description + Executive Summary (generated)
├─ src/
│  └─ generate_model.py         # Rebuilds the model and briefs
├─ requirements.txt             # xlsxwriter, python-docx
└─ README.md
```

## Notes
- Revenue recognition: City **interest** on receivable + O&M **markup**; principal flows through **cash**, not revenue.
- EAR → monthly: `(1 + r)^(1/12) - 1`. All inputs in 2025 dollars.
- Educational/portfolio project; no proprietary data.

## Optional polish
- In the repo **About** panel, add topics: `project-finance`, `valuation`, `excel-model`, `dbfom`, `utilities`.
- Create a **Release** (e.g., `v1.0.0`) for a stable snapshot you can link on your CV.
- Pin this repo on your profile (Profile → **Customize your pins**).
