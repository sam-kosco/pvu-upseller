# PVU Upseller — Claude Code Reference

This file gives you everything you need to understand and continue working on the PVU Upseller project. It lives at the root of the git clone `C:\Users\samko\pvu-upseller\` as `CLAUDE.md` (tracked in the repo) so Claude Code picks it up automatically as project context.

---

## Project Overview

**What it does:** When a tech at Duncan Aviation PVU (Provo, Utah) completes an aircraft detailing assessment via JotForm, this automation generates a professional upsell sales proposal (PDF + Word doc) and emails it to Stephen Chadbourn (GM, PVU) with the Word doc linked for editing. It also logs every submission to a tracking spreadsheet.

**Who uses it:**
- **Techs** — fill out the JotForm after inspecting an aircraft
- **Stephen Chadbourn** — receives the generated proposal, edits if needed, sends to customer
- **Sam Kosco** — sole data analyst; maintains this automation

**Business context:** FoxTrot Aviation is an aircraft detailing company (~500 employees). PVU is one location. Services offered: Private Detailing, Commercial Detailing, MRO (CIC Removal), and Facilities. This automation covers the Private Detailing upsell workflow.

---

## Repository

**Local clone (edit here):** `C:\Users\samko\pvu-upseller\` — this is the only working copy; edit and commit here.
**GitHub:** `https://github.com/sam-kosco/pvu-upseller`
**Branch:** `main`

### File Structure

```
pvu-upseller/
├── main.py                      # Entry point — called by GitHub Actions
├── parse_payload.py             # Parses raw JotForm JSON into structured dict
├── generate_upsell_pdf.py       # Generates the PDF proposal (ReportLab) + all shared config
├── generate_upsell_docx.py      # Generates the Word doc proposal (docx-js via Node)
├── photo_layout.py              # PIL photo compositor (1–5 images → grid layout)
├── service_context.md           # Knowledge base for Claude AI note rewriter
├── logo.png                     # Foxtrot Aviation logo (used in both outputs)
└── .github/workflows/
    └── generate_upsell_doc.yml  # GitHub Actions workflow definition
```

---

## Full Pipeline

```
JotForm submission
      │
      ▼
Power Automate (PVU Upseller flow)
  Step 1: Add row to tracker spreadsheet (immediate)
  Step 2: Save raw payload JSON to SharePoint staging folder
  Step 3: POST to GitHub Actions API with submission ID
  Step 4: Wait 3 minutes (for GitHub Actions to complete)
  Step 5: Get SharePoint metadata for the generated DOCX
  Step 6: Create anonymous sharing link for the DOCX
  Step 7: Get PDF file content (for email attachment)
  Step 8: Update tracker row with the DOCX sharing link
  Step 9: Send email to Sam + Stephen with PDF attached and DOCX link
      │
      ▼
GitHub Actions (generate_upsell_doc.yml)
  - Installs Python deps (reportlab, pillow, requests, anthropic, numpy)
  - Installs Node deps (docx)
  - Runs: python main.py
      │
      ▼
main.py
  1. Reads SUBMISSION_ID from env
  2. Gets Microsoft Graph API token (client credentials)
  3. Downloads payload JSON from SharePoint staging
  4. Calls parse_payload() → structured data dict
  5. Calls generate_pdf() → saves PDF to temp
  6. Calls generate_docx() → saves DOCX to temp
  7. Uploads PDF to SharePoint output folder
  8. Uploads DOCX to SharePoint output folder
  9. Deletes staging payload JSON (cleanup)
```

---

## Key Files Explained

### `parse_payload.py`
Converts the raw JotForm/Power Automate JSON body into a clean Python dict.

**JotForm field mapping:**

| Q# | Field |
|----|-------|
| 3 | Technician name |
| 4 | Tail number (auto-uppercased) |
| 5 | Indoc date |
| 6 | RTS date |
| 7 | Included services (semicolon-separated) |
| 8 | Customer/operator |
| 47 | Aircraft make |
| 48 | Aircraft model |
| 54 | Owner name |
| 55 | CIC brand — "Xylon" or "Skyde Clear" (display-only; defaults to "Xylon" when blank) |

**Upsell service field pattern** (Yes/No → Photos → Price → Notes):

| Service | Y/N | Photos | Price | Notes |
|---------|-----|--------|-------|-------|
| Brightwork | 10 | 11 | 26 | 27 |
| Ceramic Coating | 12 | 13 | 29 | 30 |
| Permagard Coating | 14 | 15 | 32 | 33 |
| Polymer Coating (displays as Nu-Glaze) | 16 | 17 | 35 | 36 |
| Interior Detail | 18 | 19 | 38 | 39 |
| Exterior Detail | 20 | 21 | 41 | 42 |
| Carpet Extraction | 22 | 23 | 44 | 45 |
| Xylon | 49 | 53 | 50 | 51 |

**Important:** The JotForm and tracker spreadsheet still use "Polymer Coating" internally. The rename to "Nu-Glaze" is display-only, controlled by `DISPLAY_NAMES` in `generate_upsell_pdf.py`.

**CIC brand (q55):** The corrosion-inhibiting compound is keyed internally as `"Xylon"` everywhere (SERVICE_MAP, tracker `XY Upsell`). Field 55 picks the **display name only** — "Xylon" or "Skyde Clear" — stored as `data["cic_type"]`. The information for the two brands is identical; only the name shown in the section title and body changes. Because there is no confirmed product/model name for Skyde Clear, the AI is told never to invent one (no "Skyde Clear 3", no "Skyde Clear Aircraft Exterior Protector") — see `CIC_BRANDS_WITHOUT_PRODUCT_NAME` in `generate_upsell_pdf.py`.

**Output dict structure:**
```python
{
    "tail": "N370EL",           # always uppercase
    "make": "Bombardier",
    "model": "Challenger 300",
    "customer": "Jet Linx",     # q8
    "owner": "Jet Linx Aviation", # q54
    "cic_type": "Xylon",        # q55 ("Xylon" or "Skyde Clear"); defaults to "Xylon"
    "technician": "Sam Kosco",
    "indoc_date": "2026-04-17",
    "rts_date": "2026-04-30",
    "included_services": ["Brightwork", "Ceramic Coating"],
    "upsells": [
        {
            "service": "Interior Detail",
            "price": "850",
            "notes": "seats look dirty",  # raw, gets AI-rewritten
            "photos": [<PIL Image>, ...]   # decoded from base64 in payload
        }
    ]
}
```

**Important:** JotForm base64-encodes photos directly in the payload (`entry["file"]`). No API auth is needed to retrieve photos — they're already in the JSON.

---

### `generate_upsell_pdf.py`
Builds the PDF proposal using ReportLab. Key design decisions:

- **Three supersections (in render order):** Paint Correction & Protective Coatings, Metal Polish & Protection, Detail Work. Order is controlled solely by the order of the `SUPERSECTIONS` list — both the PDF and DOCX iterate it, so reordering the list reorders both outputs
- **Per-supersection boilerplate:** Fixed text at the top of each section
- **Supersection example photos:** Before/after photos appear at the supersection level (not per-service). Configured via `SUPERSECTION_EXAMPLE_PHOTOS` dict
- **No per-service boilerplate:** `SERVICE_BOILERPLATE` is empty. The AI paragraph covers both scope of work and observed condition
- **AI note rewriting:** Calls Claude API (`claude-sonnet-4-6`) to write 2–3 concise sentence paragraphs that explain the scope of work and describe the specific aircraft condition. Uses `service_context.md` as system prompt and condition photos (vision)
- **Display name mapping:** `DISPLAY_NAMES = {"Polymer Coating": "Nu-Glaze"}` — renames services in PDF/DOCX output without changing JotForm or tracker. A **per-submission** override is layered on top for the CIC brand: `generate_pdf()` / `_build_doc_data()` copy `DISPLAY_NAMES` and set `display_names["Xylon"] = data["cic_type"]` (q55)
- **Fixed warranty statements:** `WARRANTY_NOTES` maps a service to an always-on sentence appended (bold) after its AI paragraph. Ceramic Coating always states a confident **3-year warranty regardless of flight hours**. The AI is told NOT to mention any warranty so this is the single source of truth
- **Layout:** Logo centered page 1 (canvas callback), small right-aligned info header every page, centered footer every page, `KeepTogether` wraps all photo blocks so captions never orphan

**Key config dicts in `generate_upsell_pdf.py`:**

```python
DISPLAY_NAMES = {"Polymer Coating": "Nu-Glaze"}

# Always-on, bolded warranty sentences appended after the AI paragraph.
# AI is instructed not to mention warranty, so these are authoritative.
WARRANTY_NOTES = {
    "Ceramic Coating": "This ceramic coating is backed by a full 3-year warranty, "
                       "regardless of flight hours.",
}

# CIC display brands (q55) with no confirmed product name — AI must not
# invent a model/product-line name for these (e.g. no "Skyde Clear 3").
CIC_BRANDS_WITHOUT_PRODUCT_NAME = {"Skyde Clear"}

SUPERSECTION_EXAMPLE_PHOTOS = {
    "Metal Polish & Protection": ("Brightwork Before.jpeg", "Brightwork After.jpeg"),
    "Paint Correction & Protective Coatings": ("Paint Coating Before.jpeg", "Paint Coating After.jpeg"),
    "Detail Work": ("Carpet Extraction Before.jpg", "Carpet Extraction After.jpg"),
}

INCLUDE_EXAMPLE_PHOTOS = {  # all False — photos now at supersection level
    "Brightwork": False, "Ceramic Coating": False, "Permagard Coating": False,
    "Polymer Coating": False, "Interior Detail": False, "Exterior Detail": False,
    "Carpet Extraction": False, "Xylon": False,
}

SERVICE_BOILERPLATE = {}  # empty — AI handles everything
```

**To add a display name override:** Add to `DISPLAY_NAMES` dict. Internal name stays the same everywhere else.

**To update supersection intro text:** Edit `SUPERSECTIONS[n]["boilerplate"]` in `generate_upsell_pdf.py`.

**To change supersection example photos:** Edit `SUPERSECTION_EXAMPLE_PHOTOS` — keys are supersection titles, values are `(before_filename, after_filename)` tuples matching files on SharePoint.

---

### `generate_upsell_docx.py`
Builds the Word doc using a Node.js docx-js script embedded as a Python string. Key points:

- Imports all config from `generate_upsell_pdf.py` (single source of truth for service lists, display names, supersection photos, etc.)
- Logo appears as centered body element on page 1 (not in header, so it's editable in Word)
- Runs the AI rewriter independently (same prompt, separate API call)
- Writes temp image files → passes paths to Node.js → Node builds the docx → Python cleans up temp files
- Uses `docx` npm package (`npm install -g docx`)

---

### `service_context.md`
The AI knowledge base. Loaded at startup by `generate_upsell_pdf.py` as the Claude API system prompt. Contains detailed information about:
- Brightwork (what it is, oxidation chemistry, product systems used)
- Xylon (corrosion inhibitor, fuel savings, UV resistance)
- Ceramic Coating (paint correction process, 3-year warranty, hydrophobic)
- Permagard (industry standard, booster requirement)
- Nu-Glaze (formerly Polymer Coating — higher quality polymer sealant, same price/warranty)
- Interior Detail (leather care, cockpit protocols, materials used)
- Exterior Detail (degreasing, wash wax, what gets cleaned)
- Carpet Extraction (hot water extraction, pre-spray, AquaPro Vac)

Stephen can edit this file freely — changes take effect on the next run with no code changes.

---

### `generate_upsell_doc.yml`
GitHub Actions workflow. Triggered via `workflow_dispatch` with input `submission_id`.

**GitHub Secrets required:**
| Secret | Description |
|--------|-------------|
| `TENANT_ID` | `ede0c57f-549f-4a90-9f8c-7ea130346f95` |
| `CLIENT_ID` | `58191600-ab56-4141-bff6-806805fcbff4` |
| `CLIENT_SECRET` | From Entra → Foxtrot Report Automation app |
| `ANTHROPIC_API_KEY` | From console.anthropic.com |

---

## SharePoint File Locations

All paths are on the **DataHub** SharePoint site.

| Description | SharePoint URL | Local Mirror (OneDrive sync) |
|-------------|---------------|------------------------------|
| Tracker spreadsheet | `Shared Documents/Report Automation/PVU/PVU Upseller.xlsx` | `C:\Users\samko\Foxtrot Aviation Services\Data Hub - Documents\Report Automation\PVU\PVU Upseller.xlsx` |
| Service example photos | `Shared Documents/Assets/Service Example Photos/` | `C:\Users\samko\Foxtrot Aviation Services\Data Hub - Documents\Assets\Service Example Photos\` |
| Generated PDFs | `Shared Documents/Flow Dumps/PVU Upsell PDFs/` | `C:\Users\samko\Foxtrot Aviation Services\Data Hub - Documents\Flow Dumps\PVU Upsell PDFs\` |
| Payload staging (auto-deleted) | `Shared Documents/Flow Dumps/PVU Upsell Payloads/` | `C:\Users\samko\Foxtrot Aviation Services\Data Hub - Documents\Flow Dumps\PVU Upsell Payloads\` |

**SharePoint Drive ID:** `b!_bzXaIx86kOufgJN3ih-BaDIDthKYuxJkJtLi1Bm5irGjCEnK-VHSpBRRm3_SDKU`

**Supersection example photo filenames** (must match exactly on SharePoint):
- `Brightwork Before.jpeg` / `Brightwork After.jpeg` — used for Metal Polish & Protection
- `Paint Coating Before.jpeg` / `Paint Coating After.jpeg` — used for Paint Correction & Protective Coatings
- `Carpet Extraction Before.jpg` / `Carpet Extraction After.jpg` — used for Detail Work

---

## Tracker Spreadsheet (`PVU Upseller.xlsx`)

**Location:** `C:\Users\samko\Foxtrot Aviation Services\Data Hub - Documents\Report Automation\PVU\PVU Upseller.xlsx`

Power Automate writes a row on each submission with:

| Column | Source | Logic |
|--------|--------|-------|
| Tail Number | q4 | Auto-uppercased |
| Make Model | q47 + q48 | Concatenated |
| Indoc Date | q5 | |
| RTS Date | q6 | |
| Form Date | PA `utcNow()` | Pacific time, MM/dd/yyyy |
| BW Upsell | q7 + q10/q26 | "Included" / price / "No attempt" |
| XY Upsell | q7 + q49/q50 | Same pattern |
| CC Upsell | q7 + q12/q29 | Same pattern |
| PGC Upsell | q7 + q14/q32 | Same pattern |
| PC Upsell | q7 + q16/q35 | Same pattern (still "PC" in tracker — displays as Nu-Glaze in PDF) |
| ID Upsell | q7 + q18/q38 | Same pattern |
| ED Upsell | q7 + q20/q41 | Same pattern |
| Sub ID | submissionID | Prefixed with `_` to prevent Excel number formatting |
| Raw Link | DOCX sharing URL | Added in second PA step after GitHub Actions completes |

**Logic for each upsell column:**
- If the service appears in q7 (already included): `"Included"`
- If Y/N field = "Yes": price from the corresponding price field
- Otherwise: `"No attempt"`

---

## Power Automate Flow

**Flow name:** PVU Upseller
**Trigger:** JotForm Enterprise — "When a response is submitted (v2)" on form ID `261044752191049`

**Step order:**
1. `Add_a_row_into_a_table` — writes tracker row immediately (runs in parallel with nothing)
2. `Create_file` — saves `[submissionID].json` to SharePoint staging (runs after step 1)
3. `HTTP` — POSTs to GitHub Actions API with `submission_id` (runs after step 2)
4. `Delay` — waits 3 minutes for GitHub Actions to finish
5. `Get_file_metadata_using_path` — gets PDF metadata from SharePoint
6. `Get_file_metadata_using_path_1` — gets DOCX metadata from SharePoint
7. `Create_sharing_link_for_a_file_or_folder` — creates anonymous edit link for DOCX
8. `Get_file_content_using_path` — downloads PDF bytes for email attachment
9. `Update_a_row` — adds DOCX sharing link back to tracker row
10. `Send_an_email_(V2)` — emails Sam + Stephen with PDF attached, DOCX linked

**Email recipients:** `samuel.kosco@foxtrotaviation.com; stephen.chadbourn@foxtrotaviation.com`
**Email subject:** `PVU Upsell Document for [Make Model] ([TAIL])`

**GitHub API endpoint:**
```
POST https://api.github.com/repos/sam-kosco/pvu-upseller/actions/workflows/generate_upsell_doc.yml/dispatches
```

**Note:** The 3-minute delay is a fixed wait. If GitHub Actions consistently takes longer (e.g. when there are many photos and API calls), increase the delay in the PA flow.

---

## Azure / Entra App Registration

**App name:** Foxtrot Report Automation
**Tenant ID:** `ede0c57f-549f-4a90-9f8c-7ea130346f95`
**Client ID:** `58191600-ab56-4141-bff6-806805fcbff4`
**Client secret:** Stored as `CLIENT_SECRET` GitHub secret (rotate annually in Entra)
**Scope used:** `https://graph.microsoft.com/.default` (client credentials flow)
**Permissions needed:** SharePoint read/write via Graph API

---

## Output File Naming

Both output files are named `[TAIL]_[submissionID]` and saved to `Flow Dumps/PVU Upsell PDFs/`:
- `N370EL_6521254158645395715.pdf`
- `N370EL_6521254158645395715.docx`

The tail is always uppercase (enforced in `parse_payload.py`).

---

## AI Integration

**Provider:** Anthropic Claude API (`claude-sonnet-4-6`)
**Key location:** `ANTHROPIC_API_KEY` GitHub secret / `console.anthropic.com`
**Cost estimate:** ~$0.01–0.02/month at 3–4 submissions/week

**What the AI does:** For each upsell service, the AI writes a 2–3 sentence (concise, not wordy) professional paragraph that:
1. Briefly explains the scope of work (what the service involves)
2. Describes the specific condition observed on this aircraft
3. Explains why the service is recommended

It receives:
- `service_context.md` as system prompt (Foxtrot's knowledge base)
- Stephen's raw field note
- Up to 3 condition photos from the JotForm (vision capability)
- The display name for the service (e.g. "Nu-Glaze" not "Polymer Coating", or the CIC brand from q55)

**Instruction guards added in `rewrite_notes()`:**
- If the resolved display name differs from the internal name, the AI is told to use the display name and never the internal name
- If `avoid_product_names` is set (CIC brand in `CIC_BRANDS_WITHOUT_PRODUCT_NAME`, e.g. Skyde Clear), the AI is told not to invent/borrow any product or model name
- If the service has a `WARRANTY_NOTES` entry, the AI is told not to mention any warranty/coverage — the fixed bold sentence is appended separately

**Where to update the AI prompt:** The `rewrite_notes()` function in `generate_upsell_pdf.py`. The system prompt is `SERVICE_CONTEXT` (loaded from `service_context.md`). The user message is built inline in `rewrite_notes()`.

---

## Common Maintenance Tasks

### Add a new upsell service
1. Add JotForm fields and note the Y/N, photos, price, notes question numbers
2. Add to `SERVICE_MAP` in `parse_payload.py`
3. Add to the appropriate `SUPERSECTIONS` entry's `"services"` list
4. Add to `INCLUDE_EXAMPLE_PHOTOS` dict (set to `False` — photos are at supersection level now)
5. If it needs a display name different from the internal name, add to `DISPLAY_NAMES`
6. Add a section to `service_context.md`
7. Update the Power Automate tracker row logic for the new column

### Rename a service in customer-facing output
Add an entry to `DISPLAY_NAMES` in `generate_upsell_pdf.py`. No other changes needed — the internal name stays the same in JotForm, `parse_payload.py`, and the tracker. The DOCX generator imports `DISPLAY_NAMES` automatically.

### Update supersection intro text
Edit the `"boilerplate"` value inside the relevant entry in `SUPERSECTIONS` in `generate_upsell_pdf.py`.

### Reorder supersections
Change the order of the entries in the `SUPERSECTIONS` list in `generate_upsell_pdf.py`. Both the PDF and DOCX iterate the list in order, so this single change reorders both outputs. (Current order: Paint Correction & Protective Coatings → Metal Polish & Protection → Detail Work.)

### Change or add a fixed warranty statement
Edit `WARRANTY_NOTES` in `generate_upsell_pdf.py` (keyed by internal service name). The sentence is appended **bold** after the AI paragraph in both outputs, and the AI is automatically told not to mention warranty for that service. `generate_upsell_docx.py` imports `WARRANTY_NOTES` — don't duplicate it.

### Add a CIC brand option (q55)
The CIC brand is display-only (`data["cic_type"]`, internal service stays "Xylon"). To support a new brand value coming from JotForm field 55: it will flow through automatically as the display name. If the new brand has **no** confirmed product/model name, add it to `CIC_BRANDS_WITHOUT_PRODUCT_NAME` in `generate_upsell_pdf.py` so the AI won't invent one, and add a short equivalence note to `service_context.md`.

### Change supersection example photos
Edit `SUPERSECTION_EXAMPLE_PHOTOS` in `generate_upsell_pdf.py`. Upload matching filenames to SharePoint `Assets/Service Example Photos/`.

### Update the AI knowledge base
Edit `service_context.md` directly. Takes effect on next run.

### Replace example photos
Upload new files to SharePoint `Assets/Service Example Photos/` with filenames matching `SUPERSECTION_EXAMPLE_PHOTOS`. OneDrive sync will push to SharePoint automatically.

### Rotate the GitHub PAT
The PAT has `workflow` scope only. When it expires:
1. Generate a new classic PAT at `github.com → Settings → Developer settings → Personal access tokens`
2. Scope: `workflow` only
3. Update the `Authorization` header in the Power Automate HTTP action

### Increase the PA wait time
If GitHub Actions is taking more than 3 minutes (can happen with many photos + multiple AI calls), edit the `Delay` step in the Power Automate flow and increase from 3 minutes.

---

## Dependencies

### Python
```
reportlab      # PDF generation
pillow         # Image processing
requests       # Graph API calls / SharePoint photo fetching
anthropic      # Claude API (note rewriting)
numpy          # Photo manipulation helpers
```

### Node.js
```
docx           # Word document generation (npm install -g docx)
```

### GitHub Actions runner
- Ubuntu latest
- Python 3.11
- Node.js (pre-installed on ubuntu-latest)

---

## Tips for Working on This Project in Claude Code

### 1. Start every session by reading key files
Claude Code doesn't retain memory between sessions. At the start of a session, either reference this file explicitly or ask Claude Code to read the files it needs:
```
Read generate_upsell_pdf.py and parse_payload.py before we start
```

### 2. Test PDF/DOCX changes locally before pushing
The smoke test at the bottom of `generate_upsell_pdf.py` (`if __name__ == "__main__":` block) generates a test document with mocked AI and SharePoint calls:
```bash
cd "C:\Users\samko\pvu-upseller"
python generate_upsell_pdf.py
```
Output goes to the system temp directory (e.g. `%TEMP%\upsell_test.pdf`).

### 3. Keep secrets out of the repo
Never put `CLIENT_SECRET`, `ANTHROPIC_API_KEY`, or the GitHub PAT in any file committed to the repo. They live only in GitHub Secrets. The local `pvu.env` file is for local testing only — do not commit it.

### 4. `generate_upsell_docx.py` imports from `generate_upsell_pdf.py`
All shared config (service lists, display names, supersection photos, Stephen's contact info, etc.) lives in `generate_upsell_pdf.py`. `generate_upsell_docx.py` imports from it. If Claude Code edits config, it should always edit `generate_upsell_pdf.py` — never duplicate values into the docx file.

### 5. The JS builder script is a Python string
The Node.js docx builder lives as the `DOCX_BUILDER_JS` string constant inside `generate_upsell_docx.py`. Claude Code will need to treat it carefully — it's JavaScript embedded in Python. When editing it, make sure to preserve the raw string delimiters (`r"""..."""`).

### 6. Power Automate flow changes require export + re-import
The PA flow is exported to `PVU Upseller/Microsoft.Flow/flows/.../definition.json`. If you change the flow in Power Automate, re-export and replace this file to keep the repo in sync. Claude Code cannot push changes directly to Power Automate — changes must be made in the Power Automate portal.

### 7. After any changes to Python files, run the smoke test
```bash
python generate_upsell_pdf.py   # generates test PDF in %TEMP%
python -m py_compile main.py    # syntax check without running
python -m py_compile generate_upsell_docx.py
```
