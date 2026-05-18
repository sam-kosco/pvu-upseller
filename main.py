"""
PVU Upsell Document Automation — Entry Point
=============================================
Called by GitHub Actions when Power Automate POSTs a JotForm submission.

Flow:
  1. Read the JotForm payload from the environment variable JOTFORM_PAYLOAD
  2. Obtain a Microsoft Graph API token using client credentials
  3. Parse the payload into structured data
  4. Generate the upsell PDF
  5. Upload the PDF to SharePoint:
       Shared Documents / Flow Dumps / PVU Upsell PDFs / [tail]_[submissionID].pdf

Environment variables (set as GitHub Secrets):
  TENANT_ID         - Azure AD tenant ID
  CLIENT_ID         - Foxtrot Report Automation app client ID
  CLIENT_SECRET     - Foxtrot Report Automation app client secret
  ANTHROPIC_API_KEY - Anthropic API key (used by generate_upsell_pdf.py)
  JOTFORM_PAYLOAD   - The full JSON payload from Power Automate (set by workflow YAML)
"""

import os
import json
import requests
import tempfile

from parse_payload import parse_payload
from generate_upsell_pdf import generate_pdf


# ─────────────────────────────────────────────────────────────
# SharePoint config
# ─────────────────────────────────────────────────────────────
DRIVE_ID      = "b!_bzXaIx86kOufgJN3ih-BaDIDthKYuxJkJtLi1Bm5irGjCEnK-VHSpBRRm3_SDKU"
OUTPUT_FOLDER = "Flow Dumps/PVU Upsell PDFs"


# ─────────────────────────────────────────────────────────────
# Graph API authentication
# ─────────────────────────────────────────────────────────────
def get_graph_token(tenant_id: str, client_id: str, client_secret: str) -> str:
    """Obtain a Microsoft Graph API bearer token via client credentials flow."""
    url  = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    data = {
        "grant_type":    "client_credentials",
        "client_id":     client_id,
        "client_secret": client_secret,
        "scope":         "https://graph.microsoft.com/.default",
    }
    resp = requests.post(url, data=data, timeout=30)
    resp.raise_for_status()
    token = resp.json().get("access_token")
    if not token:
        raise RuntimeError(f"No access_token in response: {resp.text}")
    print("Graph token obtained.")
    return token


# ─────────────────────────────────────────────────────────────
# SharePoint uploader
# ─────────────────────────────────────────────────────────────
def upload_to_sharepoint(graph_token: str, local_path: str, filename: str) -> str:
    """
    Upload a local file to OUTPUT_FOLDER on the DataHub SharePoint drive.
    Returns the SharePoint URL of the uploaded file.
    """
    sp_path = f"{OUTPUT_FOLDER}/{filename}"
    url     = (
        f"https://graph.microsoft.com/v1.0/drives/{DRIVE_ID}"
        f"/root:/{sp_path}:/content"
    )
    headers = {
        "Authorization": f"Bearer {graph_token}",
        "Content-Type":  "application/pdf",
    }

    with open(local_path, "rb") as f:
        pdf_bytes = f.read()

    resp = requests.put(url, headers=headers, data=pdf_bytes, timeout=120)
    resp.raise_for_status()

    web_url = resp.json().get("webUrl", sp_path)
    print(f"Uploaded to SharePoint: {sp_path}")
    print(f"  URL: {web_url}")
    return web_url


# ─────────────────────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────────────────────
def main():
    # ── Load secrets from environment ────────────────────────
    tenant_id     = os.environ["TENANT_ID"]
    client_id     = os.environ["CLIENT_ID"]
    client_secret = os.environ["CLIENT_SECRET"]
    # ANTHROPIC_API_KEY is read automatically by the anthropic library

    # ── Load payload ─────────────────────────────────────────
    raw_payload = os.environ.get("JOTFORM_PAYLOAD", "")
    if not raw_payload:
        raise RuntimeError("JOTFORM_PAYLOAD environment variable is empty.")

    payload = json.loads(raw_payload)

    # Power Automate wraps the JotForm body under a "body" key
    body = payload.get("body", payload)

    submission_id = str(body.get("submissionID", "unknown"))
    print(f"Processing submission ID: {submission_id}")

    # ── Get Graph token ───────────────────────────────────────
    graph_token = get_graph_token(tenant_id, client_id, client_secret)

    # ── Parse payload ─────────────────────────────────────────
    print("Parsing payload...")
    data = parse_payload(body)
    tail = data.get("tail", "unknown").replace(" ", "_")
    print(f"  Tail: {tail}")
    print(f"  Upsells: {[u['service'] for u in data.get('upsells', [])]}")

    # ── Generate PDF ──────────────────────────────────────────
    filename   = f"{tail}_{submission_id}.pdf"
    local_path = os.path.join(tempfile.gettempdir(), filename)

    print(f"Generating PDF → {local_path}")
    generate_pdf(data, graph_token=graph_token, output_path=local_path)

    # ── Upload to SharePoint ──────────────────────────────────
    upload_to_sharepoint(graph_token, local_path, filename)

    print(f"\nDone. File saved as: {filename}")


if __name__ == "__main__":
    main()
