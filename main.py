"""
PVU Upsell Document Automation — Entry Point
=============================================
Called by GitHub Actions when Power Automate POSTs a JotForm submission ID.

New flow (payload too large to pass via GitHub Actions inputs):
  1. Power Automate saves the raw payload JSON to SharePoint:
       Flow Dumps/PVU Upsell Payloads/[submissionID].json
  2. Power Automate calls GitHub Actions with only the submission ID (tiny string)
  3. This script:
       a. Gets a Graph token
       b. Downloads the payload JSON from SharePoint by submission ID
       c. Parses the payload
       d. Generates the upsell PDF
       e. Uploads the PDF to SharePoint:
            Flow Dumps/PVU Upsell PDFs/[tail]_[submissionID].pdf
       f. Deletes the staging payload JSON (cleanup)

Environment variables (set as GitHub Secrets):
  TENANT_ID         - Azure AD tenant ID
  CLIENT_ID         - Foxtrot Report Automation app client ID
  CLIENT_SECRET     - Foxtrot Report Automation app client secret
  ANTHROPIC_API_KEY - Anthropic API key (used by generate_upsell_pdf.py)
  SUBMISSION_ID     - JotForm submission ID (set by workflow YAML from PA input)
"""

import os
import json
import requests
import tempfile

from parse_payload import parse_payload
from generate_upsell_pdf import generate_pdf
from generate_upsell_docx import generate_docx


# ─────────────────────────────────────────────────────────────
# SharePoint config
# ─────────────────────────────────────────────────────────────
DRIVE_ID        = "b!_bzXaIx86kOufgJN3ih-BaDIDthKYuxJkJtLi1Bm5irGjCEnK-VHSpBRRm3_SDKU"
PAYLOAD_FOLDER  = "Flow Dumps/PVU Upsell Payloads"   # staging — Power Automate writes here
OUTPUT_FOLDER   = "Flow Dumps/PVU Upsell PDFs"        # final PDFs


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
# SharePoint helpers
# ─────────────────────────────────────────────────────────────
def download_payload(graph_token: str, submission_id: str) -> dict:
    """
    Download the raw payload JSON that Power Automate saved to SharePoint.
    Returns the parsed dict.
    """
    sp_path  = f"{PAYLOAD_FOLDER}/{submission_id}.json"
    url      = f"https://graph.microsoft.com/v1.0/drives/{DRIVE_ID}/root:/{sp_path}:/content"
    headers  = {"Authorization": f"Bearer {graph_token}"}

    resp = requests.get(url, headers=headers, timeout=60)
    if resp.status_code == 404:
        raise FileNotFoundError(
            f"Payload not found on SharePoint: {sp_path}\n"
            f"Make sure Power Automate saved it before triggering GitHub Actions."
        )
    resp.raise_for_status()
    print(f"Payload downloaded from SharePoint: {sp_path}")
    return resp.json()


def upload_pdf(graph_token: str, local_path: str, filename: str) -> str:
    """Upload the generated PDF to OUTPUT_FOLDER. Returns the SharePoint web URL."""
    sp_path = f"{OUTPUT_FOLDER}/{filename}"
    url     = f"https://graph.microsoft.com/v1.0/drives/{DRIVE_ID}/root:/{sp_path}:/content"
    headers = {
        "Authorization": f"Bearer {graph_token}",
        "Content-Type":  "application/pdf",
    }
    with open(local_path, "rb") as f:
        pdf_bytes = f.read()

    resp = requests.put(url, headers=headers, data=pdf_bytes, timeout=120)
    resp.raise_for_status()

    web_url = resp.json().get("webUrl", sp_path)
    print(f"PDF uploaded: {sp_path}")
    print(f"  URL: {web_url}")
    return web_url


def delete_payload(graph_token: str, submission_id: str) -> None:
    """Delete the staging payload JSON from SharePoint after processing."""
    sp_path = f"{PAYLOAD_FOLDER}/{submission_id}.json"
    url     = f"https://graph.microsoft.com/v1.0/drives/{DRIVE_ID}/root:/{sp_path}"
    headers = {"Authorization": f"Bearer {graph_token}"}

    resp = requests.delete(url, headers=headers, timeout=30)
    if resp.status_code in (200, 204, 404):
        print(f"Payload staging file deleted: {sp_path}")
    else:
        # Non-fatal — log and continue
        print(f"  Warning: could not delete staging file (HTTP {resp.status_code})")


# ─────────────────────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────────────────────
def main():
    # ── Load secrets ─────────────────────────────────────────
    tenant_id     = os.environ["TENANT_ID"]
    client_id     = os.environ["CLIENT_ID"]
    client_secret = os.environ["CLIENT_SECRET"]
    # ANTHROPIC_API_KEY is read automatically by the anthropic library

    submission_id = os.environ.get("SUBMISSION_ID", "").strip()
    if not submission_id:
        raise RuntimeError("SUBMISSION_ID environment variable is empty.")
    print(f"Processing submission ID: {submission_id}")

    # ── Get Graph token ───────────────────────────────────────
    graph_token = get_graph_token(tenant_id, client_id, client_secret)

    # ── Download payload from SharePoint ─────────────────────
    payload = download_payload(graph_token, submission_id)

    # Power Automate wraps the JotForm body under a "body" key
    body = payload.get("body", payload)

    # ── Parse payload ─────────────────────────────────────────
    print("Parsing payload...")
    data = parse_payload(body)
    tail = data.get("tail", "unknown").replace(" ", "_")
    print(f"  Tail: {tail}")
    print(f"  Upsells: {[u['service'] for u in data.get('upsells', [])]}")

    # ── Generate PDF ──────────────────────────────────────────
    pdf_filename = f"{tail}_{submission_id}.pdf"
    pdf_path     = os.path.join(tempfile.gettempdir(), pdf_filename)

    print(f"Generating PDF -> {pdf_path}")
    generate_pdf(data, graph_token=graph_token, output_path=pdf_path)

    # ── Generate Word doc ─────────────────────────────────────
    docx_filename = f"{tail}_{submission_id}.docx"
    docx_path     = os.path.join(tempfile.gettempdir(), docx_filename)

    print(f"Generating Word doc -> {docx_path}")
    generate_docx(data, graph_token=graph_token, output_path=docx_path)

    # ── Upload both to SharePoint ─────────────────────────────
    upload_pdf(graph_token, pdf_path, pdf_filename)
    upload_pdf(graph_token, docx_path, docx_filename)

    # ── Clean up staging file ─────────────────────────────────
    delete_payload(graph_token, submission_id)

    print(f"\nDone. Files saved: {pdf_filename} and {docx_filename}")


if __name__ == "__main__":
    main()
