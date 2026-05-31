#!/usr/bin/env python3
"""
kindle_to_onenote.py

Fetch Kindle Scribe note emails from Outlook/Microsoft 365, download the
PDF (and OCR text file when present), and create a formatted OneNote page
for each note. Processed emails are moved into a dedicated mail folder so
the Inbox acts as a work queue.

See README.md for setup, configuration, and scheduling instructions.

Requires: requests, msal, tzdata
    pip install -r requirements.txt
"""

import argparse
import atexit
import html as htmlmod
import logging
import os
import re
import sys
from datetime import datetime, timedelta
from html.parser import HTMLParser
from typing import Dict, Optional, Tuple
from zoneinfo import ZoneInfo

import msal
import requests

# --------------------------------------------------------------------------
# Configuration
#
# Every value below can be overridden with an environment variable (handy for
# cron / systemd / containers). A local ".env" file, if present next to this
# script, is loaded automatically. Copy ".env.example" to ".env" to get
# started -- nothing secret is ever committed to the repository.
# --------------------------------------------------------------------------

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def _load_dotenv(path: str) -> None:
    """Minimal .env loader (no external dependency).

    Lines look like KEY=value. Existing environment variables win, so you can
    still override anything on the command line / in the scheduler.
    """
    if not os.path.exists(path):
        return
    try:
        with open(path, "r", encoding="utf-8") as fh:
            for raw in fh:
                line = raw.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                key, value = line.split("=", 1)
                key = key.strip()
                value = value.strip().strip('"').strip("'")
                os.environ.setdefault(key, value)
    except OSError:
        pass


_load_dotenv(os.path.join(BASE_DIR, ".env"))


def _env(name: str, default: str) -> str:
    value = os.environ.get(name, default)
    return value.strip() if isinstance(value, str) else value


# Azure app registration (public client -- no secret required).
CLIENT_ID = _env("KINDLE_CLIENT_ID", "YOUR-APP-CLIENT-ID")
TENANT_ID = _env("KINDLE_TENANT_ID", "consumers")  # "consumers" = personal MS accounts
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
# MSAL adds openid/profile/offline_access automatically.
SCOPES = ["Notes.ReadWrite", "Mail.ReadWrite"]

# OneNote destination -- the *section* ID inside the "Kindle Scribe" notebook.
SECTION_ID = _env("KINDLE_SECTION_ID", "PASTE-YOUR-SECTION-ID-HERE")

# Mail handling.
KINDLE_SENDER = _env("KINDLE_SENDER", "do-not-reply@amazon.com")
DEST_FOLDER_NAME = _env("KINDLE_DEST_FOLDER", "Kindle Scribe")

# Behaviour / display.
TIMEZONE = _env("KINDLE_TIMEZONE", "America/New_York")
EXPIRY_DAYS = int(_env("KINDLE_EXPIRY_DAYS", "7"))
MAX_PER_RUN = int(_env("KINDLE_MAX_PER_RUN", "25"))
TIMEOUT = int(_env("KINDLE_HTTP_TIMEOUT", "60"))

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
USER_AGENT = "Mozilla/5.0 KindleScribeFetcher/2.0"

# Token cache, persisted next to the script (override with KINDLE_TOKEN_CACHE).
CACHE_PATH = _env("KINDLE_TOKEN_CACHE", os.path.join(BASE_DIR, "token_cache.json"))

log = logging.getLogger("kindle_to_onenote")


# --------------------------------------------------------------------------
# Token cache
# --------------------------------------------------------------------------

token_cache = msal.SerializableTokenCache()
if os.path.exists(CACHE_PATH):
    try:
        with open(CACHE_PATH, "r", encoding="utf-8") as f:
            token_cache.deserialize(f.read())
    except Exception:
        log.warning("Token cache at %s was unreadable; starting fresh.", CACHE_PATH)


def _save_cache() -> None:
    if token_cache.has_state_changed:
        try:
            with open(CACHE_PATH, "w", encoding="utf-8") as f:
                f.write(token_cache.serialize())
        except OSError as exc:
            log.error("Could not write token cache: %s", exc)


atexit.register(_save_cache)


# --------------------------------------------------------------------------
# Authentication (MSAL device-code flow)
# --------------------------------------------------------------------------

def get_access_token() -> str:
    if CLIENT_ID == "YOUR-APP-CLIENT-ID":
        raise SystemExit(
            "CLIENT_ID is not configured. Set KINDLE_CLIENT_ID in your .env "
            "file (see .env.example) before running."
        )

    app = msal.PublicClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
        token_cache=token_cache,
    )

    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
        if result and "access_token" in result:
            return result["access_token"]

    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        raise RuntimeError(f"Failed to create device flow: {flow}")
    # URL + code the user types on any browser to sign in.
    print(flow["message"], flush=True)
    result = app.acquire_token_by_device_flow(flow)  # blocks until completed
    if "access_token" not in result:
        raise RuntimeError(f"Failed to get token: {result}")
    return result["access_token"]


# --------------------------------------------------------------------------
# Microsoft Graph helpers
# --------------------------------------------------------------------------

class GraphClient:
    """Thin wrapper around the Graph REST API carrying the bearer token."""

    def __init__(self, access_token: str):
        self.token = access_token

    def _auth(self, extra: Optional[Dict] = None) -> Dict[str, str]:
        headers = {"Authorization": f"Bearer {self.token}"}
        if extra:
            headers.update(extra)
        return headers

    def get(self, url: str, params: Optional[Dict] = None,
            headers_extra: Optional[Dict] = None) -> dict:
        r = requests.get(url, headers=self._auth(headers_extra),
                         params=params, timeout=TIMEOUT)
        r.raise_for_status()
        return r.json()

    def post_json(self, url: str, payload: dict,
                  headers_extra: Optional[Dict] = None) -> dict:
        headers = self._auth({"Content-Type": "application/json"})
        if headers_extra:
            headers.update(headers_extra)
        r = requests.post(url, headers=headers, json=payload, timeout=TIMEOUT)
        r.raise_for_status()
        return r.json()

    def post_raw(self, url: str, body: bytes, content_type: str) -> requests.Response:
        headers = self._auth({"Content-Type": content_type})
        r = requests.post(url, headers=headers, data=body, timeout=TIMEOUT)
        r.raise_for_status()
        return r

    # -- Mail ---------------------------------------------------------------

    def find_kindle_message_ids(self, max_to_fetch: int = MAX_PER_RUN):
        url = f"{GRAPH_BASE}/me/mailFolders/inbox/messages"
        params = {"$search": f'"from:{KINDLE_SENDER}"', "$top": max_to_fetch}
        data = self.get(url, params=params,
                        headers_extra={"ConsistencyLevel": "eventual"})
        return [m["id"] for m in data.get("value", [])]

    def fetch_message_full(self, msg_id: str) -> dict:
        params = {"$select": "id,subject,receivedDateTime,from,body"}
        return self.get(
            f"{GRAPH_BASE}/me/messages/{msg_id}",
            params=params,
            headers_extra={"Prefer": 'outlook.body-content-type="html"'},
        )

    def ensure_folder(self) -> str:
        data = self.get(f"{GRAPH_BASE}/me/mailFolders", params={"$top": 100})
        for folder in data.get("value", []):
            if folder.get("displayName") == DEST_FOLDER_NAME:
                return folder["id"]
        created = self.post_json(f"{GRAPH_BASE}/me/mailFolders",
                                 {"displayName": DEST_FOLDER_NAME})
        return created["id"]

    def move_message(self, msg_id: str, dest_folder_id: str) -> None:
        url = f"{GRAPH_BASE}/me/messages/{msg_id}/move"
        self.post_json(url, {"destinationId": dest_folder_id})
        log.info("  Moved email to '%s'", DEST_FOLDER_NAME)

    # -- OneNote ------------------------------------------------------------

    def create_onenote_page(self, body: bytes, content_type: str) -> None:
        url = f"{GRAPH_BASE}/me/onenote/sections/{SECTION_ID}/pages"
        resp = self.post_raw(url, body, content_type)
        log.info("  OneNote page created (HTTP %s)", resp.status_code)


# --------------------------------------------------------------------------
# HTML link extraction
# --------------------------------------------------------------------------

class ATagParser(HTMLParser):
    """Collect (href, text) pairs for every <a> tag in the email body."""

    def __init__(self):
        super().__init__()
        self.in_a = False
        self.href: Optional[str] = None
        self.text_chunks = []
        self.links = []  # list of (href, text)

    def handle_starttag(self, tag, attrs):
        if tag.lower() == "a":
            self.in_a = True
            self.href = dict(attrs).get("href")
            self.text_chunks = []

    def handle_data(self, data):
        if self.in_a:
            self.text_chunks.append(data)

    def handle_endtag(self, tag):
        if tag.lower() == "a" and self.in_a:
            text = "".join(self.text_chunks).strip()
            if self.href:
                self.links.append((self.href, text))
            self.in_a = False
            self.href = None
            self.text_chunks = []


def extract_download_links(body_html: str) -> Dict[str, str]:
    """Return {"pdf": url, "txt": url?} based strictly on the <a> tag text.

    The Kindle email uses distinct anchor text for each file:
        - "Download PDF" / "Download Searchable PDF"  -> PDF
        - "Download text file"                        -> OCR text file

    Matching on the anchor text (rather than link order) is the fix for the
    old TXT bug where the text file was missed or mis-assigned. If the anchor
    text has been stripped by a mail client, we fall back to the first
    amazon.com/gp/f.html link as the PDF only.
    """
    unescaped = htmlmod.unescape(body_html)
    parser = ATagParser()
    parser.feed(unescaped)

    links: Dict[str, str] = {}
    for href, text in parser.links:
        if not href or "amazon.com/gp/f.html" not in href:
            continue
        normalized = re.sub(r"\s+", " ", (text or "").strip().lower())
        if "download text file" in normalized or "text file" in normalized:
            links.setdefault("txt", href)
        elif "searchable pdf" in normalized or ("download" in normalized and "pdf" in normalized):
            links.setdefault("pdf", href)

    # Fallback for the PDF only (never guess a TXT link).
    if "pdf" not in links:
        m = re.search(r'https://www\.amazon\.com/gp/f\.html[^"\'<>\s]+', unescaped)
        if m:
            links["pdf"] = m.group(0)

    return links


# --------------------------------------------------------------------------
# Download helpers
# --------------------------------------------------------------------------

def download_file(url: str) -> bytes:
    r = requests.get(url, allow_redirects=True,
                     headers={"User-Agent": USER_AGENT}, timeout=TIMEOUT)
    r.raise_for_status()
    return r.content


def decode_txt_bytes(b: bytes) -> str:
    for enc in ("utf-8", "utf-16", "latin-1"):
        try:
            return b.decode(enc)
        except Exception:
            continue
    return b.decode("utf-8", errors="replace")


# --------------------------------------------------------------------------
# OneNote page construction
# --------------------------------------------------------------------------

def _format_expiry(received_iso_utc: str) -> str:
    dt_received_utc = datetime.fromisoformat(received_iso_utc.replace("Z", "+00:00"))
    expires = (dt_received_utc + timedelta(days=EXPIRY_DAYS)).astimezone(ZoneInfo(TIMEZONE))
    tz_abbr = expires.strftime("%Z") or TIMEZONE
    return expires.strftime(f"%m/%d/%y %I:%M %p {tz_abbr}")


def _format_page_title(plain_title: str, received_iso_utc: str) -> str:
    dt_received_utc = datetime.fromisoformat(received_iso_utc.replace("Z", "+00:00"))
    received = dt_received_utc.astimezone(ZoneInfo(TIMEZONE))
    return f"{received.strftime('%m/%d/%y')} - {plain_title}"


def build_onenote_multipart(
    pdf_bytes: bytes,
    page_title: str,
    link_pdf: str,
    received_iso_utc: str,
    txt_bytes: Optional[bytes] = None,
    txt_text: Optional[str] = None,
) -> Tuple[bytes, str]:
    """Build the multipart/form-data body for creating a OneNote page.

    The page contains, in order:
      - the title heading
      - a bold "Download PDF" hyperlink (original Amazon link)
      - an expiration line (received + EXPIRY_DAYS days)
      - "Recognized Text" (the OCR text) when a TXT file is present
      - attachments: the PDF (always) and the TXT (when present)
      - "Note Printout": the rendered PDF preview
    """
    expires_str = _format_expiry(received_iso_utc)
    safe_title = htmlmod.escape(page_title)

    recognized_block = ""
    if txt_text:
        recognized_block = (
            "<p><b>Recognized Text</b></p>"
            f"<pre>{htmlmod.escape(txt_text)}</pre>"
        )

    attachments_html = (
        "<p><b>Attachments:</b></p>"
        '<object data="name:scribe.pdf" data-attachment="kindle.pdf" '
        'type="application/pdf"></object>'
    )
    if txt_bytes is not None:
        attachments_html += (
            '<object data="name:note.txt" data-attachment="kindle.txt" '
            'type="text/plain"></object>'
        )

    html_page = (
        "<!DOCTYPE html>"
        f"<html><head><title>{safe_title}</title></head>"
        "<body>"
        f"<h2>{safe_title}</h2>"
        f'<p><b><a href="{htmlmod.escape(link_pdf, quote=True)}">Download PDF</a></b></p>'
        f"<p><i>Expires: {expires_str}</i></p>"
        f"{recognized_block}"
        f"{attachments_html}"
        "<p><b>Note Printout</b></p>"
        '<img data-render-src="name:scribe.pdf" />'
        "</body></html>"
    )

    boundary = "KindleScribeBoundary"
    crlf = "\r\n"
    content_type = f"multipart/form-data; boundary={boundary}"

    head = (
        f"--{boundary}{crlf}"
        f'Content-Disposition: form-data; name="Presentation"{crlf}'
        f"Content-Type: text/html{crlf}{crlf}"
        f"{html_page}{crlf}"
        f"--{boundary}{crlf}"
        f'Content-Disposition: form-data; name="scribe.pdf"; filename="kindle.pdf"{crlf}'
        f"Content-Type: application/pdf{crlf}{crlf}"
    )
    body = head.encode("utf-8") + pdf_bytes

    if txt_bytes is not None:
        body += (
            f"{crlf}--{boundary}{crlf}"
            f'Content-Disposition: form-data; name="note.txt"; filename="kindle.txt"{crlf}'
            f"Content-Type: text/plain; charset=utf-8{crlf}{crlf}"
        ).encode("utf-8") + txt_bytes

    body += f"{crlf}--{boundary}--{crlf}".encode("utf-8")
    return body, content_type


# --------------------------------------------------------------------------
# Subject parsing
# --------------------------------------------------------------------------

def extract_title(subject: str) -> str:
    """Pull the quoted note name out of the Kindle email subject.

    e.g. 'You sent a file "My Notes" from your Kindle' -> 'My Notes'
    """
    m = re.search(r'"([^"]+)"', subject or "")
    title = m.group(1) if m else (subject or "Kindle Note")
    return re.sub(r'[\\/:*?"<>|]', "_", title).strip() or "Kindle Note"


# --------------------------------------------------------------------------
# Main processing loop
# --------------------------------------------------------------------------

def process_message(client: GraphClient, msg: dict, folder_id: str,
                    dry_run: bool = False) -> None:
    subject = msg.get("subject", "")
    received = msg["receivedDateTime"]
    body_html = msg["body"]["content"]
    msg_id = msg["id"]

    title = extract_title(subject)
    links = extract_download_links(body_html)
    if "pdf" not in links:
        raise RuntimeError("No PDF download link found in this email.")

    log.info("Processing: %s", title)

    pdf_bytes = download_file(links["pdf"])
    log.info("  PDF bytes: %d", len(pdf_bytes))

    txt_bytes: Optional[bytes] = None
    txt_text: Optional[str] = None
    if "txt" in links:
        txt_bytes = download_file(links["txt"])
        txt_text = decode_txt_bytes(txt_bytes)
        log.info("  TXT bytes: %d | text chars: %d", len(txt_bytes), len(txt_text))
    else:
        log.info("  No text file in this email (PDF only).")

    if dry_run:
        log.info("  [dry-run] Would create OneNote page and move email.")
        return

    page_title = _format_page_title(title, received)
    body, content_type = build_onenote_multipart(
        pdf_bytes=pdf_bytes,
        page_title=page_title,
        link_pdf=links["pdf"],
        received_iso_utc=received,
        txt_bytes=txt_bytes,
        txt_text=txt_text,
    )
    client.create_onenote_page(body, content_type)
    client.move_message(msg_id, folder_id)


def run(dry_run: bool = False) -> int:
    if SECTION_ID == "PASTE-YOUR-SECTION-ID-HERE":
        raise SystemExit(
            "SECTION_ID is not configured. Set KINDLE_SECTION_ID in your .env "
            "file (see .env.example) before running."
        )

    client = GraphClient(get_access_token())

    ids = client.find_kindle_message_ids(max_to_fetch=MAX_PER_RUN)
    if not ids:
        log.info("No Kindle emails found in Inbox.")
        return 0

    msgs = [client.fetch_message_full(m_id) for m_id in ids]
    msgs.sort(key=lambda m: m["receivedDateTime"])  # oldest -> newest

    folder_id = client.ensure_folder() if not dry_run else ""

    failures = 0
    for msg in msgs:
        try:
            process_message(client, msg, folder_id, dry_run=dry_run)
        except Exception as exc:
            failures += 1
            log.error("  ERROR processing '%s' -> %s",
                      msg.get("subject", "<no subject>"), exc)

    log.info("Done. Processed %d email(s), %d failure(s).", len(msgs), failures)
    return 1 if failures else 0


def parse_args(argv=None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Import Kindle Scribe note emails into Microsoft OneNote.")
    parser.add_argument(
        "--dry-run", action="store_true",
        help="Download and report, but do not create OneNote pages or move emails.")
    parser.add_argument(
        "-v", "--verbose", action="store_true", help="Enable debug logging.")
    return parser.parse_args(argv)


def main(argv=None) -> int:
    args = parse_args(argv)
    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        stream=sys.stdout,
    )
    return run(dry_run=args.dry_run)


if __name__ == "__main__":
    sys.exit(main())
