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
import random
import re
import secrets
import sys
import time
from datetime import datetime, timedelta
from html.parser import HTMLParser
from typing import Callable, Dict, List, Optional, Tuple
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
# MSAL adds openid/profile/offline_access automatically. MailboxSettings is
# appended below only when category color management is enabled.
SCOPES = ["Notes.ReadWrite", "Mail.ReadWrite"]

# OneNote destination -- the *section* ID inside the "Kindle Scribe" notebook.
SECTION_ID = _env("KINDLE_SECTION_ID", "PASTE-YOUR-SECTION-ID-HERE")

# Mail handling.
KINDLE_SENDER = _env("KINDLE_SENDER", "do-not-reply@amazon.com")
DEST_FOLDER_NAME = _env("KINDLE_DEST_FOLDER", "Kindle Scribe")  # override: --folder

# Outlook categories applied to processed emails (override: --tag1 / --tag2).
#   TAG1 is added when a note is picked up for processing.
#   TAG2 is added only after the OneNote page is confirmed created. TAG2 also
#   acts as the duplicate guard: if it's already present, the page exists, so a
#   re-run just retries the move instead of creating a second page.
TAG1 = _env("KINDLE_TAG1", "Kindle Scribe")
TAG2 = _env("KINDLE_TAG2", "Uploaded to OneNote")
# Colors used only when category management is enabled (Outlook preset names).
TAG1_COLOR = _env("KINDLE_TAG1_COLOR", "preset7")  # blue
TAG2_COLOR = _env("KINDLE_TAG2_COLOR", "preset4")  # green
# Assigning a category to a message always works (Mail.ReadWrite); but an
# unknown category shows without a color until it exists in the mailbox master
# list. Enabling this creates the categories (with the colors above) in the
# master list so they show colored -- which needs MailboxSettings.ReadWrite and
# a re-consent (re-run --login). Override: --manage-categories.
MANAGE_CATEGORIES = _env("KINDLE_MANAGE_CATEGORIES", "false").lower() in (
    "1", "true", "yes", "on")
if MANAGE_CATEGORIES:
    SCOPES.append("MailboxSettings.ReadWrite")

# Behaviour / display.
TIMEZONE = _env("KINDLE_TIMEZONE", "America/New_York")
EXPIRY_DAYS = int(_env("KINDLE_EXPIRY_DAYS", "7"))
MAX_PER_RUN = int(_env("KINDLE_MAX_PER_RUN", "25"))
TIMEOUT = int(_env("KINDLE_HTTP_TIMEOUT", "60"))
MAX_RETRIES = int(_env("KINDLE_MAX_RETRIES", "4"))

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
USER_AGENT = "Mozilla/5.0 KindleScribeFetcher/2.1"

# Token cache, persisted next to the script (override with KINDLE_TOKEN_CACHE).
CACHE_PATH = _env("KINDLE_TOKEN_CACHE", os.path.join(BASE_DIR, "token_cache.json"))

log = logging.getLogger("kindle_to_onenote")

# Shared HTTP session (connection pooling + a single place for defaults).
SESSION = requests.Session()
SESSION.headers["User-Agent"] = USER_AGENT


# --------------------------------------------------------------------------
# HTTP with retry/backoff
# --------------------------------------------------------------------------

# Status codes worth retrying: throttling (429) and transient server errors.
RETRY_STATUS = {429, 500, 502, 503, 504}


def _retry_after_seconds(resp: requests.Response) -> Optional[float]:
    value = resp.headers.get("Retry-After")
    if not value:
        return None
    try:
        return float(value)  # Graph sends a delta in seconds
    except ValueError:
        return None


def _backoff_seconds(attempt: int) -> float:
    return min(2 ** attempt, 30) + random.uniform(0, 0.5)


def http_request(method: str, url: str, **kwargs) -> requests.Response:
    """Perform an HTTP request, retrying transient failures with backoff.

    Retries on connection errors and on 429/5xx responses, honoring the
    server's Retry-After header when present. Returns the final response
    (the caller still decides whether to raise_for_status()).
    """
    kwargs.setdefault("timeout", TIMEOUT)
    last_exc: Optional[Exception] = None
    for attempt in range(MAX_RETRIES + 1):
        try:
            resp = SESSION.request(method, url, **kwargs)
        except requests.RequestException as exc:
            last_exc = exc
            if attempt >= MAX_RETRIES:
                raise
            delay = _backoff_seconds(attempt)
            log.warning("  %s %s failed (%s); retrying in %.1fs",
                        method, url, exc, delay)
            time.sleep(delay)
            continue

        if resp.status_code in RETRY_STATUS and attempt < MAX_RETRIES:
            delay = _retry_after_seconds(resp) or _backoff_seconds(attempt)
            log.warning("  %s %s -> HTTP %s; retrying in %.1fs",
                        method, url, resp.status_code, delay)
            time.sleep(delay)
            continue

        return resp

    # Unreachable in practice, but keeps type-checkers happy.
    assert last_exc is not None
    raise last_exc


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

def get_access_token(interactive: bool = True) -> str:
    """Return an access token, refreshing silently from the cache when possible.

    When the cache can't satisfy the request and ``interactive`` is True, the
    device-code flow runs (prints a URL + code to complete on any device). When
    ``interactive`` is False -- e.g. an unattended cron/systemd run on a
    headless box -- it raises instead of blocking on a prompt nobody can see,
    telling you to re-authenticate with ``--login``.
    """
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

    if not interactive:
        raise SystemExit(
            "No valid cached credentials and not running interactively "
            "(headless/scheduled). Re-authenticate once with:\n"
            "    python kindle_to_onenote.py --login")

    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        raise RuntimeError(f"Failed to create device flow: {flow}")
    # URL + code the user types on any browser/device to sign in.
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

    def __init__(self, token_provider: Callable[[], str]):
        # token_provider returns a fresh access token (MSAL refreshes silently).
        self._token_provider = token_provider
        self.token = token_provider()

    def _send(self, method: str, url: str, *,
              json: Optional[dict] = None,
              data: Optional[bytes] = None,
              params: Optional[Dict] = None,
              content_type: Optional[str] = None,
              headers_extra: Optional[Dict] = None) -> requests.Response:
        """Send an authenticated request, refreshing the token once on a 401.

        Covers the edge case where the access token expires mid-run (e.g. a
        long batch); MSAL hands back a refreshed token and we retry once.
        """
        def do() -> requests.Response:
            headers = {"Authorization": f"Bearer {self.token}"}
            if content_type:
                headers["Content-Type"] = content_type
            if headers_extra:
                headers.update(headers_extra)
            return http_request(method, url, headers=headers,
                                json=json, data=data, params=params)

        resp = do()
        if resp.status_code == 401:
            log.info("  Access token rejected (401); refreshing and retrying.")
            self.token = self._token_provider()
            resp = do()
        resp.raise_for_status()
        return resp

    def get(self, url: str, params: Optional[Dict] = None,
            headers_extra: Optional[Dict] = None) -> dict:
        return self._send("GET", url, params=params,
                          headers_extra=headers_extra).json()

    def post_json(self, url: str, payload: dict,
                  headers_extra: Optional[Dict] = None) -> dict:
        return self._send("POST", url, json=payload,
                          content_type="application/json",
                          headers_extra=headers_extra).json()

    def patch_json(self, url: str, payload: dict) -> dict:
        return self._send("PATCH", url, json=payload,
                          content_type="application/json").json()

    def post_raw(self, url: str, body: bytes, content_type: str) -> requests.Response:
        return self._send("POST", url, data=body, content_type=content_type)

    # -- Mail ---------------------------------------------------------------

    def find_kindle_message_ids(self, max_to_fetch: int = MAX_PER_RUN) -> List[str]:
        """Return Kindle email IDs in the inbox.

        Prefers an exact ``$filter`` on the sender address (immediate and
        precise). Note: Graph rejects ``$filter`` on ``from`` combined with
        ``$orderby`` ("restriction or sort order too complex"), so we don't
        sort here -- the caller sorts the fetched messages by received time.
        Falls back to ``$search`` if the mailbox rejects the filter.
        """
        url = f"{GRAPH_BASE}/me/mailFolders/inbox/messages"
        try:
            params = {
                "$filter": f"from/emailAddress/address eq '{KINDLE_SENDER}'",
                "$select": "id",
                "$top": max_to_fetch,
            }
            data = self.get(url, params=params)
            return [m["id"] for m in data.get("value", [])]
        except requests.HTTPError as exc:
            log.warning("$filter query failed (%s); falling back to $search.", exc)
            params = {"$search": f'"from:{KINDLE_SENDER}"', "$top": max_to_fetch}
            data = self.get(url, params=params,
                            headers_extra={"ConsistencyLevel": "eventual"})
            return [m["id"] for m in data.get("value", [])]

    def fetch_message_full(self, msg_id: str) -> dict:
        params = {"$select": "id,subject,receivedDateTime,from,body,categories"}
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

    def add_categories(self, msg_id: str, existing: List[str], *names: str) -> List[str]:
        """Add one or more categories to a message; returns the new list.

        Assigning a category works even if it isn't in the mailbox master list
        (it just shows uncolored until then). Only PATCHes if something changed.
        """
        cats = list(existing or [])
        changed = False
        for name in names:
            if name and name not in cats:
                cats.append(name)
                changed = True
        if changed:
            self.patch_json(f"{GRAPH_BASE}/me/messages/{msg_id}", {"categories": cats})
            log.info("  Categorized: %s", ", ".join(names))
        return cats

    def get_master_categories(self) -> set:
        data = self.get(f"{GRAPH_BASE}/me/outlook/masterCategories",
                        params={"$select": "displayName", "$top": 100})
        return {c.get("displayName") for c in data.get("value", [])}

    def ensure_master_categories(self, wanted: List[Tuple[str, str]]) -> None:
        """Create any missing (name, color) categories in the master list.

        Best-effort: needs MailboxSettings.ReadWrite. If that's not granted, log
        a warning and continue -- per-message categorization still works, the
        categories just won't be colored.
        """
        try:
            existing = self.get_master_categories()
        except requests.HTTPError as exc:
            log.warning("Could not read Outlook master categories (%s); "
                        "categories will be applied without managed colors. "
                        "Grant MailboxSettings.ReadWrite and re-run --login to "
                        "enable colors.", exc)
            return
        for name, color in wanted:
            if name and name not in existing:
                try:
                    self.post_json(f"{GRAPH_BASE}/me/outlook/masterCategories",
                                   {"displayName": name, "color": color})
                    log.info("Created Outlook category '%s' (%s).", name, color)
                except requests.HTTPError as exc:
                    log.warning("Could not create category '%s' (%s).", name, exc)

    def move_message(self, msg_id: str, dest_folder_id: str) -> None:
        url = f"{GRAPH_BASE}/me/messages/{msg_id}/move"
        self.post_json(url, {"destinationId": dest_folder_id})
        log.info("  Moved email to '%s'", DEST_FOLDER_NAME)

    # -- OneNote ------------------------------------------------------------

    def list_sections(self) -> List[dict]:
        url = f"{GRAPH_BASE}/me/onenote/sections"
        params = {
            "$select": "id,displayName",
            "$expand": "parentNotebook($select=displayName)",
            "$top": 100,
        }
        return self.get(url, params=params).get("value", [])

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
    """Return {"pdf": url, "txt": url?} based on each <a> tag's text.

    The Kindle email uses distinct anchor text for each file:
        - "Download PDF" / "Download Searchable PDF"  -> PDF
        - "Download text file"                        -> OCR text file

    We classify by the anchor *text*, not the URL, because Amazon's PDF and
    text-file links can use different hosts/paths (the old code required an
    `amazon.com/gp/f.html` href and silently dropped the text-file link, which
    uses a different URL). Any real http(s) link is accepted. If anchor text
    was stripped by a mail client, we fall back to the first Amazon
    `gp/f.html` URL as the PDF only (never guess a TXT link).
    """
    unescaped = htmlmod.unescape(body_html)
    parser = ATagParser()
    parser.feed(unescaped)

    links: Dict[str, str] = {}
    for href, text in parser.links:
        if not href or not href.lower().startswith(("http://", "https://")):
            continue
        normalized = re.sub(r"\s+", " ", (text or "").strip().lower())
        if "text file" in normalized or "text document" in normalized:
            links.setdefault("txt", href)
        elif "pdf" in normalized:  # "Download PDF" or "Download Searchable PDF"
            links.setdefault("pdf", href)

    # Fallback for the PDF only, if anchor text was stripped.
    if "pdf" not in links:
        m = re.search(r'https?://[^"\'<>\s]*amazon\.[^"\'<>\s]*/gp/f\.html[^"\'<>\s]*',
                      unescaped)
        if m:
            links["pdf"] = m.group(0)

    return links


# --------------------------------------------------------------------------
# Download helpers
# --------------------------------------------------------------------------

def download_file(url: str) -> bytes:
    r = http_request("GET", url, allow_redirects=True)
    r.raise_for_status()
    return r.content


def looks_like_pdf(b: bytes) -> bool:
    """A real PDF starts with the '%PDF-' magic bytes."""
    return b[:5] == b"%PDF-"


def looks_like_html(b: bytes) -> bool:
    """Detect an HTML page (e.g. an Amazon 'link expired' error page)."""
    head = b[:512].lstrip().lower()
    return head.startswith(b"<!doctype html") or head.startswith(b"<html")


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


def _make_boundary(*payloads: Optional[bytes]) -> str:
    """Return a multipart boundary guaranteed not to appear in the payloads."""
    while True:
        boundary = "KindleScribe" + secrets.token_hex(16)
        marker = f"--{boundary}".encode("utf-8")
        if all(p is None or marker not in p for p in payloads):
            return boundary


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

    # Declaring UTF-8 ensures non-ASCII OCR text/titles decode correctly, and
    # <meta name="created"> sets the page's date to when the note was emailed.
    safe_created = htmlmod.escape(received_iso_utc, quote=True)
    html_page = (
        "<!DOCTYPE html>"
        "<html><head>"
        '<meta charset="utf-8" />'
        f'<meta name="created" content="{safe_created}" />'
        f"<title>{safe_title}</title>"
        "</head>"
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

    boundary = _make_boundary(pdf_bytes, txt_bytes)
    crlf = "\r\n"
    content_type = f"multipart/form-data; boundary={boundary}"

    head = (
        f"--{boundary}{crlf}"
        f'Content-Disposition: form-data; name="Presentation"{crlf}'
        f"Content-Type: text/html; charset=utf-8{crlf}{crlf}"
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
    categories = msg.get("categories") or []

    title = extract_title(subject)

    # Idempotency guard: TAG2 is only set after a page is confirmed created. If
    # it's already present, a previous run made the page but didn't finish the
    # move -- just retry the move, never create a second page.
    if TAG2 in categories:
        log.info("Already uploaded previously: %s -- retrying move only.", title)
        if not dry_run:
            client.move_message(msg_id, folder_id)
        return

    links = extract_download_links(body_html)
    if "pdf" not in links:
        raise RuntimeError("No PDF download link found in this email.")

    log.info("Processing: %s", title)

    # Tag the email as picked up (TAG1) before doing the work, so even a note
    # that later fails to upload is still marked.
    if not dry_run:
        categories = client.add_categories(msg_id, categories, TAG1)

    pdf_bytes = download_file(links["pdf"])
    log.info("  PDF bytes: %d", len(pdf_bytes))
    if not looks_like_pdf(pdf_bytes):
        # Most likely the Amazon link expired and returned an HTML error
        # page. Raise so the email stays in the inbox for a manual re-send.
        raise RuntimeError(
            "Downloaded file is not a valid PDF (the Amazon link may have "
            "expired). Leaving the email in the inbox.")

    txt_bytes: Optional[bytes] = None
    txt_text: Optional[str] = None
    if "txt" in links:
        candidate = download_file(links["txt"])
        if looks_like_html(candidate):
            raise RuntimeError(
                "Downloaded text file looks like an HTML error page (the "
                "Amazon link may have expired). Leaving the email in the inbox.")
        txt_bytes = candidate
        txt_text = decode_txt_bytes(txt_bytes)
        log.info("  TXT bytes: %d | text chars: %d", len(txt_bytes), len(txt_text))
    else:
        log.info("  No text file in this email (PDF only).")

    if dry_run:
        log.info("  [dry-run] Would create OneNote page, categorize, and move email.")
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
    # Confirmed: stamp TAG2 *before* moving so a move failure cannot duplicate.
    client.add_categories(msg_id, categories, TAG1, TAG2)
    client.move_message(msg_id, folder_id)


def run(dry_run: bool = False, interactive: bool = True) -> int:
    if SECTION_ID == "PASTE-YOUR-SECTION-ID-HERE":
        raise SystemExit(
            "SECTION_ID is not configured. Set KINDLE_SECTION_ID in your .env "
            "file (see .env.example) before running. Tip: run with "
            "--list-sections to discover it.")

    client = GraphClient(lambda: get_access_token(interactive=interactive))

    ids = client.find_kindle_message_ids(max_to_fetch=MAX_PER_RUN)
    if not ids:
        log.info("No Kindle emails found in Inbox.")
        return 0

    msgs = [client.fetch_message_full(m_id) for m_id in ids]
    msgs.sort(key=lambda m: m["receivedDateTime"])  # oldest -> newest

    folder_id = client.ensure_folder() if not dry_run else ""
    if MANAGE_CATEGORIES and not dry_run:
        client.ensure_master_categories([(TAG1, TAG1_COLOR), (TAG2, TAG2_COLOR)])

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


def login() -> int:
    """Run the interactive device-code sign-in once and exit (setup helper)."""
    get_access_token(interactive=True)
    print("Authenticated. Token cached at:", CACHE_PATH)
    return 0


def list_sections() -> int:
    """Print every OneNote section and its ID (a setup-time helper)."""
    client = GraphClient(lambda: get_access_token(interactive=True))
    sections = client.list_sections()
    if not sections:
        print("No OneNote sections found for this account.")
        return 0
    print("Notebook / Section -> SECTION_ID")
    print("-" * 60)
    for s in sections:
        notebook = (s.get("parentNotebook") or {}).get("displayName", "?")
        print(f"{notebook} / {s.get('displayName', '?')} -> {s['id']}")
    print("\nCopy the desired ID into .env as KINDLE_SECTION_ID.")
    return 0


def parse_args(argv=None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Import Kindle Scribe note emails into Microsoft OneNote.")
    parser.add_argument(
        "--dry-run", action="store_true",
        help="Download and report, but do not create OneNote pages or move emails.")
    parser.add_argument(
        "--list-sections", action="store_true",
        help="List your OneNote sections and their IDs, then exit (setup helper).")
    parser.add_argument(
        "--login", action="store_true",
        help="Run the interactive device-code sign-in once and exit. Use this "
             "for first-time setup on a headless box, then schedule normal runs.")
    parser.add_argument(
        "--non-interactive", action="store_true",
        help="Never prompt for sign-in; fail fast if a silent refresh isn't "
             "possible. Implied automatically when not attached to a terminal.")
    parser.add_argument(
        "--tag1", metavar="NAME",
        help=f"Category added to emails when processed (default: {TAG1!r}).")
    parser.add_argument(
        "--tag2", metavar="NAME",
        help=f"Category added once the OneNote page is confirmed "
             f"(default: {TAG2!r}).")
    parser.add_argument(
        "--folder", metavar="NAME",
        help=f"Mail folder to move processed emails into "
             f"(default: {DEST_FOLDER_NAME!r}).")
    parser.add_argument(
        "--manage-categories", action="store_true",
        help="Create the categories in your Outlook master list with colors so "
             "they show colored. Needs the MailboxSettings.ReadWrite scope; "
             "re-run --login after enabling.")
    parser.add_argument(
        "-v", "--verbose", action="store_true", help="Enable debug logging.")
    return parser.parse_args(argv)


def _apply_overrides(args: argparse.Namespace) -> None:
    """Apply CLI overrides onto the module-level config globals."""
    global TAG1, TAG2, DEST_FOLDER_NAME, MANAGE_CATEGORIES
    if args.tag1:
        TAG1 = args.tag1
    if args.tag2:
        TAG2 = args.tag2
    if args.folder:
        DEST_FOLDER_NAME = args.folder
    if args.manage_categories:
        MANAGE_CATEGORIES = True
        if "MailboxSettings.ReadWrite" not in SCOPES:
            SCOPES.append("MailboxSettings.ReadWrite")


def main(argv=None) -> int:
    args = parse_args(argv)
    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        stream=sys.stdout,
    )
    _apply_overrides(args)
    if args.login:
        return login()
    if args.list_sections:
        return list_sections()
    # Only allow the interactive device-code prompt when a human is attached
    # (a TTY) and hasn't opted out. Cron/systemd runs are non-interactive.
    interactive = sys.stdin.isatty() and not args.non_interactive
    return run(dry_run=args.dry_run, interactive=interactive)


if __name__ == "__main__":
    sys.exit(main())
