# scribe-to-onenote

Automatically import **Kindle Scribe** handwritten notes into **Microsoft
OneNote**.

The Kindle Scribe (2022) can't sync to OneNote directly, but it *can* email
you your notes. Amazon's email doesn't attach the files — it contains
download links to a **PDF** and, when OCR is enabled, a **searchable PDF**
plus a **text (`.txt`) file**. This script watches your Outlook/Microsoft 365
inbox for those emails, downloads the files, creates a nicely formatted
OneNote page (with the PDF rendered inline, the recognized text, and both
files attached), and then files the email away so your inbox stays a clean
work queue.

It runs unattended on a schedule (cron, systemd timer, or Windows Task
Scheduler) on Windows, Linux, Raspberry Pi, or a small VM/container.

> ℹ️ **AI-generated project.** This code and documentation were written with AI
> assistance (originally with ChatGPT, then refactored and extended with
> Claude). See [AI disclosure & disclaimer](#ai-disclosure--disclaimer) before
> relying on it.

---

## How it works

```
Kindle Scribe  ──email──▶  Outlook Inbox  ──▶  this script  ──▶  OneNote page
                                                   │
                                                   ├─▶ tag email: "Kindle Scribe", then
                                                   │              "Uploaded to OneNote"
                                                   └─▶ move email to the "Kindle Scribe" folder
```

(The category names and the destination folder are the defaults — all
configurable; see [Categories & folder](#categories--folder).)

For each Kindle email from `do-not-reply@amazon.com`, the script:

1. Reads the subject (the quoted note name) and the received timestamp.
2. Extracts the download links from the email HTML **by anchor text**:
   - `Download PDF` / `Download Searchable PDF` → the PDF
   - `Download text file` → the OCR `.txt` file (only if that link exists)
3. Downloads the PDF (always) and the text file (when present), and tags the
   email with the **`Kindle Scribe`** Outlook category.
4. Creates a OneNote page in your configured section containing:
   - **Title:** `MM/DD/YY - <note name>`
   - A bold **Download PDF** hyperlink to the original Amazon link
   - An **Expires** line (`received + 7 days`, the Amazon link lifetime)
   - **Recognized Text** (the OCR text) — only when a `.txt` file exists
   - **Attachments:** the PDF, and the `.txt` when present
   - **Note Printout:** the PDF rendered inline
5. Tags the email with the **`Uploaded to OneNote`** category (confirmation),
   then moves it into the **Kindle Scribe** mail folder (created if needed).

The inbox is treated as the queue: every matching email still in the inbox is
processed each run, so backlogged or batched notes are all handled.

> **About the TXT fix:** earlier versions of this script sometimes missed or
> mis-assigned the text file because they guessed based on link order. This
> version identifies each link strictly by its anchor text, so the searchable
> PDF and the text file never get confused. The text file is only downloaded
> when an actual `Download text file` link is present.

### Reliability built in

- **Expired-link protection.** Amazon's download links expire (≈7 days). The
  script verifies the download is a real PDF (`%PDF-` magic bytes) before
  uploading; if a link has expired and returns an HTML error page instead, it
  raises and **leaves the email in the inbox** rather than creating a broken
  page and losing the note.
- **No duplicates.** After a page is created the email is tagged with the
  `Uploaded to OneNote` category *before* it is moved. If a later step fails
  and the email is processed again, the script sees that tag and retries only
  the move — it never creates a second page.
- **Survives throttling/blips.** Microsoft Graph (`429`/`5xx`, honoring
  `Retry-After`) and Amazon downloads are retried with exponential backoff.
- **Auto re-auth.** Access tokens are refreshed silently from the cached
  refresh token; if one expires mid-run the request is retried once with a
  fresh token (no device-code prompt).
- **Unicode-safe.** The page is sent as UTF-8, so accented/non-ASCII OCR text
  and note titles render correctly. The page's creation date is also set to
  when the note was emailed.

### Categories & folder

Each processed email is tagged with two Outlook categories and then filed into
a folder — all configurable:

| What | Default | Override |
|---|---|---|
| Category added when a note is processed | `Kindle Scribe` | `--tag1` / `KINDLE_TAG1` |
| Category added once the page is confirmed | `Uploaded to OneNote` | `--tag2` / `KINDLE_TAG2` |
| Folder processed emails are moved to | `Kindle Scribe` | `--folder` / `KINDLE_DEST_FOLDER` |

```bash
python kindle_to_onenote.py --tag1 "Scribe" --tag2 "In OneNote" --folder "Scribe Notes"
```

**Do the categories need to exist first?** No. Assigning a category to a message
works whether or not it's in your Outlook *master category list* — the script
does this with its normal `Mail.ReadWrite` permission. The only catch is that a
category not in the master list shows **without a color** until you add it
there (in Outlook, or automatically — see below).

**Want colored categories?** Add `--manage-categories` (or
`KINDLE_MANAGE_CATEGORIES=true`). The script will then create the two
categories in your master list with colors (`KINDLE_TAG1_COLOR` /
`KINDLE_TAG2_COLOR`, Outlook `presetN` names). This needs the extra
**`MailboxSettings.ReadWrite`** delegated permission — add it to your Azure app
registration and re-run `--login` to consent. If the permission is missing the
script logs a warning and carries on (emails still get categorized, just
uncolored).

You can view your current categories and their colors with this pre-filled
[Graph Explorer query for your master categories][ge-categories]:

```text
GET https://graph.microsoft.com/v1.0/me/outlook/masterCategories
```

[ge-categories]: https://developer.microsoft.com/graph/graph-explorer?request=me/outlook/masterCategories&method=GET&version=v1.0&GraphUrl=https://graph.microsoft.com

### Known limitation: 4 MB page size

Microsoft Graph caps a OneNote page-creation request at **~4 MB**. Because the
PDF is embedded in that request, a very large handwritten note (many pages, or
high-detail ink) can exceed the limit and the upload will fail with a
`413 "request too large"` error, logged for that note. The email is left in the
inbox. If you hit this, split the note into smaller PDFs on the Scribe and
re-send. (Most single notes are well under 4 MB.)

---

## Prerequisites

- Python 3.9+ (3.11+ recommended)
- A Microsoft account (personal `consumers`, or work/school) with OneNote
- An [Azure App Registration](#1-create-an-azure-app-registration) (free)
- Your Kindle configured to email notes to the same mailbox the script reads

---

## Setup

### 1. Create an Azure app registration

The script signs in **as you** (delegated permissions) and needs **no client
secret** — it uses the MSAL device-code flow. The registration exists only to
give you a **client ID** to authenticate with. It's free.

Open the app registrations page and click **+ New registration**:

> 🔗 **[Microsoft Entra → App registrations](https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade)**
> — the same page is in the [Azure portal](https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade)
> under **Microsoft Entra ID → App registrations**.

Then fill in the registration:

1. **Name** — anything you'll recognize; it's just a label shown to you at
   sign-in. For example: `Kindle to OneNote`.
2. **Supported account types** — match the account your Kindle emails go to:
   - Personal Microsoft account (outlook.com / hotmail / live) →
     **Personal Microsoft accounts only**.
   - Work/school (Microsoft 365) account → **Accounts in this organizational
     directory only**.
3. **Redirect URI** — leave it **blank**, then click **Register**.

Now configure the new app:

4. **Authentication** → turn on **Allow public client flows** → **Save**. This
   enables the device-code sign-in, and is why no client secret is needed.
5. **API permissions** → **Add a permission** → **Microsoft Graph** →
   **Delegated permissions** (delegated = the app acts on your behalf, the
   right choice for a personal tool), then add:

   | Permission | Why the script needs it | Required? |
   |---|---|---|
   | `Notes.ReadWrite` | Create the OneNote page (printout + attachments) in your section | **Yes** |
   | `Mail.ReadWrite` | Find the Kindle emails, tag them with categories, and move them to a folder | **Yes** |
   | `MailboxSettings.ReadWrite` | Create the *colored* categories in your mailbox's master list | Only for `--manage-categories` |

   Search for each name in the permission picker (copy/paste these):

   ```text
   Notes.ReadWrite
   Mail.ReadWrite
   ```

   Optional — add this only if you want colored categories
   (see [Categories & folder](#categories--folder)); you can add it later:

   ```text
   MailboxSettings.ReadWrite
   ```

   `offline_access`, `openid`, and `profile` are added automatically by MSAL —
   don't add them here. On a personal account you don't need "Grant admin
   consent"; you'll consent at first sign-in.
6. **Overview** → copy the **Application (client) ID**. That's your
   `KINDLE_CLIENT_ID`.

### 2. Install

```bash
git clone https://github.com/Mcp20091/scribe-to-onenote.git
cd scribe-to-onenote

python -m venv .venv
# Linux/macOS:
source .venv/bin/activate
# Windows (PowerShell):
# .venv\Scripts\Activate.ps1

pip install -r requirements.txt
```

> **Tip:** don't share a virtualenv between Windows and Linux. Create a
> separate `.venv` per OS and reinstall from `requirements.txt`.

### 3. Configure

```bash
cp .env.example .env
```

Only **two** variables are required:

- `KINDLE_CLIENT_ID` — the Application (client) ID from step 1
- `KINDLE_SECTION_ID` — the OneNote section to write to (see below)

Everything else has a sensible default. The full list — including the
categories, mail folder, timezone, and retry/timeout knobs — is in the
[Configuration reference](#configuration-reference). The `.env` file is
git-ignored so your IDs never get committed.

#### Finding your section ID

Once `KINDLE_CLIENT_ID` is set, run the built-in helper (this is a one-time
setup command — it signs you in, prints your sections, and exits):

```bash
python kindle_to_onenote.py --list-sections
```

Output looks like:

```text
Notebook / Section -> SECTION_ID
------------------------------------------------------------
Kindle Scribe / Quick Notes -> 0-ABC123...!456
```

Copy the `id` of the section you want (e.g. the **Quick Notes** section of a
**Kindle Scribe** notebook) into `.env` as `KINDLE_SECTION_ID`.

**Prefer a browser?** Open this pre-filled
[Graph Explorer query for your sections][ge-sections] (sign in with the same
Microsoft account, then click **Run query**) and read the `id` from the
response. The underlying request is:

```text
GET https://graph.microsoft.com/v1.0/me/onenote/sections
```

If you're not sure which notebook a section belongs to, this
[query for your notebooks][ge-notebooks] lists them too.

[ge-sections]: https://developer.microsoft.com/graph/graph-explorer?request=me/onenote/sections&method=GET&version=v1.0&GraphUrl=https://graph.microsoft.com
[ge-notebooks]: https://developer.microsoft.com/graph/graph-explorer?request=me/onenote/notebooks&method=GET&version=v1.0&GraphUrl=https://graph.microsoft.com

### 4. First run (authenticate)

Sign in once with the `--login` command:

```bash
python kindle_to_onenote.py --login
```

It prints a URL and a device code. Open the URL on **any device** (your
laptop or phone — no browser is needed on the server itself), enter the code,
and approve the requested permissions. A `token_cache.json` file is written
next to the script (git-ignored — it holds your refresh token). Subsequent
runs refresh the token silently with no prompts.

This is the [device-code flow](https://learn.microsoft.com/entra/identity-platform/v2-oauth2-device-code),
which is designed for headless machines — see [Headless / LXC / VM notes](#headless--lxc--vm-notes).

Useful flags:

```bash
python kindle_to_onenote.py --login          # one-time interactive sign-in (also use to re-auth)
python kindle_to_onenote.py --dry-run        # download + report, but don't write to OneNote or move mail
python kindle_to_onenote.py --verbose        # debug logging
python kindle_to_onenote.py --list-sections  # print your OneNote section IDs and exit (setup helper)
python kindle_to_onenote.py --tag1 NAME --tag2 NAME --folder NAME   # override categories / mail folder
python kindle_to_onenote.py --manage-categories                    # also create colored categories (needs extra scope)
```

---

## Headless / LXC / VM notes

This tool is built to run unattended on a headless box (e.g. a Proxmox LXC
container, a VM, or a Raspberry Pi). A few things that make that work cleanly:

- **No GUI or browser needed on the server.** Authentication uses the OAuth
  **device-code flow**: you run `--login` once over SSH, then complete the
  sign-in on your phone/laptop. The token cache is a plain file, so there's no
  dependency on a desktop keyring/Secret Service (which containers usually lack).
- **Scheduled runs never block on a prompt.** When the script isn't attached to
  a terminal (cron, systemd timer) it runs **non-interactively**: if the cached
  token can't be refreshed silently it exits immediately with a clear message
  telling you to re-run `--login`, instead of hanging on a device code nobody
  can see. (You can force this with `--non-interactive`.)
- **Persistent, writable cache path.** Make sure `token_cache.json` lives on a
  persistent, writable path for the user the scheduler runs as — not a tmpfs.
  Set `KINDLE_TOKEN_CACHE` to pin its location if needed, and use absolute
  paths in cron/systemd so every run finds the same cache.
- **Re-authenticating later.** If the refresh token is ever invalidated (you
  change your Microsoft password, revoke the app, or leave it idle past the
  ~90-day window), just SSH in and run `--login` again.

---

## Scheduling

### Linux / Raspberry Pi — cron

Edit your crontab (`crontab -e`) and add one of these. Use absolute paths to
the venv's Python and to the script, and redirect output to a log:

```cron
# Every 15 minutes
*/15 * * * * /home/USER/scribe-to-onenote/.venv/bin/python /home/USER/scribe-to-onenote/kindle_to_onenote.py >> /home/USER/scribe-to-onenote/kindle.log 2>&1
```

A copy lives in [`deploy/crontab.example`](deploy/crontab.example).

### Linux — systemd timer (alternative to cron)

Copy the unit files from [`deploy/`](deploy/), edit the paths/user, then:

```bash
sudo cp deploy/kindle-to-onenote.service deploy/kindle-to-onenote.timer /etc/systemd/system/
sudo systemctl daemon-reload
sudo systemctl enable --now kindle-to-onenote.timer
systemctl list-timers kindle-to-onenote.timer   # verify
journalctl -u kindle-to-onenote.service -f       # watch logs
```

### Windows — Task Scheduler

Create a Basic Task that runs:

```
Program:   C:\path\to\scribe-to-onenote\.venv\Scripts\python.exe
Arguments: C:\path\to\scribe-to-onenote\kindle_to_onenote.py
Start in:  C:\path\to\scribe-to-onenote
```

Trigger it every 15 minutes (or your preference). Run
`python kindle_to_onenote.py --login` manually once first so the device-code
sign-in can complete.

### Log rotation (Linux)

A sample is provided in [`deploy/logrotate.example`](deploy/logrotate.example).
Install it as `/etc/logrotate.d/kindle` and adjust the path.

---

## Configuration reference

All configuration is done with environment variables — there are **no settings
to edit inside the script**. Each variable can be set in your shell/scheduler
**or** in the `.env` file next to the script (copied from
[`.env.example`](.env.example)).

**Precedence (highest wins):**

```
CLI flag  >  real environment variable  >  .env file  >  built-in default
```

So a value already exported in the environment overrides the same key in
`.env`, and a command-line flag (e.g. `--folder`) overrides both.

### Required (2)

You must set these two — everything else has a working default. The script
exits with a clear message if either is missing.

| Variable | Description |
|---|---|
| `KINDLE_CLIENT_ID` | Azure app **Application (client) ID** (from [Azure setup](#1-create-an-azure-app-registration)) |
| `KINDLE_SECTION_ID` | OneNote **section** ID to create pages in (find it with `--list-sections`) |

### Optional (defaults shown)

| Variable | Default | Flag | Description |
|---|---|---|---|
| `KINDLE_TENANT_ID` | `consumers` | — | `consumers` for personal accounts; tenant ID for work/school |
| `KINDLE_SENDER` | `do-not-reply@amazon.com` | — | Sender address used to find Kindle emails |
| `KINDLE_DEST_FOLDER` | `Kindle Scribe` | `--folder` | Mail folder for processed emails (created if missing) |
| `KINDLE_TAG1` | `Kindle Scribe` | `--tag1` | Category added when a note is processed |
| `KINDLE_TAG2` | `Uploaded to OneNote` | `--tag2` | Category added after the page is confirmed (duplicate guard) |
| `KINDLE_MANAGE_CATEGORIES` | `false` | `--manage-categories` | Create the categories (colored) in the master list; needs `MailboxSettings.ReadWrite` |
| `KINDLE_TAG1_COLOR` | `preset7` | — | Outlook preset color for tag 1 (only when managing categories) |
| `KINDLE_TAG2_COLOR` | `preset4` | — | Outlook preset color for tag 2 (only when managing categories) |
| `KINDLE_TIMEZONE` | `America/New_York` | — | IANA timezone for titles/expiry |
| `KINDLE_EXPIRY_DAYS` | `7` | — | Amazon link lifetime shown on the page |
| `KINDLE_MAX_PER_RUN` | `25` | — | Max emails processed per run |
| `KINDLE_HTTP_TIMEOUT` | `60` | — | HTTP timeout (seconds) |
| `KINDLE_MAX_RETRIES` | `4` | — | Retry attempts for throttled/transient HTTP errors |
| `KINDLE_TOKEN_CACHE` | `./token_cache.json` | — | MSAL token cache path |

> **Scheduling tip:** cron/systemd often don't load your `.env` working
> directory the way an interactive shell does. Either rely on `.env` (the
> script loads it from next to itself, using absolute paths) or set the
> variables explicitly in the unit/crontab.

---

## Security notes

- `token_cache.json` contains a **refresh token** for your account. It is
  git-ignored — keep it private and never commit it.
- `.env` holds your IDs and is git-ignored.
- The script uses **delegated** permissions scoped to *your* mailbox and
  OneNote only. No client secret is stored.

---

## Troubleshooting

- **`CLIENT_ID is not configured` / `SECTION_ID is not configured`** — fill in
  `.env` (see [Configure](#3-configure)).
- **No emails processed** — the script only looks in the **Inbox** folder, so
  confirm the Kindle emails are there (not already filed by a rule) and that
  `KINDLE_SENDER` matches the actual sender. (It matches the sender with an
  exact `$filter`, falling back to `$search` if the mailbox rejects it.)
- **PDF renders but text is missing** — the email had no OCR text file. Enable
  "Convert to text (OCR)" when sending from the Kindle to get a `Download text
  file` link.
- **`413` / "request too large"** — the note exceeds Graph's ~4 MB page-creation
  limit (see [Known limitation](#known-limitation-4-mb-page-size)). Split the
  note into smaller PDFs and re-send.
- **Auth prompts every run** — the token cache isn't being saved/found. Check
  the script can write `token_cache.json` (or set `KINDLE_TOKEN_CACHE` to a
  writable path), and that the same path is used each run (matters for cron —
  use absolute paths).
- **`403`/permission errors** — make sure `Notes.ReadWrite` and
  `Mail.ReadWrite` delegated permissions are added and that you consented to
  them during the first sign-in. If you use `--manage-categories`, also add
  `MailboxSettings.ReadWrite` and re-run `--login`.

---

## Project layout

```
scribe-to-onenote/
├── kindle_to_onenote.py        # the script
├── requirements.txt            # Python dependencies
├── .env.example                # copy to .env and fill in
├── .gitignore
├── deploy/
│   ├── kindle-to-onenote.service
│   ├── kindle-to-onenote.timer
│   ├── crontab.example
│   └── logrotate.example
├── LICENSE
└── README.md
```

## AI disclosure & disclaimer

This project was created with the help of AI tools. The original script was
generated with **OpenAI's ChatGPT**, and it was later reviewed, refactored,
documented, and extended with **Anthropic's Claude** (via Claude Code). The
ideas, requirements, testing, and final decisions are the maintainer's; the AI
tools were used to draft and improve the code and docs.

A few things to keep in mind:

- **Review before you run.** AI-generated code can contain mistakes or make
  assumptions that don't fit your setup. Read the script and understand what it
  does — especially that it reads your mailbox, downloads files, and writes to
  your OneNote — before pointing it at a real account.
- **Test first.** Use `--dry-run` and a non-critical OneNote section/notebook
  until you're confident in the behavior.
- **Third-party services.** This tool talks to Microsoft Graph and Amazon's
  Kindle email links. It isn't affiliated with, endorsed by, or supported by
  Microsoft, Amazon, OpenAI, or Anthropic. Kindle, OneNote, Outlook, ChatGPT,
  and Claude are trademarks of their respective owners.
- **No warranty.** As stated in the [MIT License](LICENSE), the software is
  provided "as is", without warranty of any kind. You are responsible for how
  you use it and for keeping your own data backed up.

Contributions and fixes are welcome regardless of how the code was authored.

## License

[MIT](LICENSE)
