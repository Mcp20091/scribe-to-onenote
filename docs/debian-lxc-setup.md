# Setup on a Proxmox Debian LXC (or any Debian/Ubuntu box)

A step-by-step guide to running **scribe-to-onenote** unattended on a fresh
Debian LXC container. The same steps work on a Debian/Ubuntu VM or Raspberry Pi.

Commands assume you're `root` in the container (the Proxmox default). To run as a
non-root user instead, see [Running as a non-root user](#running-as-a-non-root-user).

## Before you start

- Complete the **Azure app registration** first and have your **client ID**
  ready — see the main [README → Create an Azure app registration](../README.md#1-create-an-azure-app-registration).
- The container needs outbound internet access (default on Proxmox).
- You'll get your **OneNote section ID** during setup (step 6) — no need to find
  it beforehand.

## 1. Install system packages

A fresh Debian LXC is minimal, so install Git, Python, venv, and certs/timezone
data:

```bash
apt update
apt install -y git python3 python3-venv python3-pip ca-certificates tzdata nano
```

## 2. Get the code

```bash
cd /opt
git clone https://github.com/Mcp20091/scribe-to-onenote.git
cd scribe-to-onenote
```

## 3. Create a virtualenv and install dependencies

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## 4. Configure

```bash
cp .env.example .env
nano .env
```

Set **`KINDLE_CLIENT_ID`** to your Azure app's client ID. Leave
`KINDLE_SECTION_ID` for the next step. (Optional: change `KINDLE_TIMEZONE` from
`America/New_York`.) Save and exit: `Ctrl+O`, `Enter`, `Ctrl+X`.

## 5. Sign in (one time, device-code flow)

```bash
python kindle_to_onenote.py --login
```

It prints a URL and a code — open the URL on your phone or laptop, enter the
code, and approve. No browser is needed inside the container. This writes
`token_cache.json` next to the script; later runs refresh the token silently.

## 6. Find your OneNote section ID

```bash
python kindle_to_onenote.py --list-sections
```

Copy the `id` of the section you want, then put it in `.env`:

```bash
nano .env      # set KINDLE_SECTION_ID=<the id you copied>
```

## 7. Test it

Send a note from your Kindle first so there's something in the Inbox, then:

```bash
python kindle_to_onenote.py --dry-run --verbose   # downloads + reports, no changes
python kindle_to_onenote.py                        # real run: creates page, tags, moves email
```

## 8. Schedule it (systemd timer, every 15 minutes)

Create the two unit files (paths are already correct for `/opt`):

```bash
cat > /etc/systemd/system/kindle-to-onenote.service <<'EOF'
[Unit]
Description=Import Kindle Scribe notes into OneNote
After=network-online.target
Wants=network-online.target

[Service]
Type=oneshot
WorkingDirectory=/opt/scribe-to-onenote
ExecStart=/opt/scribe-to-onenote/.venv/bin/python /opt/scribe-to-onenote/kindle_to_onenote.py
EOF

cat > /etc/systemd/system/kindle-to-onenote.timer <<'EOF'
[Unit]
Description=Run Kindle->OneNote import every 15 minutes

[Timer]
OnBootSec=2min
OnUnitActiveSec=15min
Persistent=true

[Install]
WantedBy=timers.target
EOF
```

Enable and start:

```bash
systemctl daemon-reload
systemctl enable --now kindle-to-onenote.timer
systemctl list-timers kindle-to-onenote.timer     # confirm it's scheduled
```

Watch a run live:

```bash
journalctl -u kindle-to-onenote.service -f
```

That's it — new Kindle emails are now imported every 15 minutes, unattended.

### Prefer cron?

Instead of the systemd timer, add this to `crontab -e` (use absolute paths):

```cron
*/15 * * * * /opt/scribe-to-onenote/.venv/bin/python /opt/scribe-to-onenote/kindle_to_onenote.py >> /opt/scribe-to-onenote/kindle.log 2>&1
```

## Maintenance

- **Re-authenticate** (if you change your Microsoft password, revoke the app, or
  it sits idle past the ~90-day refresh-token window): rerun
  `python /opt/scribe-to-onenote/kindle_to_onenote.py --login`. The scheduled
  runs are non-interactive and will exit with a clear "re-run `--login`" message
  if the token can't be refreshed — they won't hang.
- **Update the code:**
  ```bash
  cd /opt/scribe-to-onenote
  git pull
  .venv/bin/pip install -r requirements.txt   # in case dependencies changed
  ```
- **Logs:** `journalctl -u kindle-to-onenote.service` (systemd) or the
  `kindle.log` file (cron). For cron, see
  [`deploy/logrotate.example`](../deploy/logrotate.example).

## Running as a non-root user

A dedicated user is tidier than running as root:

```bash
adduser --system --group --home /opt/scribe-to-onenote kindle    # or use an existing user
chown -R kindle:kindle /opt/scribe-to-onenote
```

Then do steps 3–7 as that user (`sudo -u kindle -H bash`), and add
`User=kindle` under `[Service]` in the systemd unit. Make sure the user can
read/write `token_cache.json` and `.env` in `/opt/scribe-to-onenote`.

## Troubleshooting

See the [main README → Troubleshooting](../README.md#troubleshooting) and
[Headless / LXC / VM notes](../README.md#headless--lxc--vm-notes).
