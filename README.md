# Ray DLP — Smart Alerts MVP PoC

Bare-minimum Outlook add-in that intercepts every outbound message and shows a "blocked" dialog. No build step, no backend, no classifier — just the smallest thing that proves the integration works.

## Files

```
manifest.xml      # add-in registration — Outlook reads this to install
commands.html     # required HTML page that loads office.js + launchevent.js
launchevent.js    # the entire DLP logic: always block, show message
icon-64.png       # tile icon
icon-128.png      # high-res icon
```

## What you'll do

1. Sign up for a free Microsoft 365 Developer tenant
2. Host these 5 files at some public HTTPS URL
3. Replace `YOUR_HOST` in `manifest.xml` with that URL
4. Sideload `manifest.xml` into Outlook on the Web
5. Compose + send an email — confirm the block dialog appears

Total time: ~1–2 hours, mostly waiting for the dev-tenant signup.

---

## Step 1 — Dev tenant (10 min)

Go to **https://developer.microsoft.com/microsoft-365/dev-program**, sign in with any personal Microsoft account, pick "Instant sandbox." Wait ~5 minutes for provisioning. Note your admin email (`admin@<tenant>.onmicrosoft.com`) and password.

Log into **https://outlook.office.com** with that admin to confirm the inbox works.

## Step 2 — Hosting (pick one)

The manifest and the JS file have to be reachable over HTTPS. Two easy options:

### Option A — GitHub Pages (recommended, persistent)

```bash
cd /Users/shakedshary/Ray/dlp-poc
git init
git add .
git commit -m "Ray DLP PoC"
gh repo create ray-security/ray-dlp-poc --public --source=. --push
gh repo edit ray-security/ray-dlp-poc --enable-pages --pages-source main
```

Wait ~30 seconds. Your URL is:

```
https://ray-security.github.io/ray-dlp-poc
```

### Option B — ngrok (quick, ephemeral)

```bash
cd /Users/shakedshary/Ray/dlp-poc
python3 -m http.server 8000 &
ngrok http 8000
```

Use the `https://...ngrok-free.app` URL ngrok prints. **The URL changes every restart unless you have a paid ngrok account.**

## Step 3 — Plug your host into the manifest

```bash
# macOS sed — replace ray-security.github.io/ray-dlp-poc with your actual URL
sed -i '' "s|YOUR_HOST|ray-security.github.io/ray-dlp-poc|g" manifest.xml
```

If you used Option A, push the change:

```bash
git add manifest.xml && git commit -m "Set host URL" && git push
```

## Step 4 — Sideload into Outlook on the Web

1. Open **https://outlook.office.com** logged in as your dev-tenant admin
2. Click the gear ⚙ (top-right) → **View all Outlook settings** → **Mail** → **Customize actions**
3. Or directly: **Get Add-ins** (the puzzle-piece icon in the new toolbar)
4. **My add-ins** (left sidebar) → **Add a custom add-in** → **Add from URL**
5. Paste your manifest URL, e.g. `https://ray-security.github.io/ray-dlp-poc/manifest.xml`
6. Confirm any warning about custom add-ins
7. Wait ~30 seconds for the add-in to register

## Step 5 — Test

1. Compose a new email in Outlook on the Web
2. Type something, address it anywhere (your personal Gmail is fine)
3. Hit **Send**

You should see a dialog:

> Blocked by Ray DLP (test mode). All outbound mail is currently dropped. Contact your administrator if you need to send.

The send is cancelled. The draft stays in Drafts.

---

## When this works

You've proven:

- The add-in installs from a public manifest
- Outlook fires `OnMessageSend` when the user clicks Send
- Our handler can block the send
- The user sees a custom dialog

Next iterations on top of this (each is its own small change):

- Read the message body before deciding (`Office.context.mailbox.item.body.getAsync`)
- POST to a real classification API instead of always blocking
- Add policy rules (block if contains regex, block to specific domains, etc.)
- Host on Cloudflare Pages or your CDN instead of GitHub Pages
- Centralized deployment via M365 admin center → Integrated Apps → all users

## Troubleshooting

- **"Add-in can't be loaded"** — usually a URL in the manifest is unreachable. Test each URL in a browser: `https://YOUR_HOST/manifest.xml`, `/commands.html`, `/launchevent.js`, `/icon-64.png`. All must return 200.
- **"Sign in to your Microsoft account"** loop — you used a personal Microsoft account in the dev tenant. Use the `admin@<tenant>.onmicrosoft.com` account instead.
- **Send goes through without a dialog** — the add-in didn't register. Open the add-ins menu and confirm "Ray DLP PoC" is listed under My add-ins. Try removing and re-adding.
- **Dialog shows but a different message** — your browser cached the old `launchevent.js`. Hard-refresh (Cmd+Shift+R) or wait 5 minutes.
