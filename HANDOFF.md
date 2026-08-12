# Email Send Tool — Deployment Handoff

A self-contained tool for the ARMS support portal: paste HTML email markup, preview it, and send it via Microsoft Graph from a mailbox in your M365 tenant. This document is everything needed to lift it out of `trevorlen/arms-support-dashboard` (branch `claude/support-portal-email-send-fntkpa`) and deploy it inside your full portal repo.

**Architecture in one line:** a static page (`email/index.html`) POSTs to an Azure Static Web Apps *managed* Function (`/api/send-email`), which calls Microsoft Graph `sendMail` using client credentials — no separate Function App, App Service, or third-party email provider needed.

---

## 1. Package contents

| File | What it is | Notes for your repo |
|------|-----------|---------------------|
| `email/index.html` | The tool. Compose form, live sandboxed preview (desktop/mobile widths), drag-and-drop for `.html` files, warnings for `<script>` tags and linked stylesheets. Zero dependencies, no build step. | Copy as-is. Place anywhere under your app root; only the header's "← Support Portal" link (`href="/"`) may need updating. |
| `api/src/functions/send-email.js` | The send endpoint. Azure Functions **Node v4 programming model**, no npm dependencies beyond `@azure/functions` (uses built-in `fetch`). | Copy into your existing `api/` folder if you have one — v4 functions self-register, so no `function.json` needed. |
| `api/package.json` | Declares `@azure/functions ^4.5.0`, Node ≥ 20, and `main: src/functions/*.js`. | If your repo already has an API, just merge the dependency and make sure your `main` glob covers the new file. |
| `api/host.json` | Standard Functions host config (extension bundle `[4.*, 5.0.0)`). | Skip if you already have one. |
| `staticwebapp.config.json` | Requires the `authenticated` role for `/*` and `/api/*`, redirects 401 → Entra ID login, disables GitHub login, sets `platform.apiRuntime: node:20`. | **Merge, don't overwrite**, if your repo already has one. The non-negotiable line is `/api/*` → `allowedRoles: ["authenticated"]` — without it the send endpoint is publicly callable. |
| `index.html` | Minimal portal landing page with an Email Send card. | Only needed if your full repo doesn't already have a portal home. Otherwise just add a link/card pointing to `/email/`. |
| `.github/workflows/...yml` (2 lines changed) | `api_location: "api"` and `output_location: ""`. | Apply the same two values to your repo's SWA workflow. If your portal has a build step, keep your existing `output_location` — but then `email/` must be included in (or copied to) the build output. |

## 2. Copy the files

```
your-repo/
├── email/index.html                      ← new
├── api/
│   ├── host.json                         ← new (or keep yours)
│   ├── package.json                      ← new (or merge)
│   └── src/functions/send-email.js       ← new
├── staticwebapp.config.json              ← new (or merge)
└── .github/workflows/<your-swa>.yml      ← set api_location: "api"
```

Sanity checks after copying:

- Your SWA workflow's `api_location` must point at the folder containing `host.json`.
- If your SWA already deploys a **Bring Your Own Functions** API or uses a non-Node runtime, this function needs adapting — it assumes managed functions on Node 20.
- `platform.apiRuntime: node:20` in `staticwebapp.config.json` must not conflict with an existing value in yours.

## 3. Entra ID app registration (one-time)

The function sends mail using the client-credentials flow, so it needs its own app registration:

1. **Entra ID → App registrations → New registration.** Name it something like `arms-support-portal-mailer`, single tenant. No redirect URI needed.
2. **API permissions → Add a permission → Microsoft Graph → Application permissions → `Mail.Send`**, then **Grant admin consent** (requires a Global Admin or Privileged Role Admin).
3. **Certificates & secrets → New client secret.** Copy the value immediately — it's shown once. Note the expiry and calendar a rotation.
4. From the **Overview** blade, note the **Application (client) ID** and **Directory (tenant) ID**.

## 4. Scope the app to one mailbox (strongly recommended)

Application-level `Mail.Send` can send as **any mailbox in the tenant** by default. Restrict it to just the sending mailbox with an Exchange Online application access policy (run in Exchange Online PowerShell):

```powershell
# One-time: a mail-enabled security group containing only the sending mailbox
New-DistributionGroup -Name "Support Portal Mailer" -Type Security -Members support@edge10group.com

New-ApplicationAccessPolicy -AppId <application-client-id> `
  -PolicyScopeGroupId "Support Portal Mailer" `
  -AccessRight RestrictAccess `
  -Description "Support portal may only send as support@"

# Verify
Test-ApplicationAccessPolicy -AppId <application-client-id> -Identity support@edge10group.com
```

Policies take up to ~30 minutes to apply. A shared mailbox works fine as the sender and needs no license for sending via Graph.

## 5. Static Web App settings

Azure Portal → your Static Web App → **Environment variables** (older portals: Configuration → Application settings):

| Setting | Required | Value |
|---------|----------|-------|
| `GRAPH_TENANT_ID` | ✅ | Directory (tenant) ID |
| `GRAPH_CLIENT_ID` | ✅ | Application (client) ID |
| `GRAPH_CLIENT_SECRET` | ✅ | The client secret value |
| `MAIL_SENDER` | ✅ | Mailbox to send from, e.g. `support@edge10group.com` |
| `ALLOWED_USER_DOMAIN` | recommended | `edge10group.com` — only signed-in users on this domain may send. **Without it, any Microsoft account that can log in can send**, because SWA's built-in `aad` provider is multi-tenant. (The alternative is registering a custom Entra auth provider locked to your tenant.) |
| `MAX_RECIPIENTS` | optional | Default `50` |

The function returns a clear `500` naming any missing setting, so misconfiguration is self-diagnosing.

## 6. Deploy and verify

1. Merge to the branch your SWA workflow deploys from and let the action run.
2. Open the site — you should be bounced to an Entra ID login.
3. Open `/email/` — your email address should appear top-right.
4. Paste any HTML, check the preview renders, toggle Desktop/Mobile.
5. Click **Send test to me** — a real email should arrive from `MAIL_SENDER`, with Reply-To set to your address, and a copy in the sender mailbox's Sent Items.
6. Confirm the endpoint is protected: `curl -X POST https://<site>/api/send-email` from outside a session should return the login redirect/401, never a send.

## 7. API contract

`POST /api/send-email` — requires an authenticated SWA session (cookie); SWA injects `x-ms-client-principal`, which the function also validates.

```json
{
  "to": "jane@example.com, ops@example.com",
  "cc": "boss@example.com",
  "subject": "Weekly Support Report — w/c 10 Aug",
  "html": "<!DOCTYPE html><html>…</html>"
}
```

`to`/`cc` accept a comma/semicolon-separated string **or** an array of strings. Responses:

| Status | Body | Meaning |
|--------|------|---------|
| `200` | `{ "ok": true, "to": [...], "cc": [...], "from": "support@…" }` | Accepted by Graph (Graph returns 202; delivery is async) |
| `400` | `{ "error": "…" }` | Validation: missing recipients/subject/body, invalid address, > `MAX_RECIPIENTS`, body > 2 MB |
| `401` | `{ "error": "Not signed in." }` | No client principal |
| `403` | `{ "error": "…" }` | Signed-in user outside `ALLOWED_USER_DOMAIN` |
| `500` | `{ "error": "…(missing: …)" }` | App settings not configured |
| `502` | `{ "error": "Send failed: …" }` | Token request or Graph `sendMail` failed (Graph's message is passed through) |

## 8. Built-in behaviors and limits

- Sender is always `MAIL_SENDER`; the signed-in user becomes **Reply-To**, so replies go to the person who sent it.
- `saveToSentItems: true` — every send is auditable in the shared mailbox.
- Each send is logged in the Function logs: who sent, as what, to how many, and the subject.
- Recipient cap (`MAX_RECIPIENTS`, default 50) and a 2 MB HTML body limit — host images at public URLs rather than inlining data URIs.
- The preview iframe is sandboxed (`allow-same-origin` only, no scripts), so pasted HTML can't run code in the portal.

## 9. Troubleshooting

| Symptom | Likely cause |
|---------|--------------|
| `500 … missing: GRAPH_…` | App settings not set on the Static Web App (or set on the wrong environment/slot) |
| `502 Send failed: … ErrorAccessDenied … Access to OData is disabled` | Application access policy blocks this mailbox — check `Test-ApplicationAccessPolicy`, or wait out policy propagation |
| `502 … invalid_client` / `AADSTS7000215` | Wrong or expired client secret |
| `502 … ResourceNotFound` for the sender | `MAIL_SENDER` isn't a real mailbox in the tenant (typo, or user has no Exchange mailbox) |
| `403 Only @… accounts may send` | Signed-in account doesn't match `ALLOWED_USER_DOMAIN` |
| API returns HTML login page instead of JSON | Caller has no SWA auth session — expected for anonymous callers |
| Deploy fails: "Failed to find a default file" | Workflow `output_location` doesn't match where `index.html` lands — for a no-build repo it should be `""` |

## 10. Local development (optional)

The [SWA CLI](https://azure.github.io/static-web-apps-cli/) emulates hosting, auth, and the managed API together:

```bash
npm install -g @azure/static-web-apps-cli
cd api && npm install && cd ..
swa start . --api-location api
```

Visit `http://localhost:4280`, use the emulated login (any name/email works locally), and set the `GRAPH_*` / `MAIL_SENDER` variables in `api/local.settings.json` if you want real sends from your machine.
