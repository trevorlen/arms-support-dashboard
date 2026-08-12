# ARMS Support Portal

Internal support tools, deployed as an [Azure Static Web App](https://learn.microsoft.com/azure/static-web-apps/) on every push to `main`.

## Tools

| Tool | Path | What it does |
|------|------|--------------|
| Email Send | `/email/` | Paste an HTML email, preview it (desktop/mobile), and send it via Microsoft Graph |

## How Email Send works

- **Frontend** (`email/index.html`): a static page with a compose form, live sandboxed preview, drag-and-drop for `.html` files, and lint warnings for things email clients don't support (`<script>` tags, linked stylesheets).
- **API** (`api/src/functions/send-email.js`): a Static Web Apps managed Azure Function. It validates the request, then calls Microsoft Graph `sendMail` as a mailbox in the tenant (client-credentials flow). The signed-in portal user is set as the **Reply-To** address, and sent mail is saved to the sender mailbox's Sent Items.
- **Auth** (`staticwebapp.config.json`): the whole site and the API require a signed-in user (Entra ID login). The API additionally rejects users outside `ALLOWED_USER_DOMAIN` if that setting is configured.

## One-time setup

The frontend deploys as-is, but sending requires an Entra app registration and a few app settings.

### 1. Create an app registration (Entra ID)

1. Entra ID → **App registrations** → **New registration** (e.g. `arms-support-portal-mailer`), single tenant.
2. **API permissions** → Add → Microsoft Graph → **Application permissions** → `Mail.Send` → **Grant admin consent**.
3. **Certificates & secrets** → new client secret. Copy the value.

### 2. Restrict which mailbox the app can send from (recommended)

Application `Mail.Send` can send as *any* mailbox in the tenant by default. Scope it to just the sending mailbox with an Exchange Online application access policy:

```powershell
New-DistributionGroup -Name "Support Portal Mailer" -Type Security -Members support@edge10group.com
New-ApplicationAccessPolicy -AppId <client-id> -PolicyScopeGroupId "Support Portal Mailer" -AccessRight RestrictAccess
```

### 3. Configure the Static Web App

Azure Portal → your Static Web App → **Environment variables** (a.k.a. Configuration → Application settings):

| Setting | Value |
|---------|-------|
| `GRAPH_TENANT_ID` | Directory (tenant) ID |
| `GRAPH_CLIENT_ID` | Application (client) ID |
| `GRAPH_CLIENT_SECRET` | The client secret |
| `MAIL_SENDER` | Mailbox to send from, e.g. `support@edge10group.com` (a shared mailbox works) |
| `ALLOWED_USER_DOMAIN` | *(optional but recommended)* e.g. `edge10group.com` — only signed-in users on this domain may send |
| `MAX_RECIPIENTS` | *(optional)* default `50` |

> **Why `ALLOWED_USER_DOMAIN` matters:** the built-in `aad` login accepts accounts from any tenant. This setting is what limits sending to your own staff. (Alternatively, register a custom Entra auth provider locked to your tenant.)

### 4. Deploy

Merge to `main` — the GitHub Action builds and deploys the app and the `api/` functions together.

## HTML email tips

- Inline your CSS; `<link rel="stylesheet">` and `<script>` are stripped by most email clients (the tool warns about both).
- Host images at public URLs rather than embedding multi-megabyte data URIs — the API rejects bodies over 2 MB.
- Use the **Send test to me** button to check rendering in a real inbox before sending to a list.
