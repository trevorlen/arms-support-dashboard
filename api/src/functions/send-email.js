const { app } = require('@azure/functions');

const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
const MAX_RECIPIENTS = parseInt(process.env.MAX_RECIPIENTS || '50', 10);
const MAX_HTML_BYTES = 2 * 1024 * 1024;

function getClientPrincipal(request) {
  const header = request.headers.get('x-ms-client-principal');
  if (!header) return null;
  try {
    return JSON.parse(Buffer.from(header, 'base64').toString('utf8'));
  } catch {
    return null;
  }
}

function normalizeRecipients(value) {
  if (!value) return [];
  const list = Array.isArray(value) ? value : String(value).split(/[,;]+/);
  return list.map((e) => String(e).trim()).filter(Boolean);
}

function toGraphRecipients(addresses) {
  return addresses.map((address) => ({ emailAddress: { address } }));
}

async function getGraphToken() {
  const tenantId = process.env.GRAPH_TENANT_ID;
  const params = new URLSearchParams({
    client_id: process.env.GRAPH_CLIENT_ID,
    client_secret: process.env.GRAPH_CLIENT_SECRET,
    scope: 'https://graph.microsoft.com/.default',
    grant_type: 'client_credentials',
  });
  const res = await fetch(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: params,
  });
  const body = await res.json();
  if (!res.ok) {
    throw new Error(`Token request failed (${res.status}): ${body.error_description || body.error || 'unknown error'}`);
  }
  return body.access_token;
}

app.http('send-email', {
  methods: ['POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    const requiredSettings = ['GRAPH_TENANT_ID', 'GRAPH_CLIENT_ID', 'GRAPH_CLIENT_SECRET', 'MAIL_SENDER'];
    const missing = requiredSettings.filter((name) => !process.env[name]);
    if (missing.length) {
      context.error(`Missing app settings: ${missing.join(', ')}`);
      return {
        status: 500,
        jsonBody: { error: `The email service is not configured yet (missing: ${missing.join(', ')}). See the README for setup.` },
      };
    }

    const principal = getClientPrincipal(request);
    if (!principal || !principal.userDetails) {
      return { status: 401, jsonBody: { error: 'Not signed in.' } };
    }

    const allowedDomain = process.env.ALLOWED_USER_DOMAIN;
    if (allowedDomain && !principal.userDetails.toLowerCase().endsWith(`@${allowedDomain.toLowerCase()}`)) {
      context.warn(`Blocked send attempt by ${principal.userDetails} (outside allowed domain ${allowedDomain})`);
      return { status: 403, jsonBody: { error: `Only @${allowedDomain} accounts may send email from this tool.` } };
    }

    let payload;
    try {
      payload = await request.json();
    } catch {
      return { status: 400, jsonBody: { error: 'Request body must be JSON.' } };
    }

    const to = normalizeRecipients(payload.to);
    const cc = normalizeRecipients(payload.cc);
    const subject = String(payload.subject || '').trim();
    const html = String(payload.html || '');

    if (!to.length) return { status: 400, jsonBody: { error: 'At least one "To" recipient is required.' } };
    if (!subject) return { status: 400, jsonBody: { error: 'Subject is required.' } };
    if (!html.trim()) return { status: 400, jsonBody: { error: 'Email HTML body is empty.' } };

    const invalid = [...to, ...cc].filter((e) => !EMAIL_RE.test(e));
    if (invalid.length) {
      return { status: 400, jsonBody: { error: `Invalid email address(es): ${invalid.join(', ')}` } };
    }
    if (to.length + cc.length > MAX_RECIPIENTS) {
      return { status: 400, jsonBody: { error: `Too many recipients (max ${MAX_RECIPIENTS}).` } };
    }
    if (Buffer.byteLength(html, 'utf8') > MAX_HTML_BYTES) {
      return { status: 400, jsonBody: { error: 'Email HTML is larger than 2 MB. Host large images externally instead of inlining them.' } };
    }

    const sender = process.env.MAIL_SENDER;
    const message = {
      subject,
      body: { contentType: 'HTML', content: html },
      toRecipients: toGraphRecipients(to),
    };
    if (cc.length) message.ccRecipients = toGraphRecipients(cc);
    if (EMAIL_RE.test(principal.userDetails)) {
      message.replyTo = toGraphRecipients([principal.userDetails]);
    }

    try {
      const token = await getGraphToken();
      const res = await fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(sender)}/sendMail`, {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${token}`,
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({ message, saveToSentItems: true }),
      });

      if (res.status !== 202) {
        const errBody = await res.text();
        context.error(`Graph sendMail failed (${res.status}): ${errBody}`);
        let detail = `Graph returned ${res.status}.`;
        try {
          detail = JSON.parse(errBody).error?.message || detail;
        } catch {}
        return { status: 502, jsonBody: { error: `Send failed: ${detail}` } };
      }

      context.log(`Email sent by ${principal.userDetails} as ${sender} to ${to.length} recipient(s) (cc: ${cc.length}), subject: "${subject}"`);
      return { status: 200, jsonBody: { ok: true, to, cc, from: sender } };
    } catch (err) {
      context.error(`Send failed: ${err.message}`);
      return { status: 502, jsonBody: { error: `Send failed: ${err.message}` } };
    }
  },
});
