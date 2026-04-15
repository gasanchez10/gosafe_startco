/**
 * MailerLite REST helpers (connect.mailerlite.com) for server-side outreach.
 */

import fs from "node:fs";

const BASE = "https://connect.mailerlite.com/api";

async function mlFetch(path, token, options = {}) {
  const url = `${BASE}${path}`;
  const headers = {
    Authorization: `Bearer ${token}`,
    Accept: "application/json",
    ...(options.headers || {}),
  };
  if (options.body != null) {
    headers["Content-Type"] = "application/json";
  }
  const res = await fetch(url, {
    ...options,
    headers,
  });
  const text = await res.text();
  let data = {};
  try {
    data = text ? JSON.parse(text) : {};
  } catch {
    data = { raw: text };
  }
  if (!res.ok) {
    const err = new Error(data.message || `MailerLite ${res.status}`);
    err.status = res.status;
    err.body = data;
    throw err;
  }
  return data;
}

export function getMailerLiteToken() {
  return (
    process.env.MAILERLITE_API_KEY ||
    process.env.MAILES_API_KEY ||
    process.env.MAILERLITE_TOKEN ||
    ""
  ).trim();
}

export async function getAccountSender(token) {
  const { data } = await mlFetch("/account", token, { method: "GET" });
  return {
    fromEmail: String(data?.sender_email || ""),
    fromName: String(data?.sender_name || "Go Safe AI"),
  };
}

export async function findGroupIdByName(token, name) {
  const { data: groups } = await mlFetch("/groups", token, { method: "GET" });
  for (const g of groups || []) {
    if (g.name === name) return String(g.id);
  }
  return null;
}

export async function createGroup(token, name) {
  const { data } = await mlFetch("/groups", token, {
    method: "POST",
    body: JSON.stringify({ name }),
  });
  return String(data.id);
}

export async function ensureGroup(token, name) {
  const existing = await findGroupIdByName(token, name);
  return existing || createGroup(token, name);
}

export async function upsertSubscriber(token, email, name, groupId) {
  await mlFetch("/subscribers", token, {
    method: "POST",
    body: JSON.stringify({
      email: email.trim(),
      fields: { name: name.trim() },
      groups: [groupId],
      resubscribe: true,
    }),
  });
}

export async function createCampaign(token, { groupId, html, fromEmail, fromName, subject }) {
  const { data } = await mlFetch("/campaigns", token, {
    method: "POST",
    body: JSON.stringify({
      name: `GoSafe VC outreach - ${subject.slice(0, 40)}`,
      type: "regular",
      language_id: 9,
      groups: [groupId],
      emails: [
        {
          subject,
          from_name: fromName,
          from: fromEmail,
          reply_to: fromEmail,
          content: html,
        },
      ],
    }),
  });
  return String(data.id);
}

export async function scheduleCampaignInstant(token, campaignId) {
  await mlFetch(`/campaigns/${campaignId}/schedule`, token, {
    method: "POST",
    body: JSON.stringify({ delivery: "instant" }),
  });
}

/** Gmail blocks data:image/*;base64 in email HTML. Default: jsDelivr (free CDN mirroring public GitHub). */
const DEFAULT_OUTREACH_LOGO =
  "https://cdn.jsdelivr.net/gh/gasanchez10/gosafe_startco@main/go-safe-logo.png";

function resolveLogoSrc(publicBaseUrl) {
  const explicit = (
    process.env.OUTREACH_LOGO_URL ||
    process.env.LOGO_URL ||
    ""
  ).trim();
  if (explicit) return explicit;
  const pub = (publicBaseUrl || "").trim().replace(/\/+$/, "");
  if (pub) return `${pub}/go-safe-logo.png`;
  return DEFAULT_OUTREACH_LOGO;
}

export function buildOutreachHtml(templatePath, calendlyUrl, publicBaseUrl = "") {
  const raw = fs.readFileSync(templatePath, "utf8");
  const cal = (calendlyUrl || "https://calendly.com/").trim().replace(/\/+$/, "");
  let html = raw.replace(/__CALENDLY_URL__/g, `${cal}/`);
  const logoSrc = resolveLogoSrc(publicBaseUrl);
  html = html.replace(/__LOGO_SRC__/g, logoSrc);
  return html;
}
