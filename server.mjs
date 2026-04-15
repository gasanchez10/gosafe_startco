/**
 * Static dashboard + POST /api/send-outreach (MailerLite, server-side only).
 */
import http from "node:http";
import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import {
  buildOutreachHtml,
  createCampaign,
  createGroup,
  ensureGroup,
  getAccountSender,
  getMailerLiteToken,
  scheduleCampaignInstant,
  upsertSubscriber,
} from "./mailerlite.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, "public");
const TEMPLATE = path.join(__dirname, "gosafe_vc_outreach_email.html");
const PORT = Number(process.env.PORT || 8080);
const DEFAULT_SUBJECT =
  "GoSafe AI - healthtech pre-anestesico (Startco / 15 min si hay fit)";
const GROUP_NAME =
  process.env.MAILERLITE_GROUP_NAME?.trim() ||
  "GoSafe VC Outreach (Startco 2026)";

function calendlyBaseForPreview() {
  const u = process.env.CALENDLY_URL?.trim() || "";
  if (u.startsWith("http://") || u.startsWith("https://")) return u;
  return "https://calendly.com/";
}

function escapeHtml(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function buildEmailPreviewHtml(req) {
  const reqUrl = req.url || "/email-preview";
  let u;
  try {
    u = new URL(reqUrl, "http://local.preview");
  } catch {
    u = new URL("http://local.preview/email-preview");
  }
  const displayName = (u.searchParams.get("name") || "Carolina").slice(0, 120);
  /** Igual que el envío: sin PUBLIC_BASE_URL el logo es jsDelivr (no forzar host local, Gmail no usa base64). */
  const publicBase = process.env.PUBLIC_BASE_URL?.trim() || "";
  let html = buildOutreachHtml(TEMPLATE, calendlyBaseForPreview(), publicBase);
  html = html.replace(/\{\$name\}/g, escapeHtml(displayName));
  const strip = `<div style="font-family:system-ui,sans-serif;background:#1a1a28;color:#c8ff00;padding:10px 14px;text-align:center;font-size:12px;border-bottom:1px solid #333;">Vista previa (no env&iacute;a MailerLite). Saludo: <strong>${escapeHtml(displayName)}</strong> &mdash; prob&aacute; <code>?name=Tu+nombre</code> en la URL.</div>`;
  html = html.replace(/<body([^>]*)>/i, `<body$1>${strip}`);
  return html;
}

const MIME = {
  ".html": "text/html; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".js": "application/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".png": "image/png",
  ".ico": "image/x-icon",
  ".svg": "image/svg+xml",
  ".woff2": "font/woff2",
};

/** @type {Map<string, { n: number, reset: number }>} */
const rateByIp = new Map();
const rateBatchByIp = new Map();
const RATE_WINDOW_MS = 60 * 60 * 1000;
const RATE_MAX = 12;
const RATE_BATCH_WINDOW_MS = 60 * 60 * 1000;
const RATE_BATCH_MAX = 5;
const BATCH_RECIPIENTS_MAX = 200;

function clientIp(req) {
  const xf = req.headers["x-forwarded-for"];
  if (typeof xf === "string" && xf.length) return xf.split(",")[0].trim();
  return req.socket?.remoteAddress || "unknown";
}

function rateOk(ip) {
  const now = Date.now();
  let e = rateByIp.get(ip);
  if (!e || now > e.reset) {
    e = { n: 0, reset: now + RATE_WINDOW_MS };
    rateByIp.set(ip, e);
  }
  if (e.n >= RATE_MAX) return false;
  e.n += 1;
  return true;
}

function rateBatchOk(ip) {
  const now = Date.now();
  let e = rateBatchByIp.get(ip);
  if (!e || now > e.reset) {
    e = { n: 0, reset: now + RATE_BATCH_WINDOW_MS };
    rateBatchByIp.set(ip, e);
  }
  if (e.n >= RATE_BATCH_MAX) return false;
  e.n += 1;
  return true;
}

function safeFilePath(base, relPath) {
  const decoded = decodeURIComponent((relPath || "").split("?")[0]);
  const rel = path.normalize(decoded).replace(/^(\.\.(\/|\\|$))+/, "").replace(/^[\\/]+/, "");
  if (!rel) return null;
  const full = path.join(base, rel);
  const baseNorm = path.resolve(base) + path.sep;
  const fullNorm = path.resolve(full);
  if (!fullNorm.startsWith(baseNorm)) return null;
  return fullNorm;
}

function json(res, status, obj) {
  res.writeHead(status, {
    "Content-Type": "application/json; charset=utf-8",
    "Cache-Control": "no-store",
  });
  res.end(JSON.stringify(obj));
}

async function readJsonBody(req, limit = 400_000) {
  const chunks = [];
  let total = 0;
  for await (const chunk of req) {
    total += chunk.length;
    if (total > limit) {
      const err = new Error("body_too_large");
      err.code = "BODY_LIMIT";
      throw err;
    }
    chunks.push(chunk);
  }
  const raw = Buffer.concat(chunks).toString("utf8");
  if (!raw.trim()) return {};
  return JSON.parse(raw);
}

async function handleSendOutreach(req, res) {
  if (req.method === "OPTIONS") {
    res.writeHead(204, {
      "Access-Control-Allow-Methods": "POST, OPTIONS",
      "Access-Control-Allow-Headers": "Content-Type",
      "Access-Control-Max-Age": "86400",
    });
    res.end();
    return;
  }
  if (req.method !== "POST") {
    json(res, 405, { ok: false, error: "method_not_allowed" });
    return;
  }

  const ip = clientIp(req);
  if (!rateOk(ip)) {
    json(res, 429, { ok: false, error: "rate_limited" });
    return;
  }

  const serverSecret = process.env.OUTREACH_API_SECRET?.trim();

  let body;
  try {
    body = await readJsonBody(req);
  } catch (e) {
    const code = /** @type {NodeJS.ErrnoException} */ (e).code;
    if (code === "BODY_LIMIT") {
      json(res, 413, { ok: false, error: "payload_too_large" });
      return;
    }
    json(res, 400, { ok: false, error: "invalid_json" });
    return;
  }

  if (!serverSecret) {
    json(res, 503, {
      ok: false,
      error: "missing_outreach_secret",
      message: "Set OUTREACH_API_SECRET on Railway to enable sends.",
    });
    return;
  }

  const given = String(body.secret || "").trim();
  if (given !== serverSecret) {
    json(res, 403, { ok: false, error: "forbidden" });
    return;
  }

  const mlToken = getMailerLiteToken();
  if (!mlToken) {
    json(res, 503, {
      ok: false,
      error: "missing_mailerlite_key",
      message: "Set MAILERLITE_API_KEY or MAILES_API_KEY on Railway.",
    });
    return;
  }

  const email = String(body.email || "").trim();
  const name = String(body.name || "Contacto").trim();
  const subject = String(body.subject || DEFAULT_SUBJECT).trim().slice(0, 250);
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    json(res, 400, { ok: false, error: "invalid_email" });
    return;
  }

  try {
    const calendly = process.env.CALENDLY_URL || "https://calendly.com/";
    const publicBase = process.env.PUBLIC_BASE_URL?.trim() || "";
    const html = buildOutreachHtml(TEMPLATE, calendly, publicBase);
    const groupId = await ensureGroup(mlToken, GROUP_NAME);
    await upsertSubscriber(mlToken, email, name, groupId);
    let fromEmail = process.env.MAILERLITE_FROM_EMAIL?.trim() || "";
    let fromName = process.env.MAILERLITE_FROM_NAME?.trim() || "";
    if (!fromEmail) {
      const s = await getAccountSender(mlToken);
      fromEmail = s.fromEmail;
      fromName = fromName || s.fromName;
    }
    if (!fromEmail) {
      json(res, 502, { ok: false, error: "missing_sender_email" });
      return;
    }
    const campaignId = await createCampaign(mlToken, {
      groupId,
      html,
      fromEmail,
      fromName: fromName || "Go Safe AI",
      subject,
    });
    await scheduleCampaignInstant(mlToken, campaignId);
    json(res, 200, {
      ok: true,
      campaignId,
      groupName: GROUP_NAME,
      email,
    });
  } catch (e) {
    const st =
      typeof e === "object" &&
      e !== null &&
      "status" in e &&
      typeof /** @type {{ status: number }} */ (e).status === "number"
        ? /** @type {{ status: number }} */ (e).status
        : 502;
    const status = st >= 400 && st < 600 ? st : 502;
    const msg = /** @type {Error} */ (e).message || "mailerlite_error";
    json(res, status, {
      ok: false,
      error: "mailerlite_failed",
      message: msg,
    });
  }
}

const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

function normalizeBatchRecipients(raw) {
  if (!Array.isArray(raw)) {
    return { error: "recipients debe ser un array [{ email, name }, ...]" };
  }
  if (raw.length === 0) {
    return { error: "recipients vacío" };
  }
  if (raw.length > BATCH_RECIPIENTS_MAX) {
    return { error: `máximo ${BATCH_RECIPIENTS_MAX} contactos por envío` };
  }
  const list = [];
  const seen = new Set();
  for (const row of raw) {
    const email = String(row.email || "")
      .trim()
      .toLowerCase();
    const name = String(row.name || "").trim() || "Contacto";
    if (!EMAIL_RE.test(email)) {
      return { error: `email inválido: ${email}` };
    }
    if (seen.has(email)) continue;
    seen.add(email);
    list.push({ email, name });
  }
  if (list.length === 0) {
    return { error: "ningún email válido" };
  }
  return { list };
}

async function handleSendOutreachBatch(req, res) {
  if (req.method === "OPTIONS") {
    res.writeHead(204, {
      "Access-Control-Allow-Methods": "POST, OPTIONS",
      "Access-Control-Allow-Headers": "Content-Type",
      "Access-Control-Max-Age": "86400",
    });
    res.end();
    return;
  }
  if (req.method !== "POST") {
    json(res, 405, { ok: false, error: "method_not_allowed" });
    return;
  }

  const ip = clientIp(req);
  if (!rateBatchOk(ip)) {
    json(res, 429, { ok: false, error: "rate_limited" });
    return;
  }

  let body;
  try {
    body = await readJsonBody(req, 400_000);
  } catch (e) {
    const code = /** @type {NodeJS.ErrnoException} */ (e).code;
    if (code === "BODY_LIMIT") {
      json(res, 413, { ok: false, error: "payload_too_large" });
      return;
    }
    json(res, 400, { ok: false, error: "invalid_json" });
    return;
  }

  const serverSecret = process.env.OUTREACH_API_SECRET?.trim();
  if (!serverSecret) {
    json(res, 503, {
      ok: false,
      error: "missing_outreach_secret",
      message: "Set OUTREACH_API_SECRET on Railway to enable sends.",
    });
    return;
  }

  if (String(body.secret || "").trim() !== serverSecret) {
    json(res, 403, { ok: false, error: "forbidden" });
    return;
  }

  const norm = normalizeBatchRecipients(body.recipients);
  if (norm.error) {
    json(res, 400, { ok: false, error: "invalid_recipients", message: norm.error });
    return;
  }
  const recipients = norm.list;

  const mlToken = getMailerLiteToken();
  if (!mlToken) {
    json(res, 503, {
      ok: false,
      error: "missing_mailerlite_key",
      message: "Set MAILERLITE_API_KEY or MAILES_API_KEY on Railway.",
    });
    return;
  }

  try {
    const calendly = process.env.CALENDLY_URL || "https://calendly.com/";
    const publicBase = process.env.PUBLIC_BASE_URL?.trim() || "";
    const html = buildOutreachHtml(TEMPLATE, calendly, publicBase);
    const ts = new Date().toISOString().replace(/[:.]/g, "-").slice(0, 19);
    const groupName = `${GROUP_NAME} — batch ${ts}`;
    const groupId = await createGroup(mlToken, groupName);
    for (const r of recipients) {
      await upsertSubscriber(mlToken, r.email, r.name, groupId);
    }
    let fromEmail = process.env.MAILERLITE_FROM_EMAIL?.trim() || "";
    let fromName = process.env.MAILERLITE_FROM_NAME?.trim() || "";
    if (!fromEmail) {
      const s = await getAccountSender(mlToken);
      fromEmail = s.fromEmail;
      fromName = fromName || s.fromName;
    }
    if (!fromEmail) {
      json(res, 502, { ok: false, error: "missing_sender_email" });
      return;
    }
    const subject = String(body.subject || DEFAULT_SUBJECT).trim().slice(0, 250);
    const campaignId = await createCampaign(mlToken, {
      groupId,
      html,
      fromEmail,
      fromName: fromName || "Go Safe AI",
      subject,
    });
    await scheduleCampaignInstant(mlToken, campaignId);
    json(res, 200, {
      ok: true,
      campaignId,
      groupName,
      recipientCount: recipients.length,
    });
  } catch (e) {
    const st =
      typeof e === "object" &&
      e !== null &&
      "status" in e &&
      typeof /** @type {{ status: number }} */ (e).status === "number"
        ? /** @type {{ status: number }} */ (e).status
        : 502;
    const status = st >= 400 && st < 600 ? st : 502;
    const msg = /** @type {Error} */ (e).message || "mailerlite_error";
    json(res, status, {
      ok: false,
      error: "mailerlite_failed",
      message: msg,
    });
  }
}

const server = http.createServer(async (req, res) => {
  const pathname = (req.url || "/").split("?")[0];
  if (pathname === "/api/send-outreach") {
    await handleSendOutreach(req, res);
    return;
  }
  if (pathname === "/api/send-outreach-batch") {
    await handleSendOutreachBatch(req, res);
    return;
  }
  if (pathname === "/email-preview" && req.method === "GET") {
    try {
      const html = buildEmailPreviewHtml(req);
      res.writeHead(200, {
        "Content-Type": "text/html; charset=utf-8",
        "Cache-Control": "no-store",
      });
      res.end(html);
    } catch (e) {
      res.writeHead(500);
      res.end("Preview error");
    }
    return;
  }

  try {
    const urlPath = req.url === "/" || req.url === "" ? "/index.html" : req.url || "/index.html";
    const rel = urlPath.replace(/^\/+/, "") || "index.html";
    const filepath = safeFilePath(ROOT, rel);
    if (!filepath) {
      res.writeHead(403);
      res.end("Forbidden");
      return;
    }
    const data = await fs.readFile(filepath);
    const ext = path.extname(filepath).toLowerCase();
    res.writeHead(200, { "Content-Type": MIME[ext] || "application/octet-stream" });
    res.end(data);
  } catch (e) {
    const code = /** @type {NodeJS.ErrnoException} */ (e).code;
    if (code === "ENOENT") {
      res.writeHead(404);
      res.end("Not found");
    } else {
      res.writeHead(500);
      res.end("Server error");
    }
  }
});

server.listen(PORT, "0.0.0.0", () => {
  console.log(`listening on ${PORT}`);
});
