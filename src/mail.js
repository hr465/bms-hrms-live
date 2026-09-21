// Outgoing email via Gmail SMTP. Each company can supply its own Gmail sender (set at
// onboarding); if it hasn't, we fall back to the platform-wide GMAIL_USER/GMAIL_APP_PASSWORD.
// If neither is configured, sendMail silently no-ops (logs a warning) so the app keeps working.
// Every attempt is written to email_log so HR can see what was sent and why something failed.
const nodemailer = require("nodemailer");
const db = require("./db");

const defaultUser = process.env.GMAIL_USER;
const defaultPass = process.env.GMAIL_APP_PASSWORD;
const fromName = process.env.MAIL_FROM_NAME || "BMS HRMS";

const transporterCache = new Map(); // keyed by gmail user + password length (so a changed password is picked up)

function getTransporter(user, pass) {
  if (!user || !pass) return null;
  const key = user + "|" + pass;
  if (transporterCache.has(key)) return transporterCache.get(key);
  const t = nodemailer.createTransport({
    host: "smtp.gmail.com",
    port: 465,
    secure: true,
    family: 4, // hosting platforms often have no IPv6 route (ENETUNREACH)
    auth: { user, pass },
    connectionTimeout: 15000,
    greetingTimeout: 15000,
    socketTimeout: 30000,
  });
  transporterCache.set(key, t);
  return t;
}

async function logEmail(companyId, to, subject, status, error) {
  try {
    await db.prepare("INSERT INTO email_log(company_id,to_addr,subject,status,error) VALUES(?,?,?,?,?)")
      .run(companyId || null, String(to || "").slice(0, 200), String(subject || "").slice(0, 200), status, error ? String(error).slice(0, 400) : null);
  } catch (e) { /* logging must never break sending */ }
}

// `sender` is an optional {smtp_user, smtp_pass, name, company_id}. Falls back to the platform default account.
// Returns {ok, error}.
async function sendMailEx(to, subject, html, sender, attachments) {
  if (!to) return { ok: false, error: "No recipient email address" };
  const user = sender?.smtp_user || defaultUser;
  const pass = sender?.smtp_pass || defaultPass;
  const transporter = getTransporter(user, pass);
  if (!transporter) {
    console.warn(`[mail skipped, not configured] to=${to} subject=${subject}`);
    await logEmail(sender?.company_id, to, subject, "skipped", "No sender Gmail is configured for this company");
    return { ok: false, error: "No sender Gmail is configured" };
  }
  try {
    const from = `"${sender?.name ? sender.name + " · " : ""}${fromName}" <${user}>`;
    if (process.env.MAIL_RELAY_URL && process.env.MAIL_RELAY_SECRET) {
      // Some hosts block outbound SMTP; the relay (on a server that allows it) sends the message for us over HTTPS.
      const r = await fetch(process.env.MAIL_RELAY_URL.replace(/\/$/, "") + "/send", {
        method: "POST",
        headers: { "Content-Type": "application/json", Authorization: "Bearer " + process.env.MAIL_RELAY_SECRET },
        body: JSON.stringify({
          smtp: { user, pass }, from, to, subject, html,
          attachments: (attachments || []).map(a => ({ filename: a.filename, contentType: a.contentType, contentBase64: Buffer.from(a.content).toString("base64") })),
        }),
        signal: AbortSignal.timeout(60000),
      });
      const d = await r.json().catch(() => ({}));
      if (!r.ok || !d.ok) throw new Error(d.error || `Mail relay error ${r.status}`);
    } else {
      await transporter.sendMail({ from, to, subject, html, attachments });
    }
    await logEmail(sender?.company_id, to, subject, "sent", null);
    return { ok: true };
  } catch (e) {
    console.error("Failed to send email to", to, e.message);
    await logEmail(sender?.company_id, to, subject, "failed", `${e.code || ""} ${e.message}`.trim());
    return { ok: false, error: `${e.code || ""} ${e.message}`.trim() };
  }
}

async function sendMail(to, subject, html, sender, attachments) {
  const r = await sendMailEx(to, subject, html, sender, attachments);
  return r.ok ? true : undefined;
}

function layout(title, bodyHtml) {
  return `<div style="font-family:Arial,sans-serif;max-width:560px;margin:auto;padding:24px;color:#0f172a">
  <div style="font-size:18px;font-weight:700;color:#4f46e5;margin-bottom:18px">${process.env.PLATFORM_EMAIL_HEADER || "BMS Enterprise HRMS"}</div>
  <h2 style="margin:0 0 14px">${title}</h2>
  ${bodyHtml}
  <div style="margin-top:28px;padding-top:14px;border-top:1px solid #e5e7eb;font-size:12px;color:#64748b">This is an automated message from your organization's HRMS. Please do not reply to this email.</div>
  </div>`;
}

module.exports = { sendMail, sendMailEx, layout };
