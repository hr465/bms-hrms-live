// Outgoing email via Gmail SMTP. Each company can supply its own Gmail sender (set at
// onboarding); if it hasn't, we fall back to the platform-wide GMAIL_USER/GMAIL_APP_PASSWORD.
// If neither is configured, sendMail silently no-ops (logs a warning) so the app keeps working.
const nodemailer = require("nodemailer");

const defaultUser = process.env.GMAIL_USER;
const defaultPass = process.env.GMAIL_APP_PASSWORD;
const fromName = process.env.MAIL_FROM_NAME || "BMS HRMS";

const transporterCache = new Map(); // keyed by gmail user

function getTransporter(user, pass) {
  if (!user || !pass) return null;
  if (transporterCache.has(user)) return transporterCache.get(user);
  const t = nodemailer.createTransport({ service: "gmail", auth: { user, pass } });
  transporterCache.set(user, t);
  return t;
}

// `sender` is an optional {smtp_user, smtp_pass} — typically a company row. Falls back to the
// platform default account when the company hasn't configured its own.
async function sendMail(to, subject, html, sender, attachments) {
  if (!to) return;
  const user = sender?.smtp_user || defaultUser;
  const pass = sender?.smtp_pass || defaultPass;
  const transporter = getTransporter(user, pass);
  if (!transporter) {
    console.warn(`[mail skipped, not configured] to=${to} subject=${subject}`);
    return;
  }
  try {
    await transporter.sendMail({ from: `"${sender?.name ? sender.name + " · " : ""}${fromName}" <${user}>`, to, subject, html, attachments });
    return true;
  } catch (e) {
    console.error("Failed to send email to", to, e.message);
  }
}

function layout(title, bodyHtml) {
  return `<div style="font-family:Arial,sans-serif;max-width:560px;margin:auto;padding:24px;color:#0f172a">
  <div style="font-size:18px;font-weight:700;color:#4f46e5;margin-bottom:18px">BMS Enterprise HRMS</div>
  <h2 style="margin:0 0 14px">${title}</h2>
  ${bodyHtml}
  <div style="margin-top:28px;padding-top:14px;border-top:1px solid #e5e7eb;font-size:12px;color:#64748b">This is an automated message from your organization's HRMS. Please do not reply to this email.</div>
  </div>`;
}

module.exports = { sendMail, layout };
