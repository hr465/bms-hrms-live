// Receives a signed request from the portal and delivers it over SMTP.
// It stores nothing: the sender's Gmail account is sent with each request (over HTTPS) and never written to disk.
const http = require("http");
const crypto = require("crypto");
const nodemailer = require("nodemailer");

const PORT = Number(process.env.PORT) || 8025;
const SECRET = process.env.RELAY_SECRET || "";
if (SECRET.length < 24) { console.error("RELAY_SECRET must be set and at least 24 characters long"); process.exit(1); }

const MAX_BODY = 30 * 1024 * 1024;
const transporters = new Map();
function transporter(user, pass) {
  const key = user + "|" + pass;
  if (!transporters.has(key)) {
    transporters.set(key, nodemailer.createTransport({
      host: "smtp.gmail.com", port: 465, secure: true, family: 4,
      auth: { user, pass }, connectionTimeout: 15000, greetingTimeout: 15000, socketTimeout: 40000,
    }));
    if (transporters.size > 50) transporters.delete(transporters.keys().next().value);
  }
  return transporters.get(key);
}
const same = (a, b) => { const x = Buffer.from(String(a)), y = Buffer.from(String(b)); return x.length === y.length && crypto.timingSafeEqual(x, y); };
const reply = (res, code, obj) => { res.writeHead(code, { "Content-Type": "application/json" }); res.end(JSON.stringify(obj)); };

http.createServer((req, res) => {
  if (req.method === "GET" && req.url === "/health") return reply(res, 200, { ok: true });
  if (req.method !== "POST" || req.url !== "/send") return reply(res, 404, { ok: false, error: "Not found" });
  const auth = req.headers.authorization || "";
  if (!auth.startsWith("Bearer ") || !same(auth.slice(7), SECRET)) return reply(res, 401, { ok: false, error: "Unauthorized" });
  let size = 0; const chunks = [];
  req.on("data", c => { size += c.length; if (size > MAX_BODY) { reply(res, 413, { ok: false, error: "Too large" }); req.destroy(); } else chunks.push(c); });
  req.on("end", async () => {
    if (res.writableEnded) return;
    try {
      const m = JSON.parse(Buffer.concat(chunks).toString("utf8"));
      if (!m.smtp?.user || !m.smtp?.pass || !m.to || !m.subject) return reply(res, 400, { ok: false, error: "Missing fields" });
      await transporter(m.smtp.user, m.smtp.pass).sendMail({
        from: m.from || m.smtp.user, to: m.to, subject: m.subject, html: m.html,
        attachments: (m.attachments || []).map(a => ({ filename: a.filename, content: Buffer.from(a.contentBase64, "base64"), contentType: a.contentType })),
      });
      reply(res, 200, { ok: true });
    } catch (e) {
      console.error("send failed:", e.code || "", e.message);
      reply(res, 200, { ok: false, error: `${e.code || ""} ${e.message}`.trim() });
    }
  });
}).listen(PORT, "127.0.0.1", () => console.log("mail relay listening on 127.0.0.1:" + PORT));
