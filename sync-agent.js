/*
 * BMS HRMS — Local Biometric Sync Agent
 * ---------------------------------------------------------------------------
 * Run this on any always-on PC inside the SAME office network as the
 * biometric device. It pulls attendance logs from the device (LAN, ZKTeco/
 * eSSL protocol) and pushes them over the internet to your HRMS server —
 * which can be hosted anywhere (this same machine, your office server, or a
 * cloud server later). This is what makes biometric sync work even when the
 * HRMS is not on the same network as the device.
 *
 * Setup:
 *   1. npm install                 (from the BMS_HRMS_LIVE_v1 folder)
 *   2. Copy sync-agent.config.example.json to sync-agent.config.json and fill it in
 *      (device IP/port, and the device_id + api_key shown on the Biometric page
 *      after you register the device in the HRMS)
 *   3. node sync-agent.js
 *
 * It keeps running and re-syncs every SYNC_INTERVAL_MINUTES (default 5).
 * Stop it with Ctrl+C. Run it as a Windows scheduled task / service for
 * production so it starts automatically and keeps running in the background.
 */
const fs = require("fs");
const path = require("path");
const ZKLib = require("node-zklib");

const CONFIG_PATH = path.join(__dirname, "sync-agent.config.json");
if (!fs.existsSync(CONFIG_PATH)) {
  console.error(
    "Missing sync-agent.config.json.\n" +
    "Copy sync-agent.config.example.json to sync-agent.config.json and fill in your details first."
  );
  process.exit(1);
}
const cfg = JSON.parse(fs.readFileSync(CONFIG_PATH, "utf8"));
const {
  device_ip,
  device_port = 4370,
  server_url,       // e.g. "http://localhost:3000" or "https://hrms.yourcompany.com"
  device_id,
  api_key,
  sync_interval_minutes = 5,
  lookback_days = 4,
} = cfg;

for (const [k, v] of Object.entries({ device_ip, server_url, device_id, api_key })) {
  if (!v) {
    console.error(`sync-agent.config.json is missing "${k}"`);
    process.exit(1);
  }
}

function pad2(n) { return String(n).padStart(2, "0"); }
function localISO(d) {
  return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}T${pad2(d.getHours())}:${pad2(d.getMinutes())}:${pad2(d.getSeconds())}`;
}

async function syncOnce() {
  const stamp = new Date().toLocaleString("en-IN");
  console.log(`[${stamp}] Connecting to device ${device_ip}:${device_port} ...`);
  const zk = new ZKLib(device_ip, Number(device_port), 10000, 4000);
  try {
    await zk.createSocket();
    const result = await zk.getAttendances();
    const allLogs = result?.data || [];
    const cutoff = new Date();
    cutoff.setHours(0, 0, 0, 0);
    cutoff.setDate(cutoff.getDate() - (lookback_days - 1));
    const recent = allLogs.filter(l => l.recordTime instanceof Date && l.recordTime >= cutoff);
    console.log(`  Device has ${allLogs.length} total logs, ${recent.length} within last ${lookback_days} day(s). Pushing to server...`);

    const records = recent
      .filter(l => l.deviceUserId && l.recordTime instanceof Date)
      .map(l => ({ biometric_id: String(l.deviceUserId).trim(), punch_time: localISO(l.recordTime) }));

    const res = await fetch(`${server_url.replace(/\/$/, "")}/api/biometric/ingest`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ device_id, api_key, records }),
    });
    const data = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(data.error || `Server responded ${res.status}`);
    console.log(`  Pushed OK — ${data.matched} matched to employees, ${data.unmatched} unmatched biometric IDs.`);
  } catch (e) {
    console.error(`  FAILED: ${e?.err?.code || e?.message || e}`);
  } finally {
    try { await zk.disconnect(); } catch {}
  }
}

console.log("BMS HRMS Sync Agent starting.");
console.log(`Device: ${device_ip}:${device_port}  ->  Server: ${server_url}  (every ${sync_interval_minutes} min)`);
syncOnce();
setInterval(syncOnce, sync_interval_minutes * 60 * 1000);
