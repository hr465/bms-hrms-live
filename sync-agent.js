/*
 * BMS HRMS — Local Biometric Sync Agent
 * ---------------------------------------------------------------------------
 * Run this on an always-on PC inside the SAME office network as the biometric
 * device (for example the Udaipur office). It reads attendance punches from the
 * device over the LAN (ZKTeco / eSSL protocol) and pushes them over the
 * internet to your HRMS, which can be hosted anywhere. HR users in any city
 * then see the attendance in the portal without doing anything.
 *
 * Setup:
 *   1. Install Node.js 22 or newer on the office PC.
 *   2. Copy this folder to the PC and run:  npm install
 *   3. Copy sync-agent.config.example.json to sync-agent.config.json and fill it in
 *      (use "Sync Agent Config" on the Biometric page of the HRMS to get the values).
 *   4. Double-click start-agent.bat  (it restarts automatically if it stops).
 *      To start on boot, place a shortcut to start-agent.bat in the Windows Startup folder.
 *
 * Behaviour:
 *   - First run imports the last `lookback_days` days (default 30).
 *   - Later runs send only punches newer than the last successful sync (with a small
 *     overlap), so each cycle is fast and light on the device and the server.
 *   - If the internet or the server is down, nothing is lost: the next successful
 *     cycle sends everything that was missed.
 */
const fs = require("fs");
const path = require("path");
const ZKLib = require("node-zklib");

const CONFIG_PATH = path.join(__dirname, "sync-agent.config.json");
const STATE_PATH = path.join(__dirname, "sync-agent.state.json");
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
  server_url,       // e.g. "https://bms-hrms-live.onrender.com"
  device_id,
  api_key,
  sync_interval_minutes = 10,
  lookback_days = 30,
  batch_size = 500,
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
function loadState() {
  try { return JSON.parse(fs.readFileSync(STATE_PATH, "utf8")); } catch { return {}; }
}
function saveState(s) {
  try { fs.writeFileSync(STATE_PATH, JSON.stringify(s, null, 2)); } catch (e) { console.error("  Could not save state:", e.message); }
}
async function push(records) {
  const url = `${server_url.replace(/\/$/, "")}/api/biometric/ingest`;
  const res = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ device_id, api_key, records }),
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(data.error || `Server responded ${res.status}`);
  return data;
}

let running = false;
async function syncOnce() {
  if (running) return;
  running = true;
  const stamp = new Date().toLocaleString("en-IN");
  console.log(`[${stamp}] Connecting to device ${device_ip}:${device_port} ...`);
  const zk = new ZKLib(device_ip, Number(device_port), 10000, 4000);
  try {
    await zk.createSocket();
    const result = await zk.getAttendances();
    const allLogs = (result?.data || []).filter(l => l.deviceUserId && l.recordTime instanceof Date);

    const state = loadState();
    const cutoff = new Date();
    cutoff.setHours(0, 0, 0, 0);
    cutoff.setDate(cutoff.getDate() - (lookback_days - 1));
    let since = cutoff;
    if (state.last_punch) {
      const overlap = new Date(new Date(state.last_punch).getTime() - 24 * 60 * 60 * 1000);
      if (!isNaN(overlap) && overlap > since) since = overlap;
    }

    const fresh = allLogs.filter(l => l.recordTime >= since);
    console.log(`  Device has ${allLogs.length} logs; ${fresh.length} to send (since ${localISO(since)}).`);

    const records = fresh.map(l => ({ biometric_id: String(l.deviceUserId).trim(), punch_time: localISO(l.recordTime) }));
    let matched = 0, unmatched = 0;
    for (let i = 0; i < records.length; i += batch_size) {
      const r = await push(records.slice(i, i + batch_size));
      matched += r.matched; unmatched += r.unmatched;
    }
    if (!records.length) await push([]); // heartbeat so the portal shows the agent as online

    if (allLogs.length) {
      const newest = allLogs.reduce((m, l) => (l.recordTime > m ? l.recordTime : m), new Date(0));
      saveState({ last_punch: localISO(newest), last_success: new Date().toISOString() });
    }
    console.log(`  Sent OK — ${matched} matched to employees, ${unmatched} unmatched biometric IDs.`);
  } catch (e) {
    console.error(`  FAILED (will retry next cycle): ${e?.err?.code || e?.message || e}`);
  } finally {
    try { await zk.disconnect(); } catch {}
    running = false;
  }
}

console.log("BMS HRMS Sync Agent starting.");
console.log(`Device: ${device_ip}:${device_port}  ->  Server: ${server_url}  (every ${sync_interval_minutes} min)`);
syncOnce();
setInterval(syncOnce, sync_interval_minutes * 60 * 1000);
