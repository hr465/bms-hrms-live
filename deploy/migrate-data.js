// One-time copy of every table from the old database (SRC_URL) to the new one (DEST_URL).
// The new database must already have its tables: start the portal once against it first.
// Usage: SRC_URL=... DEST_URL=... node deploy/migrate-data.js
const { Pool } = require("pg");
const src = new Pool({ connectionString: process.env.SRC_URL, ssl: /sslmode=require/.test(process.env.SRC_URL || "") ? { rejectUnauthorized: false } : undefined });
const dst = new Pool({ connectionString: process.env.DEST_URL });
const q = id => '"' + id.replace(/"/g, '""') + '"';

(async () => {
  const tables = (await src.query("select table_name from information_schema.tables where table_schema='public' and table_type='BASE TABLE' order by 1")).rows.map(r => r.table_name);
  const have = new Set((await dst.query("select table_name from information_schema.tables where table_schema='public'")).rows.map(r => r.table_name));
  for (const t of tables) {
    if (!have.has(t)) { console.log("skip (not in new database):", t); continue; }
    const cols = (await dst.query("select column_name from information_schema.columns where table_schema='public' and table_name=$1", [t])).rows.map(r => r.column_name);
    const srcCols = new Set((await src.query("select column_name from information_schema.columns where table_schema='public' and table_name=$1", [t])).rows.map(r => r.column_name));
    const use = cols.filter(c => srcCols.has(c));
    await dst.query(`truncate table ${q(t)} restart identity cascade`);
    const rows = (await src.query(`select ${use.map(q).join(",")} from ${q(t)}`)).rows;
    const per = Math.max(1, Math.floor(60000 / Math.max(1, use.length)));
    for (let i = 0; i < rows.length; i += per) {
      const chunk = rows.slice(i, i + per), vals = [];
      const ph = chunk.map(r => "(" + use.map(c => { vals.push(r[c]); return "$" + vals.length; }).join(",") + ")").join(",");
      await dst.query(`insert into ${q(t)} (${use.map(q).join(",")}) values ${ph}`, vals);
    }
    if (use.includes("id")) await dst.query(`select setval(pg_get_serial_sequence('${t}','id'), greatest(coalesce((select max(id) from ${q(t)}),0),1), (select count(*)>0 from ${q(t)}))`).catch(() => {});
    console.log(t.padEnd(24), rows.length, "rows");
  }
  console.log("Done.");
  process.exit(0);
})().catch(e => { console.error("FAILED:", e.message); process.exit(1); });
