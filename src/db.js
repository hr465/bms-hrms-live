// Thin Postgres wrapper that mimics the better-sqlite3 API used throughout server.js
// (db.prepare(sql).run/get/all), so route handlers only need `await` added — not rewritten.
const { Pool } = require("pg");

const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: process.env.DATABASE_URL && process.env.DATABASE_URL.includes("sslmode=require")
    ? { rejectUnauthorized: false }
    : (process.env.PGSSL === "true" ? { rejectUnauthorized: false } : undefined),
});

// Tables whose primary key isn't `id` — never auto-append RETURNING id for these.
const NO_ID_TABLES = ["sessions", "password_resets", "images"];

function toPgSql(sql) {
  let text = sql.trim();
  // INSERT OR IGNORE ... -> INSERT ... ON CONFLICT DO NOTHING
  const wasInsertIgnore = /^INSERT OR IGNORE/i.test(text);
  if (wasInsertIgnore) text = text.replace(/^INSERT OR IGNORE/i, "INSERT");
  const isInsert = /^INSERT/i.test(text);
  if (wasInsertIgnore && !/ON CONFLICT/i.test(text)) {
    text = text.replace(/;?\s*$/, "") + " ON CONFLICT DO NOTHING";
  }
  const tableMatch = text.match(/^INSERT INTO\s+(\w+)/i);
  const targetsNoIdTable = tableMatch && NO_ID_TABLES.includes(tableMatch[1].toLowerCase());
  if (isInsert && !targetsNoIdTable && !/RETURNING/i.test(text)) {
    text = text.replace(/;?\s*$/, "") + " RETURNING id";
  }
  // positional ? -> $1, $2, ...
  let i = 0;
  text = text.replace(/\?/g, () => `$${++i}`);
  return text;
}

function prepare(sql) {
  const pgSql = toPgSql(sql);
  return {
    async run(...args) {
      const r = await pool.query(pgSql, args);
      return { changes: r.rowCount, lastInsertRowid: r.rows[0]?.id };
    },
    async get(...args) {
      const r = await pool.query(pgSql, args);
      return r.rows[0];
    },
    async all(...args) {
      const r = await pool.query(pgSql, args);
      return r.rows;
    },
  };
}

async function exec(sql) {
  await pool.query(sql);
}

module.exports = { prepare, exec, pool };
