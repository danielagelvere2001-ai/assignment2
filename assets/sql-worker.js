// Runs student queries with sql.js (SQLite compiled to WebAssembly) off the main thread,
// so a runaway query can be stopped by terminating the worker.
const SQLJS = "https://cdn.jsdelivr.net/npm/sql.js@1.10.3/dist/";
importScripts(SQLJS + "sql-wasm.js");

const ready = initSqlJs({ locateFile: (f) => SQLJS + f });
const original = {};  // dataset file -> pristine bytes, so "reset" undoes any UPDATE/DELETE
let db = null, current = null;

async function load(file) {
  if (!original[file]) {
    const res = await fetch("../data/" + file + ".gz");
    if (!res.ok) throw new Error("Could not download " + file + " (" + res.status + ")");
    let bytes = new Uint8Array(await res.arrayBuffer());
    // Some servers already decode .gz in transit; only gunzip if the gzip magic bytes are there.
    if (bytes[0] === 0x1f && bytes[1] === 0x8b) {
      const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream("gzip"));
      bytes = new Uint8Array(await new Response(stream).arrayBuffer());
    }
    original[file] = bytes;
  }
  const SQL = await ready;
  if (db) db.close();
  db = new SQL.Database(original[file]);
  current = file;
}

self.onmessage = async (e) => {
  const { id, action, file, sql } = e.data;
  try {
    if (action === "open") {
      if (current !== file) await load(file);
      self.postMessage({ id, ok: true });
    } else if (action === "reset") {
      await load(file);
      self.postMessage({ id, ok: true });
    } else if (action === "exec") {
      if (current !== file) await load(file);
      const t0 = performance.now();
      const results = db.exec(sql);
      const last = results.length ? results[results.length - 1] : null;
      self.postMessage({ id, ok: true, ms: Math.round(performance.now() - t0),
        result: last ? { columns: last.columns, rows: last.values } : null, statements: results.length });
    }
  } catch (err) {
    self.postMessage({ id, ok: false, error: String(err.message || err) });
  }
};
