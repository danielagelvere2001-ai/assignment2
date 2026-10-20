// Answer checking. Pure functions, no DOM.
//
// The rule this file follows: a correct answer must never be marked wrong
// because of how it was typed. Students at SDU often use Danish Excel
// (1.203,5 instead of 1,203.5; 19-05-2014 instead of 5/19/2014), copy an
// Excel date serial, or write a percentage as 7.84% instead of 0.0784. Every
// plausible reading of the input is tried, and the reading that was accepted
// is reported back, so the student can see how the input was understood.

const DA_DAYS = { mandag: "monday", tirsdag: "tuesday", onsdag: "wednesday", torsdag: "thursday",
  fredag: "friday", "lørdag": "saturday", "søndag": "sunday" };

export function numberReadings(input) {
  let s = String(input).trim().replace(/[    ]/g, "");
  if (!s) return [];
  s = s.replace(/^(dkk|kr\.?|usd|eur|\$|€)/i, "").replace(/(dkk|kr\.?|usd|eur|\$|€)$/i, "");
  let pct = false;
  if (s.endsWith("%")) { pct = true; s = s.slice(0, -1); }
  let neg = false;
  if (/^\(.*\)$/.test(s)) { neg = true; s = s.slice(1, -1); }
  if (/^[-−–]/.test(s)) { neg = true; s = s.slice(1); }
  if (!/^[\d.,]+$/.test(s) || !/\d/.test(s)) return [];

  const out = [];
  const add = (txt, reading) => {
    const v = Number(txt);
    if (Number.isFinite(v)) out.push({ value: (neg ? -v : v) / (pct ? 100 : 1), pct, reading });
  };
  const dots = (s.match(/\./g) || []).length;
  const commas = (s.match(/,/g) || []).length;
  if (!dots && !commas) add(s, "plain number");
  else if (dots && commas) {
    if (s.lastIndexOf(",") > s.lastIndexOf(".")) add(s.replace(/\./g, "").replace(",", "."), "Danish format (1.234,5)");
    else add(s.replace(/,/g, ""), "English format (1,234.5)");
  } else if (commas) {
    if (commas === 1) add(s.replace(",", "."), "decimal comma (Danish)");
    if (/^\d{1,3}(,\d{3})+$/.test(s)) add(s.replace(/,/g, ""), "comma as thousands separator (English)");
  } else {
    if (dots === 1) add(s, "decimal point (English)");
    if (/^\d{1,3}(\.\d{3})+$/.test(s)) add(s.replace(/\./g, ""), "point as thousands separator (Danish)");
  }
  return out;
}

function close(a, b, tol) {
  const t = tol ?? Math.max(1e-9, Math.abs(b) * 1e-9);
  return Math.abs(a - b) <= t + 1e-12;
}

function normText(s) {
  let t = String(s).trim().toLowerCase().replace(/^["'“”]+|["'“”.]+$/g, "").replace(/\s+/g, " ");
  return DA_DAYS[t] || t;
}

function textSet(s) {
  return String(s).split(/\s*(?:,|;|\/|&|\band\b|\bog\b)\s*/i).map(normText).filter(Boolean).sort().join("|");
}

// --- dates ---------------------------------------------------------------

function pad(n) { return String(n).padStart(2, "0"); }
function stamp(y, mo, d, h = 0, mi = 0) {
  if (mo < 1 || mo > 12 || d < 1 || d > 31 || h > 23 || mi > 59) return null;
  return `${y}-${pad(mo)}-${pad(d)} ${pad(h)}:${pad(mi)}`;
}

export function dateReadings(input) {
  const s = String(input).trim();
  const out = [];
  let m;
  if ((m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2})[:.](\d{2})(?::\d{2})?)?$/))) {
    out.push({ value: stamp(+m[1], +m[2], +m[3], +(m[4] || 0), +(m[5] || 0)), reading: "ISO date (YYYY-MM-DD)" });
  } else if ((m = s.match(/^(\d{1,2})([/.-])(\d{1,2})\2(\d{4})(?:,?\s+(\d{1,2})[:.](\d{2})(?::\d{2})?\s*(am|pm)?)?$/i))) {
    let h = +(m[5] || 0);
    if (m[7]) h = (h % 12) + (m[7].toLowerCase() === "pm" ? 12 : 0);
    const mi = +(m[6] || 0);
    if (m[2] === "/") out.push({ value: stamp(+m[4], +m[1], +m[3], h, mi), reading: "US date (M/D/YYYY)" });
    out.push({ value: stamp(+m[4], +m[3], +m[1], h, mi), reading: "European date (D-M-YYYY)" });
  } else {
    for (const n of numberReadings(s)) {
      if (n.value > 20000 && n.value < 80000) {
        const ms = Math.round((n.value - 25569) * 86400000 / 60000) * 60000;
        const d = new Date(ms);
        out.push({ value: stamp(d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate(), d.getUTCHours(), d.getUTCMinutes()),
          reading: "Excel date serial number" });
      }
    }
  }
  return out.filter((r) => r.value);
}

// --- one Excel field -----------------------------------------------------

export function checkField(field, input) {
  const raw = String(input ?? "").trim();
  if (!raw) return { status: "empty" };
  const { kind, value } = field;

  if (kind === "text") {
    const exp = String(value);
    const ok = exp.includes(",") ? textSet(raw) === textSet(exp) : normText(raw) === normText(exp);
    if (ok) return { status: "correct" };
    if (exp.includes(",") && textSet(exp).split("|").some((t) => textSet(raw).split("|").includes(t)))
      return { status: "wrong", hint: "Partly right: your list overlaps with the answer but is not complete or has extra items." };
    return { status: "wrong", hint: "Not matching. Check spelling against the data, and that you used the right filter before ranking." };
  }

  if (kind === "datetime") {
    const exp = String(value).slice(0, 16);
    const reads = dateReadings(raw);
    if (!reads.length) return { status: "unreadable", hint: "Could not read this as a date. Try 2014-05-19 14:30 or 19-05-2014 14:30." };
    const hit = reads.find((r) => r.value === exp);
    if (hit) return { status: "correct", reading: hit.reading };
    if (reads.some((r) => r.value.slice(0, 10) === exp.slice(0, 10)))
      return { status: "wrong", hint: "Right day, wrong time. Include hours and minutes." };
    return { status: "wrong", hint: "Not matching. If month and day look swapped, the file was imported with the wrong locale." };
  }

  const reads = numberReadings(raw);
  if (!reads.length) return { status: "unreadable", hint: "Could not read this as a number." };
  const exp = Number(value);
  const tol = field.tol ?? (kind === "int" ? 0 : null);
  const candidates = [];
  for (const r of reads) {
    candidates.push(r);
    // 7.84 typed for a percentage of 0.0784
    if (kind === "pct" && !r.pct && Math.abs(r.value) > 1) candidates.push({ ...r, value: r.value / 100, reading: r.reading + ", read as a percentage" });
  }
  const hit = candidates.find((r) => close(r.value, exp, tol));
  if (hit) return { status: "correct", reading: reads.length > 1 || hit.reading.includes("percentage") ? hit.reading : null };
  return { status: "wrong", hint: numericHint(candidates[0].value, exp, kind, tol) };
}

function numericHint(got, exp, kind, tol) {
  if (exp !== 0) {
    const ratio = got / exp;
    for (const [f, msg] of [[100, "about 100× the answer: a percentage/fraction mix-up?"], [0.01, "about 1/100 of the answer: a percentage/fraction mix-up?"],
      [1000, "about 1,000× the answer: check units (thousands)."], [0.001, "about 1/1,000 of the answer: check units (thousands)."],
      [1e6, "about 1,000,000× the answer: check units (millions)."], [1e-6, "about 1/1,000,000 of the answer: check units (millions)."],
      [-1, "the right size but the wrong sign: check debits/credits or how returns are recorded."]]) {
      if (Math.abs(ratio - f) < Math.abs(f) * 0.01) return "Your value is " + msg;
    }
    if (tol != null && Math.abs(got - exp) <= Math.max(tol * 20, Math.abs(exp) * 0.002))
      return "Very close. Check the rounding instruction in the question.";
  }
  const dir = got > exp ? "higher" : "lower";
  if (kind === "int")
    return `Your count is ${dir} than the reference. ${got > exp
      ? "Your filter may let in too many rows (wrong condition, duplicates, or a header/total row counted)."
      : "Your filter may be too strict, or rows were lost in a lookup or import."}`;
  return `Your value is ${dir} than the reference. Re-check which rows are included and whether blanks or text values were summed.`;
}

// --- SQL result sets ------------------------------------------------------

function cellKey(v) {
  if (v === null || v === undefined) return "∅";
  if (typeof v === "number") return Number.isInteger(v) ? String(v) : v.toFixed(2);
  if (typeof v === "bigint") return String(v);
  const s = String(v).trim();
  if (/^-?\d+(\.\d+)?$/.test(s)) { const n = Number(s); return Number.isInteger(n) ? String(n) : n.toFixed(2); }
  return s;
}

function multiset(keys) {
  const m = new Map();
  for (const k of keys) m.set(k, (m.get(k) || 0) + 1);
  return m;
}

function sameMultiset(a, b) {
  if (a.length !== b.length) return false;
  const m = multiset(a);
  for (const k of b) { const n = m.get(k); if (!n) return false; m.set(k, n - 1); }
  return true;
}

export function compareResults(expected, got) {
  const E = expected.rows, ec = expected.columns.length;
  if (!got) return { status: "wrong", messages: ["Your query returned no result table. Did the last statement end with a SELECT?"] };
  const G = got.rows, gc = got.columns.length;
  const msgs = [];

  // Map each expected column to a student column with the same values (any order, any alias).
  const colKeys = (rows, j) => rows.map((r) => cellKey(r[j]));
  const used = new Set();
  const map = [];
  for (let j = 0; j < ec; j++) {
    const ek = colKeys(E, j);
    let found = -1;
    for (let k = 0; k < gc; k++) if (!used.has(k) && sameMultiset(ek, colKeys(G, k))) { found = k; break; }
    if (found < 0 && j < gc && !used.has(j)) found = j;
    if (found < 0) for (let k = 0; k < gc; k++) if (!used.has(k)) { found = k; break; }
    map.push(found);
    if (found >= 0) used.add(found);
  }
  if (gc < ec) msgs.push(`Your result has ${gc} column${gc === 1 ? "" : "s"}; the question asks for ${ec}.`);
  if (gc > ec) msgs.push(`Your result has ${gc} columns; the question asks for ${ec}. Remove the extra column${gc - ec > 1 ? "s" : ""}.`);

  const proj = G.map((r) => map.map((k) => (k >= 0 ? r[k] : undefined)));
  const ekeys = E.map((r) => r.map(cellKey).join("¦"));
  const gkeys = proj.map((r) => r.map(cellKey).join("¦"));
  const em = multiset(ekeys);
  let missing = 0, extra = 0, dupes = 0;
  const gm = multiset(gkeys);
  for (const [k, n] of em) missing += Math.max(0, n - (gm.get(k) || 0));
  for (const [k, n] of gm) {
    const over = n - (em.get(k) || 0);
    if (over > 0) { extra += over; if (em.get(k)) dupes += over; }
  }

  if (!missing && !extra && gc === ec) {
    if (expected.order && expected.order.length) {
      const seq = (rows) => rows.map((r) => expected.order.map((j) => cellKey(r[j])).join("¦"));
      const es = seq(E), gs = seq(proj);
      const first = es.findIndex((k, i) => k !== gs[i]);
      if (first >= 0) return { status: "wrong", messages: [`All the right rows, but in the wrong order (first difference at row ${first + 1}). Check your ORDER BY, including the direction (ASC/DESC).`] };
    }
    if (map.some((k, j) => k !== j)) msgs.push("Your columns are in a different order than the question lists them. Accepted, but the asked order is easier to read.");
    return { status: "correct", messages: msgs };
  }

  // Single value: give a numeric hint
  if (E.length === 1 && ec === 1 && G.length === 1 && typeof E[0][0] === "number") {
    const n = Number(proj[0][0]);
    if (Number.isFinite(n)) {
      if (Math.abs(n - E[0][0]) < 0.01 * Math.max(1, Math.abs(E[0][0])) && cellKey(n) !== cellKey(E[0][0]))
        return { status: "wrong", messages: ["Very close. Check the rounding asked for (ROUND(x, 2))."] };
      return { status: "wrong", messages: [numericHint(n, E[0][0], Number.isInteger(E[0][0]) ? "int" : "num", null)] };
    }
  }

  msgs.unshift(`Expected ${E.length} row${E.length === 1 ? "" : "s"}; you returned ${G.length}.`);
  if (dupes) msgs.push(`${dupes} of your rows are duplicates of correct rows. Do you need DISTINCT, or is a JOIN multiplying rows?`);
  if (extra - dupes > 0) msgs.push(`${extra - dupes} row${extra - dupes === 1 ? " is" : "s are"} not in the reference result. A filter may be missing or too loose.`);
  if (missing) msgs.push(`${missing} reference row${missing === 1 ? " is" : "s are"} missing from your result. A filter may be too strict, a LIMIT too small, or an INNER JOIN dropping rows a LEFT JOIN would keep.`);
  if (!missing && !extra && gc !== ec) msgs.push("Apart from the columns, your rows match.");
  return { status: "wrong", messages: msgs };
}
