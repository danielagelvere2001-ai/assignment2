import { checkField, compareResults } from "./check.js";
import { toDanish } from "./excel-da.js";

const REPO = "danielagelvere2001-ai/assignment2";
const app = document.getElementById("app");
const S = { questions: [], byId: {}, datasets: {}, exams: [], regression: null, reviews: {}, built: "" };

// ------------------------------------------------------------------ storage (per-viewer conveniences only)
const store = {
  get(k, d) { try { const v = localStorage.getItem("ep:" + k); return v === null ? d : JSON.parse(v); } catch { return d; } },
  set(k, v) { try { localStorage.setItem("ep:" + k, JSON.stringify(v)); } catch { /* private mode */ } },
};
const solved = new Set(store.get("solved", []));
function markSolved(id) { solved.add(id); store.set("solved", [...solved]); }

// ------------------------------------------------------------------ helpers
const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
// Group thousands only for larger numbers so years and IDs below 10,000 read naturally.
const fmt = (v) => (typeof v !== "number" ? String(v)
  : v.toLocaleString("en-US", { maximumFractionDigits: 6, useGrouping: Math.abs(v) >= 10000 }));
const h = (strings, ...vals) => strings.reduce((a, s, i) => a + s + (i < vals.length ? vals[i] : ""), "");
function el(html) { const t = document.createElement("template"); t.innerHTML = html.trim(); return t.content.firstElementChild; }
function fileLink(f) { return `<a class="btn small" href="data/${esc(f)}" download>⭳ ${esc(f)}</a>`; }

// ------------------------------------------------------------------ trust: who stands behind this answer?
// Order of precedence: a human review that matches the current answer, then a matching
// instructor answer from the course materials. Anything else is shown as not yet reviewed.
export function trustOf(item) {
  const r = S.reviews[item.id];
  if (r && r.fingerprint === item.fingerprint) {
    if (r.decision === "approved") return { level: "reviewed", label: `Reviewed by ${r.reviewer}`, review: r };
    if (r.decision === "rejected") return { level: "rejected", label: "Withdrawn by reviewer", review: r };
  }
  if (r && r.fingerprint !== item.fingerprint) return { level: "stale", label: "Changed since review", review: r };
  if (item.instructor && item.instructor.agrees) return { level: "instructor", label: "Matches instructor key" };
  if (item.instructor) return { level: "conflict", label: "Instructor key disagrees" };
  return { level: "unreviewed", label: "Not yet reviewed" };
}

function badge(item) { const t = trustOf(item); return `<span class="badge ${t.level}">${esc(t.label)}</span>`; }

function trustPanel(item) {
  const t = trustOf(item), r = t.review;
  const hist = `https://github.com/${REPO}/commits/main/content/reviews.json`;
  const inst = item.instructor;
  const body = {
    reviewed: `Checked and approved by <strong>${esc(r?.reviewer)}</strong> on ${esc(r?.date)}${r?.note ? ` · “${esc(r.note)}”` : ""}. <a href="${hist}">Sign-off history</a>`,
    rejected: `A reviewer withdrew this answer on ${esc(r?.date)}${r?.note ? `: “${esc(r.note)}”` : ""}. Checking is switched off until it is corrected.`,
    stale: `This question or its data changed after ${esc(r?.reviewer)} reviewed it on ${esc(r?.date)}. The old sign-off no longer applies, so the answer needs to be reviewed again.`,
    instructor: `The computed answer matches the published course answer: ${esc(inst?.source)}.`,
    conflict: `The course answer key says <strong>${esc(fmtVal(inst?.value))}</strong> (${esc(inst?.source)}), but the data gives a different result. ${inst?.note ? esc(inst.note) : ""} Until a reviewer decides, your answer is checked against the data and both values are shown here.`,
    unreviewed: `Computed from the data by the method shown after your attempt, but no person has checked it yet. Treat a mismatch as a reason to double-check, not as proof you are wrong. <a href="#/answer-key">How answers are verified</a>`,
  }[t.level];
  return `<div class="trust ${t.level}"><span class="badge ${t.level}">${esc(t.label)}</span> ${body}</div>`;
}
const fmtVal = (v) => (Array.isArray(v) ? v.map(fmt).join(" · ") : fmt(v));

function disputeLink(item, attempt) {
  const title = `Dispute ${item.id}: ${item.prompt.slice(0, 60)}`;
  const body = `**Question:** ${item.id}\n**Answer version:** ${item.fingerprint}\n\n**My answer / result:**\n${attempt || "(paste here)"}\n\n**Why I think the reference answer is wrong:**\n`;
  return `https://github.com/${REPO}/issues/new?labels=dispute&title=${encodeURIComponent(title)}&body=${encodeURIComponent(body)}`;
}

// ------------------------------------------------------------------ SQL runner (Web Worker)
const sqlRunner = (() => {
  let worker = null, seq = 0;
  const pending = new Map();
  function start() {
    worker = new Worker("assets/sql-worker.js");
    worker.onmessage = (e) => { const p = pending.get(e.data.id); if (p) { pending.delete(e.data.id); p.resolve(e.data); } };
    worker.onerror = (e) => { for (const p of pending.values()) p.resolve({ ok: false, error: e.message || "Worker error" }); pending.clear(); };
  }
  function send(msg, timeout = 30000) {
    if (!worker) start();
    const id = ++seq;
    return new Promise((resolve) => {
      const timer = setTimeout(() => {
        pending.delete(id); worker.terminate(); worker = null;
        resolve({ ok: false, error: `Stopped after ${timeout / 1000} seconds. A JOIN without an ON condition can multiply rows; check your joins.` });
      }, timeout);
      pending.set(id, { resolve: (d) => { clearTimeout(timer); resolve(d); } });
      worker.postMessage({ id, ...msg });
    });
  }
  return {
    open: (file) => send({ action: "open", file }, 120000),
    reset: (file) => send({ action: "reset", file }, 120000),
    exec: (file, sql) => send({ action: "exec", file, sql }),
  };
})();

function resultTable(res, limit = 200) {
  if (!res) return `<p class="muted small">No rows returned.</p>`;
  const rows = res.rows.slice(0, limit);
  const isNum = res.columns.map((_, j) => rows.length && rows.every((r) => r[j] === null || typeof r[j] === "number"));
  return h`<div class="table-wrap"><table><thead><tr>${res.columns.map((c, j) => `<th class="${isNum[j] ? "num" : ""}">${esc(c)}</th>`).join("")}</tr></thead>
    <tbody>${rows.map((r) => `<tr>${r.map((v, j) => `<td class="${isNum[j] ? "num" : ""}">${v === null ? '<span class="muted">NULL</span>' : esc(fmt(v))}</td>`).join("")}</tr>`).join("")}</tbody></table></div>
    <p class="muted small">${res.rows.length} row${res.rows.length === 1 ? "" : "s"}${res.rows.length > limit ? `, first ${limit} shown` : ""}.</p>`;
}

// ------------------------------------------------------------------ question component
const attempts = {};  // id -> number of checks this visit

function renderQuestion(q, mount, { onSolved } = {}) {
  const ds = S.datasets[q.ds];
  const t = trustOf(q);
  const node = el(h`<article class="stack">
    <div class="spread"><div class="row"><span class="tag">${esc(q.subject)}</span><span class="tag">${esc(ds.case)}</span>
      <span class="tag">${esc(ds.name)}</span><span class="tag">≈ ${q.minutes} min</span></div>${badge(q)}</div>
    <h2 style="margin-top:.2em">${esc(q.id)} · ${esc(q.prompt)}</h2>
    <div class="downloads">${ds.files.map(fileLink).join("")}</div>
    ${ds.note ? `<p class="muted small">${esc(ds.note)}</p>` : ""}
    <div class="work"></div>
    <div class="outcome" aria-live="polite"></div>
    <div class="reveal"></div>
    <details><summary>Who stands behind this answer?</summary><div style="margin-top:.6rem">${trustPanel(q)}</div></details>
    <p class="small muted">Think the reference answer is wrong? <a class="dispute" href="${disputeLink(q)}" target="_blank" rel="noopener">Dispute it</a>. A person reads every dispute.</p>
  </article>`);
  mount.replaceChildren(node);
  const work = node.querySelector(".work"), outcome = node.querySelector(".outcome"), reveal = node.querySelector(".reveal");
  const dispute = node.querySelector(".dispute");
  const disabled = t.level === "rejected";

  const showRevealButton = () => {
    if (reveal.childElementCount) return;
    const b = el(`<button class="btn">Show reference answer and method</button>`);
    b.onclick = () => {
      if ((attempts[q.id] || 0) < 2 && !confirm("You have checked once. Try one more time before looking?\n\nOK = show the answer now")) return;
      reveal.replaceChildren(el(`<div class="card stack">${referenceHtml(q)}</div>`));
    };
    reveal.append(b);
  };

  const succeed = () => { markSolved(q.id); onSolved?.(q.id); showRevealButton(); };

  if (q.subject === "Excel") {
    const form = el(h`<form class="card" novalidate>
      <p class="small muted" style="margin-top:0">Solve it in Excel, then reconcile your result: enter the figures below. Danish number and date formats are fine (1.203,5 · 19-05-2014).</p>
      ${q.fields.map((f, i) => h`<label class="field" for="${q.id}-f${i}">${esc(f.label)}</label>
        <div class="field-row"><input type="text" id="${q.id}-f${i}" name="f${i}" autocomplete="off" inputmode="${f.kind === "text" ? "text" : "decimal"}">
        <span class="field-status" id="${q.id}-s${i}"></span></div>`).join("")}
      <div class="row" style="margin-top:1rem"><button class="btn primary" type="submit" ${disabled ? "disabled" : ""}>Check my answer</button></div>
    </form>`);
    form.onsubmit = (e) => {
      e.preventDefault();
      attempts[q.id] = (attempts[q.id] || 0) + 1;
      const values = q.fields.map((_, i) => form.elements["f" + i].value);
      const res = q.fields.map((f, i) => checkField(f, values[i]));
      res.forEach((r, i) => {
        const s = form.querySelector(`#${q.id}-s${i}`);
        s.className = "field-status " + r.status;
        s.textContent = r.status === "correct" ? "✓ correct" + (r.reading ? ` (read as ${r.reading})` : "") : r.status === "empty" ? "" : "✗ " + (r.hint || "not yet");
      });
      const ok = res.every((r) => r.status === "correct");
      outcome.innerHTML = ok
        ? `<div class="result correct"><strong>Correct.</strong> Every figure reconciles with the reference answer.</div>`
        : res.some((r) => r.status === "empty") ? `<div class="result info">Fill in every figure to check.</div>`
          : `<div class="result wrong"><strong>Not yet.</strong> See the notes next to each figure.</div>`;
      dispute.href = disputeLink(q, q.fields.map((f, i) => `${f.label}: ${values[i]}`).join("\n"));
      if (ok) succeed(); else showRevealButton();
    };
    work.append(form);
  } else {
    const key = "sql:" + q.id;
    const box = el(h`<div class="card stack">
      <label class="field" for="${q.id}-sql" style="margin-top:0">Your query <span class="muted small">(runs in your browser on ${esc(ds.files[0])}; Ctrl+Enter to run)</span></label>
      <textarea class="sql" id="${q.id}-sql" spellcheck="false" placeholder="SELECT …"></textarea>
      <div class="row">
        <button class="btn" data-a="run">Run</button>
        <button class="btn primary" data-a="check" ${disabled ? "disabled" : ""}>Check my result</button>
        <button class="btn small" data-a="schema">Show tables</button>
        <button class="btn small" data-a="reset" title="Reload the original database, undoing any changes">Reset database</button>
        <span class="muted small status"></span>
      </div>
      <div class="output"></div>
    </div>`);
    const ta = box.querySelector("textarea"), status = box.querySelector(".status"), output = box.querySelector(".output");
    ta.value = store.get(key, "");
    ta.oninput = () => store.set(key, ta.value);
    const file = ds.db;
    const run = async (sqlText) => {
      status.textContent = file === "pcards.db" ? "Loading the database (10 MB the first time)…" : "Running…";
      const r = await sqlRunner.exec(file, sqlText);
      status.textContent = r.ok ? `${r.ms} ms` : "";
      return r;
    };
    box.querySelector('[data-a="run"]').onclick = async () => {
      const r = await run(ta.value);
      output.innerHTML = r.ok ? resultTable(r.result) : `<div class="result wrong">${esc(r.error)}</div>`;
    };
    box.querySelector('[data-a="check"]').onclick = async () => {
      if (!ta.value.trim()) { outcome.innerHTML = `<div class="result info">Write a query first.</div>`; return; }
      attempts[q.id] = (attempts[q.id] || 0) + 1;
      const r = await run(ta.value);
      if (!r.ok) { outcome.innerHTML = `<div class="result wrong">${esc(r.error)}</div>`; showRevealButton(); return; }
      output.innerHTML = resultTable(r.result);
      const cmp = compareResults({ ...q.result, order: q.order }, r.result);
      outcome.innerHTML = cmp.status === "correct"
        ? `<div class="result correct"><strong>Correct.</strong> Your result matches the reference result (${q.result.rows.length} row${q.result.rows.length === 1 ? "" : "s"}).${cmp.messages.length ? `<ul>${cmp.messages.map((m) => `<li>${esc(m)}</li>`).join("")}</ul>` : ""}</div>`
        : `<div class="result wrong"><strong>Not yet.</strong><ul>${cmp.messages.map((m) => `<li>${esc(m)}</li>`).join("")}</ul></div>`;
      dispute.href = disputeLink(q, "```sql\n" + ta.value.trim() + "\n```");
      if (cmp.status === "correct") succeed(); else showRevealButton();
    };
    box.querySelector('[data-a="schema"]').onclick = async () => {
      const r = await run("SELECT name AS table_name, sql AS definition FROM sqlite_master WHERE type = 'table' ORDER BY name");
      output.innerHTML = r.ok ? resultTable(r.result) : `<div class="result wrong">${esc(r.error)}</div>`;
    };
    box.querySelector('[data-a="reset"]').onclick = async () => {
      status.textContent = "Resetting…"; await sqlRunner.reset(file); status.textContent = "Original database reloaded.";
    };
    ta.onkeydown = (e) => { if ((e.ctrlKey || e.metaKey) && e.key === "Enter") { e.preventDefault(); box.querySelector('[data-a="run"]').click(); } };
    work.append(box);
  }
  if (solved.has(q.id)) showRevealButton();
}

function referenceHtml(q) {
  const lang = store.get("excelLang", "en");
  let out = `<h3 style="margin-top:0">Reference answer</h3>`;
  if (q.subject === "Excel") {
    out += `<table><tbody>${q.fields.map((f) => `<tr><th>${esc(f.label)}</th><td>${esc(f.kind === "pct" ? (f.value * 100).toFixed(2) + "%" : fmt(f.value))}</td></tr>`).join("")}</tbody></table>`;
    if (q.steps) {
      out += `<div class="spread"><h3>One way to do it in Excel</h3><label class="small"><input type="checkbox" class="da" ${lang === "da" ? "checked" : ""}> Danish Excel</label></div>
        <pre class="steps">${esc(lang === "da" ? toDanish(q.steps) : q.steps)}</pre>`;
    }
    const methods = q.fields.filter((f) => f.method);
    if (methods.length) out += `<details><summary>The check that produced this answer</summary>${methods.map((f) => `<p class="small"><strong>${esc(f.label)}</strong></p><pre>${esc(f.method)}</pre>`).join("")}</details>`;
  } else {
    out += `<pre>${esc(q.sql)}</pre>${resultTable(q.result, 50)}`;
  }
  return out;
}

document.addEventListener("change", (e) => {
  if (!e.target.matches("input.da")) return;
  store.set("excelLang", e.target.checked ? "da" : "en");
  const card = e.target.closest(".card"), q = S.byId[card?.closest("[data-qid]")?.dataset.qid];
  const pre = card.querySelector("pre.steps");
  if (pre && q) pre.textContent = e.target.checked ? toDanish(q.steps) : q.steps;
});

// ------------------------------------------------------------------ pages
const pages = {};

pages.home = () => h`
  <section class="card hero">
    <h1>Practise the exam you will actually sit</h1>
    <p>The oral exam gives you one <strong>data manipulation</strong> exam (Excel and SQL on the course datasets) and one <strong>data interpretation</strong> exam (regression output). Every question here is built only from the course materials, and you can check every answer.</p>
    <p>An answer key you cannot trust is worse than none. Each answer on this site shows <strong>who stands behind it</strong>: the instructor's published key, a named human reviewer, or nobody yet. <a href="#/answer-key">How that works</a>.</p>
    <div class="row"><a class="btn primary" href="#/exams/random">Start a random manipulation exam</a><a class="btn" href="#/interpretation">Interpretation exams</a><a class="btn" href="#/questions">Browse all ${S.questions.length} questions</a></div>
  </section>
  <div class="grid" style="margin-top:1rem">
    ${["Excel", "SQL"].map((sub) => {
      const ds = Object.entries(S.datasets).filter(([, d]) => d.subject === sub);
      return `<div class="card"><h3 style="margin-top:0">${sub}</h3><ul class="small">${ds.map(([k, d]) => `<li><a href="#/questions?ds=${k}">${esc(d.name)}</a> <span class="muted">· ${esc(d.case)}</span></li>`).join("")}</ul></div>`;
    }).join("")}
    <div class="card"><h3 style="margin-top:0">Answer key status</h3>${trustSummaryHtml()}</div>
  </div>`;

function trustCounts() {
  const all = [...S.questions, ...S.regression.exams.flatMap((e) => e.questions.map((q) => ({ ...q, ds: "regression" })))];
  const c = { instructor: 0, reviewed: 0, conflict: 0, stale: 0, rejected: 0, unreviewed: 0 };
  for (const q of all) c[trustOf(q).level]++;
  return { c, total: all.length };
}
function trustSummaryHtml() {
  const { c, total } = trustCounts();
  return `<ul class="small" style="padding-left:1.1rem">
    <li><span class="badge reviewed">Reviewed</span> ${c.reviewed}</li>
    <li><span class="badge instructor">Matches instructor key</span> ${c.instructor}</li>
    <li><span class="badge conflict">Instructor key disagrees</span> ${c.conflict}</li>
    ${c.stale ? `<li><span class="badge stale">Changed since review</span> ${c.stale}</li>` : ""}
    <li><span class="badge unreviewed">Not yet reviewed</span> ${c.unreviewed}</li></ul>
    <p class="small muted">${total} answers in total. <a href="#/answer-key">Details</a></p>`;
}

pages.questions = (params) => {
  const sub = params.get("subject") || "", ds = params.get("ds") || "", tr = params.get("trust") || "", text = params.get("q") || "";
  const list = S.questions.filter((q) => (!sub || q.subject === sub) && (!ds || q.ds === ds) && (!tr || trustOf(q).level === tr)
    && (!text || (q.id + " " + q.prompt + " " + q.topics.join(" ")).toLowerCase().includes(text.toLowerCase())));
  const opt = (v, cur, label) => `<option value="${v}" ${v === cur ? "selected" : ""}>${esc(label)}</option>`;
  return h`<h1>Question bank</h1>
    <p class="muted">Every manipulation exam is made of these questions. ${solved.size ? `You have solved ${[...solved].filter((id) => S.byId[id]).length} of ${S.questions.length}.` : ""}</p>
    <form class="filters" id="qfilter">
      <select name="subject" aria-label="Subject">${opt("", sub, "All subjects")}${opt("Excel", sub, "Excel")}${opt("SQL", sub, "SQL")}</select>
      <select name="ds" aria-label="Dataset">${opt("", ds, "All datasets")}${Object.entries(S.datasets).map(([k, d]) => opt(k, ds, d.name)).join("")}</select>
      <select name="trust" aria-label="Answer status">${opt("", tr, "Any answer status")}${opt("reviewed", tr, "Reviewed")}${opt("instructor", tr, "Matches instructor key")}${opt("conflict", tr, "Instructor key disagrees")}${opt("unreviewed", tr, "Not yet reviewed")}</select>
      <input type="search" name="q" value="${esc(text)}" placeholder="Search (e.g. JOIN, overtime)" aria-label="Search">
    </form>
    <p class="small muted">${list.length} question${list.length === 1 ? "" : "s"}</p>
    <ul class="qlist">${list.map((q) => h`<li><a href="#/q/${q.id}"><span class="qid">${q.id}</span>${esc(q.prompt)}
      <div class="meta"><span class="tag">${q.subject}</span><span class="tag">${esc(S.datasets[q.ds].name)}</span>${q.topics.map((t) => `<span class="tag">${esc(t)}</span>`).join("")}${badge(q)}${solved.has(q.id) ? '<span class="solved small">✓ solved</span>' : ""}</div></a></li>`).join("")}</ul>`;
};
pages.questions.after = () => {
  const f = document.getElementById("qfilter");
  const go = () => { const p = new URLSearchParams(); for (const [k, v] of new FormData(f)) if (v) p.set(k, v); location.hash = "#/questions" + (p.size ? "?" + p : ""); };
  f.onchange = go;
  f.onsubmit = (e) => { e.preventDefault(); go(); };
  let tmo; f.elements.q.oninput = () => { clearTimeout(tmo); tmo = setTimeout(go, 400); };
};

pages.q = (params, id) => {
  const q = S.byId[id];
  if (!q) return `<h1>Question not found</h1><p><a href="#/questions">Back to the question bank</a></p>`;
  const i = S.questions.indexOf(q);
  const prev = S.questions[i - 1], next = S.questions[i + 1];
  return h`<p><a href="#/questions">← Question bank</a></p><div id="qmount" data-qid="${q.id}"></div>
    <div class="spread" style="margin-top:1.4rem">${prev ? `<a class="btn" href="#/q/${prev.id}">← ${prev.id}</a>` : "<span></span>"}${next ? `<a class="btn" href="#/q/${next.id}">${next.id} →</a>` : ""}</div>`;
};
pages.q.after = (params, id) => { const q = S.byId[id]; if (q) renderQuestion(q, document.getElementById("qmount")); };

pages.exams = () => h`<h1>Data manipulation exams</h1>
  <p>Each exam has two Excel questions on one dataset and three SQL questions on one database, with an eight-minute timer as in the oral exam. Download both files before you start the clock.</p>
  <p><a class="btn primary" href="#/exams/random">Choose one for me</a></p>
  <ul class="qlist">${S.exams.map((e) => h`<li><a href="#/exam/${e.id}"><span class="qid">${e.id}</span>${e.datasets.map((d) => esc(S.datasets[d].name)).join(" + ")}
    <div class="meta">${e.questions.map((id) => `<span class="tag">${id}${solved.has(id) ? " ✓" : ""}</span>`).join("")}</div></a></li>`).join("")}</ul>`;

pages.exam = (params, id) => {
  const e = S.exams.find((x) => x.id === id);
  if (!e) return `<h1>Exam not found</h1>`;
  const files = [...new Set(e.datasets.flatMap((d) => S.datasets[d].files))];
  return h`<p><a href="#/exams">← All exams</a></p>
    <h1>Data manipulation exam ${e.id}</h1>
    <div class="card stack">
      <div class="downloads">${files.map(fileLink).join("")}</div>
      ${timerHtml(e.minutes)}
    </div>
    <div class="exam-layout" style="margin-top:1rem">
      <nav class="exam-nav" aria-label="Exam questions">${e.questions.map((qid, i) => `<button class="btn" data-q="${qid}"><span>${i + 1}. ${qid} · ${S.byId[qid].subject}</span><span class="st">${solved.has(qid) ? "✓" : ""}</span></button>`).join("")}</nav>
      <div id="qmount" data-qid=""></div>
    </div>`;
};
pages.exam.after = (params, id) => {
  const e = S.exams.find((x) => x.id === id);
  if (!e) return;
  store.set("lastExam", id);
  wireTimer(e.minutes);
  const mount = document.getElementById("qmount"), nav = document.querySelector(".exam-nav");
  const show = (qid) => {
    mount.dataset.qid = qid;
    nav.querySelectorAll("button").forEach((b) => b.classList.toggle("current", b.dataset.q === qid));
    renderQuestion(S.byId[qid], mount, { onSolved: (sid) => { nav.querySelector(`[data-q="${sid}"] .st`).textContent = "✓"; } });
  };
  nav.onclick = (ev) => { const b = ev.target.closest("button"); if (b) show(b.dataset.q); };
  show(e.questions[0]);
};

pages.random = () => {
  const last = store.get("lastExam", "");
  const pool = S.exams.filter((e) => e.id !== last);
  location.replace("#/exam/" + pool[Math.floor(Math.random() * pool.length)].id);
  return "";
};

function timerHtml(minutes) {
  return h`<div class="timer"><div><div class="muted small">Time remaining</div><div class="clock" id="clock">${minutes}:00</div></div>
    <div><div class="muted small">Elapsed</div><div class="clock" id="elapsed">0:00</div></div>
    <div class="row"><button class="btn primary" id="t-start">Start timer</button><button class="btn" id="t-pause" disabled>Pause</button><button class="btn" id="t-reset">Reset</button></div>
    <span class="sr-only" id="t-live" aria-live="assertive"></span></div>`;
}
let timerHandle = null;
function wireTimer(minutes) {
  clearInterval(timerHandle);
  let elapsed = 0, running = false;
  const total = minutes * 60;
  const mmss = (s) => `${Math.floor(Math.abs(s) / 60)}:${String(Math.abs(s) % 60).padStart(2, "0")}`;
  const clock = document.getElementById("clock"), el2 = document.getElementById("elapsed");
  const start = document.getElementById("t-start"), pause = document.getElementById("t-pause"), live = document.getElementById("t-live");
  const draw = () => {
    const left = total - elapsed;
    clock.textContent = (left < 0 ? "−" : "") + mmss(left);
    clock.classList.toggle("low", left <= 60);
    el2.textContent = mmss(elapsed);
  };
  start.onclick = () => {
    if (running) return;
    running = true; start.disabled = true; pause.disabled = false;
    timerHandle = setInterval(() => {
      if (!document.getElementById("clock")) { clearInterval(timerHandle); return; }
      elapsed++; draw();
      if (elapsed === total) live.textContent = "Time is up.";
      if (elapsed === total - 60) live.textContent = "One minute left.";
    }, 1000);
  };
  pause.onclick = () => { running = false; clearInterval(timerHandle); start.disabled = false; pause.disabled = true; start.textContent = "Resume"; };
  document.getElementById("t-reset").onclick = () => { pause.onclick(); elapsed = 0; start.textContent = "Start timer"; live.textContent = ""; draw(); };
}

// ---- interpretation (regression)
pages.interpretation = () => h`<h1>Data interpretation exams</h1>
  <p>Regression models fitted on the Restaurant case data, merged as the case asks: one row per trading day with total sales, weather, game days and sales from a year earlier. Each exam has four questions and a five-minute timer.</p>
  <p class="small">Re-perform any model yourself: download ${fileLink(S.regression.source)} and run Excel's <em>Data → Data Analysis → Regression</em> with the variables listed in the exam.</p>
  <ul class="qlist">${S.regression.exams.map((e) => `<li><a href="#/reg/${e.id}"><span class="qid">${e.id}</span>${esc(e.title)}<div class="meta">${e.questions.map((q) => badge(q)).join("")}</div></a></li>`).join("")}</ul>`;

pages.reg = (params, id) => {
  const e = S.regression.exams.find((x) => x.id === id);
  if (!e) return `<h1>Exam not found</h1>`;
  const p = (v) => (v < 0.001 ? "&lt;0.001" : v.toFixed(3));
  return h`<p><a href="#/interpretation">← Interpretation exams</a></p>
    <h1>${e.id} · ${esc(e.title)}</h1>
    <div class="card stack">${timerHtml(5)}</div>
    <div class="card stack" style="margin-top:1rem">
      <h3 style="margin-top:0">Variables</h3>
      <ul class="small">${e.variables.map((v) => `<li><strong>${esc(v.name)}</strong>: ${esc(v.desc)}</li>`).join("")}</ul>
      <p><code>${esc(e.equation)}</code></p>
      <table class="reg-table"><thead><tr><th>Parameter</th><th>Term</th><th class="num">Estimate</th><th class="num">p-value</th></tr></thead>
      <tbody>${e.table.map((r) => `<tr><td>${r.param}</td><td>${esc(r.term)}</td><td>${fmt(r.estimate)}</td><td>${p(r.p)}</td></tr>`).join("")}</tbody></table>
      <p class="small muted">n = ${e.n} trading days · R² = ${e.r2}</p>
    </div>
    ${e.questions.map((q, i) => h`<div class="card stack" style="margin-top:1rem" data-rq="${q.id}">
      <div class="spread"><strong>Question ${i + 1}</strong>${badge(q)}</div>
      <p style="margin:0">${esc(q.prompt)}</p>
      <form class="row">${q.kind === "choice"
        ? q.options.map((o) => `<label class="btn small"><input type="radio" name="a" value="${esc(o)}"> ${esc(o)}</label>`).join("")
        : `<input type="text" name="a" autocomplete="off" inputmode="decimal" aria-label="Your answer">`}
        <button class="btn primary" type="submit">Check</button></form>
      <div class="out" aria-live="polite"></div></div>`).join("")}`;
};
pages.reg.after = (params, id) => {
  const e = S.regression.exams.find((x) => x.id === id);
  if (!e) return;
  wireTimer(5);
  document.querySelectorAll("[data-rq]").forEach((card) => {
    const q = e.questions.find((x) => x.id === card.dataset.rq);
    const form = card.querySelector("form"), out = card.querySelector(".out");
    form.onsubmit = (ev) => {
      ev.preventDefault();
      const v = form.elements.a?.value ?? new FormData(form).get("a");
      const r = q.kind === "choice" ? { status: v ? (v === q.answer ? "correct" : "wrong") : "empty" } : checkField({ kind: "num", value: q.answer, tol: q.tol }, v);
      if (r.status === "empty") { out.innerHTML = `<div class="result info">Enter an answer first.</div>`; return; }
      out.innerHTML = r.status === "correct"
        ? `<div class="result correct"><strong>Correct.</strong> ${esc(q.explain)}</div>`
        : `<div class="result wrong"><strong>Not yet.</strong> ${esc(r.hint || "")} <button class="btn small" type="button">Show the answer</button></div>`;
      out.querySelector("button")?.addEventListener("click", () => { out.innerHTML = `<div class="result info">Answer: <strong>${esc(fmt(q.answer))}</strong>. ${esc(q.explain)}</div>`; });
      if (r.status === "correct") markSolved(q.id);
    };
  });
};

// ---- answer key: how answers are verified
pages["answer-key"] = () => {
  const conflicts = S.questions.filter((q) => q.instructor && !q.instructor.agrees);
  return h`<h1>Who stands behind the answers</h1>
  <div class="card stack">
    <p style="margin-top:0">An exam-practice site is only useful if students can trust its answer key. This site does not let an AI or a script vouch for an answer. Every answer shows its source:</p>
    <ul>
      <li><span class="badge instructor">Matches instructor key</span> The answer computed from the data equals the answer published in the course materials (<em>solutions.xlsx</em> for Excel Part 1, <em>PWC answers.docx</em> for SQL Part 1).</li>
      <li><span class="badge reviewed">Reviewed by …</span> A named person checked the question, re-ran the method, and approved the answer. The sign-off is committed to <a href="https://github.com/${REPO}/commits/main/content/reviews.json">content/reviews.json</a>, so the history shows who approved what and when.</li>
      <li><span class="badge conflict">Instructor key disagrees</span> The published key and the data give different answers. The site does not quietly pick one. It shows both, explains the difference, and waits for a person to decide.</li>
      <li><span class="badge stale">Changed since review</span> A sign-off covers one exact version of the question, method and answer. If any of them changes, the old approval stops counting.</li>
      <li><span class="badge unreviewed">Not yet reviewed</span> Computed from the data with a method a reviewer can re-perform, but no person has checked it yet. A mismatch is a reason to double-check, not proof you are wrong.</li>
    </ul>
    <p>This is the <strong>preparer–reviewer control</strong> from the course. The P-card guidelines require the cardholder, approver and accountant to be three different people, and no one signs off on their own work. Here the build script prepares each answer, and a person reviews and signs it.</p>
  </div>
  <h2>Current status</h2><div class="card">${trustSummaryHtml()}</div>
  <h2>Where the instructor keys and the data disagree</h2>
  <p class="muted">Found by recomputing every published answer from the data. These are the cases a mechanical answer checker would get wrong in one direction or the other.</p>
  ${conflicts.map((q) => h`<div class="card stack" style="margin-top:.8rem"><div class="spread"><a href="#/q/${q.id}"><strong>${q.id}</strong></a>${badge(q)}</div>
    <p style="margin:0">${esc(q.prompt)}</p>
    <p class="small" style="margin:0"><strong>Key:</strong> ${esc(fmtVal(q.instructor.value))} <span class="muted">(${esc(q.instructor.source)})</span> · <strong>Data:</strong> ${esc(q.fields ? q.fields.map((f) => fmt(f.value)).join(" · ") : q.result.rows[0].map(fmt).join(" · "))}</p>
    ${q.instructor.note ? `<p class="small muted" style="margin:0">${esc(q.instructor.note)}</p>` : ""}</div>`).join("")}
  <h2>Disputes</h2>
  <p>Every question has a <em>Dispute it</em> link that opens an issue with the question, the answer version and your attempt filled in. <a href="https://github.com/${REPO}/issues?q=label%3Adispute">Open disputes</a></p>
  <h2>How answers are computed</h2>
  <p>The build script (<a href="https://github.com/${REPO}/tree/main/tools">tools/</a>) loads each course file exactly as students download it and computes every answer from it. Excel answers come from SQL queries that mirror the Excel steps; SQL answers come from running the reference query on the same database file you query in the browser. Answer key built ${esc(S.built)}.</p>`;
};

// ---- reviewer workspace
pages.review = () => {
  const drafts = store.get("drafts", {});
  const items = [...S.questions, ...S.regression.exams.flatMap((e) => e.questions.map((q) => ({ ...q, exam: e })))];
  const filter = store.get("reviewFilter", "todo");
  const todo = (q) => ["unreviewed", "conflict", "stale"].includes(trustOf(q).level);
  const shown = items.filter((q) => filter === "all" || (filter === "todo" ? todo(q) : trustOf(q).level === filter));
  return h`<h1>Reviewer workspace</h1>
  <div class="card stack">
    <p style="margin-top:0">For the teacher or a teaching assistant. Re-perform the method, then approve or withdraw the answer. Your decisions are stored in this browser until you publish them:</p>
    <ol class="small"><li>Click <strong>Download reviews.json</strong>.</li><li>On GitHub, open <a href="https://github.com/${REPO}/edit/main/content/reviews.json">content/reviews.json</a>, replace its contents and commit. The commit is your signature.</li></ol>
    <div class="row"><label for="reviewer">Your name</label><input type="text" id="reviewer" value="${esc(store.get("reviewer", ""))}" placeholder="e.g. Daniela Gelvere">
      <select id="rfilter" aria-label="Show">${[["todo", "Needs a decision"], ["all", "All answers"], ["reviewed", "Reviewed"], ["instructor", "Matches instructor key"]].map(([v, l]) => `<option value="${v}" ${v === filter ? "selected" : ""}>${l}</option>`).join("")}</select>
      <button class="btn primary" id="export">Download reviews.json (${Object.keys(drafts).length} new decision${Object.keys(drafts).length === 1 ? "" : "s"})</button></div>
  </div>
  <p class="small muted">${shown.length} answer${shown.length === 1 ? "" : "s"} shown</p>
  ${shown.slice(0, 200).map((q) => reviewCard(q, drafts[q.id])).join("")}`;
};

function reviewCard(q, draft) {
  const isReg = !!q.exam;
  let method;
  if (isReg) {
    method = `<p class="small">${esc(q.exam.title)} · ${esc(q.exam.equation)} · n = ${q.exam.n}</p><p><strong>Answer:</strong> ${esc(fmt(q.answer))}</p><p class="small muted">${esc(q.explain)}</p><p class="small">Re-perform: regression of <code>sales</code> on ${q.exam.terms.map((t) => `<code>${esc(t)}</code>`).join(", ")} in ${esc(S.regression.source)}.</p>`;
  } else if (q.subject === "Excel") {
    method = q.fields.map((f) => `<p class="small" style="margin-bottom:.2rem"><strong>${esc(f.label)}</strong> = ${esc(fmt(f.value))}</p>${f.method ? `<pre>${esc(f.method)}</pre>` : `<p class="small muted">Computed in tools/build_content.py (Python cleaning step).</p>`}`).join("") + (q.steps ? `<p class="small"><strong>Excel steps given to students:</strong> ${esc(q.steps)}</p>` : "");
  } else {
    method = `<pre>${esc(q.sql)}</pre><p class="small">${q.result.rows.length} rows · first: ${esc(q.result.rows[0].map(fmt).join(" | "))}</p><button class="btn small" data-rerun="${q.id}">Re-run reference query in the browser</button><div class="rerun"></div>`;
  }
  const inst = q.instructor ? `<p class="small"><strong>Instructor key:</strong> ${esc(fmtVal(q.instructor.value))} (${esc(q.instructor.source)})${q.instructor.note ? ` · ${esc(q.instructor.note)}` : ""}</p>` : "";
  return h`<div class="card stack" style="margin-top:.8rem" data-review="${q.id}">
    <div class="spread"><strong>${q.id}</strong><div class="row">${badge(q)}${draft ? `<span class="badge ${draft.decision === "approved" ? "reviewed" : "rejected"}">Draft: ${draft.decision}</span>` : ""}</div></div>
    <p style="margin:0">${esc(q.prompt)}</p>${inst}
    <details><summary>Method and computed answer</summary>${method}</details>
    <div class="row"><input type="text" class="note" placeholder="Note (optional, shown to students)" value="${esc(draft?.note || "")}" style="flex:1 1 260px">
      <button class="btn small" data-decide="approved">Approve</button><button class="btn small" data-decide="rejected">Withdraw</button>${draft ? `<button class="btn small" data-decide="">Clear draft</button>` : ""}</div>
  </div>`;
}

pages.review.after = () => {
  const all = [...S.questions, ...S.regression.exams.flatMap((e) => e.questions)];
  const find = (id) => all.find((q) => q.id === id);
  document.getElementById("reviewer").oninput = (e) => store.set("reviewer", e.target.value.trim());
  document.getElementById("rfilter").onchange = (e) => { store.set("reviewFilter", e.target.value); render(); };
  document.getElementById("export").onclick = () => {
    const drafts = store.get("drafts", {});
    const merged = { reviews: { ...S.reviews, ...drafts } };
    const blob = new Blob([JSON.stringify(merged, null, 1) + "\n"], { type: "application/json" });
    const a = Object.assign(document.createElement("a"), { href: URL.createObjectURL(blob), download: "reviews.json" });
    a.click(); URL.revokeObjectURL(a.href);
  };
  app.onclick = async (e) => {
    const card = e.target.closest("[data-review]");
    if (!card) return;
    const d = e.target.closest("[data-decide]"), rr = e.target.closest("[data-rerun]");
    if (d) {
      const reviewer = store.get("reviewer", "");
      const drafts = store.get("drafts", {});
      const q = find(card.dataset.review);
      if (!d.dataset.decide) delete drafts[q.id];
      else {
        if (!reviewer) { alert("Enter your name first. A sign-off needs a person."); document.getElementById("reviewer").focus(); return; }
        drafts[q.id] = { decision: d.dataset.decide, reviewer, date: new Date().toISOString().slice(0, 10), note: card.querySelector(".note").value.trim(), fingerprint: q.fingerprint };
      }
      store.set("drafts", drafts);
      const y = window.scrollY; render(); window.scrollTo(0, y);
    } else if (rr) {
      const q = S.byId[rr.dataset.rerun], box = card.querySelector(".rerun");
      box.innerHTML = `<p class="small muted">Running…</p>`;
      const r = await sqlRunner.exec(S.datasets[q.ds].db, q.sql);
      if (!r.ok) { box.innerHTML = `<div class="result wrong">${esc(r.error)}</div>`; return; }
      const cmp = compareResults({ ...q.result, order: q.order }, r.result);
      box.innerHTML = `<div class="result ${cmp.status === "correct" ? "correct" : "wrong"}">${cmp.status === "correct" ? "Re-performed in the browser: same result as the stored answer." : "Re-run differs from the stored answer: " + esc(cmp.messages.join(" "))}</div>${resultTable(r.result, 20)}`;
    }
  };
};

// ------------------------------------------------------------------ router
function route() {
  const raw = location.hash.replace(/^#\/?/, "");
  const [path, query] = raw.split("?");
  const parts = path.split("/").filter(Boolean);
  const params = new URLSearchParams(query || "");
  if (!parts.length) return ["home", params];
  if (parts[0] === "exams" && parts[1] === "random") return ["random", params];
  return [parts[0], params, parts[1]];
}

function render() {
  const [name, params, arg] = route();
  const page = pages[name] || (() => `<h1>Page not found</h1><p><a href="#/">Home</a></p>`);
  app.onclick = null;
  app.innerHTML = page(params, arg);
  page.after?.(params, arg);
  const top = { q: "questions", exam: "exams", reg: "interpretation" }[name] || name;
  document.querySelectorAll("[data-nav]").forEach((a) => a.classList.toggle("active", a.dataset.nav === top));
  const h1 = app.querySelector("h1, h2");
  document.title = (h1 ? h1.textContent.slice(0, 60) + " · " : "") + "Exam Practice";
}

let lastPath = "";
window.addEventListener("hashchange", () => {
  const path = location.hash.split("?")[0];
  render();
  if (path !== lastPath) { window.scrollTo(0, 0); app.focus({ preventScroll: true }); }
  lastPath = path;
});

async function boot() {
  const get = (f) => fetch("content/" + f, { cache: "no-cache" }).then((r) => { if (!r.ok) throw new Error(f + " " + r.status); return r.json(); });
  try {
    const [qs, ex, reg, rev] = await Promise.all([get("questions.json"), get("exams.json"), get("regression.json"), get("reviews.json")]);
    Object.assign(S, { questions: qs.questions, datasets: qs.datasets, exams: ex.exams, regression: reg, reviews: rev.reviews || {}, built: qs.built });
    for (const q of S.questions) S.byId[q.id] = q;
    lastPath = location.hash.split("?")[0];
    render();
  } catch (err) {
    app.innerHTML = `<h1>Could not load the questions</h1><p class="muted">${esc(err.message)}</p>`;
  }
}
boot();
