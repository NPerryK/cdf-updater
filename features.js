/*
 * CDF Updater - features (pure logic, no Office.js): import pre-check, grouped changes, undo,
 * CDF lookup, today-at-a-glance + history, morning-update decisions, workbook settings,
 * focused emails. Everything here is tested in tests/features.test.js.
 */
(function (root, factory) {
  var E = root.CdfEngine || (typeof require === "function" ? require("./engine.js") : null);
  var api = factory(E);
  if (typeof module === "object" && module.exports) module.exports = api;
  else root.CdfFeatures = api;
})(typeof self !== "undefined" ? self : this, function (E) {
  "use strict";

  function str(v) { return v == null ? "" : String(v); }
  function trim(v) { return str(v).replace(/ /g, " ").trim(); }
  function up(v) { return trim(v).replace(/\s+/g, " ").toUpperCase(); }
  function keyOf(v) { return E.loose(E.normalizeCdf(v)); }
  function n(v) { var x = Number(v); return isFinite(x) ? x : 0; }
  function pad(x) { return (x < 10 ? "0" : "") + x; }
  function dmy(d) { return pad(d.getDate()) + "/" + pad(d.getMonth() + 1) + "/" + d.getFullYear(); }
  function hm(d) { return pad(d.getHours()) + ":" + pad(d.getMinutes()); }
  function col(L) { return E.colToIndex(L); }

  // =============== 1. import pre-check ===============
  // p: { expect, fileName, lastModified (ms), now (ms), sheet: {rows, headers}, current: {rows, headers} | null }
  // returns [{ level: "bad" | "warn" | "ok", text }]
  var DAY = 86400000;
  function precheckImport(p) {
    var out = [], sheet = p.sheet || { rows: 0, headers: [] };
    if (p.expect && p.fileName && !E.importFileMatches(p.expect, p.fileName)) {
      out.push({ level: "warn", text: "Expected a file named " + p.expect + " — this is " + p.fileName + "." });
    }
    if (sheet.rows <= 1) { out.push({ level: "bad", text: "The file has no data rows." }); return out; }
    if (p.lastModified && p.now && p.now - p.lastModified > DAY) {
      var days = Math.floor((p.now - p.lastModified) / DAY);
      out.push({ level: "warn", text: "The file was saved " + days + " day" + (days > 1 ? "s" : "") + " ago (" + dmy(new Date(p.lastModified)) + ") — is it today's export?" });
    }
    var cur = p.current;
    if (cur && cur.rows > 1) {
      var drop = cur.rows - sheet.rows;
      if (drop > 0 && drop / cur.rows > 0.1) out.push({ level: "bad", text: "Only " + (sheet.rows - 1) + " data rows — the sheet has " + (cur.rows - 1) + " now (" + Math.min(99, Math.round(100 * drop / cur.rows)) + "% fewer)." });
      else if (drop > 0) out.push({ level: "warn", text: drop + " row" + (drop > 1 ? "s" : "") + " fewer than the sheet has now." });
      var have = {};
      (sheet.headers || []).forEach(function (h) { if (trim(h)) have[up(h)] = true; });
      var missing = (cur.headers || []).filter(function (h) { return trim(h) && !have[up(h)]; }).map(trim);
      if (missing.length) out.push({ level: "bad", text: "Column" + (missing.length > 1 ? "s" : "") + " missing in the file: " +
        (missing.length > 6 ? missing.slice(0, 5).join(", ") + " and " + (missing.length - 5) + " more" : missing.join(", ")) + "." });
    }
    if (!out.some(function (x) { return x.level !== "ok"; })) {
      out.push({ level: "ok", text: (sheet.rows - 1) + " data rows" + (cur && cur.rows > 1 ? " (now " + (cur.rows - 1) + ")" : "") + ", columns match." });
    }
    return out;
  }
  function worstLevel(list) {
    return (list || []).some(function (x) { return x.level === "bad"; }) ? "bad"
      : (list || []).some(function (x) { return x.level === "warn"; }) ? "warn" : "ok";
  }

  // =============== 2. changes grouped by CDF ===============
  function groupChangesByCdf(changes) {
    var groups = [], at = {};
    (changes || []).forEach(function (c) {
      var id = c.sheet + "|" + c.key + "|" + c.row;
      if (!at[id]) { at[id] = { key: c.key, sheet: c.sheet, row: c.row, changes: [] }; groups.push(at[id]); }
      at[id].changes.push(c);
    });
    groups.sort(function (a, b) {
      if (a.sheet !== b.sheet) return a.sheet === "GCC" ? -1 : b.sheet === "GCC" ? 1 : a.sheet < b.sheet ? -1 : 1;
      return a.row - b.row;
    });
    groups.forEach(function (g) { g.changes.sort(function (a, b) { return a.col - b.col; }); });
    return groups;
  }
  function changeFacets(changes) {
    var cols = {}, rules = {};
    (changes || []).forEach(function (c) {
      cols[c.colLetter] = (cols[c.colLetter] || 0) + 1;
      rules[c.rule] = (rules[c.rule] || 0) + 1;
    });
    function list(o, sortNum) {
      return Object.keys(o).sort(function (a, b) { return sortNum ? E.colToIndex(a) - E.colToIndex(b) : a.localeCompare(b, undefined, { numeric: true }); })
        .map(function (k) { return { id: k, n: o[k] }; });
    }
    return { columns: list(cols, true), rules: list(rules, false) };
  }
  function filterChanges(changes, f) {
    f = f || {};
    return (changes || []).filter(function (c) {
      return (!f.col || c.colLetter === f.col) && (!f.rule || c.rule === f.rule) &&
        (!f.text || up(c.key).indexOf(up(f.text)) >= 0);
    });
  }

  // =============== 3. undo the last update (from UPDATE LOG) ===============
  // log rows: [time, run name, sheet, row, key, column, before, after, note]
  var UNDO = "Undo";
  // "+3 rows · anchor 907" (rows the update added) / "-1 row · anchor 907" (rows it removed), in the log's New value
  var ROWS_RE = /^([+-])(\d+) rows? · anchor (\d+)$/;
  function undoPlan(logValues) {
    var rows = (logValues || []).slice(1), undone = {}, runs = [], byId = {};
    rows.forEach(function (r) {
      var note = trim(r[8]);
      if (trim(r[1]) === UNDO && note.indexOf("Undo of ") === 0) undone[note.slice(8)] = true;
    });
    rows.forEach(function (r, i) {
      var name = trim(r[1]);
      if (!name || name === UNDO || name === E.LOG_NOTES.newRow) return;
      var id = trim(r[0]) + " " + name;
      if (!byId[id]) { byId[id] = { id: id, stamp: trim(r[0]), name: name, entries: [], rowOps: [], removed: [], runOps: [], last: i }; runs.push(byId[id]); }
      byId[id].last = i;
      if (trim(r[2]) && trim(r[3]) && trim(r[5])) {
        var e = { sheet: trim(r[2]), row: Number(r[3]), col: trim(r[5]), key: trim(r[4]), before: str(r[6]), after: str(r[7]) };
        if (e.col === "RUN") {                     // stage 4: the run appended to BACKUPDATA (only a committed one counts anywhere)
          if (e.after === "COMMITTED") byId[id].runOps.push({ sheet: e.sheet, row: e.row, runId: e.key });
          else byId[id].runOps = byId[id].runOps.filter(function (o) { return o.runId !== e.key; });   // it ended FAILED: nothing to take back
        } else if (e.col === "ROWS") {
          var m = ROWS_RE.exec(trim(r[7]));
          if (m) byId[id].rowOps.push({ sheet: e.sheet, row: e.row, key: e.key, kind: m[1] === "+" ? "insert" : "remove", count: Number(m[2]), anchor: Number(m[3]), values: [], after: e.after });
        } else if (/^Removed row/.test(trim(r[8]))) byId[id].removed.push(e);
        else byId[id].entries.push(e);
      }
    });
    var cand = runs.filter(function (x) { return !undone[x.id] && (x.entries.length || x.rowOps.length || x.runOps.length); });
    if (!cand.length) return null;
    cand.sort(function (a, b) { return a.last - b.last; });
    var run = cand[cand.length - 1];
    function inOp(e, o) { return e.sheet === o.sheet && e.row >= o.row && e.row < o.row + o.count; }
    // the same cell written twice in one run: undo to the first "before"; cells in rows the update added go with the rows
    var seen = {}, entries = [];
    run.entries.forEach(function (e) {
      if (run.rowOps.some(function (o) { return o.kind === "insert" && inOp(e, o); })) return;
      var id = e.sheet + "!" + e.col + e.row;
      if (seen[id]) { seen[id].after = e.after; return; }
      seen[id] = e; entries.push(e);
    });
    var rowOps = run.rowOps.slice().sort(function (a, b) { return b.row - a.row; });   // bottom-up
    rowOps.forEach(function (o) { if (o.kind === "remove") o.values = run.removed.filter(function (e) { return inOp(e, o) && e.key === o.key; }); });
    return { id: run.id, stamp: run.stamp, name: run.name, entries: entries, rowOps: rowOps, runOps: run.runOps.slice() };
  }
  // the run can be taken back while its RUNS row still holds it as COMMITTED; probe = [RunID (A), RunStatus (E)] of that row
  function undoRunCheck(runOps, probes) {
    var ready = [], conflicts = [];
    (runOps || []).forEach(function (o, i) {
      var p = probes[i] || [];
      (trim(p[0]) === o.runId && trim(p[1]) === "COMMITTED" ? ready : conflicts).push(o);
    });
    return { ready: ready, conflicts: conflicts };
  }
  // the run is marked FAILED: its events stop counting (release tables, once-ever releases) and it is no baseline
  function undoRunLogRows(plan, done, stamp) {
    return done.map(function (o) { return [stamp, UNDO, o.sheet, String(o.row), o.runId, "RUN", "COMMITTED", "FAILED", "Undo of " + plan.id]; });
  }
  // can each row operation be undone? insert: the rows are still continuation rows of the same anchor (R refers to it);
  // remove: the anchor still holds the same CDF. probe = R formulas of the added rows / C of the anchor
  function undoRowCheck(rowOps, probes) {
    var ready = [], conflicts = [];
    (rowOps || []).forEach(function (o, i) {
      var p = probes[i], ok;
      if (o.kind === "insert") ok = Array.isArray(p) && p.length === o.count && p.every(function (f) { return contRef(f) === o.anchor; });
      else ok = trim(p).replace(/\s+/g, "").toUpperCase() === trim(o.key).replace(/\s+/g, "").toUpperCase();
      (ok ? ready : conflicts).push(o);
    });
    return { ready: ready, conflicts: conflicts };
  }
  function contRef(f) { var m = /^=IF\(\$([A-Z]{1,3})\$(\d+)="",\s*"",\s*\$\1\$\2\)$/.exec(trim(f)); return m ? Number(m[2]) : null; }
  function undoRowLogRows(plan, done, stamp) {
    return done.map(function (o) { return [stamp, UNDO, o.sheet, String(o.row), o.key, "ROWS", o.after, "", "Undo of " + plan.id]; });
  }
  function sameShown(a, b) {
    var x = trim(a), y = trim(b);
    if (x === y) return true;
    if (/^-?\d+(\.\d+)?$/.test(x) && /^-?\d+(\.\d+)?$/.test(y)) return Number(x) === Number(y);
    return false;
  }
  function undoCheck(entries, currentShown) {
    var ready = [], conflicts = [];
    (entries || []).forEach(function (e, i) {
      var now = currentShown[i];
      if (sameShown(now, e.after)) ready.push(e);
      else conflicts.push({ sheet: e.sheet, row: e.row, col: e.col, key: e.key, before: e.before, after: e.after, now: str(now) });
    });
    return { ready: ready, conflicts: conflicts };
  }
  function valueFromShown(s) {
    var t = str(s);
    if (/^-?\d+(\.\d+)?$/.test(trim(t))) return Number(trim(t));
    return t;
  }
  function undoLogRows(plan, done, stamp) {
    return done.map(function (e) { return [stamp, UNDO, e.sheet, String(e.row), e.key, e.col, e.after, e.before, "Undo of " + plan.id]; });
  }

  // =============== 4. CDF lookup ===============
  var C = { A: 0, B: 1, C: 2, D: 3, E: 4, F: col("F"), G: col("G"), H: col("H"), I: col("I"), J: col("J"), K: col("K"), L: col("L"), M: col("M"),
    N: col("N"), O: col("O"), P: col("P"), R: col("R"), S: col("S"), T: col("T"), U: col("U"), V: col("V"), W: col("W"), X: col("X"),
    Y: col("Y"), Z: col("Z"), AA: col("AA"), AB: col("AB") };
  var G = { A: 0, J: col("J"), AK: col("AK"), AM: col("AM"), AN: col("AN") };
  function cellOf(rows, r, c) { var row = rows[r]; return row && c < row.length ? row[c] : ""; }

  function endOfBlock(cdf, anchor) {
    var end = anchor + 1;
    while (end < cdf.length && !trim(cellOf(cdf, end, C.C)) && !trim(cellOf(cdf, end, C.B)) &&
      [C.R, C.U, C.AA, C.AB].some(function (c) { return trim(cellOf(cdf, end, c)); })) end++;
    return end;
  }

  function matchesQuery(key, q) {
    var k = trim(key);
    if (!k) return false;
    if (/^\d+$/.test(q)) return new RegExp("-0*" + Number(q) + "(\\.|$)").test(k);
    // "740.A", "CDF 740.A", "cdf#961": the number, and the revision when one is given
    var m = /^(?:CDF[\s#:_-]*)?0*(\d{1,4})(?:\.([A-Z]{1,2}|-))?$/i.exec(q);
    if (m) {
      var km = /-0*(\d{1,4})(?:\.([A-Z]{1,2}|-))?(?:\s|$)/i.exec(k);
      return !!km && Number(km[1]) === Number(m[1]) && (!m[2] || (km[2] || "").toUpperCase() === m[2].toUpperCase());
    }
    return keyOf(k) === keyOf(q) || E.loose(k) === E.loose(q);
  }
  function lookupCdf(data, query) {
    var q = trim(query);
    if (!q) return { found: false, error: "Type a CDF number, e.g. 856 or A330-1-856" };
    var cdf = data.cdf || [], hits = [];
    for (var r = 1; r < cdf.length; r++) if (matchesQuery(cellOf(cdf, r, C.C), q)) hits.push(r);
    if (!hits.length) {
      var digits = q.replace(/\D+/g, ""), sugg = [];
      if (digits) for (r = 1; r < cdf.length && sugg.length < 6; r++) {
        var k = trim(cellOf(cdf, r, C.C));
        if (k && k.replace(/\D+/g, "").indexOf(digits) >= 0) sugg.push(k);
      }
      return { found: false, error: "No CDF " + q + " in the CDF sheet", suggestions: sugg };
    }
    var a = hits[0], end = endOfBlock(cdf, a);
    var key = trim(cellOf(cdf, a, C.C));
    function f(c) { return trim(cellOf(cdf, a, c)); }
    var res = {
      found: true, key: key, row: a + 1, others: hits.slice(1).map(function (h) { return trim(cellOf(cdf, h, C.C)) + " (row " + (h + 1) + ")"; }),
      fields: { status: f(C.R), station: f(C.S), signature: f(C.T), received: cellOf(cdf, a, C.F), priority: f(C.G), description: f(C.H), notes: str(cellOf(cdf, a, C.E)),
        tags: f(C.I), dwg: f(C.J), dwgName: f(C.K), nrc: f(C.L), dwgRev: f(C.M), author: f(C.N), adviser: f(C.O), sendTo: f(C.P),
        reason: f(C.Y), classification: f(C.Z), revision: f(C.D) },
      solutions: [], gcc: null, history: []
    };
    for (r = a; r < end; r++) {
      var aa = trim(cellOf(cdf, r, C.AA)), u = trim(cellOf(cdf, r, C.U));
      if (aa || u) res.solutions.push({ row: r + 1, doc: aa, rev: trim(cellOf(cdf, r, C.AB)), mantra: u,
        v: trim(cellOf(cdf, r, C.V)), w: trim(cellOf(cdf, r, C.W)), x: trim(cellOf(cdf, r, C.X)) });
    }
    var gcc = data.gcc || [];
    for (r = 1; r < gcc.length; r++) {
      if (keyOf(cellOf(gcc, r, G.A)) === keyOf(key)) {
        res.gcc = { row: r + 1, status: trim(cellOf(gcc, r, G.J)), station: trim(cellOf(gcc, r, G.AM)),
          assignee: trim(cellOf(gcc, r, G.AN)).replace(/^'+/, ""), finalApproval: trim(cellOf(gcc, r, G.AK)) };
        break;
      }
    }
    var log = data.log || [];
    for (r = 1; r < log.length; r++) {
      if (keyOf(cellOf(log, r, 4)) === keyOf(key) && trim(cellOf(log, r, 5))) {
        res.history.push({ when: trim(cellOf(log, r, 0)), col: trim(cellOf(log, r, 5)), before: str(cellOf(log, r, 6)), after: str(cellOf(log, r, 7)),
          note: trim(cellOf(log, r, 8)), by: trim(cellOf(log, r, 1)) });
      }
    }
    res.history.reverse();
    return res;
  }

  // =============== 5. today at a glance + history ===============
  function tableRows(summary, name) {
    var t = (summary.tables || []).filter(function (x) { return x.name === name; })[0];
    var o = {};
    if (t) t.rows.forEach(function (r) { o[r[0]] = r[1]; });
    return o;
  }
  function snapshot(summary, info, now) {
    var st = tableRows(summary, "CDF STATUS"), gs = tableRows(summary, "GCC CURRENT STATUS"), pend = {};
    (summary.pendingPeople || []).forEach(function (p) { pend[p.person] = n(tableRows(summary, p.table).TOTAL); });
    var stations = {};
    Object.keys(gs).forEach(function (k) { if (k !== "TOTAL") stations[k] = n(gs[k]); });
    return {
      at: (now || new Date()).toISOString(), open: n(st.Open), released: n(st.Released), cancelled: n(st.Cancelled), total: n(st.TOTAL),
      stations: stations, pending: pend, gccRows: info ? info.gccRows : null, onlyCdf: info ? (info.onlyCdf || []).slice() : []
    };
  }
  function glance(now, prev) {
    function d(a, b) { return prev && b != null ? a - b : null; }
    var pendNow = 0, pendPrev = 0;
    Object.keys(now.pending || {}).forEach(function (k) { pendNow += n(now.pending[k]); });
    if (prev) Object.keys(prev.pending || {}).forEach(function (k) { pendPrev += n(prev.pending[k]); });
    var fa = n((now.stations || {})["9 - Final Approval"]), conv = n((now.stations || {})["3 - Conversion Engineer"]);
    var pst = prev ? prev.stations || {} : {};
    return [
      { id: "open", label: "Open", value: now.open, delta: d(now.open, prev && prev.open) },
      { id: "conv", label: "At conversion eng.", value: conv, delta: prev ? d(conv, n(pst["3 - Conversion Engineer"])) : null },
      { id: "fa", label: "At final approval", value: fa, delta: prev ? d(fa, n(pst["9 - Final Approval"])) : null },
      { id: "pending", label: "Ready to sign", value: pendNow, delta: prev ? pendNow - pendPrev : null },
      { id: "released", label: "Released", value: now.released, delta: d(now.released, prev && prev.released) }
    ];
  }
  function pushSnapshot(list, snap, max) {
    var out = (list || []).slice(), last = out[out.length - 1];
    if (last && last.at && snap.at && last.at.slice(0, 16) === snap.at.slice(0, 16)) out[out.length - 1] = snap;
    else out.push(snap);
    max = max || 60;
    return out.length > max ? out.slice(out.length - max) : out;
  }
  function previousSnapshot(list, nowIso) {
    var l = list || [];
    for (var i = l.length - 1; i >= 0; i--) if (!nowIso || l[i].at.slice(0, 16) !== nowIso.slice(0, 16)) return l[i];
    return null;
  }
  function trendSeries(list, metric, max) {
    var l = (list || []).slice(-(max || 14));
    return l.map(function (s) { return { at: s.at, value: metric === "pending" ? Object.keys(s.pending || {}).reduce(function (a, k) { return a + n(s.pending[k]); }, 0) : n(s[metric]) }; });
  }
  function stationBars(snap) {
    var st = snap.stations || {}, keys = Object.keys(st);
    var max = keys.reduce(function (m, k) { return Math.max(m, st[k]); }, 0) || 1;
    return keys.map(function (k) { return { key: k, label: k.replace(/^\d+\s*-\s*/, ""), code: (/^(\d+)/.exec(k) || [])[1] || "", value: st[k], pct: Math.round(100 * st[k] / max) }; });
  }

  // =============== 6. morning update: when to stop ===============
  // res: CdfAlgo.run result; o: { prev snapshot, maxCells (40), precheck: [{level,text}] }
  function morningDecision(res, o) {
    o = o || {};
    var reasons = [], maxCells = o.maxCells || 40;
    (o.precheck || []).forEach(function (x) { if (x.level === "bad") reasons.push({ level: "bad", text: x.text }); });
    if (res.errors && res.errors.length) reasons.push({ level: "bad", text: res.errors.length + " algorithm error" + (res.errors.length > 1 ? "s" : "") + " — " + res.errors[0].message + (res.errors.length > 1 ? " …" : "") });
    (res.checks || []).forEach(function (c) { if (!c.ok) reasons.push({ level: "bad", text: "Control total failed: " + c.name + " (" + c.detail + ")" }); });
    var cells = (res.cdfChanges || []).length + (res.gccChanges || []).length;
    // stage 1 (GCC) usually moves a few cells; stage 2 (Mantra) moves hundreds on a normal day (314 on 05/11)
    var s1 = res.stage1Cells != null ? res.stage1Cells : cells, s2 = res.stage2 || {}, s2Cells = s2.ran ? s2.cells : 0, maxMantra = o.maxMantraCells || 600;
    if (s1 > maxCells) reasons.push({ level: "warn", text: s1 + " cells would change from GCC — more than usual (" + maxCells + "). Check them before applying." });
    if (s2Cells > maxMantra) reasons.push({ level: "warn", text: s2Cells + " cells would change from Mantra — more than usual (" + maxMantra + "). Check them before applying." });
    if (s2.ran && s2.errors && s2.errors.length) reasons.push({ level: "warn", text: s2.errors.length + " Mantra solution" + (s2.errors.length > 1 ? "s" : "") + " to check — " + (s2.errors[0].key ? s2.errors[0].key + ": " : "") + s2.errors[0].message + (s2.errors.length > 1 ? " …" : "") });
    if (s2.rowsRemoved) reasons.push({ level: "warn", text: s2.rowsRemoved + " solution row" + (s2.rowsRemoved > 1 ? "s" : "") + " would be removed from sheet CDF. Check which before applying." });
    if (s2.bad) reasons.push({ level: "bad", text: s2.skipped });
    var prev = o.prev;
    if (prev && res.info) {
      if (prev.gccRows != null && res.info.gccRows < prev.gccRows) reasons.push({ level: "warn", text: (prev.gccRows - res.info.gccRows) + " CDFs fewer in GCC than in the last run (" + prev.gccRows + " → " + res.info.gccRows + ")." });
      var was = {}; (prev.onlyCdf || []).forEach(function (k) { was[k] = true; });
      var gone = (res.info.onlyCdf || []).filter(function (k) { return !was[k]; });
      if (gone.length) reasons.push({ level: "warn", text: "No longer in GCC: " + gone.join(", ") + "." });
    }
    var rep = res.reports || {}, h = res.history || {};
    if (h.bad) reasons.push({ level: "warn", text: h.skipped });
    if (h.ran && !h.out.commits) reasons.push({ level: "warn", text: "The run would be written to BACKUPDATA as FAILED — " + h.out.errors.length + " error" + (h.out.errors.length > 1 ? "s" : "") + " for sheet ERROR; its events wouldn't count and EMAIL DRAFT would say NOT CURRENT." });
    return { stop: reasons.length > 0, reasons: reasons, cells: cells, rows: (s2.rowsAdded || 0) + (s2.rowsRemoved || 0), reports: !!(rep.ran && (rep.summaryDiffers || rep.emailDiffers)),
      history: !!(h.ran && h.out.events.length) };
  }

  // =============== 7. settings kept inside the workbook ===============
  var PRIVATE_OK = ["STATION", "SUPPLIER", "FINAL_APPROVAL_RELEASE", "PENDING_TABLE", "DUPLICATE_CDF", "PERSON_GROUP", "PERSON_ALIAS", "EMAIL_GROUP", "CLAUDE_PROJECT", "REPORT"];
  function privateLines(text) {
    var bad = [];
    str(text).split(/\r?\n/).forEach(function (line, i) {
      var t = line.replace(/#.*$/, "").trim();
      if (!t) return;
      var w = t.split(/\s+/)[0].toUpperCase();
      if (PRIVATE_OK.indexOf(w) < 0) bad.push({ line: i + 1, message: w + " can't be in the workbook settings (allowed: " + PRIVATE_OK.join(", ") + ")" });
    });
    return bad;
  }
  // the CDF algorithm needs the people settings that rules.txt no longer carries (they name people): kept in the workbook
  function requirePeopleSettings(m) {
    if (!m.cdfAlgorithm || !m.cdfConfig) return m;
    var miss = [];
    if (!m.cdfConfig.finalApprovalSigner) miss.push("FINAL_APPROVAL_RELEASE (rule 10)");
    if (!(m.cdfConfig.pendingTables || []).length) miss.push("PENDING_TABLE (100.5)");
    if (miss.length) m.errors = (m.errors || []).concat([{ line: 0, setting: true,
      message: "The workbook settings need " + miss.join(" and ") + ". These lines name people, so they're kept in the workbook, not in the public rules.txt — paste them in SETUP › Workbook settings and save." }]);
    return m;
  }
  function mergeRules(pub, privText) {
    var m = JSON.parse(JSON.stringify(pub));
    m.privateLines = 0;
    if (!trim(privText)) return requirePeopleSettings(m);
    var bad = privateLines(privText);
    var priv = E.parseRules(privText);
    var errs = bad.concat(priv.errors || []).map(function (e) { return { line: e.line, message: "Workbook settings, line " + e.line + ": " + e.message }; });
    m.errors = (m.errors || []).concat(errs);
    if (bad.length) return m;
    var pc = priv.cdfConfig, mc = m.cdfConfig;
    Object.keys(pc.stations).forEach(function (k) { mc.stations[k] = pc.stations[k]; });
    Object.keys(pc.suppliers).forEach(function (k) { mc.suppliers[k] = pc.suppliers[k]; });
    if (pc.finalApprovalSigner) mc.finalApprovalSigner = pc.finalApprovalSigner;
    if (pc.pendingTables.length) mc.pendingTables = pc.pendingTables;
    pc.duplicates.forEach(function (d) { if (mc.duplicates.indexOf(d) < 0) mc.duplicates.push(d); });
    if (pc.personGroups.length) mc.personGroups = pc.personGroups;        // the workbook's people list replaces the public one
    if (pc.personAliases.length) mc.personAliases = (mc.personAliases || []).concat(pc.personAliases);
    priv.emailGroups.forEach(function (g) {
      m.emailGroups = m.emailGroups.filter(function (x) { return x.name !== g.name; });
      m.emailGroups.push(g);
    });
    if (priv.claudeProject) m.claudeProject = priv.claudeProject;
    m.reports = m.reports.concat(priv.reports);
    m.privateLines = str(privText).split(/\r?\n/).filter(function (l) { return l.replace(/#.*$/, "").trim(); }).length;
    return requirePeopleSettings(m);
  }
  // the settings lines of rules.txt, to move them into the workbook
  function extractPrivate(rulesText) {
    var keep = [], moved = [];
    str(rulesText).split(/\r?\n/).forEach(function (line) {
      var w = line.replace(/#.*$/, "").trim().split(/\s+/)[0].toUpperCase();
      if (w && ["STATION", "SUPPLIER", "FINAL_APPROVAL_RELEASE", "PENDING_TABLE", "DUPLICATE_CDF", "PERSON_GROUP", "PERSON_ALIAS", "EMAIL_GROUP", "CLAUDE_PROJECT"].indexOf(w) >= 0) moved.push(line.trim());
      else keep.push(line);
    });
    return { moved: moved.join("\n"), rest: keep.join("\n") };
  }

  // workbook storage: one custom XML part per namespace, JSON kept as URI-encoded text
  function encodeStore(ns, obj) {
    return '<cdfUpdater xmlns="' + ns + '">' + encodeURIComponent(JSON.stringify(obj)) + "</cdfUpdater>";
  }
  function decodeStore(xml) {
    var m = /<cdfUpdater[^>]*>([^<]*)<\/cdfUpdater>/.exec(str(xml));
    if (!m) return null;
    try { return JSON.parse(decodeURIComponent(m[1])); } catch (e) { return null; }
  }

  // =============== 8. focused emails ===============
  // kind: { type: "supplier", supplier } | { type: "pending", table, person }
  function focusKinds(summary) {
    var out = [];
    Object.keys((summary && summary.whereLists) || {}).forEach(function (s) {
      out.push({ id: "sup:" + s, type: "supplier", supplier: s, label: "Open at conversion eng. — " + s + " (" + summary.whereLists[s].length + ")" });
    });
    ((summary && summary.pendingPeople) || []).forEach(function (p) {
      out.push({ id: "pend:" + p.table, type: "pending", table: p.table, person: p.person, label: "Ready to sign — " + p.person });
    });
    return out;
  }
  function cellsOf(row) { return row.map(function (v) { return { v: str(v), rowspan: 1 }; }); }
  function buildFocusDoc(kind, summary, now, signature) {
    now = now || new Date();
    var asOf = "AS OF " + dmy(now) + " " + hm(now) + " (Asia/Jerusalem)";
    if (kind.type === "supplier") {
      var list = (summary.whereLists || {})[kind.supplier] || [];
      return {
        subject: "A330-300 P2F — Open CDFs at Conversion Engineer — " + kind.supplier + " — " + dmy(now), dir: "ltr",
        opening: "Hello, below are the open CDFs at Conversion Engineer for " + kind.supplier + " — " + asOf + ".",
        empty: !list.length, signature: signature || ["Best regards,"],
        tables: [{ title: "OPEN CDFS AT CONVERSION ENGINEER — " + kind.supplier, headers: ["CDF NUMBER", "SIGNATURE POSITION", "DWG.", "DWG. NAME", "PROBLEM DESCRIPTION"], widths: [105, 115, 170, 190, 380],
          rows: list.map(function (x) { return cellsOf([x.key, x.signature, x.dwg, x.title, x.desc]); }).concat([cellsOf(["TOTAL", "", "", "", String(list.length)])]) }]
      };
    }
    var t = (summary.tables || []).filter(function (x) { return x.name === kind.table; })[0];
    var rows = t ? t.rows : [];
    return {
      subject: "A330-300 P2F — CDFs ready for your signature — " + (kind.person || kind.table) + " — " + dmy(now), dir: "rtl",
      opening: "שלום " + (kind.person || "").split(" ")[0] + ", להלן החריגים שממתינים לחתימתך ב-Final Approval — " + asOf + ".",
      empty: rows.length <= 1, signature: signature || ["בברכה,"],
      tables: [{ title: kind.table, headers: t ? t.headers : ["CDF NUMBER", "SOLUTION / DOCUMENT"], rows: rows.map(cellsOf) }]
    };
  }

  // =============== 9. link with My Notes (Outlook) ===============
  // The CDF status of every CDF, for My Notes: one compact row per CDF.
  function bridgeStatus(cdf, summary, now, from, stations) {
    stations = stations || {};
    var ready = {};
    ((summary && summary.pendingPeople) || []).forEach(function (p) {
      var t = (summary.tables || []).filter(function (x) { return x.name === p.table; })[0];
      if (t) t.rows.forEach(function (r) { if (/A330/i.test(str(r[0]))) ready[keyOf(r[0])] = p.person; });
    });
    function v(x) { x = trim(x); return x === "-" ? "" : x; }
    var rows = [];
    cdf = cdf || [];
    for (var a = 1; a < cdf.length; a++) {
      var key = trim(cellOf(cdf, a, C.C));
      if (!/A330/i.test(key)) continue;
      var end = endOfBlock(cdf, a);
      var docs = [], isOpen = /^open$/i.test(trim(cellOf(cdf, a, C.R)));
      for (var r = a; r < end && docs.length < 8 && isOpen; r++) {   // documents only for open CDFs (keeps the text small)
        var doc = v(cellOf(cdf, r, C.AA)), u = v(cellOf(cdf, r, C.U));
        if (doc || u) docs.push([doc, v(cellOf(cdf, r, C.AB)), u]);
      }
      var rec = cellOf(cdf, a, C.F);
      rows.push([key, v(cellOf(cdf, a, C.R)), v(cellOf(cdf, a, C.S)), v(cellOf(cdf, a, C.T)), v(cellOf(cdf, a, C.P)),
        v(cellOf(cdf, a, C.J)), v(cellOf(cdf, a, C.K)), typeof rec === "number" ? E.serialToDateString(rec) : v(rec), docs, ready[keyOf(key)] || "",
        stations[key] ? stations[key][1] : "", stations[key] ? stations[key][2] : 0, stations[key] ? stations[key][3] || "" : "", str(cellOf(cdf, a, C.E))]);
      a = end - 1;
    }
    return { at: (now || new Date()).toISOString(), from: from || "CDF Updater", count: rows.length, rows: rows };
  }
  // ---- days at station: { key: [station, since "yyyy-mm-dd", exact 1/0, first seen "yyyy-mm-dd"] }, kept in the workbook ----
  // A new station → today (exact). First time seen → the last change of S in UPDATE LOG (exact), or else
  // "since" is the day the CDF was received (the most it can be) and "first seen" the least it can be.
  function ymdOf(d) { return d.getFullYear() + "-" + pad(d.getMonth() + 1) + "-" + pad(d.getDate()); }
  function logDay(stamp) { var m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})/.exec(trim(stamp)); return m ? m[3] + "-" + pad(+m[2]) + "-" + pad(+m[1]) : ""; }
  function trackStations(map, cdf, log, now) {
    var out = {}, changed = false, today = ymdOf(now || new Date()), lastS = {};
    (log || []).forEach(function (r, i) {
      if (!i || trim(r[5]) !== "S" || trim(r[2]) !== "CDF") return;
      var k = keyOf(r[4]), day = logDay(r[0]);
      if (k && day && (!lastS[k] || lastS[k].day <= day)) lastS[k] = { day: day, after: trim(r[7]) };
    });
    map = map || {};
    for (var a = 1; a < (cdf || []).length; a++) {
      var key = trim(cellOf(cdf, a, C.C));
      if (!/A330/i.test(key) || !/^open$/i.test(trim(cellOf(cdf, a, C.R)))) continue;
      var st = trim(cellOf(cdf, a, C.S)), was = map[key], k = keyOf(key);
      if (was && was[0] === st) { out[key] = was; continue; }
      changed = true;
      if (was) { out[key] = [st, today, 1, today]; continue; }
      if (lastS[k] && lastS[k].after === st) { out[key] = [st, lastS[k].day, 1, today]; continue; }
      var rec = cellOf(cdf, a, C.F), d = logDay(typeof rec === "number" ? E.serialToDateString(rec) : rec);
      out[key] = [st, d || today, 0, today];
    }
    if (Object.keys(map).some(function (k) { return !out[k]; })) changed = true;
    return { map: out, changed: changed };
  }
  function daysSince(day, now) {
    var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(str(day));
    if (!m) return null;
    var a = new Date(+m[1], +m[2] - 1, +m[3]), b = new Date(now || new Date()); b.setHours(0, 0, 0, 0);
    return Math.max(0, Math.round((b - a) / 864e5));
  }
  // { days, min, exact, stuck }: exact → days is the answer; else days is the most and min the least it can be
  function daysAt(entry, now) {
    if (!entry) return null;
    var days = daysSince(entry[1], now), exact = !!entry[2], min = exact ? days : daysSince(entry[3], now);
    if (days == null) return null;
    if (min == null || min > days) min = exact ? days : 0;
    return { days: days, min: min, exact: exact, stuck: (exact ? days : min) >= STUCK };
  }

  // ---- the lists behind the numbers on the display ----
  // kind: open | conv | fa | pending | released | stuck | station:<S value>
  var STUCK = 14;
  function glanceList(kind, data, now) {
    var cdf = data.cdf || [], stations = data.stations || {}, summary = data.summary, gccKeys = data.gccKeys;
    var ready = {};
    ((summary && summary.pendingPeople) || []).forEach(function (p) {
      var t = (summary.tables || []).filter(function (x) { return x.name === p.table; })[0];
      if (t) t.rows.forEach(function (r) { if (/A330/i.test(str(r[0]))) ready[keyOf(r[0])] = p.person; });
    });
    var titles = { open: "Open CDFs", conv: "At conversion engineer", fa: "At final approval", pending: "Ready to sign", released: "Released", stuck: STUCK + "+ days at the station",
      ageing: "Open at the same station for 7–13 days", noDoc: "Open without an engineering document", moved7: "Open CDFs that changed station in the last 7 days",
      released7: "CDFs released in the last 7 days" };
    var title = titles[kind] || (kind.indexOf("station:") === 0 ? "Open — " + kind.slice(8).replace(/^\d+\s*-\s*/, "") : kind);
    var rows = [];
    var recent = kind === "moved7" || kind === "released7" ? recentChanges(data.log, now) : null;
    for (var a = 1; a < cdf.length; a++) {
      var key = trim(cellOf(cdf, a, C.C));
      if (!/A330/i.test(key)) continue;
      var status = trim(cellOf(cdf, a, C.R)), st = trim(cellOf(cdf, a, C.S)), open = /^open$/i.test(status);
      var at = open ? daysAt(stations[key], now) : null;
      var hit = kind === "open" ? open : kind === "conv" ? open && /^3\s*-/.test(st) : kind === "fa" ? open && /^9\s*-/.test(st)
        : kind === "pending" ? !!ready[keyOf(key)] : kind === "released" ? /^released$/i.test(status)
        : kind === "stuck" ? !!(at && at.stuck)
        : kind === "ageing" ? !!(at && at.min >= 7 && at.min < STUCK)
        : kind === "noDoc" ? open && up(cellOf(cdf, a, C.U)) === "NO ENG. DOC." && endOfBlock(cdf, a) === a + 1
        : kind === "moved7" ? open && recent.S[key] === st
        : kind === "released7" ? /^released$/i.test(status) && recent.R[key] === "Released"
        : kind.indexOf("station:") === 0 ? open && st === kind.slice(8) : false;
      if (!hit) continue;
      rows.push({ key: key, row: a + 1, status: status, station: st, signature: trim(cellOf(cdf, a, C.T)).replace(/^-$/, ""), supplier: trim(cellOf(cdf, a, C.P)).replace(/^-$/, ""),
        days: at ? at.days : null, min: at ? at.min : null, exact: !!(at && at.exact), stuck: !!(at && at.stuck), ready: ready[keyOf(key)] || "", inGcc: gccKeys ? !!gccKeys[keyOf(key)] : true });
    }
    function sortDays(r) { return r.days == null ? -1 : r.exact ? r.days : r.min; }
    rows.sort(function (x, y) { return sortDays(y) - sortDays(x) || (y.days || 0) - (x.days || 0) || (x.row - y.row); });
    var notInGcc = rows.filter(function (r) { return !r.inGcc; }).map(function (r) { return r.key; });
    return { title: title, rows: rows, notInGcc: notInGcc };
  }
  function recentChanges(log, now) {
    var recent = { S: Object.create(null), R: Object.create(null) };
    var today = ymdOf(now), start = new Date(now);
    start.setDate(start.getDate() - 6);
    var since = ymdOf(start);
    (log || []).slice(1).forEach(function (row) {
      var col = trim(row[5]), day = logDay(row[0]), before = trim(row[6]), after = trim(row[7]);
      if (trim(row[2]) !== "CDF" || !recent[col] || day < since || day > today) return;
      var key = trim(row[4]);
      if (!before || before === after || trim(row[1]) === UNDO) delete recent[col][key];
      else recent[col][key] = after;
    });
    return recent;
  }
  function stuckCount(stations, now) {
    var exact = 0, atLeast = 0;
    Object.keys(stations || {}).forEach(function (k) {
      var d = daysAt(stations[k], now);
      if (d && d.stuck) { if (d.exact) exact++; else atLeast++; }
    });
    return { exact: exact, atLeast: atLeast, total: exact + atLeast };
  }

  // what was pasted from My Notes, in a line
  function notesInfo(data) {
    if (!data || !data.notes) return { notes: 0, cdfs: 0, open: 0 };
    var cd = {}, open = 0;
    data.notes.forEach(function (n) { (n.refs || []).forEach(function (r) { cd[r] = true; }); if (n.state !== "Done" && n.state !== "Info") open++; });
    return { notes: data.notes.length, cdfs: Object.keys(cd).length, open: open };
  }

  return {
    precheckImport: precheckImport, worstLevel: worstLevel,
    groupChangesByCdf: groupChangesByCdf, changeFacets: changeFacets, filterChanges: filterChanges,
    undoPlan: undoPlan, undoRowCheck: undoRowCheck, undoRowLogRows: undoRowLogRows, undoRunCheck: undoRunCheck, undoRunLogRows: undoRunLogRows, contRef: contRef, undoCheck: undoCheck, valueFromShown: valueFromShown, undoLogRows: undoLogRows, UNDO: UNDO,
    lookupCdf: lookupCdf,
    snapshot: snapshot, glance: glance, pushSnapshot: pushSnapshot, previousSnapshot: previousSnapshot, trendSeries: trendSeries, stationBars: stationBars,
    morningDecision: morningDecision,
    mergeRules: mergeRules, requirePeopleSettings: requirePeopleSettings, privateLines: privateLines, extractPrivate: extractPrivate, PRIVATE_OK: PRIVATE_OK,
    encodeStore: encodeStore, decodeStore: decodeStore,
    focusKinds: focusKinds, buildFocusDoc: buildFocusDoc,
    bridgeStatus: bridgeStatus, notesInfo: notesInfo,
    trackStations: trackStations, daysSince: daysSince, daysAt: daysAt, glanceList: glanceList, stuckCount: stuckCount, STUCK: STUCK
  };
});
