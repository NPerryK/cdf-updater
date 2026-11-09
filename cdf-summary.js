/*
 * CDF algorithm — stage 3: sheet CDF SUMMARY (all tables) and sheet EMAIL DRAFT (algorithm 100, 150).
 * Pure: reads the CDF sheet as stages 1–2 leave it, returns the tables and where each cell goes.
 * Office.js work is in taskpane.js.
 *
 *   100.3–100.5, 100.8  CDF STATUS, GCC CURRENT STATUS, GCC SIGNATURE POSITION, MOR / IGOR, WHERE ARE THE CDFS
 *                       (numbers from stage 1, written as the same COUNTIFS formulas the macro writes)
 *   100.7               MANTRA SIGNATURE POSITION, MANTRA WHO
 *   100.11              PERSON GROUP / DEPARTMENT
 *   100.10              REJECTED (+ REJECTED CDFS, REJECTED DOCUMENTS)
 *   100.9               PERSONS – GROUP – SIGNATURE – CDF – PENDING MANTRA DOCUMENTS
 *   100.13–100.17       the four release tables — carried over from the sheet as they are (stage 4 computes them)
 *   100.2.1             every table but CDF STATUS and the release tables: Open CDFs only, filtered before counting
 *   150                 EMAIL DRAFT: subject, Hebrew opening, six tables, signature (kept from the sheet)
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.CdfSummary = factory();
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  function str(v) { return v == null ? "" : String(v); }
  function trim(v) { return str(v).replace(/ /g, " ").replace(/\r\n?/g, "\n").trim(); }
  function upper(v) { return trim(v).toUpperCase(); }
  function isEmpty(v) { return trim(v) === ""; }
  function col(L) { var n = 0; for (var i = 0; i < L.length; i++) n = n * 26 + L.charCodeAt(i) - 64; return n - 1; }
  function letter(i) { var n = i + 1, s = ""; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
  function cell(rows, r, c) { var row = rows[r]; return row && c < row.length ? row[c] : ""; }
  function dashKey(v) { return upper(v).replace(/[‐-―\-]/g, "-").replace(/\s+/g, " "); }
  function personKey(v) { return upper(v).replace(/\s*[-‐-―]\s*/g, "-").replace(/\s+/g, " "); }

  var C = { C: col("C"), R: col("R"), U: col("U"), V: col("V"), W: col("W"), X: col("X"), AA: col("AA"), AB: col("AB") };
  var NOT_PENDING_U = ["", "-", "4 - RELEASED", "ANSWER AT GCC", "NO ENG. DOC.", "CHECK SOLUTION"];
  // 100.11: the groups in their order (others found are added after them, so nothing goes missing)
  var GROUPS = ["IAI", "IAI - Design Engineering", "IAI - Electrical Engineering", "IAI - Stress Department", "IAI - A330 Management",
    "IAI - Production Engineering", "IAI - Production Technology", "IAI - Tools Engineering", "IAI - Interior Engineering",
    "IAI - Engineering Center", "IAI - System Engineering", "ANN/SIE", "ANCRA"];
  var T = {
    status: "CDF STATUS", gcc: "GCC CURRENT STATUS", sig: "GCC SIGNATURE POSITION", where: "WHERE ARE THE CDFS",
    msig: "MANTRA SIGNATURE POSITION", who: "MANTRA WHO", groups: "PERSON GROUP / DEPARTMENT", rejected: "REJECTED",
    persons: "PERSONS – GROUP – SIGNATURE – CDF – PENDING MANTRA DOCUMENTS",
    cdf24: "CDFS RELEASED — LAST 24 HOURS", cdf7: "CDFS RELEASED — LAST 7 DAYS",
    doc24: "DOCUMENTS RELEASED / GCC ANSWERS — LAST 24 HOURS", doc7: "DOCUMENTS RELEASED / GCC ANSWERS — LAST 7 DAYS"
  };
  // 100.1.3 column widths (Excel characters)
  var WIDTHS = { A: 30, B: 10, C: 18, D: 25, E: 28, F: 45, G: 18, H: 28 };
  var DUPLICATE_RAW = "A330-1-052 (another)";

  // "A330-1-740.A" → sortable [740, 1]; "-" before letters
  function keyParts(k) {
    var m = /^A330-1-(\d+)\s*(?:\.\s*([A-Z]{1,2}|-))?(.*)$/i.exec(trim(k));
    if (!m) return [Infinity, 0, upper(k)];
    var rev = !m[2] || m[2] === "-" ? 0 : m[2].length === 1 ? m[2].toUpperCase().charCodeAt(0) - 64 : 26 + (m[2].toUpperCase().charCodeAt(0) - 64) * 26 + (m[2].toUpperCase().charCodeAt(1) - 64);
    return [Number(m[1]), rev, upper(m[3])];
  }
  function keyCmp(a, b) {
    var x = keyParts(a), y = keyParts(b);
    return x[0] - y[0] || x[1] - y[1] || (x[2] < y[2] ? -1 : x[2] > y[2] ? 1 : 0);
  }
  function txtCmp(a, b) { a = upper(a); b = upper(b); return a < b ? -1 : a > b ? 1 : 0; }

  // ---------------------------------------------------------------- people
  // the name as written in the people list (PERSON_GROUP); otherwise as in X
  function makeNames(cfg, mantra, cdf) {
    var canon = {}, alias = {};
    ((cfg && cfg.personGroups) || []).forEach(function (g) { (g.people || []).forEach(function (p) { canon[personKey(p)] = trim(p); }); });
    ((cfg && cfg.personAliases) || []).forEach(function (a) { alias[personKey(a.from)] = personKey(a.to); });
    var people = mantra && mantra.makePeople ? mantra.makePeople(cfg, cdf) : null;
    return {
      people: people,
      canon: function (x) {
        var k = personKey(x);
        if (alias[k]) k = alias[k];
        if (canon[k]) return canon[k];
        var w = k.split(" ").sort().join(" ");
        var hit = Object.keys(canon).filter(function (c) { return c.split(" ").sort().join(" ") === w; });
        if (hit.length === 1) return canon[hit[0]];
        return titleCase(x);
      },
      group: function (x, v) {
        var g = people && mantra ? mantra.groupOf(x, people) : null;
        if (g) return mantra.groupLabel(g) || "-";
        return mantra ? mantra.groupLabel(v) || "-" : "-";
      }
    };
  }
  // a name Mantra wrote in capitals and the people list doesn't have: "DANA BEN-ARI" → "Dana Ben-Ari"
  function titleCase(x) {
    var t = trim(x).replace(/\s+/g, " ");
    if (t !== t.toUpperCase() || !/[A-Z]/.test(t)) return t;
    return t.toLowerCase().replace(/(^|[\s\-'’])([a-z])/g, function (m, a, b) { return a + b.toUpperCase(); });
  }
  function mainGroup(g) { return /^IAI\b/.test(g) ? "IAI" : g; }   // 100.9.6

  // ---------------------------------------------------------------- 100.2.1: Open CDFs, documents not released
  // one entry per solution row with a pending or rejected document
  function relations(cdf, algo) {
    var out = [];
    algo.buildBlocks(cdf).forEach(function (b) {
      if (trim(cell(cdf, b.anchor, C.R)) !== "Open") return;
      var key = trim(cell(cdf, b.anchor, C.C));
      b.rows.forEach(function (r, i) {
        var U = upper(cell(cdf, r, C.U)), AA = trim(cell(cdf, r, C.AA)), AB = trim(cell(cdf, r, C.AB)) || "-";
        if (NOT_PENDING_U.indexOf(U) >= 0 || !AA || AA === "-" || upper(AA) === "ANSWER AT GCC" || upper(AA) === "CHECK SOLUTION") return;
        var W = trim(cell(cdf, r, C.W)), X = trim(cell(cdf, r, C.X)), V = trim(cell(cdf, r, C.V));
        var rejected = /^rejected\b/i.test(W);
        out.push({
          key: key, row: r + 1, sol: i, AA: AA, AB: AB, U: U, V: V, W: W, X: X, rejected: rejected,
          // 30.5: "Rejected | Check Avionics" → one position, Rejected; pending → each function
          fns: rejected ? ["Rejected"] : W && W !== "-" ? W.split(/\s*,\s*/).filter(Boolean) : [],
          persons: X && X !== "-" ? X.split(/\s*,\s*/).filter(Boolean) : [],
          rejFn: rejected ? trim(W.replace(/^rejected\s*\|?\s*/i, "")) : ""
        });
      });
    });
    return out;
  }
  function docOf(x) { return x.AA + " Rev " + x.AB; }

  // ---------------------------------------------------------------- the Mantra tables
  function countRows(map) {
    return Object.keys(map).map(function (k) { return [k, map[k].length]; })
      .sort(function (a, b) { return b[1] - a[1] || txtCmp(a[0], b[0]); });
  }
  function addOnce(map, k, v) { if (!map[k]) map[k] = []; if (map[k].indexOf(v) < 0) map[k].push(v); }

  // 100.7 MANTRA SIGNATURE POSITION: document–revision pairs per position (rejected counted under Rejected)
  function mantraSignature(rel) {
    var m = {};
    rel.forEach(function (x) { x.fns.forEach(function (f) { addOnce(m, f, docOf(x)); }); });
    return countRows(m);
  }
  // 100.7 MANTRA WHO: document–revision pairs per person (no person when the rejecter is missing)
  function mantraWho(rel, names) {
    var m = {};
    rel.forEach(function (x) { x.persons.forEach(function (p) { addOnce(m, names.canon(p), docOf(x)); }); });
    return countRows(m);
  }
  // 100.11 PERSON GROUP / DEPARTMENT: person–CDF pairs per group
  function personGroups(rel, names) {
    var m = {};
    rel.forEach(function (x) {
      x.persons.forEach(function (p) {
        var c = names.canon(p), group = names.group(c, x.V);
        if (group !== "-") addOnce(m, group, c + "|" + x.key);
      });
    });
    var rows = GROUPS.map(function (g) { return [g, (m[g] || []).length]; });
    Object.keys(m).sort(txtCmp).forEach(function (g) { if (GROUPS.indexOf(g) < 0) rows.push([g, m[g].length]); });
    return rows;
  }
  // 100.10 REJECTED: one row per CDF + document + revision; the rejecter as Mantra wrote it, or - when missing
  function rejectedRows(rel) {
    var seen = {}, rows = [];
    rel.forEach(function (x) {
      if (!x.rejected) return;
      var k = x.key + "|" + x.AA + "|" + x.AB;
      if (seen[k]) return;
      seen[k] = true;
      rows.push([x.key, "Open", x.AA, x.AB, x.persons.length ? x.X : "-", x.rejFn || "-"]);
    });
    rows.sort(function (a, b) { return keyCmp(a[0], b[0]) || txtCmp(a[2], b[2]) || txtCmp(a[3], b[3]); });
    return rows;
  }
  function uniqCount(list) { var s = {}; list.forEach(function (x) { s[x] = 1; }); return Object.keys(s).length; }

  // 100.9 PERSONS: one row per group + person + position + CDF, its documents in F; # of M per person
  function personRows(rel, names) {
    var rows = {}, order = [];
    rel.forEach(function (x) {
      x.persons.forEach(function (p, pi) {
        var person = names.canon(p), group = mainGroup(names.group(person, x.V));
        // several people and several functions come sorted apart in Mantra, so a function can't be paired
        // with one person; each person then carries all the functions of the document
        var fns = x.fns.length === 1 || x.persons.length === 1 ? x.fns : x.fns.length === x.persons.length ? [x.fns.join(", ")] : x.fns;
        fns.forEach(function (fn) {
          var k = group + "|" + personKey(person) + "|" + fn + "|" + x.key;
          if (!rows[k]) { rows[k] = { group: group, person: person, pos: fn, key: x.key, docs: [] }; order.push(k); }
          if (rows[k].docs.indexOf(docOf(x)) < 0) rows[k].docs.push(docOf(x));
        });
      });
    });
    var list = order.map(function (k) { return rows[k]; });
    list.sort(function (a, b) { return txtCmp(a.group, b.group) || txtCmp(a.person, b.person) || txtCmp(a.pos, b.pos) || keyCmp(a.key, b.key); });
    var shown = {};
    list.forEach(function (r) { var k = r.group + "|" + personKey(r.person); shown[k] = (shown[k] || []).concat(r.docs); });
    list.forEach(function (r) { r.nM = uniqCount(shown[r.group + "|" + personKey(r.person)]); });
    return list;
  }

  // ---------------------------------------------------------------- the release tables, carried over
  var RELEASE = [
    { id: "cdf24", name: T.cdf24, width: 3, end: /^TOTAL$/ },
    { id: "cdf7", name: T.cdf7, width: 3, end: /^TOTAL$/ },
    { id: "doc24", name: T.doc24, width: 6, end: /^ENGINEERING DOCUMENTS$/ },
    { id: "doc7", name: T.doc7, width: 6, end: /^ENGINEERING DOCUMENTS$/ }
  ];
  // the title row of a table: its name alone, or followed by " (AS OF …)"; with the first header under it when given
  function findTitle(values, name, c0, header) {
    var want = dashKey(name);
    for (var r = 0; r < values.length; r++) {
      var t = dashKey(cell(values, r, c0 || 0));
      if (t !== want && t.indexOf(want + " (") !== 0) continue;
      if (header && dashKey(cell(values, r + 1, c0 || 0)) !== dashKey(header)) continue;
      return r;
    }
    return -1;
  }
  // { cdf24: { title, headers, rows }, … } from the sheet as it is now
  function readReleaseTables(values) {
    var out = {};
    RELEASE.forEach(function (t) {
      var r = findTitle(values || [], t.name);
      if (r < 0) return;
      var headers = [], rows = [];
      for (var c = 0; c < t.width; c++) headers.push(trim(cell(values, r + 1, c)));
      for (var i = r + 2; i < values.length; i++) {
        var row = [];
        for (c = 0; c < t.width; c++) row.push(cell(values, i, c) == null ? "" : cell(values, i, c));
        var a = upper(row[0]);
        if (t.end.test(a) || row.every(function (v) { return isEmpty(v); })) break;
        rows.push(row.map(function (v) { return typeof v === "number" ? v : trim(v); }));
      }
      out[t.id] = { title: trim(cell(values, r, 0)), headers: headers, rows: rows, at: r + 1 };
    });
    return out;
  }

  // ---------------------------------------------------------------- the model: every table, in sheet order
  /*
   * input: { cdf (as stages 1–2 leave it), summary (stage 1's summary of that sheet), cfg, algo, mantra (CdfMantra, for groups),
   *          current (CDF SUMMARY values now: the release tables come from it) }
   */
  function build(input) {
    var cdf = input.cdf || [], s1 = input.summary, cfg = input.cfg || {}, algo = input.algo;
    var names = makeNames(cfg, input.mantra, cdf);
    var rel = relations(cdf, algo);
    var t1 = {}; (s1 ? s1.tables : []).forEach(function (t) { t1[t.name] = t; });
    var pendingNames = (s1 && s1.pendingPeople || []).map(function (p) { return p.table; });
    var multiT = false;
    algo.buildBlocks(cdf).forEach(function (b) { if (trim(cell(cdf, b.anchor, C.R)) === "Open" && /[,;\n]/.test(trim(cell(cdf, b.anchor, col("T"))))) multiT = true; });
    var persons = personRows(rel, names), rej = rejectedRows(rel);
    var tables = [];
    function from1(id, name, headers) {
      var t = t1[name];
      if (!t) return null;
      return { id: id, name: name, title: name, headers: headers || t.headers, rows: t.rows.filter(function (r) { return r[0] !== "TOTAL"; }).map(function (r) { return r.slice(); }), total: t.rows.filter(function (r) { return r[0] === "TOTAL"; })[0] };
    }
    var st = from1("status", T.status); if (st) tables.push(st);
    var gc = from1("gcc", T.gcc); if (gc) tables.push(gc);
    var sg = from1("sig", T.sig); if (sg) { sg.multiT = multiT; tables.push(sg); }
    pendingNames.forEach(function (n, i) { var p = from1("pending" + i, n); if (p) { p.side = i + 1; tables.push(p); } });
    var wh = from1("where", T.where); if (wh) tables.push(wh);
    tables.push({ id: "msig", name: T.msig, title: T.msig, headers: ["SIGNATURE POSITION", "DOCUMENTS"], rows: mantraSignature(rel) });
    tables.push({ id: "who", name: T.who, title: T.who, headers: ["PERSON", "DOCUMENTS"], rows: mantraWho(rel, names) });
    tables.push({ id: "groups", name: T.groups, title: T.groups, headers: ["GROUP / DEPARTMENT", "PERSON-CDF"], rows: personGroups(rel, names) });
    tables.push({ id: "rejected", name: T.rejected, title: T.rejected, headers: ["CDF NUMBER", "CDF STATUS", "DOCUMENT", "REVISION", "REJECTED BY", "REJECTED FUNCTION"], rows: rej,
      sums: [["REJECTED CDFS", uniqCount(rej.map(function (r) { return r[0]; }))], ["REJECTED DOCUMENTS", uniqCount(rej.map(function (r) { return r[2] + "|" + r[3]; }))]] });
    tables.push({ id: "persons", name: T.persons, title: T.persons, headers: ["PERSON", "# of M", "MAIN GROUP", "SIGNATURE POSITION", "CDF NUMBERS", "PENDING MANTRA DOCUMENTS"],
      rows: persons.map(function (p) { return [p.person, p.nM, p.group, p.pos, p.key, p.docs.join("\n")]; }), persons: persons });
    // the release tables: computed from BACKUPDATA events (stage 4) when given, else carried over from the sheet as they are
    var rt = input.release || readReleaseTables(input.current);
    RELEASE.forEach(function (d) {
      var t = rt[d.id];
      if (!t) return;
      tables.push({ id: d.id, name: d.name, title: t.title, headers: t.headers, rows: t.rows, carried: !input.release, from: t.at });
    });
    return {
      tables: tables, relations: rel, persons: persons,
      missing: RELEASE.filter(function (d) { return !rt[d.id]; }).map(function (d) { return d.name; }),
      duplicates: (cfg.duplicates && cfg.duplicates.length ? cfg.duplicates : [DUPLICATE_RAW]),
      missingGroups: persons.filter(function (p) { return p.group === "-"; }).map(function (p) { return p.person; }).filter(function (x, i, a) { return a.indexOf(x) === i; })
    };
  }

  // ---------------------------------------------------------------- layout: where each cell goes
  // cell: { r (1-based), c (0-based), v (value or "=formula"), value (what it shows), role: title|head|text|num|center|time }
  function q(s) { return '"' + str(s).replace(/"/g, '""') + '"'; }
  function kindOf(t, j) {
    if (t.id === "rejected" || /^cdf(24|7)$/.test(t.id)) return "text";
    if (/^doc(24|7)$/.test(t.id)) return j === 5 ? "time" : "text";
    if (t.id === "persons") return j === 1 ? "num" : j === 2 ? "center" : "text";
    if (/^pending/.test(t.id)) return "text";
    return j === 1 ? "num" : "text";
  }
  // the CDF sheet references are the macro's; refs.open / refs.station3 are the label cells other formulas point at
  function rowFormula(t, row, r, refs, model) {
    var a = "A" + r, cdfC = "CDF!$C:$C,\"<>\"";
    if (t.id === "status") {
      var f = "=COUNTIFS(CDF!$R:$R," + a + "," + cdfC + ")";
      if (row[0] === "Released") model.duplicates.forEach(function (d) { f += "-COUNTIFS(CDF!$C:$C," + q(d) + ",CDF!$R:$R," + a + ")"; });
      return f;
    }
    if (t.id === "gcc") return "=COUNTIFS(CDF!$R:$R," + refs.open + ",CDF!$S:$S," + a + "," + cdfC + ")";
    if (t.id === "sig") return "=COUNTIFS(CDF!$T:$T," + (t.multiT ? "\"*\"&" + a + "&\"*\"" : a) + "," + cdfC + ",CDF!$R:$R," + refs.open + ")";
    if (t.id === "where") {
      var al = refs.aliases[row[0]] || [row[0]];
      var base = "COUNTIFS(CDF!$R:$R," + refs.open + ",CDF!$S:$S," + refs.station3 + "," + cdfC + ",CDF!$P:$P,";
      return al.length > 1 ? "=SUM(" + base + "{" + al.map(q).join(",") + "}))" : "=" + base + q(al[0]) + ")";
    }
    return null;
  }
  function place(t, r0, c0, refs, model) {
    var w = t.headers.length, cells = [], merges = [];
    for (var j = 0; j < w; j++) cells.push({ r: r0, c: c0 + j, v: j ? "" : t.title, value: j ? "" : t.title, role: "title" });
    if (w > 1) merges.push({ r1: r0, c1: c0, r2: r0, c2: c0 + w - 1, title: true });
    t.headers.forEach(function (h, j) { cells.push({ r: r0 + 1, c: c0 + j, v: h, value: h, role: "head" }); });
    var first = r0 + 2, r = first;
    if (t.id === "status") refs.open = "$A$" + (first + t.rows.map(function (x) { return x[0]; }).indexOf("Open"));
    if (t.id === "gcc") refs.station3 = "$A$" + (first + t.rows.map(function (x) { return x[0]; }).indexOf("3 - Conversion Engineer"));
    t.rows.forEach(function (row) {
      row.forEach(function (v, j) {
        var f = j === 1 && refs.formulas ? rowFormula(t, row, r, refs, model) : null;
        cells.push({ r: r, c: c0 + j, v: f || v, value: v, role: kindOf(t, j) });
      });
      r++;
    });
    var last = r - 1, n = t.rows.length, L = letter(c0), L2 = letter(c0 + 1);
    function line(label, valueCol, value, formula) {
      for (var j = 0; j < w; j++) cells.push({ r: r, c: c0 + j, v: j === 0 ? label : j === valueCol ? (formula && refs.formulas ? formula : value) : "", value: j === 0 ? label : j === valueCol ? value : "", role: j === valueCol ? "num" : "text" });
      r++;
    }
    var colOf = function (j) { return letter(c0 + j); };
    // TOTAL and sums under the table — formulas over the rows above, as the macro writes them
    if (/^(status|gcc|sig|where|msig|who|groups)$/.test(t.id)) {
      var tot = t.rows.reduce(function (s, x) { return s + (Number(x[1]) || 0); }, 0);
      line("TOTAL", 1, tot, n ? "=SUM(" + L2 + first + ":" + L2 + last + ")" : null);
    } else if (/^pending/.test(t.id)) line("TOTAL", 1, n, n ? "=COUNTA(" + L + first + ":" + L + last + ")" : null);
    else if (t.id === "rejected") {
      var rng = function (cl) { return cl + first + ":" + cl + last; };
      line(t.sums[0][0], 5, t.sums[0][1], n ? "=SUMPRODUCT(1/COUNTIF(" + rng("A") + "," + rng("A") + "))" : null);
      line(t.sums[1][0], 5, t.sums[1][1], n ? "=SUMPRODUCT(1/COUNTIFS(" + rng("C") + "," + rng("C") + "," + rng("D") + "," + rng("D") + "))" : null);
    } else if (/^cdf(24|7)$/.test(t.id)) line("TOTAL", 2, uniqCount(t.rows.map(function (x) { return trim(x[0]); })), n ? "=COUNTA(A" + first + ":A" + last + ")" : null);
    else if (/^doc(24|7)$/.test(t.id)) {
      var eng = t.rows.filter(function (x) { return upper(x[0]) !== "ANSWER AT GCC"; }).length, ans = n - eng, re = r;
      line("ENGINEERING DOCUMENTS", 5, eng, n ? "=COUNTIF(A" + first + ":A" + last + ",\"<>ANSWER AT GCC\")" : null);
      line("GCC ANSWERS", 5, ans, n ? "=COUNTIF(A" + first + ":A" + last + ",\"ANSWER AT GCC\")" : null);
      line("OF WHICH LINKED TO CDFS", 5, eng, "=F" + re);
      line("TOTAL", 5, eng + ans, "=F" + re + "+F" + (re + 1));
    }
    // 100.9.5: one person's rows merged in A (person), B (# of M), C (group)
    if (t.id === "persons") {
      var i = 0;
      while (i < t.persons.length) {
        var k = i;
        while (k + 1 < t.persons.length && t.persons[k + 1].group === t.persons[i].group && personKey(t.persons[k + 1].person) === personKey(t.persons[i].person)) k++;
        if (k > i) [0, 1, 2].forEach(function (j) { merges.push({ r1: first + i, c1: c0 + j, r2: first + k, c2: c0 + j }); });
        i = k + 1;
      }
    }
    return { id: t.id, name: t.name, cells: cells, merges: merges, r1: r0, r2: r - 1, c1: c0, c2: c0 + w - 1, first: first, last: last, colOf: colOf };
  }
  // CDF SUMMARY: tables one under the other, two empty rows between; GCC SIGNATURE POSITION, MOR and IGOR side by side
  function layoutSummary(model, opts) {
    opts = opts || {};
    var refs = { formulas: opts.formulas !== false, open: "$A$3", station3: "$A$13", aliases: opts.aliases || {} };
    var placed = [], r = 1;
    var side = model.tables.filter(function (t) { return /^pending/.test(t.id); });
    model.tables.forEach(function (t) {
      if (/^pending/.test(t.id)) return;
      if (t.id === "sig") {
        var group = [place(t, r, 0, refs, model)].concat(side.map(function (p) { return place(p, r, 3 * p.side, refs, model); }));
        placed = placed.concat(group);
        r = Math.max.apply(null, group.map(function (g) { return g.r2; })) + 3;
        return;
      }
      var p = place(t, r, 0, refs, model);
      placed.push(p);
      r = p.r2 + 3;
    });
    return finish(placed, r - 3);
  }
  // EMAIL DRAFT (150.8): subject in A1, the opening in A3, the six tables from A5, two empty rows between, the signature
  var EMAIL_TABLES = ["status", "gcc", "where", "persons", "cdf24", "cdf7"];
  function layoutEmail(model, mail) {
    var refs = { formulas: false, aliases: {} }, placed = [], r = 5;
    EMAIL_TABLES.forEach(function (id) {
      var t = model.tables.filter(function (x) { return x.id === id; })[0];
      if (!t) return;
      var p = place(t, r, 0, refs, model);
      placed.push(p);
      r = p.r2 + 3;
    });
    var out = finish(placed, r - 3);
    out.cells.push({ r: 1, c: 0, v: mail.subject, value: mail.subject, role: "subject" });
    out.cells.push({ r: 3, c: 0, v: mail.opening, value: mail.opening, role: "line" });
    var sr = out.end + 2;
    (mail.signature || []).forEach(function (s, i) { out.cells.push({ r: sr + i, c: 0, v: s, value: s, role: "line" }); });
    out.end = sr + Math.max((mail.signature || []).length, 1) - 1;
    return out;
  }
  // a merged area keeps only its first cell (as Excel does); the rest are written empty
  function finish(placed, end) {
    var cells = [], merges = [];
    placed.forEach(function (p) {
      var covered = {};
      p.merges.forEach(function (m) {
        if (m.title) return;
        for (var r = m.r1; r <= m.r2; r++) for (var c = m.c1; c <= m.c2; c++) if (r !== m.r1 || c !== m.c1) covered[r + ":" + c] = true;
      });
      p.cells.forEach(function (c) { if (covered[c.r + ":" + c.c]) { c.v = ""; c.value = ""; c.covered = true; } });
      cells = cells.concat(p.cells); merges = merges.concat(p.merges);
    });
    return { tables: placed, cells: cells, merges: merges, end: end };
  }
  // the supplier aliases for WHERE ARE THE CDFS: display name → the values of CDF!P that count for it (rules.txt order)
  function supplierAliases(set) {
    var out = {};
    Object.keys(set.suppliers || {}).forEach(function (p) { var d = set.suppliers[p]; (out[d] = out[d] || []).push(p); });
    return out;
  }

  // ---------------------------------------------------------------- 150: the email
  var JLM = "Asia/Jerusalem";
  function jlm(date) {                                  // "05/11/2026", "08:10" in Israel time, whatever the PC's zone
    var p = {};
    try {
      new Intl.DateTimeFormat("en-GB", { timeZone: JLM, year: "numeric", month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit", hour12: false })
        .formatToParts(date).forEach(function (x) { p[x.type] = x.value; });
    } catch (e) { p = {}; }
    if (!p.year) {
      var pad = function (n) { return (n < 10 ? "0" : "") + n; };
      return { date: pad(date.getDate()) + "/" + pad(date.getMonth() + 1) + "/" + date.getFullYear(), time: pad(date.getHours()) + ":" + pad(date.getMinutes()) };
    }
    return { date: p.day + "/" + p.month + "/" + p.year, time: (p.hour === "24" ? "00" : p.hour) + ":" + p.minute };
  }
  function mailParts(now, signature) {
    var t = jlm(now);
    return {
      subject: "A330-300 P2F — Daily CDF Status — " + t.date,
      opening: "בוקר טוב, להלן סטטוס החריגים המעודכן — AS OF " + t.date + " " + t.time + " (" + JLM + ").",
      signature: ["בברכה,"].concat((signature || []).filter(function (s) { return trim(s) && trim(s) !== "בברכה,"; })),
      date: t.date, time: t.time
    };
  }
  // the lines after "בברכה," in the EMAIL DRAFT sheet now — the sender's name stays the user's own
  function readSignature(values) {
    values = values || [];
    for (var r = values.length - 1; r >= 0; r--) {
      if (trim(cell(values, r, 0)) === "בברכה,") {
        var out = [];
        for (var i = r + 1; i < values.length && i < r + 4; i++) if (trim(cell(values, i, 0))) out.push(trim(cell(values, i, 0)));
        return out;
      }
    }
    return [];
  }
  // the six tables for E.buildDailyEmail: { title, headers, rows: [[{v, rowspan}|null]] }, merged person cells as rowspan
  function emailDoc(model, mail) {
    var lay = layoutEmail(model, mail), tables = [];
    lay.tables.forEach(function (p) {
      var w = p.c2 - p.c1 + 1, grid = {};
      p.cells.forEach(function (c) { grid[c.r + ":" + c.c] = c; });
      var covered = {}, span = {};
      p.merges.forEach(function (m) {
        if (m.title) return;
        span[m.r1 + ":" + m.c1] = m.r2 - m.r1 + 1;
        for (var rr = m.r1 + 1; rr <= m.r2; rr++) covered[rr + ":" + m.c1] = true;
      });
      var rows = [];
      for (var r = p.r1 + 2; r <= p.r2; r++) {
        var row = [];
        for (var c = p.c1; c < p.c1 + w; c++) {
          if (covered[r + ":" + c]) { row.push(null); continue; }
          var x = grid[r + ":" + c];
          row.push({ v: x ? str(x.value) : "", rowspan: span[r + ":" + c] || 1 });
        }
        rows.push(row);
      }
      var t = model.tables.filter(function (x) { return x.id === p.id; })[0];
      tables.push({ title: t.title, headers: t.headers, rows: rows, widths: p.id === "persons" ? [130, 50, 70, 120, 100, 200] : null });
    });
    return { daily: true, subject: mail.subject, opening: mail.opening, signature: mail.signature, tables: tables };
  }
  // 150.6 an .eml draft: multipart/alternative, UTF-8, subject encoded, no attachments
  function b64(s) {
    var bytes = unescape(encodeURIComponent(s));
    if (typeof btoa === "function") return btoa(bytes);
    return Buffer.from(s, "utf8").toString("base64");
  }
  function eml(mail, html, text, to, cc) {
    var boundary = "=_cdf_" + Math.random().toString(36).slice(2);
    var wrap = function (s) { return b64(s).replace(/(.{76})/g, "$1\r\n"); };
    return [
      "X-Unsent: 1",
      "To: " + (to || []).join(", "),
      cc && cc.length ? "Cc: " + cc.join(", ") : null,
      "Subject: =?UTF-8?B?" + b64(mail.subject) + "?=",
      "MIME-Version: 1.0",
      "Content-Type: multipart/alternative; boundary=\"" + boundary + "\"",
      "",
      "--" + boundary,
      "Content-Type: text/plain; charset=UTF-8",
      "Content-Transfer-Encoding: base64",
      "",
      wrap(text),
      "--" + boundary,
      "Content-Type: text/html; charset=UTF-8",
      "Content-Transfer-Encoding: base64",
      "",
      wrap("<!doctype html><html><head><meta charset=\"utf-8\"><title>" + str(mail.subject).replace(/[<&]/g, "") + "</title></head><body>" + html + "</body></html>"),
      "--" + boundary + "--",
      ""
    ].filter(function (x) { return x !== null; }).join("\r\n");
  }

  // ---------------------------------------------------------------- compare with the sheet now
  // each table found by its title; rows compared as text (merged cells filled down); → [{ name, found, added, removed, changed }]
  function readTableAt(values, r, width, kind) {
    var rows = [], fill = [];
    for (var i = r + 2; i < values.length; i++) {
      var row = [];
      for (var c = 0; c < width; c++) { var v = cell(values, i, c); row.push(typeof v === "number" ? String(Math.round(v * 1000) / 1000) : trim(v)); }
      if (row.every(function (x) { return x === ""; })) break;
      if (kind === "persons") row = row.map(function (x, j) { if (j < 3 && x === "" && fill[j] != null) return fill[j]; fill[j] = x; return x; });
      rows.push(row);
    }
    return rows;
  }
  function compare(lay, values) {
    values = values || [];
    return lay.tables.map(function (p) {
      var w = p.c2 - p.c1 + 1, want = [];
      var grid = {};
      p.cells.forEach(function (c) { grid[c.r + ":" + c.c] = c; });
      var fill = [];
      for (var r = p.r1 + 2; r <= p.r2; r++) {
        var row = [];
        for (var c = p.c1; c <= p.c2; c++) { var x = grid[r + ":" + c]; row.push(x ? (typeof x.value === "number" ? String(x.value) : trim(x.value)) : ""); }
        if (p.id === "persons") row = row.map(function (v, j) { if (j < 3 && v === "" && fill[j] != null) return fill[j]; fill[j] = v; return v; });
        want.push(row);
      }
      var name = p.name, sheetCol = p.c1, head = grid[(p.r1 + 1) + ":" + p.c1];
      var at = findTitle(values, name, sheetCol, head ? head.value : null);
      if (at < 0) return { id: p.id, name: name, found: false, added: want.length, removed: 0, changed: [], rowsNow: 0 };
      var shifted = values.map(function (row) { return (row || []).slice(sheetCol, sheetCol + w); });
      var have = readTableAt(shifted, at, w, p.id);
      var key = function (row) { return row.join("\u0001"); };
      var hk = {}, wk = {};
      have.forEach(function (x) { hk[key(x)] = (hk[key(x)] || 0) + 1; });
      want.forEach(function (x) { wk[key(x)] = (wk[key(x)] || 0) + 1; });
      var added = [], removed = [];
      want.forEach(function (x) { var k = key(x); if (hk[k]) hk[k]--; else added.push(x); });
      have.forEach(function (x) { var k = key(x); if (wk[k]) wk[k]--; else removed.push(x); });
      // a label row whose number changed reads better as one change
      var changed = [];
      if (w === 2 || p.id === "rejected" || /^doc/.test(p.id)) {
        added.slice().forEach(function (a) {
          var i = -1;
          removed.forEach(function (rm, j) { if (i < 0 && rm[0] === a[0]) i = j; });
          if (i >= 0) { changed.push({ label: a[0], now: removed[i].slice(1).join(" · "), will: a.slice(1).join(" · ") }); removed.splice(i, 1); added.splice(added.indexOf(a), 1); }
        });
      }
      return { id: p.id, name: name, found: true, at: letter(sheetCol) + (at + 1), added: added, removed: removed, changed: changed, moved: at + 1 !== p.r1, rowsNow: have.length };
    });
  }
  function differs(c) { return !c.found || (c.added && c.added.length) || (c.removed && c.removed.length) || (c.changed && c.changed.length); }

  // ---------------------------------------------------------------- 100.12: what the sheet shows after writing
  // values read back from the written tables vs the model; a formula cell that shows "" wasn't calculated (left out)
  function readBack(lay, valuesAt) {
    var bad = [];
    lay.cells.forEach(function (c) {
      if (typeof c.value !== "number") return;
      var got = valuesAt(c.r, c.c);
      if (got === "" || got == null) return;
      if (Math.abs(Number(got) - c.value) > 1e-9) bad.push({ cell: letter(c.c) + c.r, want: c.value, got: got });
    });
    return bad;
  }

  return {
    T: T, GROUPS: GROUPS, WIDTHS: WIDTHS, EMAIL_TABLES: EMAIL_TABLES, RELEASE: RELEASE,
    keyCmp: keyCmp, keyParts: keyParts, personKey: personKey, makeNames: makeNames, titleCase: titleCase, mainGroup: mainGroup,
    relations: relations, mantraSignature: mantraSignature, mantraWho: mantraWho, personGroups: personGroups,
    rejectedRows: rejectedRows, personRows: personRows, readReleaseTables: readReleaseTables,
    build: build, layoutSummary: layoutSummary, layoutEmail: layoutEmail, supplierAliases: supplierAliases,
    jlm: jlm, mailParts: mailParts, readSignature: readSignature, emailDoc: emailDoc, eml: eml,
    compare: compare, differs: differs, readBack: readBack, letter: letter, col: col
  };
});
