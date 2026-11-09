/*
 * CDF Updater - CDF algorithm (code version of CDF_Algorithm_2026-11-05)
 *
 * STAGE 1 - GCC -> CDF
 *   00.4-00.7  key / empty R / Cancelled rules        10    GCC Final Approval + the FINAL_APPROVAL_RELEASE signer -> Released
 *   05.3/20.1  key normalisation and matching          20.2  import J,AM,AN,P,Q -> R,S,T,Y,Z
 *   20.3       explicit station mapping (+ Completed)  20.3.2 fixed correction A330-1-052 (another)
 *   100.3-100.5, 100.8  CDF SUMMARY tables that come from GCC (computed here; cdf-summary.js writes them in stage 3)
 *   130.6      control totals
 *
 * Pure logic, no Office.js: works in the task pane (window.CdfAlgo) and in Node (tests / replay).
 * Input sheets are 2D arrays of values starting at A1 (row 0 = Excel row 1).
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.CdfAlgo = factory();
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  var STAGE = 1;          // this file: stage 1. STAGES_DONE: how far the add-in goes (2: cdf-mantra.js, 3: cdf-summary.js, 4: cdf-history.js)
  var STAGES_DONE = 4;
  // the algorithm text this code implements; rules.txt states the version it expects (CDF_ALGORITHM "...")
  var ALGORITHM_VERSION = "2026-11-05";
  var STAGES = [
    { n: 1, name: "GCC → CDF (R, S, T, Y, Z) and the GCC tables of CDF SUMMARY", done: true },
    { n: 2, name: "Mantra → CDF (AA, AB, U:X), solution rows, closure date", done: true },
    { n: 3, name: "CDF SUMMARY (all tables) and EMAIL DRAFT", done: true },
    { n: 4, name: "BACKUPDATA history, releases in 24 h / 7 days, ERROR", done: true }
  ];

  function col(letter) {
    var n = 0;
    for (var i = 0; i < letter.length; i++) n = n * 26 + (letter.charCodeAt(i) - 64);
    return n - 1;
  }
  function letter(idx) {
    var n = idx + 1, s = "";
    while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }
  function str(v) { return v === null || v === undefined ? "" : String(v); }
  function trim(v) { return str(v).replace(/ /g, " ").trim(); }
  function isEmpty(v) { return trim(v) === ""; }
  function upper(v) { return trim(v).replace(/\s+/g, " ").toUpperCase(); }

  // ---------- column maps (00.1 and 20.2) ----------
  var CDF = { A: col("A"), C: col("C"), F: col("F"), H: col("H"), J: col("J"), K: col("K"), P: col("P"), Q: col("Q"), R: col("R"), S: col("S"), T: col("T"), U: col("U"),
    V: col("V"), W: col("W"), X: col("X"), Y: col("Y"), Z: col("Z"), AA: col("AA"), AB: col("AB") };
  var GCC = { A: col("A"), J: col("J"), P: col("P"), Q: col("Q"), S: col("S"), AK: col("AK"), AM: col("AM"), AN: col("AN") };

  // expected headers (00.1: verify before writing)
  var CDF_HEADERS = { C: "LINK TO GCC", R: "CDF STATUS", S: "CURRENT STATION", T: "SIGNTURE POSITION", U: "STATUS (MANTRA)",
    Y: "DEVIATION REASON", Z: "MINOR / MAJOR", AA: "CDF SOLUTION" };
  var GCC_HEADERS = { A: "CDF NUMBER", J: "CDF STATUS", P: "DEVIATION REASON", Q: "DEVIATION CLASSIFICATION",
    S: "ENGINEERING DISPOSITION", AK: "FINAL APPROVAL NAME", AM: "CURRENT STATION", AN: "CURRENT ASSIGNEE" };

  // 20.3 explicit station mapping (GCC!AM -> CDF!S)
  var STATIONS = {
    "INITIATOR": "1 - Initiator", "1 - INITIATOR": "1 - Initiator",
    "POC": "2 - POC", "POC ENG": "2 - POC", "2 - POC": "2 - POC",
    "CONVERSION ENG": "3 - Conversion Engineer", "CONVERSION ENGINEER": "3 - Conversion Engineer", "3 - CONVERSION ENGINEER": "3 - Conversion Engineer",
    "AVIONICS": "4 - AVIONICS", "4 - AVIONICS": "4 - AVIONICS",
    "FINAL APPROVAL": "9 - Final Approval", "9 - FINAL APPROVAL": "9 - Final Approval",
    "PRODUCTION ENGINEERING": "10 - Production Engineering", "10 - PRODUCTION ENGINEERING": "10 - Production Engineering",
    "PRODUCTION PLANNING": "11 - Production Planning", "11 - PRODUCTION PLANNING": "11 - Production Planning",
    "RELEASED": "Released"
  };
  var OPEN_STATIONS = ["1 - Initiator", "2 - POC", "3 - Conversion Engineer", "4 - AVIONICS", "9 - Final Approval"];
  var R_VALUES = { "OPEN": "Open", "CANCELLED": "Cancelled", "CANCELED": "Cancelled", "RELEASED": "Released" };

  var DUPLICATE_KEY = "A330-1-052(ANOTHER)";   // 20.3.2 / 100.3
  // people come from the settings (rules.txt / workbook settings), not from the code

  // 100.8 supplier display
  var SUPPLIERS = { "ANN/SIE": "ANN/SIE", "SIE/ANN": "ANN/SIE", "ANN": "ANN/SIE", "ANCRA": "ANCARA", "ANCARA": "ANCARA", "TIMAT": "TIMAT", "IAI": "IAI" };
  var SUPPLIER_ORDER = ["ANN/SIE", "ANCARA", "TIMAT", "IAI"];

  // ---------- keys (05.3 / 20.1) ----------
  function keyFull(v) { return str(v).replace(/[\s ]+/g, "").toUpperCase(); }
  function keyBase(v) { return keyFull(v).replace(/\.-$/, "").replace(/\.$/, ""); }

  function cell(rows, r, c) { var row = rows[r]; return row && c < row.length ? row[c] : ""; }

  // ---------- header check ----------
  function checkHeaders(rows, expected, sheet) {
    var errs = [];
    Object.keys(expected).forEach(function (k) {
      var h = upper(cell(rows, 0, col(k)));
      if (h.indexOf(expected[k]) < 0) errs.push(sheet + "!" + k + "1 is \"" + trim(cell(rows, 0, col(k))) + "\" — expected \"" + expected[k] + "\"");
    });
    return errs;
  }

  // ---------- CDF block map (07.2 / 100.2) ----------
  // anchor = row with a key in C; following rows up to the next anchor belong to its block
  function buildBlocks(rows) {
    var blocks = [], cur = null;
    for (var r = 1; r < rows.length; r++) {
      var key = trim(cell(rows, r, CDF.C));
      if (key) { cur = { key: key, anchor: r, rows: [r] }; blocks.push(cur); continue; }
      if (!cur) continue;
      var row = rows[r] || [];
      var hasData = [CDF.R, CDF.U, CDF.AA, CDF.AB].some(function (c) { return !isEmpty(row[c]); });
      if (hasData) cur.rows.push(r);
    }
    return blocks;
  }

  // settings from rules.txt (STATION, SUPPLIER, FINAL_APPROVAL_RELEASE, PENDING_TABLE, DUPLICATE_CDF)
  function settings(cfg) {
    cfg = cfg || {};
    var st = {}, sp = {};
    Object.keys(STATIONS).forEach(function (k) { st[k] = STATIONS[k]; });
    Object.keys(SUPPLIERS).forEach(function (k) { sp[k] = SUPPLIERS[k]; });
    Object.keys(cfg.stations || {}).forEach(function (k) { st[upper(k)] = cfg.stations[k]; });
    Object.keys(cfg.suppliers || {}).forEach(function (k) { sp[upper(k)] = cfg.suppliers[k]; });
    var order = [];
    Object.keys(sp).forEach(function (k) { if (order.indexOf(sp[k]) < 0) order.push(sp[k]); });
    var pend = cfg.pendingTables && cfg.pendingTables.length ? cfg.pendingTables : [];
    var dups = (cfg.duplicates && cfg.duplicates.length ? cfg.duplicates : [DUPLICATE_KEY]).map(keyFull);
    return { stations: st, suppliers: sp, supplierOrder: order, signer: cfg.finalApprovalSigner || null, pending: pend, duplicates: dups };
  }

  // ---------- main ----------
  /*
   * input: { cdf: [[...]], gcc: [[...]], now: Date }
   * output: { stage, errors[], cdfChanges[], gccChanges[], summary{}, checks[], info{} }
   */
  function run(input, cfg) {
    var cdf = input.cdf || [], gcc = input.gcc || [];
    var set = settings(cfg);
    function isDup(k) { return set.duplicates.indexOf(keyFull(k)) >= 0; }
    var out = { stage: STAGE, errors: [], warnings: [], cdfChanges: [], gccChanges: [], summary: null, checks: [], info: {} };

    // 00.1: verify headers before writing
    var hErr = checkHeaders(cdf, CDF_HEADERS, "CDF").concat(checkHeaders(gcc, GCC_HEADERS, "GCC"));
    if (hErr.length) { out.errors = hErr.map(function (m) { return { stage: "00.1", message: m }; }); return out; }

    var blocks = buildBlocks(cdf);

    // --- 10: GCC Final Approval + the FINAL_APPROVAL_RELEASE signer -> Released (writes GCC!AM) ---
    var gccRows = [];
    for (var g = 1; g < gcc.length; g++) {
      var gk = trim(cell(gcc, g, GCC.A));
      var am = trim(cell(gcc, g, GCC.AM));
      if (gk && set.signer && upper(am) === "FINAL APPROVAL" && upper(cell(gcc, g, GCC.AK)) === upper(set.signer)) {
        out.gccChanges.push({ sheet: "GCC", row: g + 1, col: GCC.AM, colLetter: "AM", key: gk, oldValue: am, newValue: "Released",
          rule: "10", why: "Final Approval signed by " + set.signer });
        am = "Released";
      }
      if (gk) gccRows.push({ r: g, key: gk, full: keyFull(gk), base: keyBase(gk), am: am });
    }

    // --- 20.1: match CDF!C <-> GCC!A ---
    var gFull = {}, gBase = {}, dupG = [];
    gccRows.forEach(function (x) {
      if (gFull[x.full]) dupG.push(x.key);
      gFull[x.full] = x;
      (gBase[x.base] = gBase[x.base] || []).push(x);
    });
    var cBaseCount = {};
    blocks.forEach(function (b) { var kb = keyBase(b.key); cBaseCount[kb] = (cBaseCount[kb] || 0) + 1; });
    var usedG = {};
    blocks.forEach(function (b) {
      var full = keyFull(b.key), base = keyBase(b.key);
      var m = gFull[full];
      if (!m && gBase[base] && gBase[base].length === 1 && cBaseCount[base] === 1) m = gBase[base][0];   // 20.1(2): only a unique match
      if (m) { b.gcc = m; usedG[m.full] = true; }
    });
    var onlyGcc = gccRows.filter(function (x) { return !usedG[x.full]; }).map(function (x) { return x.key; });
    var onlyCdf = blocks.filter(function (b) { return !b.gcc && !isDup(b.key); }).map(function (b) { return b.key; });
    if (dupG.length) out.errors.push({ stage: "20.1", message: "Duplicate CDF numbers in GCC: " + dupG.join(", ") });

    // --- 20.2 / 20.3 / 00.5 / 00.6: compute the anchor values ---
    function change(b, c, newValue, rule, why, src) {
      var old = cell(cdf, b.anchor, c);
      if (trim(old) === trim(newValue)) return;
      out.cdfChanges.push({ sheet: "CDF", row: b.anchor + 1, col: c, colLetter: letter(c), key: b.key,
        oldValue: old, newValue: newValue, rule: rule, why: why || "", source: src || "" });
    }
    function dash(v) { return isEmpty(v) ? "-" : trim(v); }

    blocks.forEach(function (b) {
      var R, S, T, Y, Z, src;
      if (isDup(b.key)) {                                    // 20.3.2 fixed correction
        R = trim(cell(cdf, b.anchor, CDF.R));
        if (R !== "Cancelled") {
          change(b, CDF.R, "Released", "20.3.2", "fixed correction (duplicate of A330-1-052)");
          change(b, CDF.S, "Released", "20.3.2", "fixed correction (duplicate of A330-1-052)");
        }
        b.R = R === "Cancelled" ? "Cancelled" : "Released";
        return;
      }
      if (!b.gcc) { b.R = trim(cell(cdf, b.anchor, CDF.R)); return; }   // no GCC row: nothing imported (00.5)

      var gr = b.gcc.r;
      src = "GCC row " + (gr + 1);
      var j = upper(cell(gcc, gr, GCC.J));
      R = R_VALUES[j];
      if (!R) {
        out.errors.push({ stage: "20.2", key: b.key, message: "GCC!J" + (gr + 1) + " has an unknown CDF status \"" + trim(cell(gcc, gr, GCC.J)) + "\"" });
        b.R = trim(cell(cdf, b.anchor, CDF.R));
        return;
      }
      b.R = R;
      change(b, CDF.R, R, "20.2", "GCC!J", src);

      if (R === "Cancelled") {                               // 00.6: clear S:AB
        ["S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB"].forEach(function (k) {
          change(b, CDF[k], "", "00.6", "Cancelled — S:AB cleared", src);
        });
        if (b.rows.length > 1) out.warnings.push({ stage: "00.6", key: b.key, message: "Cancelled CDF has " + (b.rows.length - 1) + " solution row(s) — left for stage 2" });
        return;
      }

      // S (20.2 AM -> S, 20.3 explicit mapping, 20.3.1 Completed)
      var am = b.gcc.am, amU = upper(am);
      if (isEmpty(am)) S = "-";
      else if (/(^|\s)COMPLETED$/.test(amU)) S = "Released";
      else S = set.stations[amU];
      if (S === undefined) {
        out.errors.push({ stage: "20.3", key: b.key, message: "GCC!AM" + (gr + 1) + " station \"" + am + "\" has no mapping" });
      } else {
        change(b, CDF.S, S, /COMPLETED$/.test(amU) ? "20.3.1" : "20.2", /COMPLETED$/.test(amU) ? "Completed → Released" : "GCC!AM \"" + am + "\"", src);
      }
      b.S = S === undefined ? trim(cell(cdf, b.anchor, CDF.S)) : S;

      T = dash(trim(cell(gcc, gr, GCC.AN)).replace(/^'+/, ""));
      Y = dash(cell(gcc, gr, GCC.P));
      Z = dash(cell(gcc, gr, GCC.Q));
      change(b, CDF.T, T, "20.2", "GCC!AN", src);
      change(b, CDF.Y, Y, "20.2", "GCC!P", src);
      change(b, CDF.Z, Z, "20.2", "GCC!Q", src);
      b.T = T;
    });

    // blocks with empty R (00.5): U:X must be empty - reported only in stage 1
    blocks.forEach(function (b) {
      if (!isEmpty(b.R)) return;
      var hasUX = b.rows.some(function (r) { return [CDF.U, CDF.V, CDF.W, CDF.X].some(function (c) { return !isEmpty(cell(cdf, r, c)); }); });
      if (hasUX) out.warnings.push({ stage: "00.5", key: b.key, message: "R is empty but U:X has values — the Mantra step empties U:X (00.5)" });
    });

    out.summary = buildSummary(blocks, cdf, set, isDup);
    out.checks = controlChecks(out.summary, blocks, gccRows, onlyGcc);
    out.info = {
      cdfBlocks: blocks.length, gccRows: gccRows.length,
      matched: blocks.filter(function (b) { return b.gcc; }).length,
      onlyGcc: onlyGcc, onlyCdf: onlyCdf
    };
    return out;
  }

  // ---------- 100: summary tables that come from GCC ----------
  function buildSummary(blocks, cdf, set, isDup) {
    var open = blocks.filter(function (b) { return b.R === "Open"; });

    // 100.3 CDF STATUS (duplicate excluded from Released and TOTAL)
    var st = { Open: 0, Cancelled: 0, Released: 0 };
    blocks.forEach(function (b) {
      if (isDup(b.key)) return;
      if (st.hasOwnProperty(b.R)) st[b.R]++;
    });
    var cdfStatus = [["Open", st.Open], ["Cancelled", st.Cancelled], ["Released", st.Released], ["TOTAL", st.Open + st.Cancelled + st.Released]];

    // 100.4 GCC CURRENT STATUS (Open only)
    var stations = OPEN_STATIONS.map(function (s) { return [s, open.filter(function (b) { return b.S === s; }).length]; });
    var otherStations = open.filter(function (b) { return OPEN_STATIONS.indexOf(b.S) < 0; }).map(function (b) { return b.key + " (" + (b.S || "empty") + ")"; });
    stations.push(["TOTAL", open.length]);

    // 100.5 GCC SIGNATURE POSITION (Open only, by name occurrence)
    var sig = {};
    open.forEach(function (b) {
      splitNames(b.T).forEach(function (n) { sig[n] = (sig[n] || 0) + 1; });
    });
    var sigRows = Object.keys(sig).sort(function (a, b) { return sig[b] - sig[a] || a.localeCompare(b); }).map(function (n) { return [n, sig[n]]; });
    var sigTotal = sigRows.reduce(function (s, x) { return s + x[1]; }, 0);
    sigRows.push(["TOTAL", sigTotal]);

    // 100.5.1 / 100.5.2 MOR / IGOR PENDING TABLES
    // the solution rows of a CDF: Mantra status (U), document (AA) and revision (AB) of each
    function docsOf(b) {
      var docs = [], us = [], answer = false, validDoc = false, lines = [];
      b.rows.forEach(function (r) {
        var u = upper(cell(cdf, r, CDF.U)), aa = trim(cell(cdf, r, CDF.AA)), ab = trim(cell(cdf, r, CDF.AB));
        lines.push("row " + (r + 1) + ": " + (aa || "(no AA)") + (aa && aa !== "-" ? " Rev " + (ab || "-") : "") + " — U " + (u || "(empty)"));
        if (u === "ANSWER AT GCC" || upper(aa) === "ANSWER AT GCC") { answer = true; return; }
        if (u && u !== "-") us.push(u);
        if (aa && aa !== "-" && upper(aa) !== "CHECK SOLUTION" && u !== "NO ENG. DOC.") {
          validDoc = true;
          var d = aa + " Rev " + (ab || "-");
          if (docs.indexOf(d) < 0) docs.push(d);
        }
      });
      var allReleased = us.length > 0 && us.every(function (u) { return /RELEASED$/.test(u); });
      return { docs: docs, us: us, answer: answer, validDoc: validDoc, allReleased: allReleased, lines: lines,
        listed: (allReleased && validDoc) || answer };
    }
    function pending(name) {
      var rows = [];
      open.forEach(function (b) {
        if (b.S !== "9 - Final Approval" || splitNames(b.T).indexOf(name) < 0) return;
        var d = docsOf(b);
        if (d.listed) {
          var shown = d.docs.slice();
          if (d.answer) shown.push("ANSWER AT GCC");
          rows.push([b.key, (d.allReleased && d.validDoc) ? shown.join("\n") : "ANSWER AT GCC"]);
        }
      });
      rows.sort(function (a, b) { return a[0].localeCompare(b[0]); });
      rows.push(["TOTAL", rows.length]);
      return rows;
    }
    // why a CDF is / isn't in a pending table today (for the cross-check)
    function explainPending(name, key) {
      var b = blocks.filter(function (x) { return keyFull(x.key) === keyFull(key); })[0] ||
              blocks.filter(function (x) { return keyBase(x.key) === keyBase(key); })[0];
      if (!b) return "not in the CDF sheet";
      if (b.R !== "Open") return "CDF status is " + (b.R || "empty") + " — only Open CDFs are listed";
      if (b.S !== "9 - Final Approval") return "station is " + (b.S || "empty") + " — the table lists 9 - Final Approval only";
      if (splitNames(b.T).indexOf(name) < 0) return "signature position is \"" + (b.T || "empty") + "\" — not " + name;
      var d = docsOf(b);
      var rows = " (" + d.lines.join("; ") + ")";
      if (d.listed) return "listed: Open, 9 - Final Approval, signature " + name + ", " + (d.answer ? "ANSWER AT GCC" : "all Mantra statuses RELEASED") + rows;
      if (!d.us.length) return "no Mantra status (U) on its solution rows" + rows;
      if (!d.allReleased) return "not all Mantra statuses (U) are RELEASED" + rows;
      return "no valid solution document in AA" + rows;
    }

    // 100.8 WHERE ARE THE CDFS (Open + station 3, supplier from CDF!P)
    var where = {}, unassigned = [], whereLists = {};
    set.supplierOrder.forEach(function (k) { where[k] = 0; whereLists[k] = []; });
    open.forEach(function (b) {
      if (b.S !== "3 - Conversion Engineer") return;
      var p = upper(cell(cdf, b.anchor, CDF.P));
      var sup = set.suppliers[p];
      if (sup) {
        where[sup]++;
        whereLists[sup].push({ key: b.key, row: b.anchor + 1, signature: trim(b.T), dwg: trim(cell(cdf, b.anchor, CDF.J)),
          title: trim(cell(cdf, b.anchor, CDF.K)), desc: trim(cell(cdf, b.anchor, CDF.H)) });
      } else unassigned.push(b.key + (p ? " (" + trim(cell(cdf, b.anchor, CDF.P)) + ")" : ""));
    });
    var whereRows = set.supplierOrder.map(function (s) { return [s, where[s]]; });
    whereRows.push(["TOTAL", set.supplierOrder.reduce(function (s, k) { return s + where[k]; }, 0)]);

    return {
      tables: [
        { name: "CDF STATUS", headers: ["CDF STATUS", "CDF'S"], rows: cdfStatus },
        { name: "GCC CURRENT STATUS", headers: ["GCC CURRENT STATUS", "CDF'S"], rows: stations },
        { name: "GCC SIGNATURE POSITION", headers: ["SIGNATURE POSITION", "CDF'S"], rows: sigRows },
        { name: "WHERE ARE THE CDFS", headers: ["SUPPLIER", "CDF'S"], rows: whereRows }
      ].concat(set.pending.map(function (p) {
        return { name: p.name, headers: ["CDF NUMBER", "SOLUTION / DOCUMENT"], rows: pending(p.person) };
      })),
      explain: function (tableName, key) {
        var p = set.pending.filter(function (x) { return x.name === tableName; })[0];
        return p ? explainPending(p.person, key) : "";
      },
      whereLists: whereLists,
      pendingPeople: set.pending.map(function (p) { return { table: p.name, person: p.person }; }),
      openCount: open.length,
      station3Open: open.filter(function (b) { return b.S === "3 - Conversion Engineer"; }).length,
      unassigned: unassigned,
      otherOpenStations: otherStations
    };
  }

  function splitNames(t) {
    var s = trim(t);
    if (!s || s === "-") return [];
    return s.split(/\s*[,;\n]\s*/).map(trim).filter(function (x) { return x && x !== "-"; });
  }

  // ---------- 130.6 control totals ----------
  function controlChecks(sum, blocks, gccRows, onlyGcc) {
    var t = {}; sum.tables.forEach(function (x) { t[x.name] = x; });
    function val(table, label) { var r = t[table].rows.filter(function (x) { return x[0] === label; })[0]; return r ? r[1] : null; }
    var checks = [];
    var inBoth = gccRows.length - onlyGcc.length;
    checks.push({ name: "CDF STATUS TOTAL = GCC CDFs that exist in CDF", ok: val("CDF STATUS", "TOTAL") === inBoth,
      detail: val("CDF STATUS", "TOTAL") + " vs " + inBoth + (onlyGcc.length ? " · only in GCC: " + onlyGcc.join(", ") : "") });
    var stSum = OPEN_STATIONS.reduce(function (s, k) { return s + val("GCC CURRENT STATUS", k); }, 0);
    checks.push({ name: "Open stations add up to Open", ok: stSum === sum.openCount,
      detail: stSum + " vs " + sum.openCount + (sum.otherOpenStations.length ? " · other: " + sum.otherOpenStations.join(", ") : "") });
    checks.push({ name: "WHERE TOTAL + unassigned = Open at station 3", ok: val("WHERE ARE THE CDFS", "TOTAL") + sum.unassigned.length === sum.station3Open,
      detail: val("WHERE ARE THE CDFS", "TOTAL") + " + " + sum.unassigned.length + " vs " + sum.station3Open + (sum.unassigned.length ? " · unassigned: " + sum.unassigned.join(", ") : "") });
    return checks;
  }

  // ---------- compare with the tables already in CDF SUMMARY ----------
  // finds each table by its title and compares row by row
  function compareWithSheet(summary, sheetRows) {
    var res = [];
    summary.tables.forEach(function (tb) {
      var pos = findTable(sheetRows, tb);
      if (!pos) { res.push({ table: tb.name, found: false, diffs: [] }); return; }
      var diffs = [], current = [];
      for (var r = pos.r + 2; r < sheetRows.length; r++) {
        var a = trim(cell(sheetRows, r, pos.c)), b = cell(sheetRows, r, pos.c + 1);
        if (!a && isEmpty(b)) break;
        current.push([a, trim(b)]);
        if (a === "TOTAL") break;
      }
      var cur = {}; current.forEach(function (x) { cur[x[0]] = x[1]; });
      var calc = {}; tb.rows.forEach(function (x) { calc[x[0]] = str(x[1]); });
      Object.keys(calc).forEach(function (k) {
        if (!cur.hasOwnProperty(k)) diffs.push({ label: k, current: "(not in the sheet)", computed: calc[k] });
        else if (norm2(cur[k]) !== norm2(calc[k])) diffs.push({ label: k, current: cur[k], computed: calc[k] });
      });
      Object.keys(cur).forEach(function (k) { if (!calc.hasOwnProperty(k)) diffs.push({ label: k, current: cur[k], computed: "(not listed today)" }); });
      if (summary.explain) diffs.forEach(function (d) { if (d.label !== "TOTAL") { var w = summary.explain(tb.name, d.label); if (w) d.why = w; } });
      res.push({ table: tb.name, found: true, at: letter(pos.c) + (pos.r + 1), diffs: diffs, rows: tb.rows.length });
    });
    return res;
  }
  function norm2(v) { var s = trim(v).replace(/\r/g, ""); return /^-?\d+(\.0+)?$/.test(s) ? String(parseFloat(s)) : s; }
  function findTable(rows, tb) {
    for (var r = 0; r < rows.length; r++) {
      var row = rows[r] || [];
      for (var c = 0; c < row.length; c++) {
        if (upper(row[c]).indexOf(tb.name) === 0 && upper(cell(rows, r + 1, c)) === tb.headers[0].toUpperCase()) return { r: r, c: c };
      }
    }
    return null;
  }

  return { STAGE: STAGE, STAGES_DONE: STAGES_DONE, ALGORITHM_VERSION: ALGORITHM_VERSION, STAGES: STAGES, run: run, settings: settings, compareWithSheet: compareWithSheet, buildBlocks: buildBlocks, keyFull: keyFull, keyBase: keyBase, col: col, letter: letter };
});
