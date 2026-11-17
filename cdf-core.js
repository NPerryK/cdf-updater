/*
 * CDF Updater 2 — shared, pure helpers (no Office.js): the pane, present.html and the tests use them.
 *   columns and keys · the CDF block map · TSV / CSV · lenient JSON · rules.txt · which sheet a file is
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.CdfCore = factory();
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  // ---------- small things ----------
  function str(v) { return v == null ? "" : String(v); }
  function trim(v) { return str(v).replace(/ /g, " ").trim(); }
  function isEmpty(v) { return v === "" || v === null || v === undefined || (typeof v === "string" && trim(v) === ""); }
  function col(L) { var n = 0; L = String(L).toUpperCase(); for (var i = 0; i < L.length; i++) n = n * 26 + L.charCodeAt(i) - 64; return n - 1; }
  function letter(i) { var n = i + 1, s = ""; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
  function norm(h) { return trim(h).toLowerCase().replace(/\s+/g, " "); }

  // ---------- dates: Excel serials (1900 system) ----------
  var DAY = 86400000, EPOCH = Date.UTC(1899, 11, 30);
  function serialOf(d) { return Math.round((Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) - EPOCH) / DAY); }   // the local calendar day
  function dateOf(serial) { var t = EPOCH + Math.floor(serial) * DAY, u = new Date(t); return new Date(u.getUTCFullYear(), u.getUTCMonth(), u.getUTCDate()); }
  // a cell as a day serial: a number, an ISO text, or dd/mm/yyyy (as in Israel); null when it isn't a date
  function daySerial(v) {
    if (typeof v === "number" && isFinite(v) && v > 0) return Math.floor(v);
    var s = trim(v), m;
    if (!s) return null;
    if ((m = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(s))) return check(+m[1], +m[2], +m[3]);
    if ((m = /^(\d{1,2})[\/.](\d{1,2})[\/.](\d{4})$/.exec(s))) return check(+m[3], +m[2], +m[1]);
    return null;
    function check(y, mo, d) { var t = Date.UTC(y, mo - 1, d), x = new Date(t); return x.getUTCMonth() === mo - 1 && x.getUTCDate() === d ? Math.round((t - EPOCH) / DAY) : null; }
  }
  function ddmm(serial) { var d = dateOf(serial); return ("0" + d.getDate()).slice(-2) + "/" + ("0" + (d.getMonth() + 1)).slice(-2); }
  function ddmmyyyy(serial) { return ddmm(serial) + "/" + dateOf(serial).getFullYear(); }
  // the Sunday that starts the week of a day (weeks run Sunday–Saturday, as at IAI)
  function weekStart(serial) { return serial - ((serial + 6) % 7); }   // serial 1 = Sun 1 Jan 1900

  // ---------- CDF keys ----------
  function keyFull(v) { return str(v).replace(/[\s ]+/g, "").toUpperCase(); }
  function keyBase(v) { return keyFull(v).replace(/\.-$/, "").replace(/\.$/, ""); }                 // 05.3: a final ".-" or "." may go
  // "A330-1-723.A" → { num: 723, rev: "A", base: "A330-1-723" }; "A330-1-723.-" → rev "-"; anything else → null
  function keyParts(v) {
    var m = /^A330-1-0*(\d{1,4})(?:\.([A-Z]{1,2}|-))?(?:[-.]|$)/.exec(keyFull(v));
    if (!m) return null;
    var num = +m[1], label = "A330-1-" + (num < 100 ? ("00" + num).slice(-3) : String(num)), rev = m[2] || "-";
    return { num: num, rev: rev, label: label, full: label + "." + rev };   // full: the way column C writes it (A330-1-978.-, A330-1-723.A)
  }
  // what a person types: "856", "740.A", "a330-1-856.a"
  function sameCdf(a, b) { var p = keyParts(a), q = keyParts(b); return p && q ? p.num === q.num && p.rev === q.rev : keyFull(a) === keyFull(b); }
  function parseQuery(q) {
    var s = keyFull(q).replace(/^A330-1-/, ""), m = /^0*(\d{1,4})(?:\.([A-Z]{1,2}|-))?$/.exec(s);
    return m ? { num: +m[1], rev: m[2] || null } : null;
  }

  // ---------- the CDF block map: an anchor (a key in C) and the solution rows under it ----------
  var C_KEY = 2, DATA_COLS = [col("R"), col("U"), col("AA"), col("AB")];
  function buildBlocks(rows) {
    var blocks = [], cur = null;
    for (var r = 1; r < (rows || []).length; r++) {
      var key = trim((rows[r] || [])[C_KEY]);
      if (key) { cur = { key: key, anchor: r, rows: [r] }; blocks.push(cur); continue; }
      if (!cur) continue;
      var row = rows[r] || [];
      if (DATA_COLS.some(function (c) { return !isEmpty(row[c]); })) cur.rows.push(r);
    }
    return blocks;
  }

  // ---------- CSV / TSV ----------
  // Excel copies a cell with a line break or a tab inside quotes ("" for a quote). A quote that isn't at the start of a
  // cell is just a character (12" PANEL), so hand-typed text and ChatGPT tables read as they look.
  function parseDelimited(text, delim) {
    text = str(text).replace(/^﻿/, "");
    if (!delim) {
      var first = text.split(/\r?\n/)[0] || "";
      delim = ["\t", ",", ";"].map(function (d) { return [d, first.split(d).length]; }).sort(function (a, b) { return b[1] - a[1]; })[0][0];
    }
    var rows = [], row = [], field = "", q = false, atStart = true;
    for (var i = 0; i < text.length; i++) {
      var c = text[i];
      if (q) {
        if (c === '"' && text[i + 1] === '"') { field += '"'; i++; }
        else if (c === '"') q = false;
        else field += c;
        continue;
      }
      if (c === '"' && atStart) { q = true; atStart = false; continue; }
      if (c === delim) { row.push(field); field = ""; atStart = true; continue; }
      if (c === "\n" || c === "\r") {
        if (c === "\r" && text[i + 1] === "\n") i++;
        row.push(field); rows.push(row); row = []; field = ""; atStart = true; continue;
      }
      field += c; atStart = false;
    }
    if (field !== "" || row.length) { row.push(field); rows.push(row); }
    return rows.filter(function (r) { return r.some(function (x) { return trim(x) !== ""; }); });
  }

  // ---------- lenient JSON (an AI's reply) → { data } or { kind, error } ----------
  function looseJson(text) {
    var s = str(text), a = s.indexOf("["), o = s.indexOf("{");
    var start = a < 0 ? o : o < 0 ? a : Math.min(a, o);
    if (start < 0) return { kind: "nojson", error: "No JSON found. Copy the code block from the AI's reply (the Copy button in the corner of the block)." };
    var close = s.charAt(start) === "[" ? "]" : "}", end = s.lastIndexOf(close);
    if (end < start) return { kind: "cut", error: "The JSON is cut off — make sure you copied the whole block" };
    var body = s.slice(start, end + 1), tries = [body, body.replace(/[“”„‟″]/g, '"').replace(/[‘’]/g, "'")];
    tries.push(tries[1].replace(/,\s*([\]}])/g, "$1"));
    tries.push(tries[2].replace(/[\t\r\n]+/g, " "));
    tries.push(tries[2].replace(/^\s*\/\/.*$/gm, "").replace(/[\t\r\n]+/g, " "));
    if (tries[1].indexOf('"') < 0) tries.push(tries[2].replace(/'((?:[^'\\]|\\.)*)'/g, function (m, x) { return JSON.stringify(x.replace(/\\'/g, "'")); })
      .replace(/\bNone\b/g, "null").replace(/\bTrue\b/g, "true").replace(/\bFalse\b/g, "false").replace(/[\t\r\n]+/g, " "));
    for (var i = 0; i < tries.length; i++) { try { return { data: JSON.parse(tries[i]) }; } catch (e) { /* next */ } }
    return { kind: "invalid", error: "The JSON is not valid — the reply is not in a format the add-in reads." };
  }

  // ---------- rules.txt ----------
  var DEFAULT_RULES = {
    imports: [
      { target: "CDF", headers: ["CDF NUM (LINK TO DEV)", "CDF STATUS (GCC)"], file: "" },
      { target: "GCC", headers: ["CDF Number", "Current Station"], file: "CDF_List_Export" },
      { target: "MANTRAREV", headers: ["Doc ID", "Release Date"], file: "" }
    ],
    rows: { sheet: "CDF", key: "C", headerRow: 1, startRow: 2, firstCol: "A", lastCol: "AH" },
    log: "UPDATE LOG", stuckDays: 14, targets: {}, exclude: ["A330-1-052 (another)"]
  };
  function words(s) { var out = [], re = /"([^"]*)"|(\S+)/g, m; while ((m = re.exec(s))) out.push(m[1] != null ? m[1] : m[2]); return out; }
  function parseRules(text) {
    var out = { imports: [], rows: JSON.parse(JSON.stringify(DEFAULT_RULES.rows)), log: DEFAULT_RULES.log, stuckDays: DEFAULT_RULES.stuckDays, targets: {}, exclude: [], errors: [] }, sawExclude = false;
    str(text).split(/\r?\n/).forEach(function (raw, i) {
      var line = raw.replace(/(^|\s)#.*$/, "").trim();
      if (!line) return;
      var w = words(line), k = w[0].toUpperCase(), at = "Line " + (i + 1) + ": ";
      if (k === "IMPORT") {
        // IMPORT "GCC" HEADERS "CDF Number" "Current Station" [FILE "CDF_List_Export"]
        var target = w[1], h = [], file = "", mode = null;
        if (!target) { out.errors.push(at + "IMPORT needs a sheet name"); return; }
        w.slice(2).forEach(function (x) {
          var u = x.toUpperCase();
          if (u === "HEADERS" || u === "FILE") { mode = u; return; }
          if (mode === "HEADERS") h.push(x); else if (mode === "FILE") file = x;
        });
        if (!h.length) { out.errors.push(at + "IMPORT \"" + target + "\" needs HEADERS — two or three column names that only this file has"); return; }
        out.imports.push({ target: target, headers: h, file: file });
      } else if (k === "ROWS_SHEET") out.rows.sheet = w[1] || out.rows.sheet;
      else if (k === "ROWS_KEY") out.rows.key = (w[1] || "C").toUpperCase();
      else if (k === "ROWS_HEADER_ROW" || k === "ROWS_START_ROW") {
        var n = parseInt(w[1], 10);
        if (!(n >= 1)) { out.errors.push(at + k + " needs a row number"); return; }
        out.rows[k === "ROWS_HEADER_ROW" ? "headerRow" : "startRow"] = n;
      } else if (k === "ROWS_COLUMNS") {
        var m = /^([A-Z]{1,2}):([A-Z]{1,2})$/i.exec(w[1] || "");
        if (!m) { out.errors.push(at + "ROWS_COLUMNS needs a range like A:AH"); return; }
        out.rows.firstCol = m[1].toUpperCase(); out.rows.lastCol = m[2].toUpperCase();
      } else if (k === "LOG") out.log = w[1] || "";
      else if (k === "STUCK_DAYS") { var d = parseInt(w[1], 10); if (d > 0) out.stuckDays = d; else out.errors.push(at + "STUCK_DAYS needs a number of days"); }
      else if (k === "EXCLUDE") { sawExclude = true; if (w[1]) out.exclude.push(w[1]); else out.errors.push(at + "EXCLUDE needs a CDF number in quotes"); }
      else if (k === "TARGET") {
        // TARGET OPEN 150 [BY 2026-12-31] — a goal line on the open chart (optional)
        var v = parseInt(w[2], 10);
        if (!w[1] || !(v >= 0)) { out.errors.push(at + "TARGET needs a name and a number, e.g. TARGET OPEN 150 BY 2026-12-31"); return; }
        out.targets[w[1].toUpperCase()] = { value: v, by: w[3] && w[3].toUpperCase() === "BY" ? daySerial(w[4]) : null };
      } else out.errors.push(at + "\"" + w[0] + "\" isn't a setting this version knows (2.0 runs no algorithm — see README)");
    });
    if (!out.imports.length) out.imports = JSON.parse(JSON.stringify(DEFAULT_RULES.imports));
    if (!sawExclude) out.exclude = DEFAULT_RULES.exclude.slice();
    return out;
  }

  // ---------- which sheet is this file? by its header row, not its name ----------
  // headers: the first row (or a row near the top) of a sheet in the file → the IMPORT it belongs to, or null
  function importFor(headerRows, imports) {
    var rows = (headerRows || []).slice(0, 5).map(function (r) { return (r || []).map(norm); });
    var hit = null;
    (imports || []).forEach(function (imp) {
      if (hit) return;
      var want = imp.headers.map(norm);
      if (rows.some(function (r) { return want.every(function (h) { return r.indexOf(h) >= 0; }); })) hit = imp;
    });
    return hit;
  }
  function fileNameFor(name, imports) {
    var n = str(name).toLowerCase(), hit = null;
    (imports || []).forEach(function (imp) { if (!hit && imp.file && n.indexOf(imp.file.toLowerCase()) >= 0) hit = imp; });
    return hit;
  }
  // a sheet's header row index (0-based) among the first rows: the row that has these headers
  function headerRow(values, must) {
    var want = (must || []).map(norm);
    for (var r = 0; r < Math.min(6, (values || []).length); r++) {
      var row = (values[r] || []).map(norm);
      if (want.every(function (h) { return row.indexOf(h) >= 0; })) return r;
    }
    return -1;
  }
  function columnIndex(headers, name) { var n = norm(name); for (var i = 0; i < (headers || []).length; i++) if (norm(headers[i]) === n) return i; return -1; }

  return {
    str: str, trim: trim, isEmpty: isEmpty, col: col, letter: letter, norm: norm,
    serialOf: serialOf, dateOf: dateOf, daySerial: daySerial, ddmm: ddmm, ddmmyyyy: ddmmyyyy, weekStart: weekStart,
    keyFull: keyFull, keyBase: keyBase, keyParts: keyParts, sameCdf: sameCdf, parseQuery: parseQuery, buildBlocks: buildBlocks,
    parseDelimited: parseDelimited, looseJson: looseJson, DEFAULT_RULES: DEFAULT_RULES, parseRules: parseRules,
    importFor: importFor, fileNameFor: fileNameFor, headerRow: headerRow, columnIndex: columnIndex
  };
});
