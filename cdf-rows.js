/*
 * CDF Updater 2 — manual rows: text copied from Excel, ChatGPT or Claude (TSV, CSV or a JSON array) → rows of sheet CDF,
 * each going where you say: insert above a row, insert below it, or write over it.
 *
 *   parse(text)                 → { rows: [[...]], kind: tsv | csv | json | header }
 *   layoutOf(rows, headers)     → how the columns map: from column A, from another column, by header names, or the
 *                                 13 fields the AI gives (CDF number, main file, ATT file, date, priority, description …)
 *   cellsOf(row, layout)        → [{ col, value }] inside the sheet's columns
 *   suggest(cells, sheet)       → where a row should go: a new revision under its CDF, the empty template row of its
 *                                 number, or after the last CDF
 *   plan(items, sheet)          → the writes in a safe order (write-overs first, then inserts from the bottom up) and
 *                                 the row each pasted row ends on
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory(require("./cdf-core.js"));
  else root.CdfRows = factory(root.CdfCore);
})(typeof self !== "undefined" ? self : this, function (K) {
  "use strict";

  var MAX_ROWS = 200;
  // the 13 fields the add-in's AI prompt asked for (1.x), in that order → CDF columns
  var AI13 = ["C", "B", "A", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O"];
  var KEY_RX = /^A330-1-\d{1,4}/i;

  // ---------- text → rows ----------
  function parse(text) {
    var s = K.str(text).replace(/^﻿/, "").replace(/ /g, " ");
    if (!s.trim()) return { rows: [], kind: "empty", error: "Nothing was pasted" };
    var fence = /```[a-zA-Z]*[ \t]*\r?\n?([\s\S]*?)```/.exec(s);
    if (fence) s = fence[1];
    // ChatGPT's "json" / "Copy code" label lines, when the text was selected rather than copied with the button
    s = s.replace(/^(?:\s*(?:json|javascript|python|csv|tsv|text|plaintext|copy code|copy|edit)\s*\r?\n)+/i, "");
    var t = s.trim();
    if ((t.charAt(0) === "[" || t.charAt(0) === "{") && t.indexOf("\t") < 0) {
      var j = K.looseJson(t);
      if (j.error) return { rows: [], kind: j.kind, error: j.error };
      var d = j.data, list = Array.isArray(d) ? d : d && Array.isArray(d.rows) ? d.rows : [d];
      if (!list.length) return { rows: [], kind: "json", error: "The JSON has no rows" };
      if (Array.isArray(list[0])) return { rows: list.map(function (r) { return r.map(cellText); }), kind: "json" };
      // objects: their keys become a header row
      var keys = [];
      list.forEach(function (o) { Object.keys(o || {}).forEach(function (k) { if (keys.indexOf(k) < 0) keys.push(k); }); });
      return { rows: [keys].concat(list.map(function (o) { return keys.map(function (k) { return cellText((o || {})[k]); }); })), kind: "json-objects" };
    }
    var rows = K.parseDelimited(s, s.indexOf("\t") >= 0 ? "\t" : null);
    if (!rows.length) return { rows: [], kind: "empty", error: "Nothing was pasted" };
    if (rows.length > MAX_ROWS + 1) return { rows: [], kind: "big", error: rows.length + " rows — paste " + MAX_ROWS + " at most at a time" };
    return { rows: rows.map(function (r) { return r.map(function (x) { return K.str(x).replace(/\r/g, ""); }); }), kind: s.indexOf("\t") >= 0 ? "tsv" : "csv" };
  }
  function cellText(v) { return v == null ? "" : typeof v === "object" ? JSON.stringify(v) : String(v); }

  // ---------- which column is which ----------
  // headers: the sheet's header row. → { mode: "header" | "ai13" | "from", start: col index, map: [col index per pasted column], header: bool, why }
  var ALIASES = { cdf_number: "C", cdf: "C", main_file: "B", att_file: "A", form_date: "F", date: "F", received_date: "F", priority: "G", description: "H",
    oem_category: "I", oem: "I", tags: "I", drawing_number: "J", drawing: "J", dwg: "J", installation: "K", dwg_name: "K", code: "L", nrc: "L",
    drawing_rev: "M", dwg_rev: "M", author: "N", technical_adviser: "O", adviser: "O", cdf_rev: "D", revision: "D", notes: "E" };
  function headerMap(first, headers) {
    var map = [], named = 0;
    first.forEach(function (h, i) {
      var n = K.norm(h), c = -1;
      (headers || []).forEach(function (sh, j) { if (c < 0 && n && K.norm(sh) === n) c = j; });
      if (c < 0) { var a = ALIASES[n.replace(/[^a-z0-9]+/g, "_").replace(/^_|_$/g, "")]; if (a) c = K.col(a); }
      if (c < 0 && /^[a-z]{1,2}$/i.test(K.trim(h))) c = K.col(K.trim(h));   // a header that is just a column letter
      map[i] = c; if (c >= 0) named++;
    });
    return { map: map, named: named };
  }
  function layoutOf(rows, headers, opt) {
    opt = opt || {};
    var first = rows[0] || [], width = rows.reduce(function (m, r) { return Math.max(m, r.length); }, 0);
    if (opt.mode === "from") return fromCol(opt.start || 0, width, false, "From column " + K.letter(opt.start || 0));
    if (opt.mode === "ai13") return ai13(false);
    // a header row: most of its cells are column names (and none is a CDF number)
    var hm = headerMap(first, headers);
    if (rows.length > 1 && hm.named >= Math.max(2, Math.ceil(first.filter(function (x) { return K.trim(x); }).length / 2)) && !first.some(function (x) { return KEY_RX.test(K.trim(x)); }))
      return { mode: "header", map: hm.map, header: true, why: "The first row names the columns (" + hm.named + " of them)" };
    var keyAt = first.map(function (x) { return K.trim(x); }).findIndex(function (x) { return KEY_RX.test(x); });
    if (width === 13 && keyAt === 0) return ai13(false);
    if (keyAt >= 0 && /-ATT\.?$/i.test(K.trim(first[keyAt]))) return fromCol(K.col("A") - keyAt, width, false, "A copy of sheet rows (starts with the ATT link in A)");
    if (keyAt >= 0) return fromCol(Math.max(0, K.col("C") - keyAt), width, false, "The CDF number lands in C");
    return fromCol(width >= 30 ? 0 : K.col("C"), width, false, width >= 30 ? "A whole sheet row: from column A" : "No CDF number found: from column C — choose another start if it's not right");
    function fromCol(start, w, header, why) { var map = []; for (var i = 0; i < w; i++) map.push(start + i); return { mode: "from", start: start, map: map, header: header, why: why }; }
    function ai13(header) { return { mode: "ai13", map: AI13.map(K.col), header: header, why: "13 fields: CDF number, main file, ATT file, date, priority, description, tags, DWG, DWG name, NRC, DWG rev, author, adviser" }; }
  }

  // ---------- one pasted row → cells ----------
  // opt: { firstCol, lastCol } (the columns rows may be written to); → { cells: [{col, value}], dropped: n }
  function cellsOf(row, layout, opt) {
    opt = opt || {};
    var lo = K.col(opt.firstCol || "A"), hi = K.col(opt.lastCol || "AH"), cells = [], dropped = 0, seen = {};
    (row || []).forEach(function (v, i) {
      var c = layout.map[i];
      if (c == null || c < 0) { if (K.trim(v)) dropped++; return; }
      if (c < lo || c > hi) { if (K.trim(v)) dropped++; return; }
      seen[c] = true;
      cells.push({ col: c, value: valueOf(v, c) });
    });
    // the 13 fields: the revision letter goes to D as well (A330-1-723.A → A, A330-1-723.- → -)
    if (layout.mode === "ai13" || layout.mode === "header") {
      var key = cells.filter(function (x) { return x.col === K.col("C"); })[0], dCol = K.col("D");
      if (key && !seen[dCol] && dCol >= lo && dCol <= hi) { var p = K.keyParts(key.value); if (p) cells.push({ col: dCol, value: p.rev }); }
    }
    cells.sort(function (a, b) { return a.col - b.col; });
    return { cells: cells, dropped: dropped };
  }
  // dates become real dates (F received date, Q closed date); everything else stays text as pasted
  var DATE_COLS = [K.col("F"), K.col("Q")];
  function valueOf(v, c) {
    var s = K.trim(v);
    if (c === K.col("C")) { var p = K.keyParts(s); if (p && /^A330-1-\d+(\.([A-Z]{1,2}|-))?$/i.test(K.keyFull(s))) return p.full; }   // the key as C writes it
    if (DATE_COLS.indexOf(c) >= 0) { var d = K.daySerial(s); if (d != null) return d; }
    return s;
  }
  function keyOf(cells) { var k = (cells || []).filter(function (x) { return x.col === K.col("C"); })[0]; return k ? K.trim(k.value) : ""; }

  // ---------- where it goes ----------
  // sheet: the CDF values (header first). → { mode: insert-below | insert-above | overwrite, row (1-based), why }
  function suggest(cells, sheet, opt) {
    opt = opt || {};
    var start = opt.startRow || 2, key = keyOf(cells), p = K.keyParts(key), blocks = K.buildBlocks(sheet || []);
    var endOf = function (b) { return b.rows[b.rows.length - 1] + 1; };
    if (p) {
      var full = blocks.filter(function (b) { return K.sameCdf(b.key, key); })[0];
      if (full) return { mode: "overwrite", row: full.anchor + 1, why: key + " is already on row " + (full.anchor + 1) + " — writing over it" };
      var same = blocks.filter(function (b) { var q = K.keyParts(b.key); return q && q.num === p.num; });
      if (same.length) {
        var lastB = same[same.length - 1];
        return { mode: "insert-below", row: endOf(lastB), why: "New revision: under " + lastB.key + " (its last row is " + endOf(lastB) + ")" };
      }
      // the empty template row prepared for this number (A/B hold its links, C is empty)
      for (var r = start - 1; r < (sheet || []).length; r++) {
        var row = sheet[r] || [];
        if (K.trim(row[2])) continue;
        var q2 = K.keyParts(K.trim(row[0]) || K.trim(row[1]));
        if (q2 && q2.num === p.num) return { mode: "overwrite", row: r + 1, why: "The empty template row of " + q2.label + " (row " + (r + 1) + ")" };
      }
      // under the CDF with the next lower number
      var lower = blocks.filter(function (b) { var q = K.keyParts(b.key); return q && q.num < p.num; });
      if (lower.length) { var lb = lower.reduce(function (a, b) { return K.keyParts(b.key).num >= K.keyParts(a.key).num ? b : a; }); return { mode: "insert-below", row: endOf(lb), why: "After " + lb.key + ", the CDF before it" }; }
    }
    var last = blocks[blocks.length - 1];
    return last ? { mode: "insert-below", row: endOf(last), why: "After the last CDF (" + last.key + ")" } : { mode: "overwrite", row: start, why: "The first row under the header" };
  }

  // ---------- the writes, in a safe order ----------
  // items: [{ cells, mode, row }] (pasted order). → { ops: [...], rows: [final row per item], errors: [] }
  //   write-overs first (row numbers are the sheet's as it is now), then inserts from the bottom up; rows inserted at the
  //   same place keep the pasted order.
  function plan(items, sheet, opt) {
    opt = opt || {};
    var errors = [], start = opt.startRow || 2, n = (sheet || []).length, over = {}, ops = [], groups = {};
    items.forEach(function (it, i) {
      var r = Math.floor(Number(it.row));
      if (!(r >= start)) { errors.push("Row " + (i + 1) + ": choose a row from " + start + " down"); return; }
      if (it.mode === "overwrite") {
        if (over[r] != null) errors.push("Rows " + (over[r] + 1) + " and " + (i + 1) + " both write over row " + r);
        over[r] = i;
        ops.push({ kind: "overwrite", row: r, cells: it.cells, item: i });
      } else {
        var at = it.mode === "insert-above" ? r : r + 1;
        if (r > n + 1 && it.mode === "insert-below") at = r + 1;
        (groups[at] = groups[at] || []).push(i);
      }
    });
    var ats = Object.keys(groups).map(Number).sort(function (a, b) { return b - a; });
    ats.forEach(function (at) { ops.push({ kind: "insert", at: at, count: groups[at].length, items: groups[at], rows: groups[at].map(function (i) { return items[i].cells; }) }); });
    // the row every item ends on, once everything is written
    var final = items.map(function () { return null; });
    ops.forEach(function (op) {
      if (op.kind === "overwrite") {
        var shift = ats.filter(function (a) { return a <= op.row; }).reduce(function (s, a) { return s + groups[a].length; }, 0);
        final[op.item] = op.row + shift;
      } else {
        var before = ats.filter(function (a) { return a < op.at; }).reduce(function (s, a) { return s + groups[a].length; }, 0);
        op.items.forEach(function (i, j) { final[i] = op.at + before + j; });
      }
    });
    return { ops: ops, rows: final, errors: errors, inserted: ats.reduce(function (s, a) { return s + groups[a].length; }, 0) };
  }
  // what is on a row now, in a few words (for the preview)
  function rowSays(sheet, r) {
    var row = (sheet || [])[r - 1];
    if (!row) return "an empty row";
    var key = K.trim(row[2]);
    if (key) return key + (K.trim(row[17]) ? " · " + K.trim(row[17]) : "");
    var t = K.keyParts(K.trim(row[0]) || K.trim(row[1]));
    if (t) return "empty template row of " + t.label;
    return row.some(function (x) { return K.trim(x); }) ? "a solution row" : "an empty row";
  }

  return { MAX_ROWS: MAX_ROWS, AI13: AI13, parse: parse, layoutOf: layoutOf, headerMap: headerMap, cellsOf: cellsOf, keyOf: keyOf, suggest: suggest, plan: plan, rowSays: rowSays, valueOf: valueOf };
});
