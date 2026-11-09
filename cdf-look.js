/*
 * CDF Updater — the look of sheet CDF.
 *
 *   110.6   Calibri 10; S:X purple #60497A with white text; a thin continuous black border on every cell and every
 *           merged area; Middle Align; Wrap Text and a fitted height; no ";;;" number format.
 *   110.7   three conditional-format rules: TRIM($R2)="Released" green #00B050, TRIM($R2)="Cancelled" red #FF0000,
 *           and =TRUE on the continuation rows' Q:X and AA:AH, which only draws a thin black top line (=FALSE when
 *           there are no continuations). No helper cells.
 *
 * Looks (the LOOK page): a request in the user's words goes to ChatGPT or Claude as a short prompt with the column
 * names only — no values from the sheet. The AI answers with a small JSON rule; the add-in checks it, counts the
 * rows it matches here, and writes it as a live conditional format next to the three rules of 110.7.
 *
 * Everything here is pure (no Office.js): taskpane.js reads and writes, the tests run it in node.
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.CdfLook = factory();
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  var LOOK_VERSION = 1;
  var LAST_COL = 33;                                   // AH
  var PURPLE = "#60497A", WHITE = "#FFFFFF", BLACK = "#000000", GREEN = "#00B050", RED = "#FF0000";
  var MAX_FORMULA = 255;                               // what Excel takes in a conditional-format formula
  var MAX_ADDRESS = 6000;                              // characters in one multi-area address (more → another rule)

  function col(L) { var n = 0; L = String(L).toUpperCase(); for (var i = 0; i < L.length; i++) n = n * 26 + L.charCodeAt(i) - 64; return n - 1; }
  function letter(i) { var n = i + 1, s = ""; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
  function isEmpty(v) { return v === "" || v === null || v === undefined || (typeof v === "string" && v.trim() === ""); }

  // ================================================================ 110.6 / 110.7
  // the last row with anything in it (1-based); 1 = only the header row
  function lastRow(values) {
    for (var r = (values || []).length - 1; r >= 0; r--) if ((values[r] || []).some(function (v) { return !isEmpty(v); })) return r + 1;
    return 0;
  }
  // the continuation rows (1-based): every row of a block after its anchor — the block map of 07.2 / 100.2
  function continuationRows(values, algo) {
    var out = [];
    algo.buildBlocks(values || []).forEach(function (b) { b.rows.slice(1).forEach(function (r) { out.push(r + 1); }); });
    return out.sort(function (a, b) { return a - b; }).filter(function (r, i, a) { return i === 0 || a[i - 1] !== r; });
  }
  // [3,4,5,9] → [[3,5],[9,9]]
  function runsOf(rows) {
    var out = [];
    (rows || []).slice().sort(function (a, b) { return a - b; }).forEach(function (r) {
      var last = out[out.length - 1];
      if (last && r === last[1] + 1) last[1] = r; else if (!last || r > last[1]) out.push([r, r]);
    });
    return out;
  }
  // the continuation rows as Q:X and AA:AH areas, in addresses of at most maxLen characters each
  function contAddresses(rows, maxLen) {
    var max = maxLen || MAX_ADDRESS, out = [], cur = "";
    runsOf(rows).forEach(function (run) {
      var part = "Q" + run[0] + ":X" + run[1] + ", AA" + run[0] + ":AH" + run[1];
      if (cur && cur.length + 2 + part.length > max) { out.push(cur); cur = ""; }
      cur = cur ? cur + ", " + part : part;
    });
    if (cur) out.push(cur);
    return out;
  }
  // what the add-in writes, from the CDF values after the update
  function lookPlan(values, algo) {
    var last = lastRow(values);
    if (last < 2) return null;
    var cont = continuationRows(values, algo), chunks = contAddresses(cont);
    return {
      last: last, all: "A1:AH" + last, body: "A2:AH" + last, purple: "S2:X" + last, cont: cont,
      rules: [
        { id: "released", areas: ["A2:AH" + last], formula: '=TRIM($R2)="Released"', fill: GREEN },
        { id: "cancelled", areas: ["A2:AH" + last], formula: '=TRIM($R2)="Cancelled"', fill: RED },
        { id: "cont", areas: chunks.length ? chunks : ["Q2:X" + last + ", AA2:AH" + last], formula: cont.length ? "=TRUE" : "=FALSE", top: BLACK }
      ]
    };
  }
  // an existing conditional format: one of the 110.7 rules (also a stray piece of one), one of the saved looks, or the user's own
  function ruleKind(formula, lookFormulas) {
    var f = String(formula == null ? "" : formula).replace(/\s+/g, "");
    if (f.charAt(0) !== "=") f = "=" + f;
    var u = f.toUpperCase(), m = /^=TRIM\(\$R\$?\d+\)="(RELEASED|CANCELLED)"$/.exec(u);
    if (m) return m[1] === "RELEASED" ? "released" : "cancelled";
    if (u === "=TRUE" || u === "=FALSE") return "cont";
    if ((lookFormulas || []).some(function (x) { return String(x).replace(/\s+/g, "").toUpperCase() === u; })) return "look";
    return null;
  }
  // existing rules: [{ formula, areas }] → what was found, for the report
  function describeExisting(items, lookFormulas) {
    var out = { released: 0, cancelled: 0, cont: 0, look: 0, ownRules: 0, pieces: { released: 0, cancelled: 0, cont: 0 } };
    (items || []).forEach(function (it) {
      var k = ruleKind(it.formula, lookFormulas);
      if (!k) { out.ownRules++; return; }
      out[k]++;
      if (out.pieces[k] != null) out.pieces[k] += Math.max(1, it.areas || 1);
    });
    out.broken = ["released", "cancelled", "cont"].filter(function (k) { return out[k] > 1 || (k !== "cont" && out.pieces[k] > 1); });
    return out;
  }
  // cells with the hidden number format ";;;" (formats: rows from sheet row r0, 1-based)
  function hiddenFormatCells(formats, r0) {
    var out = [];
    (formats || []).forEach(function (row, i) {
      (row || []).forEach(function (f, c) { if (c <= LAST_COL && String(f == null ? "" : f).replace(/\s+/g, "") === ";;;") out.push({ row: (r0 || 1) + i, col: c, a1: letter(c) + ((r0 || 1) + i) }); });
    });
    return out;
  }
  // rows: [{ row, hidden }] → runs of rows a filter doesn't hide (autofit never touches a hidden row)
  function visibleRuns(rows) {
    return runsOf((rows || []).filter(function (x) { return !x.hidden; }).map(function (x) { return x.row; }));
  }
  // Heights only grow: a row someone made taller (text in merged cells needs it — Excel's autofit skips merged cells)
  // goes back to its height; rows this update added fit their text either way.
  function restoreHeights(before, after, fresh) {
    var isFresh = {}, out = [];
    (fresh || []).forEach(function (r) { isFresh[r] = true; });
    (before || []).forEach(function (b) {
      if (b.hidden || isFresh[b.row]) return;
      var a = after[b.row];
      if (typeof a === "number" && typeof b.height === "number" && b.height > a + 0.01) out.push({ row: b.row, height: b.height });
    });
    return out;
  }

  // ================================================================ looks: a rule from ChatGPT or Claude
  var OPS = {
    is: { need: "text", say: "is" }, is_not: { need: "text", say: "is not" },
    has: { need: "text", say: "contains" }, has_not: { need: "text", say: "doesn't contain" },
    starts: { need: "text", say: "starts with" },
    empty: { need: "none", say: "is empty" }, filled: { need: "none", say: "is not empty" },
    one_of: { need: "list", say: "is one of" },
    ">": { need: "number", say: ">" }, ">=": { need: "number", say: "≥" }, "<": { need: "number", say: "<" }, "<=": { need: "number", say: "≤" },
    older_than_days: { need: "days", say: "is more than N days ago" }, within_days: { need: "days", say: "is within the last N days" },
    before: { need: "date", say: "is before" }, after: { need: "date", say: "is after" }, on: { need: "date", say: "is on" }
  };
  var OP_ALIASES = {
    "=": "is", "==": "is", equals: "is", eq: "is", is_equal: "is", "!=": "is_not", "<>": "is_not", not_equals: "is_not", ne: "is_not", isnot: "is_not",
    contains: "has", includes: "has", not_contains: "has_not", does_not_contain: "has_not", excludes: "has_not",
    starts_with: "starts", startswith: "starts", begins_with: "starts", is_empty: "empty", blank: "empty", not_empty: "filled", is_not_empty: "filled", not_blank: "filled",
    in: "one_of", any_of: "one_of", oneof: "one_of", gt: ">", gte: ">=", ge: ">=", lt: "<", lte: "<=", le: "<=",
    older_than: "older_than_days", more_than_days_ago: "older_than_days", days_ago_more_than: "older_than_days",
    newer_than_days: "within_days", in_last_days: "within_days", last_days: "within_days"
  };
  var COLOR_NAMES = {
    yellow: "#FFFF00", "light yellow": "#FFF2CC", orange: "#FFC000", "light orange": "#FCE4D6", red: "#FF0000", "light red": "#FFC7CE", "dark red": "#C00000",
    green: "#00B050", "light green": "#C6EFCE", "dark green": "#006100", blue: "#0070C0", "light blue": "#DDEBF7", "dark blue": "#002060",
    purple: "#7030A0", "light purple": "#E4DFEC", pink: "#FF66CC", gray: "#BFBFBF", grey: "#BFBFBF", "light gray": "#D9D9D9", "light grey": "#D9D9D9",
    "dark gray": "#595959", "dark grey": "#595959", black: "#000000", white: "#FFFFFF", amber: "#FFC000", gold: "#FFC000", brown: "#996633", teal: "#008080", cyan: "#00FFFF"
  };
  function hex(v) {
    if (v == null || v === "") return null;
    var s = String(v).trim().toLowerCase().replace(/[_-]+/g, " ");
    if (COLOR_NAMES[s]) return COLOR_NAMES[s];
    var m = /^#?([0-9a-f]{6})$/.exec(s);
    if (m) return "#" + m[1].toUpperCase();
    m = /^#?([0-9a-f])([0-9a-f])([0-9a-f])$/.exec(s);
    if (m) return ("#" + m[1] + m[1] + m[2] + m[2] + m[3] + m[3]).toUpperCase();
    return undefined;
  }
  function trimX(s) { return String(s).replace(/^ +| +$/g, "").replace(/ {2,}/g, " "); }    // Excel's TRIM: spaces only
  function lower(s) { return String(s).toLowerCase(); }
  function normHeader(h) { return String(h == null ? "" : h).toLowerCase().replace(/[^a-z0-9א-ת]+/g, " ").trim(); }

  // "R" / "r" / "CDF STATUS - GCC" (a header) → column index, or -1
  function columnOf(v, headers) {
    var s = String(v == null ? "" : v).trim();
    if (/^[A-Za-z]{1,2}$/.test(s)) { var c = col(s); return c <= LAST_COL ? c : -1; }
    var n = normHeader(s), hit = -1;
    (headers || []).forEach(function (h, i) { if (hit < 0 && i <= LAST_COL && n && normHeader(h) === n) hit = i; });
    return hit;
  }
  // "row" / "A:AH" / "S:X" / "H" / ["R","S"] / "R,S" → [[c1,c2]] within A:AH, merged; null when one is outside
  function paintColumns(cells, headers) {
    if (cells == null || cells === "" || /^(row|rows|whole row|entire row|all)$/i.test(String(cells).trim())) return [[0, LAST_COL]];
    var parts = Array.isArray(cells) ? cells : String(cells).split(/[,;]/), spans = [];
    for (var i = 0; i < parts.length; i++) {
      var p = String(parts[i]).trim(); if (!p) continue;
      var m = /^([A-Za-z]{1,2})\s*[:\-]\s*([A-Za-z]{1,2})$/.exec(p), a, b;
      if (m) { a = col(m[1]); b = col(m[2]); if (a > b) { var t = a; a = b; b = t; } }
      else { a = b = columnOf(p, headers); }
      if (a < 0 || b > LAST_COL) return null;
      spans.push([a, b]);
    }
    if (!spans.length) return [[0, LAST_COL]];
    spans.sort(function (x, y) { return x[0] - y[0]; });
    var out = [spans[0].slice()];
    spans.slice(1).forEach(function (s) { var l = out[out.length - 1]; if (s[0] <= l[1] + 1) l[1] = Math.max(l[1], s[1]); else out.push(s.slice()); });
    return out;
  }
  function spansText(spans) { return spans.map(function (s) { return s[0] === s[1] ? letter(s[0]) : letter(s[0]) + ":" + letter(s[1]); }).join(", "); }
  function paintAreas(spans, last) { return spans.map(function (s) { return letter(s[0]) + "2:" + letter(s[1]) + last; }); }

  function isoDate(v) {
    var m = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(String(v == null ? "" : v).trim()) || null, y, mo, d;
    if (m) { y = +m[1]; mo = +m[2]; d = +m[3]; }
    else { m = /^(\d{1,2})[\/.](\d{1,2})[\/.](\d{4})$/.exec(String(v == null ? "" : v).trim()); if (!m) return null; d = +m[1]; mo = +m[2]; y = +m[3]; }   // dd/mm/yyyy, as in Israel
    var t = Date.UTC(y, mo - 1, d), x = new Date(t);
    if (x.getUTCFullYear() !== y || x.getUTCMonth() !== mo - 1 || x.getUTCDate() !== d) return null;
    return { y: y, m: mo, d: d, serial: Math.round((t - Date.UTC(1899, 11, 30)) / 86400000) };
  }
  // today as an Excel serial, in the computer's own time zone (what TODAY() gives)
  function todaySerial(now) { var d = now || new Date(); return Math.round((Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) - Date.UTC(1899, 11, 30)) / 86400000); }

  // one condition or group from the AI → { node } or { error }
  function normWhen(raw, headers, depth) {
    if (depth > 4) return { error: "the condition is nested too deep (4 levels at most)" };
    if (Array.isArray(raw)) raw = { all: raw };
    if (!raw || typeof raw !== "object") return { error: "a condition must be an object like {\"col\": \"R\", \"op\": \"is\", \"value\": \"Open\"}" };
    var g = raw.all || raw.and || raw.any || raw.or, kind = raw.all || raw.and ? "all" : raw.any || raw.or ? "any" : null;
    if (kind) {
      if (!Array.isArray(g) || !g.length) return { error: "\"" + kind + "\" needs a list of conditions" };
      var kids = [];
      for (var i = 0; i < g.length; i++) { var k = normWhen(g[i], headers, depth + 1); if (k.error) return k; kids.push(k.node); }
      return { node: kids.length === 1 ? kids[0] : { kind: kind, kids: kids } };
    }
    if (raw.not) { var n = normWhen(raw.not, headers, depth + 1); return n.error ? n : { node: { kind: "not", kid: n.node } }; }
    var cv = raw.col != null ? raw.col : raw.column != null ? raw.column : raw.field;
    if (cv == null) return { error: "a condition has no column (\"col\": \"R\")" };
    var c = columnOf(cv, headers);
    if (c < 0) return { error: "column \"" + cv + "\" isn't in sheet CDF (A to AH)" };
    var opRaw = String(raw.op != null ? raw.op : raw.operator != null ? raw.operator : raw.is != null ? "is" : "").trim(), op = opRaw.toLowerCase().replace(/[\s-]+/g, "_");
    op = OP_ALIASES[op] || op;
    var spec = OPS[op];
    if (!spec) return { error: "\"" + opRaw + "\" isn't an operation the add-in knows (it knows " + Object.keys(OPS).join(", ") + ")" };
    var v = raw.value !== undefined ? raw.value : raw.values !== undefined ? raw.values : raw.is, node = { kind: "cond", col: c, op: op };
    if (spec.need === "text") {
      if (Array.isArray(v) || v == null || typeof v === "object") return { error: letter(c) + " " + op + ": needs one value" };
      node.value = trimX(v);
      if (op !== "is" && op !== "is_not" && node.value === "") return { error: letter(c) + " " + op + ": the value is empty" };
    } else if (spec.need === "list") {
      var list = Array.isArray(v) ? v : String(v == null ? "" : v).split(/\s*[,;|]\s*/);
      list = list.map(function (x) { return trimX(x == null ? "" : x); }).filter(function (x, i, a) { return x !== "" && a.indexOf(x) === i; });
      if (!list.length) return { error: letter(c) + " one_of: needs a list of values" };
      if (list.length > 20) return { error: letter(c) + " one_of: 20 values at most" };
      node.value = list;
    } else if (spec.need === "number" || spec.need === "days") {
      var num = typeof v === "number" ? v : Number(String(v == null ? "" : v).trim());
      if (String(v == null ? "" : v).trim() === "" || !isFinite(num)) return { error: letter(c) + " " + op + ": needs a number" };
      if (spec.need === "days" && (num < 0 || Math.floor(num) !== num || num > 36500)) return { error: letter(c) + " " + op + ": needs a whole number of days" };
      node.value = num;
    } else if (spec.need === "date") {
      var d = isoDate(v);
      if (!d) return { error: letter(c) + " " + op + ": needs a date as YYYY-MM-DD" };
      node.value = d;
    }
    return { node: node };
  }

  // ---- the condition as an Excel formula (rows relative to row 2: $R2) — the same test evalWhen makes here
  function lit(s) { return '"' + String(s).replace(/"/g, '""') + '"'; }
  function numLit(n) { var s = String(n); return /e/i.test(s) ? n.toFixed(10).replace(/0+$/, "").replace(/\.$/, "") : s; }
  function guard(r, test) { return "IF(ISNUMBER(" + r + ")," + test + ",FALSE)"; }
  function exprOf(n) {
    if (n.kind === "all") return "AND(" + n.kids.map(exprOf).join(",") + ")";
    if (n.kind === "any") return "OR(" + n.kids.map(exprOf).join(",") + ")";
    if (n.kind === "not") return "NOT(" + exprOf(n.kid) + ")";
    var r = "$" + letter(n.col) + "2", v = n.value, t = "TRIM(" + r + ")";
    switch (n.op) {
      case "is": return t + "=" + lit(v);
      case "is_not": return t + "<>" + lit(v);
      case "has": return "ISNUMBER(SEARCH(" + lit(String(v).replace(/[~*?]/g, "~$&")) + "," + r + "))";
      case "has_not": return "NOT(ISNUMBER(SEARCH(" + lit(String(v).replace(/[~*?]/g, "~$&")) + "," + r + ")))";
      case "starts": return "LEFT(" + t + "," + v.length + ")=" + lit(v);
      case "empty": return "LEN(" + t + ")=0";
      case "filled": return "LEN(" + t + ")>0";
      case "one_of": return v.length === 1 ? t + "=" + lit(v[0]) : "OR(" + v.map(function (x) { return t + "=" + lit(x); }).join(",") + ")";
      // IF, not AND: AND works out every part, and INT() of a text cell is an error that spoils the whole rule
      case ">": case ">=": case "<": case "<=": return guard(r, r + n.op + numLit(v));
      case "older_than_days": return guard(r, "TODAY()-INT(" + r + ")>" + v);
      case "within_days": return guard(r, "AND(TODAY()-INT(" + r + ")>=0,TODAY()-INT(" + r + ")<=" + v + ")");
      case "before": return guard(r, "INT(" + r + ")<DATE(" + v.y + "," + v.m + "," + v.d + ")");
      case "after": return guard(r, "INT(" + r + ")>DATE(" + v.y + "," + v.m + "," + v.d + ")");
      case "on": return guard(r, "INT(" + r + ")=DATE(" + v.y + "," + v.m + "," + v.d + ")");
    }
    throw new Error("op " + n.op);
  }
  function toFormula(node) { return "=" + exprOf(node); }

  // ---- the same test, here, on the values the add-in read (to count matches before anything is written)
  function cellText(v) { return v == null ? "" : typeof v === "boolean" ? (v ? "TRUE" : "FALSE") : String(v); }
  function evalWhen(n, row, today) {
    if (n.kind === "all") return n.kids.every(function (k) { return evalWhen(k, row, today); });
    if (n.kind === "any") return n.kids.some(function (k) { return evalWhen(k, row, today); });
    if (n.kind === "not") return !evalWhen(n.kid, row, today);
    var raw = (row || [])[n.col], s = cellText(raw), t = trimX(s), v = n.value, isNum = typeof raw === "number" && isFinite(raw);
    switch (n.op) {
      case "is": return lower(t) === lower(v);
      case "is_not": return lower(t) !== lower(v);
      case "has": return lower(s).indexOf(lower(v)) >= 0;
      case "has_not": return lower(s).indexOf(lower(v)) < 0;
      case "starts": return lower(t.slice(0, v.length)) === lower(v);
      case "empty": return t === "";
      case "filled": return t !== "";
      case "one_of": return v.some(function (x) { return lower(t) === lower(x); });
      case ">": return isNum && raw > v;
      case ">=": return isNum && raw >= v;
      case "<": return isNum && raw < v;
      case "<=": return isNum && raw <= v;
      case "older_than_days": return isNum && today - Math.floor(raw) > v;
      case "within_days": return isNum && today - Math.floor(raw) >= 0 && today - Math.floor(raw) <= v;
      case "before": return isNum && Math.floor(raw) < v.serial;
      case "after": return isNum && Math.floor(raw) > v.serial;
      case "on": return isNum && Math.floor(raw) === v.serial;
    }
    return false;
  }
  // rows (1-based) a look paints now
  function matchRows(look, values, today) {
    var out = [], last = lastRow(values);
    for (var r = 1; r < last; r++) if (evalWhen(look.when, values[r], today)) out.push(r + 1);
    return out;
  }
  // the condition in plain words, with the column headers
  function sayWhen(n, headers) {
    if (n.kind === "all") return n.kids.map(function (k) { return sayWhen(k, headers); }).join(" and ");
    if (n.kind === "any") return "(" + n.kids.map(function (k) { return sayWhen(k, headers); }).join(" or ") + ")";
    if (n.kind === "not") return "not (" + sayWhen(n.kid, headers) + ")";
    var h = String((headers || [])[n.col] || "").replace(/\s+/g, " ").trim(), name = letter(n.col) + (h ? " " + h : ""), v = n.value;
    if (n.op === "older_than_days") return name + " is more than " + v + " day" + (v === 1 ? "" : "s") + " ago";
    if (n.op === "within_days") return name + " is within the last " + v + " day" + (v === 1 ? "" : "s");
    if (OPS[n.op].need === "none") return name + " " + OPS[n.op].say;
    if (OPS[n.op].need === "date") return name + " " + OPS[n.op].say + " " + ("0" + v.d).slice(-2) + "/" + ("0" + v.m).slice(-2) + "/" + v.y;
    if (OPS[n.op].need === "list") return name + " is one of " + v.map(function (x) { return "“" + x + "”"; }).join(", ");
    if (OPS[n.op].need === "number") return name + " " + OPS[n.op].say + " " + v;
    return name + " " + OPS[n.op].say + " “" + v + "”";
  }

  // one look from the AI → { look } or { error }
  function normLook(raw, headers, i) {
    var at = "Look " + (i + 1) + ": ";
    if (!raw || typeof raw !== "object" || Array.isArray(raw)) return { error: at + "not an object" };
    var whenRaw = raw.when || raw.where || raw.condition || raw.if || raw.rule;
    if (!whenRaw && (raw.col || raw.column) && (raw.op || raw.operator)) whenRaw = raw;
    if (!whenRaw) return { error: at + "no condition (\"when\")" };
    var w = normWhen(whenRaw, headers, 0);
    if (w.error) return { error: at + w.error };
    var p = raw.paint || raw.style || raw.format || raw;
    var fill = hex(p.fill != null ? p.fill : p.background != null ? p.background : p.bg != null ? p.bg : p.color);
    var font = hex(p.font != null ? p.font : p.text != null ? p.text : p.font_color != null ? p.font_color : p.text_color);
    if (fill === undefined) return { error: at + "the fill color “" + (p.fill || p.background || p.bg || p.color) + "” isn't a color (use #RRGGBB)" };
    if (font === undefined) return { error: at + "the text color “" + (p.font || p.text || p.font_color || p.text_color) + "” isn't a color (use #RRGGBB)" };
    var bold = p.bold === true || p.bold === "true", italic = p.italic === true || p.italic === "true";
    if (!fill && !font && !bold && !italic) return { error: at + "says nothing to paint (fill, font, bold or italic)" };
    var spans = paintColumns(p.cells != null ? p.cells : p.columns != null ? p.columns : raw.cells, headers);
    if (!spans) return { error: at + "the cells to paint are outside A:AH (" + JSON.stringify(p.cells != null ? p.cells : p.columns) + ")" };
    var look = { name: String(raw.name || raw.title || "").trim().slice(0, 60) || "Look " + (i + 1), when: w.node, fill: fill || null, font: font || null, bold: bold, italic: italic, spans: spans, cells: spansText(spans) };
    look.formula = toFormula(w.node);
    if (look.formula.length > MAX_FORMULA) return { error: at + "the rule is " + look.formula.length + " characters as an Excel formula, and Excel takes " + MAX_FORMULA + " — ask for fewer values" };
    look.say = sayWhen(w.node, headers);
    return { look: look };
  }
  // the AI's reply → { looks: [...], error, kind }
  function parseLooks(text, headers, E) {
    var s = String(text || "").replace(/^﻿/, "").replace(/ /g, " ").trim();
    if (!s) return { looks: [], kind: "empty", error: "Nothing was pasted" };
    var fence = /```[a-zA-Z]*[ \t]*\r?\n?([\s\S]*?)```/.exec(s);
    if (fence) s = fence[1].trim();
    var lj = E.looseJson(s);
    if (lj.error) return { looks: [], kind: lj.kind, error: lj.error };
    var d = lj.data, list = Array.isArray(d) ? d : d && Array.isArray(d.looks) ? d.looks : d && Array.isArray(d.rules) ? d.rules : [d];
    if (d && !Array.isArray(d) && Array.isArray(d.rows)) return { looks: [], kind: "rows", error: "This is a reply with new CDF rows — paste it on the NEW ROWS page." };
    if (!list.length) return { looks: [], kind: "none", error: "The reply has no looks in it" };
    var looks = [];
    for (var i = 0; i < list.length; i++) {
      var r = normLook(list[i], headers, i);
      if (r.error) return { looks: [], kind: "invalid", error: r.error };
      looks.push(r.look);
    }
    return { looks: looks, error: null, kind: null };
  }

  // what a saved look keeps in the workbook (and reads back)
  function toStored(look) { return { name: look.name, when: look.when, fill: look.fill, font: look.font, bold: look.bold, italic: look.italic, spans: look.spans, formula: look.formula, addedAt: look.addedAt || null }; }
  function fromStored(x, headers) {
    if (!x || !x.when || !x.spans) return null;
    var look = { name: x.name, when: x.when, fill: x.fill || null, font: x.font || null, bold: !!x.bold, italic: !!x.italic, spans: x.spans, cells: spansText(x.spans), addedAt: x.addedAt || null };
    try { look.formula = toFormula(look.when); look.say = sayWhen(look.when, headers); } catch (e) { return null; }
    return look;
  }

  // ---- the prompt: the request, the columns (names and kinds only — no values), the format of the answer
  // columns: [{ letter, header, kind: text | number | date }]
  function buildPrompt(request, columns) {
    var cols = (columns || []).map(function (c) { return c.letter + " — " + String(c.header || "").replace(/\s+/g, " ").trim() + (c.kind && c.kind !== "text" ? " (" + c.kind + ")" : ""); });
    return [
      "You are helping with the look of an Excel sheet named CDF (aircraft deviation forms, one row per form or per solution document).",
      "Turn the request below into a highlighting rule for that sheet. Do not ask for the file or the data: you get the column names only, and the Excel add-in checks the rule against the sheet itself.",
      "",
      "REQUEST:",
      String(request || "").trim(),
      "",
      "COLUMNS (letter — header, kind):",
      cols.join("\n"),
      "",
      "Known values: column R (CDF status) is Open, Released or Cancelled. Dates are real Excel dates.",
      "",
      "ANSWER with one ```json code block and nothing else, in exactly this format:",
      "```json",
      "{\"cdf_look\": " + LOOK_VERSION + ", \"looks\": [",
      "  {\"name\": \"short name\",",
      "   \"when\": {\"all\": [{\"col\": \"R\", \"op\": \"is\", \"value\": \"Open\"}, {\"col\": \"S\", \"op\": \"has\", \"value\": \"Final Approval\"}]},",
      "   \"paint\": {\"cells\": \"row\", \"fill\": \"#FFC000\", \"font\": \"#000000\", \"bold\": false}}",
      "]}",
      "```",
      "Rules for the JSON:",
      "- \"col\" is a column letter from the list above.",
      "- \"op\" is one of: is, is_not, has (contains), has_not, starts, empty, filled, one_of (value is a list), >, >=, <, <= (numbers),",
      "  older_than_days and within_days (value is a whole number of days, for date columns), before, after, on (value \"YYYY-MM-DD\").",
      "- Text tests ignore upper/lower case and extra spaces.",
      "- Combine conditions with {\"all\": [...]}, {\"any\": [...]} or {\"not\": {...}}.",
      "- \"cells\" is \"row\" (A:AH), a range like \"S:X\", or a list like [\"R\", \"S\"].",
      "- Colors are #RRGGBB. \"fill\" is the cell color, \"font\" the text color; \"bold\" and \"italic\" are true or false.",
      "- Several looks in one answer are fine: one entry each in \"looks\".",
      "- Released rows stay green and Cancelled rows red (the sheet's own rules come first)."
    ].join("\n");
  }
  // column kinds from a sample of the sheet: date when most numbers carry a date format
  function columnKinds(sample, formats, isDateFormat) {
    var out = [];
    for (var c = 0; c <= LAST_COL; c++) {
      var nums = 0, dates = 0, texts = 0;
      for (var r = 0; r < (sample || []).length; r++) {
        var v = (sample[r] || [])[c];
        if (isEmpty(v)) continue;
        if (typeof v === "number") { nums++; if (isDateFormat(((formats || [])[r] || [])[c])) dates++; } else texts++;
      }
      out.push(nums && nums >= texts ? (dates * 2 >= nums ? "date" : "number") : "text");
    }
    return out;
  }

  return {
    LOOK_VERSION: LOOK_VERSION, PURPLE: PURPLE, WHITE: WHITE, BLACK: BLACK, GREEN: GREEN, RED: RED, MAX_FORMULA: MAX_FORMULA, MAX_ADDRESS: MAX_ADDRESS, OPS: OPS,
    lastRow: lastRow, continuationRows: continuationRows, runsOf: runsOf, contAddresses: contAddresses, lookPlan: lookPlan,
    ruleKind: ruleKind, describeExisting: describeExisting, hiddenFormatCells: hiddenFormatCells, visibleRuns: visibleRuns, restoreHeights: restoreHeights,
    hex: hex, columnOf: columnOf, paintColumns: paintColumns, paintAreas: paintAreas, spansText: spansText, isoDate: isoDate, todaySerial: todaySerial,
    normWhen: normWhen, toFormula: toFormula, evalWhen: evalWhen, matchRows: matchRows, sayWhen: sayWhen, normLook: normLook, parseLooks: parseLooks,
    toStored: toStored, fromStored: fromStored, buildPrompt: buildPrompt, columnKinds: columnKinds, letter: letter, col: col
  };
});
