/*
 * CDF Updater 2 — the CDF SUMMARY sheet, read as it is and shown as cards.
 *   parse(values)                 the tables of the sheet: a title row (alone in its column), a header row under it, the data
 *                                 rows, a TOTAL row and summary lines; tables side by side (A, D, G …) are read apart
 *   linkColumn(table, cdf)        which CDF column (and which CDFs: all / open / open at Conversion Engineer) the counts
 *                                 of a two-column table come from — so a click on a row lists the CDFs behind it
 *   keysFor(cdf, link, label)     the CDFs (in that scope) whose rows hold that label in that column
 *   render(parsed, opt)           the HTML of the cards: { html, lists }
 * Nothing is computed again: the numbers are the sheet's own.
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory(require("./cdf-core.js"), require("./cdf-stats.js"));
  else root.CdfSummary = factory(root.CdfCore, root.CdfStats);
})(typeof self !== "undefined" ? self : this, function (K, S) {
  "use strict";

  var KEY_RE = /A330-\d+-\d+(?:\.[A-Z0-9-]+|-)?/gi, IS_KEY = /A330-\d+-\d+/i;
  var LINK_COLS = ["R", "S", "T", "P", "U", "V", "W", "X"];

  function esc(s) { return String(s == null ? "" : s).replace(/[&<>"']/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]; }); }
  function cell(rows, r, c) { var row = rows[r]; return row && c < row.length ? row[c] : ""; }
  function blank(v) { return v == null || (typeof v === "string" && v.trim() === ""); }
  function num(v) {
    if (typeof v === "number") return isFinite(v) ? v : null;
    var s = K.trim(v);
    return /^-?\d+(\.\d+)?$/.test(s) ? Number(s) : null;
  }
  // 56.000000000000014 → 56; 12.345 → 12.35
  function tidy(v) {
    var n = num(v);
    if (n != null) return String(Math.round(n * 100) / 100);
    return K.trim(v);
  }
  function isTitle(v) { var s = K.trim(v); return !!s && typeof v === "string" && num(s) == null && !/[a-z]/.test(s.replace(/\([^)]*\)/g, "")); }

  // a title's "(AS OF 09/11/2026 09:41 Asia/Jerusalem)" → { title, asOf }
  function splitTitle(t) {
    var m = /\s*\(\s*AS OF\s+([^)]*)\)\s*$/i.exec(t);
    return m ? { title: t.slice(0, m.index).trim(), asOf: m[1].trim() } : { title: t, asOf: "" };
  }

  function parse(values) {
    var rows = values || [], tables = [], asOf = "", covered = {};
    var width = rows.reduce(function (w, r) { return Math.max(w, (r || []).length); }, 0);
    for (var r = 0; r < rows.length; r++) {
      for (var c = 0; c < width; c++) {
        var v = cell(rows, r, c);
        if (covered[r + "," + c] || !isTitle(v) || blank(cell(rows, r + 1, c)) || (r > 0 && !blank(cell(rows, r - 1, c)))) continue;
        // a title stands alone: the cell to its right is empty, or another title starts there
        var right = cell(rows, r, c + 1);
        if (!blank(right) && !isTitle(right)) continue;
        // the header: the cells under the title, left to right, up to the first empty one
        var hdr = [], e = c;
        while (e < width && !blank(cell(rows, r + 1, e))) { hdr.push(K.trim(cell(rows, r + 1, e))); e++; }
        var band = hdr.length, t = splitTitle(K.trim(v));
        if (t.asOf && !asOf) asOf = t.asOf;
        var tb = { id: "t" + tables.length, title: t.title, asOf: t.asOf, row: r + 1, col: c, head: hdr, rows: [], total: null, footer: [] };
        for (var i = r + 2; i < rows.length; i++) {
          var vals = [];
          for (var j = c; j < c + band; j++) vals.push(cell(rows, i, j));
          if (vals.every(blank)) break;
          var first = K.trim(vals[0]), mid = vals.slice(1, -1).every(blank), last = vals[band - 1];
          if (/^total$/i.test(first)) { tb.total = tidy(blank(last) ? vals.filter(function (x) { return num(x) != null; })[0] : last); continue; }
          if (band > 2 && mid && !blank(last) && isTitle(first) && !IS_KEY.test(first)) { tb.footer.push([first, tidy(last)]); continue; }
          tb.rows.push(vals.map(tidy));
        }
        for (var cr = r; cr < i; cr++) for (var cc = c; cc < c + band; cc++) covered[cr + "," + cc] = 1;
        tb.kind = band === 2 && tb.rows.length && tb.rows.every(function (x) { return num(x[1]) != null; }) ? "count" : "list";
        tables.push(tb);
      }
    }
    return { tables: tables, asOf: asOf };
  }

  // ---- the links back to sheet CDF ----
  function parts(v) { return K.str(v).split(/\s*[|,\n]\s*/).map(function (x) { return x.trim().toUpperCase(); }).filter(function (x) { return x && x !== "-"; }); }
  function matches(letter, value, label) {
    var L = label.toUpperCase(), p = parts(value);
    if (p.indexOf(L) >= 0) return true;
    if (letter === "P") return S.supplierOf(value).toUpperCase() === L;
    return false;
  }
  function blocks(cdf) { return K.buildBlocks(cdf || []); }
  // the column of sheet CDF whose values match the labels of a count table best (half of them at least)
  var SCOPES = {
    all: function () { return true; },
    open: function (a) { return /^open$/i.test(K.trim(a[K.col("R")])); },
    conv: function (a) { return /^open$/i.test(K.trim(a[K.col("R")])) && /^3\s*-|conversion/i.test(K.trim(a[K.col("S")])); }
  };
  // which column of sheet CDF a count table counts, and over which CDFs (all / open / open at Conversion Engineer):
  // the labels must be found in the column, and the scope whose counts come closest to the sheet's own wins
  // → { col: "S", scope: "open", off: 0 } or null
  function linkColumn(tb, cdf, bl) {
    if (!tb || tb.kind !== "count" || !cdf) return null;
    bl = bl || blocks(cdf);
    var rows = tb.rows.filter(function (x) { return K.trim(x[0]) && num(x[1]) != null; }), some = rows.filter(function (x) { return num(x[1]) > 0; });
    var want = some.reduce(function (t, x) { return t + num(x[1]); }, 0), best = null;
    if (!some.length) return null;
    LINK_COLS.forEach(function (L) {
      var ci = K.col(L), seen = {};
      bl.forEach(function (b) { b.rows.forEach(function (r) { var v = (cdf[r] || [])[ci]; parts(v).forEach(function (x) { seen[x] = 1; }); if (L === "P") seen[S.supplierOf(v).toUpperCase()] = 1; }); });
      if (some.filter(function (x) { return seen[K.trim(x[0]).toUpperCase()]; }).length * 2 < some.length) return;
      Object.keys(SCOPES).forEach(function (sc) {
        var link = { col: L, scope: sc }, off = rows.reduce(function (t, x) { return t + Math.abs(keysFor(cdf, link, K.trim(x[0]), bl).length - num(x[1])); }, 0) / want;
        if (!best || off < best.off - 1e-9) best = { col: L, scope: sc, off: Math.round(off * 1000) / 1000 };
      });
    });
    return best && best.off <= 0.6 ? best : null;
  }
  function keysFor(cdf, link, label, bl) {
    var ci = K.col(link.col), inScope = SCOPES[link.scope] || SCOPES.all, out = [];
    (bl || blocks(cdf)).forEach(function (b) {
      if (!inScope(cdf[b.anchor] || [])) return;
      if (b.rows.some(function (r) { return matches(link.col, (cdf[r] || [])[ci], label); })) out.push(b.key);
    });
    return out;
  }

  // ---- the cards ----
  function keyCell(v) {
    var s = esc(v), any = false;
    s = s.replace(KEY_RE, function (k) { any = true; return '<a class="sv-key" data-key="' + k + '" tabindex="0" role="button">' + k + "</a>"; });
    return any ? s.replace(/\n/g, " ") : s.replace(/\n/g, "<br>");
  }
  function render(p, opt) {
    opt = opt || {};
    var lists = {}, cdf = opt.cdf, bl = cdf ? blocks(cdf) : null, LIM = opt.limit || 12, ex = {};
    (opt.exclude || []).forEach(function (x) { ex[K.keyFull(x)] = 1; });
    if (bl) bl = bl.filter(function (b) { return !ex[K.keyFull(b.key)]; });
    if (!p || !p.tables.length) return { html: '<div class="dempty"><b>No CDF SUMMARY sheet to show.</b> The add-in shows the sheet as it is — run the algorithm so the workbook has it, then refresh.</div>', lists: lists };
    var nav = '<nav class="sv-nav">' + p.tables.map(function (t) { return '<button type="button" class="sv-chip" data-goto="' + t.id + '">' + esc(t.title) + "</button>"; }).join("") + "</nav>";
    var html = p.tables.map(function (t) {
      var sub = (t.total != null ? "<b>" + esc(t.total) + "</b> total" : "") + (t.asOf && t.asOf !== p.asOf ? (t.total != null ? " · " : "") + esc(t.asOf) : "");
      var body;
      if (t.kind === "count") {
        var link = linkColumn(t, cdf, bl), max = t.rows.reduce(function (m, x) { return Math.max(m, num(x[1]) || 0); }, 0) || 1;
        body = '<div class="sv-count">' + t.rows.map(function (x, i) {
          var n = num(x[1]) || 0, id = "", keys = link ? keysFor(cdf, link, K.trim(x[0]), bl) : [];
          if (keys.length) { id = "sum:" + t.id + ":" + i; lists[id] = { title: x[0] + " — " + t.title, keys: keys }; }
          return '<div class="sv-c' + (id ? " tap" : "") + (n ? "" : " zero") + '"' + (id ? ' data-list="' + esc(id) + '" tabindex="0" role="button"' : "") + ' data-text="' + esc((x[0] + " " + x[1]).toLowerCase()) + '">' +
            '<span class="sv-l" dir="auto">' + esc(x[0]) + '</span><span class="sv-b"><i style="width:' + (n ? Math.max(2, Math.round(100 * n / max)) : 0) + '%"></i></span><span class="sv-n">' + esc(x[1]) + "</span></div>";
        }).join("") + "</div>" + (link ? '<p class="sv-src">Counted from column ' + link.col + " of sheet CDF" + (link.scope === "open" ? ", open CDFs" : link.scope === "conv" ? ", open at Conversion Engineer" : "") + " — tap a row for its CDFs.</p>" : "");
      } else {
        body = '<div class="sv-scroll"><table class="sv-t"><thead><tr>' + t.head.map(function (h) { return "<th>" + esc(h) + "</th>"; }).join("") + "</tr></thead><tbody>" +
          t.rows.map(function (x, i) {
            var cont = blank(x[0]);
            return '<tr class="' + (cont ? "cont" : "") + (i >= LIM ? " more" : "") + '" data-text="' + esc(x.join(" ").toLowerCase()) + '">' + x.map(function (v) { return '<td dir="auto"' + (String(v).length > 30 ? ' class="long"' : "") + ">" + keyCell(v) + "</td>"; }).join("") + "</tr>";
          }).join("") + "</tbody></table></div>" +
          (t.rows.length > LIM ? '<button type="button" class="btn small sv-all" data-all="' + t.id + '">Show all ' + t.rows.length + " rows</button>" : "") +
          (t.rows.length ? "" : '<p class="sv-src">No rows.</p>');
      }
      if (t.footer.length) body += '<div class="sv-foot">' + t.footer.map(function (f) { return "<span>" + esc(f[0]) + " <b>" + esc(f[1]) + "</b></span>"; }).join("") + "</div>";
      return '<section class="dcard sv" id="sv-' + t.id + '" data-card="' + t.id + '"><div class="dhead"><h3>' + esc(t.title) + "</h3>" + (sub ? '<span class="dsub">' + sub + "</span>" : "") + "</div>" + body + "</section>";
    }).join("");
    return { html: nav + html, lists: lists };
  }

  return { parse: parse, linkColumn: linkColumn, keysFor: keysFor, render: render, tidy: tidy, splitTitle: splitTitle, isTitle: isTitle };
});
