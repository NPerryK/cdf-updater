/*
 * CDF Updater 2 — charts as inline SVG strings, for the pane and the full-screen present view. Colors come from CSS
 * (classes c-open, c-closed, c-rel, c-can, c-warn …), so light and dark themes both work. A bar or column that stands for
 * a list of CDFs carries data-list="…" and the page opens that list on a click.
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory(require("./cdf-core.js"));
  else root.CdfCharts = factory(root.CdfCore);
})(typeof self !== "undefined" ? self : this, function (K) {
  "use strict";

  function esc(s) { return String(s == null ? "" : s).replace(/[&<>"']/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]; }); }
  function f1(x) { return Math.round(x * 10) / 10; }
  function niceMax(v) { if (v <= 5) return 5; var p = Math.pow(10, Math.floor(Math.log10(v))), m = v / p; return (m <= 1 ? 1 : m <= 2 ? 2 : m <= 5 ? 5 : 10) * p; }

  // horizontal bars: items [{ label, n, list?, cls? }] — label, bar, value; opt { width, max, unit }
  function hbars(items, opt) {
    opt = opt || {};
    var w = opt.width || 320, lw = opt.labelWidth || 118, vw = 34, bh = opt.barHeight || 14, gap = 7, max = opt.max || Math.max.apply(null, items.map(function (i) { return i.n; }).concat([1]));
    var h = items.length * (bh + gap) + 2, bw = w - lw - vw - 6;
    var out = '<svg class="chart hbars" viewBox="0 0 ' + w + " " + h + '" width="100%" role="img" aria-label="' + esc(opt.title || "") + '">';
    items.forEach(function (it, i) {
      var y = i * (bh + gap) + 1, len = it.n ? Math.max(2, Math.round(bw * it.n / max)) : 0, attr = it.list ? ' data-list="' + esc(it.list) + '" tabindex="0" role="button"' : "";
      out += '<g class="hb' + (it.list ? " tap" : "") + '"' + attr + "><title>" + esc(it.label + ": " + it.n) + "</title>" +
        '<text class="t-label" x="' + (lw - 6) + '" y="' + (y + bh - 3) + '" text-anchor="end">' + esc(trunc(it.label, 20)) + "</text>" +
        '<rect class="ch-track" x="' + lw + '" y="' + y + '" width="' + bw + '" height="' + bh + '" rx="2"/>' +
        '<rect class="ch-bar ' + (it.cls || "c-open") + '" x="' + lw + '" y="' + y + '" width="' + len + '" height="' + bh + '" rx="2"/>' +
        '<text class="t-val" x="' + (lw + bw + 6) + '" y="' + (y + bh - 3) + '">' + it.n + "</text></g>";
    });
    return out + "</svg>";
  }
  function trunc(s, n) { s = String(s); return s.length > n ? s.slice(0, n - 1) + "…" : s; }

  // grouped columns per week: weeks [{ week, a, b, aList?, bList? }], two series; avg lines for the full weeks
  function columns2(weeks, opt) {
    opt = opt || {};
    var w = opt.width || 320, h = opt.height || 132, padL = 22, padB = 18, padT = 8, n = weeks.length;
    var max = niceMax(Math.max.apply(null, weeks.map(function (x) { return Math.max(x.a, x.b); }).concat([1])));
    var cw = (w - padL - 4) / n, bw = Math.max(3, Math.min(11, cw / 2 - 2)), y = function (v) { return padT + (h - padT - padB) * (1 - v / max); };
    var out = '<svg class="chart cols" viewBox="0 0 ' + w + " " + h + '" width="100%" role="img" aria-label="' + esc(opt.title || "") + '">';
    [0, max / 2, max].forEach(function (g) { out += '<line class="ch-grid" x1="' + padL + '" x2="' + (w - 2) + '" y1="' + y(g) + '" y2="' + y(g) + '"/><text class="t-axis" x="' + (padL - 4) + '" y="' + (y(g) + 3) + '" text-anchor="end">' + g + "</text>"; });
    weeks.forEach(function (x, i) {
      var cx = padL + cw * i + cw / 2, cur = i === n - 1;
      [["a", -1, opt.aCls || "c-open"], ["b", 1, opt.bCls || "c-closed"]].forEach(function (s) {
        var v = x[s[0]], list = x[s[0] + "List"], top = y(v), attr = list && v ? ' data-list="' + esc(list) + '" tabindex="0" role="button" class="tap"' : "";
        out += "<g" + attr + "><title>" + esc((opt[s[0] + "Name"] || s[0]) + " · week of " + K.ddmm(x.week) + ": " + v) + "</title>" +
          '<rect class="ch-col ' + s[2] + (cur ? " cur" : "") + '" x="' + f1(cx + (s[1] < 0 ? -bw - 1 : 1)) + '" y="' + f1(top) + '" width="' + f1(bw) + '" height="' + f1(Math.max(0, h - padB - top)) + '" rx="1.5"/></g>';
      });
      if (i % 2 === (n - 1) % 2) out += '<text class="t-axis" x="' + f1(cx) + '" y="' + (h - 5) + '" text-anchor="middle">' + K.ddmm(x.week) + "</text>";
    });
    (opt.avg || []).forEach(function (a) { if (a.v > 0) out += '<line class="ch-avg ' + a.cls + '" x1="' + padL + '" x2="' + (w - 2) + '" y1="' + f1(y(a.v)) + '" y2="' + f1(y(a.v)) + '"><title>' + esc(a.label) + "</title></line>"; });
    return out + "</svg>";
  }

  // one series of columns (documents released per week)
  function columns1(weeks, opt) {
    opt = opt || {};
    var w = opt.width || 320, h = opt.height || 96, padL = 22, padB = 18, padT = 8, n = weeks.length;
    var max = niceMax(Math.max.apply(null, weeks.map(function (x) { return x.n; }).concat([1]))), cw = (w - padL - 4) / n, bw = Math.max(4, Math.min(16, cw - 6));
    var y = function (v) { return padT + (h - padT - padB) * (1 - v / max); };
    var out = '<svg class="chart cols" viewBox="0 0 ' + w + " " + h + '" width="100%" role="img" aria-label="' + esc(opt.title || "") + '">';
    [0, max].forEach(function (g) { out += '<line class="ch-grid" x1="' + padL + '" x2="' + (w - 2) + '" y1="' + y(g) + '" y2="' + y(g) + '"/><text class="t-axis" x="' + (padL - 4) + '" y="' + (y(g) + 3) + '" text-anchor="end">' + g + "</text>"; });
    weeks.forEach(function (x, i) {
      var cx = padL + cw * i + cw / 2, top = y(x.n);
      out += "<g><title>" + esc("Week of " + K.ddmm(x.week) + ": " + x.n) + '</title><rect class="ch-col ' + (opt.cls || "c-doc") + (i === n - 1 ? " cur" : "") + '" x="' + f1(cx - bw / 2) + '" y="' + f1(top) + '" width="' + f1(bw) + '" height="' + f1(Math.max(0, h - padB - top)) + '" rx="1.5"/></g>';
      if (i % 2 === (n - 1) % 2) out += '<text class="t-axis" x="' + f1(cx) + '" y="' + (h - 5) + '" text-anchor="middle">' + K.ddmm(x.week) + "</text>";
    });
    return out + "</svg>";
  }

  // a line over days: points [{ d, v }]; opt.target { value, by } draws the goal and the straight path to it
  function line(points, opt) {
    opt = opt || {};
    var w = opt.width || 320, h = opt.height || 96, padL = 26, padB = 16, padT = 8;
    if (points.length < 2) return '<div class="chart-empty">' + esc(opt.empty || "The trend starts once there are two days of data.") + "</div>";
    var x0 = points[0].d, x1 = Math.max(points[points.length - 1].d, opt.target && opt.target.by ? opt.target.by : 0, x0 + 1);
    var vals = points.map(function (p) { return p.v; }).concat(opt.target ? [opt.target.value] : []), lo = Math.min.apply(null, vals), hi = Math.max.apply(null, vals);
    var pad = Math.max(2, Math.round((hi - lo) * 0.15)); lo = Math.max(0, lo - pad); hi = hi + pad;
    var X = function (d) { return padL + (w - padL - 6) * (d - x0) / (x1 - x0); }, Y = function (v) { return padT + (h - padT - padB) * (1 - (v - lo) / Math.max(1, hi - lo)); };
    var path = points.map(function (p, i) { return (i ? "L" : "M") + f1(X(p.d)) + " " + f1(Y(p.v)); }).join(" "), last = points[points.length - 1];
    var out = '<svg class="chart line" viewBox="0 0 ' + w + " " + h + '" width="100%" role="img" aria-label="' + esc(opt.title || "") + '">';
    [lo, hi].forEach(function (g) { out += '<line class="ch-grid" x1="' + padL + '" x2="' + (w - 2) + '" y1="' + f1(Y(g)) + '" y2="' + f1(Y(g)) + '"/><text class="t-axis" x="' + (padL - 4) + '" y="' + f1(Y(g) + 3) + '" text-anchor="end">' + Math.round(g) + "</text>"; });
    if (opt.target) {
      out += '<line class="ch-target" x1="' + padL + '" x2="' + (w - 2) + '" y1="' + f1(Y(opt.target.value)) + '" y2="' + f1(Y(opt.target.value)) + '"><title>Target ' + opt.target.value + "</title></line>";
      if (opt.target.by) out += '<line class="ch-target-path" x1="' + f1(X(last.d)) + '" y1="' + f1(Y(last.v)) + '" x2="' + f1(X(opt.target.by)) + '" y2="' + f1(Y(opt.target.value)) + '"/>';
    }
    out += '<path class="ch-ln ' + (opt.cls || "c-open") + '" d="' + path + '"/><circle class="ch-dot ' + (opt.cls || "c-open") + '" cx="' + f1(X(last.d)) + '" cy="' + f1(Y(last.v)) + '" r="3"/>';
    out += '<text class="t-axis" x="' + padL + '" y="' + (h - 3) + '">' + K.ddmm(x0) + '</text><text class="t-axis" x="' + (w - 2) + '" y="' + (h - 3) + '" text-anchor="end">' + K.ddmm(x1) + "</text>";
    return out + "</svg>";
  }

  // progress ring: pct 0–100
  function ring(pct, opt) {
    opt = opt || {};
    var s = opt.size || 92, r = s / 2 - 8, c = 2 * Math.PI * r, len = c * Math.max(0, Math.min(100, pct)) / 100;
    return '<svg class="chart ring" viewBox="0 0 ' + s + " " + s + '" width="' + s + '" height="' + s + '" role="img" aria-label="' + esc(pct + "% " + (opt.label || "")) + '">' +
      '<circle class="ch-track" cx="' + s / 2 + '" cy="' + s / 2 + '" r="' + r + '" fill="none" stroke-width="9"/>' +
      '<circle class="ch-arc ' + (opt.cls || "c-rel") + '" cx="' + s / 2 + '" cy="' + s / 2 + '" r="' + r + '" fill="none" stroke-width="9" stroke-linecap="round" stroke-dasharray="' + f1(len) + " " + f1(c) + '" transform="rotate(-90 ' + s / 2 + " " + s / 2 + ')"/>' +
      '<text class="t-ring" x="50%" y="' + (s / 2 + 2) + '" text-anchor="middle">' + pct + '%</text><text class="t-ring-sub" x="50%" y="' + (s / 2 + 15) + '" text-anchor="middle">' + esc(opt.label || "") + "</text></svg>";
  }

  // a stacked strip of the status split: [{ n, cls, label }]
  function strip(parts, opt) {
    opt = opt || {};
    var total = parts.reduce(function (s, p) { return s + p.n; }, 0) || 1, x = 0, w = opt.width || 320, h = opt.height || 10;
    return '<svg class="chart strip" viewBox="0 0 ' + w + " " + h + '" width="100%" preserveAspectRatio="none" aria-hidden="true">' + parts.map(function (p) {
      var len = w * p.n / total, s = '<rect class="' + p.cls + '" x="' + f1(x) + '" y="0" width="' + f1(len) + '" height="' + h + '"><title>' + esc(p.label + ": " + p.n) + "</title></rect>"; x += len; return s;
    }).join("") + "</svg>";
  }

  return { hbars: hbars, columns2: columns2, columns1: columns1, line: line, ring: ring, strip: strip, esc: esc, niceMax: niceMax };
});
