/*
 * CDF Updater 2 — the numbers behind the dashboard, the present view and the PowerPoint deck.
 * Read straight from the sheets as they are (CDF, GCC, MANTRAREV): no algorithm runs here, nothing is written.
 *
 *   status        Open / Cancelled / Released of every CDF in sheet CDF (its anchor rows), and % released
 *   stations      the open CDFs by current station (CDF!S), in station order
 *   suppliers     the open CDFs at Conversion Engineer by supplier (CDF!P), and all open by supplier
 *   aging         days since each open CDF was opened (GCC Open Date, else CDF received date), in buckets
 *   weekly        CDFs opened and closed in each of the last weeks (GCC Open / Close Date, else CDF F / Q)
 *   pace          closed and opened per week over the last 8 weeks, and when the open ones run out at that pace
 *   docs          the solution documents of the open CDFs: released (in MANTRAREV) or in work
 *   trend         the open count over time, from snapshots the pane keeps in the workbook (one a day)
 * Every count keeps the CDF numbers behind it (keys) so the pane can list them.
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory(require("./cdf-core.js"));
  else root.CdfStats = factory(root.CdfCore);
})(typeof self !== "undefined" ? self : this, function (K) {
  "use strict";

  var C = { C: K.col("C"), E: K.col("E"), F: K.col("F"), P: K.col("P"), Q: K.col("Q"), R: K.col("R"), S: K.col("S"), U: K.col("U"), AA: K.col("AA"), AB: K.col("AB") };
  var AGE_BUCKETS = [[0, 7, "≤ 7 d"], [8, 30, "8–30 d"], [31, 90, "1–3 mo"], [91, 180, "3–6 mo"], [181, 1e9, "> 6 mo"]];
  var SUPPLIERS = { "ANN/SIE": "ANN/SIE", "SIE/ANN": "ANN/SIE", "ANN": "ANN/SIE", "SIE": "ANN/SIE", "ANCRA": "ANCARA", "ANCARA": "ANCARA", "TIMAT": "TIMAT", "IAI": "IAI" };
  var WEEKS = 12, PACE_WEEKS = 8;

  function supplierOf(v) { var s = K.trim(v).toUpperCase(); return SUPPLIERS[s] || (s && s !== "-" ? K.trim(v) : "Unassigned"); }
  // "3 - Conversion Engineer" → 3; "Released" → 100; anything else after the numbered ones
  function stationOrder(name) { var m = /^(\d+)\s*-/.exec(K.trim(name)); return m ? +m[1] : /released/i.test(name) ? 100 : 50; }
  function statusOf(v) { var s = K.trim(v).toLowerCase(); return s === "open" ? "Open" : s === "released" ? "Released" : s === "cancelled" || s === "canceled" ? "Cancelled" : ""; }
  function group(list, keyFn) {
    var by = {}, order = [];
    list.forEach(function (x) { var k = keyFn(x); if (!by[k]) { by[k] = { name: k, n: 0, keys: [] }; order.push(k); } by[k].n++; by[k].keys.push(x.key); });
    return order.map(function (k) { return by[k]; });
  }

  // the CDFs of sheet CDF: one per anchor row (a key in C), with what the dashboard needs
  function cdfsOf(cdf) {
    return K.buildBlocks(cdf || []).map(function (b) {
      var row = cdf[b.anchor] || [], docs = b.rows.map(function (r) { return K.trim((cdf[r] || [])[C.AA]); }).filter(function (x) { return x && x !== "-"; });
      return { key: b.key, row: b.anchor + 1, status: statusOf(row[C.R]), station: K.trim(row[C.S]) || "-", supplier: supplierOf(row[C.P]),
        received: K.daySerial(row[C.F]), closed: K.daySerial(row[C.Q]), mantra: K.trim(row[C.U]), docs: docs.length, notes: K.trim(row[C.E]), blockRows: b.rows };
    });
  }
  // GCC: key base → { open, close, status } (the dates the CDF system keeps)
  function gccIndex(gcc) {
    var out = {};
    if (!gcc || !gcc.length) return out;
    var h = gcc[0] || [], iK = K.columnIndex(h, "CDF Number"), iO = K.columnIndex(h, "Open Date"), iC = K.columnIndex(h, "Close Date"), iS = K.columnIndex(h, "CDF Status");
    if (iK < 0) return out;
    Object.defineProperty(out, "latest", { value: null, writable: true, enumerable: false });
    for (var r = 1; r < gcc.length; r++) {
      var row = gcc[r] || [], k = K.keyBase(row[iK]);
      if (!k) continue;
      var x = out[k] = { open: iO < 0 ? null : K.daySerial(row[iO]), close: iC < 0 ? null : K.daySerial(row[iC]), status: iS < 0 ? "" : statusOf(row[iS]) };
      [x.open, x.close].forEach(function (d) { if (d != null && (out.latest == null || d > out.latest)) out.latest = d; });
    }
    return out;
  }

  // MANTRAREV: the released document revisions (one row per document revision). It only tells released from in work:
  // revIndex → { ids: { DOC: { REV: day } }, latest, rows }   (ids and revisions in upper case)
  function revIndex(rev) {
    if (!rev || !rev.length) return null;
    var hr = K.headerRow(rev, ["Doc ID", "Release Date"]);
    if (hr < 0) return { error: "MANTRAREV has no \"Doc ID\" and \"Release Date\" columns" };
    var h = rev[hr], iId = K.columnIndex(h, "Doc ID"), iRev = K.columnIndex(h, "Rev"), iD = K.columnIndex(h, "Release Date");
    var ids = {}, latest = null, n = 0;
    for (var r = hr + 1; r < rev.length; r++) {
      var row = rev[r] || [], id = K.trim(row[iId]).toUpperCase();
      if (!id) continue;
      var d = K.daySerial(row[iD]), rv = revOf(iRev < 0 ? "" : row[iRev]);
      (ids[id] = ids[id] || {})[rv] = d;
      n++;
      if (d != null && (latest == null || d > latest)) latest = d;
    }
    return { ids: ids, latest: latest, rows: n };
  }
  function revOf(v) { var s = K.trim(v).toUpperCase(); return s === "" ? "-" : s; }
  function revRank(r) { return r === "-" ? 0 : r.length === 1 ? r.charCodeAt(0) - 63 : 100 + r.length; }
  // one solution document of a CDF against MANTRAREV → { state: "released" | "work", released: newest released rev or "" }
  function docState(index, doc, rev) {
    var id = K.trim(doc).toUpperCase(), rv = revOf(rev), revs = index && index.ids ? index.ids[id] : null;
    if (revs && Object.prototype.hasOwnProperty.call(revs, rv)) return { state: "released", released: rv };
    var top = revs ? Object.keys(revs).sort(function (a, b) { return revRank(b) - revRank(a); })[0] : "";
    return { state: "work", released: top || "" };
  }
  function isDoc(v) { var s = K.trim(v); return !!s && s !== "-" && !/answer at gcc|no eng/i.test(s); }
  // the solution documents (AA + AB) of the open CDFs → released / in work, with the CDFs behind each
  function docsOf(cdf, open, index) {
    if (!index || index.error) return index;
    var out = { latest: index.latest, rows: index.rows, released: [], work: [], keys: { released: [], work: [] } }, seen = {};
    open.forEach(function (x) {
      x.blockRows.forEach(function (r) {
        var row = cdf[r] || [], doc = row[C.AA];
        if (!isDoc(doc)) return;
        var rv = revOf(row[C.AB]), id = K.trim(doc), tag = id.toUpperCase() + "|" + rv;
        if (seen[tag]) { if (seen[tag].keys.indexOf(x.key) < 0) seen[tag].keys.push(x.key); return; }
        var s = docState(index, id, rv), d = seen[tag] = { doc: id, rev: rv, released: s.released, keys: [x.key] };
        out[s.state].push(d);
      });
    });
    ["released", "work"].forEach(function (st) {
      var ks = {}; out[st].forEach(function (d) { d.keys.forEach(function (k) { ks[k] = 1; }); });
      out.keys[st] = Object.keys(ks);
      out[st].sort(function (a, b) { return a.doc.localeCompare(b.doc); });
    });
    return out;
  }
  function median(sorted) { var n = sorted.length, m = Math.floor(n / 2); return n % 2 ? sorted[m] : (sorted[m - 1] + sorted[m]) / 2; }
  // the Sundays of the last WEEKS weeks, oldest first, this week last
  function weekList(today, n) { var w = K.weekStart(today), out = []; for (var i = (n || WEEKS) - 1; i >= 0; i--) out.push(w - 7 * i); return out; }

  // ---------- the whole model ----------
  // sheets: { cdf, gcc, rev } values (rows, header first); opt: { now, stuckDays, targets, history: [{d, open, released}] }
  function build(sheets, opt) {
    opt = opt || {};
    var now = opt.now || new Date(), today = K.serialOf(now), cdf = sheets.cdf, gi = gccIndex(sheets.gcc);
    var out = { at: now.toISOString(), today: today, sources: { cdf: !!(cdf && cdf.length > 1), gcc: !!(sheets.gcc && sheets.gcc.length > 1), rev: !!(sheets.rev && sheets.rev.length > 1) } };
    var ri = revIndex(sheets.rev);
    if (!out.sources.cdf) { out.empty = true; return out; }
    var ex = {}; (opt.exclude || []).forEach(function (x) { ex[K.keyFull(x)] = true; });
    var all = cdfsOf(cdf).filter(function (x) { return !ex[K.keyFull(x.key)]; });   // a known duplicate (rules.txt EXCLUDE) isn't counted
    all.forEach(function (x) {
      var g = gi[K.keyBase(x.key)];
      x.opened = g && g.open != null ? g.open : x.received;
      x.closedOn = x.closed != null ? x.closed : g && g.status === "Released" ? g.close : null;
      x.inGcc = !!g;
    });
    var open = all.filter(function (x) { return x.status === "Open"; }), rel = all.filter(function (x) { return x.status === "Released"; }), can = all.filter(function (x) { return x.status === "Cancelled"; });
    var counted = open.length + rel.length + can.length;
    out.status = { open: open.length, released: rel.length, cancelled: can.length, total: counted, other: all.length - counted,
      pct: counted - can.length ? Math.round(1000 * rel.length / (counted - can.length)) / 10 : 0,
      keys: { open: open.map(k), released: rel.map(k), cancelled: can.map(k) } };

    // stations, in station order
    out.stations = group(open, function (x) { return x.station; }).sort(function (a, b) { return stationOrder(a.name) - stationOrder(b.name) || a.name.localeCompare(b.name); });
    // suppliers: at Conversion Engineer (where the supplier owes the answer), and over all open
    var conv = open.filter(function (x) { return /^3\s*-/.test(x.station) || /conversion/i.test(x.station); });
    out.suppliers = { conv: group(conv, function (x) { return x.supplier; }).sort(byN), all: group(open, function (x) { return x.supplier; }).sort(byN), convTotal: conv.length };

    // aging of the open ones
    var aged = open.filter(function (x) { return x.opened != null; }).map(function (x) { return { key: x.key, days: today - x.opened, station: x.station, supplier: x.supplier, row: x.row }; });
    out.aging = {
      buckets: AGE_BUCKETS.map(function (b) { var ks = aged.filter(function (a) { return a.days >= b[0] && a.days <= b[1]; }); return { label: b[2], n: ks.length, keys: ks.map(k) }; }),
      known: aged.length, unknown: open.length - aged.length,
      median: aged.length ? median(aged.map(function (a) { return a.days; }).sort(function (a, b) { return a - b; })) : null,
      oldest: aged.sort(function (a, b) { return b.days - a.days; }).slice(0, 8)
    };
    var stuck = opt.stuckDays || 14;
    out.aging.over = aged.filter(function (a) { return a.days > stuck * 6; }).length;   // half a year at 14 days

    // weekly opened / closed, pace and forecast
    var weeks = weekList(today), wo = {}, wc = {};
    weeks.forEach(function (w) { wo[w] = []; wc[w] = []; });
    all.forEach(function (x) {
      if (x.opened != null && wo[K.weekStart(x.opened)]) wo[K.weekStart(x.opened)].push(x.key);
      if (x.status === "Released" && x.closedOn != null && wc[K.weekStart(x.closedOn)]) wc[K.weekStart(x.closedOn)].push(x.key);
    });
    out.weekly = weeks.map(function (w) { return { week: w, opened: wo[w].length, closed: wc[w].length, openedKeys: wo[w], closedKeys: wc[w] }; });
    var full = out.weekly.slice(-PACE_WEEKS - 1, -1);   // the last full weeks (this week is still running)
    var cAvg = avg(full.map(function (x) { return x.closed; })), oAvg = avg(full.map(function (x) { return x.opened; }));
    out.pace = { weeks: full.length, closed: round1(cAvg), opened: round1(oAvg), net: round1(cAvg - oAvg) };
    if (cAvg - oAvg > 0.05) { var wk = Math.ceil(open.length / (cAvg - oAvg)); out.pace.weeksToZero = wk; out.pace.zeroOn = today + wk * 7; }
    var last7c = all.filter(function (x) { return x.status === "Released" && x.closedOn != null && today - x.closedOn >= 0 && today - x.closedOn < 7; });
    var last7o = all.filter(function (x) { return x.opened != null && today - x.opened >= 0 && today - x.opened < 7; });
    out.week = { closed: last7c.length, opened: last7o.length, closedKeys: last7c.map(k), openedKeys: last7o.map(k) };

    // documents: open CDFs that still wait for an engineering document
    var noDoc = open.filter(function (x) { return /no eng/i.test(x.mantra) || (!x.docs && /^3\s*-/.test(x.station)); });
    out.noDoc = { n: noDoc.length, keys: noDoc.map(k) };
    out.notInGcc = sheets.gcc && sheets.gcc.length > 1 ? all.filter(function (x) { return !x.inGcc && x.status !== "Cancelled"; }).map(k) : [];
    out.docs = docsOf(cdf, open, ri);

    // trend: the stored snapshots + today
    var hist = (opt.history || []).filter(function (s) { return s && s.d < today; }).concat([snapshot(out)]);
    out.trend = hist.slice(-90);
    out.targets = opt.targets || {};
    out.fresh = { gcc: gi.latest == null ? null : gi.latest, rev: out.docs && out.docs.latest != null ? out.docs.latest : null };
    return out;
  }
  function k(x) { return x.key; }
  function byN(a, b) { return b.n - a.n || a.name.localeCompare(b.name); }
  function avg(a) { return a.length ? a.reduce(function (s, x) { return s + x; }, 0) / a.length : 0; }
  function round1(x) { return Math.round(x * 10) / 10; }
  // what the pane keeps for the trend: one per day
  function snapshot(m) {
    var st = {};
    (m.stations || []).forEach(function (s) { st[s.name] = s.n; });
    return { d: m.today, open: m.status ? m.status.open : 0, released: m.status ? m.status.released : 0, st: st };
  }
  // stored list + a new snapshot → the list to store (one per day, the newest wins, 400 days at most)
  function mergeHistory(list, snap) {
    var out = (list || []).filter(function (s) { return s && s.d !== snap.d; }).concat([snap]).sort(function (a, b) { return a.d - b.d; });
    return out.slice(-400);
  }
  // change since a snapshot about a week ago: { open: -3, released: +5, since: serial }
  function deltaWeek(m) {
    var t = m.trend || [], cur = t[t.length - 1];
    if (!cur || t.length < 2) return null;
    var ref = null;
    for (var i = t.length - 2; i >= 0; i--) { ref = t[i]; if (cur.d - t[i].d >= 7) break; }
    return ref ? { open: cur.open - ref.open, released: cur.released - ref.released, since: ref.d } : null;
  }

  return { build: build, cdfsOf: cdfsOf, gccIndex: gccIndex, docsOf: docsOf, revIndex: revIndex, docState: docState, weekList: weekList, snapshot: snapshot, mergeHistory: mergeHistory, deltaWeek: deltaWeek,
    supplierOf: supplierOf, stationOrder: stationOrder, statusOf: statusOf, median: median, AGE_BUCKETS: AGE_BUCKETS, WEEKS: WEEKS, PACE_WEEKS: PACE_WEEKS };
});
