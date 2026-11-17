/*
 * CDF Updater 2 — the link with My Notes (Outlook) and days at station.
 *   bridgeStatus(cdf, now, from, stations)   the status of every CDF for My Notes (the format My Notes 3.x reads)
 *   trackStations(map, cdf, today)           { key: [station, since, exact, firstSeen] } — kept in the workbook
 *   noteLines(notes) / mergeE(cell, lines)   the Outlook notes of a CDF as lines in column E (Notes): your own text
 *                                            stays; the lines that start with ✉ are the add-in's and are replaced
 *   notesPlan(cdf, notes)                    which E cells change, old → new
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory(require("./cdf-core.js"), require("./cdf-link.js"));
  else root.CdfNotes = factory(root.CdfCore, root.CdfLink);
})(typeof self !== "undefined" ? self : this, function (K, L) {
  "use strict";
  var C = { C: K.col("C"), E: K.col("E"), F: K.col("F"), J: K.col("J"), K: K.col("K"), P: K.col("P"), R: K.col("R"), S: K.col("S"), T: K.col("T"), U: K.col("U"), AA: K.col("AA"), AB: K.col("AB") };
  var MARK = "✉ ";

  function ymd(serial) { var d = K.dateOf(serial); return d.getFullYear() + "-" + ("0" + (d.getMonth() + 1)).slice(-2) + "-" + ("0" + d.getDate()).slice(-2); }
  function v(x) { x = K.trim(x); return x === "-" ? "" : x; }

  function bridgeStatus(cdf, now, from, stations) {
    stations = stations || {};
    var rows = [], blocks = K.buildBlocks(cdf || []);
    blocks.forEach(function (b) {
      var a = cdf[b.anchor] || [], key = K.trim(a[C.C]);
      if (!/A330/i.test(key)) return;
      var docs = [], open = /^open$/i.test(K.trim(a[C.R]));
      if (open) b.rows.slice(0, 8).forEach(function (r) { var row = cdf[r] || [], d = v(row[C.AA]), u = v(row[C.U]); if (d || u) docs.push([d, v(row[C.AB]), u]); });
      var rec = K.daySerial(a[C.F]), s = stations[key];
      rows.push([key, v(a[C.R]), v(a[C.S]), v(a[C.T]), v(a[C.P]), v(a[C.J]), v(a[C.K]), rec != null ? K.ddmmyyyy(rec) : v(a[C.F]), docs, "",
        s ? s[1] : "", s ? s[2] : 0, s ? s[3] || "" : "", stripMine(K.str(a[C.E]))]);
    });
    return { at: (now || new Date()).toISOString(), from: from || "CDF Updater", count: rows.length, rows: rows };
  }

  // days at station: a new station → today (exact); first time seen → since the day it was received (the most it can be)
  function trackStations(map, cdf, today) {
    var out = {}, changed = false, t = ymd(today);
    map = map || {};
    K.buildBlocks(cdf || []).forEach(function (b) {
      var a = cdf[b.anchor] || [], key = K.trim(a[C.C]), st = K.trim(a[C.S]);
      if (!key || !/^open$/i.test(K.trim(a[C.R])) || !st || st === "-") { if (map[key]) changed = true; return; }
      var old = map[key];
      if (old && old[0] === st) { out[key] = old; return; }
      changed = true;
      if (old) out[key] = [st, t, 1, t];
      else { var rec = K.daySerial(a[C.F]); out[key] = [st, rec != null ? ymd(rec) : t, rec != null ? 0 : 1, t]; }
    });
    Object.keys(map).forEach(function (k) { if (!out[k]) changed = true; });
    return { map: out, changed: changed };
  }
  // open CDFs at the same station for more than `days` → [{ key, station, days, exact }]
  function stuck(map, today, days) {
    var t = ymd(today), out = [];
    Object.keys(map || {}).forEach(function (k) {
      // exact: the day it reached the station was seen; otherwise count from the day the add-in first saw it there
      var s = map[k], since = s[2] ? s[1] : s[3] || s[1], d = Math.round((Date.parse(t) - Date.parse(since)) / 86400000);
      if (d > days) out.push({ key: k, station: s[0], days: d, exact: !!s[2] });
    });
    return out.sort(function (a, b) { return b.days - a.days; });
  }

  // ---- notes → column E ----
  function dd(ymdStr) { var m = /^(\d{4})-(\d{2})-(\d{2})/.exec(ymdStr || ""); return m ? m[3] + "/" + m[2] : ""; }
  function noteLines(list) {
    return (list || []).filter(function (n) { return n && K.trim(n.text); }).slice(0, 12).map(function (n) {
      var st = n.state && n.state !== "Open" ? n.state : "", due = n.due ? "due " + dd(n.due) : "";
      var head = [st, due].filter(Boolean).join(", ");
      return MARK + (head ? head + ": " : "") + K.trim(n.text).replace(/\s*\n+\s*/g, " ").slice(0, 300) + (n.subject ? " (" + K.trim(n.subject).slice(0, 80) + ")" : "");
    });
  }
  function stripMine(cell) { return K.str(cell).split(/\r?\n/).filter(function (l) { return l.indexOf(MARK.trim()) !== 0; }).join("\n").replace(/\s+$/, ""); }
  function mergeE(cell, lines) {
    var own = stripMine(cell);
    return [own].concat(lines).filter(function (x) { return x !== ""; }).join("\n");
  }
  // notes: My Notes' list. → [{ row (1-based), key, old, value, n }]
  function notesPlan(cdf, notes) {
    var out = [];
    K.buildBlocks(cdf || []).forEach(function (b) {
      var a = cdf[b.anchor] || [], key = K.trim(a[C.C]);
      if (!/A330/i.test(key)) return;
      var lines = noteLines(L.notesFor(key, notes)), old = K.str(a[C.E]), val = mergeE(old, lines);
      if (val !== old.replace(/\s+$/, "") && !(val === "" && old === "")) out.push({ row: b.anchor + 1, key: key, old: old, value: val, n: lines.length });
    });
    return out;
  }

  return { MARK: MARK, bridgeStatus: bridgeStatus, trackStations: trackStations, stuck: stuck, noteLines: noteLines, stripMine: stripMine, mergeE: mergeE, notesPlan: notesPlan, ymd: ymd };
});
