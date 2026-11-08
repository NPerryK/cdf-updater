/*
 * CDF algorithm — stage 4: BACKUPDATA history, the release tables, sheet ERROR (algorithm 06, 45, 95, 100.13–100.17).
 * Pure: builds the rows to append and the tables to show; Office.js work is in taskpane.js.
 *
 *   06    BACKUPDATA-v3 — RUNS A:P, CDFS S:AD, SOLUTIONS AG:BA, EVENTS BD:BU, MANTRADOCS BX:CG; appended, never rewritten
 *   06.7  the baseline: the last FINAL of a run that is COMMITTED, PASSED, 0 errors
 *   45    events: CDF_RELEASED, DOCUMENT_RELEASED (once per document–revision, ever), GCC_ANSWER_RECEIVED, REVISION_CHANGE
 *   100.13–100.17  CDFS RELEASED / DOCUMENTS RELEASED — 24 h and 7 days, from events of committed runs
 *   95    sheet ERROR: only what failed in this run; ERROR COUNT = ValidationResult
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.CdfHistory = factory();
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  function str(v) { return v == null ? "" : String(v); }
  function trim(v) { return str(v).replace(/ /g, " ").trim(); }
  function upper(v) { return trim(v).toUpperCase(); }
  function col(L) { var n = 0; for (var i = 0; i < L.length; i++) n = n * 26 + L.charCodeAt(i) - 64; return n - 1; }
  function letter(i) { var n = i + 1, s = ""; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
  function cell(rows, r, c) { var row = rows[r]; return row && c < row.length ? row[c] : ""; }
  function pad(n, w) { n = String(n); while (n.length < (w || 2)) n = "0" + n; return n; }

  var SCHEMA = "BACKUPDATA-v3", TZ = "Asia/Jerusalem";
  var HEADERS = {
    runs: ["RunID", "RunAsOfUTC", "StartedAtUTC", "FinishedAtUTC", "RunStatus", "PreviousCommittedRunID", "TrackingStartedAtUTC", "DisplayTimeZone", "SchemaVersion",
      "StartCDFCount", "FinalCDFCount", "StartSolutionCount", "FinalSolutionCount", "NewEventCount", "ValidationResult", "ValidationStatus"],
    cdfs: ["RunID", "SnapshotType", "ObservedAtUTC", "CDFIdentity", "CDFKey", "CDFNumber", "CDFRevision", "CDFStatus", "ClosureDate", "AnchorRow", "PresentInCDF", "RecordID"],
    solutions: ["RunID", "SnapshotType", "ObservedAtUTC", "CDFIdentity", "CDFKey", "SolutionIdentity", "SolutionType", "DocumentID", "Revision", "MantraStatus",
      "GroupClassification", "SignaturePosition", "Person", "RejectCount", "SourceReference", "MatchType", "SourceOrder", "VerifiedReleaseAtUTC", "ReleaseTimeBasis", "PresentInCDF", "RecordID"],
    events: ["RunID", "EventID", "EventType", "CDFIdentity", "CDFKey", "SolutionIdentity", "DocumentID", "Revision", "PreviousValue", "NewValue",
      "PreviousObservedAtUTC", "DetectedAtUTC", "VerifiedTransitionAtUTC", "EventTimeUTC", "TimeBasis", "PreviousRunID", "Source", "Notes"],
    mantradocs: ["RunID", "SnapshotType", "ObservedAtUTC", "DocumentID", "Revision", "MantraStatus", "StatusDateLocal", "StatusDateUTC", "SourceRow", "RecordID"]
  };
  var AREAS = { runs: "A", cdfs: "S", solutions: "AG", events: "BD", mantradocs: "BX" };   // first column of each area (its RunID)
  var RELEASE_BASIS = "MANTRA DATE STATUS (Asia/Jerusalem→UTC)";
  var C = { A: col("A"), C: col("C"), D: col("D"), Q: col("Q"), R: col("R"), U: col("U"), V: col("V"), W: col("W"), X: col("X"), AA: col("AA"), AB: col("AB") };

  // ---------------------------------------------------------------- time
  function isoUtc(d) {
    d = d instanceof Date ? d : new Date(d);
    return d.getUTCFullYear() + "-" + pad(d.getUTCMonth() + 1) + "-" + pad(d.getUTCDate()) + "T" + pad(d.getUTCHours()) + ":" + pad(d.getUTCMinutes()) + ":" + pad(d.getUTCSeconds()) + "Z";
  }
  function runIdOf(d) { return "RUN-" + isoUtc(d).replace(/[-:TZ]/g, ""); }
  function parseIso(s) { var t = Date.parse(trim(s)); return isNaN(t) ? null : t; }
  // the Israel wall clock of an instant, as a UTC-based number (so the difference is the zone offset)
  // One formatter for every call (making one is the slow part), and the offset kept per UTC hour: since 1950 Israel
  // changes its clock only on whole hours, so within one UTC hour the offset is one number. Same results, far less work.
  var wallFmt = null, wallCtor = null, wallHour = {};
  function wallParts(ms) {
    if (typeof Intl === "undefined") return null;
    if (wallCtor !== Intl.DateTimeFormat) {                     // a new Intl (or a test's stub): start again
      wallCtor = Intl.DateTimeFormat; wallHour = {};
      try { wallFmt = new Intl.DateTimeFormat("en-GB", { timeZone: TZ, year: "numeric", month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit", second: "2-digit", hour12: false }); }
      catch (e) { wallFmt = null; }
    }
    if (!wallFmt) return null;
    var p = {};
    wallFmt.formatToParts(new Date(ms)).forEach(function (x) { p[x.type] = x.value; });
    return p.year ? Date.UTC(+p.year, +p.month - 1, +p.day, p.hour === "24" ? 0 : +p.hour, +p.minute, +p.second) : null;
  }
  var HOURLY_FROM = Date.UTC(1950, 0, 1);
  function jlmWall(ms) {
    try {
      var sec = ms - (((ms % 1000) + 1000) % 1000);               // the formatter drops milliseconds
      if (ms >= HOURLY_FROM) {
        var h = Math.floor(ms / 3600000), off = wallHour[h];
        if (off === undefined) { var w = wallParts(ms); if (wallCtor === Intl.DateTimeFormat && w != null) off = wallHour[h] = w - sec; }
        if (off !== undefined) return sec + off;
      } else {
        var w0 = wallParts(ms);
        if (w0 != null) return w0;
      }
    } catch (e) { /* no time zones here: the Israeli rule below */ }
    var d = new Date(ms), y = d.getUTCFullYear();
    function lastSunday(month) { var x = new Date(Date.UTC(y, month + 1, 0)); return x.getUTCDate() - x.getUTCDay(); }
    var start = Date.UTC(y, 2, lastSunday(2) - 2, 0), end = Date.UTC(y, 9, lastSunday(9), -1);   // Friday before the last Sunday of March 02:00 → last Sunday of October 02:00
    return ms + ((ms >= start && ms < end) ? 3 : 2) * 3600000;
  }
  // an Israel wall time (as a UTC-based number) → the instant
  function jlmToUtc(wallMs) {
    var guess = wallMs - 2 * 3600000, off = jlmWall(guess) - guess, utc = wallMs - off, off2 = jlmWall(utc) - utc;
    return off2 === off ? utc : wallMs - off2;
  }
  // Date Status from Mantra: an Excel serial (Israel time) or "2026-01-26T07:10:18" → { local: "26/01/2026 07:10:18", utc: ISO, ms }
  function mantraTime(v) {
    var wall = null;
    if (typeof v === "number" && isFinite(v) && v > 0) wall = Math.round((v - 25569) * 86400) * 1000;
    else {
      var s = trim(v), m;
      if ((m = /^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2}))?)?$/.exec(s))) wall = Date.UTC(+m[1], +m[2] - 1, +m[3], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0));
      else if ((m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(s))) wall = Date.UTC(+m[3], +m[2] - 1, +m[1], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0));
    }
    if (wall == null) return null;
    var d = new Date(wall), ms = jlmToUtc(wall);
    return { local: pad(d.getUTCDate()) + "/" + pad(d.getUTCMonth() + 1) + "/" + d.getUTCFullYear() + " " + pad(d.getUTCHours()) + ":" + pad(d.getUTCMinutes()) + ":" + pad(d.getUTCSeconds()), utc: isoUtc(ms), ms: ms };
  }
  // an instant → "05/11/2026 06:47" in Israel (the release tables)
  function jlmShort(ms) {
    var d = new Date(jlmWall(ms));
    return pad(d.getUTCDate()) + "/" + pad(d.getUTCMonth() + 1) + "/" + d.getUTCFullYear() + " " + pad(d.getUTCHours()) + ":" + pad(d.getUTCMinutes());
  }

  // ---------------------------------------------------------------- reading BACKUPDATA
  // rows of one area: values of its columns (RunID first); rows without a RunID are left out
  function areaRows(values, area) {
    var c0 = col(AREAS[area]), n = HEADERS[area].length, out = [];
    for (var r = 1; r < (values || []).length; r++) {
      var row = values[r] || [], id = trim(row[c0]);
      if (!id) continue;
      var x = [];
      for (var j = 0; j < n; j++) x.push(row[c0 + j] == null ? "" : row[c0 + j]);
      x.sheetRow = r + 1;
      out.push(x);
    }
    return out;
  }
  // the header row a new or existing sheet must have; → [] when it's right, else what's wrong
  function checkHeaders(headerRow) {
    var bad = [];
    Object.keys(HEADERS).forEach(function (a) {
      var c0 = col(AREAS[a]);
      HEADERS[a].forEach(function (h, j) {
        var have = trim((headerRow || [])[c0 + j]);
        if (have !== h) bad.push(letter(c0 + j) + "1 is \"" + have + "\", not \"" + h + "\"");
      });
    });
    return bad;
  }
  function headerRow() {
    var row = [];
    Object.keys(HEADERS).forEach(function (a) { var c0 = col(AREAS[a]); HEADERS[a].forEach(function (h, j) { while (row.length < c0 + j) row.push(""); row[c0 + j] = h; }); });
    return row;
  }
  // 06.7: the last FINAL of a run that is COMMITTED, PASSED, 0 errors (by RunAsOfUTC, not by row order)
  // has(runId): only a run whose snapshot is still in BACKUPDATA can be the baseline (older runs keep just their RUNS row)
  function baselineRun(runs, has) {
    var best = null;
    (runs || []).forEach(function (r) {
      if (trim(r[4]) !== "COMMITTED" || trim(r[15]) !== "PASSED" || Number(trim(r[14]) || 0) !== 0) return;
      if (has && !has(trim(r[0]))) return;
      var t = parseIso(r[1]);
      if (t == null) return;
      if (!best || t > best.t) best = { t: t, row: r };
    });
    return best ? { runId: trim(best.row[0]), at: trim(best.row[1]), tracking: trim(best.row[6]) || trim(best.row[1]), row: best.row } : null;
  }
  function committedRuns(runs) {
    var s = {};
    (runs || []).forEach(function (r) { if (trim(r[4]) === "COMMITTED") s[trim(r[0])] = true; });
    return s;
  }

  // ---------------- what BACKUPDATA keeps (Niv's decision, replaces 06.9's "append, never delete")
  // The next run needs only: the snapshot of the last committed run (its baseline) and the one before it (the baseline
  // again if the last is undone), and the events of the last days (the 24 h / 7 d tables; a document released within
  // 168 h is still in them, so "released once" holds). RUNS stays whole: one row a run.
  var KEEP_RUNS = 2, EVENT_DAYS = 8, BULK_RUNS = 2;
  // o: { runs, snap: {runId: true} (runs with a FINAL snapshot), rows: {cdfs|solutions|mantradocs: [runId of sheet row 2, 3, …]},
  //      eventRows: [event row of sheet row 2, 3, …], now (ms), current: {runId, commits} }
  // → { keep, drop, segments: {area: [[r1, r2], …] bottom-up}, rows: {area: n}, events: n, bulk, skipped }
  function retention(o) {
    var out = { keep: [], drop: [], segments: { cdfs: [], solutions: [], mantradocs: [], events: [] }, rows: { cdfs: 0, solutions: 0, mantradocs: 0, events: 0 }, bulk: false, skipped: false };
    var cur = o.current || {};
    if (!cur.commits) { out.skipped = true; return out; }       // a failed run writes no snapshot, so nothing moves
    var snap = o.snap || {}, eligible = [];
    (o.runs || []).forEach(function (r) {
      var id = trim(r[0]), t = parseIso(r[1]);
      if (!snap[id] || t == null || trim(r[4]) !== "COMMITTED" || trim(r[15]) !== "PASSED" || Number(trim(r[14]) || 0) !== 0) return;
      eligible.push({ id: id, t: t });
    });
    eligible.sort(function (a, b) { return b.t - a.t; });
    var keep = {}; eligible.slice(0, KEEP_RUNS - 1).forEach(function (x) { keep[x.id] = true; out.keep.push(x.id); });
    var seen = {}, rows = o.rows || {};
    ["cdfs", "solutions", "mantradocs"].forEach(function (a) {
      (rows[a] || []).forEach(function (id) { id = trim(id); if (id && !keep[id] && id !== cur.runId && !seen[id]) { seen[id] = true; out.drop.push(id); } });
    });
    var drop = {}; out.drop.forEach(function (id) { drop[id] = true; });
    function segments(ids, gone) {
      var segs = [], start = -1;
      for (var i = 0; i <= ids.length; i++) {
        var g = i < ids.length && gone(ids[i], i);
        if (g && start < 0) start = i;
        if (!g && start >= 0) { segs.push([start + 2, i + 1]); start = -1; }
      }
      return segs.reverse();                                    // bottom-up: the rows above keep their numbers
    }
    ["cdfs", "solutions", "mantradocs"].forEach(function (a) {
      out.segments[a] = segments(rows[a] || [], function (id) { return !!drop[trim(id)]; });
      out.segments[a].forEach(function (g) { out.rows[a] += g[1] - g[0] + 1; });
    });
    var committed = committedRuns(o.runs), since = (o.now || 0) - EVENT_DAYS * 86400000;
    committed[trim(cur.runId)] = true;
    var ev = o.eventRows || [];
    out.segments.events = segments(ev, function (e) {
      var id = trim(e && e[0]);
      if (!id) return false;
      if (!committed[id]) return true;                           // an undone or failed run's events never count
      var t = Math.max(parseIso(e[11]) || 0, parseIso(e[13]) || 0);
      return t < since;
    });
    out.segments.events.forEach(function (g) { out.rows.events += g[1] - g[0] + 1; });
    out.events = out.rows.events;
    out.bulk = out.drop.length > BULK_RUNS;                     // more than the usual one (or two, after the macro ran) — the first trim
    return out;
  }

  // a number as Excel writes it as text: 15 significant digits ("46174.2763425926")
  function excelNumber(n) { return String(Number(Number(n).toPrecision(15))); }

  // ---------------------------------------------------------------- snapshots (06.4, 06.5, 06.11)
  function cdfSnapshot(cdf, algo, runId, type, at) {
    return algo.buildBlocks(cdf).map(function (b) {
      var q = cell(cdf, b.anchor, C.Q), key = trim(cell(cdf, b.anchor, C.C));
      return [runId, type, at, "K:" + key, key, trim(cell(cdf, b.anchor, C.A)), trim(cell(cdf, b.anchor, C.D)), trim(cell(cdf, b.anchor, C.R)),
        typeof q === "number" ? excelNumber(q) : trim(q), String(b.anchor + 1), "TRUE", runId + "|" + type + "|C|K:" + key];
    });
  }
  function solutionType(AA) { var a = upper(AA); return !a || a === "-" ? "NO DOCUMENT" : a === "ANSWER AT GCC" ? "GCC ANSWER" : "ENGINEERING DOCUMENT"; }
  // the latest Mantra row of a document–revision (its Date Status decides)
  function recordOf(cat, id, rev, M) {
    var d = cat && cat.docs ? cat.docs[M.normId(id)] : null, list = d ? d.revs[M.revShow(rev)] : null;
    if (!list || !list.length) return null;
    return list.slice().sort(function (a, b) { return (b.date || 0) - (a.date || 0); })[0];
  }
  /*
   * one row per solution row of the sheet; ctx: { cat (readMantra), M (CdfMantra), matchOf(key, i, AA) → match or null,
   *   carry: { "SolutionIdentity#order": base row } — a row unchanged since the baseline keeps its match, rejects and release time }
   */
  function solutionSnapshot(cdf, algo, ctx, runId, type, at) {
    var out = [];
    algo.buildBlocks(cdf).forEach(function (b) {
      var key = trim(cell(cdf, b.anchor, C.C)), seen = {};
      b.rows.forEach(function (r, i) {
        var g = function (c) { return trim(cell(cdf, r, c)); };
        var AA = g(C.AA), AB = g(C.AB), U = g(C.U), st = solutionType(AA), eng = st === "ENGINEERING DOCUMENT";
        var sid = "K:" + key + "|" + (eng ? AA : "SOLUTION");
        seen[sid] = (seen[sid] || 0) + 1;
        var order = String(seen[sid]), base = ctx.carry ? ctx.carry[sid + "#" + order] : null;
        var row = [runId, type, at, "K:" + key, key, sid, st, eng ? AA : "", AB, U, g(C.V), g(C.W), g(C.X), "", "CDF", eng ? "" : (st === "GCC ANSWER" ? "ANSWER" : "NONE"), order, "", "", "TRUE",
          runId + "|" + type + "|S|" + sid + "|" + order];
        var same = base && trim(base[7]) === AA && trim(base[8]) === AB && trim(base[9]) === U && trim(base[10]) === row[10] && trim(base[11]) === row[11] && trim(base[12]) === row[12];
        if (eng) {
          // a row unchanged since the baseline keeps the baseline's match, rejects and release time (as the macro does);
          // a new or changed row is resolved now (stage 2's match)
          var fresh = ctx.matchOf ? ctx.matchOf(key, i, AA) : null;
          if (same) {
            row[13] = trim(base[13]); row[15] = trim(base[15]); row[17] = trim(base[17]); row[18] = trim(base[18]);
          } else {
            var match = fresh;
            var rec = ctx.cat ? recordOf(ctx.cat, AA, AB, ctx.M) : null;
            if (!match) match = rec ? "EXACT" : U === "4 - RELEASED" ? "MANTRAREV_EXACT" : U === "NO ENG. DOC." ? "NOT_FOUND_NO_CANDIDATE" : "EXACT";
            row[15] = match;
            if (rec && match !== "MANTRAREV_EXACT" && match !== "NOT_FOUND_NO_CANDIDATE") {
              row[13] = String(rec.rejects || 0);
              if (U === "4 - RELEASED") { var t = mantraTime(rec.date); if (t) { row[17] = t.utc; row[18] = RELEASE_BASIS; } }
            }
          }
        }
        out.push(row);
      });
    });
    return out;
  }
  // every Mantra row as it is (no filter); duplicates kept; the row without a Doc ID is reported, an empty one skipped
  function mantraDocs(values, runId, at) {
    values = values || [];
    var hr = -1, head = [];
    for (var r = 0; r < Math.min(values.length, 30) && hr < 0; r++) {
      var h = (values[r] || []).map(upper);
      if (h.indexOf("DOC ID") >= 0 && h.indexOf("REV") >= 0 && h.indexOf("STATUS") >= 0 && h.indexOf("DATE STATUS") >= 0) { hr = r; head = h; }
    }
    if (hr < 0) return { rows: [], errors: ["Sheet Mantra: the header row was not found"], dateCol: null };
    var ix = { id: head.indexOf("DOC ID"), rev: head.indexOf("REV"), st: head.indexOf("STATUS"), dt: head.indexOf("DATE STATUS") }, rows = [], errors = [];
    for (var i = hr + 1; i < values.length; i++) {
      var row = values[i] || [];
      if (!row.some(function (v) { return trim(v) !== ""; })) continue;
      var id = trim(row[ix.id]);
      if (!id) { errors.push("Sheet Mantra row " + (i + 1) + " has content but no Doc ID"); continue; }
      var rev = trim(row[ix.rev]) || "-", t = mantraTime(row[ix.dt]);
      rows.push([runId, "FINAL", at, id, rev, upper(row[ix.st]), t ? t.local : trim(row[ix.dt]), t ? t.utc : "", String(i + 1), runId + "|FINAL|M|" + id + "|" + rev + "|" + (i + 1)]);
    }
    return { rows: rows, errors: errors, dateCol: letter(ix.dt) };
  }

  // ---------------------------------------------------------------- 45: events
  function pairKey(id, rev) { return upper(id) + "|" + (upper(rev) || "-"); }
  // document–revision → [{ key, sid }] from the solutions of the FINAL sheet (45.4.2)
  function linksOf(solutions) {
    var out = {};
    (solutions || []).forEach(function (s) {
      if (trim(s[6]) !== "ENGINEERING DOCUMENT") return;
      var k = pairKey(s[7], s[8]), key = trim(s[4]);
      out[k] = out[k] || [];
      if (!out[k].some(function (x) { return x.key === key; })) out[k].push({ key: key, sid: "K:" + key + "|" + trim(s[7]) + "|" + (trim(s[8]) || "-") });
    });
    return out;
  }
  // the latest row of each document–revision in a MANTRADOCS snapshot
  function docStates(rows) {
    var out = {};
    (rows || []).forEach(function (r) {
      var k = pairKey(r[3], r[4]), t = parseIso(r[7]) || 0;
      if (!out[k] || t >= out[k].t) out[k] = { t: t, status: upper(r[5]), utc: trim(r[7]), row: trim(r[8]), id: trim(r[3]), rev: trim(r[4]) || "-" };
    });
    return out;
  }
  function preReleased(status, M) { return M.uOf(upper(status)) === "3 - PRE-RELEASE"; }
  /*
   * o: { runId, at (ISO), base: { runId, at, cdfs, solutions, mantradocs } (rows of the baseline FINAL, or null), final: { cdfs, solutions },
   *      mantradocs (this run's FINAL), history (event rows of committed runs), gccRow: key → GCC sheet row, dateCol ("G"), cat, M }
   */
  function detectEvents(o) {
    var M = o.M, out = [], n = 0, base = o.base, at = o.at, atMs = parseIso(at);
    if (!base) return out;                                      // the first run: nothing to compare with
    function add(type, f) {
      n++;
      out.push([o.runId, o.runId + "-E" + pad(n, 4), type, f.cdfId || "-", f.key || "-", f.sid || "-", f.doc || "-", f.rev || "-", f.prev, f.now,
        f.prevAt == null ? base.at : f.prevAt, at, f.verified || "", f.time || at, f.basis || "DETECTED AT", base.runId, f.source, f.notes]);
    }
    var baseCdf = {}; (base.cdfs || []).forEach(function (r) { baseCdf[trim(r[4])] = r; });
    var baseSol = {}; (base.solutions || []).forEach(function (r) { baseSol[trim(r[5]) + "#" + trim(r[16])] = r; });
    var released = {}; (o.history || []).forEach(function (e) { if (trim(e[2]) === "DOCUMENT_RELEASED") released[pairKey(e[6], e[7])] = true; });
    // 45.3 CDF_RELEASED: Open → Released in R
    (o.final.cdfs || []).forEach(function (r) {
      var key = trim(r[4]), was = baseCdf[key];
      if (!was || trim(was[7]) !== "Open" || trim(r[7]) !== "Released") return;
      add("CDF_RELEASED", { cdfId: "K:" + key, key: key, prev: "Open", now: "Released", source: o.gccRow && o.gccRow[key] ? "GCC!J" + o.gccRow[key] : "CDF!R" + trim(r[9]), notes: "Previous value observed in " + base.runId });
    });
    // 45.4.1 DOCUMENT_RELEASED: once per document–revision, ever; by Mantra row
    var links = linksOf(o.final.solutions), before = docStates(base.mantradocs), now = docStates(o.mantradocs), seen = {};
    Object.keys(now).map(function (k) { return [k, now[k]]; }).sort(function (a, b) { return Number(a[1].row) - Number(b[1].row); }).forEach(function (kv) {
      var k = kv[0], d = kv[1];
      if (d.status !== "RELEASED" || released[k] || seen[k]) return;
      // 1. PRE-RELEASE in the baseline and RELEASED now; 2. RELEASED now with Date Status in the last 168 hours
      var b = before[k], observed = !!(b && preReleased(b.status, M)), recent = !!(d.t && atMs && d.t <= atMs && d.t > atMs - 168 * 3600000);
      if (!observed && !recent) return;
      seen[k] = true;
      var l = links[k] || [];
      var notes = observed ? "Observed " + b.status + " (3 - PRE-RELEASE) in previous committed snapshot" : "Previous status per source-history rule (inferred, not observed)";
      if (!l.length) notes += "; UNLINKED – history only";
      add("DOCUMENT_RELEASED", {
        cdfId: l.length ? l.map(function (x) { return "K:" + x.key; }).join("; ") : "-", key: l.length ? l.map(function (x) { return x.key; }).join("; ") : "-",
        sid: l.length ? l.map(function (x) { return x.sid; }).join("; ") : "-", doc: d.id, rev: d.rev, prev: "3 - PRE-RELEASE", now: "4 - RELEASED",
        prevAt: observed ? base.at : "", verified: d.utc, time: d.utc, basis: "VERIFIED TRANSITION TIME", source: "Mantra!" + (o.dateCol || "G") + d.row, notes: notes
      });
    });
    // 45.4 GCC_ANSWER_RECEIVED: NO ENG. DOC. → ANSWER AT GCC in the same solution
    (o.final.solutions || []).forEach(function (s) {
      if (trim(s[6]) !== "GCC ANSWER") return;
      var was = baseSol[trim(s[5]) + "#" + trim(s[16])], key = trim(s[4]);
      if (!was || upper(was[9]) !== "NO ENG. DOC.") return;
      add("GCC_ANSWER_RECEIVED", { cdfId: "K:" + key, key: key, sid: trim(s[5]), doc: "ANSWER AT GCC", rev: "-", prev: "NO ENG. DOC.", now: "ANSWER AT GCC",
        source: o.gccRow && o.gccRow[key] ? "GCC!S" + o.gccRow[key] : "GCC!S", notes: "Previous value observed in " + base.runId });
    });
    // 45.5 REVISION_CHANGE: another revision in the same solution; none when the old one no longer exists (REVISION_REGRESSION)
    (o.final.solutions || []).forEach(function (s) {
      if (trim(s[6]) !== "ENGINEERING DOCUMENT") return;
      var was = baseSol[trim(s[5]) + "#" + trim(s[16])];
      if (!was || trim(was[6]) !== "ENGINEERING DOCUMENT" || upper(was[7]) !== upper(s[7])) return;
      var prev = trim(was[8]) || "-", rev = trim(s[8]) || "-";
      if (upper(prev) === upper(rev)) return;
      var rec = o.cat ? recordOf(o.cat, s[7], rev, M) : null;
      if (o.cat && !recordOf(o.cat, s[7], prev, M) && !(o.revDocs && o.revDocs[pairKey(s[7], prev)])) return;   // REVISION_REGRESSION: internal only
      var key = trim(s[4]);
      add("REVISION_CHANGE", { cdfId: "K:" + key, key: key, sid: "K:" + key + "|" + trim(s[7]) + "|" + prev, doc: trim(s[7]), rev: rev, prev: prev, now: rev,
        source: "Mantra:" + trim(s[7]) + "!row" + (rec ? rec.row : ""), notes: "Highest Mantra revision rule (30.3)" });
    });
    return out;
  }

  // ---------------------------------------------------------------- 100.13–100.17: the release tables
  /*
   * events: every event row of committed runs (and this run's, when it commits); solutions: FINAL of this run (current links);
   * cat: readMantra (the determining record must still be RELEASED); asOf: ms
   */
  function releaseTables(events, solutions, cat, M, asOf) {
    var links = linksOf(solutions);
    var title = function (name) { return name + " (AS OF " + jlmShort(asOf) + " " + TZ + ")"; };
    function within(e, hours) { var t = parseIso(e[13]); return t != null && t <= asOf && t > asOf - hours * 3600000 ? t : null; }
    function cdfTable(name, hours) {
      var last = {};
      (events || []).forEach(function (e) {
        if (trim(e[2]) !== "CDF_RELEASED") return;
        var t = within(e, hours), key = trim(e[4]);
        if (t == null || !key || key === "-") return;
        if (!last[key] || t > last[key].t) last[key] = { t: t, e: e };
      });
      var rows = Object.keys(last).map(function (k) { return last[k]; }).sort(function (a, b) { return b.t - a.t || (a.e[4] < b.e[4] ? -1 : a.e[4] > b.e[4] ? 1 : 0); })
        .map(function (x) { return [trim(x.e[4]), trim(x.e[8]), trim(x.e[9])]; });
      return { title: title(name), headers: ["CDF NUMBER", "PREVIOUS STATUS", "NEW STATUS"], rows: rows };
    }
    var count = {};
    (events || []).forEach(function (e) { if (trim(e[2]) === "DOCUMENT_RELEASED") { var k = pairKey(e[6], e[7]); count[k] = (count[k] || 0) + 1; } });
    function docTable(name, hours) {
      var list = [];
      (events || []).forEach(function (e) {
        var type = trim(e[2]), t = within(e, hours);
        if (t == null) return;
        if (type === "DOCUMENT_RELEASED") {
          var k = pairKey(e[6], e[7]), l = links[k] || [];
          if (count[k] !== 1 || trim(e[14]) !== "VERIFIED TRANSITION TIME" || !l.length) return;
          var rec = cat ? recordOf(cat, e[6], e[7], M) : null;
          if (cat && (!rec || rec.status !== "RELEASED")) return;          // the record that decides is still RELEASED
          list.push({ t: t, k: trim(e[6]), row: [trim(e[6]), trim(e[7]) || "-", l.map(function (x) { return x.key; }).join("\n"), trim(e[8]), trim(e[9]), jlmShort(t)] });
        } else if (type === "GCC_ANSWER_RECEIVED") {
          var key = trim(e[4]);
          if (!key || key === "-" || !trim(e[5]) || trim(e[5]) === "-") return;
          list.push({ t: t, k: key, row: ["ANSWER AT GCC", "-", key, trim(e[8]), trim(e[9]), jlmShort(t)] });
        }
      });
      list.sort(function (a, b) { return b.t - a.t || (a.k < b.k ? -1 : a.k > b.k ? 1 : 0); });
      return { title: title(name), headers: ["DOCUMENT / SOLUTION", "REVISION", "CDF NUMBERS", "PREVIOUS STATUS", "NEW STATUS", "TRANSITION / RELEASE TIME"], rows: list.map(function (x) { return x.row; }) };
    }
    return {
      cdf24: cdfTable("CDFS RELEASED — LAST 24 HOURS", 24), cdf7: cdfTable("CDFS RELEASED — LAST 7 DAYS", 168),
      doc24: docTable("DOCUMENTS RELEASED / GCC ANSWERS — LAST 24 HOURS", 24), doc7: docTable("DOCUMENTS RELEASED / GCC ANSWERS — LAST 7 DAYS", 168)
    };
  }

  // ---------------------------------------------------------------- 06.3: the RUNS row
  function runRow(o) {
    var errs = o.errors || 0;
    return [o.runId, o.at, o.started || o.at, o.finished || o.at, errs ? "FAILED" : "COMMITTED", o.base ? o.base.runId : "", o.base ? o.base.tracking : o.at, TZ, SCHEMA,
      String(o.startCdfs), String(o.finalCdfs), String(o.startSolutions), String(o.finalSolutions), String(o.events), String(errs), errs ? "FAILED" : "PASSED"];
  }

  // ---------------------------------------------------------------- 95: sheet ERROR
  var ERROR_HEADERS = ["RUN ID", "DETECTED AT", "STAGE", "SHEET", "CELL / RANGE", "CDF NUMBER", "DOCUMENT ID", "REVISION", "ERROR CODE", "EXPECTED", "ACTUAL", "ERROR DETAILS", "REQUIRED ACTION", "STATUS", "SEVERITY"];
  var STAGE_ORDER = ["00.1", "00.1.2", "10", "20", "25", "25/30", "30", "110", "100", "130", "06", "45", "95"];
  /*
   * errors: [{ severity (ERROR|BLOCKER), stage, sheet, cell, cdf, doc, rev, code, expected, actual, details, action }]
   * → the rows, deduplicated (95.6) and sorted (95.7)
   */
  function errorRows(errors, runId, detectedAt) {
    var seen = {}, list = [];
    (errors || []).forEach(function (e) {
      var id = [e.stage, e.code, e.cdf || "", e.doc || "", e.rev || "", e.cell || ""].join("|");
      if (seen[id]) return;
      seen[id] = true;
      list.push(e);
    });
    var so = function (s) { var i = STAGE_ORDER.indexOf(trim(s)); return i < 0 ? 99 : i; };
    list.sort(function (a, b) {
      var sa = a.severity === "BLOCKER" ? 0 : 1, sb = b.severity === "BLOCKER" ? 0 : 1;
      return sa - sb || so(a.stage) - so(b.stage) || (str(a.sheet) < str(b.sheet) ? -1 : str(a.sheet) > str(b.sheet) ? 1 : 0) ||
        (str(a.cdf) < str(b.cdf) ? -1 : str(a.cdf) > str(b.cdf) ? 1 : 0) || (str(a.doc) < str(b.doc) ? -1 : str(a.doc) > str(b.doc) ? 1 : 0) || (str(a.cell) < str(b.cell) ? -1 : 1);
    });
    return list.map(function (e) {
      return [runId, detectedAt, str(e.stage), str(e.sheet), str(e.cell), str(e.cdf), str(e.doc), str(e.rev), str(e.code), str(e.expected), str(e.actual), str(e.details), str(e.action), "UNRESOLVED", e.severity === "BLOCKER" ? "BLOCKER" : "ERROR"];
    });
  }
  // the sheet: title A1:O1, headers in row 2, the errors from row 3, then RUN ID / AS OF / ERROR COUNT (95.8)
  function errorLayout(rows, runId, asOfText) {
    var cells = [], r = 1;
    cells.push({ r: 1, c: 0, v: "ALGORITHM RUN ERRORS", role: "title" });
    for (var j = 1; j < 15; j++) cells.push({ r: 1, c: j, v: "", role: "title" });
    ERROR_HEADERS.forEach(function (h, j) { cells.push({ r: 2, c: j, v: h, role: "head" }); });
    rows.forEach(function (row, i) { row.forEach(function (v, j) { cells.push({ r: 3 + i, c: j, v: v, role: "text" }); }); });
    r = rows.length ? 3 + rows.length + 1 : 4;
    var info = [["RUN ID", runId], ["AS OF", asOfText], ["ERROR COUNT", rows.length]];
    info.forEach(function (x, i) { cells.push({ r: r + i, c: 0, v: x[0], role: "label" }); cells.push({ r: r + i, c: 1, v: x[1], role: i === 2 ? "num" : "label" }); });
    var end = r + info.length - 1;
    if (!rows.length) { end++; cells.push({ r: end, c: 0, v: "NO ERRORS FOUND", role: "label" }); }
    return { cells: cells, merges: [{ r1: 1, c1: 0, r2: 1, c2: 14 }], records: rows.length ? { r1: 3, r2: 2 + rows.length } : null, end: end, info: { r1: r, r2: end } };
  }
  // what each stage leaves for sheet ERROR (only what failed; the benign states of 95.4 are not errors)
  function collectErrors(o) {
    var out = [];
    (o.stage1Errors || []).forEach(function (e) {
      out.push({ severity: e.stage === "00.1" ? "BLOCKER" : "ERROR", stage: e.stage || "10", sheet: e.sheet || "CDF", cell: e.cell || "", cdf: e.key || "", code: e.code || (e.stage === "00.1" ? "HEADER_MISMATCH" : "IMPORT_ERROR"),
        expected: e.expected || "", actual: e.actual || "", details: e.message || "", action: e.action || "Fix the source and run again." });
    });
    (o.checks || []).forEach(function (c) {
      if (c.ok) return;
      out.push({ severity: "ERROR", stage: "130", sheet: "CDF SUMMARY", cdf: "", code: "CONTROL_TOTAL_FAILED", expected: "", actual: "", details: c.name + " — " + c.detail, action: "Check the CDF and GCC rows the control total names." });
    });
    (o.stage2Errors || []).forEach(function (e) {
      var m = /^(.*?): (AMBIGUOUS|DRAWING_CONFLICT|[A-Z_]+)?/.exec(e.message || "") || [];
      out.push({ severity: e.stage === "00.1.2" ? "BLOCKER" : "ERROR", stage: e.stage || "25/30", sheet: e.stage === "00.1.2" ? "Mantra" : "CDF", cdf: e.key || "", doc: m[1] && /^\d{3}-/.test(m[1]) ? m[1] : "",
        code: e.stage === "00.1.2" ? "MANTRA_UNREADABLE" : e.stage === "110" ? "ROW_LAYOUT" : (m[2] || "SOLUTION_UNRESOLVED"), details: e.message || "", action: e.stage === "110" ? "Fix the solution rows of the block by hand." : "Decide the document by hand (GCC answer or Mantra) and run again." });
    });
    (o.readBack || []).forEach(function (b) {
      out.push({ severity: "ERROR", stage: "100", sheet: "CDF SUMMARY", cell: b.cell, code: "SUMMARY_TOTAL_MISMATCH", expected: String(b.want), actual: String(b.got), details: "The sheet shows another number than computed", action: "Set calculation to Automatic and run again." });
    });
    (o.historyErrors || []).forEach(function (e) {
      out.push({ severity: e.severity || "ERROR", stage: "06", sheet: "BACKUPDATA", cell: e.cell || "", code: e.code || "BACKUPDATA_STRUCTURE", details: e.message, action: e.action || "Fix sheet BACKUPDATA by hand." });
    });
    var said = {}; (o.stage2Errors || []).forEach(function (e) { said[trim(e.message)] = true; });
    (o.mantraRowErrors || []).forEach(function (m) {
      if (said[trim(m)]) return;                                // stage 2 already stopped on it
      out.push({ severity: "ERROR", stage: "06", sheet: "Mantra", cell: (/row (\d+)/.exec(m) || [])[1] ? "row " + /row (\d+)/.exec(m)[1] : "", code: "MANTRA_ROW_WITHOUT_ID", details: m, action: "Fill in the Doc ID or clear the row in Mantra." });
    });
    return out;
  }

  // ---------------------------------------------------------------- everything for one run
  /*
   * o: { now (Date), started (Date), algo, M (CdfMantra), runs (RUNS rows), base: { cdfs, solutions, mantradocs } rows of the baseline FINAL,
   *      history (EVENTS rows), startCdf, finalCdf, mantra (values), mantraRev (values), gcc (values), matchOf, errors (for sheet ERROR) }
   */
  function buildRun(o) {
    var M = o.M, algo = o.algo, at = isoUtc(o.now), runId = runIdOf(o.now);
    var base = baselineRun(o.runs, o.hasSnapshot || null), committed = committedRuns(o.runs);
    var baseRows = base && o.base ? { runId: base.runId, at: base.at, tracking: base.tracking, cdfs: o.base.cdfs, solutions: o.base.solutions, mantradocs: o.base.mantradocs } : null;
    var cat = o.cat || M.readMantra(o.mantra), rev = o.rev || M.readMantraRev(o.mantraRev || []);   // stage 2's, when given
    var revDocs = {}; Object.keys(rev.docs || {}).forEach(function (k) { rev.docs[k].forEach(function (x) { revDocs[pairKey(x.id, x.rev)] = true; }); });
    var carry = {}; (baseRows ? baseRows.solutions : []).forEach(function (r) { carry[trim(r[5]) + "#" + trim(r[16])] = r; });
    var startC = cdfSnapshot(o.startCdf, algo, runId, "START", at), finalC = cdfSnapshot(o.finalCdf, algo, runId, "FINAL", at);
    var startS = solutionSnapshot(o.startCdf, algo, { type: "START", cat: cat, M: M, carry: carry, matchOf: o.matchOf || null }, runId, "START", at);
    var finalS = solutionSnapshot(o.finalCdf, algo, { type: "FINAL", cat: cat, M: M, carry: carry, matchOf: o.matchOf || null }, runId, "FINAL", at);
    var md = mantraDocs(o.mantra, runId, at);
    var gFull = {}, keyRow = {};                                // CDF → its row in GCC (the event's source cell)
    for (var g = 1; g < (o.gcc || []).length; g++) { var k = trim(o.gcc[g][0]); if (k && gFull[algo.keyFull(k)] == null) gFull[algo.keyFull(k)] = g + 1; }
    finalC.forEach(function (r) { var hit = gFull[algo.keyFull(r[4])]; if (hit) keyRow[r[4]] = hit; });
    var history = (o.history || []).filter(function (e) { return committed[trim(e[0])]; });
    var events = detectEvents({ runId: runId, at: at, base: baseRows, final: { cdfs: finalC, solutions: finalS }, mantradocs: md.rows, history: history, gccRow: keyRow, dateCol: md.dateCol, cat: cat, M: M, revDocs: revDocs });
    var errIn = {}; Object.keys(o.errors || {}).forEach(function (k) { errIn[k] = o.errors[k]; });   // no Object.assign: Office 2016/2019 run IE11
    errIn.mantraRowErrors = md.errors;
    var errorsList = collectErrors(errIn);
    var errRows = errorRows(errorsList, runId, at);
    var commits = errRows.length === 0;
    var release = releaseTables(history.concat(commits ? events : []), finalS, cat, M, o.now.getTime());
    var run = runRow({ runId: runId, at: at, started: o.started ? isoUtc(o.started) : at, finished: at, base: base, startCdfs: startC.length, finalCdfs: finalC.length,
      startSolutions: startS.length, finalSolutions: finalS.length, events: events.length, errors: errRows.length });
    return {
      runId: runId, at: at, base: base, commits: commits, run: run, cdfs: startC.concat(finalC), solutions: startS.concat(finalS), events: events, mantradocs: md.rows,
      release: release, errors: errRows, counts: { cdfs: startC.length + finalC.length, solutions: startS.length + finalS.length, events: events.length, mantradocs: md.rows.length }
    };
  }

  return {
    SCHEMA: SCHEMA, TZ: TZ, HEADERS: HEADERS, AREAS: AREAS, ERROR_HEADERS: ERROR_HEADERS, RELEASE_BASIS: RELEASE_BASIS,
    isoUtc: isoUtc, runIdOf: runIdOf, excelNumber: excelNumber, parseIso: parseIso, jlmWall: jlmWall, jlmToUtc: jlmToUtc, mantraTime: mantraTime, jlmShort: jlmShort,
    areaRows: areaRows, checkHeaders: checkHeaders, headerRow: headerRow, baselineRun: baselineRun, committedRuns: committedRuns,
    retention: retention, KEEP_RUNS: KEEP_RUNS, EVENT_DAYS: EVENT_DAYS,
    cdfSnapshot: cdfSnapshot, solutionSnapshot: solutionSnapshot, solutionType: solutionType, recordOf: recordOf, mantraDocs: mantraDocs,
    pairKey: pairKey, linksOf: linksOf, docStates: docStates, detectEvents: detectEvents, releaseTables: releaseTables, runRow: runRow,
    errorRows: errorRows, errorLayout: errorLayout, collectErrors: collectErrors, buildRun: buildRun, letter: letter, col: col
  };
});
