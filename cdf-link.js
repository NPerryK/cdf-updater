/* CDF link — shared by CDF Updater (Excel) and My Notes (Outlook).
 * The same file is in both add-ins, so both read CDF numbers the same way.
 * Excel → Outlook: the CDF status of every CDF.  Outlook → Excel: your notes that mention a CDF.
 * The data moves by copy and paste, and also through localStorage when both add-ins share it
 * (same address, nperryk.github.io) — then nothing needs to be pasted.
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.CdfLink = factory();
})(this, function () {
  "use strict";
  var V = 1;
  var KIND_STATUS = "cdf-updater/status", KIND_NOTES = "my-notes/cdf-notes";
  var LS_STATUS = "cdfBridge.status.v1", LS_NOTES = "cdfBridge.notes.v1";

  function str(v) { return v == null ? "" : String(v); }
  function ts(x) { return x ? Date.parse(x) || 0 : 0; }
  function pad3(n) { n = String(n); while (n.length < 3) n = "0" + n; return n; }

  // CDF numbers in free text: "A330-1-961.-", "A330-1-740.A", "CDF 740.A", "CDF#961", "CDF-052".
  // Not: dates ("CDF 2026-11-05"), "CDF_List_Export", "CDFs", "A330-300".
  // A list keeps going after the first number: "CDFs: 961, 740.A and 279" (3–4 digit numbers only).
  var RE = /(^|[^A-Za-z0-9])(?:A330-?1-|CDFs?[\s#:_-]*(?:A330-?1-)?)0*(\d{1,4})(?![\d\/]|-\d)(?:\.(?:([A-Z]{1,2})(?![A-Za-z0-9])|(-)))?/gi;
  var MORE = /\s*(?:,|&|\/|\band\b|ו-?)\s*0*(\d{3,4})(?![\d\/]|-\d)(?:\.(?:([A-Z]{1,2})(?![A-Za-z0-9])|(-)))?/iy;
  function refs(text) {
    var out = [], seen = {}, m, s = str(text);
    function add(num, r1, r2) {
      var n = Number(num);
      if (!n || n > 1999) return;
      var rev = r1 ? r1.toUpperCase() : r2 ? "-" : "";
      var id = n + "|" + rev;
      if (seen[id]) return;
      seen[id] = true;
      out.push({ n: n, rev: rev, label: label(n, rev) });
    }
    RE.lastIndex = 0;
    while ((m = RE.exec(s))) {
      add(m[2], m[3], m[4]);
      var mm;
      MORE.lastIndex = RE.lastIndex;
      while ((mm = MORE.exec(s))) { add(mm[1], mm[2], mm[3]); RE.lastIndex = MORE.lastIndex; }
    }
    // "CDF 961" and "A330-1-961.-" in one text are the same CDF: keep the more exact one
    return out.filter(function (r) { return r.rev || !out.some(function (o) { return o.n === r.n && o.rev; }); });
  }
  function label(n, rev) { return "A330-1-" + pad3(n) + (rev ? "." + rev : ""); }
  // "A330-1-961.-" → { n: 961, rev: "-" };  "A330-1-052 (another)" → { n: 52, rev: "", extra: "(another)" }
  function keyParts(key) {
    var m = /A330-?1-0*(\d{1,4})(?:\.([A-Z]{1,2}|-))?\s*(.*)$/i.exec(str(key).trim());
    return m ? { n: Number(m[1]), rev: m[2] ? m[2].toUpperCase() : "", extra: m[3] || "" } : null;
  }
  // a reference without a revision matches every revision of that number
  function matches(ref, key) {
    var k = typeof key === "string" ? keyParts(key) : key;
    if (!k || !ref || k.n !== ref.n) return false;
    return !ref.rev || ref.rev === k.rev;
  }

  // ---------- packing for the clipboard ----------
  function encode(kind, payload) {
    var o = { kind: kind, v: V };
    Object.keys(payload || {}).forEach(function (k) { o[k] = payload[k]; });
    return JSON.stringify(o);
  }
  var NAMES = {};
  NAMES[KIND_STATUS] = "the CDF status from CDF Updater (Excel)";
  NAMES[KIND_NOTES] = "notes from My Notes (Outlook)";
  var PASTE_INTO = {};
  PASTE_INTO[KIND_STATUS] = "My Notes in Outlook";
  PASTE_INTO[KIND_NOTES] = "CDF Updater in Excel";
  // → { ok, data } or { ok: false, error }
  function decode(text, want) {
    var s = str(text).trim();
    if (!s) return { ok: false, error: "Nothing was pasted." };
    var a = s.indexOf("{"), b = s.lastIndexOf("}");
    if (a < 0) return { ok: false, error: "This isn't " + NAMES[want] + ". Copy it again and paste the whole text." };
    if (b < a) return { ok: false, error: "The text was cut off. Copy it again and paste the whole text." };
    var o;
    try { o = JSON.parse(s.slice(a, b + 1)); } catch (e) {
      return { ok: false, error: "The text was cut off or changed (" + (e.message || "not valid") + "). Copy it again and paste the whole text." };
    }
    if (!o || typeof o !== "object" || !o.kind) return { ok: false, error: "This isn't " + NAMES[want] + "." };
    if (o.kind !== want) {
      return { ok: false, error: NAMES[o.kind] ? "This is " + NAMES[o.kind] + " — paste it into " + PASTE_INTO[o.kind] + "." : "This isn't " + NAMES[want] + "." };
    }
    if (Number(o.v) > V) return { ok: false, error: "This was made by a newer version. Update this add-in first." };
    if (want === KIND_STATUS && !Array.isArray(o.rows)) return { ok: false, error: "The CDF status has no rows." };
    if (want === KIND_NOTES && !Array.isArray(o.notes)) return { ok: false, error: "The notes list is missing." };
    return { ok: true, data: o };
  }
  // which package a pasted text is (cheap, before parsing): KIND_STATUS, KIND_NOTES or ""
  function kindOf(text) {
    var m = /"kind"\s*:\s*"([^"]+)"/.exec(str(text).slice(0, 400));
    return m && (m[1] === KIND_STATUS || m[1] === KIND_NOTES) ? m[1] : "";
  }
  // the newer of two packages (pasted vs. shared automatically)
  function newer(a, b) {
    if (!a) return b || null;
    if (!b) return a;
    return ts(b.at) > ts(a.at) ? b : a;
  }
  function readShared(storage, key, kind) {
    try {
      var raw = storage && storage.getItem(key);
      if (!raw) return null;
      var r = decode(raw, kind);
      return r.ok ? r.data : null;
    } catch (e) { return null; }
  }
  function writeShared(storage, key, text) {
    try { storage.setItem(key, text); return true; } catch (e) { return false; }
  }

  // ---------- status rows (Excel → Outlook) ----------
  // since: the day the CDF reached its station (yyyy-mm-dd); exact: 1 when known — when 0, "since" is the day it was
  // received (the most it can be) and "seen" the day CDF Updater first saw it there (the least it can be)
  var COLS = ["key", "status", "station", "signature", "supplier", "dwg", "dwgName", "received", "docs", "ready", "since", "exact", "seen", "notes"];
  function rowObj(r) { var o = {}; COLS.forEach(function (c, i) { o[c] = r[i]; }); return o; }
  // every status row a set of references points to
  function statusFor(refList, data) {
    var out = [], seen = {};
    (refList || []).forEach(function (ref) {
      var hit = false;
      (data && data.rows || []).forEach(function (r) {
        if (!matches(ref, r[0])) return;
        hit = true;
        if (!seen[r[0]]) { seen[r[0]] = true; out.push(rowObj(r)); }
      });
      if (!hit && !seen[ref.label]) { seen[ref.label] = true; out.push({ key: ref.label, missing: true }); }
    });
    return out;
  }
  function ageDays(at, now) {
    var t = ts(at);
    if (!t) return null;
    var a = new Date(t), b = new Date(now || Date.now());
    a.setHours(0, 0, 0, 0); b.setHours(0, 0, 0, 0);
    return Math.round((b - a) / 864e5);
  }

  // ---------- the path of a CDF (same picture in both add-ins) ----------
  var PHASES = [
    { id: "INIT", label: "Initiator" }, { id: "POC", label: "POC" }, { id: "ENG", label: "Engineering" },
    { id: "F/A", label: "Final approval" }, { id: "REL", label: "Released" }
  ];
  // "Open" + "3 - Conversion Engineer" → { index: 2 }; Released → 4; Cancelled → { index: -1, cancelled: true }
  function phaseOf(status, station) {
    var st = str(status).trim().toLowerCase(), m = /^(\d+)/.exec(str(station).trim());
    if (st === "cancelled") return { index: -1, cancelled: true };
    if (st === "released") return { index: 4 };
    if (!m) return { index: /released/i.test(str(station)) ? 4 : -1 };
    var n = Number(m[1]);
    return { index: n <= 1 ? 0 : n === 2 ? 1 : n <= 8 ? 2 : n === 9 ? 3 : 4 };
  }
  var STUCK_DAYS = 14;
  function ymd(d) { return d.getFullYear() + "-" + ("0" + (d.getMonth() + 1)).slice(-2) + "-" + ("0" + d.getDate()).slice(-2); }
  function parseDay(x) {
    var m = /^(\d{4})-(\d{2})-(\d{2})/.exec(str(x)) , m2 = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(str(x).trim());
    if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
    if (m2) return new Date(+m2[3], +m2[2] - 1, +m2[1]);
    return null;
  }
  // days at the current station: { days, min, exact, stuck } — when not exact, days is the most it can be, min the least
  function stationDays(row, now) {
    if (!row || !/^open$/i.test(str(row.status).trim())) return null;
    var exact = !!row.since && !!Number(row.exact);
    var since = parseDay(row.since) || parseDay(row.received);
    if (!since) return null;
    var b = new Date(now || Date.now()); b.setHours(0, 0, 0, 0);
    function d(x) { return Math.max(0, Math.round((b - x) / 864e5)); }
    var days = d(since), seen = exact ? null : parseDay(row.seen), min = exact ? days : seen ? Math.min(d(seen), days) : 0;
    return { days: days, min: min, exact: exact, stuck: (exact ? days : min) >= STUCK_DAYS };
  }

  // ---------- emails that CDF Updater prepared ----------
  // "A330-300 P2F — Open CDFs at Conversion Engineer — ANN/SIE — 06/11/2026" → { kind: "supplier", who: "ANN/SIE" }
  // "A330-300 P2F — CDFs ready for your signature — Dana Ben-Ari — 06/11/2026" → { kind: "sign", who: "Dana Ben-Ari" }
  function sentKind(subject) {
    var s = str(subject).replace(/^\s*((re|fw|fwd|השב|העבר)\s*:\s*)+/i, "").trim(), m;
    if ((m = /^A330-300 P2F\s+[—–-]\s+Open CDFs at Conversion Engineer\s+[—–-]\s+(.+?)\s+[—–-]\s+\d{1,2}\/\d{1,2}\/\d{4}\s*$/i.exec(s))) return { kind: "supplier", who: m[1].trim() };
    if ((m = /^A330-300 P2F\s+[—–-]\s+CDFs ready for your signature\s+[—–-]\s+(.+?)\s+[—–-]\s+\d{1,2}\/\d{1,2}\/\d{4}\s*$/i.exec(s))) return { kind: "sign", who: m[1].trim() };
    if (/^A330-300 P2F\s+[—–-]\s+Daily CDF Status\b/i.test(s)) return { kind: "daily", who: "" };
    return null;
  }
  // Israeli work week: Sunday–Thursday
  function addWorkdays(date, n) {
    var d = new Date(date); d.setHours(0, 0, 0, 0);
    while (n > 0) { d.setDate(d.getDate() + 1); if (d.getDay() !== 5 && d.getDay() !== 6) n--; }
    return ymd(d);
  }

  // ---------- notes (Outlook → Excel) ----------
  // notes: [{ refs: ["A330-1-961.-"], text, cat, state, due, updated, subject, from, received, link }]
  function notesFor(key, notes) {
    var k = keyParts(key);
    if (!k) return [];
    return (notes || []).filter(function (n) {
      return (n.refs || []).some(function (r) { var p = keyParts(r); return p && matches(p, k); });
    }).sort(function (a, b) {
      var od = (a.state === "Done") - (b.state === "Done");
      return od || (b.updated || 0) - (a.updated || 0);
    });
  }
  // open notes whose due date has passed (or is today), by CDF
  function dueNotes(notes, now) {
    var t = new Date(now || Date.now()), today = t.getFullYear() + "-" + ("0" + (t.getMonth() + 1)).slice(-2) + "-" + ("0" + t.getDate()).slice(-2);
    return (notes || []).filter(function (n) { return n.state !== "Done" && n.state !== "Info" && n.due && n.due <= today; })
      .sort(function (a, b) { return a.due < b.due ? -1 : a.due > b.due ? 1 : 0; });
  }

  // open notes whose every CDF is now released or cancelled — time to close them
  // notes: [{ refs: [...labels], state }]; data: a status package
  function closable(notes, data) {
    if (!data) return [];
    return (notes || []).filter(function (n) {
      if (n.state === "Done" || n.state === "Info" || !(n.refs || []).length) return false;
      return n.refs.every(function (r) {
        var p = keyParts(r), rows = p ? statusFor([{ n: p.n, rev: p.rev, label: r }], data) : [];
        return rows.length && rows.every(function (x) { return !x.missing && /^(released|cancelled)$/i.test(str(x.status).trim()); });
      });
    });
  }
  // CDF numbers in notes that are not in the sheet (typos like CDF 9610) → [{ ref, notes: n }]
  function unknownRefs(notes, data) {
    if (!data) return [];
    var count = {};
    (notes || []).forEach(function (n) {
      (n.refs || []).forEach(function (r) {
        var p = keyParts(r);
        if (p && statusFor([{ n: p.n, rev: p.rev, label: r }], data)[0].missing) count[r] = (count[r] || 0) + 1;
      });
    });
    return Object.keys(count).sort().map(function (r) { return { ref: r, notes: count[r] }; });
  }

  return {
    V: V, KIND_STATUS: KIND_STATUS, KIND_NOTES: KIND_NOTES, LS_STATUS: LS_STATUS, LS_NOTES: LS_NOTES, COLS: COLS,
    refs: refs, label: label, keyParts: keyParts, matches: matches,
    encode: encode, decode: decode, kindOf: kindOf, newer: newer, readShared: readShared, writeShared: writeShared,
    statusFor: statusFor, rowObj: rowObj, ageDays: ageDays, notesFor: notesFor, dueNotes: dueNotes,
    PHASES: PHASES, phaseOf: phaseOf, STUCK_DAYS: STUCK_DAYS, stationDays: stationDays, parseDay: parseDay, ymd: ymd,
    sentKind: sentKind, addWorkdays: addWorkdays, closable: closable, unknownRefs: unknownRefs
  };
});
