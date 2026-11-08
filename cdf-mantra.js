/*
 * CDF Updater - CDF algorithm, STAGE 2: Mantra -> CDF
 *   00.1.1 / 00.1.2  Mantra mapping and the current document catalog
 *   20.4             decoding the GCC answer (GCC!S) into documents and requested revisions
 *   25               document identity (exact, approved corrections, H -101, automatic, MANTRAREV, not found)
 *   30               determining revision, U (status), V/W/X (group, signature position, person)
 *   30.8 / 40        closing date Q
 *   110.3 / 120      one solution row per document; V from the person
 *
 * Pure logic, no Office.js: works in the task pane (window.CdfMantra) and in Node (tests / replay).
 * Sheets are 2D arrays of values starting at A1 (row 0 = Excel row 1).
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.CdfMantra = factory();
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  function str(v) { return v === null || v === undefined ? "" : String(v); }
  function trim(v) { return str(v).replace(/[   ]/g, " ").trim(); }
  function isEmpty(v) { return trim(v) === ""; }
  function upper(v) { return trim(v).replace(/\s+/g, " ").toUpperCase(); }
  function col(letter) { var n = 0; for (var i = 0; i < letter.length; i++) n = n * 26 + (letter.charCodeAt(i) - 64); return n - 1; }
  function cell(rows, r, c) { var row = rows[r]; return row && c < row.length ? row[c] : ""; }

  // 25.3: normalise for comparison only — trim, NBSP → space, upper case, typographic dashes → "-"
  var normMemo = typeof Map === "function" ? new Map() : null;  // the same ids come back thousands of times: remember them
  function normId(v) {
    if (typeof v !== "string" || !normMemo) return trim(v).replace(/[‐-―−]/g, "-").toUpperCase();
    var hit = normMemo.get(v);
    if (hit === undefined) { hit = trim(v).replace(/[‐-―−]/g, "-").toUpperCase(); if (normMemo.size < 200000) normMemo.set(v, hit); }
    return hit;
  }

  // ---------------------------------------------------------------- 30.2 revision order
  // "-" < A < B < … < Z < AA < AB < …; NEW → "-"; anything else is unknown (reported, never guessed)
  function revKey(rev) {
    var r = upper(rev);
    if (r === "" || r === "-" || r === "NEW") return 0;
    if (!/^[A-Z]{1,3}$/.test(r)) return null;
    var n = 0;
    for (var i = 0; i < r.length; i++) n = n * 26 + (r.charCodeAt(i) - 64);
    return n;
  }
  function revShow(rev) { var r = upper(rev); return r === "" || r === "NEW" ? "-" : r; }

  // ---------------------------------------------------------------- dates
  // Mantra dates come as Excel serial numbers (Office.js) or text (ISO / dd/mm/yyyy hh:mm). Local time = Asia/Jerusalem.
  function toSerial(v) {
    if (typeof v === "number" && isFinite(v)) return v;
    var s = trim(v), m;
    if (!s) return null;
    if ((m = /^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2}))?)?/.exec(s))) return serialOf(+m[1], +m[2], +m[3], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0));
    if ((m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(s))) return serialOf(+m[3], +m[2], +m[1], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0));
    return null;
  }
  function serialOf(y, mo, d, h, mi, se) { return (Date.UTC(y, mo - 1, d, h, mi, se) - Date.UTC(1899, 11, 30)) / 864e5; }
  function dayOf(serial) { return serial == null ? null : Math.floor(serial + 1e-9); }

  // ---------------------------------------------------------------- 00.1.1 / 00.1.2 Mantra catalog
  var MANTRA_HEADERS = { id: "DOC ID", rev: "REV", status: "STATUS", date: "DATE STATUS", rejects: "# REJECTS",
    rejBy: "REJECTED BY EMP", rejFn: "REJECTED BY FUNCTION", pendEmp: "PENDING TO EMP", pendFn: "PENDING TO FUNCTION" };
  function findHeaderRow(values, need) {
    for (var r = 0; r < Math.min(values.length, 10); r++) {
      var row = (values[r] || []).map(upper);
      if (need.every(function (h) { return row.indexOf(h) >= 0; })) return r;
    }
    return -1;
  }
  function readMantra(values) {
    values = values || [];
    var out = { ok: false, errors: [], docs: {}, count: 0, headerRow: -1 };
    var hr = findHeaderRow(values, ["DOC ID", "REV", "STATUS", "DATE STATUS"]);
    if (hr < 0) { out.errors.push("Sheet Mantra: the header row (Doc ID, Rev, Status, Date Status) was not found"); return out; }
    out.headerRow = hr;
    var head = (values[hr] || []).map(upper), idx = {};
    Object.keys(MANTRA_HEADERS).forEach(function (k) { idx[k] = head.indexOf(MANTRA_HEADERS[k]); });
    var missing = Object.keys(idx).filter(function (k) { return idx[k] < 0; }).map(function (k) { return MANTRA_HEADERS[k]; });
    if (missing.length) { out.errors.push("Sheet Mantra: missing columns " + missing.join(", ")); return out; }
    for (var r = hr + 1; r < values.length; r++) {
      var row = values[r] || [];
      if (!row.some(function (v) { return !isEmpty(v); })) continue;              // an empty row is not a record
      var id = trim(row[idx.id]);
      if (!id) { out.errors.push("Sheet Mantra row " + (r + 1) + " has content but no Doc ID"); continue; }   // integrity failure
      var n = normId(id), rec = {
        row: r + 1, id: id, rev: revShow(row[idx.rev]), status: upper(row[idx.status]), date: toSerial(row[idx.date]),
        rejects: Number(trim(row[idx.rejects]) || 0), rejBy: trim(row[idx.rejBy]), rejFn: trim(row[idx.rejFn]),
        pendEmp: trim(row[idx.pendEmp]), pendFn: trim(row[idx.pendFn])
      };
      var d = out.docs[n] || (out.docs[n] = { id: id, norm: n, revs: {} });
      (d.revs[rec.rev] = d.revs[rec.rev] || []).push(rec);
      out.count++;
    }
    out.ok = !out.errors.length && out.count > 0;
    if (!out.count) out.errors.push("Sheet Mantra has no documents");
    return out;
  }
  // MANTRAREV: documented releases (the fallback source) — Doc ID, Rev, Release Date, Function, Approver, Sign Date
  function readMantraRev(values) {
    values = values || [];
    var out = { docs: {}, count: 0, errors: [] };
    var hr = findHeaderRow(values, ["DOC ID", "REV"]);
    if (hr < 0) { if (values.length) out.errors.push("Sheet MANTRAREV: the header row was not found"); return out; }
    var head = (values[hr] || []).map(upper);
    var ix = { id: head.indexOf("DOC ID"), rev: head.indexOf("REV"), rel: head.indexOf("RELEASE DATE"), fn: head.indexOf("FUNCTION"), app: head.indexOf("APPROVER"), sign: head.indexOf("SIGN DATE") };
    for (var r = hr + 1; r < values.length; r++) {
      var row = values[r] || [], id = trim(row[ix.id]);
      if (!id) continue;
      var n = normId(id);
      (out.docs[n] = out.docs[n] || []).push({ row: r + 1, id: id, rev: revShow(row[ix.rev]), release: toSerial(row[ix.rel]),
        fn: upper(ix.fn >= 0 ? row[ix.fn] : ""), approver: upper(ix.app >= 0 ? row[ix.app] : ""), sign: toSerial(ix.sign >= 0 ? row[ix.sign] : "") });
      out.count++;
    }
    return out;
  }

  // ---------------------------------------------------------------- 20.4 decoding the GCC answer
  // a document id: 330-53-14-J2400, 330-00-00-JC136-R01, 330-92-80-H411VU-101 … (25.3)
  var DOC_RE = /(^|[^0-9A-Za-z])(\d{2,3})-(\d{2})-(\d{2})-([A-Z]{1,3}\d{2,5}(?:\.\d{1,3})?[A-Z]{0,3})((?:-R\d{2})?)((?:-\d{3}[A-Z]?)?)(?![0-9A-Za-z])/gi;
  // 20.4.3 fixed corrections of the source text
  var FIXES = [
    { from: "330-53-13-J1.025-R01", to: "330-53-13-J1025-R01" },
    { from: "30-53-14-J1799-001", to: "330-53-14-J1799-001" }
  ];
  // 20.4.8 manual correction for one CDF
  var MANUAL = [{ key: "A330-1-353", from: "330-92-80-JF1313-R02", to: "330-92-80-JF131-R02" }];
  var ANSWER_WORDS = /\b(ACCEPT(ED|ABLE)?|AS IS|ATTACHED FILE|ATTACHED|WIRE|ECP|AMM|ELECTRIC(AL)?|SEE CDF|A330-1-\d+|CDF\s*\d+|USE AS IS|OK|APPROVED|REPAIR|REWORK|INSTALL|REPLACE|REMOVE|PER)\b/i;

  function baseKey(key) { return str(key).replace(/[\s ]+/g, "").toUpperCase().replace(/\.-$/, "").replace(/\.$/, "").replace(/\(.*\)$/, ""); }

  // requested revision right after an id: "Rev. C", "REV "B"", "(REV X)", "Revision D", or a single letter " B"
  function revAfter(text, at) {
    var rest = text.slice(at, at + 40), m;
    if ((m = /^\s*[,;]?\s*\(?\s*REV(?:ISION|\.)?\s*[:.]?\s*["'“”]?([A-Z]{1,2}|-|NEW)["'“”]?\s*\)?(?![A-Za-z0-9])/i.exec(rest))) return { rev: revShow(m[1]), len: m[0].length };
    if ((m = /^\s+([A-Z])(?![A-Za-z0-9])/.exec(rest)) && !/^\s+[AI]\s+[a-z]/.test(rest)) return { rev: m[1], len: m[0].length };
    return null;
  }

  // → { kind: "empty" | "answer" | "docs", docs: [{ original, id, requested, itemSuffix }], notes: [] }
  function decodeAnswer(text, cdfKey) {
    var src = str(text), notes = [], applied = [];
    if (isEmpty(src)) return { kind: "empty", docs: [], notes: notes };
    var t = src.replace(/[‐-―−]/g, "-").replace(/(\d)(REV)/gi, "$1 $2");   // "-R01REV A" → "-R01 REV A"
    FIXES.forEach(function (f) {
      var re = new RegExp("(^|[^0-9A-Za-z])" + f.from.replace(/[.]/g, "\\.") + "(?![0-9A-Za-z])", "gi");
      if (re.test(t)) { t = t.replace(re, "$1" + f.to); notes.push("20.4.3 " + f.from + " → " + f.to); applied.push(f.to.toUpperCase()); }
    });
    MANUAL.forEach(function (f) {
      if (baseKey(cdfKey) === f.key && t.toUpperCase().indexOf(f.from) >= 0) { t = t.replace(new RegExp(f.from, "gi"), f.to); notes.push("20.4.8 " + f.from + " → " + f.to); applied.push(f.to.toUpperCase()); }
    });
    var docs = [], m;
    DOC_RE.lastIndex = 0;
    while ((m = DOC_RE.exec(t))) {
      var air = m[2];
      if (air.length !== 3) continue;                  // "30-53-…" without a fix is not an id
      var core = m[5].toUpperCase(), rnn = m[6].toUpperCase(), item = m[7].toUpperCase();
      if (core.indexOf(".") >= 0) continue;           // damaged characters (J1.025) — only the fixed list repairs them
      var full = air + "-" + m[3] + "-" + m[4] + "-" + core + rnn;
      var end = m.index + m[0].length;
      var rq = revAfter(t, end);
      // 20.4.5: a single letter right after the id (also after an item suffix) is the requested revision
      var d = { original: (full + item), id: full, requested: rq ? rq.rev : "", itemSuffix: item.replace(/^-/, ""), at: m.index + m[1].length };
      if (applied.some(function (to) { return (full + item).indexOf(to) >= 0; })) d.fixed = true;   // the id came out of a source fix
      docs.push(d);
    }
    if (docs.length) return { kind: "docs", docs: docs, notes: notes };
    return { kind: "answer", docs: [], notes: notes };   // 20.4.2: a verbal / contractual answer without a document
  }

  // ---------------------------------------------------------------- 25 identity
  // explicit corrections (25.5): CDF base key + missing original → approved id
  var EXPLICIT = [
    ["A330-1-740", "330-92-80-JF015-R01", "330-92-30-JF015-R01"],
    ["A330-1-281", "330-91-00-JC158", "330-00-00-JC158"],
    ["A330-1-281", "330-00-00-JC138-R01", "330-00-00-JC138"],
    ["A330-1-518", "330-91-00-JC154", "330-00-00-JC154"],
    ["A330-1-758", "330-91-00-JC154", "330-00-00-JC154"],
    ["A330-1-854", "330-91-00-JC154", "330-00-00-JC154"],
    ["A330-1-855", "330-91-00-JC154", "330-00-00-JC154"],
    ["A330-1-828", "330-91-00-JC120", "330-00-00-JC120"],
    ["A330-1-544", "330-00-00-JC119", "330-91-00-JC119"],
    ["A330-1-733", "330-53-17-J4251", "330-56-13-J4251"],
    ["A330-1-588", "390-53-14-J2397", "330-53-14-J2397"],
    ["A330-1-156", "330-53-17-J5011-R03", "330-53-17-J5011"],
    ["A330-1-664", "330-92-80-JF120-R02", "330-92-80-JF120"],
    ["A330-1-769", "330-53-14-J8005-R01", "330-53-14-J8005"],
    ["A330-1-789", "330-53-15-J2654-R02", "330-53-15-J2654"],
    ["A330-1-852", "330-92-80-JF103-R03", "330-92-80-JF103-R02"],
    ["A330-1-545", "330-92-80-JF114-R04", "330-92-30-JF114-R04"]
  ];
  var partsMemo = typeof Map === "function" ? new Map() : null;
  function parts(id) {
    var n = normId(id);
    if (partsMemo && partsMemo.has(n)) return partsMemo.get(n);  // read only by the callers: one object per id is safe
    var m = /^(\d{3})-(\d{2})-(\d{2})-([A-Z]{1,3})(\d{3,5})((?:[A-Z]{1,3})?)((?:-R\d{2})?)((?:-\d{3})?)$/.exec(n);
    var p = m ? { air: m[1], ata1: m[2], ata2: m[3], core: m[4] + m[5] + m[6], suffix: m[7].replace(/^-/, ""), item: m[8].replace(/^-/, "") } : null;
    if (partsMemo && partsMemo.size < 200000) partsMemo.set(n, p);
    return p;
  }
  // the catalog's ids, in Object.keys order, read once per catalog (the catalog never changes once read), and the ids
  // grouped by aircraft | core | suffix | item — 25.6's candidates, in the same order a filter over all ids gives
  var keysOf = typeof WeakMap === "function" ? new WeakMap() : null, groupsOf = typeof WeakMap === "function" ? new WeakMap() : null;
  function docKeys(cat) {
    if (!keysOf) return Object.keys(cat.docs);
    var k = keysOf.get(cat.docs);
    if (!k) keysOf.set(cat.docs, (k = Object.keys(cat.docs)));
    return k;
  }
  function groupKey(q) { return q.air + "|" + q.core + "|" + q.suffix + "|" + q.item; }
  function sameCore(cat, p) {
    if (!groupsOf) return docKeys(cat).filter(function (k) { var q = parts(k); return q && groupKey(q) === groupKey(p); });
    var g = groupsOf.get(cat.docs);
    if (!g) {
      g = {};
      docKeys(cat).forEach(function (k) { var q = parts(k); if (q) { var gk = "#" + groupKey(q); (g[gk] = g[gk] || []).push(k); } });
      groupsOf.set(cat.docs, g);
    }
    return (g["#" + groupKey(p)] || []).slice();
  }
  // one edit: replace, insert, delete or swap two neighbours
  function oneEdit(a, b) {
    if (a === b) return false;
    var la = a.length, lb = b.length;
    if (Math.abs(la - lb) > 1) return false;
    if (la === lb) {
      var diff = [];
      for (var i = 0; i < la; i++) if (a[i] !== b[i]) diff.push(i);
      if (diff.length === 1) return true;
      return diff.length === 2 && diff[1] === diff[0] + 1 && a[diff[0]] === b[diff[1]] && a[diff[1]] === b[diff[0]];
    }
    var s = la < lb ? a : b, l = la < lb ? b : a;
    for (var j = 0, k = 0, skipped = false; k < l.length; k++) {
      if (s[j] === l[k]) { j++; continue; }
      if (skipped) return false;
      skipped = true;
    }
    return true;
  }

  // → { id, match, candidates, notes }
  function resolveIdentity(doc, cat, rev, ctx) {
    var o = normId(doc.id), key = baseKey(ctx.key);
    var res = { original: doc.id, id: o, match: "", candidates: [], notes: [] };
    // 25.1: the original exists → never replaced; one the decoder fixed (20.4.3) is an approved manual correction
    if (cat.docs[o]) { res.match = doc.fixed ? "CORRECTED_MANUAL" : "EXACT"; return res; }
    var ex = EXPLICIT.filter(function (e) { return e[0] === key && e[1] === o; })[0];
    if (ex) {
      if (cat.docs[ex[2]]) { res.id = ex[2]; res.match = "CORRECTED_MANUAL"; return res; }
      res.match = "APPROVED_TARGET_MISSING"; res.candidates = [ex[2]]; return res;
    }
    var p = parts(o);
    // approved equivalent: an H number written without its letters ("H5716-101" ↔ "H5716VB-101") — one id only
    if (p && /^H\d+$/.test(p.core)) {
      var pre = p.air + "-" + p.ata1 + "-" + p.ata2 + "-" + p.core, tail = (p.suffix ? "-" + p.suffix : "") + (p.item ? "-" + p.item : "");
      var tailRe = new RegExp("^[A-Z]{1,3}" + tail.replace(/-/g, "\\-") + "$");
      var eq = docKeys(cat).filter(function (k) { return k.indexOf(pre) === 0 && tailRe.test(k.slice(pre.length)); });
      if (eq.length === 1) { res.id = eq[0]; res.match = "EXISTING_EQUIVALENT"; return res; }
      if (eq.length > 1) { res.match = "AMBIGUOUS_EXISTING"; res.candidates = eq.slice(0, 5); return res; }
      if (!p.item) {
        var eqp = docKeys(cat).filter(function (k) { return k.indexOf(pre) === 0 && /^[A-Z]{1,3}-1\d\d$/.test(k.slice(pre.length)); });
        var stems = eqp.map(function (k) { return k.replace(/-1\d\d$/, ""); }).filter(function (x, i, a) { return a.indexOf(x) === i; });
        var with101 = eqp.filter(function (k) { return /-101$/.test(k); });
        if (stems.length === 1 && with101.length === 1) { res.id = with101[0]; res.match = "EXISTING_EQUIVALENT_DEFAULT_101"; return res; }
        if (eqp.length) { res.match = "AMBIGUOUS_EXISTING"; res.candidates = eqp.slice(0, 5); return res; }
      }
    }
    // 25.5.1 H parallels -1xx: base missing → -101 when it exists, else ambiguous
    if (p && /^H/.test(p.core) && !p.item) {
      var par = docKeys(cat).filter(function (k) { return k.indexOf(o + "-1") === 0 && /^-1\d\d$/.test(k.slice(o.length)); });
      if (par.length) {
        if (cat.docs[o + "-101"]) { res.id = o + "-101"; res.match = "EXISTING_EQUIVALENT_DEFAULT_101"; return res; }
        if (par.length === 1) { res.id = par[0]; res.match = "EXISTING_EQUIVALENT"; return res; }
        res.match = "AMBIGUOUS_EXISTING"; res.candidates = par; return res;
      }
    }
    // 20.4.1 item suffix (-001 / -201) is not part of the id: the id without it
    if (p && p.item) {
      var noItem = o.slice(0, o.length - p.item.length - 1);
      var r2 = resolveIdentity({ id: noItem }, cat, rev, ctx);
      r2.original = doc.id;
      if (r2.match) return r2;
    }
    // 25.6 automatic correction: same aircraft, core and suffix; exactly one candidate; one ATA segment one edit away
    if (p) {
      var cands = sameCore(cat, p);                             // same aircraft, core, suffix and item
      if (cands.length > 1) { res.match = "AMBIGUOUS"; res.candidates = cands.slice(0, 5); return res; }
      if (cands.length === 1) {
        var q = parts(cands[0]);
        var d1 = q.ata1 !== p.ata1, d2 = q.ata2 !== p.ata2;
        var okEdit = (d1 !== d2) && (d1 ? oneEdit(p.ata1, q.ata1) : oneEdit(p.ata2, q.ata2));
        // 25.7: the drawing (CDF!J) with the same aircraft and core is evidence for the ATA
        var dw = parts(str(ctx.drawing).replace(/-\d{3}(-[A-Z]{2})?$/i, "").replace(/-[A-Z]{2}$/i, ""));
        var conflict = dw && dw.air === p.air && dw.core.replace(/[A-Z]+$/, "") === p.core.replace(/[A-Z]+$/, "") && (dw.ata1 + dw.ata2) !== (q.ata1 + q.ata2);
        var revConflict = rev && rev.docs[o] && rev.docs[o].length;
        if (okEdit && !conflict && !revConflict) { res.id = cands[0]; res.match = "CORRECTED_AUTO"; return res; }
        res.candidates = cands;
        if (conflict) { res.match = "DRAWING_CONFLICT"; return res; }
      }
    }
    // 25.4(7): the source in MANTRAREV
    if (rev && rev.docs[o] && rev.docs[o].length) { res.match = "MANTRAREV_EXACT"; return res; }
    if (p && p.item) { res.id = o.slice(0, o.length - p.item.length - 1); }
    res.match = "NOT_FOUND_NO_CANDIDATE";                                           // 25.8.1
    return res;
  }

  // ---------------------------------------------------------------- 30 revision, status, signatures
  var U_OF = { "IN-WORK": "2 - IN WORK", "IN-WORK WF": "2 - IN WORK", "IN-WORK REJECTED": "2 - IN WORK",
    "PRE-RELEASED": "3 - PRE-RELEASE", "PRE-RELEASED WF": "3 - PRE-RELEASE", "PRE-RELEASED REJECTED": "2 - IN WORK", "PRE-RELEASED WF DELETED": "3 - PRE-RELEASE",
    "RELEASED": "4 - RELEASED" };
  function uOf(status) {
    if (U_OF[status]) return U_OF[status];
    if (/REJECTED/.test(status)) return "2 - IN WORK";              // 30.5: Rejected → 2 - IN WORK
    if (/RELEASED$/.test(status) && !/PRE-RELEASED/.test(status)) return "4 - RELEASED";
    if (/PRE-RELEASE/.test(status)) return "3 - PRE-RELEASE";
    if (/IN-?WORK|REJECTED/.test(status)) return "2 - IN WORK";
    return null;
  }
  // 30.3: the highest revision in Mantra; 30.3.2: duplicates of one doc–revision → the latest Date Status
  function selectRecord(id, cat) {
    var d = cat.docs[id];
    if (!d) return { error: "not in Mantra" };
    var best = null, bestKey = -1, unknown = [];
    Object.keys(d.revs).forEach(function (rv) {
      var k = revKey(rv);
      if (k == null) { unknown.push(rv); return; }
      if (k > bestKey) { bestKey = k; best = rv; }
    });
    if (best == null) return { error: "revision structure not recognised: " + unknown.join(", ") };
    var recs = d.revs[best].slice().sort(function (a, b) { return (b.date || 0) - (a.date || 0); });
    var conflict = recs.length > 1 && recs[0].date === recs[1].date && (recs[0].status !== recs[1].status || recs[0].pendEmp !== recs[1].pendEmp);
    return { rev: best, record: recs[0], conflict: conflict, unknown: unknown };
  }

  // 120: person → group (V)
  function personKey(name) { return upper(name).replace(/[\t ]+/g, " ").replace(/[‐-―−]/g, "-").replace(/\s*-\s*/g, "-").replace(/\s+/g, " "); }
  function makePeople(cfg) {
    var map = {}, alias = {};
    ((cfg && cfg.personGroups) || []).forEach(function (g) {
      (g.people || []).forEach(function (p) {
        var k = personKey(p);
        if (map[k] && map[k] !== g.group) map[k] = "AMBIGUOUS"; else map[k] = g.group;
      });
    });
    ((cfg && cfg.personAliases) || []).forEach(function (a) { alias[personKey(a.from)] = personKey(a.to); });
    return { map: map, alias: alias, empty: !Object.keys(map).length };
  }
  function groupOf(name, people) {
    if (isEmpty(name) || trim(name) === "-") return "-";
    if (people.empty) return null;                                     // no dictionary: V is left as it is
    var k = personKey(name);
    if (people.alias[k]) k = people.alias[k];
    if (people.map[k]) return people.map[k];
    // "Levi Dana" → "Dana Levi": the same two words in another order (only exact words, no similarity)
    var w = k.split(" ").sort().join(" ");
    var hits = Object.keys(people.map).filter(function (x) { return x.split(" ").sort().join(" ") === w; });
    var groups = hits.map(function (x) { return people.map[x]; }).filter(function (g, i, a) { return a.indexOf(g) === i; });
    return groups.length === 1 ? groups[0] : groups.length > 1 ? "AMBIGUOUS" : "UNMAPPED";
  }
  // 30.5: U, V, W, X of one record
  function uvwx(rec, people) {
    var U = uOf(rec.status);
    if (!U) return { error: "Mantra status \"" + rec.status + "\" is not mapped", U: "", V: "-", W: "-", X: "-" };
    if (U === "4 - RELEASED") return { U: U, V: "-", W: "-", X: "-" };                    // 30.5.1
    if (rec.rejects > 0) {                                                                 // 30.5.2
      if (isEmpty(rec.rejBy) || isEmpty(rec.rejFn)) return { U: U, V: "-", W: "Rejected", X: "-", note: "REJECTED_BY_MISSING_IN_SOURCE" };   // 30.5.2.1
      var rb = uniqList(rec.rejBy).join(", ");
      return { U: U, V: uniqList(rec.rejBy).map(function (e) { return groupOf(e, people); }).join(","), W: "Rejected | " + uniqList(rec.rejFn).join(", "), X: rb };
    }
    // 30.5.3 Pending; several people come comma-separated — the same person twice is one (120.5)
    var emps = uniqList(rec.pendEmp), fns = uniqList(rec.pendFn);
    var X = emps.length ? emps.join(", ") : "-", W = fns.length ? fns.join(", ") : "-";
    var groups = emps.map(function (e) { return groupOf(e, people); });
    if (groups.some(function (g) { return g === null; })) groups = null;
    var V = !emps.length ? "-" : groups ? uniqList(groups.join(",")).join(", ") : null;
    return { U: U, V: V, W: W, X: X };
  }
  function uniqList(v) {
    var seen = {}, out = [];
    str(v).split(",").forEach(function (x) { x = trim(x); var k = personKey(x); if (x && !seen[k]) { seen[k] = true; out.push(x); } });
    return out;
  }

  // ---------------------------------------------------------------- the whole stage for every CDF
  var COLS = { Q: col("Q"), R: col("R"), U: col("U"), V: col("V"), W: col("W"), X: col("X"), AA: col("AA"), AB: col("AB"), J: col("J"), C: col("C"), B: col("B") };
  var GS = col("S");   // GCC!S — the answer (ENGINEERING DISPOSITION)
  var GI = col("I");   // GCC!I — the drawing (compared with CDF!J before closing, 40.1)
  function highestRev(list) {
    var best = null, bk = -1;
    list.forEach(function (x) { var k = revKey(x.rev); if (k != null && k > bk) { bk = k; best = x; } });
    return best;
  }
  // one solution → { AA, AB, U, V, W, X, match, release (serial day), released, error }
  function solveDoc(doc, ctx, cat, rev, people) {
    var r = resolveIdentity(doc, cat, rev, ctx), o = { AA: r.id, AB: "-", U: "", V: "-", W: "-", X: "-", match: r.match, candidates: r.candidates, original: doc.original, requested: doc.requested };
    if (cat.docs[r.id] && /^(EXACT|CORRECTED_MANUAL|CORRECTED_AUTO|EXISTING_EQUIVALENT|EXISTING_EQUIVALENT_DEFAULT_101)$/.test(r.match)) {
      var sel = selectRecord(r.id, cat);
      if (sel.error) { o.U = "CHECK SOLUTION"; o.error = sel.error; return o; }
      var x = uvwx(sel.record, people);
      o.AB = sel.rev; o.U = x.U; o.V = x.V; o.W = x.W; o.X = x.X; o.note = x.note; o.record = sel.record;
      if (x.error) o.error = x.error;
      if (sel.conflict) o.error = "two Mantra rows for " + r.id + " rev " + sel.rev + " with the same Date Status disagree";
      o.released = o.U === "4 - RELEASED";
      o.releaseDay = o.released ? dayOf(sel.record.date) : null;
      return o;
    }
    if (r.match === "MANTRAREV_EXACT") {                       // 30.9: documented in MANTRAREV only — no invented signatures
      var hr = highestRev(rev.docs[normId(doc.id)] || []);
      o.AA = normId(doc.id); o.AB = hr ? hr.rev : "-"; o.U = "4 - RELEASED"; o.released = true;
      o.releaseDay = hr ? dayOf(hr.release) : null;
      return o;
    }
    if (r.match === "NOT_FOUND_NO_CANDIDATE") { o.U = "NO ENG. DOC."; return o; }   // 25.8.1
    o.AA = normId(doc.id); o.U = "CHECK SOLUTION";                                   // 25.8: ambiguity or conflict
    o.error = r.match + (r.candidates && r.candidates.length ? " — candidates: " + r.candidates.slice(0, 5).join(", ") : "");
    return o;
  }

  /*
   * input: { cdf, gcc, mantra, mantraRev, now }  (cdf already updated by stage 1)
   * algo:  CdfAlgo (block map and key matching)
   * → { blocks: [{ key, anchor, rows: [sheet rows], want: [{AA,AB,U,V,W,X}], Q, kind, notes }], changes, structure, errors, info }
   */
  function run(input, cfg, algo) {
    var cdf = input.cdf || [], gcc = input.gcc || [], out = { stage: 2, changes: [], structure: [], errors: [], warnings: [], blocks: [], info: {} };
    var cat = readMantra(input.mantra), rev = readMantraRev(input.mantraRev), people = makePeople(cfg);
    // the catalogs, for stage 4 to use instead of reading Mantra again (not enumerable: the result itself stays as it was)
    try { Object.defineProperty(out, "catalogs", { value: { cat: cat, rev: rev }, enumerable: false }); } catch (e) { /* very old engine */ }
    if (!cat.ok) { cat.errors.forEach(function (e) { out.errors.push({ stage: "00.1.2", message: e }); }); return out; }   // a partial catalog can't say a document is missing
    out.info.catalog = { docs: docKeys(cat).length, rows: cat.count, mantraRev: rev.count };
    var blocks = algo.buildBlocks(cdf), gFull = {}, gBase = {};
    for (var g = 1; g < gcc.length; g++) {
      var k = trim(gcc[g][0]);
      if (!k) continue;
      gFull[algo.keyFull(k)] = g; (gBase[algo.keyBase(k)] = gBase[algo.keyBase(k)] || []).push(g);
    }
    var cBase = {};
    blocks.forEach(function (b) { var kb = algo.keyBase(b.key); cBase[kb] = (cBase[kb] || 0) + 1; });
    function put(b, i, c, v, rule, why) {
      var r = b.rows[i];
      if (r == null) return;
      var old = cell(cdf, r, c);
      if (trim(old) === trim(v) && !(typeof v === "number" && typeof old !== "number" && trim(old) !== "")) return;
      if (typeof v === "number" && typeof old === "number" && Math.floor(old) === Math.floor(v)) return;
      out.changes.push({ sheet: "CDF", row: r + 1, col: c, colLetter: algo.letter(c), key: b.key, oldValue: old, newValue: v, rule: rule, why: why || "" });
    }
    blocks.forEach(function (b) {
      var R = trim(cell(cdf, b.anchor, COLS.R)), gi = gFull[algo.keyFull(b.key)];
      if (gi == null && (gBase[algo.keyBase(b.key)] || []).length === 1 && cBase[algo.keyBase(b.key)] === 1) gi = gBase[algo.keyBase(b.key)][0];
      var blk = { key: b.key, anchor: b.anchor, rows: b.rows.slice(), want: [], kind: "", notes: [] };
      out.blocks.push(blk);
      if (R === "Cancelled") {                                                   // 00.6: S:AB cleared — no Mantra
        blk.kind = "cancelled";
        blk.want = [{ AA: "", AB: "", U: "", V: "", W: "", X: "" }];
      } else if (!R) {                                                           // 00.5: R empty ⇒ U:X empty
        blk.kind = "no-r";
        blk.want = b.rows.map(function (r) { return { AA: trim(cell(cdf, r, COLS.AA)), AB: trim(cell(cdf, r, COLS.AB)), U: "", V: "", W: "", X: "" }; });
      } else if (gi == null) {
        blk.kind = "no-gcc";                                                     // nothing to import: left as it is
        return;
      } else {
        var dec = decodeAnswer(cell(gcc, gi, GS), b.key);
        blk.notes = dec.notes;
        if (dec.kind === "empty") { blk.kind = "empty"; blk.want = [{ AA: "-", AB: "-", U: "NO ENG. DOC.", V: "-", W: "-", X: "-" }]; }   // no answer yet
        else if (dec.kind === "answer") { blk.kind = "answer"; blk.want = [{ AA: "ANSWER AT GCC", AB: "-", U: "ANSWER AT GCC", V: "-", W: "-", X: "-" }]; }
        else {
          blk.kind = "docs";
          var ctx = { key: b.key, drawing: cell(cdf, b.anchor, COLS.J) };
          blk.want = dec.docs.map(function (d) { return solveDoc(d, ctx, cat, rev, people); });
          blk.want.forEach(function (w) { if (w.error) out.errors.push({ stage: "25/30", key: b.key, message: w.original + ": " + w.error }); });
          // 30.8 / 40: Q — R Released (GCC!J), the drawing agrees with GCC (CDF!J = GCC!I), every document found in
          // Mantra and its determining record Released, all on one day. Never from MANTRAREV alone (40.3).
          var days = blk.want.map(function (w) { return w.released && w.match !== "MANTRAREV_EXACT" ? w.releaseDay : undefined; });
          var sameDrawing = upper(cell(cdf, b.anchor, COLS.J)) === upper(cell(gcc, gi, GI));
          if (R === "Released" && days.length && days.every(function (x) { return x != null; })) {
            var uniq = days.filter(function (x, i, a) { return a.indexOf(x) === i; });
            if (!sameDrawing) blk.notes.push("the drawing differs from GCC — Q not set (40.1)");
            else if (uniq.length === 1) blk.Q = uniq[0];
            else blk.notes.push("release days differ (" + uniq.length + ") — Q not set");
          } else if (R === "Released" && blk.want.some(function (w) { return w.match === "MANTRAREV_EXACT"; })) blk.notes.push("a document is only in MANTRAREV — Q not set (40.3)");
        }
      }
      // compare with the sheet, row by row (one row per solution)
      if (blk.want.length !== b.rows.length) {
        out.structure.push({ key: b.key, row: b.anchor + 1, have: b.rows.length, need: blk.want.length });
      }
      var n = Math.min(blk.want.length, b.rows.length);
      for (var i = 0; i < n; i++) {
        var w = blk.want[i], rule = blk.kind === "docs" ? (w.match || "30") : blk.kind === "cancelled" ? "00.6" : blk.kind === "no-r" ? "00.5" : blk.kind === "answer" ? "20.4.2" : "120.9";
        put(b, i, COLS.AA, w.AA, rule, "solution " + (i + 1));
        put(b, i, COLS.AB, w.AB, rule, "revision");
        put(b, i, COLS.U, w.U, rule, "status");
        if (w.V != null && !/(^|,)null(,|$)/.test(String(w.V))) put(b, i, COLS.V, w.V, rule, "group");
        put(b, i, COLS.W, w.W, rule, "signature position");
        put(b, i, COLS.X, w.X, rule, "person");
      }
      // Q is the day the CDF closed: written once, when it is empty (an existing date is never moved)
      if (blk.Q != null && isEmpty(cell(cdf, b.anchor, COLS.Q))) put(b, 0, COLS.Q, blk.Q, "40", "closing date");
    });
    return out;
  }

  // ---------------------------------------------------------------- 110: solution rows
  // A block with more solutions than rows gets whole rows below its last row; one with fewer loses its extra
  // continuation rows. Work goes bottom-up (110.2), so a row operation never moves a block that is still to come.
  // A continuation row shares the anchor: A:P and Y:Z merged with it, the rest as =IF($R$7="","",$R$7) (110.4).
  // MERGE_COLS / REF_COLS are what the macro built on 05/11 for A330-1-701, 881.A and 921 (and rows 7–8); a block
  // that already has continuation rows is copied from its own first continuation row instead.
  var MERGE_COLS = ["A", "B", "C", "D", "F", "G", "I", "J", "K", "L", "M", "N", "P", "Y", "Z"];
  var REF_COLS = ["E", "H", "O", "R", "S", "T", "AC", "AD", "AE", "AF", "AG", "AH"];
  var SOLUTION_COLS = ["U", "V", "W", "X", "AA", "AB"];
  function contFormula(letter, anchorRow) { var a = "$" + letter + "$" + anchorRow; return "=IF(" + a + "=\"\",\"\"," + a + ")"; }
  // =IF($R$845="","",$R$845) → 845 ; anything else → null
  function contAnchor(formula) {
    var m = /^=IF\(\$([A-Z]{1,3})\$(\d+)="",\s*"",\s*\$\1\$\2\)$/.exec(trim(formula));
    return m ? Number(m[2]) : null;
  }

  /*
   * out: the result of run(); sheetRows: the CDF values run() saw (to log removed rows)
   * → { ops: [row operations, bottom-up], changes: [every cell, in the row numbers after the operations], rowsAdded, rowsRemoved }
   *   insert: { kind, key, anchor, last, at, count, template }  — rows go in at row `at` (= last + 1), 1-based, before any operation
   *   remove: { kind, key, anchor, first, count, removed: [{ row, values: {U..AB} }] }
   */
  function layout(out, sheetRows) {
    var ops = [], byAnchor = {};
    out.blocks.forEach(function (b) { byAnchor[b.anchor + 1] = b; });
    var multi = out.blocks.filter(function (b) { return b.rows.length > 1; }).map(function (b) { return b.rows[1] + 1; });
    out.structure.forEach(function (s) {
      var b = byAnchor[s.row];
      if (!b) return;
      var last = b.rows[b.rows.length - 1] + 1;
      if (s.need > s.have) {
        // a continuation row to copy the look from: this block's own, else the nearest one above (bottom-up keeps it in place)
        var own = b.rows.length > 1 ? b.rows[1] + 1 : null, above = null;
        multi.forEach(function (r) { if (r < s.row) above = r; });
        ops.push({ kind: "insert", key: b.key, anchor: s.row, last: last, at: last + 1, count: s.need - s.have, have: s.have, need: s.need, own: own, template: own || above });
      } else if (s.need < s.have) {
        var tail = b.rows.slice(s.need - 1);
        if (tail.some(function (r, i) { return i && r !== tail[i - 1] + 1; })) {      // only whole continuation rows right under the kept ones
          out.errors.push({ stage: "110", key: b.key, message: "the block has " + s.have + " rows but " + s.need + " solution" + (s.need > 1 ? "s" : "") + ", and its extra rows are not together under the others — fix the rows by hand" });
          return;
        }
        var gone = b.rows.slice(s.need).map(function (r) {
          var v = {};
          SOLUTION_COLS.forEach(function (L) { v[L] = cell(sheetRows || [], r, col(L)); });
          return { row: r + 1, values: v };
        });
        ops.push({ kind: "remove", key: b.key, anchor: s.row, first: gone[0].row, last: gone[gone.length - 1].row, count: gone.length, have: s.have, need: s.need, removed: gone });
      }
    });
    ops.sort(function (a, b) { return b.anchor - a.anchor; });
    function shift(r) {                                   // a row number before the operations → after them
      var d = 0;
      ops.forEach(function (o) {
        if (o.kind === "insert" && r >= o.at) d += o.count;
        if (o.kind === "remove" && r > o.last) d -= o.count;
      });
      return r + d;
    }
    function removed(r) { return ops.some(function (o) { return o.kind === "remove" && r >= o.first && r <= o.last; }); }
    var changes = [];
    out.changes.forEach(function (c) {
      if (removed(c.row)) return;
      var x = {}; Object.keys(c).forEach(function (k) { x[k] = c[k]; });
      x.sourceRow = c.row; x.row = shift(c.row);
      changes.push(x);
    });
    ops.forEach(function (o) {
      o.finalAnchor = shift(o.anchor);
      if (o.kind === "remove") { o.finalFirst = o.finalAnchor + o.need; return; }
      o.finalFirst = shift(o.last) + 1;
      var b = byAnchor[o.anchor];
      for (var i = o.have; i < o.need; i++) {
        var w = b.want[i], row = o.finalFirst + (i - o.have);
        SOLUTION_COLS.forEach(function (L) {
          var v = w[L];
          if (v == null || /(^|,)null(,|$)/.test(String(v)) || v === "") return;   // V without a people list stays empty
          changes.push({ sheet: "CDF", row: row, col: col(L), colLetter: L, key: o.key, oldValue: "", newValue: v, rule: w.match || "110.2", why: "solution " + (i + 1) + " (new row)", newRow: true });
        });
      }
    });
    // Q, the closing date of a block that also grows, is on the anchor: already in out.changes
    changes.sort(function (a, b) { return a.row - b.row || a.col - b.col; });
    return {
      ops: ops, changes: changes, shift: shift, isRemoved: removed,
      rowsAdded: ops.filter(function (o) { return o.kind === "insert"; }).reduce(function (s, o) { return s + o.count; }, 0),
      rowsRemoved: ops.filter(function (o) { return o.kind === "remove"; }).reduce(function (s, o) { return s + o.count; }, 0)
    };
  }

  // stage 1 + stage 2 as one update: every CDF cell in the row numbers after the row operations. A cell both
  // stages write keeps stage 1's "before" (the value in the sheet now) and stage 2's "after".
  function combine(stage1, plan) {
    var at = {}, out = [];
    (stage1 || []).forEach(function (c) {
      if (plan.isRemoved(c.row)) return;
      var x = {}; Object.keys(c).forEach(function (k) { x[k] = c[k]; });
      x.sourceRow = c.row; x.row = plan.shift(c.row);
      at[x.row + ":" + x.col] = x; out.push(x);
    });
    plan.changes.forEach(function (c) {
      var prev = at[c.row + ":" + c.col];
      if (!prev) { out.push(c); return; }
      prev.newValue = c.newValue; prev.rule = c.rule; prev.why = c.why;
    });
    return out.filter(function (c) { return !(trim(c.oldValue) === trim(c.newValue) && typeof c.oldValue === typeof c.newValue); })
      .sort(function (a, b) { return a.row - b.row || a.col - b.col; });
  }
  function plural(n, w) { return n + " " + w + (n === 1 ? "" : "s"); }
  // the log lines of the row operations; "+3 rows · anchor 907" is what undo reads back
  function rowOpLogRows(ops, stamp, name, show) {
    var rows = [];
    (ops || []).forEach(function (o) {
      if (o.kind === "insert") {
        rows.push([stamp, name, "CDF", String(o.finalFirst), o.key, "ROWS", "", "+" + plural(o.count, "row") + " · anchor " + o.finalAnchor,
          "Rule 110.2 — solution " + (o.count > 1 ? "rows " + (o.have + 1) + "–" + o.need : "row " + o.need) + " added under the block"]);
        return;
      }
      rows.push([stamp, name, "CDF", String(o.finalFirst), o.key, "ROWS", "", "-" + plural(o.count, "row") + " · anchor " + o.finalAnchor,
        "Rule 110.2 — " + plural(o.count, "extra solution row") + " removed (" + o.have + " rows, " + o.need + " solution" + (o.need === 1 ? "" : "s") + ")"]);
      o.removed.forEach(function (rr, j) {
        SOLUTION_COLS.forEach(function (L) {
          var v = rr.values[L];
          if (trim(v) === "") return;
          rows.push([stamp, name, "CDF", String(o.finalFirst + j), o.key, L, show ? show(v) : str(v), "", "Removed row (rule 110.2)"]);
        });
      });
    });
    return rows;
  }
  var ROWS_RE = /^([+-])(\d+) rows? · anchor (\d+)$/;
  function parseRowsNote(v) { var m = ROWS_RE.exec(trim(v)); return m ? { kind: m[1] === "+" ? "insert" : "remove", count: Number(m[2]), anchor: Number(m[3]) } : null; }

  // the CDF sheet after the whole update, in memory: row operations (bottom-up, as the add-in does them), then every
  // change in the row numbers after them. New continuation rows show their anchor in REF_COLS, as the formulas do.
  function finalValues(values, ops, changes) {
    var out = (values || []).map(function (r) { return (r || []).slice(); });
    (ops || []).forEach(function (o) {
      if (o.kind === "remove") { out.splice(o.first - 1, o.count); return; }
      var add = [];
      for (var i = 0; i < o.count; i++) {
        var nr = [], a = out[o.anchor - 1] || [];
        REF_COLS.forEach(function (L) { var c = col(L); while (nr.length < c) nr.push(""); nr[c] = a[c] == null ? "" : a[c]; });
        add.push(nr);
      }
      while (out.length < o.at - 1) out.push([]);
      Array.prototype.splice.apply(out, [o.at - 1, 0].concat(add));
    });
    return applyInMemory(out, changes);
  }

  // stage 1's CDF changes applied to a copy of the sheet values, so stage 2 reads the sheet as stage 1 leaves it
  function applyInMemory(values, changes) {
    var copy = (values || []).map(function (r) { return r.slice(); });
    (changes || []).forEach(function (c) {
      var r = c.row - 1;
      while (copy.length <= r) copy.push([]);
      while (copy[r].length <= c.col) copy[r].push("");
      copy[r][c.col] = c.newValue;
    });
    return copy;
  }

  return {
    MERGE_COLS: MERGE_COLS, REF_COLS: REF_COLS, SOLUTION_COLS: SOLUTION_COLS, contFormula: contFormula, contAnchor: contAnchor,
    layout: layout, applyInMemory: applyInMemory, finalValues: finalValues, combine: combine, rowOpLogRows: rowOpLogRows, parseRowsNote: parseRowsNote,
    normId: normId, revKey: revKey, revShow: revShow, toSerial: toSerial, dayOf: dayOf,
    readMantra: readMantra, readMantraRev: readMantraRev, decodeAnswer: decodeAnswer, parts: parts, oneEdit: oneEdit,
    resolveIdentity: resolveIdentity, selectRecord: selectRecord, uOf: uOf, uvwx: uvwx, makePeople: makePeople, groupOf: groupOf, personKey: personKey,
    EXPLICIT: EXPLICIT, baseKey: baseKey, col: col, cell: cell, trim: trim, upper: upper, solveDoc: solveDoc, run: run
  };
});
