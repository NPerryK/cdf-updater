// Run: node tests/link.test.js   (synthetic data only)
// The link with My Notes: cdf-link.js (shared with the Outlook add-in) and the Excel-side builders.
// At least 5 cases for every function.
const assert = require("assert");
const L = require("../cdf-link.js");
const F = require("../features.js");
const A = require("../cdf-algo.js");

let passed = 0, failed = 0;
const perFn = {};
function test(fn, name, body) {
  perFn[fn] = (perFn[fn] || 0) + 1;
  try { body(); passed++; console.log("  ✓ " + fn + ": " + name); }
  catch (e) { failed++; console.error("  ✗ " + fn + ": " + name + "\n    " + e.stack); process.exitCode = 1; }
}
const labels = (t) => L.refs(t).map((r) => r.label);

// ---------------- refs ----------------
test("refs", "full numbers with a revision", () => assert.deepStrictEqual(labels("A330-1-961.- and A330-1-740.A"), ["A330-1-961.-", "A330-1-740.A"]));
test("refs", "CDF 740.A, CDF#52, CDF-052 (one CDF)", () => assert.deepStrictEqual(labels("Re: CDF 740.A / cdf#52 / CDF-052"), ["A330-1-740.A", "A330-1-052"]));
test("refs", "file names: the ATT file keeps its revision", () => assert.deepStrictEqual(labels("A330-1-961.-ATT.pdf"), ["A330-1-961.-"]));
test("refs", "not dates, exports, plurals or the aircraft type", () =>
  assert.deepStrictEqual(labels("CDF 2026-11-05 · CDF_List_Export (19).xlsx · CDFs for 05/11 · A330-300 P2F · CDF 05/11"), []));
test("refs", "a number without a revision gives way to the exact one", () => assert.deepStrictEqual(labels("CDF 961 … A330-1-961.-"), ["A330-1-961.-"]));
test("refs", "file extensions are not revisions", () => assert.deepStrictEqual(labels("CDF 961.pdf, CDF 962.xlsx"), ["A330-1-961", "A330-1-962"]));
test("refs", "empty and odd input", () => { assert.deepStrictEqual(L.refs(""), []); assert.deepStrictEqual(L.refs(null), []); assert.deepStrictEqual(L.refs(12), []); });
test("refs", "numbers above 1999 are ignored", () => assert.deepStrictEqual(labels("CDF 2500 and CDF 1250"), ["A330-1-1250"]));

test("refs", "a list after CDFs: keeps going", () => assert.deepStrictEqual(labels("Waiting for ANN/SIE — 23 open CDFs: 961, 740.A, 279 and 052"), ["A330-1-961", "A330-1-740.A", "A330-1-279", "A330-1-052"]));
test("refs", "a list stops at small numbers and words", () => { assert.deepStrictEqual(labels("CDF 961, 2 drawings"), ["A330-1-961"]); assert.deepStrictEqual(labels("CDF 740.A / 741 ו-962"), ["A330-1-740.A", "A330-1-741", "A330-1-962"]); });
// ---------------- label / keyParts / matches ----------------
test("label", "pads to three digits", () => assert.strictEqual(L.label(52, ""), "A330-1-052"));
test("label", "keeps the revision", () => assert.strictEqual(L.label(740, "A"), "A330-1-740.A"));
test("label", "original revision", () => assert.strictEqual(L.label(961, "-"), "A330-1-961.-"));
test("label", "four digits", () => assert.strictEqual(L.label(1250, ""), "A330-1-1250"));
test("label", "number given as text", () => assert.strictEqual(L.label("7", "B"), "A330-1-007.B"));
test("keyParts", "sheet key", () => assert.deepStrictEqual(L.keyParts("A330-1-961.-"), { n: 961, rev: "-", extra: "" }));
test("keyParts", "the duplicate's suffix", () => assert.deepStrictEqual(L.keyParts("A330-1-052 (another)"), { n: 52, rev: "", extra: "(another)" }));
test("keyParts", "lower case revision", () => assert.strictEqual(L.keyParts("a330-1-740.a").rev, "A"));
test("keyParts", "not a CDF key", () => { assert.strictEqual(L.keyParts("TOTAL"), null); assert.strictEqual(L.keyParts(""), null); });
test("keyParts", "spaces around", () => assert.strictEqual(L.keyParts("  A330-1-100.B ").n, 100));
test("matches", "no revision matches every revision", () => { const r = L.refs("CDF 740")[0]; assert.ok(L.matches(r, "A330-1-740.-") && L.matches(r, "A330-1-740.A")); });
test("matches", "a revision matches only itself", () => { const r = L.refs("CDF 740.A")[0]; assert.ok(L.matches(r, "A330-1-740.A") && !L.matches(r, "A330-1-740.-")); });
test("matches", "different number", () => assert.ok(!L.matches(L.refs("CDF 740")[0], "A330-1-741.-")));
test("matches", "parts object as key", () => assert.ok(L.matches({ n: 9, rev: "" }, { n: 9, rev: "C" })));
test("matches", "bad input", () => { assert.ok(!L.matches(null, "A330-1-1.-")); assert.ok(!L.matches({ n: 1, rev: "" }, "TOTAL")); });

// ---------------- encode / decode ----------------
const ST = { at: "2026-11-06T06:10:00.000Z", from: "CDF Updater 1.5.0", count: 1, rows: [["A330-1-961.-", "Open", "3 - Conversion Engineer", "", "ANN/SIE", "D-1", "INSTL", "26/10/2026", [], ""]] };
test("encode", "kind and version first", () => { const o = JSON.parse(L.encode(L.KIND_STATUS, ST)); assert.strictEqual(o.kind, "cdf-updater/status"); assert.strictEqual(o.v, 1); });
test("encode", "keeps the payload", () => assert.strictEqual(JSON.parse(L.encode(L.KIND_STATUS, ST)).rows[0][0], "A330-1-961.-"));
test("encode", "one line (survives chat and email)", () => assert.ok(!/\n/.test(L.encode(L.KIND_NOTES, { notes: [{ text: "a\nb" }] }))));
test("encode", "Hebrew text round trip", () => assert.strictEqual(JSON.parse(L.encode(L.KIND_NOTES, { notes: [{ text: "לבדוק מול מור" }] })).notes[0].text, "לבדוק מול מור"));
test("encode", "empty payload", () => assert.deepStrictEqual(JSON.parse(L.encode(L.KIND_NOTES)), { kind: "my-notes/cdf-notes", v: 1 }));
test("decode", "reads what encode wrote", () => { const r = L.decode(L.encode(L.KIND_STATUS, ST), L.KIND_STATUS); assert.ok(r.ok); assert.strictEqual(r.data.count, 1); });
test("decode", "text around it is fine", () => assert.ok(L.decode("Here it is:\n" + L.encode(L.KIND_STATUS, ST) + "\nthanks", L.KIND_STATUS).ok));
test("decode", "pasted into the wrong add-in", () => assert.match(L.decode(L.encode(L.KIND_STATUS, ST), L.KIND_NOTES).error, /This is the CDF status from CDF Updater \(Excel\) — paste it into My Notes in Outlook/));
test("decode", "cut off", () => {
  assert.match(L.decode(L.encode(L.KIND_STATUS, ST).slice(0, 60) + "}", L.KIND_STATUS).error, /cut off or changed/);
  assert.match(L.decode(L.encode(L.KIND_STATUS, ST).slice(0, 60), L.KIND_STATUS).error, /^The text was cut off\./);
});
test("decode", "nothing / plain text", () => { assert.match(L.decode("  ", L.KIND_NOTES).error, /Nothing was pasted/); assert.match(L.decode("hello", L.KIND_NOTES).error, /isn't notes from My Notes/); });
test("decode", "newer version", () => assert.match(L.decode(JSON.stringify({ kind: L.KIND_NOTES, v: 2, notes: [] }), L.KIND_NOTES).error, /newer version/));
test("decode", "missing rows / notes", () => { assert.ok(!L.decode(JSON.stringify({ kind: L.KIND_STATUS, v: 1 }), L.KIND_STATUS).ok); assert.ok(!L.decode(JSON.stringify({ kind: L.KIND_NOTES, v: 1 }), L.KIND_NOTES).ok); });
test("decode", "unknown kind", () => assert.match(L.decode('{"kind":"x"}', L.KIND_NOTES).error, /^This isn't notes from My Notes/));

test("kindOf", "status", () => assert.strictEqual(L.kindOf(L.encode(L.KIND_STATUS, ST)), L.KIND_STATUS));
test("kindOf", "notes, with spaces after the colon", () => assert.strictEqual(L.kindOf('{"kind": "my-notes/cdf-notes", "v": 1}'), L.KIND_NOTES));
test("kindOf", "plain text", () => assert.strictEqual(L.kindOf("A330-1-961.-"), ""));
test("kindOf", "another kind", () => assert.strictEqual(L.kindOf('{"kind":"other"}'), ""));
test("kindOf", "empty", () => { assert.strictEqual(L.kindOf(""), ""); assert.strictEqual(L.kindOf(null), ""); });
// ---------------- newer / shared storage ----------------
const a = { at: "2026-11-05T06:00:00Z" }, b = { at: "2026-11-06T06:00:00Z" };
test("newer", "later wins", () => assert.strictEqual(L.newer(a, b), b));
test("newer", "order doesn't matter", () => assert.strictEqual(L.newer(b, a), b));
test("newer", "one missing", () => { assert.strictEqual(L.newer(null, a), a); assert.strictEqual(L.newer(a, null), a); });
test("newer", "both missing", () => assert.strictEqual(L.newer(null, undefined), null));
test("newer", "same time keeps the first", () => assert.strictEqual(L.newer(a, { at: a.at }), a));
function mem() { const m = {}; return { getItem: (k) => (k in m ? m[k] : null), setItem: (k, v) => { m[k] = String(v); }, m }; }
test("writeShared", "writes", () => { const s = mem(); assert.ok(L.writeShared(s, L.LS_STATUS, "x")); assert.strictEqual(s.m[L.LS_STATUS], "x"); });
test("writeShared", "storage that throws", () => assert.strictEqual(L.writeShared({ setItem() { throw new Error("quota"); } }, "k", "v"), false));
test("writeShared", "no storage", () => assert.strictEqual(L.writeShared(null, "k", "v"), false));
test("writeShared", "overwrites", () => { const s = mem(); L.writeShared(s, "k", "1"); L.writeShared(s, "k", "2"); assert.strictEqual(s.m.k, "2"); });
test("writeShared", "keys are separate", () => { const s = mem(); L.writeShared(s, L.LS_STATUS, "a"); L.writeShared(s, L.LS_NOTES, "b"); assert.deepStrictEqual(Object.keys(s.m).sort(), [L.LS_NOTES, L.LS_STATUS].sort()); });
test("readShared", "reads a package", () => { const s = mem(); L.writeShared(s, L.LS_STATUS, L.encode(L.KIND_STATUS, ST)); assert.strictEqual(L.readShared(s, L.LS_STATUS, L.KIND_STATUS).count, 1); });
test("readShared", "nothing there", () => assert.strictEqual(L.readShared(mem(), L.LS_STATUS, L.KIND_STATUS), null));
test("readShared", "wrong kind", () => { const s = mem(); L.writeShared(s, "k", L.encode(L.KIND_NOTES, { notes: [] })); assert.strictEqual(L.readShared(s, "k", L.KIND_STATUS), null); });
test("readShared", "broken text", () => { const s = mem(); L.writeShared(s, "k", "{oops"); assert.strictEqual(L.readShared(s, "k", L.KIND_STATUS), null); });
test("readShared", "storage that throws", () => assert.strictEqual(L.readShared({ getItem() { throw new Error("denied"); } }, "k", L.KIND_STATUS), null));

// ---------------- statusFor / ageDays ----------------
const ROWS = { rows: [["A330-1-740.-", "Released"], ["A330-1-740.A", "Open", "3 - Conversion Engineer"], ["A330-1-961.-", "Open"], ["A330-1-052", "Open"], ["A330-1-052 (another)", "Cancelled"]] };
test("statusFor", "a number gives all its revisions", () => assert.deepStrictEqual(L.statusFor(L.refs("CDF 740"), ROWS).map((x) => x.key), ["A330-1-740.-", "A330-1-740.A"]));
test("statusFor", "fields by name", () => { const r = L.statusFor(L.refs("CDF 740.A"), ROWS)[0]; assert.strictEqual(r.status, "Open"); assert.strictEqual(r.station, "3 - Conversion Engineer"); });
test("statusFor", "not in the status → missing", () => assert.deepStrictEqual(L.statusFor(L.refs("CDF 999"), ROWS), [{ key: "A330-1-999", missing: true }]));
test("statusFor", "no double rows", () => assert.strictEqual(L.statusFor(L.refs("CDF 961 and A330-1-961.- and CDF#961"), ROWS).length, 1));
test("statusFor", "the duplicate 052", () => assert.deepStrictEqual(L.statusFor(L.refs("CDF 052"), ROWS).map((x) => x.status), ["Open", "Cancelled"]));
test("statusFor", "no status pasted", () => assert.deepStrictEqual(L.statusFor(L.refs("CDF 1"), null), [{ key: "A330-1-001", missing: true }]));
const NOW = new Date(2026, 10, 6, 15, 0);
test("ageDays", "same day", () => assert.strictEqual(L.ageDays(new Date(2026, 10, 6, 8, 10).toISOString(), NOW), 0));
test("ageDays", "yesterday", () => assert.strictEqual(L.ageDays(new Date(2026, 10, 5, 23, 0).toISOString(), NOW), 1));
test("ageDays", "a week", () => assert.strictEqual(L.ageDays(new Date(2026, 9, 30, 9, 0).toISOString(), NOW), 7));
test("ageDays", "missing", () => assert.strictEqual(L.ageDays(null, NOW), null));
test("ageDays", "not a date", () => assert.strictEqual(L.ageDays("soon", NOW), null));

// ---------------- notesFor / dueNotes ----------------
const NOTES = [
  { refs: ["A330-1-740"], text: "any revision", state: "Open", updated: 5 },
  { refs: ["A330-1-740.A"], text: "rev A", state: "Overdue", due: "2026-11-04", updated: 9 },
  { refs: ["A330-1-740.-"], text: "original", state: "Done", updated: 20 },
  { refs: ["A330-1-961.-", "A330-1-740.A"], text: "two CDFs", state: "Due today", due: "2026-11-06", updated: 7 },
  { refs: ["A330-1-500"], text: "info", state: "Info", due: "2026-11-01", updated: 1 },
];
test("notesFor", "rev A: its own notes and the any-revision note", () => assert.deepStrictEqual(L.notesFor("A330-1-740.A", NOTES).map((n) => n.text), ["rev A", "two CDFs", "any revision"]));
test("notesFor", "done notes last", () => assert.deepStrictEqual(L.notesFor("A330-1-740.-", NOTES).map((n) => n.text), ["any revision", "original"]));
test("notesFor", "a note on two CDFs shows on both", () => assert.deepStrictEqual(L.notesFor("A330-1-961.-", NOTES).map((n) => n.text), ["two CDFs"]));
test("notesFor", "no notes", () => assert.deepStrictEqual(L.notesFor("A330-1-111.-", NOTES), []));
test("notesFor", "bad key or no list", () => { assert.deepStrictEqual(L.notesFor("TOTAL", NOTES), []); assert.deepStrictEqual(L.notesFor("A330-1-740.A", null), []); });
test("dueNotes", "overdue and today, oldest first", () => assert.deepStrictEqual(L.dueNotes(NOTES, NOW).map((n) => n.text), ["rev A", "two CDFs"]));
test("dueNotes", "info and done are not follow-ups", () => assert.ok(!L.dueNotes(NOTES, NOW).some((n) => n.state === "Info" || n.state === "Done")));
test("dueNotes", "future due dates wait", () => assert.deepStrictEqual(L.dueNotes([{ state: "Open", due: "2026-11-09" }], NOW), []));
test("dueNotes", "no due date", () => assert.deepStrictEqual(L.dueNotes([{ state: "Open" }], NOW), []));
test("dueNotes", "nothing pasted", () => assert.deepStrictEqual(L.dueNotes(null, NOW), []));

// ---------------- bridgeStatus / notesInfo (Excel side) ----------------
const C = A.col;
function row(obj, w) { const r = new Array(w || 30).fill(""); for (const [k, v] of Object.entries(obj)) r[C(k)] = v; return r; }
const CDF = [
  row({ C: "CDF NUM" }),
  row({ C: "A330-1-961.-", F: 46321, R: "Open", S: "3 - Conversion Engineer", T: "-", P: "ANN/SIE", J: "D-961", K: "INSTL MID", U: "NO ENG. DOC.", AA: "-" }),
  row({ R: "Open", U: "4 - RELEASED", AA: "330-92-80-R01", AB: "A" }),
  row({ C: "A330-1-740.-", R: "Released", S: "Released", P: "IAI", U: "4 - RELEASED", AA: "DOC-740" }),
  row({ A: "link", B: "link" }),
  row({ C: "A330-1-279.-", R: "Open", S: "9 - Final Approval", T: "Dana Ben-Ari", U: "4 - RELEASED", AA: "DOC-279" }),
];
const SUM = { pendingPeople: [{ table: "DANA PENDING TABLE", person: "Dana Ben-Ari" }], tables: [{ name: "DANA PENDING TABLE", rows: [["A330-1-279.-", "DOC-279"], ["TOTAL", 1]] }] };
const BS = F.bridgeStatus(CDF, SUM, new Date(Date.UTC(2026, 10, 6, 6, 10)), "CDF Updater 1.5.0");
test("bridgeStatus", "one row per CDF, solution rows folded in", () => assert.deepStrictEqual(BS.rows.map((r) => r[0]), ["A330-1-961.-", "A330-1-740.-", "A330-1-279.-"]));
test("bridgeStatus", "fields in the shared column order", () => {
  const o = L.rowObj(BS.rows[0]);
  assert.deepStrictEqual([o.status, o.station, o.signature, o.supplier, o.dwg, o.dwgName, o.received], ["Open", "3 - Conversion Engineer", "", "ANN/SIE", "D-961", "INSTL MID", "26/10/2026"]);
});
test("bridgeStatus", "documents with their Mantra status (open CDFs only)", () => {
  assert.deepStrictEqual(BS.rows[0][8], [["", "", "NO ENG. DOC."], ["330-92-80-R01", "A", "4 - RELEASED"]]);
  assert.deepStrictEqual(BS.rows[1][8], []);
});
test("bridgeStatus", "ready to sign from the pending tables", () => { assert.strictEqual(BS.rows[2][9], "Dana Ben-Ari"); assert.strictEqual(BS.rows[0][9], ""); });
test("bridgeStatus", "time, source and count", () => { assert.strictEqual(BS.at, "2026-11-06T06:10:00.000Z"); assert.strictEqual(BS.from, "CDF Updater 1.5.0"); assert.strictEqual(BS.count, 3); });
test("bridgeStatus", "no summary, no sheet", () => { assert.strictEqual(F.bridgeStatus(CDF, null).rows[2][9], ""); assert.deepStrictEqual(F.bridgeStatus(null).rows, []); });
test("bridgeStatus", "round trip through the clipboard text", () => {
  const r = L.decode(L.encode(L.KIND_STATUS, BS), L.KIND_STATUS);
  assert.ok(r.ok); assert.strictEqual(L.statusFor(L.refs("Re: CDF 279"), r.data)[0].ready, "Dana Ben-Ari");
});
test("notesInfo", "counts notes, CDFs and open ones", () => assert.deepStrictEqual(F.notesInfo({ notes: NOTES }), { notes: 5, cdfs: 5, open: 3 }));
test("notesInfo", "nothing pasted", () => assert.deepStrictEqual(F.notesInfo(null), { notes: 0, cdfs: 0, open: 0 }));
test("notesInfo", "empty list", () => assert.deepStrictEqual(F.notesInfo({ notes: [] }), { notes: 0, cdfs: 0, open: 0 }));
test("notesInfo", "a CDF in two notes counts once", () => assert.strictEqual(F.notesInfo({ notes: [{ refs: ["A"] }, { refs: ["A"] }] }).cdfs, 1));
test("notesInfo", "notes without refs", () => assert.deepStrictEqual(F.notesInfo({ notes: [{ state: "Done" }] }), { notes: 1, cdfs: 0, open: 0 }));

// ---------------- phaseOf / stationDays / parseDay / ymd ----------------
test("phaseOf", "initiator and POC", () => { assert.strictEqual(L.phaseOf("Open", "1 - Initiator").index, 0); assert.strictEqual(L.phaseOf("Open", "2 - POC").index, 1); });
test("phaseOf", "engineering stations 3–8", () => { assert.strictEqual(L.phaseOf("Open", "3 - Conversion Engineer").index, 2); assert.strictEqual(L.phaseOf("Open", "4 - AVIONICS").index, 2); });
test("phaseOf", "final approval", () => assert.strictEqual(L.phaseOf("Open", "9 - Final Approval").index, 3));
test("phaseOf", "released, even when production steps follow", () => { assert.strictEqual(L.phaseOf("Released", "10 - Production Engineering").index, 4); assert.strictEqual(L.phaseOf("", "Released").index, 4); });
test("phaseOf", "cancelled and unknown", () => { assert.deepStrictEqual(L.phaseOf("Cancelled", ""), { index: -1, cancelled: true }); assert.strictEqual(L.phaseOf("Open", "").index, -1); });
const D6 = new Date(2026, 10, 6, 15, 0);
test("stationDays", "exact date", () => assert.deepStrictEqual(L.stationDays({ status: "Open", since: "2026-10-20", exact: 1 }, D6), { days: 17, min: 17, exact: true, stuck: true }));
test("stationDays", "not exact: most and least", () => assert.deepStrictEqual(L.stationDays({ status: "Open", since: "2026-10-26", exact: 0, seen: "2026-11-06" }, D6), { days: 11, min: 0, exact: false, stuck: false }));
test("stationDays", "not exact but seen there 14+ days → stuck", () => assert.strictEqual(L.stationDays({ status: "Open", since: "2026-08-01", exact: 0, seen: "2026-10-20" }, D6).stuck, true));
test("stationDays", "falls back to the received date", () => assert.strictEqual(L.stationDays({ status: "Open", received: "26/10/2026" }, D6).days, 11));
test("stationDays", "only for open CDFs", () => { assert.strictEqual(L.stationDays({ status: "Released", since: "2026-10-01", exact: 1 }, D6), null); assert.strictEqual(L.stationDays(null, D6), null); });
test("parseDay", "yyyy-mm-dd", () => assert.strictEqual(L.parseDay("2026-11-06").getDate(), 6));
test("parseDay", "dd/mm/yyyy", () => assert.strictEqual(L.parseDay("26/10/2026").getMonth(), 9));
test("parseDay", "ISO with time", () => assert.strictEqual(L.parseDay("2026-11-06T06:10:00Z").getFullYear(), 2026));
test("parseDay", "empty", () => assert.strictEqual(L.parseDay(""), null));
test("parseDay", "not a date", () => assert.strictEqual(L.parseDay("soon"), null));
test("ymd", "pads", () => assert.strictEqual(L.ymd(new Date(2026, 0, 5)), "2026-01-05"));
test("ymd", "end of year", () => assert.strictEqual(L.ymd(new Date(2026, 11, 31)), "2026-12-31"));
test("ymd", "round trip with parseDay", () => assert.strictEqual(L.ymd(L.parseDay("2026-11-06")), "2026-11-06"));
test("ymd", "local time, not UTC", () => assert.strictEqual(L.ymd(new Date(2026, 10, 6, 23, 59)), "2026-11-06"));
test("ymd", "leap day", () => assert.strictEqual(L.ymd(new Date(2028, 1, 29)), "2028-02-29"));

// ---------------- sentKind / addWorkdays ----------------
test("sentKind", "supplier email", () => assert.deepStrictEqual(L.sentKind("A330-300 P2F — Open CDFs at Conversion Engineer — ANN/SIE — 06/11/2026"), { kind: "supplier", who: "ANN/SIE" }));
test("sentKind", "signature email", () => assert.deepStrictEqual(L.sentKind("A330-300 P2F — CDFs ready for your signature — Dana Ben-Ari — 06/11/2026"), { kind: "sign", who: "Dana Ben-Ari" }));
test("sentKind", "a reply keeps it (RE:, השב:)", () => { assert.strictEqual(L.sentKind("RE: A330-300 P2F — Open CDFs at Conversion Engineer — IAI — 1/1/2027").who, "IAI"); assert.strictEqual(L.sentKind("השב: A330-300 P2F — CDFs ready for your signature — Tom Bar — 06/11/2026").kind, "sign"); });
test("sentKind", "the daily email", () => assert.strictEqual(L.sentKind("A330-300 P2F — Daily CDF Status — 06/11/2026").kind, "daily"));
test("sentKind", "anything else", () => { assert.strictEqual(L.sentKind("RE: CDF 961"), null); assert.strictEqual(L.sentKind(""), null); });
test("addWorkdays", "Thursday + 3 → Tuesday (Friday and Saturday skipped)", () => assert.strictEqual(L.addWorkdays(new Date(2026, 10, 5), 3), "2026-11-10"));
test("addWorkdays", "Sunday + 1 → Monday", () => assert.strictEqual(L.addWorkdays(new Date(2026, 10, 8), 1), "2026-11-09"));
test("addWorkdays", "Friday + 1 → Sunday", () => assert.strictEqual(L.addWorkdays(new Date(2026, 10, 6), 1), "2026-11-08"));
test("addWorkdays", "zero days", () => assert.strictEqual(L.addWorkdays(new Date(2026, 10, 6), 0), "2026-11-06"));
test("addWorkdays", "a full week", () => assert.strictEqual(L.addWorkdays(new Date(2026, 10, 8), 5), "2026-11-15"));

// ---------------- closable / unknownRefs ----------------
const ST2 = { rows: [["A330-1-740.-", "Released"], ["A330-1-740.A", "Released"], ["A330-1-961.-", "Open"], ["A330-1-100.-", "Cancelled"]] };
test("closable", "every CDF released or cancelled", () => assert.deepStrictEqual(L.closable([{ refs: ["A330-1-740.A", "A330-1-100"], state: "Open", text: "x" }], ST2).length, 1));
test("closable", "one still open → keep", () => assert.deepStrictEqual(L.closable([{ refs: ["A330-1-740.A", "A330-1-961.-"], state: "Open" }], ST2), []));
test("closable", "a number without revision: all its revisions released", () => assert.strictEqual(L.closable([{ refs: ["A330-1-740"], state: "Overdue" }], ST2).length, 1));
test("closable", "done, info, unknown CDF, no refs", () => assert.deepStrictEqual(L.closable([{ refs: ["A330-1-740.A"], state: "Done" }, { refs: ["A330-1-740.A"], state: "Info" }, { refs: ["A330-1-999"], state: "Open" }, { refs: [], state: "Open" }], ST2), []));
test("closable", "no status", () => assert.deepStrictEqual(L.closable([{ refs: ["A330-1-740.A"], state: "Open" }], null), []));
test("unknownRefs", "a typo", () => assert.deepStrictEqual(L.unknownRefs([{ refs: ["A330-1-9610"] }], ST2), [{ ref: "A330-1-9610", notes: 1 }]));
test("unknownRefs", "counts notes per number", () => assert.deepStrictEqual(L.unknownRefs([{ refs: ["A330-1-555"] }, { refs: ["A330-1-555", "A330-1-961.-"] }], ST2), [{ ref: "A330-1-555", notes: 2 }]));
test("unknownRefs", "known numbers are fine", () => assert.deepStrictEqual(L.unknownRefs([{ refs: ["A330-1-740", "A330-1-961.-"] }], ST2), []));
test("unknownRefs", "a revision that isn't there", () => assert.deepStrictEqual(L.unknownRefs([{ refs: ["A330-1-961.B"] }], ST2), [{ ref: "A330-1-961.B", notes: 1 }]));
test("unknownRefs", "no status or notes", () => { assert.deepStrictEqual(L.unknownRefs([{ refs: ["A330-1-1"] }], null), []); assert.deepStrictEqual(L.unknownRefs(null, ST2), []); });

// ---------------- trackStations / daysSince / daysAt / glanceList / stuckCount (Excel side) ----------------
const TR_CDF = [
  row({ C: "CDF NUM" }),
  row({ C: "A330-1-001.-", R: "Open", S: "3 - Conversion Engineer", F: 46300 }),
  row({ C: "A330-1-002.-", R: "Open", S: "9 - Final Approval", F: 46310 }),
  row({ C: "A330-1-003.-", R: "Released", S: "Released" }),
];
const LOG = [["When", "By", "Sheet", "Row", "CDF", "Col", "Before", "After", "Note"], ["01/11/2026 08:10", "x", "CDF", 3, "A330-1-002.-", "S", "3 - Conversion Engineer", "9 - Final Approval", ""]];
const NOV6 = new Date(2026, 10, 6);
test("trackStations", "first time: log date (exact) or received date (not exact)", () => {
  const t = F.trackStations(null, TR_CDF, LOG, NOV6);
  assert.deepStrictEqual(t.map["A330-1-002.-"], ["9 - Final Approval", "2026-11-01", 1, "2026-11-06"]);
  assert.deepStrictEqual(t.map["A330-1-001.-"], ["3 - Conversion Engineer", "2026-10-05", 0, "2026-11-06"]);
});
test("trackStations", "only open CDFs", () => assert.ok(!("A330-1-003.-" in F.trackStations(null, TR_CDF, [], NOV6).map)));
test("trackStations", "same station → kept, nothing changed", () => { const a = F.trackStations(null, TR_CDF, [], NOV6); const b = F.trackStations(a.map, TR_CDF, [], new Date(2026, 10, 9)); assert.strictEqual(b.changed, false); assert.deepStrictEqual(b.map, a.map); });
test("trackStations", "a new station → today, exact", () => {
  const a = F.trackStations(null, TR_CDF, [], NOV6).map, moved = TR_CDF.map((r) => r.slice()); moved[1][C("S")] = "9 - Final Approval";
  assert.deepStrictEqual(F.trackStations(a, moved, [], new Date(2026, 10, 9)).map["A330-1-001.-"], ["9 - Final Approval", "2026-11-09", 1, "2026-11-09"]);
});
test("trackStations", "released since → dropped, and that is a change", () => {
  const a = F.trackStations(null, TR_CDF, [], NOV6).map, rel = TR_CDF.map((r) => r.slice()); rel[2][C("R")] = "Released";
  const b = F.trackStations(a, rel, [], NOV6); assert.strictEqual(b.changed, true); assert.ok(!("A330-1-002.-" in b.map));
});
test("trackStations", "a log change to another station isn't used", () => assert.strictEqual(F.trackStations(null, TR_CDF, [LOG[0], ["01/11/2026 08:10", "x", "CDF", 2, "A330-1-001.-", "S", "", "2 - POC", ""]], NOV6).map["A330-1-001.-"][2], 0));
test("daysSince", "days", () => assert.strictEqual(F.daysSince("2026-11-01", NOV6), 5));
test("daysSince", "today", () => assert.strictEqual(F.daysSince("2026-11-06", new Date(2026, 10, 6, 23)), 0));
test("daysSince", "future → 0", () => assert.strictEqual(F.daysSince("2026-12-01", NOV6), 0));
test("daysSince", "bad", () => { assert.strictEqual(F.daysSince("", NOV6), null); assert.strictEqual(F.daysSince("01/11/2026", NOV6), null); });
test("daysSince", "across months", () => assert.strictEqual(F.daysSince("2026-10-06", NOV6), 31));
test("daysAt", "exact", () => assert.deepStrictEqual(F.daysAt(["x", "2026-10-20", 1, "2026-11-06"], NOV6), { days: 17, min: 17, exact: true, stuck: true }));
test("daysAt", "not exact, just seen", () => assert.deepStrictEqual(F.daysAt(["x", "2026-10-05", 0, "2026-11-06"], NOV6), { days: 32, min: 0, exact: false, stuck: false }));
test("daysAt", "not exact, seen there 14 days ago", () => assert.strictEqual(F.daysAt(["x", "2026-10-05", 0, "2026-10-23"], NOV6).stuck, true));
test("daysAt", "missing", () => assert.strictEqual(F.daysAt(null, NOV6), null));
test("daysAt", "first seen missing (older data)", () => assert.deepStrictEqual(F.daysAt(["x", "2026-10-05", 0], NOV6), { days: 32, min: 0, exact: false, stuck: false }));
const GL = { cdf: TR_CDF, summary: SUM, stations: F.trackStations(null, TR_CDF, LOG, NOV6).map };
test("glanceList", "open, longest first", () => { const l = F.glanceList("open", GL, NOV6); assert.strictEqual(l.title, "Open CDFs"); assert.deepStrictEqual(l.rows.map((r) => r.key), ["A330-1-002.-", "A330-1-001.-"]); });
test("glanceList", "by station", () => { assert.deepStrictEqual(F.glanceList("conv", GL, NOV6).rows.map((r) => r.key), ["A330-1-001.-"]); assert.deepStrictEqual(F.glanceList("station:9 - Final Approval", GL, NOV6).rows.map((r) => r.key), ["A330-1-002.-"]); });
test("glanceList", "released", () => assert.deepStrictEqual(F.glanceList("released", GL, NOV6).rows.map((r) => r.key), ["A330-1-003.-"]));
test("glanceList", "days and exactness on each row", () => { const r = F.glanceList("fa", GL, NOV6).rows[0]; assert.deepStrictEqual([r.days, r.exact, r.row], [5, true, 3]); });
test("glanceList", "stuck uses the least it can be", () => { assert.deepStrictEqual(F.glanceList("stuck", GL, NOV6).rows, []); assert.strictEqual(F.glanceList("stuck", GL, new Date(2026, 10, 25)).rows.length, 2); });
test("glanceList", "CDFs not in GCC are named", () => assert.deepStrictEqual(F.glanceList("open", Object.assign({ gccKeys: { "A330-1-001": true } }, GL), NOV6).notInGcc, ["A330-1-002.-"]));
test("stuckCount", "nothing on day one", () => assert.deepStrictEqual(F.stuckCount(GL.stations, NOV6), { exact: 0, atLeast: 0, total: 0 }));
test("stuckCount", "two weeks later", () => assert.deepStrictEqual(F.stuckCount(GL.stations, new Date(2026, 10, 20)), { exact: 1, atLeast: 1, total: 2 }));
test("stuckCount", "exact only", () => assert.deepStrictEqual(F.stuckCount({ a: ["x", "2026-10-01", 1, "2026-11-06"] }, NOV6), { exact: 1, atLeast: 0, total: 1 }));
test("stuckCount", "empty", () => assert.deepStrictEqual(F.stuckCount({}, NOV6), { exact: 0, atLeast: 0, total: 0 }));
test("stuckCount", "no map", () => assert.deepStrictEqual(F.stuckCount(null, NOV6), { exact: 0, atLeast: 0, total: 0 }));
test("bridgeStatus", "carries the station date", () => {
  const b = F.bridgeStatus(TR_CDF, null, NOV6, "x", GL.stations), o = L.rowObj(b.rows[1]);
  assert.deepStrictEqual([o.since, o.exact, o.seen], ["2026-11-01", 1, "2026-11-06"]);
  assert.strictEqual(L.stationDays(o, NOV6).days, 5);
});

const counts = Object.entries(perFn).filter(([, n]) => n < 5);
if (counts.length) process.exitCode = 1;
console.log(`\n${passed} tests passed, ${failed} failed` + (counts.length ? " — fewer than 5 cases: " + counts.map(([f, n]) => f + " (" + n + ")").join(", ") : " — every function has 5+ cases"));
