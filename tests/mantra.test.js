// Run: node tests/mantra.test.js   (synthetic data only — no real names or documents)
// Stage 2 of the CDF algorithm: Mantra → CDF (cdf-mantra.js), the row operations (110), their log and undo,
// the people lines of the workbook settings. At least 5 cases for every function.
const assert = require("assert");
const M = require("../cdf-mantra.js");
const A = require("../cdf-algo.js");
const E = require("../engine.js");
const F = require("../features.js");

let passed = 0, failed = 0;
const perFn = {};
function test(fn, name, body) {
  perFn[fn] = (perFn[fn] || 0) + 1;
  try { body(); passed++; console.log("  ✓ " + fn + ": " + name); }
  catch (e) { failed++; console.error("  ✗ " + fn + ": " + name + "\n    " + e.stack); process.exitCode = 1; }
}
const C = A.col;

// ---------------------------------------------------------------- fixtures
const MANTRA_HEAD = ["Doc ID", "Rev", "Status", "Date Status", "# Rejects", "Rejected By Emp", "Rejected By Function", "Pending To Emp", "Pending To Function"];
const MANTRA = [["Mantra export"], MANTRA_HEAD,
  ["330-53-14-J2400", "A", "Released", "2026-10-01 09:00", 0, "", "", "", ""],
  ["330-53-14-J2400", "B", "In-Work WF", "2026-10-20 11:00", 0, "", "", "Dana Levi, Dana Levi", "Check Design (SIE)"],
  ["330-53-14-J2401", "-", "Released", "2026-10-03 10:00", 0, "", "", "", ""],
  ["330-53-15-J2654", "C", "Pre-Released Rejected", "2026-10-05", 1, "Ron Avi", "Stress", "", ""],
  ["330-53-16-J4521", "A", "Released", "2026-10-03 16:00", 0, "", "", "", ""],
  ["330-00-00-JC158", "-", "Released", "2026-09-01", 0, "", "", "", ""],
  ["330-92-80-H5716VB-101", "A", "Pre-Released WF", "2026-10-10", 0, "", "", "Tom Bar", "Approval"],
  ["330-92-80-H6000-101", "-", "Released", "2026-08-01", 0, "", "", "", ""],
  ["330-92-80-H6000-102", "-", "Released", "2026-08-01", 0, "", "", "", ""]];
const REV = [["Doc ID", "Rev", "Release Date", "Function", "Approver", "Sign Date"], ["330-00-00-J0117", "A", "2025-01-01", "", "", ""], ["330-00-00-J0117", "B", "2025-02-02", "", "", ""]];
const CAT = M.readMantra(MANTRA), MREV = M.readMantraRev(REV);
function row(key, R, extra) { const r = new Array(34).fill(""); r[C("C")] = key; r[C("R")] = R; Object.keys(extra || {}).forEach((k) => { r[C(k)] = extra[k]; }); return r; }
const CDF = [new Array(34).fill("h"),
  row("A330-1-001.-", "Released", { U: "3 - PRE-RELEASE", AA: "330-53-14-J2401", AB: "-" }),
  row("A330-1-002.-", "Open", {}),
  row("A330-1-003.-", "Cancelled", { U: "2 - IN WORK", AA: "330-1", AB: "A" }),
  row("A330-1-004.-", "", { U: "2 - IN WORK", V: "IAI", AA: "330-53-14-J2400", AB: "B" }),
  row("A330-1-005.-", "Open", {}),
  row("A330-1-006.-", "Open", {}),
  row("A330-1-007.-", "Open", { U: "2 - IN WORK", AA: "330-53-14-J2400", AB: "B" }),
  row("", "Open", { U: "4 - RELEASED", AA: "330-53-14-J2401", AB: "-" }),
  row("", "Open", { U: "4 - RELEASED", AA: "330-53-16-J4521", AB: "A" })];
const g = (k, s) => { const r = new Array(40).fill(""); r[0] = k; r[18] = s; return r; };
const GCC = [new Array(40).fill("hdr"), g("A330-1-001.-", "330-53-14-J2401"), g("A330-1-002.-", "See 330-53-14-J2400 Rev B and 330-53-14-J2401"), g("A330-1-003.-", "x"),
  g("A330-1-005.-", ""), g("A330-1-006.-", "Accepted as is"), g("A330-1-007.-", "330-53-14-J2400")];
const CFG = { personGroups: [{ group: "IAI - Design Engineering", people: ["Dana Levi"] }, { group: "IAI - Stress Department", people: ["Ron Avi"] }, { group: "Vendor", people: ["Tom Bar"] }], personAliases: [{ from: "Levi Dana", to: "Dana Levi" }] };
const RUN = () => M.run({ cdf: CDF, gcc: GCC, mantra: MANTRA, mantraRev: REV }, CFG, A);
const PEOPLE = M.makePeople(CFG);
const ctx = (key, drawing) => ({ key: key || "A330-1-100.-", drawing: drawing || "" });

// ---------------------------------------------------------------- small helpers
test("normId", "trims and upper-cases", () => assert.strictEqual(M.normId("  330-53-14-j2400 "), "330-53-14-J2400"));
test("normId", "typographic dashes become -", () => assert.strictEqual(M.normId("330–53—14‑J2400"), "330-53-14-J2400"));
test("normId", "no-break spaces", () => assert.strictEqual(M.normId(" 330-53-14-J2400 "), "330-53-14-J2400"));
test("normId", "empty and null", () => { assert.strictEqual(M.normId(""), ""); assert.strictEqual(M.normId(null), ""); });
test("normId", "numbers", () => assert.strictEqual(M.normId(330), "330"));

test("revKey", "- and NEW are the first revision", () => { assert.strictEqual(M.revKey("-"), 0); assert.strictEqual(M.revKey("new"), 0); assert.strictEqual(M.revKey(""), 0); });
test("revKey", "letters in order", () => assert.ok(M.revKey("A") < M.revKey("B") && M.revKey("B") < M.revKey("Z")));
test("revKey", "AA comes after Z", () => assert.ok(M.revKey("AA") > M.revKey("Z") && M.revKey("AB") > M.revKey("AA")));
test("revKey", "unknown → null", () => { assert.strictEqual(M.revKey("A1"), null); assert.strictEqual(M.revKey("REV-B"), null); });
test("revKey", "case and spaces don't matter", () => assert.strictEqual(M.revKey(" c "), M.revKey("C")));

test("revShow", "empty → -", () => assert.strictEqual(M.revShow(""), "-"));
test("revShow", "NEW → -", () => assert.strictEqual(M.revShow("New"), "-"));
test("revShow", "letters upper-cased", () => assert.strictEqual(M.revShow(" b "), "B"));
test("revShow", "- stays", () => assert.strictEqual(M.revShow("-"), "-"));
test("revShow", "two letters", () => assert.strictEqual(M.revShow("aa"), "AA"));

test("toSerial", "serial numbers pass through", () => assert.strictEqual(M.toSerial(46300.5), 46300.5));
test("toSerial", "ISO date", () => assert.strictEqual(M.toSerial("2026-11-05"), 46331));
test("toSerial", "ISO date and time", () => assert.strictEqual(M.toSerial("2026-11-05T12:00"), 46331.5));
test("toSerial", "dd/mm/yyyy hh:mm", () => assert.strictEqual(M.toSerial("05/11/2026 18:00"), 46331.75));
test("toSerial", "nonsense and empty → null", () => { assert.strictEqual(M.toSerial("soon"), null); assert.strictEqual(M.toSerial(""), null); assert.strictEqual(M.toSerial(NaN), null); });

test("dayOf", "drops the time", () => assert.strictEqual(M.dayOf(46331.75), 46331));
test("dayOf", "whole day stays", () => assert.strictEqual(M.dayOf(46331), 46331));
test("dayOf", "floating error just under a day", () => assert.strictEqual(M.dayOf(46331.9999999999), 46332));
test("dayOf", "null → null", () => assert.strictEqual(M.dayOf(null), null));
test("dayOf", "undefined → null", () => assert.strictEqual(M.dayOf(undefined), null));

test("personKey", "upper case, single spaces", () => assert.strictEqual(M.personKey("  dana   levi "), "DANA LEVI"));
test("personKey", "tabs", () => assert.strictEqual(M.personKey("Dana\tLevi"), "DANA LEVI"));
test("personKey", "hyphens without spaces", () => assert.strictEqual(M.personKey("Ben - Ari"), "BEN-ARI"));
test("personKey", "typographic dash", () => assert.strictEqual(M.personKey("Ben–Ari"), "BEN-ARI"));
test("personKey", "empty", () => assert.strictEqual(M.personKey(""), ""));

test("baseKey", "drops .-", () => assert.strictEqual(M.baseKey("A330-1-961.-"), "A330-1-961"));
test("baseKey", "keeps a letter revision", () => assert.strictEqual(M.baseKey("A330-1-740.A"), "A330-1-740.A"));
test("baseKey", "spaces and case", () => assert.strictEqual(M.baseKey(" a330-1-961 .- "), "A330-1-961"));
test("baseKey", "trailing dot", () => assert.strictEqual(M.baseKey("A330-1-961."), "A330-1-961"));
test("baseKey", "a note in brackets", () => assert.strictEqual(M.baseKey("A330-1-961(old)"), "A330-1-961"));

// ---------------------------------------------------------------- the sheets
test("readMantra", "finds the header row under a title", () => { assert.strictEqual(CAT.ok, true); assert.strictEqual(CAT.headerRow, 1); });
test("readMantra", "groups revisions of one document", () => assert.deepStrictEqual(Object.keys(CAT.docs["330-53-14-J2400"].revs).sort(), ["A", "B"]));
test("readMantra", "reads dates, rejects and people", () => { const r = CAT.docs["330-53-15-J2654"].revs.C[0]; assert.deepStrictEqual([r.status, r.rejects, r.rejBy, r.date], ["PRE-RELEASED REJECTED", 1, "Ron Avi", 46300]); });
test("readMantra", "no header → not ok", () => { const r = M.readMantra([["a", "b"]]); assert.strictEqual(r.ok, false); assert.ok(/header row/.test(r.errors[0])); });
test("readMantra", "a missing column is named", () => { const r = M.readMantra([["Doc ID", "Rev", "Status", "Date Status"], ["x", "A", "Released", ""]]); assert.strictEqual(r.ok, false); assert.ok(/# REJECTS/.test(r.errors[0])); });
test("readMantra", "a row with content but no id is an error", () => { const r = M.readMantra([MANTRA_HEAD, ["", "A", "Released", "2026-01-01", 0, "", "", "", ""]]); assert.strictEqual(r.ok, false); });
test("readMantra", "empty sheet", () => assert.strictEqual(M.readMantra([]).ok, false));

test("readMantraRev", "counts rows", () => assert.strictEqual(MREV.count, 2));
test("readMantraRev", "groups by document", () => assert.strictEqual(MREV.docs["330-00-00-J0117"].length, 2));
test("readMantraRev", "release date as serial", () => assert.strictEqual(MREV.docs["330-00-00-J0117"][1].release, M.toSerial("2025-02-02")));
test("readMantraRev", "no header → error, nothing read", () => { const r = M.readMantraRev([["x"]]); assert.strictEqual(r.count, 0); assert.strictEqual(r.errors.length, 1); });
test("readMantraRev", "empty sheet → no error", () => assert.deepStrictEqual(M.readMantraRev([]).errors, []));

// ---------------------------------------------------------------- 20.4 the answer
test("decodeAnswer", "empty answer", () => assert.strictEqual(M.decodeAnswer("  ", "A330-1-001.-").kind, "empty"));
test("decodeAnswer", "words only → answer", () => assert.strictEqual(M.decodeAnswer("Accepted as is per attached file", "A330-1-001.-").kind, "answer"));
test("decodeAnswer", "two documents and a requested revision", () => {
  const d = M.decodeAnswer("See 330-53-14-J2400 Rev B and 330-53-14-J2401", "A330-1-002.-");
  assert.deepStrictEqual(d.docs.map((x) => [x.id, x.requested]), [["330-53-14-J2400", "B"], ["330-53-14-J2401", ""]]);
});
test("decodeAnswer", "item suffix is not part of the id", () => { const d = M.decodeAnswer("330-53-14-J1799-001", "A330-1-001.-").docs[0]; assert.deepStrictEqual([d.id, d.itemSuffix], ["330-53-14-J1799", "001"]); });
test("decodeAnswer", "fixed source error 30-53-… → 330-53-…", () => { const d = M.decodeAnswer("30-53-14-J1799-001", "A330-1-001.-"); assert.strictEqual(d.docs[0].id, "330-53-14-J1799"); assert.ok(d.notes[0].indexOf("20.4.3") === 0); });
test("decodeAnswer", "the fix doesn't touch a longer id", () => assert.strictEqual(M.decodeAnswer("330-53-14-J1799-001", "A330-1-001.-").notes.length, 0));
test("decodeAnswer", "R01REV A → R01 with revision A", () => { const d = M.decodeAnswer("330-53-14-J1035-R01REV A", "A330-1-001.-").docs[0]; assert.deepStrictEqual([d.id, d.requested], ["330-53-14-J1035-R01", "A"]); });
test("decodeAnswer", "manual fix for one CDF only", () => {
  assert.strictEqual(M.decodeAnswer("330-92-80-JF1313-R02", "A330-1-353.-").docs[0].id, "330-92-80-JF131-R02");
  assert.strictEqual(M.decodeAnswer("330-92-80-JF1313-R02", "A330-1-354.-").docs[0].id, "330-92-80-JF1313-R02");
});
test("decodeAnswer", "damaged characters are not an id", () => assert.strictEqual(M.decodeAnswer("330-53-99-J1.077", "A330-1-001.-").kind, "answer"));

test("parts", "plain id", () => assert.deepStrictEqual(M.parts("330-53-14-J2400"), { air: "330", ata1: "53", ata2: "14", core: "J2400", suffix: "", item: "" }));
test("parts", "with -R01 and item", () => { const p = M.parts("330-00-00-JC136-R01-001"); assert.deepStrictEqual([p.core, p.suffix, p.item], ["JC136", "R01", "001"]); });
test("parts", "H number with letters", () => assert.strictEqual(M.parts("330-92-80-H5716VB-101").core, "H5716VB"));
test("parts", "not an id → null", () => assert.strictEqual(M.parts("hello"), null));
test("parts", "lower case is normalised", () => assert.strictEqual(M.parts("330-53-14-j2400").core, "J2400"));

test("oneEdit", "one letter replaced", () => assert.strictEqual(M.oneEdit("16", "17"), true));
test("oneEdit", "two neighbours swapped", () => assert.strictEqual(M.oneEdit("53", "35"), true));
test("oneEdit", "one inserted", () => assert.strictEqual(M.oneEdit("5", "53"), true));
test("oneEdit", "two edits", () => assert.strictEqual(M.oneEdit("12", "34"), false));
test("oneEdit", "the same is not an edit", () => assert.strictEqual(M.oneEdit("14", "14"), false));

// ---------------------------------------------------------------- 25 identity
test("resolveIdentity", "exact", () => assert.deepStrictEqual(pick(M.resolveIdentity({ id: "330-53-14-J2400" }, CAT, MREV, ctx())), ["330-53-14-J2400", "EXACT"]));
test("resolveIdentity", "approved correction of one CDF", () => assert.deepStrictEqual(pick(M.resolveIdentity({ id: "330-91-00-JC158" }, CAT, MREV, ctx("A330-1-281.-"))), ["330-00-00-JC158", "CORRECTED_MANUAL"]));
test("resolveIdentity", "the same correction doesn't apply to another CDF", () => assert.notStrictEqual(M.resolveIdentity({ id: "330-91-00-JC158" }, CAT, MREV, ctx("A330-1-282.-")).match, "CORRECTED_MANUAL"));
test("resolveIdentity", "H written without its letters", () => assert.deepStrictEqual(pick(M.resolveIdentity({ id: "330-92-80-H5716-101" }, CAT, MREV, ctx())), ["330-92-80-H5716VB-101", "EXISTING_EQUIVALENT"]));
test("resolveIdentity", "H base without -1xx → -101", () => assert.deepStrictEqual(pick(M.resolveIdentity({ id: "330-92-80-H6000" }, CAT, MREV, ctx())), ["330-92-80-H6000-101", "EXISTING_EQUIVALENT_DEFAULT_101"]));
// 1.9.2 keeps each catalog's ids (and the candidates of 25.6) once read: the answers must stay the same, catalog by catalog
test("resolveIdentity", "asked twice, the same answer (ids kept from the first time)", () => { const a = pick(M.resolveIdentity({ id: "330-53-17-J4521" }, CAT, MREV, ctx())), b = pick(M.resolveIdentity({ id: "330-53-17-J4521" }, CAT, MREV, ctx())); assert.deepStrictEqual(a, b); });
test("resolveIdentity", "another catalog with a second candidate → AMBIGUOUS there, still corrected in the first", () => {
  const mh = [["Doc ID", "Rev", "Status", "Date Status", "# Rejects", "Rejected By Emp", "Rejected By Function", "Pending To Emp", "Pending To Function"]];
  const two = M.readMantra(mh.concat([["330-53-16-J4521", "-", "RELEASED", "2026-01-01", 0, "", "", "", ""], ["330-53-18-J4521", "-", "RELEASED", "2026-01-01", 0, "", "", "", ""]]));
  assert.strictEqual(M.resolveIdentity({ id: "330-53-17-J4521" }, two, MREV, ctx()).match, "AMBIGUOUS");
  assert.strictEqual(M.resolveIdentity({ id: "330-53-17-J4521" }, CAT, MREV, ctx()).match, "CORRECTED_AUTO");
});
test("resolveIdentity", "a catalog without the id's core → no candidate", () => {
  const mh = [["Doc ID", "Rev", "Status", "Date Status", "# Rejects", "Rejected By Emp", "Rejected By Function", "Pending To Emp", "Pending To Function"]];
  const none = M.readMantra(mh.concat([["330-11-11-J0001", "-", "RELEASED", "2026-01-01", 0, "", "", "", ""]]));
  assert.ok(/NOT_FOUND/.test(M.resolveIdentity({ id: "330-53-17-J4521" }, none, MREV, ctx()).match));
});
test("parts", "asked twice → the same parts", () => assert.deepStrictEqual(M.parts("330-00-00-JC136-R01-001"), M.parts("330-00-00-jc136-R01-001")));
test("resolveIdentity", "one ATA digit off → corrected automatically", () => assert.deepStrictEqual(pick(M.resolveIdentity({ id: "330-53-17-J4521" }, CAT, MREV, ctx())), ["330-53-16-J4521", "CORRECTED_AUTO"]));
test("resolveIdentity", "…unless the drawing says otherwise", () => assert.strictEqual(M.resolveIdentity({ id: "330-53-17-J4521" }, CAT, MREV, ctx("A330-1-100.-", "330-53-17-J4521-001-SM")).match, "DRAWING_CONFLICT"));
test("resolveIdentity", "only in MANTRAREV", () => assert.strictEqual(M.resolveIdentity({ id: "330-00-00-J0117" }, CAT, MREV, ctx()).match, "MANTRAREV_EXACT"));
test("resolveIdentity", "nowhere → not found", () => assert.strictEqual(M.resolveIdentity({ id: "330-11-11-J9999" }, CAT, MREV, ctx()).match, "NOT_FOUND_NO_CANDIDATE"));
test("resolveIdentity", "item suffix falls back to the id", () => assert.deepStrictEqual(pick(M.resolveIdentity({ id: "330-53-14-J2400-001" }, CAT, MREV, ctx())), ["330-53-14-J2400", "EXACT"]));
function pick(r) { return [r.id, r.match]; }

// ---------------------------------------------------------------- 30 revision, status, people
test("selectRecord", "the highest revision", () => assert.strictEqual(M.selectRecord("330-53-14-J2400", CAT).rev, "B"));
test("selectRecord", "not in Mantra", () => assert.ok(M.selectRecord("330-11-11-J9999", CAT).error));
test("selectRecord", "two rows of one revision → the latest Date Status", () => {
  const cat = M.readMantra([MANTRA_HEAD, ["X-1", "A", "In-Work", "2026-01-01", 0, "", "", "P1", "F"], ["X-1", "A", "Released", "2026-02-01", 0, "", "", "", ""]]);
  assert.strictEqual(M.selectRecord("X-1", cat).record.status, "RELEASED");
});
test("selectRecord", "same date, different status → conflict", () => {
  const cat = M.readMantra([MANTRA_HEAD, ["X-1", "A", "In-Work", "2026-01-01", 0, "", "", "P1", "F"], ["X-1", "A", "Released", "2026-01-01", 0, "", "", "", ""]]);
  assert.strictEqual(M.selectRecord("X-1", cat).conflict, true);
});
test("selectRecord", "revision that isn't letters → error", () => { const cat = M.readMantra([MANTRA_HEAD, ["X-2", "1", "Released", "2026-01-01", 0, "", "", "", ""]]); assert.ok(/not recognised/.test(M.selectRecord("X-2", cat).error)); });

test("uOf", "released", () => assert.strictEqual(M.uOf("RELEASED"), "4 - RELEASED"));
test("uOf", "pre-released wf", () => assert.strictEqual(M.uOf("PRE-RELEASED WF"), "3 - PRE-RELEASE"));
test("uOf", "pre-released rejected is back in work", () => assert.strictEqual(M.uOf("PRE-RELEASED REJECTED"), "2 - IN WORK"));
test("uOf", "in-work", () => assert.strictEqual(M.uOf("IN-WORK WF"), "2 - IN WORK"));
test("uOf", "unknown → null", () => assert.strictEqual(M.uOf("ARCHIVED"), null));

test("uvwx", "released → dashes", () => assert.deepStrictEqual(M.uvwx({ status: "RELEASED", rejects: 0 }, PEOPLE), { U: "4 - RELEASED", V: "-", W: "-", X: "-" }));
test("uvwx", "rejected: group, function, person", () => assert.deepStrictEqual(M.uvwx({ status: "PRE-RELEASED REJECTED", rejects: 1, rejBy: "Ron Avi", rejFn: "Stress" }, PEOPLE), { U: "2 - IN WORK", V: "IAI - Stress Department", W: "Rejected | Stress", X: "Ron Avi" }));
test("uvwx", "rejected without who → W Rejected", () => { const x = M.uvwx({ status: "IN-WORK REJECTED", rejects: 2, rejBy: "", rejFn: "" }, PEOPLE); assert.deepStrictEqual([x.W, x.X, x.note], ["Rejected", "-", "REJECTED_BY_MISSING_IN_SOURCE"]); });
test("uvwx", "pending: the same person twice is one", () => assert.deepStrictEqual(M.uvwx({ status: "IN-WORK WF", rejects: 0, pendEmp: "Dana Levi, Dana Levi", pendFn: "Check Design (SIE)" }, PEOPLE), { U: "2 - IN WORK", V: "IAI - Design Engineering", W: "Check Design (SIE)", X: "Dana Levi" }));
test("uvwx", "no people list → V null (left as it is)", () => assert.strictEqual(M.uvwx({ status: "IN-WORK WF", rejects: 0, pendEmp: "Dana Levi", pendFn: "F" }, M.makePeople({})).V, null));
test("uvwx", "unmapped status → error", () => assert.ok(M.uvwx({ status: "ARCHIVED", rejects: 0 }, PEOPLE).error));

test("makePeople", "empty config", () => assert.strictEqual(M.makePeople({}).empty, true));
test("makePeople", "names are keys", () => assert.strictEqual(PEOPLE.map["DANA LEVI"], "IAI - Design Engineering"));
test("makePeople", "one name in two groups → AMBIGUOUS", () => assert.strictEqual(M.makePeople({ personGroups: [{ group: "A", people: ["X Y"] }, { group: "B", people: ["X Y"] }] }).map["X Y"], "AMBIGUOUS"));
test("makePeople", "aliases", () => assert.strictEqual(PEOPLE.alias["LEVI DANA"], "DANA LEVI"));
test("makePeople", "null config", () => assert.strictEqual(M.makePeople(null).empty, true));

test("groupOf", "known", () => assert.strictEqual(M.groupOf("dana levi", PEOPLE), "IAI - Design Engineering"));
test("groupOf", "alias", () => assert.strictEqual(M.groupOf("Levi Dana", PEOPLE), "IAI - Design Engineering"));
test("groupOf", "two words the other way round", () => assert.strictEqual(M.groupOf("Avi Ron", PEOPLE), "IAI - Stress Department"));
test("groupOf", "unknown → UNMAPPED", () => assert.strictEqual(M.groupOf("Nobody Here", PEOPLE), "UNMAPPED"));
test("groupOf", "- and empty → -", () => { assert.strictEqual(M.groupOf("-", PEOPLE), "-"); assert.strictEqual(M.groupOf("", PEOPLE), "-"); });
test("groupOf", "no list → null", () => assert.strictEqual(M.groupOf("Dana Levi", M.makePeople({})), null));

test("solveDoc", "exact and released", () => { const o = M.solveDoc({ id: "330-53-14-J2401", original: "330-53-14-J2401" }, ctx(), CAT, MREV, PEOPLE); assert.deepStrictEqual([o.AA, o.AB, o.U, o.released, o.releaseDay], ["330-53-14-J2401", "-", "4 - RELEASED", true, 46298]); });
test("solveDoc", "pending person", () => { const o = M.solveDoc({ id: "330-53-14-J2400", original: "330-53-14-J2400" }, ctx(), CAT, MREV, PEOPLE); assert.deepStrictEqual([o.AB, o.U, o.X], ["B", "2 - IN WORK", "Dana Levi"]); });
test("solveDoc", "MANTRAREV only → its highest revision, released", () => { const o = M.solveDoc({ id: "330-00-00-J0117", original: "330-00-00-J0117" }, ctx(), CAT, MREV, PEOPLE); assert.deepStrictEqual([o.AB, o.U, o.V], ["B", "4 - RELEASED", "-"]); });
test("solveDoc", "not found → NO ENG. DOC.", () => assert.strictEqual(M.solveDoc({ id: "330-11-11-J9999", original: "330-11-11-J9999" }, ctx(), CAT, MREV, PEOPLE).U, "NO ENG. DOC."));
test("solveDoc", "conflict → CHECK SOLUTION with an error", () => { const o = M.solveDoc({ id: "330-53-17-J4521", original: "330-53-17-J4521" }, ctx("A330-1-100.-", "330-53-17-J4521-001-SM"), CAT, MREV, PEOPLE); assert.deepStrictEqual([o.U, /DRAWING_CONFLICT/.test(o.error)], ["CHECK SOLUTION", true]); });

// ---------------------------------------------------------------- run
const OUT = RUN();
const kinds = () => OUT.blocks.map((b) => b.key + ":" + b.kind);
test("run", "every kind of block", () => assert.deepStrictEqual(kinds(), ["A330-1-001.-:docs", "A330-1-002.-:docs", "A330-1-003.-:cancelled", "A330-1-004.-:no-r", "A330-1-005.-:empty", "A330-1-006.-:answer", "A330-1-007.-:docs"]));
test("run", "released block gets its closing date once", () => { const q = OUT.changes.filter((c) => c.colLetter === "Q"); assert.deepStrictEqual(q.map((c) => [c.key, c.newValue]), [["A330-1-001.-", 46298]]); });
test("run", "an existing closing date is never moved", () => { const cdf = CDF.map((r) => r.slice()); cdf[1][C("Q")] = 46000; const o = M.run({ cdf, gcc: GCC, mantra: MANTRA, mantraRev: REV }, CFG, A); assert.strictEqual(o.changes.filter((c) => c.colLetter === "Q").length, 0); });
test("run", "no closing date when the drawing differs from GCC (40.1)", () => {
  const cdf = CDF.map((r) => r.slice()); cdf[1][C("J")] = "330-53-14-J2401-001-SM";
  const o = M.run({ cdf, gcc: GCC, mantra: MANTRA, mantraRev: REV }, CFG, A);
  assert.strictEqual(o.changes.filter((c) => c.colLetter === "Q").length, 0); assert.ok(o.blocks[0].notes.some((n) => /drawing differs/.test(n)));
});
test("run", "no closing date from MANTRAREV alone (40.3)", () => {
  const gcc = GCC.map((r) => r.slice()); gcc[1][18] = "330-00-00-J0117";
  const o = M.run({ cdf: CDF, gcc, mantra: MANTRA, mantraRev: REV }, CFG, A);
  assert.strictEqual(o.changes.filter((c) => c.colLetter === "Q").length, 0); assert.ok(o.blocks[0].notes.some((n) => /MANTRAREV/.test(n)));
});
test("run", "no closing date when release days differ", () => {
  const gcc = GCC.map((r) => r.slice()); gcc[1][18] = "330-53-14-J2401 and 330-53-14-J2400 Rev A";
  const mantra = MANTRA.filter((r) => !(r[0] === "330-53-14-J2400" && r[1] === "B"));
  const o = M.run({ cdf: CDF, gcc, mantra, mantraRev: REV }, CFG, A);
  assert.strictEqual(o.changes.filter((c) => c.colLetter === "Q").length, 0); assert.ok(o.blocks[0].notes.some((n) => /release days differ/.test(n)));
});
test("run", "structure: one block short, one block long", () => assert.deepStrictEqual(OUT.structure, [{ key: "A330-1-002.-", row: 3, have: 1, need: 2 }, { key: "A330-1-007.-", row: 8, have: 3, need: 1 }]));
test("run", "cancelled → S:AB empty; no R → U:X empty", () => {
  const by = (k, L) => OUT.changes.filter((c) => c.key === k && c.colLetter === L)[0];
  assert.deepStrictEqual([by("A330-1-003.-", "AA").newValue, by("A330-1-003.-", "AA").rule, by("A330-1-004.-", "U").newValue, by("A330-1-004.-", "U").rule], ["", "00.6", "", "00.5"]);
  assert.strictEqual(by("A330-1-004.-", "AA"), undefined);
});
test("run", "no answer / words only", () => {
  const v = (k, L) => OUT.changes.filter((c) => c.key === k && c.colLetter === L)[0].newValue;
  assert.deepStrictEqual([v("A330-1-005.-", "U"), v("A330-1-006.-", "AA"), v("A330-1-006.-", "U")], ["NO ENG. DOC.", "ANSWER AT GCC", "ANSWER AT GCC"]);
});
test("run", "a Mantra sheet it can't read stops the stage", () => { const o = M.run({ cdf: CDF, gcc: GCC, mantra: [["x"]], mantraRev: REV }, CFG, A); assert.strictEqual(o.errors[0].stage, "00.1.2"); assert.strictEqual(o.changes.length, 0); });
test("run", "a CDF not in GCC is left as it is", () => { const o = M.run({ cdf: CDF, gcc: GCC.slice(0, 2), mantra: MANTRA, mantraRev: REV }, CFG, A); assert.strictEqual(o.blocks[1].kind, "no-gcc"); assert.ok(!o.changes.some((c) => c.key === "A330-1-002.-")); });
test("run", "no people list → V isn't touched", () => { const o = M.run({ cdf: CDF, gcc: GCC, mantra: MANTRA, mantraRev: REV }, {}, A); assert.ok(!o.changes.some((c) => c.key === "A330-1-002.-" && c.colLetter === "V")); });

// ---------------------------------------------------------------- 110 row operations
const PLAN = M.layout(OUT, CDF);
test("layout", "bottom-up: the removal below first, then the insert above", () => assert.deepStrictEqual(PLAN.ops.map((o) => [o.kind, o.key]), [["remove", "A330-1-007.-"], ["insert", "A330-1-002.-"]]));
test("layout", "insert: at the row under the block, row numbers after every operation", () => { const o = PLAN.ops[1]; assert.deepStrictEqual([o.at, o.count, o.finalAnchor, o.finalFirst], [4, 1, 3, 4]); });
test("layout", "remove: the extra rows and their values, for the log", () => { const o = PLAN.ops[0]; assert.deepStrictEqual([o.first, o.last, o.finalAnchor, o.finalFirst, o.removed[1].values.AA], [9, 10, 9, 10, "330-53-16-J4521"]); });
test("layout", "rows under an insert move down", () => { const c = PLAN.changes.filter((x) => x.key === "A330-1-006.-")[0]; assert.deepStrictEqual([c.sourceRow, c.row], [7, 8]); });
test("layout", "the new row carries its solution", () => assert.deepStrictEqual(PLAN.changes.filter((c) => c.newRow).map((c) => c.colLetter + "=" + c.newValue + "@" + c.row), ["U=4 - RELEASED@4", "V=-@4", "W=-@4", "X=-@4", "AA=330-53-14-J2401@4", "AB=-@4"]));
test("layout", "no cells for removed rows", () => assert.ok(!PLAN.changes.some((c) => c.key === "A330-1-007.-" && c.sourceRow >= 9)));
test("layout", "shift / isRemoved", () => { assert.deepStrictEqual([PLAN.shift(2), PLAN.shift(4), PLAN.shift(11)], [2, 5, 10]); assert.deepStrictEqual([PLAN.isRemoved(9), PLAN.isRemoved(8)], [true, false]); });
test("layout", "counts", () => assert.deepStrictEqual([PLAN.rowsAdded, PLAN.rowsRemoved], [1, 2]));
test("layout", "extra rows not together → error, no removal", () => {
  const o = { blocks: [{ key: "K", anchor: 1, rows: [1, 2, 4], want: [{}] }], structure: [{ key: "K", row: 2, have: 3, need: 1 }], changes: [], errors: [] };
  const p = M.layout(o, []); assert.strictEqual(p.ops.length, 0); assert.strictEqual(o.errors[0].stage, "110");
});
test("layout", "template: the block's own continuation row, else the nearest one above", () => {
  const o = { blocks: [{ key: "A", anchor: 1, rows: [1, 2], want: [] }, { key: "B", anchor: 3, rows: [3], want: [{}, { U: "x" }] }, { key: "C", anchor: 5, rows: [5, 6], want: [{}, {}, { U: "y" }] }],
    structure: [{ key: "B", row: 4, have: 1, need: 2 }, { key: "C", row: 6, have: 2, need: 3 }], changes: [], errors: [] };
  const p = M.layout(o, []); assert.deepStrictEqual(p.ops.map((x) => [x.key, x.template, x.own]), [["C", 7, 7], ["B", 3, null]]);
});

test("contFormula", "E", () => assert.strictEqual(M.contFormula("E", 7), '=IF($E$7="","",$E$7)'));
test("contFormula", "two-letter column", () => assert.strictEqual(M.contFormula("AH", 907), '=IF($AH$907="","",$AH$907)'));
test("contFormula", "reads back", () => assert.strictEqual(M.contAnchor(M.contFormula("R", 1197)), 1197));
test("contFormula", "every REF column", () => assert.ok(M.REF_COLS.every((L) => M.contAnchor(M.contFormula(L, 5)) === 5)));
test("contFormula", "MERGE and REF columns don't overlap and cover A:P, R:T, Y:Z, AC:AH", () => {
  const all = M.MERGE_COLS.concat(M.REF_COLS).sort();
  assert.strictEqual(new Set(all).size, all.length);
  ["A", "E", "H", "O", "P", "R", "T", "Y", "Z", "AC", "AH"].forEach((L) => assert.ok(all.indexOf(L) >= 0, L));
  ["Q", "U", "V", "W", "X", "AA", "AB"].forEach((L) => assert.ok(all.indexOf(L) < 0, L));
});

test("contAnchor", "anchor row", () => assert.strictEqual(M.contAnchor('=IF($R$845="","",$R$845)'), 845));
test("contAnchor", "two different cells → null", () => assert.strictEqual(M.contAnchor('=IF($R$845="","",$S$845)'), null));
test("contAnchor", "relative reference → null", () => assert.strictEqual(M.contAnchor('=IF(R845="","",R845)'), null));
test("contAnchor", "a value → null", () => { assert.strictEqual(M.contAnchor("Released"), null); assert.strictEqual(M.contAnchor(""), null); });
test("contAnchor", "spaces after commas", () => assert.strictEqual(M.contAnchor('=IF($E$7="", "", $E$7)'), 7));

test("applyInMemory", "writes changes into a copy", () => { const v = [["a", "b"]], c = M.applyInMemory(v, [{ row: 1, col: 1, newValue: "x" }]); assert.deepStrictEqual([c[0][1], v[0][1]], ["x", "b"]); });
test("applyInMemory", "grows rows and columns", () => assert.deepStrictEqual(M.applyInMemory([], [{ row: 2, col: 2, newValue: 1 }]), [[], ["", "", 1]]));
test("applyInMemory", "no changes → equal copy", () => assert.deepStrictEqual(M.applyInMemory([[1, 2]], []), [[1, 2]]));
test("applyInMemory", "null values and changes", () => assert.deepStrictEqual(M.applyInMemory(null, null), []));
test("applyInMemory", "the last change wins", () => assert.strictEqual(M.applyInMemory([["a"]], [{ row: 1, col: 0, newValue: "x" }, { row: 1, col: 0, newValue: "y" }])[0][0], "y"));

test("combine", "stage 1 cells move with the rows", () => { const c = M.combine([{ row: 7, col: C("S"), colLetter: "S", key: "A330-1-006.-", oldValue: "a", newValue: "b" }], PLAN).filter((x) => x.colLetter === "S")[0]; assert.deepStrictEqual([c.sourceRow, c.row], [7, 8]); });
test("combine", "stage 1 cells in removed rows are dropped", () => assert.ok(!M.combine([{ row: 9, col: C("S"), colLetter: "S", key: "K", oldValue: "a", newValue: "b" }], PLAN).some((c) => c.colLetter === "S")));
test("combine", "one cell written by both: before from stage 1, after from stage 2", () => {
  const c = M.combine([{ row: 2, col: C("U"), colLetter: "U", key: "A330-1-001.-", oldValue: "1 - X", newValue: "3 - PRE-RELEASE", rule: "00.6" }], PLAN).filter((x) => x.key === "A330-1-001.-" && x.colLetter === "U");
  assert.strictEqual(c.length, 1); assert.deepStrictEqual([c[0].oldValue, c[0].newValue], ["1 - X", "4 - RELEASED"]);
});
test("combine", "a cell that ends where it started is dropped", () => assert.ok(!M.combine([{ row: 2, col: C("U"), colLetter: "U", key: "K", oldValue: "4 - RELEASED", newValue: "x" }], PLAN).some((c) => c.row === 2 && c.colLetter === "U")));
test("combine", "sorted by row and column", () => { const c = M.combine([], PLAN); assert.ok(c.every((x, i) => !i || c[i - 1].row < x.row || (c[i - 1].row === x.row && c[i - 1].col <= x.col))); });

test("rowOpLogRows", "insert: one line, read back by undo", () => { const r = M.rowOpLogRows([PLAN.ops[1]], "s", "n"); assert.deepStrictEqual(r.map((x) => [x[2], x[3], x[4], x[5], x[7]]), [["CDF", "4", "A330-1-002.-", "ROWS", "+1 row · anchor 3"]]); });
test("rowOpLogRows", "remove: one line + the values of the removed rows", () => { const r = M.rowOpLogRows([PLAN.ops[0]], "s", "n"); assert.strictEqual(r[0][7], "-2 rows · anchor 9"); assert.strictEqual(r.filter((x) => x[8] === "Removed row (rule 110.2)").length, 6); });   // U, AA, AB of two rows
test("rowOpLogRows", "empty values of removed rows aren't logged", () => { const o = Object.assign({}, PLAN.ops[0], { removed: [{ row: 9, values: { U: "x", V: "", W: "", X: "", AA: "", AB: "" } }], count: 1 }); assert.strictEqual(M.rowOpLogRows([o], "s", "n").length, 2); });
test("rowOpLogRows", "the note says which solutions", () => { const o = Object.assign({}, PLAN.ops[1], { have: 1, need: 4, count: 3 }); assert.ok(/solution rows 2–4/.test(M.rowOpLogRows([o], "s", "n")[0][8])); });
test("rowOpLogRows", "no operations", () => { assert.deepStrictEqual(M.rowOpLogRows([], "s", "n"), []); assert.deepStrictEqual(M.rowOpLogRows(null, "s", "n"), []); });
test("rowOpLogRows", "a show() for values", () => { const r = M.rowOpLogRows([PLAN.ops[0]], "s", "n", (v) => "<" + v + ">"); assert.strictEqual(r[1][6], "<4 - RELEASED>"); });

test("parseRowsNote", "added", () => assert.deepStrictEqual(M.parseRowsNote("+3 rows · anchor 907"), { kind: "insert", count: 3, anchor: 907 }));
test("parseRowsNote", "removed, one row", () => assert.deepStrictEqual(M.parseRowsNote("-1 row · anchor 12"), { kind: "remove", count: 1, anchor: 12 }));
test("parseRowsNote", "not a rows note", () => assert.strictEqual(M.parseRowsNote("4 - RELEASED"), null));
test("parseRowsNote", "spaces around", () => assert.deepStrictEqual(M.parseRowsNote("  +2 rows · anchor 5 "), { kind: "insert", count: 2, anchor: 5 }));
test("parseRowsNote", "round trip with the log", () => { const r = M.rowOpLogRows(PLAN.ops, "s", "n").filter((x) => x[5] === "ROWS"); assert.deepStrictEqual(r.map((x) => M.parseRowsNote(x[7]).kind), ["remove", "insert"]); });

// ---------------------------------------------------------------- undo of row operations (features.js)
const LOGH = ["Run time", "Algorithm", "Sheet", "Row", "Key (CDF)", "Column", "Old value", "New value", "Note"];
const LOG = [LOGH].concat(M.rowOpLogRows(PLAN.ops, "06/11 07:00", "CDF algorithm · stage 2"), PLAN.changes.map((c) => ["06/11 07:00", "CDF algorithm · stage 2", "CDF", String(c.row), c.key, c.colLetter, String(c.oldValue), String(c.newValue), "Rule"]));
const UP = F.undoPlan(LOG);
test("undoPlan", "row operations come back bottom-up", () => assert.deepStrictEqual(UP.rowOps.map((o) => [o.kind, o.row, o.count, o.anchor]), [["remove", 10, 2, 9], ["insert", 4, 1, 3]]));
test("undoPlan", "cells of added rows go with the rows", () => assert.ok(!UP.entries.some((e) => e.row === 4)));
test("undoPlan", "values of removed rows ride with their operation", () => assert.strictEqual(UP.rowOps[0].values.length, 6));
test("undoPlan", "removed-row values don't mix with real cells of the same address", () => assert.ok(!UP.entries.some((e) => e.row >= 10 && e.key === "A330-1-007.-")));
test("undoPlan", "a run of only row operations can be undone", () => { const p = F.undoPlan([LOGH].concat(M.rowOpLogRows([PLAN.ops[1]], "t", "n"))); assert.strictEqual(p.entries.length, 0); assert.strictEqual(p.rowOps.length, 1); });

test("undoRowCheck", "added rows still refer to their anchor → ready", () => assert.strictEqual(F.undoRowCheck([UP.rowOps[1]], [[M.contFormula("R", 3)]]).ready.length, 1));
test("undoRowCheck", "added rows changed since → conflict", () => assert.strictEqual(F.undoRowCheck([UP.rowOps[1]], [["Open"]]).conflicts.length, 1));
test("undoRowCheck", "wrong number of rows → conflict", () => assert.strictEqual(F.undoRowCheck([UP.rowOps[1]], [[]]).conflicts.length, 1));
test("undoRowCheck", "removed rows: the anchor still holds the CDF → ready", () => assert.strictEqual(F.undoRowCheck([UP.rowOps[0]], ["A330-1-007.-"]).ready.length, 1));
test("undoRowCheck", "removed rows: another CDF at the anchor → conflict", () => assert.strictEqual(F.undoRowCheck([UP.rowOps[0]], ["A330-1-008.-"]).conflicts.length, 1));

test("undoRowLogRows", "one line per row operation", () => assert.strictEqual(F.undoRowLogRows(UP, UP.rowOps, "t").length, 2));
test("undoRowLogRows", "marks the run as undone", () => { const l = LOG.concat(F.undoRowLogRows(UP, UP.rowOps, "t"), F.undoLogRows(UP, UP.entries, "t")); assert.strictEqual(F.undoPlan(l), null); });
test("undoRowLogRows", "column ROWS, the note as written", () => { const r = F.undoRowLogRows(UP, [UP.rowOps[1]], "t")[0]; assert.deepStrictEqual([r[1], r[5], r[6]], [F.UNDO, "ROWS", "+1 row · anchor 3"]); });
test("undoRowLogRows", "nothing done → nothing logged", () => assert.deepStrictEqual(F.undoRowLogRows(UP, [], "t"), []));
test("undoRowLogRows", "refers to the run", () => assert.strictEqual(F.undoRowLogRows(UP, UP.rowOps, "t")[0][8], "Undo of " + UP.id));

test("contRef", "anchor row", () => assert.strictEqual(F.contRef('=IF($R$9="","",$R$9)'), 9));
test("contRef", "a value → null", () => assert.strictEqual(F.contRef("Open"), null));
test("contRef", "different cells → null", () => assert.strictEqual(F.contRef('=IF($R$9="","",$R$8)'), null));
test("contRef", "agrees with cdf-mantra", () => assert.strictEqual(F.contRef(M.contFormula("AH", 77)), M.contAnchor(M.contFormula("AH", 77))));
test("contRef", "empty", () => assert.strictEqual(F.contRef(""), null));

// ---------------------------------------------------------------- people lines in the settings (engine.js / features.js)
const PL = 'PERSON_GROUP "Design" "Dana Levi; Tom Bar"\nPERSON_GROUP "Stress" "Ron Avi"\nPERSON_ALIAS "Levi Dana" = "Dana Levi"';
test("PERSON_GROUP", "groups and names", () => assert.deepStrictEqual(E.parseRules(PL).cdfConfig.personGroups, [{ group: "Design", people: ["Dana Levi", "Tom Bar"] }, { group: "Stress", people: ["Ron Avi"] }]));
test("PERSON_GROUP", "the same group on two lines adds up", () => assert.deepStrictEqual(E.parseRules('PERSON_GROUP "D" "A B"\nPERSON_GROUP "D" "C D"').cdfConfig.personGroups, [{ group: "D", people: ["A B", "C D"] }]));
test("PERSON_GROUP", "no names → error with the line", () => { const r = E.parseRules('PERSON_GROUP "D" ""'); assert.strictEqual(r.errors.length, 1); assert.strictEqual(r.errors[0].line, 1); });
test("PERSON_GROUP", "aliases", () => assert.deepStrictEqual(E.parseRules(PL).cdfConfig.personAliases, [{ from: "Levi Dana", to: "Dana Levi" }]));
test("PERSON_GROUP", "allowed in the workbook settings, replaces the public list", () => {
  assert.deepStrictEqual(F.privateLines(PL), []);
  const m = F.mergeRules(E.parseRules('PERSON_GROUP "Old" "X Y"'), PL);
  assert.deepStrictEqual(m.cdfConfig.personGroups.map((x) => x.group), ["Design", "Stress"]);
});
test("PERSON_GROUP", "moved out of rules.txt with the other settings", () => assert.strictEqual(F.extractPrivate(PL + "\nLOG \"UPDATE LOG\"").moved.split("\n").length, 3));
test("PERSON_GROUP", "drives V in run()", () => { const o = M.run({ cdf: CDF, gcc: GCC, mantra: MANTRA, mantraRev: REV }, E.parseRules(PL).cdfConfig, A); assert.strictEqual(o.changes.filter((c) => c.key === "A330-1-002.-" && c.colLetter === "V")[0].newValue, "Design"); });

// ---------------------------------------------------------------- the morning routine with stage 2 (features.js)
const RES = (s2, s1) => ({ errors: [], checks: [], cdfChanges: new Array((s1 || 0) + (s2.cells || 0)).fill({}), gccChanges: [], stage1Cells: s1 || 0, stage2: s2, info: null });
test("morningDecision", "a normal Mantra day doesn't stop", () => assert.strictEqual(F.morningDecision(RES({ ran: true, cells: 314, errors: [], rowsAdded: 5 })).stop, false));
test("morningDecision", "many GCC cells still stop", () => assert.strictEqual(F.morningDecision(RES({ ran: true, cells: 10, errors: [] }, 60)).stop, true));
test("morningDecision", "Mantra solutions to check stop", () => assert.ok(F.morningDecision(RES({ ran: true, cells: 3, errors: [{ key: "K", message: "AMBIGUOUS" }] })).reasons[0].text.indexOf("1 Mantra solution to check") === 0));
test("morningDecision", "rows to remove stop", () => assert.ok(/removed/.test(F.morningDecision(RES({ ran: true, cells: 3, errors: [], rowsRemoved: 2 })).reasons[0].text)));
test("morningDecision", "an unreadable Mantra sheet stops (bad); a missing one doesn't", () => {
  assert.strictEqual(F.morningDecision(RES({ skipped: "Sheet Mantra can't be read", bad: true })).reasons[0].level, "bad");
  assert.strictEqual(F.morningDecision(RES({ skipped: "no sheet", missing: true })).stop, false);
});
test("morningDecision", "rows count as work", () => assert.strictEqual(F.morningDecision(RES({ ran: true, cells: 0, errors: [], rowsAdded: 2 })).rows, 2));

const counts = Object.entries(perFn).filter(([, n]) => n < 5);
if (counts.length) process.exitCode = 1;
console.log(`\n${passed} tests passed, ${failed} failed` + (counts.length ? " — fewer than 5 cases: " + counts.map(([f, n]) => f + " (" + n + ")").join(", ") : " — every function has 5+ cases"));
