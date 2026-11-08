// Run: node tests/features.test.js   (synthetic data only)
// At least 5 cases for every function in features.js and for the new engine/algorithm functions.
const assert = require("assert");
const fs = require("fs");
const path = require("path");
const E = require("../engine.js");
const F = require("../features.js");
const A = require("../cdf-algo.js");
// rules.txt no longer names people (FINAL_APPROVAL_RELEASE, PENDING_TABLE live in the workbook settings): invented ones for the tests
const TEST_PEOPLE = '\nFINAL_APPROVAL_RELEASE "Tom Bar"\nPENDING_TABLE "DANA PENDING TABLE" "Dana Ben-Ari"\nPENDING_TABLE "TOM PENDING TABLE" "Tom Bar"\n';
function rulesWithPeople() { return require("fs").readFileSync(require("path").join(__dirname, "..", "rules.txt"), "utf8") + TEST_PEOPLE; }


let passed = 0, failed = 0;
const perFn = {};
function test(fn, name, body) {
  perFn[fn] = (perFn[fn] || 0) + 1;
  try { body(); passed++; console.log("  ✓ " + fn + ": " + name); }
  catch (e) { failed++; console.error("  ✗ " + fn + ": " + name + "\n    " + e.stack); process.exitCode = 1; }
}
const NOW = new Date(2026, 10, 6, 9, 30).getTime();
const C = A.col;
function row(obj, w) { const r = new Array(w || 34).fill(""); for (const [k, v] of Object.entries(obj)) r[C(k)] = v; return r; }

// ---------------- precheckImport / worstLevel ----------------
const H = ["CDF Number", "CDF Status", "Current Station"];
const cur = { rows: 1001, headers: H };
test("precheckImport", "a good file passes", () => {
  const r = F.precheckImport({ expect: "CDF_List_Export", fileName: "CDF_List_Export (19).xlsx", lastModified: NOW - 3600e3, now: NOW, sheet: { rows: 1006, headers: H }, current: cur });
  assert.deepStrictEqual(r.map((x) => x.level), ["ok"]); assert.match(r[0].text, /1005 data rows \(now 1000\)/);
});
test("precheckImport", "an empty file stops", () => {
  const r = F.precheckImport({ sheet: { rows: 1, headers: H }, current: cur });
  assert.strictEqual(r[0].level, "bad"); assert.match(r[0].text, /no data rows/);
});
test("precheckImport", "an old file is a warning", () => {
  const r = F.precheckImport({ lastModified: NOW - 3 * 86400e3, now: NOW, sheet: { rows: 1001, headers: H }, current: cur });
  assert.strictEqual(r[0].level, "warn"); assert.match(r[0].text, /3 days ago \(03\/11\/2026\)/);
});
test("precheckImport", "many rows fewer stops, a few is a warning", () => {
  assert.strictEqual(F.precheckImport({ sheet: { rows: 500, headers: H }, current: cur })[0].level, "bad");
  const w = F.precheckImport({ sheet: { rows: 998, headers: H }, current: cur });
  assert.strictEqual(w[0].level, "warn"); assert.match(w[0].text, /3 rows fewer/);
});
test("precheckImport", "missing columns stop; other order / case is fine", () => {
  const r = F.precheckImport({ sheet: { rows: 1001, headers: ["cdf number", "Current Station"] }, current: cur });
  assert.strictEqual(r[0].level, "bad"); assert.match(r[0].text, /Column missing in the file: CDF Status/);
  assert.strictEqual(F.precheckImport({ sheet: { rows: 1001, headers: [" current station", "CDF STATUS", "cdf number"] }, current: cur })[0].level, "ok");
});
test("precheckImport", "wrong file name is a warning; no current sheet is fine", () => {
  const r = F.precheckImport({ expect: "650", fileName: "CDF_List_Export.xlsx", sheet: { rows: 10, headers: H }, current: null });
  assert.deepStrictEqual(r.map((x) => x.level), ["warn"]); assert.match(r[0].text, /Expected a file named 650/);
});
test("precheckImport", "a nearly empty file says 99%, not 100%", () => {
  const r = F.precheckImport({ sheet: { rows: 5, headers: H }, current: cur });
  assert.match(r[0].text, /Only 4 data rows — the sheet has 1000 now \(99% fewer\)/);
});
test("precheckImport", "a long list of missing columns is cut short", () => {
  const many = H.concat(["A1", "A2", "A3", "A4", "A5", "A6", "A7"]);
  const r = F.precheckImport({ sheet: { rows: 1001, headers: ["CDF Number"] }, current: { rows: 1001, headers: many } });
  assert.match(r[0].text, /^Columns missing in the file: CDF Status, Current Station, A1, A2, A3 and 4 more\.$/);
});
test("worstLevel", "bad wins", () => assert.strictEqual(F.worstLevel([{ level: "ok" }, { level: "bad" }, { level: "warn" }]), "bad"));
test("worstLevel", "warn over ok", () => assert.strictEqual(F.worstLevel([{ level: "ok" }, { level: "warn" }]), "warn"));
test("worstLevel", "ok only", () => assert.strictEqual(F.worstLevel([{ level: "ok" }]), "ok"));
test("worstLevel", "empty list is ok", () => assert.strictEqual(F.worstLevel([]), "ok"));
test("worstLevel", "undefined is ok", () => assert.strictEqual(F.worstLevel(undefined), "ok"));

// ---------------- grouped changes ----------------
const CH = [
  { sheet: "CDF", row: 9, col: 18, colLetter: "S", key: "X-2", oldValue: "a", newValue: "b", rule: "20.2" },
  { sheet: "CDF", row: 4, col: 19, colLetter: "T", key: "X-1", oldValue: "-", newValue: "Mor", rule: "20.2" },
  { sheet: "CDF", row: 4, col: 17, colLetter: "R", key: "X-1", oldValue: "Open", newValue: "Released", rule: "20.2" },
  { sheet: "GCC", row: 30, col: 38, colLetter: "AM", key: "X-1", oldValue: "Final Approval", newValue: "Released", rule: "10" },
  { sheet: "CDF", row: 9, col: 18, colLetter: "S", key: "X-2", oldValue: "b", newValue: "c", rule: "20.3.1" }
];
test("groupChangesByCdf", "one group per CDF and sheet", () => assert.strictEqual(F.groupChangesByCdf(CH).length, 3));
test("groupChangesByCdf", "GCC first, then CDF by row", () => assert.deepStrictEqual(F.groupChangesByCdf(CH).map((g) => g.sheet + g.row), ["GCC30", "CDF4", "CDF9"]));
test("groupChangesByCdf", "columns in sheet order inside a group", () => assert.deepStrictEqual(F.groupChangesByCdf(CH)[1].changes.map((c) => c.colLetter), ["R", "T"]));
test("groupChangesByCdf", "empty input", () => assert.deepStrictEqual(F.groupChangesByCdf([]), []));
test("groupChangesByCdf", "does not change the input order", () => { const copy = CH.slice(); F.groupChangesByCdf(CH); assert.deepStrictEqual(CH, copy); });
test("changeFacets", "counts per column in sheet order", () => assert.deepStrictEqual(F.changeFacets(CH).columns, [{ id: "R", n: 1 }, { id: "S", n: 2 }, { id: "T", n: 1 }, { id: "AM", n: 1 }]));
test("changeFacets", "counts per rule in natural order", () => assert.deepStrictEqual(F.changeFacets(CH).rules.map((r) => r.id), ["10", "20.2", "20.3.1"]));
test("changeFacets", "empty", () => assert.deepStrictEqual(F.changeFacets([]), { columns: [], rules: [] }));
test("changeFacets", "undefined", () => assert.deepStrictEqual(F.changeFacets(), { columns: [], rules: [] }));
test("changeFacets", "rule counts add up", () => assert.strictEqual(F.changeFacets(CH).rules.reduce((s, r) => s + r.n, 0), CH.length));
test("filterChanges", "by column", () => assert.strictEqual(F.filterChanges(CH, { col: "S" }).length, 2));
test("filterChanges", "by rule", () => assert.strictEqual(F.filterChanges(CH, { rule: "10" })[0].sheet, "GCC"));
test("filterChanges", "by CDF text", () => assert.strictEqual(F.filterChanges(CH, { text: "x-1" }).length, 3));
test("filterChanges", "column + rule together", () => assert.strictEqual(F.filterChanges(CH, { col: "S", rule: "20.3.1" }).length, 1));
test("filterChanges", "no filter keeps all", () => assert.strictEqual(F.filterChanges(CH, {}).length, 5));

// ---------------- undo ----------------
const LOGH = ["Run time", "Algorithm", "Sheet", "Row", "Key", "Column", "Before", "After", "Note"];
const LOG = [LOGH,
  ["05/11/2026 08:10", "CDF algorithm 2026-11-05 · stage 1", "CDF", "4", "X-1", "R", "Open", "Released", ""],
  ["05/11/2026 08:10", "CDF algorithm 2026-11-05 · stage 1", "GCC", "", "X-9", "", "", "", "Not found in base sheet"],
  ["06/11/2026 09:00", "New row", "CDF", "12", "X-7", "C", "", "X-7", "New row"],
  ["06/11/2026 09:20", "CDF algorithm 2026-11-05 · stage 1", "CDF", "9", "X-2", "S", "1 - Initiator", "2 - POC", ""],
  ["06/11/2026 09:20", "CDF algorithm 2026-11-05 · stage 1", "CDF", "9", "X-2", "T", "-", "Mor", ""],
  ["06/11/2026 09:20", "CDF algorithm 2026-11-05 · stage 1", "GCC", "30", "X-1", "AM", "Final Approval", "Released", "Rule 10"]];
test("undoPlan", "takes the last algorithm run", () => { const p = F.undoPlan(LOG); assert.strictEqual(p.stamp, "06/11/2026 09:20"); assert.strictEqual(p.entries.length, 3); });
test("undoPlan", "skips new rows and info lines", () => assert.ok(F.undoPlan(LOG).entries.every((e) => e.col && e.row)));
test("undoPlan", "a run already undone is skipped", () => {
  const p = F.undoPlan(LOG);
  const l2 = LOG.concat(F.undoLogRows(p, p.entries, "06/11/2026 09:40"));
  const p2 = F.undoPlan(l2); assert.strictEqual(p2.stamp, "05/11/2026 08:10"); assert.strictEqual(p2.entries.length, 1);
});
test("undoPlan", "nothing to undo", () => { assert.strictEqual(F.undoPlan([LOGH]), null); assert.strictEqual(F.undoPlan(null), null); });
test("undoPlan", "the same cell twice in a run goes back to the first value", () => {
  const p = F.undoPlan([LOGH, ["t", "Algo", "CDF", "5", "K", "S", "a", "b", ""], ["t", "Algo", "CDF", "5", "K", "S", "b", "c", ""]]);
  assert.deepStrictEqual(p.entries.map((e) => [e.before, e.after]), [["a", "c"]]);
});
test("undoCheck", "unchanged cells are ready", () => assert.strictEqual(F.undoCheck([{ after: "2 - POC" }], ["2 - POC"]).ready.length, 1));
test("undoCheck", "cells changed since are conflicts", () => { const r = F.undoCheck([{ after: "2 - POC", sheet: "CDF", row: 9, col: "S" }], ["Released"]); assert.strictEqual(r.conflicts[0].now, "Released"); });
test("undoCheck", "numbers compare as numbers", () => assert.strictEqual(F.undoCheck([{ after: "5" }], [5.0]).ready.length, 1));
test("undoCheck", "spaces around don't matter", () => assert.strictEqual(F.undoCheck([{ after: "Open" }], [" Open "]).ready.length, 1));
test("undoCheck", "mixed", () => { const r = F.undoCheck([{ after: "a" }, { after: "b" }, { after: "" }], ["a", "x", ""]); assert.deepStrictEqual([r.ready.length, r.conflicts.length], [2, 1]); });
test("valueFromShown", "integer", () => assert.strictEqual(F.valueFromShown("46000"), 46000));
test("valueFromShown", "decimal", () => assert.strictEqual(F.valueFromShown("1.5"), 1.5));
test("valueFromShown", "text stays text", () => assert.strictEqual(F.valueFromShown("2 - POC"), "2 - POC"));
test("valueFromShown", "empty stays empty", () => assert.strictEqual(F.valueFromShown(""), ""));
test("valueFromShown", "a date string stays text", () => assert.strictEqual(F.valueFromShown("05/11/2026"), "05/11/2026"));
test("undoLogRows", "one log row per cell", () => assert.strictEqual(F.undoLogRows(F.undoPlan(LOG), F.undoPlan(LOG).entries, "s").length, 3));
test("undoLogRows", "before / after are swapped", () => { const r = F.undoLogRows(F.undoPlan(LOG), F.undoPlan(LOG).entries, "s")[0]; assert.deepStrictEqual([r[6], r[7]], ["2 - POC", "1 - Initiator"]); });
test("undoLogRows", "run name is Undo", () => assert.ok(F.undoLogRows(F.undoPlan(LOG), F.undoPlan(LOG).entries, "s").every((r) => r[1] === "Undo")));
test("undoLogRows", "note names the undone run", () => assert.strictEqual(F.undoLogRows(F.undoPlan(LOG), F.undoPlan(LOG).entries, "s")[0][8], "Undo of 06/11/2026 09:20 CDF algorithm 2026-11-05 · stage 1"));
test("undoLogRows", "nothing done, nothing logged", () => assert.deepStrictEqual(F.undoLogRows(F.undoPlan(LOG), [], "s"), []));

// ---------------- lookup ----------------
const CDF = [row({ C: "LINK TO GCC" }),
  row({ C: "A330-1-855.-", R: "Open", S: "2 - POC" }),
  row({ C: "A330-1-856.-", D: "-", H: "HOLES MISSING", J: "DWG-1", R: "Open", S: "9 - Final Approval", T: "Dana Ben-Ari", U: "4 - RELEASED", AA: "DOC-1", AB: "-" }),
  row({ R: "Open", U: "3 - IN WORK", AA: "DOC-2", AB: "A" }),
  row({ C: "A330-1-856.A", R: "Open" }),
  row({ B: "A330-1-857.-" })];
const GCC = [row({ A: "CDF Number" }, 42), row({ A: "A330-1-856.-", J: "Open", AM: "Final Approval", AN: "'Dana Ben-Ari" }, 42)];
const LG = [LOGH, ["05/11 08:10", "Algo", "CDF", "3", "A330-1-856.-", "T", "-", "Dana Ben-Ari", ""], ["06/11 09:20", "Algo", "CDF", "3", "A330-1-856.-", "S", "3", "9 - Final Approval", ""]];
test("lookupCdf", "by full number", () => { const r = F.lookupCdf({ cdf: CDF, gcc: GCC, log: LG }, "A330-1-856.-"); assert.strictEqual(r.row, 3); assert.strictEqual(r.fields.signature, "Dana Ben-Ari"); });
test("lookupCdf", "by short number; a revision is listed as another hit", () => { const r = F.lookupCdf({ cdf: CDF }, "856"); assert.strictEqual(r.key, "A330-1-856.-"); assert.deepStrictEqual(r.others, ["A330-1-856.A (row 5)"]); });
test("lookupCdf", "number with a revision: 856.A and CDF 856.A", () => {
  assert.strictEqual(F.lookupCdf({ cdf: CDF }, "856.A").key, "A330-1-856.A");
  assert.strictEqual(F.lookupCdf({ cdf: CDF }, "CDF 856.A").key, "A330-1-856.A");
  assert.strictEqual(F.lookupCdf({ cdf: CDF }, "cdf#856").key, "A330-1-856.-");
});
test("lookupCdf", "a revision that isn't there", () => { const r = F.lookupCdf({ cdf: CDF }, "856.C"); assert.strictEqual(r.found, false); assert.ok(r.suggestions.includes("A330-1-856.A")); });
test("lookupCdf", "solution rows of the block", () => assert.deepStrictEqual(F.lookupCdf({ cdf: CDF }, "856").solutions.map((s) => [s.row, s.doc, s.rev, s.mantra]), [[3, "DOC-1", "-", "4 - RELEASED"], [4, "DOC-2", "A", "3 - IN WORK"]]));
test("lookupCdf", "GCC row and history (newest first)", () => {
  const r = F.lookupCdf({ cdf: CDF, gcc: GCC, log: LG }, "a330-1-856");
  assert.strictEqual(r.gcc.assignee, "Dana Ben-Ari"); assert.deepStrictEqual(r.history.map((h) => h.col), ["S", "T"]);
});
test("lookupCdf", "not found suggests close numbers", () => { const r = F.lookupCdf({ cdf: CDF }, "A330-1-85"); assert.strictEqual(r.found, false); assert.deepStrictEqual(r.suggestions, ["A330-1-855.-", "A330-1-856.-", "A330-1-856.A"]); });
test("lookupCdf", "empty query", () => assert.match(F.lookupCdf({ cdf: CDF }, " ").error, /Type a CDF number/));

// ---------------- glance / history ----------------
const cfg = E.parseRules(rulesWithPeople()).cdfConfig;
const SUM = { tables: [
  { name: "CDF STATUS", rows: [["Open", 224], ["Cancelled", 35], ["Released", 747], ["TOTAL", 1006]] },
  { name: "GCC CURRENT STATUS", rows: [["1 - Initiator", 8], ["2 - POC", 11], ["3 - Conversion Engineer", 103], ["4 - AVIONICS", 1], ["9 - Final Approval", 101], ["TOTAL", 224]] },
  { name: "DANA PENDING TABLE", rows: [["X", "d"], ["TOTAL", 12]] }, { name: "TOM PENDING TABLE", rows: [["TOTAL", 17]] }],
  pendingPeople: [{ table: "DANA PENDING TABLE", person: "Dana Ben-Ari" }, { table: "TOM PENDING TABLE", person: "Tom Bar" }] };
const S1 = F.snapshot(SUM, { gccRows: 1006, onlyCdf: ["A"] }, new Date(NOW));
test("snapshot", "status counts", () => assert.deepStrictEqual([S1.open, S1.released, S1.cancelled, S1.total], [224, 747, 35, 1006]));
test("snapshot", "stations without TOTAL", () => assert.strictEqual(Object.keys(S1.stations).length, 5));
test("snapshot", "pending per person", () => assert.deepStrictEqual(S1.pending, { "Dana Ben-Ari": 12, "Tom Bar": 17 }));
test("snapshot", "GCC rows and only-in-CDF list", () => { assert.strictEqual(S1.gccRows, 1006); assert.deepStrictEqual(S1.onlyCdf, ["A"]); });
test("snapshot", "time stamp", () => assert.strictEqual(S1.at, new Date(NOW).toISOString()));
const S0 = Object.assign({}, S1, { open: 223, released: 745, stations: Object.assign({}, S1.stations, { "9 - Final Approval": 99 }), pending: { "Dana Ben-Ari": 13, "Tom Bar": 18 } });
test("glance", "five metrics", () => assert.deepStrictEqual(F.glance(S1, S0).map((m) => m.id), ["open", "conv", "fa", "pending", "released"]));
test("glance", "deltas against the previous run", () => assert.deepStrictEqual(F.glance(S1, S0).map((m) => m.delta), [1, 0, 2, -2, 2]));
test("glance", "no previous run -> no deltas", () => assert.ok(F.glance(S1, null).every((m) => m.delta === null)));
test("glance", "values", () => assert.deepStrictEqual(F.glance(S1, null).map((m) => m.value), [224, 103, 101, 29, 747]));
test("glance", "missing stations count as 0", () => assert.strictEqual(F.glance(Object.assign({}, S1, { stations: {} }), null)[2].value, 0));
const at = (m) => ({ at: new Date(NOW + m * 60000).toISOString(), open: m });
test("pushSnapshot", "adds", () => assert.strictEqual(F.pushSnapshot([at(0)], at(5)).length, 2));
test("pushSnapshot", "same minute replaces", () => assert.strictEqual(F.pushSnapshot([at(0)], Object.assign(at(0), { open: 99 }))[0].open, 99));
test("pushSnapshot", "keeps the newest max entries", () => { const l = F.pushSnapshot([at(0), at(1), at(2)], at(3), 3); assert.deepStrictEqual(l.map((x) => x.open), [1, 2, 3]); });
test("pushSnapshot", "empty list", () => assert.strictEqual(F.pushSnapshot(null, at(1)).length, 1));
test("pushSnapshot", "does not change the input", () => { const l = [at(0)]; F.pushSnapshot(l, at(9)); assert.strictEqual(l.length, 1); });
test("previousSnapshot", "last one when it is older", () => assert.strictEqual(F.previousSnapshot([at(0), at(1)], at(5).at).open, 1));
test("previousSnapshot", "skips the current minute", () => assert.strictEqual(F.previousSnapshot([at(0), at(1)], at(1).at).open, 0));
test("previousSnapshot", "none", () => assert.strictEqual(F.previousSnapshot([], at(1).at), null));
test("previousSnapshot", "only the current one", () => assert.strictEqual(F.previousSnapshot([at(1)], at(1).at), null));
test("previousSnapshot", "no time given -> the last", () => assert.strictEqual(F.previousSnapshot([at(0), at(1)]).open, 1));
test("trendSeries", "last N values", () => assert.deepStrictEqual(F.trendSeries([at(1), at(2), at(3)], "open", 2).map((x) => x.value), [2, 3]));
test("trendSeries", "pending sums people", () => assert.strictEqual(F.trendSeries([S1], "pending")[0].value, 29));
test("trendSeries", "empty", () => assert.deepStrictEqual(F.trendSeries([], "open"), []));
test("trendSeries", "missing metric is 0", () => assert.strictEqual(F.trendSeries([at(1)], "released")[0].value, 0));
test("trendSeries", "keeps time stamps", () => assert.strictEqual(F.trendSeries([at(1)], "open")[0].at, at(1).at));
test("stationBars", "one bar per station", () => assert.strictEqual(F.stationBars(S1).length, 5));
test("stationBars", "longest bar is 100%", () => assert.strictEqual(F.stationBars(S1).find((b) => b.code === "3").pct, 100));
test("stationBars", "labels without the number", () => assert.strictEqual(F.stationBars(S1)[2].label, "Conversion Engineer"));
test("stationBars", "percent rounded", () => assert.strictEqual(F.stationBars(S1).find((b) => b.code === "9").pct, 98));
test("stationBars", "no stations", () => assert.deepStrictEqual(F.stationBars({ stations: {} }), []));

// ---------------- morning decision ----------------
const RES = { errors: [], checks: [{ ok: true, name: "x", detail: "" }], cdfChanges: new Array(6).fill({}), gccChanges: [], info: { gccRows: 1006, onlyCdf: ["A"] } };
test("morningDecision", "a normal day goes on", () => assert.deepStrictEqual(F.morningDecision(RES, { prev: S1 }), { stop: false, reasons: [], cells: 6, rows: 0, reports: false, history: false }));
test("morningDecision", "errors stop", () => assert.ok(F.morningDecision(Object.assign({}, RES, { errors: [{ message: "bad station" }] })).stop));
test("morningDecision", "failed control total stops", () => assert.match(F.morningDecision(Object.assign({}, RES, { checks: [{ ok: false, name: "TOTAL", detail: "1 vs 2" }] })).reasons[0].text, /Control total failed: TOTAL/));
test("morningDecision", "too many cells stop", () => assert.match(F.morningDecision(Object.assign({}, RES, { cdfChanges: new Array(41).fill({}) })).reasons[0].text, /41 cells/));
test("morningDecision", "CDFs gone from GCC stop", () => {
  const d = F.morningDecision(Object.assign({}, RES, { info: { gccRows: 1000, onlyCdf: ["A", "B"] } }), { prev: S1 });
  assert.deepStrictEqual(d.reasons.map((r) => r.text), ["6 CDFs fewer in GCC than in the last run (1006 → 1000).", "No longer in GCC: B."]);
});
test("morningDecision", "a bad file check stops", () => assert.ok(F.morningDecision(RES, { precheck: [{ level: "bad", text: "no rows" }] }).stop));
test("morningDecision", "the limit can be raised", () => assert.ok(!F.morningDecision(Object.assign({}, RES, { cdfChanges: new Array(41).fill({}) }), { maxCells: 100 }).stop));

// ---------------- workbook settings ----------------
const PUB = E.parseRules(rulesWithPeople());
test("mergeRules", "no workbook settings -> same rules", () => { const m = F.mergeRules(PUB, ""); assert.deepStrictEqual(m.cdfConfig, PUB.cdfConfig); assert.strictEqual(m.privateLines, 0); });
test("mergeRules", "stations are added / overridden", () => { const m = F.mergeRules(PUB, 'STATION "Weird" = "2 - POC"\nSTATION "POC" = "4 - AVIONICS"'); assert.strictEqual(m.cdfConfig.stations.Weird, "2 - POC"); assert.strictEqual(m.cdfConfig.stations.POC, "4 - AVIONICS"); });
test("mergeRules", "pending tables replace the public ones", () => assert.deepStrictEqual(F.mergeRules(PUB, 'PENDING_TABLE "A PENDING TABLE" "Person A"').cdfConfig.pendingTables, [{ name: "A PENDING TABLE", person: "Person A" }]));
test("mergeRules", "signer, groups and project", () => {
  const m = F.mergeRules(PUB, 'FINAL_APPROVAL_RELEASE "Person B"\nEMAIL_GROUP "Team" "a@b.com"\nCLAUDE_PROJECT "https://claude.ai/project/abc-1"');
  assert.deepStrictEqual([m.cdfConfig.finalApprovalSigner, m.emailGroups[0].name, m.claudeProject, m.privateLines], ["Person B", "Team", "https://claude.ai/project/abc-1", 3]);
});
test("mergeRules", "other commands are refused", () => { const m = F.mergeRules(PUB, 'IMPORT 3 "X"'); assert.match(m.errors.slice(-1)[0].message, /Workbook settings, line 1: IMPORT can't be in the workbook settings/); });
test("mergeRules", "the public rules are not changed", () => { const before = JSON.stringify(PUB); F.mergeRules(PUB, 'STATION "Z" = "2 - POC"'); assert.strictEqual(JSON.stringify(PUB), before); });
test("privateLines", "allowed lines", () => assert.deepStrictEqual(F.privateLines('STATION "a" = "b"\n# note\n\nSUPPLIER "x" = "y"'), []));
test("privateLines", "a forbidden line", () => assert.strictEqual(F.privateLines('LOG "X"')[0].line, 1));
test("privateLines", "lower case is still checked", () => assert.strictEqual(F.privateLines('import 1 "x"').length, 1));
test("privateLines", "comments only", () => assert.deepStrictEqual(F.privateLines("# a\n   # b"), []));
test("privateLines", "line numbers", () => assert.strictEqual(F.privateLines('STATION "a" = "b"\nNEWROW C = CDF')[0].line, 2));
const RAW = E.parseRules(require("fs").readFileSync(require("path").join(__dirname, "..", "rules.txt"), "utf8"));
test("requirePeopleSettings", "rules.txt alone (no people) → one settings error naming both lines", () => { const m = F.mergeRules(RAW, ""); const e = m.errors.filter((x) => x.setting); assert.ok(e.length === 1 && /FINAL_APPROVAL_RELEASE/.test(e[0].message) && /PENDING_TABLE/.test(e[0].message)); });
test("requirePeopleSettings", "only the signer given → asks for PENDING_TABLE only", () => { const e = F.mergeRules(RAW, 'FINAL_APPROVAL_RELEASE "Tom Bar"').errors.filter((x) => x.setting); assert.ok(e.length === 1 && !/FINAL_APPROVAL_RELEASE/.test(e[0].message) && /PENDING_TABLE/.test(e[0].message)); });
test("requirePeopleSettings", "both in the workbook settings → no error", () => assert.deepStrictEqual(F.mergeRules(RAW, 'FINAL_APPROVAL_RELEASE "Tom Bar"\nPENDING_TABLE "DANA PENDING TABLE" "Dana Ben-Ari"').errors, []));
test("requirePeopleSettings", "rules without CDF_ALGORITHM don't need them", () => assert.deepStrictEqual(F.requirePeopleSettings(E.parseRules('UPDATE "X"\nMATCH UPDATE.A = BASE.C\nCOPY UPDATE.B -> BASE.D')).errors, []));
test("requirePeopleSettings", "the message says where to paste them", () => assert.match(F.mergeRules(RAW, "").errors[0].message, /SETUP › Workbook settings/));
test("requirePeopleSettings", "the public rules.txt names nobody", () => { const t = require("fs").readFileSync(require("path").join(__dirname, "..", "rules.txt"), "utf8"); assert.doesNotMatch(t.replace(/#.*$/gm, ""), /FINAL_APPROVAL_RELEASE|PENDING_TABLE/); });
test("extractPrivate", "moves the settings lines", () => { const x = F.extractPrivate(rulesWithPeople()); assert.match(x.moved, /PENDING_TABLE "DANA PENDING TABLE"/); assert.doesNotMatch(x.rest, /^PENDING_TABLE/m); });
test("extractPrivate", "keeps imports and new-row lines", () => { const x = F.extractPrivate('IMPORT 1 "GCC"\nSTATION "a" = "b"\nNEWROW C = CDF'); assert.strictEqual(x.rest, 'IMPORT 1 "GCC"\nNEWROW C = CDF'); });
test("extractPrivate", "comments stay in the public file", () => assert.strictEqual(F.extractPrivate("# STATION is explained here").moved, ""));
test("extractPrivate", "the moved lines are valid workbook settings", () => assert.deepStrictEqual(F.privateLines(F.extractPrivate(rulesWithPeople()).moved), []));
test("extractPrivate", "public + moved parse to the same settings", () => {
  const t = rulesWithPeople(), x = F.extractPrivate(t);
  assert.deepStrictEqual(F.mergeRules(E.parseRules(x.rest), x.moved).cdfConfig, PUB.cdfConfig);
});
test("encodeStore/decodeStore", "round trip", () => assert.deepStrictEqual(F.decodeStore(F.encodeStore("urn:x", { a: [1, 2], b: "שלום" })), { a: [1, 2], b: "שלום" }));
test("encodeStore/decodeStore", "XML-unsafe text survives", () => assert.deepStrictEqual(F.decodeStore(F.encodeStore("urn:x", { t: "<a> & ]]> \"q\"" })), { t: "<a> & ]]> \"q\"" }));
test("encodeStore/decodeStore", "the namespace is on the root element", () => assert.match(F.encodeStore("urn:cdf-updater:settings:v1", {}), /^<cdfUpdater xmlns="urn:cdf-updater:settings:v1">/));
test("encodeStore/decodeStore", "garbage -> null", () => { assert.strictEqual(F.decodeStore("<x/>"), null); assert.strictEqual(F.decodeStore("<cdfUpdater>%E0%A4%A</cdfUpdater>"), null); });
test("encodeStore/decodeStore", "empty -> null", () => assert.strictEqual(F.decodeStore(""), null));

// ---------------- focused emails ----------------
const cdfSheet = [row({ C: "LINK TO GCC", P: "SEND TO", R: "CDF STATUS", S: "CURRENT STATION", T: "SIGNTURE POSITION", U: "STATUS (MANTRA)", Y: "DEVIATION REASON", Z: "MINOR / MAJOR", AA: "CDF SOLUTION" }),
  row({ C: "X-001.-", P: "ANN", H: "DESC 1", J: "DWG-1", K: "INSTL 1", R: "Open", S: "3 - Conversion Engineer", T: "Person S" }),
  row({ C: "X-002.-", P: "IAI", R: "Open", S: "3 - Conversion Engineer", T: "-" }),
  row({ C: "X-003.-", P: "SIE/ANN", R: "Open", S: "3 - Conversion Engineer", T: "-" })];
const gccSheet = [row({ A: "CDF Number", J: "CDF Status", P: "Deviation Reason", Q: "Deviation Classification", S: "Engineering Disposition", AK: "Final Approval Name", AM: "Current Station", AN: "Current Assignee" }, 42),
  row({ A: "X-001.-", J: "Open", AM: "Conversion ENG", AN: "Person S" }, 42), row({ A: "X-002.-", J: "Open", AM: "Conversion ENG", AN: "'-" }, 42), row({ A: "X-003.-", J: "Open", AM: "Conversion ENG", AN: "'-" }, 42)];
const R = A.run({ cdf: cdfSheet, gcc: gccSheet }, cfg);
test("focusKinds", "one per supplier and per pending table", () => assert.deepStrictEqual(F.focusKinds(R.summary).map((k) => k.id), ["sup:ANN/SIE", "sup:ANCARA", "sup:TIMAT", "sup:IAI", "pend:DANA PENDING TABLE", "pend:TOM PENDING TABLE"]));
test("focusKinds", "supplier label has the count", () => assert.match(F.focusKinds(R.summary)[0].label, /ANN\/SIE \(2\)/));
test("focusKinds", "no summary -> none", () => assert.deepStrictEqual(F.focusKinds(null), []));
test("focusKinds", "pending kinds carry the person", () => assert.strictEqual(F.focusKinds(R.summary)[4].person, "Dana Ben-Ari"));
test("focusKinds", "summary lists match the WHERE table", () => assert.strictEqual(R.summary.whereLists["ANN/SIE"].length, R.summary.tables.find((t) => t.name === "WHERE ARE THE CDFS").rows[0][1]));
test("buildFocusDoc", "supplier table rows + TOTAL", () => { const d = F.buildFocusDoc({ type: "supplier", supplier: "ANN/SIE" }, R.summary, new Date(NOW)); assert.deepStrictEqual(d.tables[0].rows.map((r) => r[0].v), ["X-001.-", "X-003.-", "TOTAL"]); });
test("buildFocusDoc", "supplier email is English, left to right", () => { const d = F.buildFocusDoc({ type: "supplier", supplier: "IAI" }, R.summary, new Date(NOW)); assert.strictEqual(d.dir, "ltr"); assert.match(d.opening, /^Hello, below are the open CDFs at Conversion Engineer for IAI — AS OF 06\/11\/2026 09:30/); });
test("buildFocusDoc", "supplier with nothing open is empty", () => assert.ok(F.buildFocusDoc({ type: "supplier", supplier: "TIMAT" }, R.summary, new Date(NOW)).empty));
test("buildFocusDoc", "pending email is Hebrew and greets the person", () => { const d = F.buildFocusDoc({ type: "pending", table: "DANA PENDING TABLE", person: "Dana Ben-Ari" }, R.summary, new Date(NOW)); assert.strictEqual(d.dir, "rtl"); assert.match(d.opening, /^שלום Dana,/); assert.ok(d.empty); });
test("buildFocusDoc", "renders with the daily email look", () => { const m = E.buildDailyEmail(F.buildFocusDoc({ type: "supplier", supplier: "ANN/SIE" }, R.summary, new Date(NOW))); assert.ok(m.html.includes('colspan="5"') && m.html.includes("#B10001") && m.html.includes('dir="ltr"')); });
test("buildFocusDoc", "subject has the date", () => assert.match(F.buildFocusDoc({ type: "pending", table: "TOM PENDING TABLE", person: "x" }, R.summary, new Date(NOW)).subject, / — 06\/11\/2026$/));

// ---------------- engine: daily email pieces (5+ each) ----------------
const V = [["T"], [], ["opening"], [], ["CDF STATUS"], ["CDF STATUS", "CDF'S"], ["Open", 3], ["TOTAL", 3], [], ["PERSONS – X"], ["PERSON", "N", "D"], ["P1", 2, "d1"], ["", "", "d2"], ["P2", 1, "d3"], [], ["Sig"]];
test("readSheetTables", "finds by title start", () => assert.strictEqual(E.readSheetTables(V, ["PERSONS"])[0].title, "PERSONS – X"));
test("readSheetTables", "stops at TOTAL", () => assert.strictEqual(E.readSheetTables(V, ["CDF STATUS"])[0].rows.length, 2));
test("readSheetTables", "stops at an empty row", () => assert.strictEqual(E.readSheetTables(V, ["PERSONS"])[0].rows.length, 3));
test("readSheetTables", "rowspan for continuation rows", () => { const r = E.readSheetTables(V, ["PERSONS"])[0].rows; assert.deepStrictEqual([r[0][0].rowspan, r[1][0], r[1][2].v], [2, null, "d2"]); });
test("readSheetTables", "missing table is skipped", () => assert.deepStrictEqual(E.readSheetTables(V, ["NOPE", "PERSONS"]).map((t) => t.title), ["PERSONS – X"]));
test("readSheetTables", "header width from the header row", () => assert.deepStrictEqual(E.readSheetTables(V, ["PERSONS"])[0].headers, ["PERSON", "N", "D"]));
const DOC = { subject: "S", opening: "בוקר טוב — AS OF 06/11/2026 09:30 (Asia/Jerusalem).", signature: ["בברכה,", "N"], tables: [{ title: "T1", headers: ["A", "B"], rows: [[{ v: "x", rowspan: 1 }, { v: "7", rowspan: 1 }]] }] };
test("buildDailyEmail", "subject is kept", () => assert.strictEqual(E.buildDailyEmail(DOC).subject, "S"));
test("buildDailyEmail", "AS OF stays left-to-right inside Hebrew", () => assert.ok(E.buildDailyEmail(DOC).html.includes('<span dir="ltr">AS OF 06/11/2026 09:30 (Asia/Jerusalem)</span>')));
test("buildDailyEmail", "numbers are centred", () => assert.ok(/text-align:center;">7<\/td>/.test(E.buildDailyEmail(DOC).html)));
test("buildDailyEmail", "text is escaped", () => assert.ok(E.buildDailyEmail(Object.assign({}, DOC, { opening: "<b>x</b>" })).html.includes("&lt;b&gt;x&lt;/b&gt;")));
test("buildDailyEmail", "plain text version", () => assert.strictEqual(E.buildDailyEmail(DOC).text, "בוקר טוב — AS OF 06/11/2026 09:30 (Asia/Jerusalem).\r\n\r\nT1\r\nA\tB\r\nx\t7\r\n\r\n\r\nבברכה,\r\nN"));
test("buildDailyEmail", "column widths when the table gives them", () => {
  const h = E.buildDailyEmail(Object.assign({}, DOC, { tables: [Object.assign({}, DOC.tables[0], { widths: [120, null] })] })).html;
  assert.ok(h.includes('<th width="120" style=') && h.includes("width:120px;min-width:120px;"));
  assert.strictEqual((h.match(/ width="/g) || []).length, 1);
});
test("buildFocusDoc", "supplier table has column widths, the long description widest", () => {
  const w = F.buildFocusDoc({ type: "supplier", supplier: "ANN/SIE" }, R.summary, new Date(NOW)).tables[0].widths;
  assert.strictEqual(w.length, 5); assert.strictEqual(Math.max.apply(null, w), w[4]);
});
test("buildDailyEmail", "left-to-right emails", () => assert.ok(E.buildDailyEmail(Object.assign({}, DOC, { dir: "ltr" })).html.includes('dir="ltr" style="font-family:Calibri,Arial,sans-serif;font-size:10pt;margin:0 0 12px;text-align:left"')));

const counts = Object.entries(perFn).filter(([, n]) => n < 5);
console.log(`\n${passed} tests passed, ${failed} failed` + (counts.length ? " — fewer than 5 cases: " + counts.map(([f, n]) => f + " (" + n + ")").join(", ") : " — every function has 5+ cases"));
if (counts.length) process.exitCode = 1;
