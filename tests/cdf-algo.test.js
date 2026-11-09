// Run: node tests/cdf-algo.test.js   (synthetic data only - no project data in this repository)
const assert = require("assert");
const fs = require("fs");
const path = require("path");
const A = require("../cdf-algo.js");
const E = require("../engine.js");
// rules.txt no longer names people (FINAL_APPROVAL_RELEASE, PENDING_TABLE live in the workbook settings): invented ones for the tests
const TEST_PEOPLE = '\nFINAL_APPROVAL_RELEASE "Tom Bar"\nPENDING_TABLE "DANA PENDING TABLE" "Dana Ben-Ari"\nPENDING_TABLE "TOM PENDING TABLE" "Tom Bar"\n';
function rulesWithPeople() { return require("fs").readFileSync(require("path").join(__dirname, "..", "rules.txt"), "utf8") + TEST_PEOPLE; }


let passed = 0;
function test(name, fn) {
  try { fn(); passed++; console.log("  ✓ " + name); }
  catch (e) { console.error("  ✗ " + name + "\n    " + e.stack); process.exitCode = 1; }
}
const C = A.col;
function row(obj, width) {
  const r = new Array(width).fill("");
  for (const [k, v] of Object.entries(obj)) r[C(k)] = v;
  return r;
}
const cdfHeader = row({ C: "LINK TO GCC", P: "SEND TO", R: "CDF STATUS\n(GCC)", S: "CURRENT STATION (GCC)", T: "SIGNTURE POSITION\n(GCC)",
  U: "STATUS (MANTRA)", Y: "DEVIATION REASON", Z: "MINOR / MAJOR", AA: "CDF SOLUTION", AB: "CDF SOLUTION\nREVISION" }, 34);
const gccHeader = row({ A: "CDF Number", J: "CDF Status", P: "Deviation Reason", Q: "Deviation Classification", S: "Engineering Disposition",
  AK: "Final Approval Name", AM: "Current Station", AN: "Current Assignee" }, 42);
const cfg = E.parseRules(rulesWithPeople()).cdfConfig;

// CDF sheet: 5 CDFs, one with two solution rows
const cdf = [
  cdfHeader,
  row({ C: "X-001.-", P: "IAI", R: "Open", S: "3 - Conversion Engineer", T: "-", Y: "-", Z: "-" }, 34),
  row({ C: "X-002.-", P: "ANN", R: "Open", S: "9 - Final Approval", T: "Dana Ben-Ari", Y: "Design Error", Z: "Minor", U: "4 - RELEASED", AA: "DOC-1", AB: "A" }, 34),
  row({ R: "Open", S: "9 - Final Approval", U: "4 - RELEASED", AA: "DOC-2", AB: "-" }, 34),            // solution row of X-002
  row({ C: "X-003.", P: "SIE/ANN", R: "Open", S: "3 - Conversion Engineer", T: "-", Y: "-", Z: "-" }, 34), // key without "-"
  row({ C: "X-004.-", R: "Open", S: "1 - Initiator", T: "-", U: "-", AA: "DOC-9", Y: "x", Z: "y" }, 34),
  row({ C: "A330-1-052 (another)", R: "Released", S: "Released" }, 34),
];
const gcc = [
  gccHeader,
  row({ A: "X-001.-", J: "Open", AM: "Conversion ENG", AN: "'-" }, 42),
  row({ A: "X-002.-", J: "Open", AM: "Final Approval", AN: "Dana Ben-Ari", P: "Design Error", Q: "Minor" }, 42),
  row({ A: "X-003.-", J: "Open", AM: "Conversion ENG", AN: "'-" }, 42),
  row({ A: "X-004.-", J: "Open", AM: "Initiator", AN: "'-", P: "x", Q: "y" }, 42),
];
function clone(x) { return JSON.parse(JSON.stringify(x)); }

console.log("stage 1: GCC -> CDF");
test("unchanged data -> no changes, checks pass", () => {
  const r = A.run({ cdf, gcc }, cfg);
  assert.deepStrictEqual(r.errors, []);
  assert.strictEqual(r.cdfChanges.length, 0, JSON.stringify(r.cdfChanges));
  assert.ok(r.checks.every((c) => c.ok), JSON.stringify(r.checks));
  assert.strictEqual(r.info.matched, 4);
  assert.deepStrictEqual(r.info.onlyCdf, []);
});
test("known blank CDFs do not count as missing from GCC; active CDFs still do", () => {
  const extra = [row({ C: "A330-1-763.-" }, 34), row({ C: "A330-1-878" }, 34),
    row({ C: "A330-1-878.A", R: "Open" }, 34), row({ C: "A330-1-999.-" }, 34)];
  assert.deepStrictEqual(A.run({ cdf: cdf.concat(extra), gcc }, cfg).info.onlyCdf, ["A330-1-878.A", "A330-1-999.-"]);
  extra[0][A.col("R")] = "Open";
  assert.deepStrictEqual(A.run({ cdf: cdf.concat(extra), gcc }, cfg).info.onlyCdf, ["A330-1-763.-", "A330-1-878.A", "A330-1-999.-"]);
});
test("summary tables", () => {
  const t = Object.fromEntries(A.run({ cdf, gcc }, cfg).summary.tables.map((x) => [x.name, x.rows]));
  assert.deepStrictEqual(t["CDF STATUS"], [["Open", 4], ["Cancelled", 0], ["Released", 0], ["TOTAL", 4]], "duplicate excluded");
  assert.deepStrictEqual(t["GCC CURRENT STATUS"].find((x) => x[0] === "3 - Conversion Engineer"), ["3 - Conversion Engineer", 2]);
  assert.deepStrictEqual(t["GCC SIGNATURE POSITION"], [["Dana Ben-Ari", 1], ["TOTAL", 1]]);
  assert.deepStrictEqual(t["DANA PENDING TABLE"], [["X-002.-", "DOC-1 Rev A\nDOC-2 Rev -"], ["TOTAL", 1]]);
  assert.deepStrictEqual(t["WHERE ARE THE CDFS"], [["ANN/SIE", 1], ["ANCARA", 0], ["TIMAT", 0], ["IAI", 1], ["TOTAL", 2]]);
});
test("GCC changes flow into R, S, T, Y, Z (anchor rows only)", () => {
  const g = clone(gcc);
  g[1][C("J")] = "Released"; g[1][C("AM")] = "Completed";                  // 20.3.1
  g[2][C("AN")] = "Dana Ben-Ari, Tom Bar"; g[2][C("P")] = "";          // two signatures, empty reason -> "-"
  g[4][C("J")] = "Cancelled";                                               // 00.6
  const r = A.run({ cdf, gcc: g }, cfg);
  const by = Object.fromEntries(r.cdfChanges.map((c) => [c.colLetter + c.row, c]));
  assert.strictEqual(by.R2.newValue, "Released");
  assert.strictEqual(by.S2.newValue, "Released"); assert.strictEqual(by.S2.rule, "20.3.1");
  assert.strictEqual(by.T3.newValue, "Dana Ben-Ari, Tom Bar");
  assert.strictEqual(by.Y3.newValue, "-");
  assert.ok(!r.cdfChanges.some((c) => c.row === 4), "solution row is never written");
  ["S", "T", "U", "Y", "Z", "AA"].forEach((k) => assert.strictEqual(by[k + "6"].newValue, "", "Cancelled clears " + k));
  const sig = Object.fromEntries(r.summary.tables.find((x) => x.name === "GCC SIGNATURE POSITION").rows);
  assert.strictEqual(sig["Tom Bar"], 1);
});
test("rule 10: Final Approval signed by the FINAL_APPROVAL_RELEASE person -> GCC!AM Released", () => {
  const g = clone(gcc);
  g[2][C("AK")] = "Tom Bar";
  const r = A.run({ cdf, gcc: g }, cfg);
  assert.deepStrictEqual(r.gccChanges.map((c) => [c.row, c.colLetter, c.oldValue, c.newValue]), [[3, "AM", "Final Approval", "Released"]]);
  assert.strictEqual(r.cdfChanges.find((c) => c.row === 3 && c.colLetter === "S").newValue, "Released");
  assert.ok(!r.cdfChanges.some((c) => c.colLetter === "R"), "R stays Open (AM alone is not a release)");
  assert.ok(!r.checks.find((c) => /stations/.test(c.name)).ok, "Open CDF in a non-open station is flagged");
});
test("unknown station is an error and is not written", () => {
  const g = clone(gcc); g[1][C("AM")] = "Mystery";
  const r = A.run({ cdf, gcc: g }, cfg);
  assert.ok(r.errors.some((e) => e.stage === "20.3" && /Mystery/.test(e.message)));
  assert.ok(!r.cdfChanges.some((c) => c.row === 2 && c.colLetter === "S"));
});
test("settings from rules.txt can add a station", () => {
  const g = clone(gcc); g[1][C("AM")] = "Mystery";
  const r = A.run({ cdf, gcc: g }, { ...cfg, stations: { ...cfg.stations, Mystery: "3 - Conversion Engineer" } });
  assert.deepStrictEqual(r.errors, []);
});
test("unknown status, keys only on one side, duplicate GCC keys", () => {
  const g = clone(gcc);
  g[3][C("J")] = "On hold";
  g.push(row({ A: "X-099.-", J: "Open", AM: "Initiator" }, 42));
  g.push(row({ A: "X-001.-", J: "Open", AM: "Initiator" }, 42));
  const r = A.run({ cdf, gcc: g }, cfg);
  assert.ok(r.errors.some((e) => /unknown CDF status/.test(e.message)));
  assert.ok(r.errors.some((e) => /Duplicate CDF numbers in GCC/.test(e.message)));
  assert.deepStrictEqual(r.info.onlyGcc, ["X-099.-"]);
});
test("wrong headers stop everything", () => {
  const c2 = clone(cdf); c2[0][C("R")] = "Something else";
  const r = A.run({ cdf: c2, gcc }, cfg);
  assert.ok(r.errors.length && r.errors[0].stage === "00.1");
  assert.strictEqual(r.cdfChanges.length, 0);
  assert.strictEqual(r.summary, null);
});
test("compareWithSheet finds tables by title and lists differences", () => {
  const r = A.run({ cdf, gcc }, cfg);
  const sheet = [["CDF STATUS"], ["CDF STATUS", "CDF'S"], ["Open", 3], ["Cancelled", 0], ["Released", 0], ["TOTAL", 3]];
  const cmp = A.compareWithSheet(r.summary, sheet);
  const st = cmp.find((t) => t.table === "CDF STATUS");
  assert.deepStrictEqual(st.diffs.map((d) => [d.label, d.current, d.computed]), [["Open", "3", "4"], ["TOTAL", "3", "4"]]);
  assert.strictEqual(cmp.find((t) => t.table === "WHERE ARE THE CDFS").found, false);
});
test("cross-check says why a CDF left / joined a pending table", () => {
  const sheet = [["DANA PENDING TABLE"], ["CDF NUMBER", "SOLUTION / DOCUMENT"], ["X-002.-", "DOC-1 Rev A\nDOC-2 Rev -"], ["TOTAL", 1]];
  const c2 = clone(cdf); c2[3][C("U")] = "3 - IN WORK";                         // solution row of X-002 not released
  const why1 = A.compareWithSheet(A.run({ cdf: c2, gcc }, cfg).summary, sheet).find((t) => t.table === "DANA PENDING TABLE").diffs.find((d) => d.label === "X-002.-");
  assert.strictEqual(why1.computed, "(not listed today)");
  assert.match(why1.why, /not all Mantra statuses \(U\) are RELEASED .*row 4: DOC-2 Rev - — U 3 - IN WORK/);
  const g = clone(gcc); g[2][C("AN")] = "Tom Bar";                        // signature moved
  const why2 = A.compareWithSheet(A.run({ cdf, gcc: g }, cfg).summary, sheet).find((t) => t.table === "DANA PENDING TABLE").diffs.find((d) => d.label === "X-002.-");
  assert.match(why2.why, /signature position is "Tom Bar" — not Dana Ben-Ari/);
  const ok = A.compareWithSheet(A.run({ cdf, gcc }, cfg).summary, [["DANA PENDING TABLE"], ["CDF NUMBER", "SOLUTION / DOCUMENT"], ["TOTAL", 0]])
    .find((t) => t.table === "DANA PENDING TABLE").diffs.find((d) => d.label === "X-002.-");
  assert.strictEqual(ok.current, "(not in the sheet)"); assert.match(ok.why, /^listed: Open, 9 - Final Approval, signature Dana Ben-Ari/);
});
test("rules.txt: CDF_ALGORITHM and settings parse; version matches the code", () => {
  const rules = E.parseRules(rulesWithPeople());
  assert.deepStrictEqual(rules.errors, []);
  assert.strictEqual(rules.cdfAlgorithm.version, A.ALGORITHM_VERSION);
  assert.ok(E.parseRules('CDF_ALGORITHM "x"\nUPDATE "A"\nMATCH UPDATE.A = BASE.C\nCOPY UPDATE.B -> BASE.C').errors.some((e) => /either/.test(e.message)));
});

console.log(`\n${passed} tests passed` + (process.exitCode ? " — SOME FAILED" : ""));
