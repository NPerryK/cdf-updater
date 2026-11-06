// Run: node tests/engine.test.js
const assert = require("assert");
const fs = require("fs");
const path = require("path");
const E = require("../engine.js");

let passed = 0;
function test(name, fn) {
  try { fn(); passed++; console.log("  ✓ " + name); }
  catch (e) { console.error("  ✗ " + name + "\n    " + e.stack); process.exitCode = 1; }
}

const C = (col) => E.colToIndex(col);

// build a sparse row: {C:"CDF-1", T:"3 - CONV. ENG."}
function row(obj, width = 30) {
  const r = new Array(width).fill("");
  for (const [k, v] of Object.entries(obj)) r[C(k)] = v;
  return r;
}

const rules = E.parseRules(fs.readFileSync(path.join(__dirname, "..", "rules.txt"), "utf8"));
const algo = rules.algorithms[0];
const NOW = new Date(2026, 10, 5, 14, 30, 0); // 05/11/2026
const TODAY = E.dateToSerial(NOW, false);

console.log("parser");
test("rules.txt parses with no errors", () => {
  assert.deepStrictEqual(rules.errors, []);
});
test("imports", () => {
  assert.strictEqual(rules.imports.length, 2);
  assert.deepStrictEqual(rules.imports.map((i) => [i.slot, i.target, i.from]), [[1, "INTEGRATOR", null], [2, "SHEET2", null]]);
});
test("algorithm settings", () => {
  assert.strictEqual(rules.algorithms.length, 1);
  assert.strictEqual(algo.baseFirst, true);
  assert.strictEqual(algo.update, "INTEGRATOR");
  assert.strictEqual(algo.baseStartRow, 3);
  assert.strictEqual(algo.updateStartRow, 2);
  assert.strictEqual(algo.match.update.col, "A");
  assert.strictEqual(algo.match.base.col, "C");
  assert.strictEqual(algo.actions.length, 4);
  assert.strictEqual(algo.onChange.length, 1);
  assert.strictEqual(rules.log, "UPDATE LOG");
});
test("colToIndex / indexToCol", () => {
  assert.strictEqual(C("A"), 0); assert.strictEqual(C("T"), 19); assert.strictEqual(C("AA"), 26); assert.strictEqual(C("AB"), 27);
  for (let i = 0; i < 800; i++) assert.strictEqual(E.colToIndex(E.indexToCol(i)), i);
});
test("syntax errors reported with line numbers", () => {
  const r = E.parseRules('UPDATE "X"\nMATCH UPDATE.A = BASE.C\nCOPY UPDATE.M => BASE.T\nFOO 1\nIF BASE.T == "x" COPY UPDATE.A -> BASE.B\nCOPY UPDATE.A -> UPDATE.B');
  assert.deepStrictEqual(r.errors.map((e) => e.line), [3, 4, 5, 6, 1]);
});
test("no-space arrow and NOT CONTAINS / AND / SET text", () => {
  const r = E.parseRules('UPDATE "U"\nMATCH BASE.C = UPDATE.A\nIF UPDATE.M NOT CONTAINS "CANCEL" AND BASE.T <> EMPTY COPY UPDATE.M->BASE.T\nSET BASE.Z = "ok"');
  assert.deepStrictEqual(r.errors, []);
  assert.strictEqual(r.algorithms[0].actions[0].conds.length, 2);
  assert.strictEqual(r.algorithms[0].match.update.col, "A");
});

console.log("algorithm (CDF scenario)");
const header1 = row({ C: "CDF NUM", T: "STATUS (GCC)", R: "DATE", X: "NAME", AA: "Q", AB: "R" });
const header2 = row({});
const base = {
  values: [
    header1, header2,
    row({ C: "A330-1-001", T: "3 - CONV. ENG.", R: 46000 }),             // row 3 -> status change, not final
    row({ C: "A330-1-002", T: "3 - CONV. ENG.", R: 46000, X: "old" }),   // row 4 -> becomes FINAL APPROVAL
    row({ C: "A330-1-003", T: "9 - FINAL APPROVAL", R: 46000, X: "Dana", AA: "q", AB: "r" }), // row 5 -> no change
    row({ C: "A330-1-004", T: "RELEASED", R: 46000 }),                   // row 6 -> not in update
    row({ C: " a330-1-005 ", T: "2 - POC ENG.", R: 46000 }),             // row 7 -> key normalisation
    row({ C: "A330-1-002", T: "3 - CONV. ENG.", R: 46000 }),             // row 8 -> duplicate base key
  ],
  formats: null,
};
const update = {
  values: [
    row({ A: "CDF", M: "STATUS", N: "NAME", Q: "QQ", R: "RR" }),
    row({ A: "A330-1-001", M: "4 - STRESS", N: "Yossi", Q: "q1", R: "r1" }),
    row({ A: "A330-1-002", M: "9 - final approval ", N: "Avi Cohen", Q: 46100, R: "r2" }),
    row({ A: "A330-1-003", M: "9 - FINAL APPROVAL", N: "Dana", Q: "q", R: "r" }),
    row({ A: "A330-1-005", M: "3 - CONV. ENG.", N: "", Q: "", R: "" }),
    row({ A: "A330-1-999", M: "1 - INITIATOR" }),                        // not in base
    row({}),                                                              // empty row ignored
  ],
  formats: null,
};
update.formats = update.values.map((r) => r.map(() => "General"));
update.formats[2][C("Q")] = "dd/mm/yyyy";

const res = E.computeChanges(algo, base, update, { now: NOW });
const byCell = {};
res.changes.forEach((c) => (byCell[c.colLetter + c.row] = c));

test("status copied M -> T for a non-final row, and only T + R change", () => {
  assert.strictEqual(byCell.T3.newValue, "4 - STRESS");
  assert.strictEqual(byCell.T3.oldValue, "3 - CONV. ENG.");
  assert.ok(!byCell.X3 && !byCell.AA3 && !byCell.AB3);
  assert.strictEqual(byCell.R3.newValue, TODAY);
  assert.strictEqual(byCell.R3.numberFormat, "dd/mm/yyyy");
  assert.strictEqual(byCell.R3.onChange, true);
});
test("FINAL APPROVAL (case/space-insensitive) copies N->X, Q->AA, R->AB", () => {
  assert.strictEqual(byCell.T4.newValue, "9 - final approval ");
  assert.strictEqual(byCell.X4.newValue, "Avi Cohen");
  assert.strictEqual(byCell.AA4.newValue, 46100);
  assert.strictEqual(byCell.AA4.numberFormat, "dd/mm/yyyy", "date format travels with copied value");
  assert.strictEqual(byCell.AB4.newValue, "r2");
  assert.strictEqual(byCell.R4.newValue, TODAY);
});
test("duplicate base key: both rows updated", () => {
  assert.ok(byCell.T8 && byCell.X8 && byCell.R8);
});
test("identical row -> no change and no date", () => {
  assert.ok(!res.changes.some((c) => c.row === 5));
});
test("key normalisation (spaces/case) matches row 7", () => {
  assert.strictEqual(byCell.T7.newValue, "3 - CONV. ENG.");
});
test("unmatched update keys reported", () => {
  assert.deepStrictEqual(res.stats.unmatched.map((u) => u.key), ["A330-1-999"]);
});
test("stats", () => {
  assert.strictEqual(res.stats.updateRows, 5);
  assert.strictEqual(res.stats.matchedKeys, 4);
  assert.strictEqual(res.stats.matchedRows, 5);
  assert.strictEqual(res.stats.changedRows, 4); // rows 3,4,7,8
});
test("running again on the updated data changes nothing (idempotent)", () => {
  const after = { values: base.values.map((r) => r.slice()), formats: null };
  res.changes.forEach((c) => (after.values[c.row - 1][c.col] = c.newValue));
  const again = E.computeChanges(algo, after, update, { now: NOW });
  assert.strictEqual(again.changes.length, 0);
});
test("SKIP_EMPTY YES keeps existing values when update cell is blank", () => {
  const r = E.parseRules('UPDATE "U"\nMATCH UPDATE.A = BASE.C\nSKIP_EMPTY YES\nCOPY UPDATE.M -> BASE.T');
  const b = { values: [row({}), row({ C: "K", T: "keep" })] };
  const u = { values: [row({}), row({ A: "K", M: "" })] };
  assert.strictEqual(E.computeChanges(r.algorithms[0], b, u, { now: NOW }).changes.length, 0);
  const r2 = E.parseRules('UPDATE "U"\nMATCH UPDATE.A = BASE.C\nCOPY UPDATE.M -> BASE.T');
  assert.strictEqual(E.computeChanges(r2.algorithms[0], b, u, { now: NOW }).changes.length, 1);
});
test("numbers vs numeric text are treated as equal", () => {
  const r = E.parseRules('UPDATE "U"\nMATCH UPDATE.A = BASE.C\nCOPY UPDATE.M -> BASE.T');
  const b = { values: [row({}), row({ C: 1574, T: 5 })] };
  const u = { values: [row({}), row({ A: "1574", M: "5" })] };
  const out = E.computeChanges(r.algorithms[0], b, u, { now: NOW });
  assert.strictEqual(out.stats.matchedKeys, 1);
  assert.strictEqual(out.changes.length, 0);
});
test("multiple ALGORITHM blocks", () => {
  const r = E.parseRules('ALGORITHM "a"\nUPDATE "U1"\nMATCH UPDATE.A = BASE.C\nCOPY UPDATE.B -> BASE.D\nALGORITHM "b"\nBASE "Main"\nUPDATE "U2"\nMATCH UPDATE.A = BASE.C\nSET BASE.E = NOW');
  assert.deepStrictEqual(r.errors, []);
  assert.deepStrictEqual(r.algorithms.map((a) => [a.name, a.update, a.base]), [["a", "U1", null], ["b", "U2", "Main"]]);
});

console.log("new CDF row");
test("rules.txt NEWROW block parsed", () => {
  const nr = rules.newRow;
  assert.strictEqual(nr.key.col, "C");
  assert.strictEqual(nr.startRow, 3, "defaults to BASE_START_ROW");
  assert.strictEqual(nr.cells.length, 12);
  assert.deepStrictEqual(nr.cells.find((c) => c.col === "H").source, { kind: "ai", field: "DESCRIPTION", key: "description" });
  assert.deepStrictEqual(nr.cells.find((c) => c.col === "A").source, { kind: "link", which: "ATT", prefix: "" });
  assert.deepStrictEqual(E.parseRules('AI_MODEL "x"').errors, [], "old AI_MODEL line still accepted");
});
test("NEWROW errors", () => {
  const r = E.parseRules('NEWROW C = AI.FOO\nNEWROW D = SOMETHING\nNEWROW E = LINK X\nNEWROW F = "x"\nNEWROW F = "y"');
  assert.deepStrictEqual(r.errors.map((e) => e.line), [1, 2, 3, 5, 4]); // last: missing key (reported on first cell line)
});
test("CDF from file names", () => {
  const t = (n) => E.parseCdfFileName(n);
  assert.deepStrictEqual(t("A330-1-961.-.pdf"), { cdf: "A330-1-961", isAtt: false });
  assert.deepStrictEqual(t("A330-1-961.-ATT.pdf"), { cdf: "A330-1-961", isAtt: true });
  assert.deepStrictEqual(t("A330-1-682.A.pdf"), { cdf: "A330-1-682.A", isAtt: false });
  assert.deepStrictEqual(t("A330-1-682.A-ATT.pdf"), { cdf: "A330-1-682.A", isAtt: true });
  assert.deepStrictEqual(t("A330-1-682.AATT.pdf"), { cdf: "A330-1-682.A", isAtt: true });
  assert.strictEqual(t("scan.pdf").cdf, null);
  assert.strictEqual(E.normalizeCdf("a330-1-961.-"), "A330-1-961");
});
const ai = {
  cdf_number: "A330-1-961.-", priority: "ROUTINE", description: "HOLES MISSING IN FRAME 38. NO ENGINEERING DISPOSITION PROVIDED YET.",
  oem_category: "OEM STRUCTURE", drawing_number: "DWG-1", installation: "INSTL - TEST", code: "C1025", initiator: "Test Person", flags: [],
};
test("buildNewRowCells fills every mapped column", () => {
  const b = E.buildNewRowCells(rules.newRow, { ai, mainName: "A330-1-961.-.pdf", attName: "A330-1-961.-ATT.pdf", now: NOW });
  const by = Object.fromEntries(b.cells.map((c) => [c.col, c]));
  assert.strictEqual(b.cdf, "A330-1-961");
  assert.deepStrictEqual(b.warnings, []);
  assert.strictEqual(by.C.value, "A330-1-961");
  assert.strictEqual(by.F.value, TODAY); assert.strictEqual(by.F.display, "05/11/2026"); assert.strictEqual(by.F.format, "dd/mm/yyyy");
  assert.strictEqual(by.H.value, ai.description);
  assert.strictEqual(by.M.value, "-");
  assert.deepStrictEqual(by.A.link, { address: "A330-1-961.-ATT.pdf", textToDisplay: "A330-1-961.-ATT.pdf" });
  assert.deepStrictEqual(by.B.link, { address: "A330-1-961.-.pdf", textToDisplay: "A330-1-961.-.pdf" });
});
test("buildNewRowCells warns on mismatched files and handles missing ATT", () => {
  const b = E.buildNewRowCells(rules.newRow, { ai: { ...ai, cdf_number: "A330-1-962" }, mainName: "A330-1-961.-.pdf", attName: "A330-1-700.-ATT.pdf", now: NOW });
  assert.strictEqual(b.warnings.length, 2);
  const noAtt = E.buildNewRowCells(rules.newRow, { ai, mainName: "A330-1-961.-.pdf", attName: null, now: NOW });
  assert.strictEqual(noAtt.cells.find((c) => c.col === "A").link, null);
});
test("findNewRowTarget: existing row vs append after last", () => {
  const keys = ["CDF NUM", "", "A330-1-001", "A330-1-961", "", "A330-1-003", "", ""];
  assert.deepStrictEqual(E.findNewRowTarget(keys, 3, "A330-1-961"), { row: 4, existing: true });
  assert.deepStrictEqual(E.findNewRowTarget(keys, 3, "a330-1-961.-"), { row: 4, existing: true });
  assert.deepStrictEqual(E.findNewRowTarget(keys, 3, "A330-1-999"), { row: 7, existing: false });
  assert.deepStrictEqual(E.findNewRowTarget(["H", ""], 3, "X"), { row: 3, existing: false });
});
test("date parsing round-trip", () => {
  assert.strictEqual(E.parseDateString("05/11/2026"), TODAY);
  assert.strictEqual(E.serialToDateString(TODAY), "05/11/2026");
  assert.strictEqual(E.parseDateString("31/13/2026"), null);
});

console.log("paste from claude.ai");
test("parses a fenced JSON array with text around it", () => {
  const txt = 'Here are the rows:\n```json\n[{"main_file":"A330-1-961.-.pdf","att_file":"A330-1-961.-ATT.pdf","cdf_number":"A330-1-961.-","description":"X","flags":["check frame"]},\n {"main_file":"A330-1-682.A.pdf","att_file":"","cdf_number":"A330-1-682.A","description":"Y","flags":[]}]\n```\nLet me know.';
  const r = E.parseAiPaste(txt);
  assert.strictEqual(r.error, null);
  assert.strictEqual(r.items.length, 2);
  assert.strictEqual(r.items[0].att_file, "A330-1-961.-ATT.pdf");
  assert.deepStrictEqual(r.items[0].flags, ["check frame"]);
  assert.strictEqual(r.items[1].priority, "", "missing keys become empty strings");
});
test("single object, smart quotes, trailing comma, alias keys", () => {
  const r = E.parseAiPaste('{“cdf”: “A330-1-700”, “drawing”: “DWG-1”, “Flags”: “one”,}');
  assert.strictEqual(r.error, null);
  assert.strictEqual(r.items[0].cdf_number, "A330-1-700");
  assert.strictEqual(r.items[0].drawing_number, "DWG-1");
  assert.deepStrictEqual(r.items[0].flags, ["one"]);
});
test("clear errors for empty / non-JSON / broken / unrelated JSON", () => {
  assert.ok(E.parseAiPaste("").error);
  assert.ok(/לא נמצא JSON/.test(E.parseAiPaste("A330-1-961  ROUTINE  DESCRIPTION").error));
  assert.ok(/לא תקין/.test(E.parseAiPaste('[{"cdf_number": "A330-1-9}]').error));
  assert.ok(/לא נמצאו שורות/.test(E.parseAiPaste('[{"foo": 1}]').error));
});
test("pasted item -> row cells use the file names for CDF and links", () => {
  const it = E.parseAiPaste('[{"main_file":"A330-1-961.-.pdf","att_file":"A330-1-961.-ATT.pdf","cdf_number":"A330-1-961.-","priority":"ROUTINE"}]').items[0];
  const b = E.buildNewRowCells(rules.newRow, { ai: it, mainName: it.main_file, attName: it.att_file, now: NOW });
  assert.strictEqual(b.cdf, "A330-1-961");
  assert.strictEqual(b.cells.find((c) => c.col === "A").link.address, "A330-1-961.-ATT.pdf");
  assert.strictEqual(b.cells.find((c) => c.col === "G").value, "ROUTINE");
});

console.log("email + reports");
test("REPORT and EMAIL_GROUP parse", () => {
  assert.strictEqual(rules.reports.length, 2);
  assert.deepStrictEqual(rules.reports[0].columns.map((c) => c.col), ["C", "T", "X"]);
  const r = E.parseRules('EMAIL_GROUP "team" "a@x.com; b@y.org"\nEMAIL_GROUP "bad" "nope"\nREPORT "x" WHERE UPDATE.A = "1" COLUMNS C\nREPORT "y" SHEET "S" FROM_ROW 5 WHERE BASE.T CONTAINS "FINAL" AND BASE.X <> EMPTY COLUMNS C,T');
  assert.deepStrictEqual(r.errors.map((e) => e.line), [2, 3]);
  assert.deepStrictEqual(r.emailGroups[0].addresses, ["a@x.com", "b@y.org"]);
  assert.strictEqual(r.reports[0].fromRow, 5);
  assert.strictEqual(r.reports[0].conds.length, 2);
});
test("runReport filters rows", () => {
  const out = E.runReport(rules.reports[0], base.values);
  assert.deepStrictEqual(out.rows.map((r) => r.row), [5]);
  assert.deepStrictEqual(out.rows[0].cells, ["A330-1-003", "9 - FINAL APPROVAL", "Dana"]);
});
test("lastRunFromLog picks the latest run and today's new rows", () => {
  const N = E.LOG_NOTES;
  const log = [
    ["hdr"],
    ["04/11/2026 09:00", "a", "S", "3", "A330-1-001", "T", "1", "2", ""],
    ["05/11/2026 10:00", "a", "S", "3", "A330-1-001", "T", "2", "3", ""],
    ["05/11/2026 10:00", "a", "S", "3", "A330-1-001", "R", "x", "y", N.auto],
    ["05/11/2026 10:00", "a", "U", "7", "A330-1-999", "", "", "", N.notFound],
    ["05/11/2026 11:00", N.newRow, "S", "9", "A330-1-961", "C", "", "A330-1-961", N.newRow],
    ["04/11/2026 11:00", N.newRow, "S", "8", "A330-1-960", "C", "", "A330-1-960", N.newRow],
  ];
  const run = E.lastRunFromLog(log, "05/11/2026");
  assert.strictEqual(run.stamp, "05/11/2026 10:00");
  assert.deepStrictEqual(run.changes.map((c) => [c.key, c.col, c.newValue]), [["A330-1-001", "T", "3"]]);
  assert.deepStrictEqual(run.notFound.map((c) => c.key), ["A330-1-999"]);
  assert.deepStrictEqual(run.newRows.map((c) => c.key), ["A330-1-961"]);
});
test("buildEmail + mailto", () => {
  const m = E.buildEmail({ title: "CDF status update – 05/11/2026", intro: "1 change.", sections: [{ heading: "", headers: ["CDF", "Field"], rows: [["A330-1-001", "<STATUS> & co"]] }] });
  assert.ok(m.text.includes("A330-1-001 | <STATUS> & co"));
  assert.ok(m.html.includes("&lt;STATUS&gt; &amp; co") && !m.html.includes("<STATUS>"));
  const url = E.mailtoUrl(["a@x.com", "b@y.org"], ["c@z.com"], "Subj & more", "line1\r\nline2");
  assert.strictEqual(url, "mailto:a@x.com;b@y.org?cc=c%40z.com&subject=Subj%20%26%20more&body=line1%0D%0Aline2");
});

console.log(`\n${passed} tests passed` + (process.exitCode ? " — SOME FAILED" : ""));
