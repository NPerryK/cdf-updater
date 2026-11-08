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
// a NEWROW spec that also writes the PDF links (the real rules.txt leaves A and B to the prepared rows)
const linkRows = E.parseRules('NEWROW_KEY C\nNEWROW_START_ROW 2\nNEWROW A = LINK ATT ""\nNEWROW B = LINK MAIN ""\nNEWROW C = CDF\nNEWROW F = TODAY\n' +
  'NEWROW G = AI.PRIORITY\nNEWROW H = AI.DESCRIPTION\nNEWROW I = AI.OEM\nNEWROW J = AI.DRAWING\nNEWROW K = AI.INSTALLATION\nNEWROW L = AI.CODE\nNEWROW M = "-"\nNEWROW O = AI.INITIATOR').newRow;
const generic = E.parseRules(fs.readFileSync(path.join(__dirname, "generic-rules.txt"), "utf8"));
const algo = generic.algorithms[0];
const NOW = new Date(2026, 10, 5, 14, 30, 0); // 05/11/2026
const TODAY = E.dateToSerial(NOW, false);

console.log("parser");
test("rules.txt parses with no errors", () => {
  assert.deepStrictEqual(rules.errors, []);
});
test("imports", () => {
  assert.strictEqual(rules.imports.length, 2);
  assert.deepStrictEqual(rules.imports.map((i) => [i.slot, i.target, i.file, i.from]), [[1, "GCC", "CDF_List_Export", null], [2, "Mantra", "650", null]]);
});
test("algorithm settings", () => {
  assert.strictEqual(generic.algorithms.length, 1);
  assert.strictEqual(rules.algorithms.length, 0, "rules.txt uses CDF_ALGORITHM");
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

test("IMPORT options: FILE/FROM in any order, errors", () => {
  const r = E.parseRules('IMPORT 1 "A" FROM "Data" FILE "X_Y"\nIMPORT 2 "B" FILE ""\nIMPORT 3 "C" FOO "x"');
  assert.deepStrictEqual([r.imports[0].file, r.imports[0].from], ["X_Y", "Data"]);
  assert.deepStrictEqual(r.errors.map((e) => e.line), [2, 3]);
});
test("file names route to the right import slot", () => {
  const imp = rules.imports;
  assert.strictEqual(E.matchImportSlot(imp, "CDF_List_Export.xlsx"), 1);
  assert.strictEqual(E.matchImportSlot(imp, "cdf list export (3).xlsx"), 1);
  assert.strictEqual(E.matchImportSlot(imp, "650.xlsx"), 2);
  assert.strictEqual(E.matchImportSlot(imp, "Report_650_Nov.xlsx"), 2);
  assert.strictEqual(E.matchImportSlot(imp, "CDF_List_Export_2026-06-50.xlsx"), 1, "longest name wins over a stray 650");
  assert.strictEqual(E.matchImportSlot(imp, "something else.xlsx"), null);
  assert.ok(E.importFileMatches("CDF_List_Export", "CDF-List-Export.xlsm"));
  assert.ok(!E.importFileMatches("650", "INTEGRATOR EXCEL.xlsx"));
});
test("all parser messages are English", () => {
  const r = E.parseRules('FOO\nIMPORT 1 "A" BAR\nMATCH UPDATE.A = UPDATE.B\nCOPY UPDATE.A => BASE.B\nNEWROW C = AI.NOPE\nREPORT "x"\nIF BASE.A ?? "1" COPY UPDATE.A -> BASE.B');
  r.errors.forEach((e) => assert.ok(!/[\u0590-\u05FF]/.test(e.message), e.message));
  assert.ok(r.errors.length >= 7);
  assert.strictEqual(E.describeAction(algo.actions[1]), 'If BASE.T = "9 - FINAL APPROVAL" — Copy UPDATE.N → BASE.X');
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
  assert.strictEqual(nr.startRow, 2);
  assert.strictEqual(nr.cells.length, 12);
  assert.deepStrictEqual(nr.cells.find((c) => c.col === "H").source, { kind: "ai", field: "DESCRIPTION", key: "description" });
  assert.ok(!nr.cells.some((c) => c.col === "A" || c.col === "B"), "the prepared links in A and B are left alone");
  assert.deepStrictEqual(linkRows.cells.find((c) => c.col === "A").source, { kind: "link", which: "ATT", prefix: "" });
  assert.deepStrictEqual(E.parseRules('AI_MODEL "x"').errors, [], "old AI_MODEL line still accepted");
});
test("NEWROW errors", () => {
  const r = E.parseRules('NEWROW C = AI.FOO\nNEWROW D = SOMETHING\nNEWROW E = LINK X\nNEWROW F = "x"\nNEWROW F = "y"');
  assert.deepStrictEqual(r.errors.map((e) => e.line), [1, 2, 3, 5, 4]); // last: missing key (reported on first cell line)
});
test("CDF from file names", () => {
  const t = (n) => E.parseCdfFileName(n);
  assert.deepStrictEqual(t("A330-1-961.-.pdf"), { cdf: "A330-1-961", printed: "A330-1-961.-", isAtt: false });
  assert.deepStrictEqual(t("A330-1-961.-ATT.pdf"), { cdf: "A330-1-961", printed: "A330-1-961.-", isAtt: true });
  assert.deepStrictEqual(t("A330-1-682.A.pdf"), { cdf: "A330-1-682.A", printed: "A330-1-682.A", isAtt: false });
  assert.deepStrictEqual(t("A330-1-682.A-ATT.pdf"), { cdf: "A330-1-682.A", printed: "A330-1-682.A", isAtt: true });
  assert.deepStrictEqual(t("A330-1-682.AATT.pdf"), { cdf: "A330-1-682.A", printed: "A330-1-682.A", isAtt: true });
  assert.strictEqual(t("scan.pdf").cdf, null);
  assert.strictEqual(E.normalizeCdf("a330-1-961.-"), "A330-1-961");
});
const ai = {
  cdf_number: "A330-1-961.-", priority: "ROUTINE", description: "HOLES MISSING IN FRAME 38. NO ENGINEERING DISPOSITION PROVIDED YET.",
  oem_category: "OEM STRUCTURE", drawing_number: "DWG-1", installation: "INSTL - TEST", code: "C1025", initiator: "Test Person", flags: [],
};
test("buildNewRowCells fills every mapped column", () => {
  const b = E.buildNewRowCells(linkRows, { ai, mainName: "A330-1-961.-.pdf", attName: "A330-1-961.-ATT.pdf", now: NOW });
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
  const b = E.buildNewRowCells(linkRows, { ai: { ...ai, cdf_number: "A330-1-962" }, mainName: "A330-1-961.-.pdf", attName: "A330-1-700.-ATT.pdf", now: NOW });
  assert.strictEqual(b.warnings.length, 2);
  const noAtt = E.buildNewRowCells(linkRows, { ai, mainName: "A330-1-961.-.pdf", attName: null, now: NOW });
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
  assert.ok(/No JSON found/.test(E.parseAiPaste("A330-1-961  ROUTINE  DESCRIPTION").error));
  assert.ok(/not valid/.test(E.parseAiPaste('[{"cdf_number": "A330-1-9}]').error));
  assert.ok(/No CDF rows/.test(E.parseAiPaste('[{"foo": 1}]').error));
});
test("pasted item -> row cells use the file names for CDF and links", () => {
  const it = E.parseAiPaste('[{"main_file":"A330-1-961.-.pdf","att_file":"A330-1-961.-ATT.pdf","cdf_number":"A330-1-961.-","priority":"ROUTINE"}]').items[0];
  const b = E.buildNewRowCells(linkRows, { ai: it, mainName: it.main_file, attName: it.att_file, now: NOW });
  assert.strictEqual(b.cdf, "A330-1-961");
  assert.strictEqual(b.cells.find((c) => c.col === "A").link.address, "A330-1-961.-ATT.pdf");
  assert.strictEqual(b.cells.find((c) => c.col === "G").value, "ROUTINE");
});
test("several CDFs at once: each new one takes the next free row, repeats share a row", () => {
  const keys = ["CDF NUM", "A330-1-001", "A330-1-961", "", "A330-1-003", ""];
  const t = E.assignNewRowTargets(keys, 2, ["A330-1-970", "A330-1-961.-", "A330-1-971", "A330-1-970.-"]);
  assert.deepStrictEqual(t.map((x) => [x.row, x.existing]), [[6, false], [3, true], [7, false], [6, true]]);
  assert.strictEqual(t[3].dupOf, 0);
  assert.strictEqual(keys.length, 6, "the sheet's key list is not changed");
});
test("a tab-separated row instead of JSON is recognised (Claude used another format)", () => {
  const row = "A330-1-973\tOpen\t05 Nov 2026\tC1039\t43075600\t330-25-23-JC855-001-SM\t330-25-23-JC855\tA\t-\tINSTL - CEILING FR41 - FR53.3 [MCC]\tC41-C53.3, Y1047.8 / -1047.8\tShort edge margin";
  const r = E.parseAiPaste(row);
  assert.strictEqual(r.kind, "table");
  assert.ok(/table row/.test(r.error));
  assert.strictEqual(E.parseAiPaste("```\n" + row + "\n```").kind, "table", "also inside a code block");
  assert.strictEqual(E.parseAiPaste('[{"cdf_number":"A330-1-1","description":"A\tB\tC\tD"}]').kind, null, "tabs inside JSON are fine");
  assert.strictEqual(E.parseAiPaste("A330-1-961 ROUTINE TEXT").kind, "nojson");
});
// a sheet like the real CDF sheet: C = key, B = prepared CDF numbers (rows for future CDFs)
//  row: 1 hdr | 2 A-001 | 3 (solution) | 4 A-002 | 5 A-002.A | 6 (solution) | 7 A-003 | 8 prep 004 | 9 prep 005 | 10 prep 006
const KEYS  = ["LINK TO GCC", "A330-1-001.-", "", "A330-1-002.-", "A330-1-002.A", "", "A330-1-003.-", "", "", ""];
const MATCH = ["LINK TO GCC FILE", "A330-1-001.-", "", "A330-1-002.-", "", "", "A330-1-003.-", "A330-1-004.-", "A330-1-005.-", "A330-1-006.-"];
test("target: existing / prepared row / revision insert / end", () => {
  const f = (cdf) => E.findNewRowTarget(KEYS, 2, cdf, MATCH);
  assert.deepStrictEqual(f("A330-1-002.-"), { row: 4, existing: true });
  assert.deepStrictEqual(f("A330-1-002"), { row: 4, existing: true }, "with or without .-");
  assert.deepStrictEqual(f("A330-1-005.-"), { row: 9, existing: false, prepared: true }, "the row prepared for 005, not the next free one");
  assert.deepStrictEqual(f("A330-1-001.A"), { row: 4, existing: false, insert: true, after: "A330-1-001.-" }, "after 001 and its solution row");
  assert.deepStrictEqual(f("A330-1-002.B"), { row: 7, existing: false, insert: true, after: "A330-1-002.A" }, "after the last revision's block");
  assert.deepStrictEqual(f("A330-1-099.-"), { row: 11, existing: false }, "no prepared row -> after the last used row");
  assert.deepStrictEqual(E.findNewRowTarget(KEYS, 2, "A330-1-099.A", MATCH), { row: 11, existing: false }, "revision of an unknown CDF -> end");
});
test("batch: rows shift under an inserted row; writeRow follows the order of writing", () => {
  const t = E.planNewRowTargets({ keys: KEYS, match: MATCH }, 2, ["A330-1-005.-", "A330-1-001.A", "A330-1-004.-", "A330-1-001.B"]);
  // 005: prepared row 9, written first; two rows are inserted above it later -> ends in row 11
  // 001.A: inserted at row 4 (after 001 + its solution row)
  // 004: prepared row 8, which is row 9 by then (after the first insert); ends in row 10
  // 001.B: inserted after 001.A -> row 5
  assert.deepStrictEqual(t.map((x) => [x.writeRow, x.row, !!x.prepared, !!x.insert]),
    [[9, 11, true, false], [4, 4, false, true], [9, 10, true, false], [5, 5, false, true]]);
  assert.strictEqual(t[3].after, "A330-1-001.A", "001.B goes after 001.A");
});
test("CDF FULL / CDF REV and the new AI fields", () => {
  const r = E.parseRules('NEWROW_SHEET "CDF"\nNEWROW_KEY C\nNEWROW_MATCH B\nNEWROW_START_ROW 2\nNEWROW C = CDF FULL\nNEWROW D = CDF REV\nNEWROW M = AI.DWG_REV\nNEWROW N = AI.AUTHOR\nNEWROW O = AI.ADVISER\nNEWROW P = AI.INITIATOR');
  assert.deepStrictEqual(r.errors, []);
  assert.strictEqual(r.newRow.match.col, "B");
  const it = E.parseAiPaste('[{"cdf_number":"A330-1-973.-","drawing_rev":"A","initiator":"Name A\\n050-1","technical_adviser":"Name B"}]').items[0];
  assert.strictEqual(it.author, "Name A\n050-1", "initiator is read as author");
  const b = E.buildNewRowCells(r.newRow, { ai: it, now: NOW });
  const v = Object.fromEntries(b.cells.map((c) => [c.col, c.value]));
  assert.deepStrictEqual(v, { C: "A330-1-973.-", D: "-", M: "A", N: "Name A\n050-1", O: "Name B", P: "Name A\n050-1" });
  const rv = E.buildNewRowCells(r.newRow, { ai: { cdf_number: "a330-1-682.a" }, now: NOW });
  assert.strictEqual(rv.cells[0].value, "A330-1-682.A"); assert.strictEqual(rv.cells[1].value, "A");
  const fromFile = E.buildNewRowCells(r.newRow, { ai: {}, mainName: "A330-1-961.-.pdf", now: NOW });
  assert.strictEqual(fromFile.cells[0].value, "A330-1-961.-", "from the file name when the form number is missing");
});
test("rules.txt: the new row follows the CDF sheet header", () => {
  const r = E.parseRules(fs.readFileSync(path.join(__dirname, "..", "rules.txt"), "utf8"));
  assert.deepStrictEqual(r.errors, []);
  const m = Object.fromEntries(r.newRow.cells.map((c) => [c.col, c.source.kind === "ai" ? "AI." + c.source.field : c.source.kind]));
  assert.deepStrictEqual(m, { C: "cdf_full", D: "cdf_rev", F: "today", G: "AI.PRIORITY", H: "AI.DESCRIPTION", I: "AI.OEM", J: "AI.DRAWING",
    K: "AI.INSTALLATION", L: "AI.CODE", M: "AI.DWG_REV", N: "AI.AUTHOR", O: "AI.ADVISER" });
  assert.strictEqual(r.newRow.sheet.name, "CDF"); assert.strictEqual(r.newRow.match.col, "B");
});
test("the older project format (cdf, nrc, jc, dwgver, title, summary ...) is read too", () => {
  const r = E.parseAiPaste('[{"cdf":"A330-1-973","status":"Open","open":"-","nrc":"C1039","cs":"1","jc":"330-25-23-JC855-001-SM","dwg":"330-25-23-JC855","dwgver":"A","title":"INSTL - X","summary":"TEXT","initiator":"Name","notes":"check X"}]');
  const it = r.items[0];
  assert.deepStrictEqual([it.cdf_number, it.code, it.drawing_number, it.drawing_rev, it.installation, it.description, it.author, it.form_date],
    ["A330-1-973", "C1039", "330-25-23-JC855-001-SM", "A", "INSTL - X", "TEXT", "Name", ""]);
  assert.deepStrictEqual(it.flags, ["check X", "Not used (not in the add-in's format): status, cs"]);
});
test("prompt_version is kept from the reply", () => {
  const r = E.parseAiPaste('[{"cdf_number":"A330-1-1","prompt_version":"ABC123"},{"cdf_number":"A330-1-2"}]');
  assert.strictEqual(r.items[0].prompt_version, "ABC123");
  assert.strictEqual(r.items[1].prompt_version, "");
});
test("CLAUDE_PROJECT parses and is checked", () => {
  assert.strictEqual(E.parseRules('CLAUDE_PROJECT "https://claude.ai/project/0193-abc"').claudeProject, "https://claude.ai/project/0193-abc");
  assert.ok(/project's address/.test(E.parseRules('CLAUDE_PROJECT "https://claude.ai/new"').errors[0].message));
  assert.strictEqual(E.parseRules(fs.readFileSync(path.join(__dirname, "..", "rules.txt"), "utf8")).claudeProject, null, "commented out by default");
});
test("prompt: version code is stable, changes with the text, and is asked for in the reply", () => {
  const P = require("../prompt.js").CdfPrompt || globalThis.CdfPrompt;
  const base = fs.readFileSync(path.join(__dirname, "..", "ai-prompt.txt"), "utf8");
  const a = P.build(base), b = P.build(base.replace(/\n/g, "\r\n")), c = P.build(base + "\nOne more rule.");
  assert.match(a.version, /^[0-9A-F]{6}$/);
  assert.strictEqual(a.version, b.version, "line endings don't matter");
  assert.notStrictEqual(a.version, c.version);
  assert.ok(a.text.includes('prompt_version: always "' + a.version + '"') && a.text.includes('"prompt_version": "' + a.version + '"'));
  assert.ok(a.text.includes("UPPERCASE English") && a.text.includes("OUTPUT FORMAT"));
  const fix = P.fixupMessage(base);
  assert.ok(fix.startsWith("Your previous answer is not in the format") && fix.includes(a.text), "fix-up message carries the full instructions");
});

console.log("email + reports");
test("EMAIL DRAFT tables are read with their merged cells (rowspan)", () => {
  const v = [["A330-300 P2F — Daily CDF Status — 05/11/2026"], [], ["בוקר טוב ... AS OF 05/11/2026 08:10 (Asia/Jerusalem)."], [],
    ["PERSONS – GROUP – SIGNATURE – CDF – PENDING MANTRA DOCUMENTS"],
    ["PERSON", "# of M", "MAIN GROUP", "SIGNATURE POSITION", "CDF NUMBERS", "PENDING MANTRA DOCUMENTS"],
    ["P1", 1, "IAI", "Rejected", "A330-1-1.-", "DOC-1 Rev -"],
    ["P2", 2, "IAI", "Check Design", "A330-1-2.-", "DOC-2 Rev A\r\nDOC-3 Rev -"],
    ["", "", "", "Check Design", "A330-1-3.-", "DOC-4 Rev B"],
    [], [],
    ["CDFS RELEASED - LAST 24 HOURS (AS OF 05/11/2026 08:10 Asia/Jerusalem)"], ["CDF NUMBER", "PREVIOUS STATUS", "NEW STATUS"],
    ["A330-1-9.-", "Open", "Released"], ["TOTAL", "", 1], ["בברכה,"], [], ["Name"]];
  const t = E.readSheetTables(v, ["PERSONS", "CDFS RELEASED — LAST 24", "CDFS RELEASED — LAST 7"]);
  assert.strictEqual(t.length, 2, "dash kinds don't matter; a missing table is skipped");
  assert.deepStrictEqual(t[0].rows.map((r) => r.map((c) => c && (c.v.slice(0, 5) + (c.rowspan > 1 ? "^" + c.rowspan : "")))),
    [["P1", "1", "IAI", "Rejec", "A330-", "DOC-1"], ["P2^2", "2^2", "IAI^2", "Check", "A330-", "DOC-2"], [null, null, null, "Check", "A330-", "DOC-4"]]);
  assert.strictEqual(t[0].rows[1][5].v, "DOC-2 Rev A\nDOC-3 Rev -");
  assert.deepStrictEqual(t[1].rows.map((r) => r.map((c) => c.v)), [["A330-1-9.-", "Open", "Released"], ["TOTAL", "", "1"]], "TOTAL ends the table, its empty cell is not merged");
  assert.strictEqual(t[1].endRow, 14);
  const m = E.buildDailyEmail({ subject: "S", opening: "בוקר טוב", signature: ["בברכה,", "Name"],
    tables: [{ title: "CDF STATUS", headers: ["CDF STATUS", "CDF'S"], rows: [[{ v: "Open", rowspan: 1 }, { v: "3", rowspan: 1 }]] }].concat(t) });
  assert.ok(m.html.includes('colspan="2"') && m.html.includes('colspan="6"') && m.html.includes('colspan="3"'), "title rows span the table");
  assert.ok(m.html.includes('rowspan="2"') && m.html.includes("DOC-2 Rev A<br>DOC-3 Rev -"));
  assert.ok(m.html.includes("background:#B10001;color:#FFFFFF;font-weight:bold") && m.html.includes("border:1px solid #BCBCBC"), "red headers, grey borders");
  assert.ok(m.html.includes('dir="rtl"') && m.html.includes('dir="ltr"'), "Hebrew text RTL, tables LTR");
  assert.strictEqual((m.html.match(/<br><br>/g) || []).length, 2, "two empty lines between tables");
  assert.ok(m.text.startsWith("בוקר טוב") && m.text.endsWith("Name"));
});
test("the add-in version is the same everywhere (badge, file versions, manifest)", () => {
  const html = fs.readFileSync(path.join(__dirname, "..", "taskpane.html"), "utf8");
  const badge = /id="appVersion"[^>]*>v([\d.]+)</.exec(html)[1];
  const vs = [...html.matchAll(/\?v=([\d.]+)"/g)].map((m) => m[1]);
  assert.ok(vs.length >= 6 && vs.every((x) => x === badge), "every ?v= equals the badge " + badge + ": " + vs);
  const mf = /<Version>([\d.]+)<\/Version>/.exec(fs.readFileSync(path.join(__dirname, "..", "manifest.xml"), "utf8"))[1];
  assert.strictEqual(mf, badge + ".0");
});
test("REPORT and EMAIL_GROUP parse", () => {
  assert.strictEqual(rules.reports.length, 2);
  assert.deepStrictEqual(generic.reports[0].columns.map((c) => c.col), ["C", "T", "X"]);
  assert.deepStrictEqual(rules.reports[0].columns.map((c) => c.col), ["C", "S", "T"]);
  assert.strictEqual(rules.reports[0].conds.length, 3);
  const r = E.parseRules('EMAIL_GROUP "team" "a@x.com; b@y.org"\nEMAIL_GROUP "bad" "nope"\nREPORT "x" WHERE UPDATE.A = "1" COLUMNS C\nREPORT "y" SHEET "S" FROM_ROW 5 WHERE BASE.T CONTAINS "FINAL" AND BASE.X <> EMPTY COLUMNS C,T');
  assert.deepStrictEqual(r.errors.map((e) => e.line), [2, 3]);
  assert.deepStrictEqual(r.emailGroups[0].addresses, ["a@x.com", "b@y.org"]);
  assert.strictEqual(r.reports[0].fromRow, 5);
  assert.strictEqual(r.reports[0].conds.length, 2);
});
test("runReport filters rows", () => {
  const out = E.runReport(generic.reports[0], base.values);
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
test("log written by the old Hebrew version is still read", () => {
  const log = [["hdr"],
    ["05/11/2026 10:00", "a", "U", "7", "A330-1-999", "", "", "", "\u05DC\u05D0 \u05E0\u05DE\u05E6\u05D0 \u05D1\u05D2\u05D9\u05DC\u05D9\u05D5\u05DF \u05D4\u05D1\u05E1\u05D9\u05E1"],
    ["05/11/2026 11:00", "x", "S", "9", "A330-1-961", "C", "", "A330-1-961", "\u05E9\u05D5\u05E8\u05D4 \u05D7\u05D3\u05E9\u05D4"]];
  const run = E.lastRunFromLog(log, "05/11/2026");
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
