// Run: node tests/core.test.js — cdf-core.js, cdf-rows.js and cdf-notes.js (5+ cases each function)
const assert = require("assert");
const K = require("../cdf-core.js"), R = require("../cdf-rows.js"), N = require("../cdf-notes.js");
let passed = 0, failed = 0; const perFn = {};
function test(fn, name, f) { perFn[fn] = (perFn[fn] || 0) + 1; try { f(); passed++; console.log("  ✓ " + fn + ": " + name); } catch (e) { failed++; console.error("  ✗ " + fn + ": " + name + "\n    " + e.stack); process.exitCode = 1; } }
const C = K.col;
function row(o, w) { const r = new Array(w || 34).fill(""); for (const [k, v] of Object.entries(o)) r[C(k)] = v; return r; }
const HEAD = new Array(34).fill("").map((_, i) => "H" + K.letter(i)); HEAD[C("C")] = "LINK TO GCC"; HEAD[C("F")] = "RECEIVED DATE"; HEAD[C("H")] = "PROBLEM DESCRIPTION";
const SHEET = [HEAD,
  row({ A: "A330-1-901-ATT.", C: "A330-1-901.-", R: "Open", S: "9 - Final Approval", AA: "D1" }),
  row({ R: "Open", AA: "D2" }),
  row({ C: "A330-1-902.-", R: "Open", S: "3 - Conversion Engineer" }),
  row({ C: "A330-1-902.A", R: "Open" }),
  row({ A: "A330-1-905.-ATT.", B: "A330-1-905.-" }),
  row({ A: "A330-1-906.-ATT.", B: "A330-1-906.-" })];

// ---------- core ----------
test("col/letter", "A, Z, AA, AH", () => { assert.deepStrictEqual(["A", "Z", "AA", "AH"].map(C), [0, 25, 26, 33]); assert.deepStrictEqual([0, 25, 26, 33].map(K.letter), ["A", "Z", "AA", "AH"]); });
test("col/letter", "lower case", () => assert.strictEqual(C("ah"), 33));
test("col/letter", "round trip to ZZ", () => { for (let i = 0; i < 702; i++) assert.strictEqual(C(K.letter(i)), i); });
test("col/letter", "AAA", () => assert.strictEqual(K.letter(702), "AAA"));
test("col/letter", "B", () => assert.strictEqual(C("B"), 1));
test("daySerial", "a serial", () => assert.strictEqual(K.daySerial(46335.7), 46335));
test("daySerial", "ISO", () => assert.strictEqual(K.daySerial("2026-11-09"), 46335));
test("daySerial", "dd/mm/yyyy (Israel)", () => assert.strictEqual(K.daySerial("08/11/2026"), 46334));
test("daySerial", "not a date", () => { assert.strictEqual(K.daySerial("31/02/2026"), null); assert.strictEqual(K.daySerial("soon"), null); assert.strictEqual(K.daySerial(""), null); });
test("daySerial", "zero and negatives aren't dates", () => { assert.strictEqual(K.daySerial(0), null); assert.strictEqual(K.daySerial(-3), null); });
test("weekStart", "a Monday → the Sunday before", () => assert.strictEqual(K.weekStart(46335), 46334));
test("weekStart", "a Sunday is its own start", () => assert.strictEqual(K.weekStart(46334), 46334));
test("weekStart", "a Saturday", () => assert.strictEqual(K.weekStart(46340), 46334));
test("weekStart", "next Sunday", () => assert.strictEqual(K.weekStart(46341), 46341));
test("weekStart", "always a Sunday", () => { for (let d = 46000; d < 46100; d++) assert.strictEqual(K.dateOf(K.weekStart(d)).getDay(), 0); });
test("keyParts", "revision letter", () => assert.deepStrictEqual(K.keyParts("A330-1-723.A"), { num: 723, rev: "A", label: "A330-1-723", full: "A330-1-723.A" }));
test("keyParts", "no suffix = revision -", () => assert.strictEqual(K.keyParts("A330-1-978").full, "A330-1-978.-"));
test("keyParts", "the ATT link text", () => assert.strictEqual(K.keyParts("A330-1-901-ATT.").num, 901));
test("keyParts", "padding to 3 digits, 4 digits stay", () => { assert.strictEqual(K.keyParts("A330-1-7.-").label, "A330-1-007"); assert.strictEqual(K.keyParts("A330-1-1001.-").label, "A330-1-1001"); });
test("keyParts", "not a CDF", () => { assert.strictEqual(K.keyParts("hello"), null); assert.strictEqual(K.keyParts(""), null); });
test("sameCdf", "with and without .-", () => assert.ok(K.sameCdf("A330-1-978", "A330-1-978.-")));
test("sameCdf", "different revisions", () => assert.ok(!K.sameCdf("A330-1-723.A", "A330-1-723.-")));
test("sameCdf", "spaces and case", () => assert.ok(K.sameCdf(" a330-1-723.a ", "A330-1-723.A")));
test("sameCdf", "different numbers", () => assert.ok(!K.sameCdf("A330-1-723.-", "A330-1-724.-")));
test("sameCdf", "not CDF keys: compared as they are", () => { assert.ok(K.sameCdf("A330-1-052 (another)", "a330-1-052 (another)")); assert.ok(!K.sameCdf("x", "y")); });
test("parseQuery", "a number", () => assert.deepStrictEqual(K.parseQuery("856"), { num: 856, rev: null }));
test("parseQuery", "number.rev", () => assert.deepStrictEqual(K.parseQuery("740.a"), { num: 740, rev: "A" }));
test("parseQuery", "a full key", () => assert.deepStrictEqual(K.parseQuery("A330-1-856.A"), { num: 856, rev: "A" }));
test("parseQuery", "leading zeros", () => assert.strictEqual(K.parseQuery("007").num, 7));
test("parseQuery", "words", () => assert.strictEqual(K.parseQuery("open"), null));
test("buildBlocks", "anchors and their solution rows", () => assert.deepStrictEqual(K.buildBlocks(SHEET).map((b) => [b.key, b.rows]), [["A330-1-901.-", [1, 2]], ["A330-1-902.-", [3]], ["A330-1-902.A", [4]]]));
test("buildBlocks", "template rows don't join a block", () => assert.ok(!K.buildBlocks(SHEET).some((b) => b.rows.includes(5))));
test("buildBlocks", "empty / header only", () => { assert.deepStrictEqual(K.buildBlocks([]), []); assert.deepStrictEqual(K.buildBlocks([HEAD]), []); });
test("buildBlocks", "rows before the first anchor", () => assert.deepStrictEqual(K.buildBlocks([HEAD, row({ R: "Open" }), row({ C: "K" })]).map((b) => b.rows), [[2]]));
test("buildBlocks", "null input", () => assert.deepStrictEqual(K.buildBlocks(null), []));
test("parseDelimited", "TSV", () => assert.deepStrictEqual(K.parseDelimited("a\tb\n1\t2"), [["a", "b"], ["1", "2"]]));
test("parseDelimited", "Excel's quoted cell with a line break", () => assert.deepStrictEqual(K.parseDelimited('x\t"line 1\nline ""2"""\ty', "\t"), [["x", 'line 1\nline "2"', "y"]]));
test("parseDelimited", "a quote inside a cell is a character (12\" PANEL)", () => assert.deepStrictEqual(K.parseDelimited('12" PANEL\tB', "\t"), [['12" PANEL', "B"]]));
test("parseDelimited", "CSV with ; and CRLF", () => assert.deepStrictEqual(K.parseDelimited("a;b\r\n1;2\r\n"), [["a", "b"], ["1", "2"]]));
test("parseDelimited", "empty lines go, BOM goes", () => assert.deepStrictEqual(K.parseDelimited("﻿a\tb\n\n\t\n1\t2"), [["a", "b"], ["1", "2"]]));
test("looseJson", "plain", () => assert.deepStrictEqual(K.looseJson('{"a":1}').data, { a: 1 }));
test("looseJson", "smart quotes and a trailing comma", () => assert.deepStrictEqual(K.looseJson("[“a”, “b”,]").data, ["a", "b"]));
test("looseJson", "Python style", () => assert.deepStrictEqual(K.looseJson("[['a', None, True]]").data, [["a", null, true]]));
test("looseJson", "no JSON", () => assert.strictEqual(K.looseJson("hello").kind, "nojson"));
test("looseJson", "cut off", () => assert.strictEqual(K.looseJson('{"a": [1, 2').kind, "cut"));
test("parseRules", "the shipped rules.txt has no errors", () => { const r = K.parseRules(require("fs").readFileSync(__dirname + "/../rules.txt", "utf8")); assert.deepStrictEqual(r.errors, []); assert.deepStrictEqual(r.imports.map((i) => i.target), ["CDF", "GCC", "MANTRAREV"]); });
test("parseRules", "defaults when empty", () => { const r = K.parseRules(""); assert.strictEqual(r.imports.length, 3); assert.deepStrictEqual(r.exclude, ["A330-1-052 (another)"]); assert.strictEqual(r.log, "UPDATE LOG"); });
test("parseRules", "an old 1.x line is reported, not obeyed", () => assert.ok(/isn't a setting this version knows/.test(K.parseRules('FINAL_APPROVAL_RELEASE "X"').errors[0])));
test("parseRules", "IMPORT without HEADERS", () => assert.ok(/needs HEADERS/.test(K.parseRules('IMPORT "GCC" FILE "x"').errors[0])));
test("parseRules", "TARGET, STUCK_DAYS, ROWS_COLUMNS, EXCLUDE", () => {
  const r = K.parseRules('TARGET OPEN 150 BY 2026-12-31\nSTUCK_DAYS 10\nROWS_COLUMNS B:Z\nEXCLUDE "A330-1-1.-"');
  assert.deepStrictEqual(r.targets.OPEN, { value: 150, by: K.daySerial("2026-12-31") }); assert.strictEqual(r.stuckDays, 10); assert.strictEqual(r.rows.firstCol + r.rows.lastCol, "BZ"); assert.deepStrictEqual(r.exclude, ["A330-1-1.-"]);
});
const IMPS = K.parseRules("").imports;
test("importFor", "GCC by its headers", () => assert.strictEqual(K.importFor([["CDF Number", "Site", "Current Station"]], IMPS).target, "GCC"));
test("importFor", "MANTRAREV, header in the 2nd row", () => assert.strictEqual(K.importFor([["report"], ["Project", "Doc ID", "Rev", "Release Date"]], IMPS).target, "MANTRAREV"));
test("importFor", "CDF", () => assert.strictEqual(K.importFor([["CDF NUM (LINK TO DEV)", "x", "CDF STATUS (GCC)"]], IMPS).target, "CDF"));
test("importFor", "case and spaces", () => assert.strictEqual(K.importFor([["cdf  number", "CURRENT STATION "]], IMPS).target, "GCC"));
test("importFor", "the 650 report (Mantra status) is not MANTRAREV", () => assert.strictEqual(K.importFor([["Project", "Doc ID", "Rev", "Status", "Date Status"]], IMPS), null));
test("fileNameFor", "part of the name", () => assert.strictEqual(K.fileNameFor("CDF_List_Export (3).xlsx", IMPS).target, "GCC"));
test("fileNameFor", "case", () => assert.strictEqual(K.fileNameFor("cdf_list_export.xlsx", IMPS).target, "GCC"));
test("fileNameFor", "no FILE → no match", () => assert.strictEqual(K.fileNameFor("CDF.xlsx", IMPS), null));
test("fileNameFor", "MANTRAREV by name", () => assert.strictEqual(K.fileNameFor("mantrarev 0911.xlsx", K.parseRules(require("fs").readFileSync(__dirname + "/../rules.txt", "utf8")).imports).target, "MANTRAREV"));
test("fileNameFor", "nothing", () => assert.strictEqual(K.fileNameFor("", IMPS), null));

// ---------- rows ----------
const AI = ["A330-1-902.B", "", "", "08/11/2026", "ROUTINE", "DESC", "TAGS", "DWG-1", "DWG NAME", "C1", "-", "", "Ron Avi"];
test("parse", "TSV from Excel", () => { const p = R.parse(AI.join("\t")); assert.strictEqual(p.kind, "tsv"); assert.strictEqual(p.rows[0].length, 13); });
test("parse", "a JSON array of arrays (ChatGPT)", () => { const p = R.parse("json\nCopy code\n" + JSON.stringify([AI])); assert.strictEqual(p.kind, "json"); assert.strictEqual(p.rows[0][0], "A330-1-902.B"); });
test("parse", "JSON objects → a header row", () => { const p = R.parse('[{"cdf_number": "A330-1-1.-", "priority": "ROUTINE"}]'); assert.deepStrictEqual(p.rows, [["cdf_number", "priority"], ["A330-1-1.-", "ROUTINE"]]); });
test("parse", "a ``` fence", () => assert.strictEqual(R.parse("```\n" + AI.join("\t") + "\n```").rows.length, 1));
test("parse", "empty / too many rows", () => { assert.strictEqual(R.parse(" ").kind, "empty"); assert.strictEqual(R.parse(new Array(205).fill("a\tb").join("\n")).kind, "big"); });
test("layoutOf", "13 AI fields", () => assert.strictEqual(R.layoutOf([AI], HEAD).mode, "ai13"));
test("layoutOf", "a header row by the sheet's names", () => { const l = R.layoutOf([["LINK TO GCC", "PROBLEM DESCRIPTION"], ["A330-1-1.-", "x"]], HEAD); assert.strictEqual(l.mode, "header"); assert.deepStrictEqual(l.map, [C("C"), C("H")]); });
test("layoutOf", "a whole sheet row (starts with the ATT link)", () => { const l = R.layoutOf([row({ A: "A330-1-950-ATT.", C: "A330-1-950.-" })], HEAD); assert.strictEqual(l.map[0], 0); });
test("layoutOf", "the CDF number lands in C", () => { const l = R.layoutOf([["A330-1-950.-", "x", "y"]], HEAD); assert.strictEqual(l.map[0], C("C")); });
test("layoutOf", "a chosen start", () => assert.strictEqual(R.layoutOf([["a", "b"]], HEAD, { mode: "from", start: C("H") }).map[1], C("I")));
test("cellsOf", "13 fields → C, D (revision), F as a date …", () => { const c = R.cellsOf(AI, R.layoutOf([AI], HEAD)).cells; const at = (L) => c.find((x) => x.col === C(L)).value; assert.strictEqual(at("C"), "A330-1-902.B"); assert.strictEqual(at("D"), "B"); assert.strictEqual(at("F"), 46334); assert.strictEqual(at("O"), "Ron Avi"); });
test("cellsOf", "the key is written in full", () => assert.strictEqual(R.cellsOf(["A330-1-978"], { mode: "from", map: [C("C")] }).cells[0].value, "A330-1-978.-"));
test("cellsOf", "values outside the columns are dropped and counted", () => { const r = R.cellsOf(["a", "b"], { mode: "from", map: [C("AH"), C("AH") + 1] }); assert.strictEqual(r.cells.length, 1); assert.strictEqual(r.dropped, 1); });
test("cellsOf", "a text date in F that isn't a date stays text", () => assert.strictEqual(R.cellsOf(["soon"], { mode: "from", map: [C("F")] }).cells[0].value, "soon"));
test("cellsOf", "unmapped columns are dropped", () => assert.strictEqual(R.cellsOf(["x", "y"], { mode: "header", map: [-1, C("H")] }).dropped, 1));
const cellsFor = (key) => [{ col: C("C"), value: key }];
test("suggest", "a new revision goes under the CDF's last revision", () => { const s = R.suggest(cellsFor("A330-1-902.B"), SHEET); assert.deepStrictEqual([s.mode, s.row], ["insert-below", 5]); });
test("suggest", "the same CDF → write over it", () => { const s = R.suggest(cellsFor("A330-1-902"), SHEET); assert.deepStrictEqual([s.mode, s.row], ["overwrite", 4]); });
test("suggest", "a new number with a template row → that row", () => { const s = R.suggest(cellsFor("A330-1-906.-"), SHEET); assert.deepStrictEqual([s.mode, s.row], ["overwrite", 7]); });
test("suggest", "a new number without a template row → after the CDF before it", () => { const s = R.suggest(cellsFor("A330-1-903.-"), SHEET); assert.deepStrictEqual([s.mode, s.row], ["insert-below", 5]); });
test("suggest", "no CDF number → after the last CDF", () => { const s = R.suggest([{ col: C("H"), value: "x" }], SHEET); assert.deepStrictEqual([s.mode, s.row], ["insert-below", 5]); });
test("plan", "write-overs first, then inserts from the bottom up", () => { const p = R.plan([{ cells: [], mode: "insert-below", row: 2 }, { cells: [], mode: "overwrite", row: 6 }, { cells: [], mode: "insert-below", row: 4 }], SHEET); assert.deepStrictEqual(p.ops.map((o) => o.kind + (o.row || o.at)), ["overwrite6", "insert5", "insert3"]); });
test("plan", "final rows account for the inserts above", () => assert.deepStrictEqual(R.plan([{ cells: [], mode: "insert-below", row: 2 }, { cells: [], mode: "overwrite", row: 6 }, { cells: [], mode: "insert-below", row: 4 }], SHEET).rows, [3, 8, 6]));
test("plan", "two inserts at the same place keep the pasted order", () => assert.deepStrictEqual(R.plan([{ cells: [], mode: "insert-below", row: 4 }, { cells: [], mode: "insert-below", row: 4 }], SHEET).rows, [5, 6]));
test("plan", "insert above", () => assert.deepStrictEqual(R.plan([{ cells: [], mode: "insert-above", row: 4 }], SHEET).rows, [4]));
test("plan", "two rows over the same row / a row above the start", () => { assert.ok(/both write over row 6/.test(R.plan([{ cells: [], mode: "overwrite", row: 6 }, { cells: [], mode: "overwrite", row: 6 }], SHEET).errors[0])); assert.ok(/from 2 down/.test(R.plan([{ cells: [], mode: "overwrite", row: 1 }], SHEET).errors[0])); });
test("rowSays", "a CDF", () => assert.strictEqual(R.rowSays(SHEET, 2), "A330-1-901.- · Open"));
test("rowSays", "a template row", () => assert.strictEqual(R.rowSays(SHEET, 6), "empty template row of A330-1-905"));
test("rowSays", "a solution row", () => assert.strictEqual(R.rowSays(SHEET, 3), "a solution row"));
test("rowSays", "past the end", () => assert.strictEqual(R.rowSays(SHEET, 50), "an empty row"));
test("rowSays", "an empty row", () => assert.strictEqual(R.rowSays([HEAD, row({})], 2), "an empty row"));

// ---------- notes ----------
const NOTES = [{ refs: ["A330-1-901.-"], text: "Ask for the drawing", state: "Open", due: "2026-11-12", subject: "RE: 901", updated: 2 }, { refs: ["A330-1-902.-"], text: "Done here", state: "Done", subject: "", updated: 1 }];
test("noteLines", "state, due, text, subject", () => assert.deepStrictEqual(N.noteLines([NOTES[0]]), ["✉ due 12/11: Ask for the drawing (RE: 901)"]));
test("noteLines", "Done shows", () => assert.deepStrictEqual(N.noteLines([NOTES[1]]), ["✉ Done: Done here"]));
test("noteLines", "line breaks fold", () => assert.deepStrictEqual(N.noteLines([{ text: "a\n\nb" }]), ["✉ a b"]));
test("noteLines", "empty notes go", () => assert.deepStrictEqual(N.noteLines([{ text: " " }, null]), []));
test("noteLines", "12 at most", () => assert.strictEqual(N.noteLines(new Array(20).fill({ text: "x" })).length, 12));
test("mergeE", "own text stays, notes after it", () => assert.strictEqual(N.mergeE("mine", ["✉ a"]), "mine\n✉ a"));
test("mergeE", "old ✉ lines are replaced", () => assert.strictEqual(N.mergeE("mine\n✉ old", ["✉ new"]), "mine\n✉ new"));
test("mergeE", "no notes → only own text", () => assert.strictEqual(N.mergeE("mine\n✉ old", []), "mine"));
test("mergeE", "empty cell", () => assert.strictEqual(N.mergeE("", ["✉ a"]), "✉ a"));
test("mergeE", "a ✉ in the middle of own text isn't touched", () => assert.strictEqual(N.mergeE("write to ✉ box", []), "write to ✉ box"));
const NS = [HEAD, row({ C: "A330-1-901.-", E: "mine", R: "Open" }), row({ C: "A330-1-902.-", R: "Open" }), row({ C: "A330-1-903.-", E: "✉ old", R: "Open" })];
test("notesPlan", "only cells that change", () => assert.deepStrictEqual(N.notesPlan(NS, NOTES).map((p) => p.row), [2, 3, 4]));
test("notesPlan", "the new value", () => assert.strictEqual(N.notesPlan(NS, NOTES)[0].value, "mine\n✉ due 12/11: Ask for the drawing (RE: 901)"));
test("notesPlan", "a CDF whose note went away loses its ✉ line", () => assert.strictEqual(N.notesPlan(NS, NOTES)[2].value, ""));
test("notesPlan", "applied → nothing more to do", () => { const s2 = NS.map((r) => r.slice()); N.notesPlan(NS, NOTES).forEach((p) => { s2[p.row - 1][C("E")] = p.value; }); assert.deepStrictEqual(N.notesPlan(s2, NOTES), []); });
test("notesPlan", "no notes, no ✉ lines → nothing", () => assert.deepStrictEqual(N.notesPlan([HEAD, row({ C: "A330-1-1.-", E: "x" })], []), []));
const TR = [HEAD, row({ C: "A330-1-001.-", R: "Open", S: "3 - Conversion Engineer", F: 46300 }), row({ C: "A330-1-002.-", R: "Open", S: "9 - Final Approval" }), row({ C: "A330-1-003.-", R: "Released", S: "Released" })];
test("trackStations", "first time: since the received date, not exact", () => assert.deepStrictEqual(N.trackStations(null, TR, 46335).map["A330-1-001.-"], ["3 - Conversion Engineer", "2026-10-05", 0, "2026-11-09"]));
test("trackStations", "first time without a date: today, exact", () => assert.deepStrictEqual(N.trackStations(null, TR, 46335).map["A330-1-002.-"], ["9 - Final Approval", "2026-11-09", 1, "2026-11-09"]));
test("trackStations", "only open CDFs", () => assert.ok(!("A330-1-003.-" in N.trackStations(null, TR, 46335).map)));
test("trackStations", "same station → unchanged", () => { const a = N.trackStations(null, TR, 46335); const b = N.trackStations(a.map, TR, 46340); assert.strictEqual(b.changed, false); assert.deepStrictEqual(b.map, a.map); });
test("trackStations", "a move → today, exact", () => { const a = N.trackStations(null, TR, 46335).map, m = TR.map((r) => r.slice()); m[1][C("S")] = "9 - Final Approval"; assert.deepStrictEqual(N.trackStations(a, m, 46340).map["A330-1-001.-"], ["9 - Final Approval", "2026-11-14", 1, "2026-11-14"]); });
const MAP = N.trackStations(null, TR, 46335).map;
test("stuck", "nothing on day one", () => assert.deepStrictEqual(N.stuck(MAP, 46335, 14), []));
test("stuck", "not exact: counted from the day first seen", () => assert.deepStrictEqual(N.stuck(MAP, 46350, 14).map((x) => [x.key, x.days, x.exact]), [["A330-1-001.-", 15, false], ["A330-1-002.-", 15, true]]));
test("stuck", "exactly the limit isn't stuck", () => assert.deepStrictEqual(N.stuck(MAP, 46349, 14), []));
test("stuck", "longest first", () => { const m = Object.assign({}, MAP, { x: ["s", "2026-10-01", 1, "2026-10-01"] }); assert.strictEqual(N.stuck(m, 46350, 14)[0].key, "x"); });
test("stuck", "empty map", () => assert.deepStrictEqual(N.stuck(null, 46350, 14), []));
test("bridgeStatus", "notes in E go without the ✉ lines", () => assert.strictEqual(N.bridgeStatus([HEAD, row({ C: "A330-1-1.-", E: "mine\n✉ from outlook", R: "Open" })]).rows[0][13], "mine"));
test("bridgeStatus", "days at station", () => { const b = N.bridgeStatus(TR, new Date(), "x", MAP); assert.deepStrictEqual(b.rows[0].slice(10, 13), ["2026-10-05", 0, "2026-11-09"]); });
test("bridgeStatus", "received as dd/mm/yyyy", () => assert.strictEqual(N.bridgeStatus(TR).rows[0][7], "05/10/2026"));
test("bridgeStatus", "only A330 keys", () => assert.deepStrictEqual(N.bridgeStatus([HEAD, row({ C: "A330-1-052 (another)", R: "Released" }), row({ C: "junk" })]).rows.map((r) => r[0]), ["A330-1-052 (another)"]));
test("bridgeStatus", "empty", () => assert.deepStrictEqual(N.bridgeStatus(null).rows, []));

const few = Object.entries(perFn).filter(([, n]) => n < 5);
if (few.length) process.exitCode = 1;
console.log(`\n${passed} tests passed, ${failed} failed` + (few.length ? " — fewer than 5 cases: " + few.map(([f, n]) => f + " (" + n + ")").join(", ") : " — every function has 5+ cases"));
