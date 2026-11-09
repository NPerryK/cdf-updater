// Run: node tests/look.test.js — the look of sheet CDF (110.6 / 110.7) and the looks from ChatGPT / Claude
const assert = require("assert");
const L = require("../cdf-look.js");
const A = require("../cdf-algo.js");
const E = require("../engine.js");

let passed = 0;
function test(name, fn) {
  try { fn(); passed++; console.log("  ✓ " + name); }
  catch (e) { console.error("  ✗ " + name + "\n    " + e.stack); process.exitCode = 1; }
}
const C = (l) => L.col(l);
function row(obj) { const r = new Array(34).fill(""); for (const [k, v] of Object.entries(obj)) r[C(k)] = v; return r; }
const HEAD = new Array(34).fill("").map((_, i) => "H" + L.letter(i));
HEAD[C("R")] = "CDF STATUS - GCC"; HEAD[C("S")] = "GCC CURRENT STATUS"; HEAD[C("F")] = "OPEN DATE"; HEAD[C("C")] = "LINK TO GCC";
// an invented sheet: 901 (2 continuation rows), 902 (none), 903 (1), then template rows with A only
function sheet() {
  return [HEAD,
    row({ C: "A330-1-901.-", R: "Open", S: "9 - Final Approval", F: 46300, U: "IN-WORK", AA: "D1" }),
    row({ R: "Open", U: "IN-WORK", AA: "D2" }),
    row({ R: "Open", U: "RELEASED", AA: "D3" }),
    row({ C: "A330-1-902.-", R: "Released", F: 46200, U: "RELEASED", AA: "D4" }),
    row({ C: "A330-1-903.A", R: "Cancelled", F: 46310, AA: "D5" }),
    row({ R: "Cancelled", AA: "D6" }),
    row({ A: "A330-1-904.-ATT." }),
    row({ A: "A330-1-905.-ATT." })];
}
const TODAY = 46330;

console.log("lastRow");
test("last row with anything", () => assert.strictEqual(L.lastRow(sheet()), 9));
test("trailing empty rows don't count", () => assert.strictEqual(L.lastRow(sheet().concat([row({}), row({ B: "  " })])), 9));
test("only the header", () => assert.strictEqual(L.lastRow([HEAD]), 1));
test("nothing", () => { assert.strictEqual(L.lastRow([]), 0); assert.strictEqual(L.lastRow(null), 0); });
test("a gap in the middle doesn't stop it", () => assert.strictEqual(L.lastRow([HEAD, row({}), row({ C: "X" })]), 3));

console.log("continuationRows");
test("rows after each anchor", () => assert.deepStrictEqual(L.continuationRows(sheet(), A), [3, 4, 7]));
test("anchors are never continuations", () => { const c = L.continuationRows(sheet(), A); [2, 5, 6].forEach((r) => assert.ok(c.indexOf(r) < 0)); });
test("template rows (A only) are not continuations", () => { const c = L.continuationRows(sheet(), A); assert.ok(c.indexOf(8) < 0 && c.indexOf(9) < 0); });
test("no blocks → none", () => assert.deepStrictEqual(L.continuationRows([HEAD], A), []));
test("a row before the first anchor is not one", () => assert.deepStrictEqual(L.continuationRows([HEAD, row({ R: "Open" }), row({ C: "K" })], A), []));

console.log("runsOf");
test("joins neighbours", () => assert.deepStrictEqual(L.runsOf([3, 4, 5, 9]), [[3, 5], [9, 9]]));
test("sorts first", () => assert.deepStrictEqual(L.runsOf([9, 3, 4]), [[3, 4], [9, 9]]));
test("duplicates are one", () => assert.deepStrictEqual(L.runsOf([3, 3, 4]), [[3, 4]]));
test("empty", () => assert.deepStrictEqual(L.runsOf([]), []));
test("single", () => assert.deepStrictEqual(L.runsOf([7]), [[7, 7]]));

console.log("contAddresses");
test("Q:X then AA:AH for each run", () => assert.deepStrictEqual(L.contAddresses([3, 4, 7]), ["Q3:X4, AA3:AH4, Q7:X7, AA7:AH7"]));
test("none → no address", () => assert.deepStrictEqual(L.contAddresses([]), []));
test("split by length", () => {
  const rows = []; for (let r = 2; r < 2000; r += 2) rows.push(r);
  const out = L.contAddresses(rows, 500);
  assert.ok(out.length > 1); out.forEach((a) => assert.ok(a.length <= 500));
  assert.strictEqual(out.join(", ").split(", ").length, rows.length * 2);
});
test("every row is in exactly one address", () => {
  const out = L.contAddresses([5, 6, 7, 20], 25), all = out.join(", ");
  assert.strictEqual((all.match(/Q5:X7/g) || []).length, 1); assert.strictEqual((all.match(/AA20:AH20/g) || []).length, 1);
});
test("default limit keeps ~250 rows in one rule", () => { const rows = []; for (let r = 2; r < 1000; r += 4) rows.push(r); assert.strictEqual(L.contAddresses(rows).length, 1); });

console.log("lookPlan");
test("ranges", () => { const p = L.lookPlan(sheet(), A); assert.strictEqual(p.all, "A1:AH9"); assert.strictEqual(p.body, "A2:AH9"); assert.strictEqual(p.purple, "S2:X9"); });
test("the three rules of 110.7", () => {
  const p = L.lookPlan(sheet(), A);
  assert.deepStrictEqual(p.rules.map((r) => r.formula), ['=TRIM($R2)="Released"', '=TRIM($R2)="Cancelled"', "=TRUE"]);
  assert.strictEqual(p.rules[0].fill, "#00B050"); assert.strictEqual(p.rules[1].fill, "#FF0000"); assert.strictEqual(p.rules[2].top, "#000000");
});
test("continuation rule on their Q:X and AA:AH", () => assert.deepStrictEqual(L.lookPlan(sheet(), A).rules[2].areas, ["Q3:X4, AA3:AH4, Q7:X7, AA7:AH7"]));
test("no continuations → =FALSE", () => {
  const p = L.lookPlan([HEAD, row({ C: "K1", R: "Open" }), row({ C: "K2", R: "Open" })], A);
  assert.strictEqual(p.rules[2].formula, "=FALSE"); assert.deepStrictEqual(p.rules[2].areas, ["Q2:X3, AA2:AH3"]);
});
test("nothing below the header → no plan", () => { assert.strictEqual(L.lookPlan([HEAD], A), null); assert.strictEqual(L.lookPlan([], A), null); });

console.log("ruleKind");
test("Released / Cancelled, any row", () => { assert.strictEqual(L.ruleKind('=TRIM($R2)="Released"'), "released"); assert.strictEqual(L.ruleKind('=TRIM($R1250)="Cancelled"'), "cancelled"); });
test("spaces and case", () => assert.strictEqual(L.ruleKind(' = trim( $R$2 ) = "released" '), "released"));
test("=TRUE / =FALSE are the continuation rule", () => { assert.strictEqual(L.ruleKind("=TRUE"), "cont"); assert.strictEqual(L.ruleKind("FALSE"), "cont"); });
test("a saved look", () => assert.strictEqual(L.ruleKind('=TRIM($R2)="Open"', ['=TRIM($R2)="Open"']), "look"));
test("anything else is the user's own", () => { assert.strictEqual(L.ruleKind('=$Q2>TODAY()'), null); assert.strictEqual(L.ruleKind('=TRIM($S2)="Released"'), null); assert.strictEqual(L.ruleKind(""), null); });

console.log("describeExisting");
const REAL = [{ formula: '=TRIM($R2)="Released"', areas: 5 }, { formula: '=TRIM($R2)="Cancelled"', areas: 5 }, { formula: "=TRUE", areas: 260 },
  { formula: '=TRIM($R1250)="Released"', areas: 1 }, { formula: '=TRIM($R1250)="Cancelled"', areas: 1 }, { formula: '=TRIM($R1250)="Released"', areas: 1 }, { formula: '=TRIM($R1250)="Cancelled"', areas: 1 }];
test("the 05/11 workbook: Released in 3 rules / 7 pieces", () => { const d = L.describeExisting(REAL); assert.strictEqual(d.released, 3); assert.strictEqual(d.pieces.released, 7); });
test("broken lists the split rules", () => assert.deepStrictEqual(L.describeExisting(REAL).broken, ["released", "cancelled"]));
test("own rules are counted, not ours", () => { const d = L.describeExisting(REAL.concat([{ formula: "=$Q2>1", areas: 1 }])); assert.strictEqual(d.ownRules, 1); });
test("a clean set is not broken", () => assert.deepStrictEqual(L.describeExisting([{ formula: '=TRIM($R2)="Released"', areas: 1 }, { formula: "=TRUE", areas: 300 }]).broken, []));
test("looks are recognised", () => assert.strictEqual(L.describeExisting([{ formula: "=X", areas: 1 }], ["=X"]).look, 1));

console.log("hiddenFormatCells");
test("finds ;;;", () => assert.deepStrictEqual(L.hiddenFormatCells([["General", ";;;"]], 2).map((x) => x.a1), ["B2"]));
test("with spaces", () => assert.strictEqual(L.hiddenFormatCells([[" ; ; ; "]], 5)[0].a1, "A5"));
test("other formats are fine", () => assert.deepStrictEqual(L.hiddenFormatCells([["0;-0;;@", "m/d/yyyy", ";;"]], 2), []));
test("beyond AH is not ours", () => { const r = new Array(40).fill("General"); r[36] = ";;;"; assert.deepStrictEqual(L.hiddenFormatCells([r], 2), []); });
test("row numbers follow r0", () => assert.deepStrictEqual(L.hiddenFormatCells([[], [], [";;;"]], 10).map((x) => x.row), [12]));

console.log("visibleRuns");
test("hidden rows split runs", () => assert.deepStrictEqual(L.visibleRuns([{ row: 2 }, { row: 3, hidden: true }, { row: 4 }, { row: 5 }]), [[2, 2], [4, 5]]));
test("all visible", () => assert.deepStrictEqual(L.visibleRuns([{ row: 2 }, { row: 3 }]), [[2, 3]]));
test("all hidden", () => assert.deepStrictEqual(L.visibleRuns([{ row: 2, hidden: true }]), []));
test("empty", () => assert.deepStrictEqual(L.visibleRuns([]), []));
test("unsorted input", () => assert.deepStrictEqual(L.visibleRuns([{ row: 5 }, { row: 4 }]), [[4, 5]]));

console.log("restoreHeights");
test("a taller row goes back", () => assert.deepStrictEqual(L.restoreHeights([{ row: 2, height: 51 }], { 2: 12.75 }), [{ row: 2, height: 51 }]));
test("a row that grew stays grown", () => assert.deepStrictEqual(L.restoreHeights([{ row: 2, height: 12.75 }], { 2: 25.5 }), []));
test("fresh rows fit their text", () => assert.deepStrictEqual(L.restoreHeights([{ row: 2, height: 18.75 }], { 2: 12.75 }, [2]), []));
test("hidden rows are left alone", () => assert.deepStrictEqual(L.restoreHeights([{ row: 2, height: 51, hidden: true }], { 2: 12.75 }), []));
test("no reading after → nothing", () => assert.deepStrictEqual(L.restoreHeights([{ row: 2, height: 51 }], {}), []));

console.log("hex");
test("#RRGGBB", () => assert.strictEqual(L.hex("#ffc000"), "#FFC000"));
test("without # and short", () => { assert.strictEqual(L.hex("ffc000"), "#FFC000"); assert.strictEqual(L.hex("#f80"), "#FF8800"); });
test("names", () => { assert.strictEqual(L.hex("Orange"), "#FFC000"); assert.strictEqual(L.hex("light-green"), "#C6EFCE"); });
test("empty → null", () => { assert.strictEqual(L.hex(""), null); assert.strictEqual(L.hex(null), null); });
test("not a color → undefined", () => { assert.strictEqual(L.hex("sunset"), undefined); assert.strictEqual(L.hex("#12345"), undefined); });

console.log("columnOf / paintColumns");
test("letters", () => { assert.strictEqual(L.columnOf("r"), C("R")); assert.strictEqual(L.columnOf("AH"), 33); });
test("beyond AH", () => assert.strictEqual(L.columnOf("AI"), -1));
test("a header name", () => assert.strictEqual(L.columnOf("cdf status – gcc", HEAD), C("R")));
test("row / default", () => { assert.deepStrictEqual(L.paintColumns("row"), [[0, 33]]); assert.deepStrictEqual(L.paintColumns(undefined), [[0, 33]]); });
test("ranges and lists merge", () => { assert.deepStrictEqual(L.paintColumns(["S", "T", "R"]), [[C("R"), C("T")]]); assert.deepStrictEqual(L.paintColumns("A:C, H"), [[0, 2], [7, 7]]); });
test("outside → null", () => { assert.strictEqual(L.paintColumns("A:AZ"), null); assert.strictEqual(L.paintColumns(["ZZ"]), null); });
test("paintAreas from row 2", () => assert.deepStrictEqual(L.paintAreas([[0, 2], [7, 7]], 99), ["A2:C99", "H2:H99"]));

console.log("isoDate / todaySerial");
test("ISO", () => assert.deepStrictEqual(L.isoDate("2026-11-05"), { y: 2026, m: 11, d: 5, serial: 46331 }));
test("dd/mm/yyyy", () => assert.strictEqual(L.isoDate("05/11/2026").serial, 46331));
test("not a date", () => { assert.strictEqual(L.isoDate("2026-02-30"), null); assert.strictEqual(L.isoDate("soon"), null); });
test("today in local time", () => assert.strictEqual(L.todaySerial(new Date(2026, 10, 5, 23, 59)), 46331));
test("1900 base", () => assert.strictEqual(L.todaySerial(new Date(1900, 0, 1)), 2));

console.log("normWhen");
test("a single condition", () => assert.deepStrictEqual(L.normWhen({ col: "R", op: "is", value: " Open " }, HEAD, 0).node, { kind: "cond", col: C("R"), op: "is", value: "Open" }));
test("aliases", () => { const n = L.normWhen({ column: "S", operator: "contains", value: "Final" }, HEAD, 0).node; assert.strictEqual(n.op, "has"); assert.strictEqual(n.col, C("S")); });
test("groups and a list as all", () => { const n = L.normWhen([{ col: "R", op: "is", value: "Open" }, { any: [{ col: "S", op: "empty" }, { not: { col: "F", op: ">", value: 1 } }] }], HEAD, 0).node; assert.strictEqual(n.kind, "all"); assert.strictEqual(n.kids[1].kind, "any"); assert.strictEqual(n.kids[1].kids[1].kind, "not"); });
test("unknown op says what is known", () => assert.ok(/isn't an operation.*older_than_days/.test(L.normWhen({ col: "R", op: "between", value: 1 }, HEAD, 0).error)));
test("bad column / missing number / bad date", () => {
  assert.ok(/column "ZZ"/.test(L.normWhen({ col: "ZZ", op: "is", value: 1 }, HEAD, 0).error));
  assert.ok(/needs a number/.test(L.normWhen({ col: "F", op: ">", value: "x" }, HEAD, 0).error));
  assert.ok(/whole number of days/.test(L.normWhen({ col: "F", op: "older_than_days", value: 2.5 }, HEAD, 0).error));
  assert.ok(/YYYY-MM-DD/.test(L.normWhen({ col: "F", op: "before", value: "next week" }, HEAD, 0).error));
});
test("one_of: list, dedupe, limit", () => {
  assert.deepStrictEqual(L.normWhen({ col: "R", op: "in", value: ["Open", "Open", "Released"] }, HEAD, 0).node.value, ["Open", "Released"]);
  assert.deepStrictEqual(L.normWhen({ col: "R", op: "one_of", value: "a, b|c" }, HEAD, 0).node.value, ["a", "b", "c"]);
  assert.ok(/20 values/.test(L.normWhen({ col: "R", op: "one_of", value: new Array(21).fill(0).map((_, i) => "v" + i) }, HEAD, 0).error));
});

console.log("toFormula");
const F = (w) => L.toFormula(L.normWhen(w, HEAD, 0).node);
test("is / is_not", () => { assert.strictEqual(F({ col: "R", op: "is", value: "Open" }), '=TRIM($R2)="Open"'); assert.strictEqual(F({ col: "R", op: "is_not", value: "Open" }), '=TRIM($R2)<>"Open"'); });
test("has escapes wildcards and quotes", () => assert.strictEqual(F({ col: "H", op: "has", value: 'a*b?"c~' }), '=ISNUMBER(SEARCH("a~*b~?""c~~",$H2))'));
test("dates", () => {
  assert.strictEqual(F({ col: "F", op: "older_than_days", value: 7 }), "=IF(ISNUMBER($F2),TODAY()-INT($F2)>7,FALSE)");
  assert.strictEqual(F({ col: "F", op: "before", value: "2026-11-01" }), "=IF(ISNUMBER($F2),INT($F2)<DATE(2026,11,1),FALSE)");
});
test("groups", () => assert.strictEqual(F({ all: [{ col: "R", op: "is", value: "Open" }, { not: { col: "S", op: "empty" } }] }), '=AND(TRIM($R2)="Open",NOT(LEN(TRIM($S2))=0))'));
test("one_of and numbers", () => {
  assert.strictEqual(F({ col: "R", op: "one_of", value: ["a", "b"] }), '=OR(TRIM($R2)="a",TRIM($R2)="b")');
  assert.strictEqual(F({ col: "F", op: ">=", value: 0.5 }), "=IF(ISNUMBER($F2),$F2>=0.5,FALSE)");
});

console.log("evalWhen / matchRows");
const N = (w) => L.normWhen(w, HEAD, 0).node;
test("is ignores case and extra spaces", () => assert.ok(L.evalWhen(N({ col: "R", op: "is", value: "open" }), row({ R: "  Open " }), TODAY)));
test("has / has_not", () => { const r = row({ S: "9 - Final Approval" }); assert.ok(L.evalWhen(N({ col: "S", op: "has", value: "final" }), r, TODAY)); assert.ok(!L.evalWhen(N({ col: "S", op: "has_not", value: "final" }), r, TODAY)); });
test("days need a real date", () => {
  assert.ok(L.evalWhen(N({ col: "F", op: "older_than_days", value: 7 }), row({ F: 46300 }), TODAY));
  assert.ok(!L.evalWhen(N({ col: "F", op: "older_than_days", value: 7 }), row({ F: "01/01/2020" }), TODAY));
  assert.ok(L.evalWhen(N({ col: "F", op: "within_days", value: 3 }), row({ F: 46329.7 }), TODAY));
});
test("empty / filled", () => { assert.ok(L.evalWhen(N({ col: "S", op: "empty" }), row({ S: "   " }), TODAY)); assert.ok(L.evalWhen(N({ col: "S", op: "filled" }), row({ S: 0 }), TODAY)); });
test("matchRows on the sheet", () => {
  const look = L.normLook({ when: { all: [{ col: "R", op: "is", value: "Open" }, { col: "S", op: "has", value: "Final Approval" }] }, paint: { fill: "orange" } }, HEAD, 0).look;
  assert.deepStrictEqual(L.matchRows(look, sheet(), TODAY), [2]);
  const all = L.normLook({ when: { col: "R", op: "filled" }, paint: { fill: "orange" } }, HEAD, 0).look;
  assert.deepStrictEqual(L.matchRows(all, sheet(), TODAY), [2, 3, 4, 5, 6, 7]);
});

console.log("sayWhen");
test("plain words with headers", () => assert.strictEqual(L.sayWhen(N({ col: "R", op: "is", value: "Open" }), HEAD), "R CDF STATUS - GCC is “Open”"));
test("and / or / not", () => assert.strictEqual(L.sayWhen(N({ any: [{ col: "S", op: "empty" }, { not: { col: "R", op: "is", value: "x" } }] }), HEAD), "(S GCC CURRENT STATUS is empty or not (R CDF STATUS - GCC is “x”))"));
test("days", () => assert.strictEqual(L.sayWhen(N({ col: "F", op: "older_than_days", value: 1 }), HEAD), "F OPEN DATE is more than 1 day ago"));
test("dates as dd/mm/yyyy", () => assert.strictEqual(L.sayWhen(N({ col: "F", op: "after", value: "2026-11-01" }), HEAD), "F OPEN DATE is after 01/11/2026"));
test("lists", () => assert.strictEqual(L.sayWhen(N({ col: "R", op: "one_of", value: ["A", "B"] }), []), "R is one of “A”, “B”"));

console.log("normLook");
test("defaults: whole row, a name", () => { const l = L.normLook({ when: { col: "R", op: "is", value: "Open" }, paint: { fill: "#ffc000" } }, HEAD, 2).look; assert.strictEqual(l.cells, "A:AH"); assert.strictEqual(l.name, "Look 3"); assert.strictEqual(l.fill, "#FFC000"); });
test("flat form (condition and paint at the top)", () => { const l = L.normLook({ name: "x", col: "R", op: "is", value: "Open", fill: "yellow", cells: "S:X" }, HEAD, 0).look; assert.strictEqual(l.cells, "S:X"); assert.strictEqual(l.fill, "#FFFF00"); });
test("nothing to paint", () => assert.ok(/nothing to paint/.test(L.normLook({ when: { col: "R", op: "filled" }, paint: {} }, HEAD, 0).error)));
test("bad color / cells", () => { assert.ok(/isn't a color/.test(L.normLook({ when: { col: "R", op: "filled" }, paint: { fill: "sunset" } }, HEAD, 0).error)); assert.ok(/outside A:AH/.test(L.normLook({ when: { col: "R", op: "filled" }, paint: { fill: "red", cells: "AK" } }, HEAD, 0).error)); });
test("too long for Excel", () => assert.ok(/Excel takes 255/.test(L.normLook({ when: { col: "R", op: "one_of", value: new Array(20).fill(0).map((_, i) => "a long status value " + i) }, paint: { bold: true } }, HEAD, 0).error)));

console.log("parseLooks");
const ANSWER = '{"cdf_look": 1, "looks": [{"name": "Stuck at FA", "when": {"all": [{"col": "R", "op": "is", "value": "Open"}, {"col": "S", "op": "has", "value": "Final Approval"}]}, "paint": {"cells": "row", "fill": "#FFC000"}}]}';
test("the format of the prompt", () => { const r = L.parseLooks(ANSWER, HEAD, E); assert.strictEqual(r.error, null); assert.strictEqual(r.looks[0].name, "Stuck at FA"); });
test("ChatGPT copy with 'json / Copy code' and a fence", () => { assert.strictEqual(L.parseLooks("json\nCopy code\n" + ANSWER, HEAD, E).looks.length, 1); assert.strictEqual(L.parseLooks("Here:\n```json\n" + ANSWER + "\n```\nDone.", HEAD, E).looks.length, 1); });
test("a single look, a list, smart quotes, trailing comma", () => {
  assert.strictEqual(L.parseLooks('{"when": {"col": "R", "op": "is", "value": "Open"}, "paint": {"bold": true},}', HEAD, E).looks.length, 1);
  assert.strictEqual(L.parseLooks('[{“when”: {“col”: “R”, “op”: “empty”}, “paint”: {“font”: “red”}}]', HEAD, E).looks.length, 1);
});
test("errors say which look", () => assert.ok(/^Look 2: /.test(L.parseLooks('[{"when":{"col":"R","op":"empty"},"paint":{"bold":true}},{"when":{"col":"R","op":"nope"},"paint":{"bold":true}}]', HEAD, E).error)));
test("empty / not JSON / a NEW ROWS reply", () => {
  assert.strictEqual(L.parseLooks("  ", HEAD, E).kind, "empty");
  assert.strictEqual(L.parseLooks("make it orange", HEAD, E).kind, "nojson");
  assert.strictEqual(L.parseLooks('{"rows": [{"cdf_number": "A330-1-1"}]}', HEAD, E).kind, "rows");
});

console.log("toStored / fromStored");
const LK = L.parseLooks(ANSWER, HEAD, E).looks[0];
test("round trip keeps the formula", () => assert.strictEqual(L.fromStored(JSON.parse(JSON.stringify(L.toStored(LK))), HEAD).formula, LK.formula));
test("round trip keeps the words", () => assert.strictEqual(L.fromStored(L.toStored(LK), HEAD).say, LK.say));
test("cells text rebuilt", () => assert.strictEqual(L.fromStored(L.toStored(LK), HEAD).cells, "A:AH"));
test("broken entry → null", () => { assert.strictEqual(L.fromStored(null), null); assert.strictEqual(L.fromStored({ name: "x" }), null); });
test("an unknown op stored → null", () => assert.strictEqual(L.fromStored({ when: { kind: "cond", col: 1, op: "zzz" }, spans: [[0, 1]] }), null));

console.log("buildPrompt / columnKinds");
const P = L.buildPrompt("orange: open at Final Approval", [{ letter: "R", header: "CDF STATUS\nGCC", kind: "text" }, { letter: "F", header: "OPEN DATE", kind: "date" }]);
test("has the request", () => assert.ok(P.indexOf("orange: open at Final Approval") > 0));
test("columns one per line with kind", () => { assert.ok(P.indexOf("R — CDF STATUS GCC\n") > 0); assert.ok(P.indexOf("F — OPEN DATE (date)") > 0); });
test("asks for one json block in the add-in's format", () => { assert.ok(/```json/.test(P)); assert.ok(P.indexOf('"cdf_look": 1') > 0); });
test("the example in the prompt is itself a valid answer", () => { const ex = /```json\n([\s\S]*?)```/.exec(P)[1]; assert.strictEqual(L.parseLooks(ex, HEAD, E).error, null); });
test("columnKinds: date / number / text", () => {
  const isDate = (f) => /d|y/.test(String(f || ""));
  assert.deepStrictEqual(L.columnKinds([[46300, 5, "a"], [46301, 6, "b"]], [["m/d/yyyy", "General", "General"], ["m/d/yyyy", "0", "@"]], isDate).slice(0, 3), ["date", "number", "text"]);
});

console.log("\n" + passed + " tests passed");
