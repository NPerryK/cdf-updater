// Run: node tests/summary.test.js — cdf-summary.js (the CDF SUMMARY sheet as cards), 5+ cases each function
const assert = require("assert");
const K = require("../cdf-core.js"), V = require("../cdf-summary.js");
let passed = 0, failed = 0; const perFn = {};
function test(fn, name, f) { perFn[fn] = (perFn[fn] || 0) + 1; try { f(); passed++; console.log("  ✓ " + fn + ": " + name); } catch (e) { failed++; console.error("  ✗ " + fn + ": " + name + "\n    " + e.stack); process.exitCode = 1; } }
const C = K.col;
function row(o) { const r = new Array(34).fill(""); for (const [k, v] of Object.entries(o)) r[C(k)] = v; return r; }
// sheet CDF (invented people)
const CDF = [new Array(34).fill("h"),
  row({ C: "A330-1-101.-", R: "Open", S: "3 - Conversion Engineer", P: "ANN", T: "Dana Levi", W: "Rejected", X: "TOM BAR" }),
  row({ R: "Open", W: "APP Design | Rejected", X: "RON AVI" }),
  row({ C: "A330-1-102.-", R: "Open", S: "9 - Final Approval", P: "TIMAT", T: "Ron Avi", W: "APP Design", X: "TOM BAR" }),
  row({ C: "A330-1-103.A", R: "Open", S: "3 - Conversion Engineer", P: "SIE", T: "Dana Levi" }),
  row({ C: "A330-1-104.-", R: "Released", S: "Released", P: "ANN", T: "-" }),
  row({ C: "A330-1-105.-", R: "Cancelled", P: "IAI" })];
// sheet CDF SUMMARY: tables down the sheet, three side by side, a TOTAL row, summary lines, merged person rows
const E = "";
const SUM = [
  ["CDF STATUS", E, E, E, E, E, E, E],
  ["CDF STATUS", "CDF'S", E, E, E, E, E, E],
  ["Open", 3, E, E, E, E, E, E], ["Cancelled", 1, E, E, E, E, E, E], ["Released", 1, E, E, E, E, E, E], ["TOTAL", 5, E, E, E, E, E, E],
  [E, E, E, E, E, E, E, E],
  ["GCC SIGNATURE POSITION", E, E, "DANA PENDING TABLE", E, E, "RON PENDING TABLE", E],
  ["SIGNATURE POSITION", "CDF'S", E, "CDF NUMBER", "SOLUTION / DOCUMENT", E, "CDF NUMBER", "SOLUTION / DOCUMENT"],
  ["Dana Levi", 2, E, "A330-1-101.-", "330-53-14-J1 Rev A", E, "A330-1-102.-", "ANSWER AT GCC"],
  ["Ron Avi", 1, E, "A330-1-103.A", "ANSWER AT GCC", E, "TOTAL", 1],
  ["TOTAL", 3, E, "TOTAL", 2, E, E, E],
  [E, E, E, E, E, E, E, E],
  ["WHERE ARE THE CDFS", E, E, E, E, E, E, E],
  ["SUPPLIER", "CDF'S", E, E, E, E, E, E], ["ANN/SIE", 2, E, E, E, E, E, E], ["TIMAT", 0, E, E, E, E, E, E], ["TOTAL", 2, E, E, E, E, E, E],
  [E, E, E, E, E, E, E, E],
  ["REJECTED", E, E, E, E, E, E, E],
  ["CDF NUMBER", "CDF STATUS", "DOCUMENT", "REVISION", "REJECTED BY", "REJECTED FUNCTION", E, E],
  ["A330-1-101.-", "Open", "330-53-14-J1", "A", "TOM BAR", "Check Design", E, E],
  ["REJECTED CDFS", E, E, E, E, 1.0000000000000002, E, E],
  ["REJECTED DOCUMENTS", E, E, E, E, 56.000000000000014, E, E],
  [E, E, E, E, E, E, E, E],
  ["PERSONS – GROUP", E, E, E, E, E, E, E],
  ["PERSON", "# of M", "MAIN GROUP", "SIGNATURE POSITION", "CDF NUMBERS", "PENDING MANTRA DOCUMENTS", E, E],
  ["Tom Bar", 2, "IAI", "APP Design", "A330-1-101.-", "330-1 Rev A", E, E],
  [E, E, E, "APP Design", "A330-1-102.-\nA330-1-103.A", "330-2 Rev -", E, E],
  ["ANN VENDOR", 1, "ANN/SIE", "Rejected", "A330-1-102.-", "330-3 Rev B", E, E],
  [E, E, E, E, E, E, E, E],
  ["CDFS RELEASED — LAST 7 DAYS (AS OF 09/11/2026 09:41 Asia/Jerusalem)", E, E, E, E, E, E, E],
  ["CDF NUMBER", "PREVIOUS STATUS", "NEW STATUS", E, E, E, E, E],
  ["A330-1-104.-", "Open", "Released", E, E, E, E, E],
  ["TOTAL", E, 1, E, E, E, E, E]];
const P = V.parse(SUM), T = (title) => P.tables.find((t) => t.title === title);

test("parse", "every table, side by side ones apart", () => assert.deepStrictEqual(P.tables.map((t) => t.title), ["CDF STATUS", "GCC SIGNATURE POSITION", "DANA PENDING TABLE", "RON PENDING TABLE", "WHERE ARE THE CDFS", "REJECTED", "PERSONS – GROUP", "CDFS RELEASED — LAST 7 DAYS"]));
test("parse", "header, rows and TOTAL", () => { const t = T("CDF STATUS"); assert.deepStrictEqual([t.head, t.rows.length, t.total, t.kind], [["CDF STATUS", "CDF'S"], 3, "5", "count"]); });
test("parse", "a side table keeps its own columns and TOTAL", () => { const t = T("DANA PENDING TABLE"); assert.deepStrictEqual([t.col, t.head, t.rows, t.total, t.kind], [3, ["CDF NUMBER", "SOLUTION / DOCUMENT"], [["A330-1-101.-", "330-53-14-J1 Rev A"], ["A330-1-103.A", "ANSWER AT GCC"]], "2", "list"]); });
test("parse", "summary lines, rounded", () => assert.deepStrictEqual(T("REJECTED").footer, [["REJECTED CDFS", "1"], ["REJECTED DOCUMENTS", "56"]]));
test("parse", "merged rows stay, an upper-case name inside a table is not a title", () => { const t = T("PERSONS – GROUP"); assert.strictEqual(t.rows.length, 3); assert.strictEqual(t.rows[2][0], "ANN VENDOR"); });
test("parse", "AS OF taken off the title", () => { assert.strictEqual(T("CDFS RELEASED — LAST 7 DAYS").asOf, "09/11/2026 09:41 Asia/Jerusalem"); assert.strictEqual(P.asOf, "09/11/2026 09:41 Asia/Jerusalem"); });
test("parse", "TOTAL in a middle column", () => assert.strictEqual(T("CDFS RELEASED — LAST 7 DAYS").total, "1"));
test("parse", "empty sheet → no tables", () => { assert.deepStrictEqual(V.parse([]).tables, []); assert.deepStrictEqual(V.parse(null).tables, []); });

const bl = K.buildBlocks(CDF);
test("linkColumn", "status counts over all CDFs from R", () => assert.deepStrictEqual(V.linkColumn(T("CDF STATUS"), CDF), { col: "R", scope: "all", off: 0 }));
test("linkColumn", "signature position from T", () => assert.strictEqual(V.linkColumn(T("GCC SIGNATURE POSITION"), CDF, bl).col, "T"));
test("linkColumn", "suppliers: open at Conversion Engineer, ANN and SIE are ANN/SIE", () => assert.deepStrictEqual(V.linkColumn(T("WHERE ARE THE CDFS"), CDF, bl), { col: "P", scope: "conv", off: 0 }));
test("linkColumn", "a list table has no link", () => assert.strictEqual(V.linkColumn(T("REJECTED"), CDF, bl), null));
test("linkColumn", "labels not in sheet CDF → no link", () => assert.strictEqual(V.linkColumn({ kind: "count", rows: [["Nobody", "4"], ["Noone", "2"]] }, CDF, bl), null));
test("linkColumn", "no sheet CDF → no link", () => assert.strictEqual(V.linkColumn(T("CDF STATUS"), null), null));

test("keysFor", "status", () => assert.deepStrictEqual(V.keysFor(CDF, { col: "R", scope: "all" }, "Open", bl), ["A330-1-101.-", "A330-1-102.-", "A330-1-103.A"]));
test("keysFor", "any row of the block counts", () => assert.deepStrictEqual(V.keysFor(CDF, { col: "X", scope: "open" }, "Ron Avi", bl), ["A330-1-101.-"]));
test("keysFor", "values split on |", () => assert.deepStrictEqual(V.keysFor(CDF, { col: "W", scope: "open" }, "Rejected", bl), ["A330-1-101.-"]));
test("keysFor", "case does not matter", () => assert.deepStrictEqual(V.keysFor(CDF, { col: "X", scope: "open" }, "tom bar", bl), ["A330-1-101.-", "A330-1-102.-"]));
test("keysFor", "the scope limits", () => assert.deepStrictEqual(V.keysFor(CDF, { col: "P", scope: "conv" }, "ANN/SIE", bl), ["A330-1-101.-", "A330-1-103.A"]));
test("keysFor", "no match → none", () => assert.deepStrictEqual(V.keysFor(CDF, { col: "T", scope: "all" }, "Nobody", bl), []));

const R = V.render(P, { cdf: CDF });
test("render", "a card for every table, and a chip to jump to it", () => { assert.strictEqual((R.html.match(/class="dcard sv"/g) || []).length, 8); assert.strictEqual((R.html.match(/data-goto=/g) || []).length, 8); });
test("render", "count rows open their CDFs", () => { const id = Object.keys(R.lists).find((k) => R.lists[k].title.indexOf("Open — CDF STATUS") === 0); assert.strictEqual(R.lists[id].keys.length, 3); });
test("render", "every data-list has its list", () => (R.html.match(/data-list="([^"]+)"/g) || []).forEach((m) => assert.ok(R.lists[m.slice(11, -1)], m)));
test("render", "CDF numbers in cells are links, two in one cell too", () => { assert.ok(R.html.indexOf('data-key="A330-1-103.A"') > 0); assert.ok(/A330-1-102.-<\/a> <a class="sv-key" data-key="A330-1-103.A"/.test(R.html)); });
test("render", "a zero row is drawn faint, not tappable", () => assert.ok(/class="sv-c zero"/.test(R.html)));
test("render", "long tables fold", () => { const r = V.render(P, { cdf: CDF, limit: 2 }); assert.ok(r.html.indexOf("Show all 3 rows") > 0 && /class=" more"/.test(r.html)); });
test("render", "escaped", () => assert.ok(V.render(V.parse([["<B>"], ["<x>", "y"], ["a", "b"]])).html.indexOf("&lt;x&gt;") > 0));
test("render", "no tables → says so", () => assert.ok(V.render({ tables: [] }).html.indexOf("No CDF SUMMARY") > 0));
test("render", "excluded CDFs are left out of the lists", () => { const r = V.render(P, { cdf: CDF, exclude: ["A330-1-104.-"] }); const id = Object.keys(r.lists).find((k) => r.lists[k].title.indexOf("Released — CDF STATUS") === 0); assert.ok(!id); });

test("tidy", "float noise", () => assert.strictEqual(V.tidy(56.000000000000014), "56"));
test("tidy", "two decimals", () => assert.strictEqual(V.tidy(12.345), "12.35"));
test("tidy", "numeric text", () => assert.strictEqual(V.tidy("52.00000000000001"), "52"));
test("tidy", "text trimmed", () => assert.strictEqual(V.tidy("  Open "), "Open"));
test("tidy", "empty", () => assert.strictEqual(V.tidy(""), ""));
test("splitTitle", "AS OF", () => assert.deepStrictEqual(V.splitTitle("X (AS OF 01/02/2026 10:00)"), { title: "X", asOf: "01/02/2026 10:00" }));
test("splitTitle", "none", () => assert.deepStrictEqual(V.splitTitle("CDF STATUS"), { title: "CDF STATUS", asOf: "" }));
test("splitTitle", "lower case as of", () => assert.strictEqual(V.splitTitle("Y (as of today)").asOf, "today"));
test("splitTitle", "other brackets stay", () => assert.strictEqual(V.splitTitle("A (B)").title, "A (B)"));
test("splitTitle", "dash in title", () => assert.strictEqual(V.splitTitle("CDFS RELEASED — LAST 24 HOURS (AS OF z)").title, "CDFS RELEASED — LAST 24 HOURS"));
test("isTitle", "upper case", () => assert.ok(V.isTitle("WHERE ARE THE CDFS")));
test("isTitle", "a name in mixed case is not", () => assert.ok(!V.isTitle("Dana Levi")));
test("isTitle", "a number is not", () => assert.ok(!V.isTitle(12) && !V.isTitle("12")));
test("isTitle", "the (AS OF …) part may hold lower case", () => assert.ok(V.isTitle("CDFS RELEASED (AS OF 09/11 Asia/Jerusalem)")));
test("isTitle", "empty is not", () => assert.ok(!V.isTitle("") && !V.isTitle(null)));

const few = Object.entries(perFn).filter(([, n]) => n < 5);
if (few.length) process.exitCode = 1;
console.log(`\n${passed} tests passed, ${failed} failed` + (few.length ? " — fewer than 5 cases: " + few.map(([f, n]) => f + " (" + n + ")").join(", ") : " — every function has 5+ cases"));
