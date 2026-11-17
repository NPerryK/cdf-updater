// Run: node tests/stats.test.js — cdf-stats.js, cdf-charts.js, cdf-dash.js, cdf-deck.js (5+ cases each function)
const assert = require("assert");
const K = require("../cdf-core.js"), S = require("../cdf-stats.js"), G = require("../cdf-charts.js"), D = require("../cdf-dash.js"), DK = require("../cdf-deck.js");
let passed = 0, failed = 0; const perFn = {};
function test(fn, name, f) { perFn[fn] = (perFn[fn] || 0) + 1; try { f(); passed++; console.log("  ✓ " + fn + ": " + name); } catch (e) { failed++; console.error("  ✗ " + fn + ": " + name + "\n    " + e.stack); process.exitCode = 1; } }
const C = K.col, T = 46335, NOW = new Date(2026, 10, 9, 10, 0);     // Mon 09/11/2026
function row(o) { const r = new Array(34).fill(""); for (const [k, v] of Object.entries(o)) r[C(k)] = v; return r; }
const CDF = [new Array(34).fill("h"),
  row({ C: "A330-1-001.-", R: "Open", S: "3 - Conversion Engineer", P: "ANN", F: T - 10, U: "NO ENG. DOC." }),
  row({ C: "A330-1-002.-", R: "Open", S: "9 - Final Approval", P: "TIMAT", F: T - 100, AA: "2", AB: "B" }),
  row({ R: "Open", AA: "3", AB: "B" }),
  row({ C: "A330-1-003.-", R: "Released", S: "Released", F: T - 50, Q: T - 2, AA: "1", AB: "A" }),
  row({ C: "A330-1-004.-", R: "Cancelled", F: T - 60 }),
  row({ C: "A330-1-005.-", R: "Open", S: "1 - Initiator", P: "-", F: T - 1, AA: "NEW-9", AB: "-" }),
  row({ C: "A330-1-052 (another)", R: "Released" })];
const GCCH = ["CDF Number", "CDF Status", "Open Date", "Close Date", "Current Station"];
const GCC = [GCCH, ["A330-1-001.-", "Open", T - 12, "", "Conversion ENG"], ["A330-1-002.-", "Open", T - 200, "", "Final Approval"], ["A330-1-003.-", "Released", T - 55, T - 2, "Released"], ["A330-1-009.-", "Released", T - 9, T - 8, "Released"]];
const REVH = ["Project", "Doc ID", "Rev", "Doc Title", "Release Date", "Days Cal", "Doc Type"];
const REV = [REVH, ["A330", "1", "A", "ONE", T - 1, 10, "XD"], ["A330", "2", "B", "TWO", T - 5, 20, "DWG"], ["A330", "3", "A", "THREE", T - 20, 30, "XD"], ["A330", "4", "A", "OLD", T - 400, 5, "XD"]];
const M = S.build({ cdf: CDF, gcc: GCC, rev: REV }, { now: NOW, exclude: ["A330-1-052 (another)"] });

test("build", "status from the anchors, the duplicate excluded", () => assert.deepStrictEqual([M.status.open, M.status.released, M.status.cancelled, M.status.total], [3, 1, 1, 5]));
test("build", "% released of the not-cancelled", () => assert.strictEqual(M.status.pct, 25));
test("build", "stations in station order", () => assert.deepStrictEqual(M.stations.map((s) => s.name + ":" + s.n), ["1 - Initiator:1", "3 - Conversion Engineer:1", "9 - Final Approval:1"]));
test("build", "aging uses the GCC open date first", () => assert.deepStrictEqual(M.aging.oldest.map((o) => o.key + ":" + o.days), ["A330-1-002.-:200", "A330-1-001.-:12", "A330-1-005.-:1"]));
test("build", "aging buckets", () => assert.deepStrictEqual(M.aging.buckets.map((b) => b.n), [1, 1, 0, 0, 1]));
test("build", "suppliers at Conversion Engineer", () => assert.deepStrictEqual(M.suppliers.conv.map((s) => s.name + ":" + s.n), ["ANN/SIE:1"]));
test("build", "an unknown supplier is Unassigned", () => assert.ok(M.suppliers.all.some((s) => s.name === "Unassigned")));
test("build", "released this week from the closed date", () => assert.deepStrictEqual(M.week.closedKeys, ["A330-1-003.-"]));
test("build", "opened in the last 7 days", () => assert.deepStrictEqual(M.week.openedKeys, ["A330-1-005.-"]));
test("build", "waiting for an engineering document", () => assert.deepStrictEqual(M.noDoc.keys, ["A330-1-001.-"]));
test("build", "not in GCC", () => assert.deepStrictEqual(M.notInGcc, ["A330-1-005.-"]));
test("build", "freshness of GCC and MANTRAREV", () => { assert.strictEqual(M.fresh.gcc, T - 2); assert.strictEqual(M.docs.latest, T - 1); assert.strictEqual(M.fresh.rev, T - 1); });
test("build", "no CDF sheet → empty", () => { const e = S.build({ cdf: null, gcc: GCC, rev: REV }, { now: NOW }); assert.ok(e.empty); assert.ok(!e.status); });
test("build", "weekly: 12 weeks, this week last", () => { assert.strictEqual(M.weekly.length, 12); assert.strictEqual(M.weekly[11].week, K.weekStart(T)); });
test("build", "trend: stored days + today", () => { const m = S.build({ cdf: CDF }, { now: NOW, history: [{ d: T - 1, open: 9, released: 1 }, { d: T, open: 99, released: 0 }] }); assert.deepStrictEqual(m.trend.map((t) => t.open), [9, 3]); });
test("cdfsOf", "one per anchor", () => assert.strictEqual(S.cdfsOf(CDF).length, 6));
test("cdfsOf", "documents counted over the block", () => assert.strictEqual(S.cdfsOf(CDF)[1].docs, 2));
test("cdfsOf", "row numbers are 1-based", () => assert.strictEqual(S.cdfsOf(CDF)[0].row, 2));
test("cdfsOf", "status words normalised", () => assert.deepStrictEqual(S.cdfsOf([CDF[0], row({ C: "K", R: " open " }), row({ C: "L", R: "Canceled" })]).map((x) => x.status), ["Open", "Cancelled"]));
test("cdfsOf", "empty", () => assert.deepStrictEqual(S.cdfsOf(null), []));
test("gccIndex", "by key base", () => assert.strictEqual(S.gccIndex(GCC)["A330-1-002"].open, T - 200));
test("gccIndex", "close and status", () => assert.deepStrictEqual(S.gccIndex(GCC)["A330-1-003"], { open: T - 55, close: T - 2, status: "Released" }));
test("gccIndex", "latest date seen", () => assert.strictEqual(S.gccIndex(GCC).latest, T - 2));
test("gccIndex", "no CDF Number column", () => assert.deepStrictEqual(S.gccIndex([["x"], ["y"]]), {}));
test("gccIndex", "empty", () => assert.deepStrictEqual(S.gccIndex(null), {}));
const RI = S.revIndex(REV);
test("revIndex", "documents and their released revisions", () => assert.deepStrictEqual(Object.keys(RI.ids).sort(), ["1", "2", "3", "4"]));
test("revIndex", "the newest release date", () => assert.strictEqual(RI.latest, T - 1));
test("revIndex", "upper case, blank rev is -", () => { const x = S.revIndex([["Doc ID", "Rev", "Release Date"], ["ab-1 ", "", T], ["AB-1", "a", T]]); assert.deepStrictEqual(Object.keys(x.ids["AB-1"]).sort(), ["-", "A"]); });
test("revIndex", "no header → error", () => assert.ok(S.revIndex([["a", "b"], [1, 2]]).error));
test("revIndex", "no sheet → null", () => assert.strictEqual(S.revIndex(null), null));
test("docState", "the revision is in MANTRAREV → released", () => assert.deepStrictEqual(S.docState(RI, "2", "B"), { state: "released", released: "B" }));
test("docState", "another revision released → in work, says which", () => assert.deepStrictEqual(S.docState(RI, "3", "B"), { state: "work", released: "A" }));
test("docState", "not in MANTRAREV → in work, nothing released", () => assert.deepStrictEqual(S.docState(RI, "X-1", "-"), { state: "work", released: "" }));
test("docState", "case and spaces", () => assert.strictEqual(S.docState(RI, " 2 ", "b").state, "released"));
test("docState", "the newest released revision", () => { const x = S.revIndex([["Doc ID", "Rev", "Release Date"], ["D", "-", T], ["D", "C", T], ["D", "A", T]]); assert.strictEqual(S.docState(x, "D", "E").released, "C"); });
test("docsOf", "released and in work, open CDFs only", () => assert.deepStrictEqual([M.docs.released.map((d) => d.doc), M.docs.work.map((d) => d.doc)], [["2"], ["3", "NEW-9"]]));
test("docsOf", "the CDFs behind each", () => assert.deepStrictEqual([M.docs.keys.released, M.docs.keys.work], [["A330-1-002.-"], ["A330-1-002.-", "A330-1-005.-"]]));
test("docsOf", "no-document markers skipped", () => assert.ok(!M.docs.work.concat(M.docs.released).some((d) => /^-$|ANSWER|NO ENG/i.test(d.doc))));
test("docsOf", "a document shared by two CDFs counts once", () => { const c = CDF.concat([row({ C: "A330-1-007.-", R: "Open", AA: "3", AB: "b" })]); const m = S.build({ cdf: c, rev: REV }, { now: NOW }); assert.deepStrictEqual(m.docs.work.find((d) => d.doc === "3").keys, ["A330-1-002.-", "A330-1-007.-"]); });
test("docsOf", "no MANTRAREV → null", () => assert.strictEqual(S.build({ cdf: CDF }, { now: NOW }).docs, null));
test("weekList", "12 Sundays, last is this week", () => { const w = S.weekList(T); assert.strictEqual(w.length, 12); assert.strictEqual(w[11], 46334); assert.strictEqual(w[0], 46334 - 77); });
test("weekList", "a custom count", () => assert.strictEqual(S.weekList(T, 4).length, 4));
test("weekList", "7 days apart", () => { const w = S.weekList(T); for (let i = 1; i < w.length; i++) assert.strictEqual(w[i] - w[i - 1], 7); });
test("weekList", "a Sunday today", () => assert.strictEqual(S.weekList(46334)[11], 46334));
test("weekList", "a Saturday", () => assert.strictEqual(S.weekList(46340)[11], 46334));
test("snapshot", "open, released, stations", () => assert.deepStrictEqual(S.snapshot(M), { d: T, open: 3, released: 1, st: { "1 - Initiator": 1, "3 - Conversion Engineer": 1, "9 - Final Approval": 1 } }));
test("snapshot", "empty model", () => assert.deepStrictEqual(S.snapshot({ today: T }), { d: T, open: 0, released: 0, st: {} }));
test("snapshot", "day is the model's day", () => assert.strictEqual(S.snapshot(Object.assign({}, M, { today: 5 })).d, 5));
test("snapshot", "no stations", () => assert.deepStrictEqual(S.snapshot({ today: 1, status: { open: 2, released: 3 } }).st, {}));
test("snapshot", "JSON safe", () => assert.doesNotThrow(() => JSON.stringify(S.snapshot(M))));
test("mergeHistory", "one per day, newest wins", () => assert.deepStrictEqual(S.mergeHistory([{ d: 1, open: 5 }, { d: 2, open: 6 }], { d: 2, open: 7 }), [{ d: 1, open: 5 }, { d: 2, open: 7 }]));
test("mergeHistory", "sorted", () => assert.deepStrictEqual(S.mergeHistory([{ d: 3 }, { d: 1 }], { d: 2 }).map((x) => x.d), [1, 2, 3]));
test("mergeHistory", "400 at most", () => assert.strictEqual(S.mergeHistory(Array.from({ length: 450 }, (_, i) => ({ d: i })), { d: 1000 }).length, 400));
test("mergeHistory", "from nothing", () => assert.deepStrictEqual(S.mergeHistory(null, { d: 1 }), [{ d: 1 }]));
test("mergeHistory", "junk entries go", () => assert.deepStrictEqual(S.mergeHistory([null, { d: 1 }], { d: 2 }).map((x) => x.d), [1, 2]));
test("deltaWeek", "vs a week ago", () => assert.deepStrictEqual(S.deltaWeek({ trend: [{ d: 1, open: 10, released: 5 }, { d: 8, open: 7, released: 9 }] }), { open: -3, released: 4, since: 1 }));
test("deltaWeek", "the oldest when less than a week", () => assert.strictEqual(S.deltaWeek({ trend: [{ d: 5, open: 10, released: 5 }, { d: 8, open: 7, released: 9 }] }).since, 5));
test("deltaWeek", "the closest one a week or more back", () => assert.strictEqual(S.deltaWeek({ trend: [{ d: 1, open: 1, released: 0 }, { d: 2, open: 2, released: 0 }, { d: 10, open: 3, released: 0 }] }).since, 2));
test("deltaWeek", "one day → null", () => assert.strictEqual(S.deltaWeek({ trend: [{ d: 1, open: 1, released: 0 }] }), null));
test("deltaWeek", "no trend → null", () => assert.strictEqual(S.deltaWeek({}), null));
test("supplierOf", "aliases", () => assert.deepStrictEqual(["ANN", "SIE/ANN", "ancra", "Timat", "IAI"].map(S.supplierOf), ["ANN/SIE", "ANN/SIE", "ANCARA", "TIMAT", "IAI"]));
test("supplierOf", "- and empty → Unassigned", () => assert.deepStrictEqual(["-", "", null].map(S.supplierOf), ["Unassigned", "Unassigned", "Unassigned"]));
test("supplierOf", "another name stays", () => assert.strictEqual(S.supplierOf("Boeing"), "Boeing"));
test("supplierOf", "spaces", () => assert.strictEqual(S.supplierOf(" ANN/SIE "), "ANN/SIE"));
test("supplierOf", "SIE alone", () => assert.strictEqual(S.supplierOf("SIE"), "ANN/SIE"));
test("stationOrder", "numbered", () => assert.ok(S.stationOrder("2 - POC") < S.stationOrder("10 - Production Engineering")));
test("stationOrder", "released last", () => assert.strictEqual(S.stationOrder("Released"), 100));
test("stationOrder", "unknown in between", () => assert.strictEqual(S.stationOrder("-"), 50));
test("stationOrder", "9 before 10", () => assert.ok(S.stationOrder("9 - Final Approval") < S.stationOrder("10 - X")));
test("stationOrder", "1", () => assert.strictEqual(S.stationOrder("1 - Initiator"), 1));

// ---------- charts ----------
test("hbars", "one bar per item, values shown", () => { const s = G.hbars([{ label: "A", n: 3 }, { label: "B", n: 1 }]); assert.strictEqual((s.match(/class="ch-bar /g) || []).length, 2); assert.ok(s.indexOf(">3</text>") > 0); });
test("hbars", "a list id makes it clickable", () => assert.ok(G.hbars([{ label: "A", n: 3, list: "x" }]).indexOf('data-list="x"') > 0));
test("hbars", "labels are escaped", () => assert.ok(G.hbars([{ label: "<b>", n: 1 }]).indexOf("&lt;b&gt;") > 0));
test("hbars", "zero is an empty bar", () => assert.ok(/class="ch-bar c-open" x="\d+" y="\d+" width="0"/.test(G.hbars([{ label: "A", n: 0 }, { label: "B", n: 2 }]))));
test("hbars", "long labels are cut", () => assert.ok(G.hbars([{ label: "x".repeat(40), n: 1 }]).indexOf("…") > 0));
test("columns2", "two columns per week", () => assert.strictEqual((G.columns2([{ week: T, a: 1, b: 2 }, { week: T + 7, a: 3, b: 0 }]).match(/class="ch-col /g) || []).length, 4));
test("columns2", "average lines", () => assert.strictEqual((G.columns2([{ week: T, a: 1, b: 2 }], { avg: [{ v: 1, cls: "c-open", label: "x" }, { v: 0, cls: "c-rel", label: "y" }] }).match(/ch-avg/g) || []).length, 1));
test("columns2", "this week is marked", () => assert.ok(G.columns2([{ week: T, a: 1, b: 2 }]).indexOf(" cur") > 0));
test("columns2", "clickable when there's a list", () => assert.ok(G.columns2([{ week: T, a: 1, b: 0, aList: "l" }]).indexOf('data-list="l"') > 0));
test("columns2", "all zero still draws", () => assert.ok(G.columns2([{ week: T, a: 0, b: 0 }]).indexOf("<svg") === 0));
test("line", "a path through the points", () => assert.ok(/<path class="ch-ln c-open" d="M[\d.]+ [\d.]+ L/.test(G.line([{ d: 1, v: 3 }, { d: 2, v: 5 }]))));
test("line", "one point → a message", () => assert.ok(G.line([{ d: 1, v: 3 }]).indexOf("chart-empty") > 0));
test("line", "a target line", () => assert.ok(G.line([{ d: 1, v: 3 }, { d: 2, v: 5 }], { target: { value: 1, by: 10 } }).indexOf("ch-target-path") > 0));
test("line", "custom empty text", () => assert.ok(G.line([], { empty: "later" }).indexOf("later") > 0));
test("line", "dates on the axis", () => assert.ok(G.line([{ d: T - 1, v: 3 }, { d: T, v: 5 }]).indexOf("09/11") > 0));
test("ring", "the % in the middle", () => assert.ok(G.ring(76.1).indexOf(">76.1%</text>") > 0));
test("ring", "0%", () => assert.ok(/stroke-dasharray="0 /.test(G.ring(0))));
test("ring", "over 100 is full", () => { const s = G.ring(130), m = /stroke-dasharray="([\d.]+) ([\d.]+)"/.exec(s); assert.strictEqual(m[1], m[2]); });
test("ring", "size", () => assert.ok(G.ring(5, { size: 50 }).indexOf('width="50"') > 0));
test("ring", "label", () => assert.ok(G.ring(5, { label: "released" }).indexOf(">released<") > 0));
test("ring", "a small ring keeps only the %", () => assert.ok(G.ring(5, { label: "released", size: 62 }).indexOf("released<") < 0));
test("strip", "one part per status", () => assert.strictEqual((G.strip([{ n: 1, cls: "a", label: "A" }, { n: 3, cls: "b", label: "B" }]).match(/<rect/g) || []).length, 2));
test("strip", "widths add up", () => { const s = G.strip([{ n: 1, cls: "a", label: "A" }, { n: 3, cls: "b", label: "B" }], { width: 100 }); assert.ok(s.indexOf('width="25"') > 0 && s.indexOf('width="75"') > 0); });
test("strip", "all zero", () => assert.ok(G.strip([{ n: 0, cls: "a", label: "A" }]).indexOf("<svg") === 0));
test("strip", "titles for hover", () => assert.ok(G.strip([{ n: 2, cls: "a", label: "Open" }]).indexOf("<title>Open: 2</title>") > 0));
test("strip", "aria-hidden (the numbers are in the text)", () => assert.ok(G.strip([{ n: 2, cls: "a", label: "x" }]).indexOf('aria-hidden="true"') > 0));
test("niceMax", "small", () => assert.strictEqual(G.niceMax(3), 5));
test("niceMax", "12 → 20", () => assert.strictEqual(G.niceMax(12), 20));
test("niceMax", "42 → 50", () => assert.strictEqual(G.niceMax(42), 50));
test("niceMax", "108 → 200", () => assert.strictEqual(G.niceMax(108), 200));
test("niceMax", "100 stays", () => assert.strictEqual(G.niceMax(100), 100));

// ---------- dashboard ----------
const R0 = D.render(M, { mode: "pane" });
test("render", "the cards", () => assert.deepStrictEqual(R0.cards.map((c) => c.id), ["status", "stations", "flow", "aging", "suppliers", "attention", "trend", "docs"]));
test("render", "every data-list has its list", () => { (R0.html.match(/data-list="([^"]+)"/g) || []).forEach((m) => { const id = m.slice(11, -1).replace(/&amp;/g, "&"); assert.ok(R0.lists[id], id); }); });
test("render", "the lists carry the keys", () => assert.deepStrictEqual(R0.lists["status:open"].keys, ["A330-1-001.-", "A330-1-002.-", "A330-1-005.-"]));
test("render", "empty model", () => assert.ok(D.render({ empty: true }).html.indexOf("No CDF sheet") > 0));
test("render", "solution documents card", () => assert.ok(/revisions in work/.test(R0.html) && R0.lists["doc:work"].keys.length === 2));
test("render", "present mode draws wider charts", () => assert.ok(D.render(M, { mode: "present" }).html.indexOf('viewBox="0 0 520') > 0));
test("paceText", "growing", () => assert.ok(/grows by about/.test(D.paceText({ pace: { weeks: 8, closed: 2, opened: 5, net: -3 }, status: { open: 10 } }))));
test("paceText", "a date when they run out", () => assert.ok(/around 21\/12\/2026/.test(D.paceText({ pace: { weeks: 8, closed: 5, opened: 2, net: 3, weeksToZero: 6, zeroOn: T + 42 }, status: { open: 18 } }))));
test("paceText", "steady", () => assert.ok(/holds steady/.test(D.paceText({ pace: { weeks: 8, closed: 2, opened: 2, net: 0 }, status: { open: 10 } }))));
test("paceText", "no pace", () => assert.strictEqual(D.paceText({ pace: { weeks: 0 } }), ""));
test("paceText", "says the weeks", () => assert.ok(/last 8 full weeks/.test(D.paceText({ pace: { weeks: 8, closed: 2, opened: 2, net: 0 }, status: { open: 1 } }))));
test("summaryText", "the date and the numbers", () => { const t = D.summaryText(M); assert.ok(t.indexOf("09/11/2026") > 0 && t.indexOf("Open 3") > 0 && t.indexOf("(25%)") > 0); });
test("summaryText", "stations", () => assert.ok(D.summaryText(M).indexOf("Conversion Engineer 1") > 0));
test("summaryText", "empty", () => assert.strictEqual(D.summaryText({ empty: true }), ""));
test("summaryText", "suppliers at Conversion Engineer", () => assert.ok(D.summaryText(M).indexOf("At Conversion Engineer: ANN/SIE 1") > 0));
test("summaryText", "the pace", () => assert.ok(/a week/.test(D.summaryText(M))));
test("summaryHtml", "a table", () => assert.ok(D.summaryHtml(M).indexOf("<table") > 0));
test("summaryHtml", "escaped", () => assert.ok(D.summaryHtml(Object.assign({}, M, { stations: [{ name: "<x>", n: 1 }] })).indexOf("&lt;x&gt;") > 0));
test("summaryHtml", "empty", () => assert.strictEqual(D.summaryHtml({ empty: true }), ""));
test("summaryHtml", "inline styles only (emails drop <style>)", () => assert.ok(D.summaryHtml(M).indexOf("<style") < 0));
test("summaryHtml", "stations listed", () => assert.ok(D.summaryHtml(M).indexOf("at 9 - Final Approval") > 0));
test("stationShort", "drops the number", () => assert.strictEqual(D.stationShort("3 - Conversion Engineer"), "Conversion Engineer"));
test("stationShort", "released", () => assert.strictEqual(D.stationShort("Released"), "Released"));
test("stationShort", "10", () => assert.strictEqual(D.stationShort("10 - Production Engineering"), "Production Engineering"));
test("stationShort", "-", () => assert.strictEqual(D.stationShort("-"), "-"));
test("stationShort", "no spaces", () => assert.strictEqual(D.stationShort("2-POC"), "POC"));

// ---------- deck (with the real PptxGenJS when it's installed next to the tests) ----------
let Pptx = null; try { Pptx = require(process.env.PPTXGENJS || "pptxgenjs"); } catch (e) { /* the deck tests below use a recorder */ }
function recorder() { const log = []; function P() { this.slides = []; } P.prototype.addSlide = function () { const s = { charts: [], texts: [], tables: [], addChart: (t, d, o) => s.charts.push({ t, d, o }), addText: (t) => s.texts.push(t), addShape: () => {}, addTable: (r) => s.tables.push(r) }; this.slides.push(s); return s; }; P.prototype.ChartType = { bar: "bar", line: "line", doughnut: "doughnut" }; P.prototype.ShapeType = { rect: "rect" }; return P; }
const P = recorder();
test("deck build", "7 slides with a trend, 6 without", () => { assert.strictEqual(DK.build(M, P).slides.length, 6); const m2 = Object.assign({}, M, { trend: [{ d: T - 1, open: 2 }, { d: T, open: 3 }] }); assert.strictEqual(DK.build(m2, P).slides.length, 7); });
test("deck build", "native charts with the model's numbers", () => { const d = DK.build(M, P); assert.deepStrictEqual(d.slides[0].charts[0].d[0].values, [1, 3, 1]); });
test("deck build", "no MANTRAREV → no documents slide", () => assert.strictEqual(DK.build(Object.assign({}, M, { docs: null }), P).slides.length, 5));
test("deck build", "empty model → one slide", () => assert.strictEqual(DK.build({ empty: true, today: T }, P).slides.length, 1));
test("deck build", "the real library writes a .pptx", () => { if (!Pptx) return; const d = DK.build(M, Pptx); assert.strictEqual(typeof d.write, "function"); });
test("deck fileName", "by date", () => assert.strictEqual(DK.fileName(M), "CDF_Status_20261109.pptx"));
test("deck fileName", "another day", () => assert.strictEqual(DK.fileName({ today: 46000 }), "CDF_Status_" + K.ddmmyyyy(46000).split("/").reverse().join("") + ".pptx"));
test("deck fileName", ".pptx", () => assert.ok(/\.pptx$/.test(DK.fileName(M))));
test("deck fileName", "no spaces", () => assert.ok(!/\s/.test(DK.fileName(M))));
test("deck fileName", "8 digits", () => assert.ok(/_\d{8}\.pptx$/.test(DK.fileName(M))));

const few = Object.entries(perFn).filter(([, n]) => n < 5);
if (few.length) process.exitCode = 1;
console.log(`\n${passed} tests passed, ${failed} failed` + (few.length ? " — fewer than 5 cases: " + few.map(([f, n]) => f + " (" + n + ")").join(", ") : " — every function has 5+ cases"));
