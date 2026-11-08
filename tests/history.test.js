// Run: node tests/history.test.js   (synthetic data only — no real names or documents)
// Stage 4 of the CDF algorithm: BACKUPDATA, the events, the release tables and sheet ERROR (cdf-history.js). At least 5 cases for every function.
const assert = require("assert");
const H = require("../cdf-history.js");
const A = require("../cdf-algo.js");
const M = require("../cdf-mantra.js");
const F = require("../features.js");

let passed = 0, failed = 0;
const perFn = {};
function test(fn, name, body) {
  perFn[fn] = (perFn[fn] || 0) + 1;
  try { body(); passed++; console.log("  ✓ " + fn + ": " + name); }
  catch (e) { failed++; console.error("  ✗ " + fn + ": " + name + "\n    " + e.stack); process.exitCode = 1; }
}
const C = A.col;
const T = (s) => Date.parse(s);

// ---------------------------------------------------------------- fixtures: yesterday and today of a small workbook
function row(key, R, x) { const r = new Array(28).fill(""); r[C("A")] = key ? key.replace(/\.[^.]*$/, "") : ""; r[C("C")] = key; r[C("D")] = key ? key.split(".").pop() : ""; r[C("R")] = R; Object.keys(x || {}).forEach((k) => { r[C(k)] = x[k]; }); return r; }
function cont(R, x) { return row("", R, x); }
const HEAD = new Array(28).fill("h");
const CDF_Y = [HEAD,
  row("A330-1-001.-", "Open", { U: "3 - PRE-RELEASE", V: "IAI - Design Engineering", W: "APP Design", X: "Dana Levi", AA: "330-11-11-J0001", AB: "A" }),
  row("A330-1-002.-", "Open", { U: "NO ENG. DOC.", V: "-", W: "-", X: "-", AA: "-", AB: "-" }),
  row("A330-1-003.-", "Open", { U: "2 - IN WORK", V: "IAI - Stress Department", W: "Rejected | Stress", X: "Ron Avi", AA: "330-22-22-J0002", AB: "A" }),
  row("A330-1-004.-", "Released", { Q: 46174.27634259259, U: "4 - RELEASED", V: "-", W: "-", X: "-", AA: "330-33-33-J0003", AB: "B" }),
  cont("Released", { U: "4 - RELEASED", V: "-", W: "-", X: "-", AA: "330-33-33-J0003", AB: "B" })
];
const CDF_T = [HEAD,
  row("A330-1-001.-", "Released", { Q: 46331, U: "4 - RELEASED", V: "-", W: "-", X: "-", AA: "330-11-11-J0001", AB: "A" }),
  row("A330-1-002.-", "Open", { U: "ANSWER AT GCC", V: "-", W: "-", X: "-", AA: "ANSWER AT GCC", AB: "-" }),
  row("A330-1-003.-", "Open", { U: "2 - IN WORK", V: "IAI - Stress Department", W: "Rejected | Stress", X: "Ron Avi", AA: "330-22-22-J0002", AB: "B" }),
  row("A330-1-004.-", "Released", { Q: 46174.27634259259, U: "4 - RELEASED", V: "-", W: "-", X: "-", AA: "330-33-33-J0003", AB: "B" }),
  cont("Released", { U: "4 - RELEASED", V: "-", W: "-", X: "-", AA: "330-33-33-J0003", AB: "B" })
];
const MH = ["Doc ID", "Rev", "Status", "Date Status", "# Rejects", "Rejected By Emp", "Rejected By Function", "Pending To Emp", "Pending To Function"];
const mrow = (id, rev, st, dt, rej) => [id, rev, st, dt, rej || 0, "", "", "", ""];
const MANTRA_Y = [MH,
  mrow("330-11-11-J0001", "A", "PRE-RELEASED WF", "2026-11-01T10:00:00"),
  mrow("330-22-22-J0002", "A", "IN-WORK", "2026-10-20T09:00:00", 2),
  mrow("330-22-22-J0002", "B", "IN-WORK", "2026-11-03T09:00:00"),
  mrow("330-33-33-J0003", "B", "RELEASED", "2026-10-01T08:00:00"),
  mrow("330-99-99-J0009", "A", "PRE-RELEASED", "2026-10-30T08:00:00")
];
const MANTRA_T = [MH,
  mrow("330-11-11-J0001", "A", "RELEASED", "2026-11-05T07:30:00", 1),
  mrow("330-22-22-J0002", "A", "IN-WORK", "2026-10-20T09:00:00", 2),
  mrow("330-22-22-J0002", "B", "IN-WORK", "2026-11-03T09:00:00"),
  mrow("330-33-33-J0003", "B", "RELEASED", "2026-10-01T08:00:00"),
  mrow("330-99-99-J0009", "A", "RELEASED", "2026-11-04T12:00:00")
];
const gccRow = (k) => { const r = new Array(20).fill(""); r[0] = k; return r; };
const GCC = [["CDF"], gccRow("A330-1-004.-"), gccRow("A330-1-001.-"), gccRow("A330-1-002.-"), gccRow("A330-1-003.-")];
const NOW_Y = new Date("2026-11-04T06:00:00Z"), NOW_T = new Date("2026-11-05T06:00:00Z");
function run(now, startCdf, finalCdf, mantra, runs, base, history, errors) {
  return H.buildRun({ now, started: now, algo: A, M, runs: runs || [], base: base || null, history: history || [], startCdf, finalCdf, mantra, mantraRev: [], gcc: GCC, errors: errors || {} });
}
const finals = (r) => ({ cdfs: r.cdfs.filter((x) => x[1] === "FINAL"), solutions: r.solutions.filter((x) => x[1] === "FINAL"), mantradocs: r.mantradocs });
const Y = run(NOW_Y, CDF_Y, CDF_Y, MANTRA_Y);
const TD = run(NOW_T, CDF_T, CDF_T, MANTRA_T, [Y.run], finals(Y), Y.events);
const evOf = (r, type) => r.events.filter((e) => e[2] === type);
const CAT_T = M.readMantra(MANTRA_T);

// ---------------------------------------------------------------- time
test("isoUtc", "a Date → ISO without milliseconds", () => assert.strictEqual(H.isoUtc(new Date("2026-11-05T05:57:14.789Z")), "2026-11-05T05:57:14Z"));
test("isoUtc", "a number (ms) works too", () => assert.strictEqual(H.isoUtc(T("2026-01-02T03:04:05Z")), "2026-01-02T03:04:05Z"));
test("isoUtc", "always UTC, whatever the zone in the string", () => assert.strictEqual(H.isoUtc(new Date("2026-11-05T08:00:00+02:00")), "2026-11-05T06:00:00Z"));
test("isoUtc", "pads month, day and time", () => assert.strictEqual(H.isoUtc(new Date(Date.UTC(2026, 0, 1, 0, 0, 0))), "2026-01-01T00:00:00Z"));
test("isoUtc", "the end of a year", () => assert.strictEqual(H.isoUtc(new Date(Date.UTC(2026, 11, 31, 23, 59, 59))), "2026-12-31T23:59:59Z"));

test("runIdOf", "RUN- and the UTC time without separators", () => assert.strictEqual(H.runIdOf(new Date("2026-11-05T05:57:14Z")), "RUN-20261105055714"));
test("runIdOf", "the macro's format (the 04/11 run)", () => assert.strictEqual(H.runIdOf(new Date("2026-11-04T06:52:08Z")), "RUN-20261104065208"));
test("runIdOf", "milliseconds don't show", () => assert.strictEqual(H.runIdOf(new Date("2026-11-05T05:57:14.999Z")), "RUN-20261105055714"));
test("runIdOf", "Israel time in → UTC id out", () => assert.strictEqual(H.runIdOf(new Date("2026-11-05T07:57:14+02:00")), "RUN-20261105055714"));
test("runIdOf", "two seconds apart → two ids", () => assert.notStrictEqual(H.runIdOf(new Date(T("2026-11-05T05:57:14Z"))), H.runIdOf(new Date(T("2026-11-05T05:57:15Z")))));

test("parseIso", "an ISO UTC string → ms", () => assert.strictEqual(H.parseIso("2026-11-05T05:57:14Z"), T("2026-11-05T05:57:14Z")));
test("parseIso", "spaces around are fine", () => assert.strictEqual(H.parseIso("  2026-11-05T05:57:14Z "), T("2026-11-05T05:57:14Z")));
test("parseIso", "empty → null", () => assert.strictEqual(H.parseIso(""), null));
test("parseIso", "text → null", () => assert.strictEqual(H.parseIso("yesterday"), null));
test("parseIso", "null → null", () => assert.strictEqual(H.parseIso(null), null));

test("jlmWall", "winter: UTC+2", () => assert.strictEqual(H.jlmWall(T("2026-11-05T06:00:00Z")) - T("2026-11-05T06:00:00Z"), 2 * 3600000));
test("jlmWall", "summer: UTC+3", () => assert.strictEqual(H.jlmWall(T("2026-07-01T06:00:00Z")) - T("2026-07-01T06:00:00Z"), 3 * 3600000));
test("jlmWall", "the Friday DST starts: before and after 02:00", () => assert.deepStrictEqual([H.jlmWall(T("2026-03-26T23:30:00Z")) - T("2026-03-26T23:30:00Z"), H.jlmWall(T("2026-03-27T00:30:00Z")) - T("2026-03-27T00:30:00Z")], [2 * 3600000, 3 * 3600000]));
test("jlmWall", "the Sunday DST ends: before and after 02:00", () => assert.deepStrictEqual([H.jlmWall(T("2026-10-24T22:30:00Z")) - T("2026-10-24T22:30:00Z"), H.jlmWall(T("2026-10-24T23:30:00Z")) - T("2026-10-24T23:30:00Z")], [3 * 3600000, 2 * 3600000]));
test("jlmWall", "no time zones (Intl fails) → the Israeli rule gives the same", () => {
  const real = Intl.DateTimeFormat;
  const ts = ["2026-11-05T06:00:00Z", "2026-07-01T06:00:00Z", "2026-03-26T23:30:00Z", "2026-03-27T00:30:00Z", "2026-10-24T22:30:00Z", "2026-10-24T23:30:00Z"].map(T);
  const want = ts.map((t) => H.jlmWall(t));
  Intl.DateTimeFormat = function () { throw new Error("no ICU"); };
  try { assert.deepStrictEqual(ts.map((t) => H.jlmWall(t)), want); } finally { Intl.DateTimeFormat = real; }
});

test("jlmWall", "asked again for the same hour → the same (the offset is kept per hour)", () => { const t = T("2026-11-05T06:12:34Z"); assert.strictEqual(H.jlmWall(t), H.jlmWall(t + 1000) - 1000); });
test("jlmWall", "milliseconds are dropped, as the formatter does", () => assert.strictEqual(H.jlmWall(T("2026-11-05T06:00:00Z") + 999), H.jlmWall(T("2026-11-05T06:00:00Z"))));
test("jlmToUtc", "winter wall 07:30 → 05:30Z", () => assert.strictEqual(H.jlmToUtc(Date.UTC(2026, 10, 5, 7, 30)), T("2026-11-05T05:30:00Z")));
test("jlmToUtc", "summer wall 12:00 → 09:00Z", () => assert.strictEqual(H.jlmToUtc(Date.UTC(2026, 6, 1, 12, 0)), T("2026-07-01T09:00:00Z")));
test("jlmToUtc", "the inverse of jlmWall (winter)", () => { const t = T("2026-12-15T10:11:12Z"); assert.strictEqual(H.jlmToUtc(H.jlmWall(t)), t); });
test("jlmToUtc", "the inverse of jlmWall (summer)", () => { const t = T("2026-08-15T10:11:12Z"); assert.strictEqual(H.jlmToUtc(H.jlmWall(t)), t); });
test("jlmToUtc", "the day after DST starts", () => assert.strictEqual(H.jlmToUtc(Date.UTC(2026, 2, 28, 9, 0)), T("2026-03-28T06:00:00Z")));

test("mantraTime", "an ISO local string → local text and UTC", () => assert.deepStrictEqual(H.mantraTime("2026-11-05T07:30:00"), { local: "05/11/2026 07:30:00", utc: "2026-11-05T05:30:00Z", ms: T("2026-11-05T05:30:00Z") }));
test("mantraTime", "an Excel serial (Israel time)", () => assert.strictEqual(H.mantraTime(46331.3125).utc, "2026-11-05T05:30:00Z"));
test("mantraTime", "dd/mm/yyyy hh:mm", () => assert.strictEqual(H.mantraTime("05/11/2026 07:30").local, "05/11/2026 07:30:00"));
test("mantraTime", "a date only is midnight local", () => assert.strictEqual(H.mantraTime("2026-11-05").utc, "2026-11-04T22:00:00Z"));
test("mantraTime", "text, empty, zero → null", () => assert.deepStrictEqual([H.mantraTime("soon"), H.mantraTime(""), H.mantraTime(0)], [null, null, null]));

test("jlmShort", "winter", () => assert.strictEqual(H.jlmShort(T("2026-11-05T05:30:00Z")), "05/11/2026 07:30"));
test("jlmShort", "summer", () => assert.strictEqual(H.jlmShort(T("2026-07-01T05:30:00Z")), "01/07/2026 08:30"));
test("jlmShort", "crosses midnight into the next day", () => assert.strictEqual(H.jlmShort(T("2026-11-04T22:30:00Z")), "05/11/2026 00:30"));
test("jlmShort", "no seconds", () => assert.strictEqual(H.jlmShort(T("2026-11-05T05:30:59Z")), "05/11/2026 07:30"));
test("jlmShort", "the AS OF of the 05/11 run", () => assert.strictEqual(H.jlmShort(T("2026-11-05T06:10:00Z")), "05/11/2026 08:10"));

test("excelNumber", "15 significant digits, as Excel writes it as text", () => assert.strictEqual(H.excelNumber(46174.27634259259), "46174.2763425926"));
test("excelNumber", "a whole number stays whole", () => assert.strictEqual(H.excelNumber(46174), "46174"));
test("excelNumber", "floating noise is gone", () => assert.strictEqual(H.excelNumber(0.1 + 0.2), "0.3"));
test("excelNumber", "text numbers work", () => assert.strictEqual(H.excelNumber("12.50"), "12.5"));
test("excelNumber", "zero", () => assert.strictEqual(H.excelNumber(0), "0"));

test("letter", "A", () => assert.strictEqual(H.letter(0), "A"));
test("letter", "Z", () => assert.strictEqual(H.letter(25), "Z"));
test("letter", "AA", () => assert.strictEqual(H.letter(26), "AA"));
test("letter", "BD (EVENTS)", () => assert.strictEqual(H.letter(55), "BD"));
test("letter", "CG (the last column)", () => assert.strictEqual(H.letter(84), "CG"));
test("col", "A → 0", () => assert.strictEqual(H.col("A"), 0));
test("col", "S → 18", () => assert.strictEqual(H.col("S"), 18));
test("col", "AG → 32", () => assert.strictEqual(H.col("AG"), 32));
test("col", "BX → 75", () => assert.strictEqual(H.col("BX"), 75));
test("col", "the inverse of letter", () => [0, 1, 25, 26, 52, 84].forEach((i) => assert.strictEqual(H.col(H.letter(i)), i)));

// ---------------------------------------------------------------- the sheet BACKUPDATA
test("headerRow", "85 columns, A:CG", () => assert.strictEqual(H.headerRow().length, 85));
test("headerRow", "each area starts with RunID in its column", () => Object.keys(H.AREAS).forEach((a) => assert.strictEqual(H.headerRow()[H.col(H.AREAS[a])], "RunID")));
test("headerRow", "the gaps between areas are empty", () => assert.deepStrictEqual([H.headerRow()[16], H.headerRow()[17], H.headerRow()[30]], ["", "", ""]));
test("headerRow", "RUNS ends at P with ValidationStatus", () => assert.strictEqual(H.headerRow()[H.col("P")], "ValidationStatus"));
test("headerRow", "EVENTS ends at BU with Notes", () => assert.strictEqual(H.headerRow()[H.col("BU")], "Notes"));

test("checkHeaders", "its own header row is right", () => assert.deepStrictEqual(H.checkHeaders(H.headerRow()), []));
test("checkHeaders", "a wrong header names its cell", () => { const h = H.headerRow(); h[H.col("BF")] = "Type"; assert.deepStrictEqual(H.checkHeaders(h), ['BF1 is "Type", not "EventType"']); });
test("checkHeaders", "an empty row → every header", () => assert.strictEqual(H.checkHeaders([]).length, 77));
test("checkHeaders", "spaces around a header are fine", () => { const h = H.headerRow(); h[0] = " RunID "; assert.deepStrictEqual(H.checkHeaders(h), []); });
test("checkHeaders", "an area moved one column → reported", () => { const h = H.headerRow(); h.splice(18, 0, ""); assert.ok(H.checkHeaders(h).length > 30); });

test("areaRows", "the columns of one area, RunID first", () => { const v = [H.headerRow(), H.headerRow().map((x, i) => "v" + i)]; assert.deepStrictEqual(H.areaRows(v, "events")[0].slice(0, 2), ["v55", "v56"]); });
test("areaRows", "rows without a RunID in the area are left out", () => { const r = new Array(85).fill(""); r[0] = "RUN-1"; assert.deepStrictEqual([H.areaRows([[], r], "runs").length, H.areaRows([[], r], "cdfs").length], [1, 0]); });
test("areaRows", "the sheet row is kept", () => { const r = new Array(85).fill(""); r[18] = "RUN-1"; assert.strictEqual(H.areaRows([[], [], r], "cdfs")[0].sheetRow, 3); });
test("areaRows", "as wide as the area", () => { const r = new Array(85).fill("x"); assert.strictEqual(H.areaRows([[], r], "solutions")[0].length, 21); });
test("areaRows", "short rows are padded with empty values", () => { const r = ["RUN-1"]; assert.deepStrictEqual(H.areaRows([[], r], "runs")[0].slice(0, 3), ["RUN-1", "", ""]); });

const RUNROW = (id, at, st, err, val) => [id, at, at, at, st, "", at, "Asia/Jerusalem", "BACKUPDATA-v3", "1", "1", "1", "1", "0", String(err), val];
test("baselineRun", "the last committed, passed run with no errors", () => assert.strictEqual(H.baselineRun([RUNROW("R1", "2026-11-03T06:00:00Z", "COMMITTED", 0, "PASSED"), RUNROW("R2", "2026-11-04T06:00:00Z", "COMMITTED", 0, "PASSED")]).runId, "R2"));
test("baselineRun", "by RunAsOfUTC, not by row order", () => assert.strictEqual(H.baselineRun([RUNROW("R2", "2026-11-04T06:00:00Z", "COMMITTED", 0, "PASSED"), RUNROW("R1", "2026-11-03T06:00:00Z", "COMMITTED", 0, "PASSED")]).runId, "R2"));
test("baselineRun", "a FAILED run is never the baseline", () => assert.strictEqual(H.baselineRun([RUNROW("R1", "2026-11-03T06:00:00Z", "COMMITTED", 0, "PASSED"), RUNROW("R2", "2026-11-04T06:00:00Z", "FAILED", 2, "FAILED")]).runId, "R1"));
test("baselineRun", "a run with errors isn't either, even marked COMMITTED", () => assert.strictEqual(H.baselineRun([RUNROW("R1", "2026-11-03T06:00:00Z", "COMMITTED", 0, "PASSED"), RUNROW("R2", "2026-11-04T06:00:00Z", "COMMITTED", 1, "PASSED")]).runId, "R1"));
test("baselineRun", "none → null", () => assert.strictEqual(H.baselineRun([RUNROW("R2", "2026-11-04T06:00:00Z", "FAILED", 2, "FAILED")]), null));
test("baselineRun", "TrackingStartedAtUTC carried; the run time when it's empty", () => { const r = RUNROW("R1", "2026-11-03T06:00:00Z", "COMMITTED", 0, "PASSED"); r[6] = ""; assert.strictEqual(H.baselineRun([r]).tracking, "2026-11-03T06:00:00Z"); });

test("baselineRun", "only a run whose snapshot is still in BACKUPDATA", () => assert.strictEqual(H.baselineRun([RUNROW("R1", "2026-11-03T06:00:00Z", "COMMITTED", 0, "PASSED"), RUNROW("R2", "2026-11-04T06:00:00Z", "COMMITTED", 0, "PASSED")], (id) => id === "R1").runId, "R1"));
test("baselineRun", "no run with a snapshot → null (history starts again)", () => assert.strictEqual(H.baselineRun([RUNROW("R1", "2026-11-03T06:00:00Z", "COMMITTED", 0, "PASSED")], () => false), null));

// ---------------------------------------------------------------- what BACKUPDATA keeps
const NOWR = T("2026-11-07T06:00:00Z");
const RUNS7 = [RUNROW("R1", "2026-10-30T06:00:00Z", "COMMITTED", 0, "PASSED"), RUNROW("R2", "2026-11-05T06:00:00Z", "COMMITTED", 0, "PASSED"), RUNROW("R3", "2026-11-06T06:00:00Z", "COMMITTED", 0, "PASSED")];
const ROWS7 = { cdfs: ["R1", "R1", "R2", "R2", "R3", "R3"], solutions: ["R1", "R2", "R2", "R3"], mantradocs: ["R2", "R2", "R3", "R3", "R3"] };
const ev = (run, detected, time) => { const e = new Array(18).fill(""); e[0] = run; e[11] = detected; e[13] = time; return e; };
const EV7 = [ev("R1", "2026-10-29T06:00:00Z", "2026-10-28T06:00:00Z"), ev("R2", "2026-11-05T06:00:00Z", "2026-11-05T06:00:00Z"), ev("RX", "2026-11-06T06:00:00Z", "2026-11-06T06:00:00Z"), ev("R3", "2026-11-06T06:00:00Z", "2026-10-20T06:00:00Z")];
const RET = H.retention({ runs: RUNS7, snap: { R1: true, R2: true, R3: true }, rows: ROWS7, eventRows: EV7, now: NOWR, current: { runId: "R4", commits: true } });
test("retention", "keeps the last committed run (with this one: two)", () => assert.deepStrictEqual(RET.keep, ["R3"]));
test("retention", "drops every other run's snapshot", () => assert.deepStrictEqual(RET.drop.sort(), ["R1", "R2"]));
test("retention", "segments are sheet rows, bottom-up", () => assert.deepStrictEqual([RET.segments.cdfs, RET.segments.mantradocs], [[[2, 5]], [[2, 3]]]));
test("retention", "row counts per area", () => assert.deepStrictEqual([RET.rows.cdfs, RET.rows.solutions, RET.rows.mantradocs], [4, 3, 2]));
test("retention", "events: older than 8 days, or of a run that isn't committed, go; detected recently stays", () => assert.deepStrictEqual([RET.segments.events, RET.events], [[[4, 4], [2, 2]], 2]));
test("retention", "two runs to drop is the usual rolling trim; more is the first (bulk) trim", () => {
  assert.strictEqual(RET.bulk, false);
  const many = H.retention({ runs: RUNS7.concat([RUNROW("R0", "2026-10-20T06:00:00Z", "COMMITTED", 0, "PASSED")]), snap: { R0: true, R1: true, R2: true, R3: true }, rows: { cdfs: ["R0", "R1", "R2", "R3"] }, eventRows: [], now: NOWR, current: { runId: "R4", commits: true } });
  assert.deepStrictEqual([many.bulk, many.drop.length], [true, 3]);
});
test("retention", "a failed run moves nothing", () => { const r = H.retention({ runs: RUNS7, snap: { R1: true, R2: true, R3: true }, rows: ROWS7, eventRows: EV7, now: NOWR, current: { runId: "R4", commits: false } }); assert.deepStrictEqual([r.skipped, r.drop, r.segments.cdfs], [true, [], []]); });
test("retention", "an undone run (FAILED) isn't kept; the run before it is", () => {
  const runs = RUNS7.map((r) => r.slice()); runs[2][4] = "FAILED";
  const r = H.retention({ runs, snap: { R1: true, R2: true, R3: true }, rows: ROWS7, eventRows: [], now: NOWR, current: { runId: "R4", commits: true } });
  assert.deepStrictEqual([r.keep, r.drop.sort()], [["R2"], ["R1", "R3"]]);
});
test("retention", "a run without a snapshot can't be kept as the baseline", () => assert.deepStrictEqual(H.retention({ runs: RUNS7, snap: { R1: true, R2: true }, rows: ROWS7, eventRows: [], now: NOWR, current: { runId: "R4", commits: true } }).keep, ["R2"]));
test("retention", "empty rows split segments and are left alone", () => assert.deepStrictEqual(H.retention({ runs: RUNS7, snap: { R3: true }, rows: { cdfs: ["R1", "", "R1", "R3"] }, eventRows: [], now: NOWR, current: { runId: "R4", commits: true } }).segments.cdfs, [[4, 4], [2, 2]]));
test("retention", "the first run in a new BACKUPDATA: nothing to drop", () => assert.deepStrictEqual(H.retention({ runs: [], snap: {}, rows: {}, eventRows: [], now: NOWR, current: { runId: "R1", commits: true } }).drop, []));

test("committedRuns", "committed runs are in", () => assert.deepStrictEqual(H.committedRuns([RUNROW("R1", "x", "COMMITTED", 0, "PASSED")]), { R1: true }));
test("committedRuns", "failed runs are out", () => assert.deepStrictEqual(H.committedRuns([RUNROW("R1", "x", "FAILED", 1, "FAILED")]), {}));
test("committedRuns", "empty", () => assert.deepStrictEqual(H.committedRuns([]), {}));
test("committedRuns", "spaces around the id are trimmed", () => assert.deepStrictEqual(H.committedRuns([RUNROW(" R1 ", "x", "COMMITTED", 0, "PASSED")]), { R1: true }));
test("committedRuns", "an undone run (marked FAILED afterwards) no longer counts", () => { const r = RUNROW("R1", "x", "COMMITTED", 0, "PASSED"); r[4] = "FAILED"; assert.deepStrictEqual(H.committedRuns([r]), {}); });

// ---------------------------------------------------------------- snapshots
const CS = H.cdfSnapshot(CDF_Y, A, "RUN-1", "FINAL", "2026-11-04T06:00:00Z");
test("cdfSnapshot", "one row per CDF, not per solution row", () => assert.strictEqual(CS.length, 4));
test("cdfSnapshot", "identity, key, number, revision, status", () => assert.deepStrictEqual(CS[0].slice(0, 8), ["RUN-1", "FINAL", "2026-11-04T06:00:00Z", "K:A330-1-001.-", "A330-1-001.-", "A330-1-001", "-", "Open"]));
test("cdfSnapshot", "a closure date keeps its time, as Excel text", () => assert.strictEqual(CS[3][8], "46174.2763425926"));
test("cdfSnapshot", "AnchorRow is the sheet row", () => assert.deepStrictEqual(CS.map((r) => r[9]), ["2", "3", "4", "5"]));
test("cdfSnapshot", "PresentInCDF and RecordID", () => assert.deepStrictEqual(CS[1].slice(10), ["TRUE", "RUN-1|FINAL|C|K:A330-1-002.-"]));
test("cdfSnapshot", "a text closure date is kept as text", () => { const v = [HEAD, row("A330-1-009.-", "Released", { Q: "12/06/2025", AA: "-" })]; assert.strictEqual(H.cdfSnapshot(v, A, "R", "START", "t")[0][8], "12/06/2025"); });

const SS = H.solutionSnapshot(CDF_Y, A, { cat: M.readMantra(MANTRA_Y), M }, "RUN-1", "FINAL", "t");
test("solutionSnapshot", "one row per solution row", () => assert.strictEqual(SS.length, 5));
test("solutionSnapshot", "types: engineering document, no document", () => assert.deepStrictEqual(SS.map((r) => r[6]).slice(0, 2), ["ENGINEERING DOCUMENT", "NO DOCUMENT"]));
test("solutionSnapshot", "SolutionIdentity: the document, or SOLUTION", () => assert.deepStrictEqual([SS[0][5], SS[1][5]], ["K:A330-1-001.-|330-11-11-J0001", "K:A330-1-002.-|SOLUTION"]));
test("solutionSnapshot", "the same document twice in a CDF → SourceOrder 1 and 2", () => assert.deepStrictEqual([SS[3][16], SS[4][16]], ["1", "2"]));
test("solutionSnapshot", "a released document gets its Mantra time as VerifiedReleaseAtUTC", () => assert.deepStrictEqual([SS[3][17], SS[3][18]], ["2026-10-01T05:00:00Z", H.RELEASE_BASIS]));
test("solutionSnapshot", "RejectCount from Mantra; MatchType EXACT", () => assert.deepStrictEqual([SS[2][13], SS[2][15]], ["2", "EXACT"]));
test("solutionSnapshot", "MatchType NONE / ANSWER for non-engineering rows", () => { const s = H.solutionSnapshot(CDF_T, A, { cat: CAT_T, M }, "R", "FINAL", "t"); assert.deepStrictEqual([SS[1][15], s[1][15]], ["NONE", "ANSWER"]); });
test("solutionSnapshot", "an unchanged row keeps the baseline's match, rejects and release time", () => {
  const base = SS[0].slice(); base[13] = "7"; base[15] = "CORRECTED_MANUAL";
  const s = H.solutionSnapshot(CDF_Y, A, { cat: M.readMantra(MANTRA_Y), M, carry: { [base[5] + "#1"]: base } }, "R", "FINAL", "t");
  assert.deepStrictEqual([s[0][13], s[0][15]], ["7", "CORRECTED_MANUAL"]);
});
test("solutionSnapshot", "a changed row is resolved now (stage 2's match)", () => {
  const base = SS[2].slice(); base[15] = "OLD";
  const s = H.solutionSnapshot(CDF_T, A, { cat: CAT_T, M, carry: { [base[5] + "#1"]: base }, matchOf: () => "CORRECTED_MANUAL" }, "R", "FINAL", "t");
  assert.strictEqual(s[2][15], "CORRECTED_MANUAL");
});

test("solutionType", "a document id", () => assert.strictEqual(H.solutionType("330-11-11-J0001"), "ENGINEERING DOCUMENT"));
test("solutionType", "ANSWER AT GCC", () => assert.strictEqual(H.solutionType("ANSWER AT GCC"), "GCC ANSWER"));
test("solutionType", "a dash", () => assert.strictEqual(H.solutionType("-"), "NO DOCUMENT"));
test("solutionType", "empty", () => assert.strictEqual(H.solutionType(""), "NO DOCUMENT"));
test("solutionType", "lower case answer", () => assert.strictEqual(H.solutionType("answer at gcc"), "GCC ANSWER"));

test("recordOf", "the latest Date Status of the doc–revision", () => { const v = [MH, mrow("330-11-11-J0001", "A", "PRE-RELEASED", "2026-11-01T10:00:00"), mrow("330-11-11-J0001", "A", "RELEASED", "2026-11-05T07:30:00")]; assert.strictEqual(H.recordOf(M.readMantra(v), "330-11-11-J0001", "A", M).status, "RELEASED"); });
test("recordOf", "a document Mantra doesn't have → null", () => assert.strictEqual(H.recordOf(CAT_T, "330-00-00-J9999", "A", M), null));
test("recordOf", "a revision Mantra doesn't have → null", () => assert.strictEqual(H.recordOf(CAT_T, "330-11-11-J0001", "Z", M), null));
test("recordOf", "the id is normalised (dashes, case)", () => assert.ok(H.recordOf(CAT_T, "330‑11‑11‑j0001", "A", M)));
test("recordOf", "no catalog → null", () => assert.strictEqual(H.recordOf(null, "330-11-11-J0001", "A", M), null));

const MD = H.mantraDocs(MANTRA_T, "RUN-2", "2026-11-05T06:00:00Z");
test("mantraDocs", "every row as it is", () => assert.strictEqual(MD.rows.length, 5));
test("mantraDocs", "status upper-case, local and UTC time, source row", () => assert.deepStrictEqual(MD.rows[0].slice(3, 9), ["330-11-11-J0001", "A", "RELEASED", "05/11/2026 07:30:00", "2026-11-05T05:30:00Z", "2"]));
test("mantraDocs", "RecordID", () => assert.strictEqual(MD.rows[0][9], "RUN-2|FINAL|M|330-11-11-J0001|A|2"));
test("mantraDocs", "the Date Status column letter", () => assert.strictEqual(MD.dateCol, "D"));
test("mantraDocs", "a row with content but no Doc ID → an error; an empty row is skipped", () => { const r = H.mantraDocs([MH, ["", "A", "IN-WORK", ""], [], mrow("330-1-1-J1", "", "IN-WORK", "")], "R", "t"); assert.deepStrictEqual([r.rows.length, r.errors.length, r.rows[0][4]], [1, 1, "-"]); });
test("mantraDocs", "no header row → an error, nothing else", () => assert.deepStrictEqual(H.mantraDocs([["x"]], "R", "t").rows, []));

// ---------------------------------------------------------------- events
test("pairKey", "upper-case id and revision", () => assert.strictEqual(H.pairKey("330-1-1-j1", "a"), "330-1-1-J1|A"));
test("pairKey", "no revision → -", () => assert.strictEqual(H.pairKey("330-1-1-J1", ""), "330-1-1-J1|-"));
test("pairKey", "spaces trimmed", () => assert.strictEqual(H.pairKey(" 330-1-1-J1 ", " B "), "330-1-1-J1|B"));
test("pairKey", "null revision", () => assert.strictEqual(H.pairKey("330-1-1-J1", null), "330-1-1-J1|-"));
test("pairKey", "two revisions → two keys", () => assert.notStrictEqual(H.pairKey("X", "A"), H.pairKey("X", "B")));

test("linksOf", "a document links to its CDF", () => assert.deepStrictEqual(H.linksOf(SS)["330-11-11-J0001|A"], [{ key: "A330-1-001.-", sid: "K:A330-1-001.-|330-11-11-J0001|A" }]));
test("linksOf", "the same CDF twice counts once", () => assert.strictEqual(H.linksOf(SS)["330-33-33-J0003|B"].length, 1));
test("linksOf", "non-engineering rows have no links", () => assert.strictEqual(H.linksOf(SS)["-|-"], undefined));
test("linksOf", "one document in two CDFs → both", () => { const s = SS.slice(); const x = SS[0].slice(); x[4] = "A330-1-077.-"; s.push(x); assert.strictEqual(H.linksOf(s)["330-11-11-J0001|A"].length, 2); });
test("linksOf", "empty", () => assert.deepStrictEqual(H.linksOf([]), {}));

test("docStates", "one state per doc–revision", () => assert.strictEqual(Object.keys(H.docStates(MD.rows)).length, 5));
test("docStates", "the latest Date Status wins", () => { const r = H.mantraDocs([MH, mrow("X-1", "A", "RELEASED", "2026-11-05T07:00:00"), mrow("X-1", "A", "PRE-RELEASED", "2026-11-01T07:00:00")], "R", "t").rows; assert.strictEqual(H.docStates(r)["X-1|A"].status, "RELEASED"); });
test("docStates", "the same time → the later row", () => { const r = H.mantraDocs([MH, mrow("X-1", "A", "IN-WORK", "2026-11-05T07:00:00"), mrow("X-1", "A", "RELEASED", "2026-11-05T07:00:00")], "R", "t").rows; assert.strictEqual(H.docStates(r)["X-1|A"].row, "3"); });
test("docStates", "keeps the UTC time and row", () => assert.deepStrictEqual([H.docStates(MD.rows)["330-11-11-J0001|A"].utc, H.docStates(MD.rows)["330-11-11-J0001|A"].row], ["2026-11-05T05:30:00Z", "2"]));
test("docStates", "empty", () => assert.deepStrictEqual(H.docStates([]), {}));

test("detectEvents", "the first run has nothing to compare with", () => assert.deepStrictEqual(Y.events, []));
test("detectEvents", "CDF_RELEASED: Open → Released, the GCC cell as source", () => { const e = evOf(TD, "CDF_RELEASED"); assert.deepStrictEqual([e.length, e[0][4], e[0][8], e[0][9], e[0][16], e[0][14]], [1, "A330-1-001.-", "Open", "Released", "GCC!J3", "DETECTED AT"]); });
test("detectEvents", "DOCUMENT_RELEASED observed PRE-RELEASE in the baseline", () => { const e = evOf(TD, "DOCUMENT_RELEASED")[0]; assert.deepStrictEqual([e[6], e[7], e[12], e[13], e[14], e[16]], ["330-11-11-J0001", "A", "2026-11-05T05:30:00Z", "2026-11-05T05:30:00Z", "VERIFIED TRANSITION TIME", "Mantra!D2"]); assert.ok(/Observed PRE-RELEASED WF \(3 - PRE-RELEASE\)/.test(e[17])); });
test("detectEvents", "an unlinked document is history only", () => { const e = evOf(TD, "DOCUMENT_RELEASED")[1]; assert.deepStrictEqual([e[6], e[3], e[4]], ["330-99-99-J0009", "-", "-"]); assert.ok(/UNLINKED – history only/.test(e[17])); });
test("detectEvents", "GCC_ANSWER_RECEIVED: NO ENG. DOC. → ANSWER AT GCC", () => { const e = evOf(TD, "GCC_ANSWER_RECEIVED"); assert.deepStrictEqual([e.length, e[0][4], e[0][8], e[0][9], e[0][16]], [1, "A330-1-002.-", "NO ENG. DOC.", "ANSWER AT GCC", "GCC!S4"]); });
test("detectEvents", "REVISION_CHANGE A → B in the same solution", () => { const e = evOf(TD, "REVISION_CHANGE"); assert.deepStrictEqual([e.length, e[0][6], e[0][8], e[0][9]], [1, "330-22-22-J0002", "A", "B"]); });
test("detectEvents", "the order: CDF, documents (by Mantra row), GCC answers, revisions; ids E0001…", () => assert.deepStrictEqual(TD.events.map((e) => [e[1], e[2]]), [
  ["RUN-20261105060000-E0001", "CDF_RELEASED"], ["RUN-20261105060000-E0002", "DOCUMENT_RELEASED"], ["RUN-20261105060000-E0003", "DOCUMENT_RELEASED"], ["RUN-20261105060000-E0004", "GCC_ANSWER_RECEIVED"], ["RUN-20261105060000-E0005", "REVISION_CHANGE"]]));
test("detectEvents", "a document released once is never released again", () => {
  const again = run(new Date("2026-11-06T06:00:00Z"), CDF_T, CDF_T, MANTRA_T, [Y.run, TD.run], finals(TD), TD.events);
  assert.deepStrictEqual(evOf(again, "DOCUMENT_RELEASED"), []);
});
test("detectEvents", "RELEASED within 168 hours and not seen before → inferred", () => {
  const base = finals(Y); base.mantradocs = base.mantradocs.filter((r) => r[3] !== "330-99-99-J0009");
  const r = run(NOW_T, CDF_T, CDF_T, MANTRA_T, [Y.run], base, []);
  assert.ok(/inferred, not observed/.test(evOf(r, "DOCUMENT_RELEASED").filter((e) => e[6] === "330-99-99-J0009")[0][17]));
});
test("detectEvents", "RELEASED long ago and not seen as PRE-RELEASE → no event", () => assert.strictEqual(TD.events.filter((e) => e[6] === "330-33-33-J0003").length, 0));
test("detectEvents", "the old revision gone from Mantra and MANTRAREV → no REVISION_CHANGE (regression)", () => {
  const m = MANTRA_T.filter((r) => !(r[0] === "330-22-22-J0002" && r[1] === "A"));
  assert.deepStrictEqual(evOf(run(NOW_T, CDF_T, CDF_T, m, [Y.run], finals(Y), []), "REVISION_CHANGE"), []);
});

// ---------------------------------------------------------------- release tables
const RT = TD.release;
test("releaseTables", "CDFs released in 24 h", () => assert.deepStrictEqual(RT.cdf24.rows, [["A330-1-001.-", "Open", "Released"]]));
test("releaseTables", "titles carry AS OF in Israel time", () => assert.strictEqual(RT.cdf7.title, "CDFS RELEASED — LAST 7 DAYS (AS OF 05/11/2026 08:00 Asia/Jerusalem)"));
test("releaseTables", "documents: linked ones and GCC answers, latest first", () => assert.deepStrictEqual(RT.doc24.rows, [
  ["ANSWER AT GCC", "-", "A330-1-002.-", "NO ENG. DOC.", "ANSWER AT GCC", "05/11/2026 08:00"],
  ["330-11-11-J0001", "A", "A330-1-001.-", "3 - PRE-RELEASE", "4 - RELEASED", "05/11/2026 07:30"]]));
test("releaseTables", "an event older than 24 h is in the 7-day table only", () => {
  const ev = TD.events.map((e) => e.slice()); ev[0][13] = "2026-11-02T06:00:00Z";
  const r = H.releaseTables(ev, TD.solutions.filter((x) => x[1] === "FINAL"), CAT_T, M, NOW_T.getTime());
  assert.deepStrictEqual([r.cdf24.rows.length, r.cdf7.rows.length], [0, 1]);
});
test("releaseTables", "older than 7 days → in neither", () => {
  const ev = TD.events.map((e) => e.slice()); ev[0][13] = "2026-10-20T06:00:00Z";
  assert.strictEqual(H.releaseTables(ev, TD.solutions, CAT_T, M, NOW_T.getTime()).cdf7.rows.length, 0);
});
test("releaseTables", "a CDF released twice in the window shows once (the latest)", () => {
  const e2 = TD.events[0].slice(); e2[13] = "2026-11-04T06:00:00Z";
  assert.strictEqual(H.releaseTables([TD.events[0], e2], TD.solutions, CAT_T, M, NOW_T.getTime()).cdf7.rows.length, 1);
});
test("releaseTables", "a document whose record is no longer RELEASED is left out", () => {
  const m = MANTRA_T.map((r) => r.slice()); m[1][2] = "IN-WORK";
  assert.strictEqual(H.releaseTables(TD.events, TD.solutions, M.readMantra(m), M, NOW_T.getTime()).doc24.rows.filter((r) => r[0] === "330-11-11-J0001").length, 0);
});
test("releaseTables", "a document released twice in history is left out", () => {
  const ev = TD.events.concat([TD.events[1].slice()]);
  assert.strictEqual(H.releaseTables(ev, TD.solutions, CAT_T, M, NOW_T.getTime()).doc7.rows.filter((r) => r[0] === "330-11-11-J0001").length, 0);
});
test("releaseTables", "no events → empty tables with their headers", () => { const r = H.releaseTables([], [], CAT_T, M, NOW_T.getTime()); assert.deepStrictEqual([r.doc7.rows, r.doc7.headers.length, r.cdf24.headers.length], [[], 6, 3]); });

// ---------------------------------------------------------------- RUNS row, sheet ERROR
test("runRow", "no errors → COMMITTED / PASSED", () => assert.deepStrictEqual([Y.run[4], Y.run[14], Y.run[15]], ["COMMITTED", "0", "PASSED"]));
test("runRow", "the counts, as text", () => assert.deepStrictEqual(TD.run.slice(9, 14), ["4", "4", "5", "5", "5"]));
test("runRow", "the previous committed run and when tracking started", () => assert.deepStrictEqual([TD.run[5], TD.run[6]], [Y.runId, Y.at]));
test("runRow", "the first run tracks from itself", () => assert.deepStrictEqual([Y.run[5], Y.run[6]], ["", Y.at]));
test("runRow", "errors → FAILED / FAILED with the count", () => assert.deepStrictEqual(H.runRow({ runId: "R", at: "t", errors: 3, startCdfs: 1, finalCdfs: 1, startSolutions: 1, finalSolutions: 1, events: 0 }).slice(14), ["3", "FAILED"]));
test("runRow", "time zone and schema", () => assert.deepStrictEqual(Y.run.slice(7, 9), ["Asia/Jerusalem", "BACKUPDATA-v3"]));

const ER = [{ severity: "ERROR", stage: "25/30", sheet: "CDF", cdf: "A330-1-002.-", code: "AMBIGUOUS", details: "two documents" },
  { severity: "BLOCKER", stage: "00.1", sheet: "CDF", cell: "C1", code: "HEADER_MISMATCH", details: "C1" },
  { severity: "ERROR", stage: "10", sheet: "CDF", cdf: "A330-1-001.-", code: "IMPORT_ERROR", details: "x" },
  { severity: "ERROR", stage: "25/30", sheet: "CDF", cdf: "A330-1-002.-", code: "AMBIGUOUS", details: "two documents again" }];
const ERR = H.errorRows(ER, "RUN-9", "2026-11-05T06:00:00Z");
test("errorRows", "the same error twice → once", () => assert.strictEqual(ERR.length, 3));
test("errorRows", "BLOCKER first", () => assert.strictEqual(ERR[0][14], "BLOCKER"));
test("errorRows", "then by stage order", () => assert.deepStrictEqual(ERR.map((r) => r[2]), ["00.1", "10", "25/30"]));
test("errorRows", "15 columns, UNRESOLVED", () => assert.deepStrictEqual([ERR[1].length, ERR[1][13]], [15, "UNRESOLVED"]));
test("errorRows", "run id and detection time on each", () => assert.deepStrictEqual(ERR[2].slice(0, 2), ["RUN-9", "2026-11-05T06:00:00Z"]));
test("errorRows", "none → none", () => assert.deepStrictEqual(H.errorRows([], "R", "t"), []));

test("errorLayout", "no errors: headers, then the info block from row 4 and NO ERRORS FOUND", () => { const l = H.errorLayout([], "R", "05/11/2026 08:00"); assert.deepStrictEqual([l.records, l.info, l.end, l.cells.filter((c) => c.r === 7)[0].v], [null, { r1: 4, r2: 7 }, 7, "NO ERRORS FOUND"]); });
test("errorLayout", "with errors: rows from 3, a blank row, then the info block", () => { const l = H.errorLayout(ERR, "R", "x"); assert.deepStrictEqual([l.records, l.info, l.end], [{ r1: 3, r2: 5 }, { r1: 7, r2: 9 }, 9]); });
test("errorLayout", "title merged over A1:O1", () => { const l = H.errorLayout([], "R", "x"); assert.deepStrictEqual([l.cells[0].v, l.merges], ["ALGORITHM RUN ERRORS", [{ r1: 1, c1: 0, r2: 1, c2: 14 }]]); });
test("errorLayout", "the 15 headers in row 2", () => assert.deepStrictEqual(H.errorLayout([], "R", "x").cells.filter((c) => c.r === 2).map((c) => c.v), H.ERROR_HEADERS));
test("errorLayout", "ERROR COUNT is a number", () => { const c = H.errorLayout(ERR, "R", "x").cells.filter((x) => x.role === "num")[0]; assert.deepStrictEqual([c.v, c.r], [3, 9]); });

test("collectErrors", "a stage 1 header error is a BLOCKER", () => assert.deepStrictEqual(H.collectErrors({ stage1Errors: [{ stage: "00.1", message: "C1 is wrong" }] }).map((e) => [e.severity, e.code]), [["BLOCKER", "HEADER_MISMATCH"]]));
test("collectErrors", "a failed control total", () => assert.strictEqual(H.collectErrors({ checks: [{ ok: true }, { ok: false, name: "CDF STATUS", detail: "6 ≠ 7" }] })[0].code, "CONTROL_TOTAL_FAILED"));
test("collectErrors", "stage 2: AMBIGUOUS keeps its code and document", () => { const e = H.collectErrors({ stage2Errors: [{ stage: "25/30", key: "A330-1-002.-", message: "330-11-11-J0001: AMBIGUOUS — two documents" }] })[0]; assert.deepStrictEqual([e.code, e.doc, e.cdf], ["AMBIGUOUS", "330-11-11-J0001", "A330-1-002.-"]); });
test("collectErrors", "a CDF SUMMARY that reads back other numbers", () => { const e = H.collectErrors({ readBack: [{ cell: "B5", want: 6, got: 7 }] })[0]; assert.deepStrictEqual([e.code, e.expected, e.actual], ["SUMMARY_TOTAL_MISMATCH", "6", "7"]); });
test("collectErrors", "BACKUPDATA and Mantra row problems", () => assert.deepStrictEqual(H.collectErrors({ historyErrors: [{ message: "x" }], mantraRowErrors: ["Sheet Mantra row 12 has content but no Doc ID"] }).map((e) => [e.sheet, e.cell]), [["BACKUPDATA", ""], ["Mantra", "row 12"]]));
test("collectErrors", "a Mantra row stage 2 already stopped on is reported once, as its BLOCKER", () => { const m = "Sheet Mantra row 9 has content but no Doc ID"; const e = H.collectErrors({ stage2Errors: [{ stage: "00.1.2", message: m }], mantraRowErrors: [m] }); assert.deepStrictEqual(e.map((x) => [x.severity, x.code]), [["BLOCKER", "MANTRA_UNREADABLE"]]); });
test("collectErrors", "nothing failed → nothing", () => assert.deepStrictEqual(H.collectErrors({ checks: [{ ok: true }] }), []));

// ---------------------------------------------------------------- one run, end to end
test("buildRun", "START and FINAL snapshots of CDFs and solutions", () => assert.deepStrictEqual([TD.cdfs.length, TD.solutions.length, TD.counts.cdfs, TD.counts.solutions], [8, 10, 8, 10]));
test("buildRun", "the run id and time from now", () => assert.deepStrictEqual([TD.runId, TD.at], ["RUN-20261105060000", "2026-11-05T06:00:00Z"]));
test("buildRun", "the baseline is the last committed run", () => assert.strictEqual(TD.base.runId, Y.runId));
test("buildRun", "MANTRADOCS: every Mantra row", () => assert.strictEqual(TD.mantradocs.length, 5));
test("buildRun", "an error fails the run, and its own events don't reach the release tables", () => {
  const r = run(NOW_T, CDF_T, CDF_T, MANTRA_T, [Y.run], finals(Y), [], { checks: [{ ok: false, name: "X", detail: "y" }] });
  assert.deepStrictEqual([r.commits, r.run[4], r.errors.length, r.events.length, r.release.cdf24.rows.length], [false, "FAILED", 1, 5, 0]);
});
test("buildRun", "events of a FAILED run in history don't count", () => {
  const bad = TD.run.slice(); bad[4] = "FAILED";
  const r = run(new Date("2026-11-05T07:00:00Z"), CDF_T, CDF_T, MANTRA_T, [Y.run, bad], finals(Y), TD.events);
  assert.strictEqual(evOf(r, "DOCUMENT_RELEASED").length, 2);           // detected again: the failed run's releases never happened
});
test("buildRun", "the baseline only from a run with its snapshot (hasSnapshot)", () => {
  const r = H.buildRun({ now: NOW_T, started: NOW_T, algo: A, M, runs: [Y.run], base: null, history: [], startCdf: CDF_T, finalCdf: CDF_T, mantra: MANTRA_T, mantraRev: [], gcc: GCC, errors: {}, hasSnapshot: () => false });
  assert.deepStrictEqual([r.base, r.events.length, r.run[5]], [null, 0, ""]);
});
test("buildRun", "the errors given aren't changed (no Object.assign: IE11)", () => { const e = { checks: [] }; run(NOW_T, CDF_T, CDF_T, MANTRA_T, [Y.run], finals(Y), [], e); assert.deepStrictEqual(Object.keys(e), ["checks"]); });
test("buildRun", "a Mantra row without a Doc ID fails the run", () => {
  const m = MANTRA_T.concat([["", "A", "IN-WORK", "", 0, "", "", "", ""]]);
  assert.strictEqual(run(NOW_T, CDF_T, CDF_T, m, [Y.run], finals(Y), []).run[4], "FAILED");
});

// ---------------------------------------------------------------- undo of a run (features.js)
const LOGH = ["When", "By", "Sheet", "Row", "CDF", "Column", "Before", "After", "Note"];
const st = "05/11/2026 08:00:00", nm = "CDF algorithm 2.0 · stage 4";
const LOG = [LOGH, [st, nm, "CDF", "2", "A330-1-001.-", "R", "Open", "Released", ""], [st, nm, "BACKUPDATA", "3", "RUN-20261105060000", "RUN", "", "COMMITTED", "Stage 4: …"], [st, nm, "ERROR", "", "", "", "", "", "Rebuilt"]];
test("undoPlan", "the RUN line becomes a run to take back", () => assert.deepStrictEqual(F.undoPlan(LOG).runOps, [{ sheet: "BACKUPDATA", row: 3, runId: "RUN-20261105060000" }]));
test("undoPlan", "the RUN line is not a cell to put back", () => assert.deepStrictEqual(F.undoPlan(LOG).entries.map((e) => e.col), ["R"]));
test("undoPlan", "a run that ended FAILED has nothing to take back", () => assert.deepStrictEqual(F.undoPlan(LOG.concat([[st, nm, "BACKUPDATA", "3", "RUN-20261105060000", "RUN", "COMMITTED", "FAILED", "read back"]])).runOps, []));
test("undoPlan", "a run with only a RUN line can be undone", () => assert.strictEqual(F.undoPlan([LOGH, LOG[2]]).runOps.length, 1));
test("undoPlan", "an undone update isn't offered again", () => assert.strictEqual(F.undoPlan(LOG.concat([[st, F.UNDO, "BACKUPDATA", "3", "RUN-20261105060000", "RUN", "COMMITTED", "FAILED", "Undo of " + st + " " + nm]])), null));

const RO = [{ sheet: "BACKUPDATA", row: 3, runId: "RUN-1" }];
test("undoRunCheck", "the row still holds the run as COMMITTED → ready", () => assert.strictEqual(F.undoRunCheck(RO, [["RUN-1", "COMMITTED"]]).ready.length, 1));
test("undoRunCheck", "another run in that row → left", () => assert.strictEqual(F.undoRunCheck(RO, [["RUN-2", "COMMITTED"]]).conflicts.length, 1));
test("undoRunCheck", "already FAILED → left", () => assert.strictEqual(F.undoRunCheck(RO, [["RUN-1", "FAILED"]]).conflicts.length, 1));
test("undoRunCheck", "no BACKUPDATA (no probe) → left", () => assert.strictEqual(F.undoRunCheck(RO, [[]]).conflicts.length, 1));
test("undoRunCheck", "spaces around are fine", () => assert.strictEqual(F.undoRunCheck(RO, [[" RUN-1 ", " COMMITTED "]]).ready.length, 1));

test("undoRunLogRows", "one log row per run", () => assert.strictEqual(F.undoRunLogRows({ id: "x" }, RO, "s").length, 1));
test("undoRunLogRows", "COMMITTED → FAILED in column RUN", () => assert.deepStrictEqual(F.undoRunLogRows({ id: "x" }, RO, "s")[0].slice(2, 8), ["BACKUPDATA", "3", "RUN-1", "RUN", "COMMITTED", "FAILED"]));
test("undoRunLogRows", "Undo of the update id", () => assert.strictEqual(F.undoRunLogRows({ id: "x y" }, RO, "s")[0][8], "Undo of x y"));
test("undoRunLogRows", "by the undo name and stamp", () => assert.deepStrictEqual(F.undoRunLogRows({ id: "x" }, RO, "S1")[0].slice(0, 2), ["S1", F.UNDO]));
test("undoRunLogRows", "none → none", () => assert.deepStrictEqual(F.undoRunLogRows({ id: "x" }, [], "s"), []));

const base0 = { errors: [], checks: [], cdfChanges: [], gccChanges: [] };
test("morningDecision", "new events count as work, without stopping", () => { const d = F.morningDecision(Object.assign({ history: { ran: true, out: { events: [1], commits: true, errors: [] } } }, base0)); assert.deepStrictEqual([d.stop, d.history], [false, true]); });
test("morningDecision", "no events → no history work", () => assert.strictEqual(F.morningDecision(Object.assign({ history: { ran: true, out: { events: [], commits: true, errors: [] } } }, base0)).history, false));
test("morningDecision", "a run that would fail stops the morning", () => { const d = F.morningDecision(Object.assign({ history: { ran: true, out: { events: [], commits: false, errors: [1, 2] } } }, base0)); assert.ok(d.stop && /FAILED — 2 errors/.test(d.reasons[0].text)); });
test("morningDecision", "a BACKUPDATA with wrong headers stops it too", () => assert.ok(F.morningDecision(Object.assign({ history: { skipped: "headers", bad: true } }, base0)).stop));
test("morningDecision", "no stage 4 at all", () => assert.deepStrictEqual([F.morningDecision(base0).stop, F.morningDecision(base0).history], [false, false]));

const low = Object.entries(perFn).filter(([, n]) => n < 5);
if (low.length) process.exitCode = 1;
console.log(`\n${passed} tests passed, ${failed} failed` + (low.length ? " — fewer than 5 cases: " + low.map(([f, n]) => f + " (" + n + ")").join(", ") : " — every function has 5+ cases"));
