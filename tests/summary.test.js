// Run: node tests/summary.test.js   (synthetic data only — no real names or documents)
// Stage 3 of the CDF algorithm: CDF SUMMARY (all tables) and EMAIL DRAFT (cdf-summary.js). At least 5 cases for every function.
const assert = require("assert");
const S = require("../cdf-summary.js");
const A = require("../cdf-algo.js");
const M = require("../cdf-mantra.js");
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
function row(key, R, x) { const r = new Array(34).fill(""); r[C("C")] = key; r[C("R")] = R; Object.keys(x || {}).forEach((k) => { r[C(k)] = x[k]; }); return r; }
function cont(R, x) { return row("", R, x); }
const CDF = [new Array(34).fill("h"),
  row("A330-1-010.-", "Open", { S: "3 - Conversion Engineer", T: "-", P: "ANN", U: "3 - PRE-RELEASE", V: "IAI - Design Engineering", W: "APP Design", X: "DANA LEVI", AA: "330-53-14-J2400", AB: "B" }),
  cont("Open", { U: "2 - IN WORK", V: "IAI - Stress Department", W: "Rejected | Stress", X: "RON AVI", AA: "330-53-15-J2654", AB: "C" }),
  cont("Open", { U: "4 - RELEASED", V: "-", W: "-", X: "-", AA: "330-53-14-J2401", AB: "-" }),
  row("A330-1-002.A", "Open", { S: "9 - Final Approval", T: "Tom Bar", U: "3 - PRE-RELEASE", V: "-", W: "Rejected", X: "-", AA: "330-92-80-JF133-R01", AB: "A" }),
  cont("Open", { U: "3 - PRE-RELEASE", V: "IAI - Design Engineering", W: "APP Design", X: "DANA LEVI", AA: "330-53-14-J2400", AB: "B" }),
  row("A330-1-003.-", "Released", { S: "9 - Final Approval", U: "2 - IN WORK", V: "IAI - Design Engineering", W: "APP Design", X: "DANA LEVI", AA: "330-11-11-J0001", AB: "A" }),
  row("A330-1-004.-", "Cancelled", {}),
  row("A330-1-100.-", "Open", { S: "3 - Conversion Engineer", T: "-", P: "TIMAT", U: "ANSWER AT GCC", AA: "ANSWER AT GCC", AB: "-" }),
  row("A330-1-011.-", "Open", { S: "1 - Initiator", T: "Ann Vendor", U: "3 - PRE-RELEASE", V: "ANN/SIE", W: "Check Design (SIE)", X: "ANN VENDOR", AA: "330-53-13-J1565", AB: "D" })
];
const g = (k, R, S2, T2) => { const r = new Array(40).fill(""); r[0] = k; r[C("J")] = R; return r; };
const CFG = { personGroups: [{ group: "IAI - Design Engineering", people: ["Dana Levi"] }, { group: "IAI - Stress Department", people: ["Ron Avi"] }, { group: "ANN/SIE", people: ["Ann Vendor"] }], personAliases: [] };
const NAMES = S.makeNames(CFG, M);
const REL = S.relations(CDF, A);
// the stage 1 summary of this sheet, as stage 1 hands it over
const S1 = { tables: [
  { name: "CDF STATUS", headers: ["CDF STATUS", "CDF'S"], rows: [["Open", 4], ["Cancelled", 1], ["Released", 1], ["TOTAL", 6]] },
  { name: "GCC CURRENT STATUS", headers: ["GCC CURRENT STATUS", "CDF'S"], rows: [["1 - Initiator", 1], ["2 - POC", 0], ["3 - Conversion Engineer", 2], ["4 - AVIONICS", 0], ["9 - Final Approval", 1], ["TOTAL", 4]] },
  { name: "GCC SIGNATURE POSITION", headers: ["SIGNATURE POSITION", "CDF'S"], rows: [["Ann Vendor", 1], ["Tom Bar", 1], ["TOTAL", 2]] },
  { name: "WHERE ARE THE CDFS", headers: ["SUPPLIER", "CDF'S"], rows: [["ANN/SIE", 1], ["ANCARA", 0], ["TIMAT", 1], ["IAI", 0], ["TOTAL", 2]] },
  { name: "MOR PENDING TABLE", headers: ["CDF NUMBER", "SOLUTION / DOCUMENT"], rows: [["A330-1-002.A", "ANSWER AT GCC"], ["TOTAL", 1]] },
  { name: "IGOR PENDING TABLE", headers: ["CDF NUMBER", "SOLUTION / DOCUMENT"], rows: [["TOTAL", 0]] }
], pendingPeople: [{ table: "MOR PENDING TABLE", person: "X" }, { table: "IGOR PENDING TABLE", person: "Y" }] };
// the release tables as the sheet has them now
const CURRENT = [["CDFS RELEASED — LAST 24 HOURS (AS OF 05/11/2026 08:10 Asia/Jerusalem)"], ["CDF NUMBER", "PREVIOUS STATUS", "NEW STATUS"], ["A330-1-003.-", "Open", "Released"], ["TOTAL", "", 1], [], [],
  ["CDFS RELEASED — LAST 7 DAYS (AS OF 05/11/2026 08:10 Asia/Jerusalem)"], ["CDF NUMBER", "PREVIOUS STATUS", "NEW STATUS"], ["A330-1-003.-", "Open", "Released"], ["A330-1-009.-", "Open", "Released"], ["TOTAL", "", 2], [], [],
  ["DOCUMENTS RELEASED / GCC ANSWERS — LAST 24 HOURS (AS OF 05/11/2026 08:10 Asia/Jerusalem)"], ["DOCUMENT / SOLUTION", "REVISION", "CDF NUMBERS", "PREVIOUS STATUS", "NEW STATUS", "TRANSITION / RELEASE TIME"],
  ["330-11-11-J0001", "A", "A330-1-003.-", "3 - PRE-RELEASE", "4 - RELEASED", "05/11/2026 06:47"], ["ANSWER AT GCC", "-", "A330-1-100.-", "NO ENG. DOC.", "ANSWER AT GCC", "04/11/2026 08:52"],
  ["ENGINEERING DOCUMENTS", "", "", "", "", 1], ["GCC ANSWERS", "", "", "", "", 1], ["OF WHICH LINKED TO CDFS", "", "", "", "", 1], ["TOTAL", "", "", "", "", 2], [], [],
  ["DOCUMENTS RELEASED / GCC ANSWERS — LAST 7 DAYS (AS OF 05/11/2026 08:10 Asia/Jerusalem)"], ["DOCUMENT / SOLUTION", "REVISION", "CDF NUMBERS", "PREVIOUS STATUS", "NEW STATUS", "TRANSITION / RELEASE TIME"],
  ["ENGINEERING DOCUMENTS", "", "", "", "", 0]];
const MODEL = S.build({ cdf: CDF, summary: S1, cfg: CFG, algo: A, mantra: M, current: CURRENT });
const ALIASES = S.supplierAliases(A.settings({}));
const LAY = S.layoutSummary(MODEL, { aliases: ALIASES });
const MAIL = S.mailParts(new Date("2026-11-05T06:10:00Z"), ["Your Name"]);
const LE = S.layoutEmail(MODEL, MAIL);
const tbl = (id) => MODEL.tables.filter((t) => t.id === id)[0];
const placed = (lay, id) => lay.tables.filter((t) => t.id === id)[0];
const at = (lay, a) => lay.cells.filter((c) => S.letter(c.c) + c.r === a)[0];

// ---------------------------------------------------------------- keys and names
test("keyParts", "number and revision", () => assert.deepStrictEqual(S.keyParts("A330-1-740.A").slice(0, 2), [740, 1]));
test("keyParts", "- is the first revision", () => assert.strictEqual(S.keyParts("A330-1-740.-")[1], 0));
test("keyParts", "two-letter revision after Z", () => assert.ok(S.keyParts("A330-1-1.AA")[1] > S.keyParts("A330-1-1.Z")[1]));
test("keyParts", "a note after the key", () => assert.strictEqual(S.keyParts("A330-1-052 (another)")[0], 52));
test("keyParts", "not a key → at the end", () => assert.strictEqual(S.keyParts("hello")[0], Infinity));

test("keyCmp", "by number, not text", () => assert.ok(S.keyCmp("A330-1-99.-", "A330-1-100.-") < 0));
test("keyCmp", "revision within a number", () => assert.ok(S.keyCmp("A330-1-740.-", "A330-1-740.A") < 0));
test("keyCmp", "equal keys", () => assert.strictEqual(S.keyCmp("A330-1-740.A", "a330-1-740.a"), 0));
test("keyCmp", "1001 after 999", () => assert.ok(S.keyCmp("A330-1-1001.-", "A330-1-999.-") > 0));
test("keyCmp", "sorts a list", () => assert.deepStrictEqual(["A330-1-10.-", "A330-1-2.A", "A330-1-2.-"].sort(S.keyCmp), ["A330-1-2.-", "A330-1-2.A", "A330-1-10.-"]));

test("personKey", "case and spaces", () => assert.strictEqual(S.personKey("  dana   levi "), "DANA LEVI"));
test("personKey", "no-break space", () => assert.strictEqual(S.personKey("Dana Levi"), "DANA LEVI"));
test("personKey", "hyphen", () => assert.strictEqual(S.personKey("Ben - Ari"), "BEN-ARI"));
test("personKey", "en dash", () => assert.strictEqual(S.personKey("Ben–Ari"), "BEN-ARI"));
test("personKey", "empty", () => assert.strictEqual(S.personKey(null), ""));

test("makeNames", "the name as in the people list", () => assert.strictEqual(NAMES.canon("DANA LEVI"), "Dana Levi"));
test("makeNames", "words the other way round", () => assert.strictEqual(NAMES.canon("LEVI DANA"), "Dana Levi"));
test("makeNames", "an alias", () => assert.strictEqual(S.makeNames({ personGroups: CFG.personGroups, personAliases: [{ from: "D. Levi", to: "Dana Levi" }] }, M).canon("d. levi"), "Dana Levi"));
test("makeNames", "unknown in capitals → as a name", () => assert.strictEqual(NAMES.canon("NEW PERSON"), "New Person"));
test("titleCase", "capitals → a name", () => assert.strictEqual(S.titleCase("DANA LEVI"), "Dana Levi"));
test("titleCase", "after a hyphen too", () => assert.strictEqual(S.titleCase("DANA BEN-ARI"), "Dana Ben-Ari"));
test("titleCase", "already mixed case stays", () => assert.strictEqual(S.titleCase("Dana van Levi"), "Dana van Levi"));
test("titleCase", "extra spaces", () => assert.strictEqual(S.titleCase("  ANN   VENDOR "), "Ann Vendor"));
test("titleCase", "no letters", () => assert.strictEqual(S.titleCase("-"), "-"));
test("makeNames", "group from the list; without one, from V; else dash", () => {
  assert.strictEqual(NAMES.group("Dana Levi", "x"), "IAI - Design Engineering");
  const none = S.makeNames({}, M);
  assert.strictEqual(none.group("Dana Levi", "ANN/SIE"), "ANN/SIE");
  assert.strictEqual(none.group("Dana Levi", "A, B"), "-");
});

test("mainGroup", "an IAI department → IAI", () => assert.strictEqual(S.mainGroup("IAI - Stress Department"), "IAI"));
test("mainGroup", "IAI alone", () => assert.strictEqual(S.mainGroup("IAI"), "IAI"));
test("mainGroup", "ANN/SIE", () => assert.strictEqual(S.mainGroup("ANN/SIE"), "ANN/SIE"));
test("mainGroup", "ANCRA stays apart from ANN/SIE", () => assert.strictEqual(S.mainGroup("ANCRA"), "ANCRA"));
test("mainGroup", "a word that starts with IAI letters isn't IAI", () => assert.strictEqual(S.mainGroup("IAIX"), "IAIX"));

// ---------------------------------------------------------------- 100.2.1 relations
test("relations", "Open CDFs only (Released and Cancelled left out)", () => assert.deepStrictEqual([...new Set(REL.map((x) => x.key))], ["A330-1-010.-", "A330-1-002.A", "A330-1-011.-"]));
test("relations", "released documents and GCC answers aren't pending", () => assert.ok(!REL.some((x) => x.AA === "330-53-14-J2401" || x.AA === "ANSWER AT GCC")));
test("relations", "rejected: one position, Rejected; the function kept apart", () => { const r = REL.filter((x) => x.AA === "330-53-15-J2654")[0]; assert.deepStrictEqual([r.rejected, r.fns, r.rejFn, r.persons], [true, ["Rejected"], "Stress", ["RON AVI"]]); });
test("relations", "rejecter missing in Mantra → no person", () => { const r = REL.filter((x) => x.AA === "330-92-80-JF133-R01")[0]; assert.deepStrictEqual([r.rejected, r.persons, r.rejFn], [true, [], ""]); });
test("relations", "several pending functions", () => { const cdf = [CDF[0], row("A330-1-020.-", "Open", { U: "3 - PRE-RELEASE", W: "Check Design, Production ENG", X: "A B, C D", AA: "330-1-1-J1", AB: "A" })]; assert.deepStrictEqual(S.relations(cdf, A)[0].fns, ["Check Design", "Production ENG"]); });

// ---------------------------------------------------------------- the Mantra tables
const SIG = S.mantraSignature(REL);
test("mantraSignature", "document–revision pairs per position", () => assert.deepStrictEqual(SIG, [["Rejected", 2], ["APP Design", 1], ["Check Design (SIE)", 1]]));
test("mantraSignature", "the same document in two CDFs counts once", () => assert.strictEqual(SIG.filter((x) => x[0] === "APP Design")[0][1], 1));
test("mantraSignature", "sorted by count, then by name", () => { const r = S.mantraSignature([{ fns: ["B"], AA: "1", AB: "-" }, { fns: ["A"], AA: "2", AB: "-" }, { fns: ["C"], AA: "3", AB: "-" }, { fns: ["C"], AA: "4", AB: "-" }]); assert.deepStrictEqual(r.map((x) => x[0]), ["C", "A", "B"]); });
test("mantraSignature", "a document pending at two positions counts at both", () => assert.deepStrictEqual(S.mantraSignature([{ fns: ["A", "B"], AA: "1", AB: "-" }]), [["A", 1], ["B", 1]]));
test("mantraSignature", "nothing pending → no rows", () => assert.deepStrictEqual(S.mantraSignature([]), []));

const WHO = S.mantraWho(REL, NAMES);
test("mantraWho", "people by the people list's spelling", () => assert.deepStrictEqual(WHO.map((x) => x[0]), ["Ann Vendor", "Dana Levi", "Ron Avi"]));
test("mantraWho", "one document in two CDFs is one", () => assert.strictEqual(WHO.filter((x) => x[0] === "Dana Levi")[0][1], 1));
test("mantraWho", "no row for a missing rejecter", () => assert.ok(!WHO.some((x) => x[0] === "-")));
test("mantraWho", "count first", () => assert.deepStrictEqual(S.mantraWho([{ persons: ["B"], AA: "1", AB: "-" }, { persons: ["A"], AA: "2", AB: "-" }, { persons: ["B"], AA: "3", AB: "-" }], NAMES), [["B", 2], ["A", 1]]));
test("mantraWho", "empty", () => assert.deepStrictEqual(S.mantraWho([], NAMES), []));

const GRP = S.personGroups(REL, NAMES);
test("personGroups", "all the groups in order, zeros included", () => assert.deepStrictEqual(GRP.map((x) => x[0]), S.GROUPS));
test("personGroups", "person–CDF pairs", () => assert.strictEqual(GRP.filter((x) => x[0] === "IAI - Design Engineering")[0][1], 2));
test("personGroups", "confirmed groups are counted and missing groups are omitted", () => {
  const r = S.personGroups([{ persons: ["One Person"], key: "K", V: "Other confirmed group" }, { persons: ["Nobody"], key: "J", V: "UNMAPPED" }], NAMES);
  assert.deepStrictEqual(r.slice(S.GROUPS.length), [["Other confirmed group", 1]]);
});
test("personGroups", "ANN/SIE", () => assert.strictEqual(GRP.filter((x) => x[0] === "ANN/SIE")[0][1], 1));
test("personGroups", "the total adds up to the pairs", () => assert.strictEqual(GRP.reduce((s, x) => s + x[1], 0), 4));

const REJ = S.rejectedRows(REL);
test("rejectedRows", "one row per CDF + document + revision, sorted by CDF", () => assert.deepStrictEqual(REJ.map((r) => r[0] + " " + r[2]), ["A330-1-002.A 330-92-80-JF133-R01", "A330-1-010.- 330-53-15-J2654"]));
test("rejectedRows", "the rejecter as Mantra wrote it, and the function", () => assert.deepStrictEqual(REJ[1].slice(4), ["RON AVI", "Stress"]));
test("rejectedRows", "missing rejecter → - and -", () => assert.deepStrictEqual(REJ[0].slice(4), ["-", "-"]));
test("rejectedRows", "status is Open", () => assert.ok(REJ.every((r) => r[1] === "Open")));
test("rejectedRows", "the same document twice in one CDF is one row", () => assert.strictEqual(S.rejectedRows([{ rejected: true, key: "A330-1-1.-", AA: "D", AB: "A", persons: [], rejFn: "" }, { rejected: true, key: "A330-1-1.-", AA: "D", AB: "A", persons: [], rejFn: "" }]).length, 1));

const PR = S.personRows(REL, NAMES);
test("personRows", "group, then person, then position, then CDF", () => assert.deepStrictEqual(PR.map((p) => p.group + "/" + p.person + "/" + p.pos + "/" + p.key), ["ANN/SIE/Ann Vendor/Check Design (SIE)/A330-1-011.-", "IAI/Dana Levi/APP Design/A330-1-002.A", "IAI/Dana Levi/APP Design/A330-1-010.-", "IAI/Ron Avi/Rejected/A330-1-010.-"]));
test("personRows", "# of M: the person's document–revision pairs, once", () => assert.strictEqual(PR.filter((p) => p.person === "Dana Levi")[0].nM, 1));
test("personRows", "documents of a row in F", () => assert.deepStrictEqual(PR[3].docs, ["330-53-15-J2654 Rev C"]));
test("personRows", "MAIN GROUP: departments as IAI", () => assert.ok(PR.filter((p) => p.person === "Ron Avi")[0].group === "IAI"));
test("personRows", "two documents of one person, position and CDF share the row", () => { const r = S.personRows([{ persons: ["P"], fns: ["F"], key: "K", AA: "D1", AB: "-" }, { persons: ["P"], fns: ["F"], key: "K", AA: "D2", AB: "-" }], NAMES); assert.deepStrictEqual([r.length, r[0].docs, r[0].nM], [1, ["D1 Rev -", "D2 Rev -"], 2]); });
test("personRows", "several people and functions can't be paired → each carries all", () => { const r = S.personRows([{ persons: ["A", "B"], fns: ["F1", "F2"], key: "K", AA: "D", AB: "-", V: "" }], NAMES); assert.deepStrictEqual(r.map((x) => x.pos), ["F1, F2", "F1, F2"]); });

// ---------------------------------------------------------------- the release tables
const RT = S.readReleaseTables(CURRENT);
test("readReleaseTables", "all four found", () => assert.deepStrictEqual(Object.keys(RT).sort(), ["cdf24", "cdf7", "doc24", "doc7"]));
test("readReleaseTables", "the title keeps its AS OF", () => assert.ok(/AS OF 05\/11\/2026 08:10/.test(RT.cdf24.title)));
test("readReleaseTables", "rows stop at TOTAL", () => assert.strictEqual(RT.cdf7.rows.length, 2));
test("readReleaseTables", "document tables stop at ENGINEERING DOCUMENTS", () => assert.deepStrictEqual(RT.doc24.rows.map((r) => r[0]), ["330-11-11-J0001", "ANSWER AT GCC"]));
test("readReleaseTables", "an empty table", () => assert.deepStrictEqual(RT.doc7.rows, []));
test("readReleaseTables", "not in the sheet → missing", () => assert.deepStrictEqual(S.readReleaseTables([["x"]]), {}));

// ---------------------------------------------------------------- build
test("build", "every table in sheet order", () => assert.deepStrictEqual(MODEL.tables.map((t) => t.id), ["status", "gcc", "sig", "pending0", "pending1", "where", "msig", "who", "groups", "rejected", "persons", "cdf24", "cdf7", "doc24", "doc7"]));
test("build", "stage 1's tables come without their TOTAL (it's written as a formula)", () => assert.ok(!tbl("status").rows.some((r) => r[0] === "TOTAL")));
test("build", "the release tables carried over", () => assert.deepStrictEqual([tbl("cdf24").carried, tbl("cdf24").rows.length], [true, 1]));
test("build", "missing release tables are named", () => assert.deepStrictEqual(S.build({ cdf: CDF, summary: S1, cfg: CFG, algo: A, mantra: M, current: [] }).missing.length, 4));
test("build", "rejected sums", () => assert.deepStrictEqual(tbl("rejected").sums, [["REJECTED CDFS", 2], ["REJECTED DOCUMENTS", 2]]));
test("build", "people without a confirmed group are named", () => {
  const cdf = CDF.map((r) => r.slice());
  cdf[1][C("V")] = "UNMAPPED";
  cdf[1][C("X")] = "Unknown Person";
  const m = S.build({ cdf, summary: S1, cfg: CFG, algo: A, mantra: M, current: [] });
  assert.ok(m.missingGroups.includes("Unknown Person"));
});

// ---------------------------------------------------------------- layout
test("layoutSummary", "CDF STATUS at A1, two empty rows between tables", () => { const a = placed(LAY, "status"), b = placed(LAY, "gcc"); assert.deepStrictEqual([a.r1, a.r2, b.r1], [1, 6, 9]); });
test("layoutSummary", "GCC SIGNATURE POSITION, MOR and IGOR start on one row, side by side", () => { const r = ["sig", "pending0", "pending1"].map((id) => placed(LAY, id)); assert.deepStrictEqual(r.map((p) => [p.r1, S.letter(p.c1)]), [[19, "A"], [19, "D"], [19, "G"]]); });
test("layoutSummary", "the macro's formulas", () => {
  assert.strictEqual(at(LAY, "B3").v, '=COUNTIFS(CDF!$R:$R,A3,CDF!$C:$C,"<>")');
  assert.strictEqual(at(LAY, "B5").v, '=COUNTIFS(CDF!$R:$R,A5,CDF!$C:$C,"<>")-COUNTIFS(CDF!$C:$C,"A330-1-052 (another)",CDF!$R:$R,A5)');
  assert.strictEqual(at(LAY, "B6").v, "=SUM(B3:B5)");
  assert.strictEqual(at(LAY, "B11").v, '=COUNTIFS(CDF!$R:$R,$A$3,CDF!$S:$S,A11,CDF!$C:$C,"<>")');
});
test("layoutSummary", "WHERE: several spellings of a supplier counted together", () => { const w = placed(LAY, "where"); const c = at(LAY, "B" + w.first); assert.strictEqual(c.v, '=SUM(COUNTIFS(CDF!$R:$R,$A$3,CDF!$S:$S,$A$13,CDF!$C:$C,"<>",CDF!$P:$P,{"ANN/SIE","SIE/ANN","ANN"}))'); });
test("layoutSummary", "values only when asked", () => assert.strictEqual(S.layoutSummary(MODEL, { formulas: false, aliases: ALIASES }).cells.filter((c) => /^=/.test(String(c.v))).length, 0));
test("layoutSummary", "persons: A, B, C merged over one person's rows; the rest empty", () => {
  const p = placed(LAY, "persons"), m = LAY.merges.filter((x) => x.r1 === p.first + 1 && x.c1 <= 2);
  assert.strictEqual(m.length, 3); assert.strictEqual(at(LAY, "A" + (p.first + 2)).v, "");
});
test("layoutSummary", "REJECTED sums under the table, value in F", () => { const p = placed(LAY, "rejected"); assert.strictEqual(at(LAY, "F" + (p.last + 1)).v, "=SUMPRODUCT(1/COUNTIF(A" + p.first + ":A" + p.last + ",A" + p.first + ":A" + p.last + "))"); });
test("layoutSummary", "an empty table's TOTAL is 0, not a formula over no rows", () => { const p = placed(LAY, "pending1"); assert.deepStrictEqual([at(LAY, "G" + p.first).v, at(LAY, "H" + p.first).v], ["TOTAL", 0]); });
test("layoutSummary", "document tables: the four sums", () => { const p = placed(LAY, "doc24"); assert.deepStrictEqual([0, 1, 2, 3].map((i) => at(LAY, "A" + (p.last + 1 + i)).v), ["ENGINEERING DOCUMENTS", "GCC ANSWERS", "OF WHICH LINKED TO CDFS", "TOTAL"]); });
test("layoutSummary", "the GCC SIGNATURE formula looks inside lists when T has several names", () => { const m = JSON.parse(JSON.stringify(MODEL)); m.tables[2].multiT = true; assert.ok(/"\*"&A21&"\*"/.test(S.layoutSummary(m, { aliases: ALIASES }).cells.filter((c) => S.letter(c.c) + c.r === "B21")[0].v)); });

test("layoutEmail", "subject A1, opening A3, first table at A5", () => assert.deepStrictEqual([at(LE, "A1").v, at(LE, "A3").v.slice(0, 9), placed(LE, "status").r1], [MAIL.subject, "בוקר טוב,", 5]));
test("layoutEmail", "exactly the six tables, in order", () => assert.deepStrictEqual(LE.tables.map((t) => t.id), S.EMAIL_TABLES));
test("layoutEmail", "values, no formulas", () => assert.strictEqual(LE.cells.filter((c) => /^=/.test(String(c.v))).length, 0));
test("layoutEmail", "the signature two rows under the last table", () => { const last = LE.tables[LE.tables.length - 1]; assert.deepStrictEqual([at(LE, "A" + (last.r2 + 2)).v, at(LE, "A" + (last.r2 + 3)).v], ["בברכה,", "Your Name"]); });
test("layoutEmail", "the same persons table as CDF SUMMARY", () => { const a = placed(LE, "persons"), b = placed(LAY, "persons"); assert.strictEqual(a.r2 - a.r1, b.r2 - b.r1); });

test("supplierAliases", "the default spellings", () => assert.deepStrictEqual(ALIASES["ANN/SIE"], ["ANN/SIE", "SIE/ANN", "ANN"]));
test("supplierAliases", "ANCARA", () => assert.deepStrictEqual(ALIASES.ANCARA, ["ANCRA", "ANCARA"]));
test("supplierAliases", "one spelling", () => assert.deepStrictEqual(ALIASES.TIMAT, ["TIMAT"]));
test("supplierAliases", "from SUPPLIER lines", () => assert.ok(S.supplierAliases(A.settings({ suppliers: { "NEWCO": "NEW" } })).NEW.indexOf("NEWCO") >= 0));
test("supplierAliases", "empty settings", () => assert.deepStrictEqual(S.supplierAliases({}), {}));

// ---------------------------------------------------------------- 150: the email
test("jlm", "Israel time in winter (UTC+2)", () => assert.deepStrictEqual(S.jlm(new Date("2026-11-05T06:10:00Z")), { date: "05/11/2026", time: "08:10" }));
test("jlm", "Israel time in summer (UTC+3)", () => assert.strictEqual(S.jlm(new Date("2026-07-01T06:10:00Z")).time, "09:10"));
test("jlm", "the date turns at midnight in Israel", () => assert.strictEqual(S.jlm(new Date("2026-11-05T22:30:00Z")).date, "06/11/2026"));
test("jlm", "midnight is 00, not 24", () => assert.strictEqual(S.jlm(new Date("2026-11-05T22:00:00Z")).time, "00:00"));
test("jlm", "leading zeros", () => assert.strictEqual(S.jlm(new Date("2026-01-02T01:05:00Z")).date, "02/01/2026"));

test("mailParts", "subject (150.4)", () => assert.strictEqual(MAIL.subject, "A330-300 P2F — Daily CDF Status — 05/11/2026"));
test("mailParts", "opening with AS OF and the zone", () => assert.strictEqual(MAIL.opening, "בוקר טוב, להלן סטטוס החריגים המעודכן — AS OF 05/11/2026 08:10 (Asia/Jerusalem)."));
test("mailParts", "signature: בברכה, then the name", () => assert.deepStrictEqual(MAIL.signature, ["בברכה,", "Your Name"]));
test("mailParts", "no name → בברכה, alone", () => assert.deepStrictEqual(S.mailParts(new Date(), []).signature, ["בברכה,"]));
test("mailParts", "בברכה, isn't doubled", () => assert.deepStrictEqual(S.mailParts(new Date(), ["בברכה,", "X"]).signature, ["בברכה,", "X"]));

test("readSignature", "the lines after בברכה,", () => assert.deepStrictEqual(S.readSignature([["a"], ["בברכה,"], [], ["Your Name"]]), ["Your Name"]));
test("readSignature", "no signature", () => assert.deepStrictEqual(S.readSignature([["a"]]), []));
test("readSignature", "empty sheet", () => assert.deepStrictEqual(S.readSignature(null), []));
test("readSignature", "the last בברכה, wins", () => assert.deepStrictEqual(S.readSignature([["בברכה,"], ["Old"], ["בברכה,"], ["New"]]), ["New"]));
test("readSignature", "at most three lines", () => assert.strictEqual(S.readSignature([["בברכה,"], ["a"], ["b"], ["c"], ["d"]]).length, 3));

const DOC = S.emailDoc(MODEL, MAIL);
test("emailDoc", "six tables for the HTML", () => assert.strictEqual(DOC.tables.length, 6));
test("emailDoc", "merged person cells → rowspan", () => { const t = DOC.tables[3]; assert.ok(t.rows.some((r) => r[0] && r[0].rowspan > 1)); assert.ok(t.rows.some((r) => r[0] === null)); });
test("emailDoc", "numbers as text", () => assert.strictEqual(DOC.tables[0].rows[0][1].v, "4"));
test("emailDoc", "subject, opening, signature", () => assert.deepStrictEqual([DOC.subject, DOC.signature.length], [MAIL.subject, 2]));
test("emailDoc", "builds as an email", () => { const m = E.buildDailyEmail(DOC); assert.ok(/rowspan="2"/.test(m.html) && /#B10001/.test(m.html)); });

const EML = S.eml(MAIL, "<table><tr><td>ok</td></tr></table>", "ok", ["a@b.c"], []);
test("eml", "an unsent draft", () => assert.ok(/^X-Unsent: 1\r\n/.test(EML)));
test("eml", "subject encoded in UTF-8", () => { const m = /Subject: =\?UTF-8\?B\?([^?]+)\?=/.exec(EML); assert.strictEqual(Buffer.from(m[1], "base64").toString("utf8"), MAIL.subject); });
test("eml", "text and HTML parts", () => assert.ok(/multipart\/alternative/.test(EML) && /text\/plain; charset=UTF-8/.test(EML) && /text\/html; charset=UTF-8/.test(EML)));
test("eml", "recipients when given", () => assert.ok(/\r\nTo: a@b\.c\r\n/.test(EML)));
test("eml", "the HTML inside decodes", () => { const parts = EML.split(/--=_cdf_\w+/); const html = Buffer.from(parts[2].split("\r\n\r\n")[1].replace(/\r\n/g, ""), "base64").toString("utf8"); assert.ok(/<td>ok<\/td>/.test(html)); });

// ---------------------------------------------------------------- compare and read back
const GRID = (lay) => { const g = []; lay.cells.forEach((c) => { (g[c.r - 1] = g[c.r - 1] || [])[c.c] = c.value; }); for (let i = 0; i < g.length; i++) g[i] = g[i] || []; return g; };
test("compare", "the sheet it writes compares the same", () => assert.ok(S.compare(LAY, GRID(LAY)).every((c) => !S.differs(c) && !c.moved)));
test("compare", "a table missing from the sheet", () => assert.ok(S.compare(LAY, []).every((c) => !c.found)));
test("compare", "a changed count reads as one change", () => { const g = GRID(LAY); g[2][1] = 5; const c = S.compare(LAY, g).filter((x) => x.id === "status")[0]; assert.deepStrictEqual(c.changed[0], { label: "Open", now: "5", will: "4" }); });
test("compare", "a person row more and one less", () => { const g = GRID(LAY); const p = placed(LAY, "persons"); g[p.first - 1][4] = "A330-1-999.-"; const c = S.compare(LAY, g).filter((x) => x.id === "persons")[0]; assert.deepStrictEqual([c.added.length, c.removed.length], [1, 1]); });
test("compare", "merged person cells are read filled down", () => { const g = GRID(LAY); const c = S.compare(LAY, g).filter((x) => x.id === "persons")[0]; assert.strictEqual(c.rowsNow, placed(LAY, "persons").r2 - placed(LAY, "persons").r1 - 1); });
test("compare", "a table further down only moves", () => { const g = [[], [], []].concat(GRID(LAY)); assert.ok(S.compare(LAY, g).every((c) => !S.differs(c) && c.moved)); });
test("compare", "Rejected in another table isn't the REJECTED table", () => { const g = GRID(LAY); const c = S.compare(LAY, g).filter((x) => x.id === "rejected")[0]; assert.strictEqual(c.at, "A" + placed(LAY, "rejected").r1); });

test("differs", "not found", () => assert.ok(S.differs({ found: false })));
test("differs", "rows in", () => assert.ok(S.differs({ found: true, added: [1], removed: [], changed: [] })));
test("differs", "rows out", () => assert.ok(S.differs({ found: true, added: [], removed: [1], changed: [] })));
test("differs", "a changed number", () => assert.ok(S.differs({ found: true, added: [], removed: [], changed: [1] })));
test("differs", "the same", () => assert.ok(!S.differs({ found: true, added: [], removed: [], changed: [] })));

test("readBack", "all the same → nothing", () => { const g = GRID(LAY); assert.deepStrictEqual(S.readBack(LAY, (r, c) => (g[r - 1] || [])[c]), []); });
test("readBack", "another number is reported with its cell", () => { const g = GRID(LAY); g[2][1] = 9; assert.deepStrictEqual(S.readBack(LAY, (r, c) => (g[r - 1] || [])[c])[0], { cell: "B3", want: 4, got: 9 }); });
test("readBack", "a formula not calculated (empty) is left out", () => assert.deepStrictEqual(S.readBack(LAY, () => ""), []));
test("readBack", "text cells aren't compared", () => { const g = GRID(LAY); g[0][0] = "other title"; assert.deepStrictEqual(S.readBack(LAY, (r, c) => (g[r - 1] || [])[c]), []); });
test("readBack", "numbers as text are numbers", () => { const g = GRID(LAY); g[2][1] = "4"; assert.deepStrictEqual(S.readBack(LAY, (r, c) => (g[r - 1] || [])[c]), []); });

// ---------------------------------------------------------------- the sheet after the update, in memory (cdf-mantra.js)
test("finalValues", "rows go in under the anchor, showing it in R", () => { const v = M.finalValues([["h"], row("K", "Open")], [{ kind: "insert", anchor: 2, at: 3, count: 1 }], []); assert.deepStrictEqual([v.length, v[2][C("R")]], [3, "Open"]); });
test("finalValues", "changes land in the row numbers after the operations", () => { const v = M.finalValues([["h"], ["a"], ["b"]], [{ kind: "insert", anchor: 2, at: 3, count: 1 }], [{ row: 4, col: 0, newValue: "B" }]); assert.strictEqual(v[3][0], "B"); });
test("finalValues", "removed rows go", () => assert.deepStrictEqual(M.finalValues([["h"], ["a"], ["b"], ["c"]], [{ kind: "remove", first: 3, last: 3, count: 1 }], []).map((r) => r[0]), ["h", "a", "c"]));
test("finalValues", "no operations → the changes only", () => assert.strictEqual(M.finalValues([["h"], ["a"]], [], [{ row: 2, col: 0, newValue: "x" }])[1][0], "x"));
test("finalValues", "the original is left as it is", () => { const v = [["h"], ["a"]]; M.finalValues(v, [{ kind: "remove", first: 2, last: 2, count: 1 }], []); assert.strictEqual(v.length, 2); });

test("morningDecision", "reports to rebuild count as work, without stopping", () => { const d = F.morningDecision({ errors: [], checks: [], cdfChanges: [], gccChanges: [], reports: { ran: true, summaryDiffers: false, emailDiffers: true } }); assert.deepStrictEqual([d.stop, d.reports], [false, true]); });
test("morningDecision", "nothing to rebuild", () => assert.strictEqual(F.morningDecision({ errors: [], checks: [], cdfChanges: [], gccChanges: [], reports: { ran: true, summaryDiffers: false, emailDiffers: false } }).reports, false));
test("morningDecision", "a skipped stage 3 isn't work", () => assert.strictEqual(F.morningDecision({ errors: [], checks: [], cdfChanges: [], gccChanges: [], reports: { skipped: "x" } }).reports, false));
test("morningDecision", "no stage 3 at all", () => assert.strictEqual(F.morningDecision({ errors: [], checks: [], cdfChanges: [], gccChanges: [] }).reports, false));
test("morningDecision", "a changed CDF SUMMARY", () => assert.strictEqual(F.morningDecision({ errors: [], checks: [], cdfChanges: [], gccChanges: [], reports: { ran: true, summaryDiffers: true } }).reports, true));

const low = Object.entries(perFn).filter(([, n]) => n < 5);
if (low.length) process.exitCode = 1;
console.log(`\n${passed} tests passed, ${failed} failed` + (low.length ? " — fewer than 5 cases: " + low.map(([f, n]) => f + " (" + n + ")").join(", ") : " — every function has 5+ cases"));
