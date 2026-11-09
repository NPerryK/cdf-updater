# End-to-end test of the LOOK page and the look of sheet CDF (110.6–110.7), in a browser against the simulated Excel
# (mock-office.js), on an invented workbook. Made-up names only.
#   python tests/e2e/look.py
import json, os, sys, mimetypes, re, urllib.parse
from playwright.sync_api import sync_playwright

HERE = os.path.dirname(os.path.abspath(__file__))
ADDIN = os.path.abspath(os.path.join(HERE, "..", ".."))
HEAD = json.load(open(os.path.join(HERE, "headers.json"), encoding="utf-8"))
NS_S = "urn:cdf-updater:settings:v1"
errors = []
def check(cond, msg):
    print(("  ✓ " if cond else "  ✗ ") + msg)
    if not cond: errors.append(msg)
def col(L):
    n = 0
    for ch in L: n = n * 26 + ord(ch) - 64
    return n - 1
def enc(ns, obj): return '<cdfUpdater xmlns="%s">%s</cdfUpdater>' % (ns, urllib.parse.quote(json.dumps(obj, ensure_ascii=False), safe="-_.!~*'()"))

def cdf_row(n, R, x):
    r = [""] * 34
    k = "A330-1-%03d" % n
    r[col("A")], r[col("B")], r[col("C")], r[col("D")], r[col("F")], r[col("G")], r[col("R")] = k + "-ATT.", k + "-", k + ".-", "-", 46300, "ROUTINE", R
    for c, v in x.items(): r[col(c)] = v
    return r
def cont_row(x):
    r = [""] * 34
    for c, v in x.items(): r[col(c)] = v
    return r
def gcc_row(n, status, station):
    r = [""] * len(HEAD["gcc"]); g = {h: i for i, h in enumerate(HEAD["gcc"])}
    r[g["CDF Number"]], r[g["Site Name"]], r[g["CDF Status"]], r[g["Open Date"]] = "A330-1-%03d.-" % n, "IAI", status, "46300"
    r[g["Current Station"]], r[g["Current Assignee"]], r[g["New Engineering Doc"]] = station, "'-", "No"
    return r
def mantra_row(doc, rev, status, date):
    h = HEAD["mantra"]; r = [None] * len(h)
    r[h.index("Project")], r[h.index("Doc ID")], r[h.index("Rev")], r[h.index("Doc Title")], r[h.index("Status")], r[h.index("Date Status")], r[h.index("Author Name")] = "A330", doc, rev, "INVENTED DOCUMENT", status, date, "Dana Levi"
    return r

# 901: anchor + 2 solution rows; 902 open; 903 released; 904 cancelled; then two template rows
CDF = [HEAD["cdf"],
       cdf_row(901, "Open", {"S": "9 - Final Approval", "U": "4 - RELEASED", "V": "-", "W": "-", "X": "-", "AA": "330-11-11-J0901", "AB": "A", "H": "A LONG DESCRIPTION " * 6}),
       cont_row({"R": "Open", "U": "4 - RELEASED", "V": "-", "W": "-", "X": "-", "AA": "330-11-11-J0911", "AB": "A"}),
       cont_row({"R": "Open", "U": "4 - RELEASED", "V": "-", "W": "-", "X": "-", "AA": "330-11-11-J0912", "AB": "B"}),
       cdf_row(902, "Open", {"S": "3 - Conversion Engineer", "U": "NO ENG. DOC.", "V": "-", "W": "-", "X": "-", "AA": "-", "AB": "-", "F": 46000}),
       cdf_row(903, "Released", {"S": "Released", "Q": 46200, "U": "4 - RELEASED", "V": "-", "W": "-", "X": "-", "AA": "330-33-33-J0903", "AB": "B"}),
       cdf_row(904, "Cancelled", {"S": "Cancelled", "U": "-", "V": "-", "W": "-", "X": "-", "AA": "-", "AB": "-"}),
       cont_row({"A": "A330-1-905.-ATT.", "B": "A330-1-905.-"}),
       cont_row({"A": "A330-1-906.-ATT.", "B": "A330-1-906.-"})]
F = [[] for _ in CDF]
for r in (2, 3):   # 0-based rows 2,3 are 901's continuation rows: references to the anchor in row 2
    F[r] = [None] * 34
    for L in ("E", "H", "O", "R", "S", "T", "AC", "AD", "AE", "AF", "AG", "AH"): F[r][col(L)] = '=IF($%s$2="","",$%s$2)' % (L, L)
MERGES = [{"r1": 1, "r2": 3, "c1": col(L), "c2": col(L)} for L in ("A", "B", "C", "D", "F", "G", "I", "J", "K", "L", "M", "N", "P", "Y", "Z")]
FORMATS = [["General"] * 34 for _ in CDF]; FORMATS[5][col("L")] = ";;;"
for r in range(1, 7): FORMATS[r][col("F")] = "m/d/yyyy"
# what the 05/11 workbook has: Released / Cancelled in pieces, stray copies, the continuation rule — and one rule of the user's own
CFS = [{"formula": '=TRIM($R2)="Released"', "areas": "A2:AH4, A6:AH9, A5:B5", "fill": "#00B050"},
       {"formula": '=TRIM($R2)="Cancelled"', "areas": "A2:AH4, A6:AH9", "fill": "#FF0000"},
       {"formula": "=TRUE", "areas": "Q3:X3, AA3:AH3", "top": "#000000"},
       {"formula": '=TRIM($R5)="Released"', "areas": "C5", "fill": "#00B050"},
       {"formula": "=$Q2>46100", "areas": "Q2:Q9", "font": "#0000FF"}]
HEIGHTS = [None, 12.75, 12.75, 12.75, 60, 12.75, 12.75, 12.75, 12.75]   # row 5 made tall by hand (merged text)
GCC = [HEAD["gcc"], gcc_row(901, "Open", "Final Approval"), gcc_row(902, "Open", "Conversion ENG"), gcc_row(903, "Released", "Released"), gcc_row(904, "Cancelled", "Cancelled")]
MANTRA = [[None] * len(HEAD["mantra"]), HEAD["mantra"],
          mantra_row("330-11-11-J0901", "A", "RELEASED", "2026-11-04T07:30:00"), mantra_row("330-11-11-J0911", "A", "RELEASED", "2026-11-04T07:30:00"),
          mantra_row("330-11-11-J0912", "B", "RELEASED", "2026-11-04T07:30:00"), mantra_row("330-33-33-J0903", "B", "RELEASED", "2026-10-01T08:00:00")]
PEOPLE = 'PERSON_GROUP "IAI - Design Engineering" "Dana Levi; Ron Avi"\nFINAL_APPROVAL_RELEASE "Tom Bar"\nPENDING_TABLE "DANA PENDING TABLE" "Dana Ben-Ari"\n'
SHEETS = [{"name": "CDF", "values": CDF, "f": F, "merges": MERGES, "formats": FORMATS, "cfs": CFS, "heights": HEIGHTS, "hidden": [8]},
          {"name": "GCC", "values": GCC}, {"name": "Mantra", "values": MANTRA}, {"name": "MANTRAREV", "values": [HEAD["mantraRev"]]}]

ANSWER_GPT = """json
Copy code
{"cdf_look": 1, "looks": [
  {"name": "Old open CDFs",
   "when": {"all": [{"col": "R", "op": "is", "value": "Open"}, {"col": "F", "op": "older_than_days", "value": 200}]},
   "paint": {"cells": "A:H", "fill": "#FFC000", "bold": true}}
]}"""
ANSWER_CLAUDE = """Here is the rule:
```json
{"cdf_look": 1, "looks": [{"name": "At Final Approval", "when": {"col": "S", "op": "has", "value": "final approval"}, "paint": {"cells": ["R", "S"], "font": "dark red", "italic": true}}]}
```"""

def boot(browser, extra_init="", parts=None, sheets=None):
    page = browser.new_page(viewport={"width": 380, "height": 900})
    page.on("pageerror", lambda e: (print("   [pageerror]", e), errors.append(str(e))))
    def handle(route):
        url = route.request.url
        if "appsforoffice.microsoft.com" in url: return route.fulfill(path=os.path.join(HERE, "mock-office.js"), content_type="application/javascript")
        if url.startswith("https://addin.test/"):
            fp = os.path.join(ADDIN, url[len("https://addin.test/"):].split("?")[0] or "taskpane.html")
            if os.path.isfile(fp): return route.fulfill(path=fp, content_type=(mimetypes.guess_type(fp)[0] or "text/plain") + ("; charset=utf-8" if fp.endswith(".txt") else ""))
            return route.fulfill(status=404, body="nf")
        return route.abort()
    page.route("**/*", handle)
    page.clock.set_fixed_time("2026-11-05T06:00:00Z")
    page.add_init_script("window.__fixtures={}; window.__initialSheets=%s; window.__xmlParts=%s; window.__opened=[]; %s" % (json.dumps(sheets or SHEETS), json.dumps(parts or [{"id": "x0", "xml": enc(NS_S, {"privateRules": PEOPLE})}]), extra_init))
    page.add_init_script("window.open = function (u) { window.__opened.push(u); }; Object.defineProperty(navigator, 'clipboard', { value: { writeText: function (t) { window.__clip = t; return Promise.resolve(); } } });")
    page.goto("https://addin.test/taskpane.html")
    page.wait_for_selector("#rulesPill.ok, #rulesPill.bad", timeout=20000)
    return page
def cfs(page): return page.evaluate("window.__wb.get('CDF').cfs.map(r => ({ f: r.custom.rule.formula, a: r.areas.map(b => [b.r1+1,b.c1,b.r2+1,b.c2]), fill: r.custom.format.fill.color, font: r.custom.format.font.color, bold: r.custom.format.font.bold, italic: r.custom.format.font.italic, top: (r.custom.format.borders.items.EdgeTop||{}).color }))")
def addr(a):
    def L(i):
        s = ""; i += 1
        while i: i, m = divmod(i - 1, 26); s = chr(65 + m) + s
        return s
    return "%s%d:%s%d" % (L(a[1]), a[0], L(a[3]), a[2])
def areas(rule): return ", ".join(addr(a) for a in rule["a"])
def fmt_writes(page, sheet, prop):
    return page.evaluate("(window.__wb.fmt||[]).filter(x => x.sheet === %s).map(x => ({ r: x.r, c: x.c, nr: x.nr, nc: x.nc, v: %s })).filter(x => x.v !== undefined && x.v !== null)" % (json.dumps(sheet), prop))
def heights(page): return page.evaluate("window.__wb.get('CDF').heights.slice(0, 10)")

with sync_playwright() as p:
    browser = p.chromium.launch()

    print("\nRepaint CDF on a sheet with broken rules")
    page = boot(browser)
    check(page.is_visible("#tab-look"), "a LOOK key next to the other four")
    page.click("#tab-look")
    check(page.is_visible("#repaintBtn") and page.is_enabled("#repaintBtn"), "Repaint CDF is ready")
    page.evaluate("window.__wb.syncs = 0; window.__wb.fmt = []; window.__wb.borders = []")
    page.click("#repaintBtn"); page.wait_for_selector("#repaintMsg .look-report", timeout=20000)
    rep = page.inner_text("#repaintMsg")
    print("   " + rep.replace("\n", "\n   "))
    R = cfs(page)
    ours = [r for r in R if r["f"] != "=$Q2>46100"]
    check([r["f"] for r in ours] == ['=TRIM($R2)="Released"', '=TRIM($R2)="Cancelled"', "=TRUE"], "exactly the three rules of 110.7, in order: %s" % [r["f"] for r in ours])
    check(areas(ours[0]) == "A2:AH9" and areas(ours[1]) == "A2:AH9", "Released / Cancelled on one range each: " + areas(ours[0]))
    check(ours[0]["fill"] == "#00B050" and ours[1]["fill"] == "#FF0000", "green and red fills")
    check(areas(ours[2]) == "Q3:X4, AA3:AH4" and ours[2]["top"] == "#000000" and not ours[2]["fill"], "the line over 901's two solution rows: " + areas(ours[2]))
    own = [r for r in R if r["f"] == "=$Q2>46100"]
    check(len(own) == 1 and areas(own[0]) == "Q2:Q9", "the user's own rule is kept as it was")
    check(R.index(own[0]) == 3, "…below the add-in's rules")
    check("were split" in rep and "Kept 1 conditional format of your own" in rep, "the report says what was split and what was kept")
    fills = fmt_writes(page, "CDF", "x.f.fill.color")
    check(any(x["v"] == "#60497A" and x["r"] == 1 and x["c"] == col("S") and x["nc"] == 6 and x["nr"] == 8 for x in fills), "S2:X9 purple")
    check(not any(x["v"] == "#60497A" and x["r"] == 0 for x in fills), "the header row keeps its own fill")
    fonts = fmt_writes(page, "CDF", "x.f.font.name")
    check(any(x["v"] == "Calibri" and x["r"] == 0 and x["nr"] == 9 and x["nc"] == 34 for x in fonts), "Calibri on A1:AH9")
    colors = fmt_writes(page, "CDF", "x.f.font.color")
    check(all(x["c"] == col("S") for x in colors), "font color changes only in S:X (links in A and C keep theirs)")
    b = page.evaluate("window.__wb.borders.filter(x => x.sheet === 'CDF').map(x => [x.b.side, x.b.style, x.b.weight, x.b.color, x.r, x.nr, x.nc])")
    check(sorted(set(x[0] for x in b)) == sorted(["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideVertical", "InsideHorizontal"]) and all(x[1:4] == ["Continuous", "Thin", "#000000"] for x in b), "thin black lines on every side, inside too")
    check(page.evaluate("window.__wb.get('CDF').formats[5][%d]" % col("L")) == "General", "the ;;; cell shows its value again")
    h = heights(page)
    check(h[4] == 60, "row 5 (made tall by hand) keeps 60 pt: %s" % h[4])
    check(h[1] > 12.75, "row 2 grew to fit its long description: %s" % h[1])
    check(page.evaluate("window.__wb.autofitHidden || 0") == 0, "autofit never touches the row a filter hides")
    check(page.evaluate("window.__wb.get('CDF').hidden[7]") is True, "the filtered row is still hidden")
    n = page.evaluate("window.__wb.syncs")
    check(n <= 8, "%d Excel round trips (3 to read the sheet, 4 to paint)" % n)

    print("\nRepaint again: nothing piles up")
    page.click("#repaintBtn"); page.wait_for_timeout(300); page.wait_for_selector("#repaintMsg .look-report", timeout=20000)
    R2 = cfs(page)
    check(len(R2) == 4 and [r["f"] for r in R2] == [r["f"] for r in R], "still 3 + 1 rules, same order")
    check("were split" not in page.inner_text("#repaintMsg"), "nothing split this time")

    print("\nAsk ChatGPT: copy the request, paste the answer, add it")
    page.click("#lookCopy")
    check("Write first" in page.inner_text("#lookCopyMsg"), "an empty request asks for words first")
    page.fill("#lookAsk", "orange: open CDFs opened more than 200 days ago")
    page.click("#lookCopy"); page.wait_for_selector("#lookCopyMsg.ok", timeout=10000)
    clip = page.evaluate("window.__clip")
    check("orange: open CDFs opened more than 200 days ago" in clip, "the request is in the copied text")
    check("R — " in clip and "F — " in clip and "(date)" in clip, "with the column letters, names and kinds")
    check("A330-1-901" not in clip and "330-11-11-J0901" not in clip and "Dana" not in clip and "LONG DESCRIPTION" not in clip, "and no value from the sheet")
    page.click("#openChatGpt"); page.click("#openClaudeLook")
    check(page.evaluate("window.__opened") == ["https://chatgpt.com/", "https://claude.ai/new"], "ChatGPT and Claude open in the browser")
    page.fill("#lookPaste", ANSWER_GPT); page.wait_for_selector("#lookPreview .look-item", timeout=10000)
    pv = page.inner_text("#lookPreview")
    check("Old open CDFs" in pv and "more than 200 days ago" in pv, "the rule in plain words")
    check("1 row now: 902" in pv, "and what it paints now: " + re.sub(r"\s+", " ", pv)[:160])
    page.click("#lookAdd"); page.wait_for_selector("#lookAddMsg .look-report", timeout=20000)
    R3 = cfs(page)
    lk = [r for r in R3 if "TODAY()" in r["f"]]
    check(len(lk) == 1 and areas(lk[0]) == "A2:H9" and lk[0]["fill"] == "#FFC000" and lk[0]["bold"] is True, "written as a live rule on A2:H9: %s" % (lk and lk[0]["f"]))
    check(R3.index(lk[0]) == 3, "after Released / Cancelled / solution rows, before the user's own")
    check("Old open CDFs" in page.inner_text("#savedLooks") and page.inner_text("#looksCount") == "1", "listed under Looks in this workbook")
    store = page.evaluate("window.__xmlParts.filter(p => p.xml.indexOf('looks:v1') > 0).length")
    check(store == 1, "kept in the workbook (custom XML)")

    print("\nAsk Claude too: a fenced answer with a list of cells")
    page.fill("#lookPaste", ANSWER_CLAUDE); page.wait_for_selector("#lookPreview .look-item", timeout=10000)
    check("3 rows now: 901" in page.inner_text("#lookPreview") and "row 3" not in page.inner_text("#lookPreview"), "counted: " + re.sub(r"\s+", " ", page.inner_text("#lookPreview"))[:120])
    page.click("#lookAdd"); page.wait_for_selector("#lookAddMsg .look-report", timeout=20000)
    lk2 = [r for r in cfs(page) if "SEARCH" in r["f"]]
    check(len(lk2) == 1 and areas(lk2[0]) == "R2:S9" and lk2[0]["font"] == "#C00000" and lk2[0]["italic"] is True, "R2:S9, dark red italic text")

    print("\nA bad answer says why")
    page.fill("#lookPaste", '{"looks":[{"when":{"col":"R","op":"between","value":[1,2]},"paint":{"fill":"red"}}]}'); page.wait_for_selector("#lookPreview .note.bad", timeout=10000)
    check("isn't an operation" in page.inner_text("#lookPreview"), page.inner_text("#lookPreview")[:90])

    print("\nThe looks survive reopening the workbook and an Apply")
    parts = page.evaluate("window.__xmlParts")
    sheets_now = page.evaluate("""window.__wb.sheets.map(s => ({ name: s.name, values: s.values, f: s.f, merges: s.merges, formats: s.formats, heights: s.heights,
      hidden: s.hidden.map((h, i) => h ? i + 1 : 0).filter(Boolean),
      cfs: s.cfs.map(r => { const L = (i) => { let n = i + 1, t = ''; while (n) { const m = (n - 1) % 26; t = String.fromCharCode(65 + m) + t; n = Math.floor((n - 1) / 26); } return t; };
        const f = r.custom.format; return { formula: r.custom.rule.formula, areas: r.areas.map(b => L(b.c1) + (b.r1 + 1) + ':' + L(b.c2) + (b.r2 + 1)).join(', '), fill: f.fill.color, font: f.font.color, bold: f.font.bold, italic: f.font.italic, top: (f.borders.items.EdgeTop || {}).color }; }) }))""")
    page.close()
    page = boot(browser, parts=parts, sheets=sheets_now)
    page.click("#tab-look")
    check(page.inner_text("#looksCount") == "2", "2 looks read back from the workbook")
    page.click("#tab-today")
    page.click("#previewBtn"); page.wait_for_selector("#results .verdict", timeout=60000)
    check("At Apply, the look of rows 2–7" in page.inner_text("#results") and "2 saved looks" in page.inner_text("#results"), "Check changes says what Apply paints")
    page.click("#applyBtn"); page.wait_for_selector("#undoNow", timeout=60000)
    res = page.inner_text("#results")
    check("Sheet look ✓" in res and "2 saved looks painted again" in res, "Apply painted the sheet and the looks")
    R4 = cfs(page)
    vals = page.evaluate("window.__wb.get('CDF').values.length")
    check(sum(1 for r in R4 if "TODAY()" in r["f"] or "SEARCH" in r["f"]) == 2 and len(R4) == 6, "3 rules + 2 looks + the user's own, no duplicates: %d rules" % len(R4))
    contr = [r for r in R4 if r["f"] in ("=TRUE", "=FALSE")][0]
    check(contr["f"] == "=FALSE" and areas(contr) == "Q2:X7, AA2:AH7", "Apply removed 901's two extra solution rows: the line rule follows (=FALSE, no solution rows): " + areas(contr))
    check(areas([r for r in R4 if "Released" in r["f"]][0]) == "A2:AH7", "Released / Cancelled now end at row 7")

    print("\nRemove a look")
    page.click("#tab-look")
    page.click("[data-unlook='0']"); page.wait_for_selector("#repaintMsg .look-report", timeout=20000)
    R5 = cfs(page)
    check(not any("TODAY()" in r["f"] for r in R5) and any("SEARCH" in r["f"] for r in R5), "its rule is gone, the other stays")
    check(page.inner_text("#looksCount") == "1", "one look left")
    page.close()

    print("\nExcel without RangeAreas (ExcelApi 1.9): the line goes on each range")
    page = boot(browser, "window.__noApi = { 'ExcelApi 1.9': true };")
    page.click("#tab-look"); page.click("#repaintBtn"); page.wait_for_selector("#repaintMsg .look-report", timeout=20000)
    R6 = cfs(page)
    check(sum(1 for r in R6 if r["f"] == "=TRUE") == 2 and "can't put one rule on several ranges" in page.inner_text("#repaintMsg"), "two =TRUE rules (Q3:X4 and AA3:AH4), and the report says why")
    browser.close()

print("\n" + ("ALL OK" if not errors else "%d FAILED:\n  " % len(errors) + "\n  ".join(errors)))
sys.exit(1 if errors else 0)
