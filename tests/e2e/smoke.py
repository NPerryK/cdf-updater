# End-to-end smoke test of the add-in in a browser, against a simulated Excel (mock-office.js), on an invented workbook:
# three CDFs, three Mantra documents, made-up names. Three "days" — the first run creates BACKUPDATA, the second records an
# event, the third lets go of the first run's snapshot (BACKUPDATA keeps the last two committed runs) — then an undo.
#   pip install playwright && python -m playwright install chromium
#   python tests/e2e/smoke.py
import json, os, sys, mimetypes, urllib.parse
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
def gcc_row(n, status, station, assignee="'-", answer="No"):
    r = [""] * len(HEAD["gcc"])
    g = {h: i for i, h in enumerate(HEAD["gcc"])}
    r[g["CDF Number"]], r[g["Site Name"]], r[g["CDF Status"]], r[g["Open Date"]] = "A330-1-%03d.-" % n, "IAI", status, "46300"
    r[g["Current Station"]], r[g["Current Assignee"]], r[g["New Engineering Doc"]] = station, assignee, answer
    return r
def mantra_row(doc, rev, status, date):
    h = HEAD["mantra"]; r = [None] * len(h)
    r[h.index("Project")], r[h.index("Doc ID")], r[h.index("Rev")], r[h.index("Doc Title")], r[h.index("Status")], r[h.index("Date Status")], r[h.index("Author Name")] = "A330", doc, rev, "INVENTED DOCUMENT", status, date, "Dana Levi"
    return r

CDF = [HEAD["cdf"],
       cdf_row(901, "Open", {"S": "9 - Final Approval", "U": "4 - RELEASED", "V": "-", "W": "-", "X": "-", "AA": "330-11-11-J0901", "AB": "A"}),
       cdf_row(902, "Open", {"S": "3 - Conversion Engineer", "U": "NO ENG. DOC.", "V": "-", "W": "-", "X": "-", "AA": "-", "AB": "-"}),
       cdf_row(903, "Released", {"S": "Released", "Q": 46200, "U": "4 - RELEASED", "V": "-", "W": "-", "X": "-", "AA": "330-33-33-J0903", "AB": "B"})]
GCC = [HEAD["gcc"], gcc_row(901, "Open", "Final Approval"), gcc_row(902, "Open", "Conversion ENG"), gcc_row(903, "Released", "Released")]
MANTRA = [[None] * len(HEAD["mantra"]), HEAD["mantra"],
          mantra_row("330-11-11-J0901", "A", "RELEASED", "2026-11-04T07:30:00"),
          mantra_row("330-22-22-J0902", "A", "IN-WORK", "2026-10-20T09:00:00"),
          mantra_row("330-33-33-J0903", "B", "RELEASED", "2026-10-01T08:00:00")]
PEOPLE = 'PERSON_GROUP "IAI - Design Engineering" "Dana Levi; Ron Avi"\nPERSON_GROUP "Vendor" "Ann Vendor"\n'
# rules 10 and 100.5 name people, so they live in the workbook settings too (rules.txt doesn't carry them)
SIGNERS = 'FINAL_APPROVAL_RELEASE "Tom Bar"\nPENDING_TABLE "DANA PENDING TABLE" "Dana Ben-Ari"\n'
SHEETS = [{"name": "CDF", "values": CDF}, {"name": "GCC", "values": GCC}, {"name": "Mantra", "values": MANTRA}, {"name": "MANTRAREV", "values": [HEAD["mantraRev"]]}]

def boot(browser, priv=PEOPLE + SIGNERS):
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
    page.add_init_script("window.__fixtures={}; window.__initialSheets=%s; window.__xmlParts=%s;" % (json.dumps(SHEETS), json.dumps([{"id": "x0", "xml": enc(NS_S, {"privateRules": priv})}])))
    page.goto("https://addin.test/taskpane.html")
    page.wait_for_selector("#rulesPill.ok, #rulesPill.bad", timeout=20000)
    return page
def check_changes(page):
    page.evaluate("document.getElementById('results').innerHTML = ''")
    page.click("#previewBtn"); page.wait_for_selector("#results .verdict", timeout=60000)
    return page.inner_text("#results .verdict").replace("\n", " ")
def apply(page):
    page.click("#applyBtn"); page.wait_for_selector("#undoNow", timeout=60000)
def runs(page): return page.evaluate("window.__wb.get('BACKUPDATA').values.slice(1).filter(r => r[0]).map(r => [r[0], r[4]])")
def snap_runs(page): return page.evaluate("[...new Set(window.__wb.get('BACKUPDATA').values.slice(1).map(r => r[%d]).filter(Boolean))].sort()" % col("S"))

with sync_playwright() as p:
    browser = p.chromium.launch()

    print("\nThe people settings are missing from the workbook → the pane stops and says what to paste")
    page = boot(browser, priv=PEOPLE)
    check("bad" in page.get_attribute("#rulesPill", "class"), "the rules pill is red: " + page.inner_text("#rulesPill"))
    page.click("#tab-setup")
    er = page.inner_text("#rulesErrors")
    check("Workbook settings: The workbook settings need FINAL_APPROVAL_RELEASE (rule 10) and PENDING_TABLE (100.5)" in er, "the message: " + er[:110])
    check(page.is_disabled("#previewBtn"), "Check changes waits for them")
    page.fill("#privText", PEOPLE + SIGNERS); page.click("#privSave"); page.wait_for_selector("#rulesPill.ok", timeout=10000)
    check(page.is_enabled("#previewBtn"), "pasted and saved → ready")
    page.close()

    page = boot(browser)

    print("\nDay 1 — the first run creates BACKUPDATA")
    page.evaluate("window.__wb.syncs = 0")
    v = check_changes(page)
    n = page.evaluate("window.__wb.syncs")
    check(n <= 8, "Check changes reads the whole workbook in %d Excel round trips (1.9.1: 31 on the real workbook)" % n)
    check("BACKUPDATA is new" in page.inner_text("#results .hist-s"), "History card: BACKUPDATA is new")
    apply(page)
    r1 = runs(page)
    check(len(r1) == 1 and r1[0][1] == "COMMITTED", "one committed run: %s" % r1)
    for s in ("CDF SUMMARY", "ERROR", "EMAIL DRAFT"):
        check(page.evaluate("!!window.__wb.get(%s)" % json.dumps(s)), "sheet %s written" % s)
    check(page.evaluate("window.__wb.get('ERROR').values[6][0]") == "NO ERRORS FOUND", "ERROR: NO ERRORS FOUND")
    check(page.evaluate("window.__wb.active") == "CDF SUMMARY", "ends on CDF SUMMARY")

    print("\nDay 2 — a document released in the last 168 h is recorded once")
    page.clock.set_fixed_time("2026-11-06T06:00:00Z")
    v = check_changes(page)
    ev = page.inner_text("#results .hist-ev")
    check("1 document released" in ev, "History: " + ev.replace("\n", " "))
    apply(page)
    check(len(runs(page)) == 2 and snap_runs(page) == sorted(x[0] for x in runs(page)), "two runs, both with their snapshot")

    print("\nDay 3 — CDF 902 closes in GCC; BACKUPDATA lets go of day 1's snapshot")
    page.clock.set_fixed_time("2026-11-07T06:00:00Z")
    page.evaluate("(() => { const g = window.__wb.get('GCC').values; const h = g[0]; g[2][h.indexOf('CDF Status')] = 'Released'; g[2][h.indexOf('Current Station')] = 'Released'; g[2][h.indexOf('Close Date')] = '46333'; })()")
    v = check_changes(page)
    tr = page.inner_text("#results .trim-roll") if page.locator("#results .trim-roll").count() else ""
    check("Lets go of 1 older run's snapshot" in tr, "rolling trim, no question asked: " + tr)
    apply(page)
    rr = runs(page)
    check(len(rr) == 3, "RUNS keeps all three runs")
    check(snap_runs(page) == sorted([rr[1][0], rr[2][0]]), "snapshots of the last two runs only: %s" % snap_runs(page))

    print("\nUndo — the run is marked FAILED, the baseline goes back to day 2")
    page.click("#tab-setup"); page.click("#undoBtn"); page.wait_for_selector("#undoBox .confirm", timeout=30000)
    page.click("#undoGo"); page.wait_for_selector("#undoBox .note.ok", timeout=30000)
    check(runs(page)[2][1] == "FAILED", "day 3's run is FAILED")
    page.click("#tab-today")
    check_changes(page)
    check(page.inner_text("#results .hist-s").startswith("since " + rr[1][0]), "the baseline is day 2's run again")
    browser.close()

print("\n" + ("ALL OK" if not errors else "%d FAILED:\n  " % len(errors) + "\n  ".join(errors)))
sys.exit(1 if errors else 0)
