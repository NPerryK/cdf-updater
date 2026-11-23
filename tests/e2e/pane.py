# End-to-end test of CDF Updater 2 in a browser, against the simulated Excel (mock-office.js), on an invented workbook.
# Made-up names only.   pip install playwright openpyxl && python -m playwright install chromium && python tests/e2e/pane.py
import json, os, sys, mimetypes, re, tempfile, urllib.parse
from playwright.sync_api import sync_playwright
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
ADDIN = os.path.abspath(os.path.join(HERE, "..", ".."))
HEAD = json.load(open(os.path.join(HERE, "headers.json"), encoding="utf-8"))
errors = []
def check(cond, msg):
    print(("  ✓ " if cond else "  ✗ ") + msg)
    if not cond: errors.append(msg)
def col(L):
    n = 0
    for ch in L: n = n * 26 + ord(ch) - 64
    return n - 1
TODAY = 46335          # Mon 09/11/2026 (the clock below)
import datetime
def ddmmyyyy(n): return (datetime.date(1899, 12, 30) + datetime.timedelta(days=n)).strftime("%d/%m/%Y")
def cdf_row(n, R, x, rev="-"):
    r = [""] * 34
    k = "A330-1-%03d" % n
    r[col("A")], r[col("B")], r[col("C")], r[col("D")], r[col("F")], r[col("G")], r[col("R")] = k + "-ATT.", k + "-", k + "." + rev, rev, TODAY - 40, "ROUTINE", R
    for c, v in x.items(): r[col(c)] = v
    return r
def cont(x):
    r = [""] * 34
    for c, v in x.items(): r[col(c)] = v
    return r
CDF = [HEAD["cdf"],
       cdf_row(901, "Open", {"S": "9 - Final Approval", "P": "ANN", "U": "4 - RELEASED", "AA": "330-11-11-J0901", "AB": "A", "E": "call the supplier"}),
       cont({"R": "Open", "U": "4 - RELEASED", "AA": "330-11-11-J0911", "AB": "A"}),
       cdf_row(902, "Open", {"S": "3 - Conversion Engineer", "P": "ANN", "U": "NO ENG. DOC.", "AA": "-", "F": TODAY - 200}),
       cdf_row(903, "Released", {"S": "Released", "Q": TODAY - 3, "U": "4 - RELEASED", "AA": "330-33-33-J0903"}),
       cdf_row(904, "Cancelled", {"S": "Cancelled"}),
       cdf_row(906, "Open", {"S": "3 - Conversion Engineer", "P": "TIMAT", "F": TODAY - 2}),
       cont({"A": "A330-1-905.-ATT.", "B": "A330-1-905.-"}),      # the empty template row of 905
       cont({"A": "A330-1-907.-ATT.", "B": "A330-1-907.-"})]
MERGES = [{"r1": 1, "r2": 2, "c1": col(L), "c2": col(L)} for L in ("A", "B", "C", "D", "F", "G")]
def gcc_row(n, status, station, opened, closed=""):
    r = [""] * len(HEAD["gcc"]); g = {h: i for i, h in enumerate(HEAD["gcc"])}
    r[g["CDF Number"]], r[g["CDF Status"]], r[g["Open Date"]], r[g["Current Station"]] = "A330-1-%03d.-" % n, status, opened, station
    if "Close Date" in g: r[g["Close Date"]] = closed
    return r
GCC = [HEAD["gcc"], gcc_row(901, "Open", "Final Approval", TODAY - 30), gcc_row(902, "Open", "Conversion ENG", TODAY - 200), gcc_row(903, "Released", "Released", TODAY - 60, TODAY - 3),
       gcc_row(904, "Cancelled", "Cancelled", TODAY - 90), gcc_row(906, "Open", "Conversion ENG", TODAY - 2)]
REVH = ["Project", "Doc ID", "Rev", "Doc Title", "EO ID", "Doc Model", "Author", "Author Dept", "Release Date", "Sign Date", "Template", "Function", "Approver", "WF ID", "Days Cal", "Days Wd", "Days Last Cal", "Days Last Wd", "Doc Type", "EO Action", "EO Reason"]
def rev_row(i, day, days):
    r = [""] * len(REVH); r[0], r[1], r[2], r[3], r[8], r[14], r[18], r[19] = "A330", str(1000 + i), "A", "INVENTED DOCUMENT %d" % i, day, days, "XD" if i % 2 else "DWG", "RELEASE NEW DOCUMENT"
    return r
REV = [REVH] + [rev_row(i, TODAY - (i % 20), 10 + i % 9) for i in range(30)]
REV[1][1] = "330-11-11-J0901"        # 901's first solution document: rev A released; its second (J0911) is in work
E = ""
SUMMARY = [["CDF STATUS", E, E, E, E], ["CDF STATUS", "CDF'S", E, E, E], ["Open", 3, E, E, E], ["Cancelled", 1, E, E, E], ["Released", 1, E, E, E], ["TOTAL", 5, E, E, E], [E] * 5,
           ["WHERE ARE THE CDFS", E, E, "DANA PENDING TABLE", E], ["SUPPLIER", "CDF'S", E, "CDF NUMBER", "SOLUTION / DOCUMENT"], ["ANN/SIE", 1, E, "A330-1-901.-", "330-11-11-J0911 Rev A"], ["TIMAT", 1, E, "TOTAL", 1], ["TOTAL", 2, E, E, E], [E] * 5,
           ["REJECTED", E, E, E, E], ["CDF NUMBER", "CDF STATUS", "DOCUMENT", "REVISION", "REJECTED BY"], ["A330-1-906.-", "Open", "330-66-00-J0906", "A", "TOM BAR"],
           ["REJECTED CDFS", E, E, E, 1.0000000000000002]]
SHEETS = [{"name": "CDF", "values": CDF, "merges": MERGES}, {"name": "GCC", "values": GCC}, {"name": "CDF SUMMARY", "values": SUMMARY}]
NEW_GCC = [HEAD["gcc"]] + GCC[1:] + [gcc_row(908, "Open", "Initiator", TODAY)]

def xlsx(path, sheet, rows):
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = sheet
    for r in rows: ws.append([None if v == "" else v for v in r])
    wb.save(path)

def boot(browser, init=""):
    page = browser.new_page(viewport={"width": 380, "height": 1000}, accept_downloads=True)
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
    page.clock.set_fixed_time("2026-11-09T08:00:00Z")
    page.add_init_script("window.__fixtures=%s; window.__initialSheets=%s; window.__xmlParts=[]; %s" % (json.dumps({"CDF_List_Export": NEW_GCC, "Released": REV}), json.dumps(SHEETS), init))
    page.add_init_script("Object.defineProperty(navigator, 'clipboard', { value: { writeText: function (t) { window.__clip = t; return Promise.resolve(); }, write: function (items) { window.__clipRich = true; return Promise.resolve(); } } });")
    page.goto("https://addin.test/taskpane.html")
    page.wait_for_selector("#dash .dcard", timeout=20000)
    return page
def sheet(page, name): return page.evaluate("(window.__wb.get(%s) || {}).values || null" % json.dumps(name))

with sync_playwright() as p:
    browser = p.chromium.launch()
    tmp = tempfile.mkdtemp()

    print("\nSTATUS — the dashboard reads the sheets as they are")
    page = boot(browser)
    check(page.inner_text("#rulesPill") == "RULES OK", "rules.txt reads without errors")
    kp = page.inner_text("#dash .kpis")
    check(re.search(r"3\s*Open", kp) and re.search(r"1\s*Released", kp) and re.search(r"1\s*Cancelled", kp), "Open 3 · Released 1 · Cancelled 1: " + kp.replace("\n", " "))
    check(re.search(r"\b25(\.0)?%|\b25%", kp) is not None, "25% released (1 of 4 not cancelled): " + kp.replace("\n", " "))
    st = page.inner_text("[data-card=stations]")
    check("Conversion Engineer" in st and "Final Approval" in st, "stations: " + st.replace("\n", " ")[:120])
    check("SHEETS 2/3" in page.inner_text("#sheetsLamp"), "the lamp says MANTRAREV is missing: " + page.inner_text("#sheetsLamp"))
    check("No MANTRAREV sheet yet" in page.inner_text("[data-card=docs]"), "the solution documents card says to import MANTRAREV")
    fs = page.evaluate("getComputedStyle(document.querySelector('#dash .kpi .kn')).fontSize")
    check(fs == "21px", "the status card's numbers are smaller: " + fs)
    ag = page.inner_text("[data-card=aging]")
    check("902" in ag and "200 d" in ag, "oldest open: 902, 200 days")
    check("2 open with no engineering document" in page.inner_text("[data-card=attention]").replace("\n", " "), "902 (NO ENG. DOC.) and 906 (no document at Conversion Engineer) wait for an engineering document")
    gl = page.inner_text("#glance")
    check("OPEN" in gl and "RELEASED" in gl, "the display on top shows the numbers")
    page.click("#dash .kpi.c-open"); page.wait_for_selector("#list:not(.hidden)")
    check(page.inner_text("#listCount") == "3 CDFs" and "A330-1-902.-" in page.inner_text("#listBody"), "a click on Open lists the 3 CDFs")
    page.click("#listBody [data-key='A330-1-902.-']")
    check(page.evaluate("window.__wb.active") == "CDF", "a click on a CDF goes to its row")
    page.click("#listClose")
    trend = page.evaluate("window.__xmlParts.filter(p => p.xml.indexOf('trend:v2') > 0).length")
    check(trend == 1, "today's point for the trend is kept in the workbook")

    print("\nPresent, PowerPoint, Copy")
    page.click("#dashPresent")
    check(page.evaluate("(window.__dialogs || []).some(u => /present\\.html\\?v=2\\.3\\.0$/.test(u))"), "Present opens present.html in an Office dialog")
    handed = page.evaluate("JSON.parse(localStorage.getItem('cdfUpdater.present.v2')).model.status.open")
    check(handed == 3, "the model is handed to the present view")
    with page.expect_download(timeout=30000) as dl: page.click("#dashDeck")
    d = dl.value; path = os.path.join(tmp, d.suggested_filename); d.save_as(path)
    check(d.suggested_filename == "CDF_Status_20261109.pptx" and os.path.getsize(path) > 20000, "PowerPoint deck saved: %s (%d KB)" % (d.suggested_filename, os.path.getsize(path) // 1024))
    import zipfile
    z = zipfile.ZipFile(path); slides = [n for n in z.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", n)]; charts = [n for n in z.namelist() if re.match(r"ppt/charts/chart\d+\.xml$", n)]
    check(len(slides) >= 5 and len(charts) >= 5, "%d slides with %d native charts" % (len(slides), len(charts)))
    page.click("#dashCopy"); page.wait_for_timeout(200)
    check(page.evaluate("window.__clipRich === true"), "Copy puts the status table on the clipboard")

    print("\nIMPORT — files go to their sheets by their headers")
    xlsx(os.path.join(tmp, "CDF_List_Export.xlsx"), "CDF_List_Export", NEW_GCC)
    xlsx(os.path.join(tmp, "export 2026-11-09.xlsx"), "Released", REV)          # a name that says nothing — the headers tell
    xlsx(os.path.join(tmp, "notes.xlsx"), "Sheet1", [["a", "b"], [1, 2]])
    page.click("#tab-import")
    page.set_input_files("#dropInput", [os.path.join(tmp, f) for f in ("CDF_List_Export.xlsx", "export 2026-11-09.xlsx", "notes.xlsx")])
    page.wait_for_selector(".slot2.has", timeout=10000); page.wait_for_timeout(300)
    sl = page.inner_text("#slots")
    check("export 2026-11-09.xlsx" in sl and re.search(r"MANTRAREV[\s\S]*export 2026-11-09", sl) is not None, "the MANTRAREV export is recognised by its headers, not its name")
    check("not in the workbook yet — will be created" in sl, "MANTRAREV doesn't exist yet: it will be created")
    check("neither its headers nor its name match" in sl, "an unrelated file asks which sheet")
    check(page.inner_text("#importBtn") == "Replace GCC, MANTRAREV", "the button names the sheets: " + page.inner_text("#importBtn"))
    print("\nIMPORT — straight onto a sheet's box")
    M650 = [[""] * 7, ["Project", "Doc ID", "Rev", "Status", "Date Status", "Pending To Emp", "Due Date"]] + [["A330", "330-1-%d" % i, "A", "IN-WORK", TODAY - 2 - i, "Tom Bar", TODAY + 40] for i in range(6)]
    xlsx(os.path.join(tmp, "650 (29).xlsx"), "650", M650)
    page.evaluate("m => { window.__fixtures['650'] = m; }", M650)
    check(page.query_selector("[data-slot=Mantra]") is not None and page.query_selector("[data-slot=CDF]") is None, "the boxes are Mantra, GCC, MANTRAREV — sheet CDF isn't replaced here")
    with page.expect_file_chooser() as fc: page.click("[data-slot=Mantra] .s-file")
    fc.value.set_files(os.path.join(tmp, "650 (29).xlsx")); page.wait_for_timeout(400)
    mt = page.inner_text("[data-slot=Mantra]")
    check("650 (29).xlsx" in mt and "dropped on this box" in mt, "a click on the Mantra box chooses the 650 file: " + mt.replace("\n", " ")[:120])
    page.evaluate("""() => { const dt = new DataTransfer(); dt.items.add(new File(["a,b\\n1,2"], "odd.csv", { type: "text/csv" }));
      document.querySelector("[data-slot=GCC]").dispatchEvent(new DragEvent("drop", { dataTransfer: dt, bubbles: true, cancelable: true })); }""")
    page.wait_for_timeout(400)
    gt = page.inner_text("[data-slot=GCC]")
    check("odd.csv" in gt and "check it's the right file" in gt, "a file dropped on the GCC box goes there, with a warning when its headers don't match")
    page.click("[data-slot=GCC] [data-clear='odd.csv']"); page.wait_for_timeout(200)
    page.set_input_files("#dropInput", os.path.join(tmp, "CDF_List_Export.xlsx")); page.wait_for_timeout(400)
    check(page.inner_text("#importBtn") == "Replace Mantra, GCC, MANTRAREV", "the button names the sheets: " + page.inner_text("#importBtn"))
    page.click("#importBtn"); page.wait_for_selector("#importConfirm:not(.hidden)")
    check("will be replaced" in page.inner_text("#importConfirmText") and "New sheets: Mantra, MANTRAREV" in page.inner_text("#importConfirmText"), "confirm: GCC replaced, Mantra and MANTRAREV new: " + page.inner_text("#importConfirmText")[:150])
    page.click("#importYes"); page.wait_for_selector("#importMsg .note.ok", timeout=20000)
    check(len(sheet(page, "GCC")) == len(NEW_GCC) and sheet(page, "MANTRAREV") is not None, "GCC replaced (%d rows), MANTRAREV created" % len(sheet(page, "GCC")))
    check(sheet(page, "UPDATE LOG") is None, "no UPDATE LOG sheet is added by the import")
    check(page.evaluate("window.__xmlParts.some(p => p.xml.indexOf('urn:cdf-updater:imports:v1') >= 0)"), "when each sheet was replaced is kept inside the workbook")
    page.wait_for_timeout(300)
    up = page.inner_text("[data-slot=Mantra] .s-upd")
    check(re.search(r"Replaced \d\d/\d\d/\d{4} \d\d:\d\d", up) and "data to " + ddmmyyyy(TODAY - 2) in up, "under the box: when it was replaced and its newest date (Due Date left out): " + up)
    page.click("#tab-status"); page.wait_for_timeout(300)
    check("SHEETS 3/3" in page.inner_text("#sheetsLamp"), "all three sheets there now")
    dc = page.inner_text("[data-card=docs]")
    check(re.search(r"1\s*revisions released", dc) and re.search(r"1\s*revisions in work", dc), "solution documents of the open CDFs: 1 released, 1 in work: " + dc.replace("\n", " ")[:140])
    check("330-11-11-J0911" in dc and "no revision released" in dc, "the one in work is listed: J0911, no revision released yet")

    print("\nSUMMARY — sheet CDF SUMMARY as cards")
    page.click("#tab-summary"); page.wait_for_selector("#sumView .dcard.sv", timeout=20000)
    titles = page.eval_on_selector_all("#sumView .dcard.sv h3", "e => e.map(x => x.textContent)")
    check(titles == ["CDF STATUS", "WHERE ARE THE CDFS", "DANA PENDING TABLE", "REJECTED"], "every table is a card, side by side ones too: " + str(titles))
    check("4 tables" in page.inner_text("#sumMeta"), "the meta line counts the tables: " + page.inner_text("#sumMeta"))
    check("REJECTED CDFS 1" in page.inner_text("#sv-t3").replace("\n", " "), "summary lines rounded (1.0000000000000002 → 1)")
    check("column R" in page.inner_text("#sv-t0") and "column P" in page.inner_text("#sv-t1"), "the counts are linked to their columns in sheet CDF")
    page.click("#sv-t0 .sv-c.tap >> nth=0"); page.wait_for_selector("#list:not(.hidden)")
    check(page.inner_text("#listCount") == "3 CDFs", "a click on Open lists its 3 CDFs: " + page.inner_text("#listCount"))
    page.click("#listClose")
    page.click("#sv-t2 .sv-key"); page.wait_for_timeout(200)
    check(page.evaluate("window.__wb.active") == "CDF", "a CDF number in a table goes to its row")
    page.fill("#sumFilter", "tom bar"); page.wait_for_timeout(100)
    vis = page.eval_on_selector_all("#sumView .dcard.sv", "e => e.filter(x => x.offsetParent !== null).map(x => x.querySelector('h3').textContent)")
    check(vis == ["REJECTED"], "the filter leaves only what holds the text: " + str(vis))
    page.fill("#sumFilter", "")

    print("\nROWS — TSV, each row where you say")
    page.click("#tab-rows")
    tsv = "\n".join(["\t".join(["A330-1-902.A", "", "", "08/11/2026", "ROUTINE", "NEW REVISION OF 902", "OEM AVIONICS", "330-92-80-JF151-001-AV", "INSTL - X", "C0776", "-", "", "Ron Avi"]),
                     "\t".join(["A330-1-905", "", "", "09/11/2026", "URGENT", "THE NEW 905", "TAGS", "330-11-00-J0001", "INSTL - Y", "C0999", "A", "Dana Levi", ""])])
    page.fill("#rowsPaste", tsv); page.wait_for_selector(".row-item", timeout=10000)
    pv = page.inner_text("#rowsPreview")
    check("13 fields" in pv, "the 13 AI fields are recognised")
    check("New revision: under A330-1-902.-" in pv, "902.A → a new revision under 902")
    check("empty template row of A330-1-905" in pv, "905 → its empty template row")
    check("A330-1-905.-" in pv, "the key is written in full (A330-1-905 → A330-1-905.-)")
    # a place inside 901's merged rows is refused
    page.select_option("[data-mode='0']", "insert-above"); page.fill("[data-row='0']", "3"); page.dispatch_event("[data-row='0']", "change")
    check("inside the rows of A330-1-901.-" in page.inner_text("#rowsPreview") and page.is_disabled("#rowsWrite"), "inserting inside 901's merged rows is refused")
    page.select_option("[data-mode='0']", "insert-below"); page.fill("[data-row='0']", "4"); page.dispatch_event("[data-row='0']", "change")
    check(len(page.query_selector_all("[data-link]")) == 2, "a GCC link box under every row")
    page.fill("[data-link='0']", "gcc/902A"); page.dispatch_event("[data-link='0']", "change")
    check("isn't a web link" in page.inner_text("#rowsPreview") and page.is_disabled("#rowsWrite"), "a link that isn't http(s) is refused")
    page.fill("[data-link='0']", "https://gcc.example/cdf/902A"); page.dispatch_event("[data-link='0']", "change")
    check(not page.is_disabled("#rowsWrite"), "a web link is fine")
    before = sheet(page, "CDF")
    page.click("#rowsWrite"); page.wait_for_selector("#rowsMsg .note.ok", timeout=20000)
    after = sheet(page, "CDF")
    check(len(after) == len(before) + 1, "one row inserted (%d → %d)" % (len(before), len(after)))
    check(after[4][col("C")] == "A330-1-902.A" and after[4][col("D")] == "A" and after[4][col("H")] == "NEW REVISION OF 902", "902.A on row 5, under 902: C, D, H written")
    check(after[4][col("F")] == 46334, "the date is a real date (08/11/2026 → 46334)")
    r905 = [r for r in after if r[col("C")] == "A330-1-905.-"]
    check(len(r905) == 1 and r905[0][col("A")] == "A330-1-905.-ATT." and r905[0][col("G")] == "URGENT", "905 written over its template row; its links in A/B stay (empty cells don't erase)")
    check(page.evaluate("window.__wb.calls.some(c => /copyFrom .* Formats -> 5/.test(c))"), "the new row takes the look of the row above")
    lk = page.evaluate("window.__wb.get('CDF').links || {}")
    check(lk.get("5:2", {}).get("address") == "https://gcc.example/cdf/902A" and lk["5:2"].get("textToDisplay") == "A330-1-902.A", "902.A's CDF number (C5) links to GCC: " + json.dumps(lk))
    check(len(lk) == 1, "the row without a link gets none")
    check(sheet(page, "UPDATE LOG") is None, "still no UPDATE LOG sheet (the add-in never adds it)")
    page.click("#rowsUndo"); page.wait_for_selector("#rowsUndoMsg .note.ok", timeout=20000)
    back = sheet(page, "CDF")
    check(len(back) == len(before) and [r[col("C")] for r in back] == [r[col("C")] for r in before] and back[7][col("G")] == before[7][col("G")], "Undo: the row is gone and 905's template row is back as it was")

    print("\nMy Notes → column E")
    notes = {"kind": "my-notes/cdf-notes", "v": 1, "at": "2026-11-09T07:00:00.000Z", "notes": [
        {"refs": ["A330-1-901.-"], "text": "Ask Ann for the drawing", "state": "Open", "due": "2026-11-12", "subject": "RE: 901 drawing", "updated": 2},
        {"refs": ["A330-1-906.-"], "text": "Waiting for TIMAT", "state": "Open", "subject": "906", "updated": 1}]}
    packed = page.evaluate("CdfLink.encode(CdfLink.KIND_NOTES, %s)" % json.dumps({k: v for k, v in notes.items() if k not in ("kind", "v")}))
    page.click("#tab-setup"); page.fill("#linkPaste", packed); page.wait_for_function("document.getElementById('linkMsg').textContent.indexOf('Wrote') >= 0", timeout=20000, polling=200)
    cd = sheet(page, "CDF"); e901 = cd[1][col("E")]; e906 = [r for r in cd if r[col("C")] == "A330-1-906.-"][0][col("E")]
    check(e901.startswith("call the supplier\n✉ due 12/11: Ask Ann for the drawing"), "901: own text stays, the note is added: " + repr(e901))
    check(e906 == "✉ Waiting for TIMAT (906)", "906: " + repr(e906))
    page.fill("#linkPaste", packed); page.wait_for_function("document.getElementById('linkMsg').textContent.indexOf('already up to date') >= 0", timeout=20000, polling=200)
    check(sheet(page, "CDF")[1][col("E")] == e901, "pasting again changes nothing")
    page.click("#linkCopy"); page.wait_for_timeout(200)
    st = page.evaluate("CdfLink.decode(window.__clip, CdfLink.KIND_STATUS)")
    row901 = [r for r in st["data"]["rows"] if r[0] == "A330-1-901.-"][0]
    check(st["ok"] and row901[13] == "call the supplier", "the status for My Notes carries E without the ✉ lines (no echo back to Outlook)")

    print("\nLOOK — repaint still works")
    page.click("#tab-look"); page.click("#repaintBtn"); page.wait_for_selector("#repaintMsg .look-report", timeout=20000)
    cfs = page.evaluate("window.__wb.get('CDF').cfs.map(r => r.custom.rule.formula)")
    check(cfs[:3] == ['=TRIM($R2)="Released"', '=TRIM($R2)="Cancelled"', "=TRUE"], "the three rules of 110.7")
    page.close()

    print("\nAn empty workbook")
    page = browser.new_page(viewport={"width": 380, "height": 900})
    page.on("pageerror", lambda e: (print("   [pageerror]", e), errors.append(str(e))))
    page.route("**/*", lambda route: route.fulfill(path=os.path.join(HERE, "mock-office.js"), content_type="application/javascript") if "appsforoffice" in route.request.url else
               (route.fulfill(path=os.path.join(ADDIN, route.request.url[len("https://addin.test/"):].split("?")[0] or "taskpane.html")) if route.request.url.startswith("https://addin.test/") and os.path.isfile(os.path.join(ADDIN, route.request.url[len("https://addin.test/"):].split("?")[0] or "taskpane.html")) else route.abort()))
    page.add_init_script("window.__fixtures={}; window.__initialSheets=[]; window.__xmlParts=[];")
    page.goto("https://addin.test/taskpane.html"); page.wait_for_selector("#dash .dempty", timeout=20000); page.wait_for_timeout(500)
    check("No CDF sheet to read yet" in page.inner_text("#dash"), "says to import the CDF file")
    check("SHEETS 0/3" in page.inner_text("#sheetsLamp"), "lamp: 0/3")
    browser.close()

print("\n" + ("ALL OK" if not errors else "%d FAILED:\n  " % len(errors) + "\n  ".join(errors)))
sys.exit(1 if errors else 0)
