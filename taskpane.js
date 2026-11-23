/* CDF Updater 2 — task pane (Office.js glue + UI). No algorithm runs here: the pane replaces sheets, adds rows by hand,
   shows and presents the status, paints the sheet's look and links with My Notes. */
(function () {
  "use strict";

  var K = window.CdfCore, S = window.CdfStats, G = window.CdfCharts, D = window.CdfDash, R = window.CdfRows, N = window.CdfNotes, L = window.CdfLink, V = window.CdfSummary;
  var READ_BUDGET = 150000;
  var LS_RULES = "cdfUpdater.localRules.v2", THEME_KEY = "cdfUpdater.theme.v1", PRESENT_KEY = "cdfUpdater.present.v2";
  var NS_LOOKS = "urn:cdf-updater:looks:v1", NS_TREND = "urn:cdf-updater:trend:v2", NS_STATIONS = "urn:cdf-updater:stations:v1", NS_NOTES = "urn:cdf-updater:mynotes:v1", NS_IMPORTS = "urn:cdf-updater:imports:v1";

  var state = {
    inExcel: false, canImport: false, busy: false, page: "status",
    rules: null, rulesText: "", source: "github", localName: null,
    model: null, lists: {}, data: null, readAt: 0, trend: [], stations: {},
    files: [], importDone: {},                 // [{ file, target, sheet, rows, why }]
    rowsItems: null, rowsLayout: null, rowsUndo: null,
    looks: [], lookValues: null, lookCols: null, lookDraft: null,
    myNotes: null, notesPlan: null, statusPack: null,
    sum: null, sumLists: {}, sumAt: 0
  };
  var $ = function (id) { return document.getElementById(id); };

  // ---------------- startup ----------------
  function start(info) {
    state.inExcel = !!(info && window.Office && info.host === Office.HostType.Excel);
    setLamp("excelLamp", state.inExcel ? "EXCEL" : "NO EXCEL", state.inExcel ? "ok" : "warn", state.inExcel ? "Connected to this workbook" : "Opened outside Excel — workbook actions are off");
    if (state.inExcel) state.canImport = Office.context.requirements.isSetSupported("ExcelApi", "1.13");
    else showBanner("Opened outside Excel: you can check the settings file; everything else works inside Excel.", false);
    applyTheme();
    wire();
    restoreLocalRules();
    loadRules().then(loadStores).then(function () {
      if (!state.inExcel) { renderDashEmpty("Open the add-in inside Excel to see the status."); return; }
      watchCdf();
      return refreshDash();
    });
  }
  if (window.Office && Office.onReady) Office.onReady(start);
  else document.addEventListener("DOMContentLoaded", function () { start(null); });

  function wire() {
    Array.prototype.forEach.call(document.querySelectorAll(".ecp .key"), function (k, i, all) {
      k.addEventListener("click", function () { showPage(k.getAttribute("data-page")); });
      k.addEventListener("keydown", function (ev) {
        var d = ev.key === "ArrowRight" ? 1 : ev.key === "ArrowLeft" ? -1 : 0;
        if (!d) return;
        var n = all[(i + d + all.length) % all.length]; n.focus(); showPage(n.getAttribute("data-page"));
      });
    });
    // STATUS
    $("dashRefresh").addEventListener("click", function () { refreshDash(true); });
    $("dashPresent").addEventListener("click", present);
    $("dashDeck").addEventListener("click", makeDeck);
    $("dashCopy").addEventListener("click", copySummary);
    $("dash").addEventListener("click", onDashClick);
    $("dash").addEventListener("keydown", function (ev) { if (ev.key === "Enter" || ev.key === " ") onDashClick(ev); });
    $("glance").addEventListener("click", onDashClick);
    // SUMMARY
    $("sumRefresh").addEventListener("click", function () { refreshSummary(true); });
    $("sumFilter").addEventListener("input", filterSummary);
    $("sumView").addEventListener("click", onSumClick);
    $("sumView").addEventListener("keydown", function (ev) { if (ev.key === "Enter" || ev.key === " ") onSumClick(ev); });
    // IMPORT
    var drop = $("drop"), input = $("dropInput");
    input.addEventListener("change", function () { takeFiles(input.files); input.value = ""; });
    drop.addEventListener("dragover", function (ev) { ev.preventDefault(); drop.classList.add("drag"); });
    drop.addEventListener("dragleave", function () { drop.classList.remove("drag"); });
    drop.addEventListener("drop", function (ev) { ev.preventDefault(); drop.classList.remove("drag"); takeFiles(ev.dataTransfer.files); });
    $("slotInput").addEventListener("change", function () { var i = $("slotInput"); if (state.slotTarget) takeFiles(i.files, state.slotTarget); i.value = ""; });
    drop.addEventListener("keydown", function (ev) { if (ev.key === "Enter" || ev.key === " ") { ev.preventDefault(); input.click(); } });
    $("importBtn").addEventListener("click", askImport);
    $("importYes").addEventListener("click", doImport);
    $("importNo").addEventListener("click", function () { $("importConfirm").classList.add("hidden"); refreshButtons(); });
    // ROWS
    $("rowsPaste").addEventListener("input", function () { clearTimeout(state.rowsTimer); state.rowsTimer = setTimeout(readRows, 250); });
    $("rowsLayout").addEventListener("change", function () { $("rowsStart").classList.toggle("hidden", $("rowsLayout").value !== "from"); readRows(); });
    $("rowsStart").innerHTML = Array.apply(null, Array(34)).map(function (x, i) { return '<option value="' + i + '"' + (i === 2 ? " selected" : "") + ">from " + K.letter(i) + "</option>"; }).join("");
    $("rowsStart").addEventListener("change", readRows);
    $("rowsWrite").addEventListener("click", writeRows);
    $("rowsUndo").addEventListener("click", undoRows);
    // LOOK
    $("repaintBtn").addEventListener("click", function () { repaintNow({}, "repaintMsg", "Sheet look ✓"); });
    $("lookCopy").addEventListener("click", copyLookRequest);
    $("openChatGpt").addEventListener("click", function () { openUrl("https://chatgpt.com/"); });
    $("openClaudeLook").addEventListener("click", function () { openUrl("https://claude.ai/new"); });
    $("lookPaste").addEventListener("paste", onLookPaste);
    $("lookPaste").addEventListener("input", function () { clearTimeout(state.lookTimer); state.lookTimer = setTimeout(readLookAnswer, 400); });
    // SETUP
    $("linkCopy").addEventListener("click", function () { copyStatusForMyNotes($("linkCopy")); });
    $("linkPaste").addEventListener("input", function () { var t = $("linkPaste").value; if (t.trim()) takeMyNotes(t); });
    $("reloadRules").addEventListener("click", function () { loadRules(); });
    $("localRules").addEventListener("change", onLocalRules);
    $("useGithub").addEventListener("click", function () { state.source = "github"; try { localStorage.removeItem(LS_RULES); } catch (e) { /* ignore */ } loadRules(); });
    Array.prototype.forEach.call(document.querySelectorAll("[data-theme-choice]"), function (b) {
      b.addEventListener("click", function () { try { localStorage.setItem(THEME_KEY, b.getAttribute("data-theme-choice")); } catch (e) { /* ignore */ } applyTheme(); renderDash(); });
    });
    $("rulesPill").addEventListener("click", function () { showPage("setup"); });
    // lists and lookup
    $("listClose").addEventListener("click", closeList);
    $("lookupOpen").addEventListener("click", openLookup);
    $("lookupClose").addEventListener("click", closeLookup);
    $("lookupInput").addEventListener("keydown", function (ev) { if (ev.key === "Enter") doLookup($("lookupInput").value); if (ev.key === "Escape") closeLookup(); });
    document.addEventListener("keydown", function (ev) {
      if ((ev.ctrlKey || ev.metaKey) && (ev.key === "f" || ev.key === "F")) { ev.preventDefault(); openLookup(); }
      if (ev.key === "Escape") { closeList(); closeLookup(); }
    });
    window.addEventListener("storage", function (ev) { if (L && ev.key === L.LS_NOTES) adoptSharedNotes(); });
    window.addEventListener("focus", adoptSharedNotes);
  }

  // ---------------- rules.txt ----------------
  function restoreLocalRules() {
    try { var x = JSON.parse(localStorage.getItem(LS_RULES) || "null"); if (x && x.text) { state.source = "local"; state.localName = x.name; state.rulesText = x.text; } } catch (e) { /* ignore */ }
  }
  function loadRules() {
    var p = state.source === "local" ? Promise.resolve(state.rulesText)
      : fetch("rules.txt?v=" + Date.now(), { cache: "no-store" }).then(function (r) { if (!r.ok) throw new Error("Couldn't read rules.txt (HTTP " + r.status + ")"); return r.text(); });
    return p.then(applyRulesText).catch(function (e) {
      applyRulesText("");   // the defaults still work
      setPill("NO RULES", "warn"); $("rulesErrors").innerHTML = '<div class="err">' + esc(errText(e)) + " — using the built-in settings.</div>";
    });
  }
  function applyRulesText(text) {
    state.rulesText = text;
    state.rules = K.parseRules(text);
    var n = state.rules.errors.length;
    setPill(n ? "RULES " + n + " ERR" : "RULES OK", n ? "bad" : "ok");
    $("rulesErrors").innerHTML = state.rules.errors.map(function (e) { return '<div class="err">' + esc(e) + "</div>"; }).join("");
    $("rulesText").textContent = text;
    $("rulesSource").textContent = state.source === "local" ? "From your computer: " + (state.localName || "rules.txt") : "From GitHub (rules.txt next to the add-in)";
    $("useGithub").classList.toggle("hidden", state.source !== "local");
    renderSlots(); refreshButtons();
  }
  function onLocalRules(ev) {
    var f = ev.target.files[0]; ev.target.value = "";
    if (!f) return;
    f.text().then(function (text) {
      state.source = "local"; state.localName = f.name;
      try { localStorage.setItem(LS_RULES, JSON.stringify({ name: f.name, text: text })); } catch (e) { /* ignore */ }
      applyRulesText(text); refreshDash(true);
    });
  }

  // ---------------- workbook stores (custom XML parts travel with the file) ----------------
  function storesSupported() { try { return state.inExcel && Office.context.requirements.isSetSupported("ExcelApi", "1.5"); } catch (e) { return false; } }
  function encodeStore(ns, obj) { return '<cdfUpdater xmlns="' + ns + '">' + encodeURIComponent(JSON.stringify(obj)) + "</cdfUpdater>"; }
  function decodeStore(xml) { var m = /<cdfUpdater[^>]*>([^<]*)<\/cdfUpdater>/.exec(String(xml || "")); if (!m) return null; try { return JSON.parse(decodeURIComponent(m[1])); } catch (e) { return null; } }
  function readStore(ns) {
    if (!storesSupported()) return Promise.resolve(null);
    return Excel.run(function (ctx) {
      var part = ctx.workbook.customXmlParts.getByNamespace(ns).getOnlyItemOrNullObject();
      return ctx.sync().then(function () {
        if (part.isNullObject) return null;
        var x = part.getXml();
        return ctx.sync().then(function () { return decodeStore(x.value); });
      });
    }).catch(function () { return null; });
  }
  function writeStore(ns, obj) {
    if (!storesSupported()) return Promise.reject(new Error("This Excel version can't keep data in the workbook (ExcelApi 1.5 needed)"));
    return Excel.run(function (ctx) {
      var parts = ctx.workbook.customXmlParts.getByNamespace(ns);
      parts.load("items");
      return ctx.sync().then(function () {
        var xml = encodeStore(ns, obj);
        if (parts.items.length) { parts.items[0].setXml(xml); for (var i = 1; i < parts.items.length; i++) parts.items[i].delete(); }
        else ctx.workbook.customXmlParts.add(xml);
        return ctx.sync().then(function () { return true; });
      });
    });
  }
  function loadStores() {
    return Promise.all([readStore(NS_LOOKS), readStore(NS_TREND), readStore(NS_STATIONS), readStore(NS_NOTES), readStore(NS_IMPORTS)]).then(function (x) {
      state.imported = (x[4] && x[4].map) || {};
      loadLooks(x[0]);
      state.trend = (x[1] && x[1].list) || [];
      state.stations = (x[2] && x[2].map) || {};
      state.myNotes = (x[3] && x[3].data) || null;
      renderLink(); adoptSharedNotes();
    });
  }

  // ---------------- reading sheets ----------------
  function roundTrips(ctx, jobs, budget) {
    var groups = [], cur = [], n = 0;
    jobs.forEach(function (j) { if (cur.length && n + j.cells > budget) { groups.push(cur); cur = []; n = 0; } cur.push(j); n += j.cells; });
    if (cur.length) groups.push(cur);
    return groups.reduce(function (p, g) {
      return p.then(function () { g.forEach(function (j) { j.queue(); }); return ctx.sync().then(function () { g.forEach(function (j) { if (j.done) j.done(); }); }); });
    }, Promise.resolve());
  }
  // names → { sheets: { name: { values } | null } } — whole used ranges, from A1, in a few round trips
  function readSheets(ctx, names, opt) {
    var wss = ctx.workbook.worksheets, sh = {}, used = {}, out = { sheets: {} };
    names.forEach(function (n) { sh[n] = wss.getItemOrNullObject(n); sh[n].load("name"); });
    return ctx.sync().then(function () {
      names.forEach(function (n) { if (!sh[n].isNullObject) { used[n] = sh[n].getUsedRangeOrNullObject(true); used[n].load("rowIndex,rowCount,columnIndex,columnCount"); } });
      return ctx.sync();
    }).then(function () {
      var jobs = [];
      names.forEach(function (n) {
        if (sh[n].isNullObject) { out.sheets[n] = null; return; }
        var u = used[n], d = out.sheets[n] = { values: [] };
        if (u.isNullObject) return;
        var rows = u.rowIndex + u.rowCount, cols = u.columnIndex + u.columnCount, CH = Math.max(1, Math.min(rows, Math.floor(READ_BUDGET / Math.max(1, cols))));
        for (var r = 0; r < rows; r += CH) (function (st, count) {
          var rg;
          jobs.push({ cells: count * cols, queue: function () { rg = sh[n].getRangeByIndexes(st, 0, count, cols); rg.load("values"); }, done: function () { Array.prototype.push.apply(d.values, rg.values); } });
        })(r, Math.min(CH, rows - r));
      });
      return roundTrips(ctx, jobs, READ_BUDGET);
    }).then(function () { return out; });
  }
  function readAll() {
    return Excel.run(function (ctx) {
      var r = state.rules || K.DEFAULT_RULES, names = ["CDF"];
      r.imports.forEach(function (i) { if (names.indexOf(i.target) < 0) names.push(i.target); });
      if (r.log && names.indexOf(r.log) < 0) names.push(r.log);
      return readSheets(ctx, names).then(function (rd) {
        var byTarget = {};
        names.forEach(function (n) { byTarget[n] = rd.sheets[n] ? rd.sheets[n].values : null; });
        state.data = { cdf: byTarget.CDF, gcc: byTarget.GCC, rev: byTarget.MANTRAREV, all: byTarget, imported: K.lastImports(r.log ? byTarget[r.log] : null) };
        state.readAt = Date.now();
        return state.data;
      });
    });
  }

  // ---------------- STATUS: the dashboard ----------------
  function refreshDash(manual) {
    if (!state.inExcel) return Promise.resolve();
    if (manual) setStatus("Reading the sheets…");
    return readAll().then(function (d) {
      var now = new Date(), today = K.serialOf(now);
      var tr = N.trackStations(state.stations, d.cdf || [], today);
      state.stations = tr.map;
      var r = state.rules || K.parseRules("");
      state.model = S.build({ cdf: d.cdf, gcc: d.gcc, rev: d.rev }, { now: now, exclude: r.exclude, stuckDays: r.stuckDays, targets: r.targets, history: state.trend });
      var m = state.model;
      m.stuck = N.stuck(state.stations, today, r.stuckDays);
      renderDash(); renderGlance(); renderSheetsLamp(); refreshButtons();
      // keep one point a day for the trend, and the stations for "days at station"
      if (!m.empty && storesSupported()) {
        var snap = S.snapshot(m), old = state.trend.filter(function (s) { return s.d === snap.d; })[0];
        if (!old || old.open !== snap.open || old.released !== snap.released) {
          state.trend = S.mergeHistory(state.trend, snap);
          writeStore(NS_TREND, { list: state.trend }).catch(function () { /* the trend waits for the next time */ });
        }
        if (tr.changed) writeStore(NS_STATIONS, { map: state.stations, at: now.toISOString() }).catch(function () { /* kept for this session */ });
      }
      refreshStatusPack(); renderLink();
      $("trendInfo").textContent = state.trend.length ? "The trend line has " + state.trend.length + " day" + (state.trend.length === 1 ? "" : "s") + " in this workbook, since " + K.ddmmyyyy(state.trend[0].d) + "." : "";
      if (manual) dashMsg("ok", "Read " + K.ddmmyyyy(today) + " " + hhmm(now) + ".");
    }).catch(function (e) { dashMsg("bad", errText(e)); }).then(function () { if (manual) clearStatus(); });
  }
  function dashMsg(cls, text) { var m = $("dashMsg"); m.className = "msg dash-msg " + (cls || ""); m.textContent = text; if (cls === "ok") setTimeout(function () { if (m.textContent === text) m.textContent = ""; }, 4000); }
  function renderDashEmpty(text) { $("dash").innerHTML = '<div class="dempty">' + esc(text) + "</div>"; $("glance").innerHTML = '<div class="ecam-empty">' + esc(text) + "</div>"; }
  function renderDash() {
    if (!state.model) return;
    var r = D.render(state.model, { mode: "pane" });
    state.lists = r.lists;
    $("dash").innerHTML = r.html + stuckCard(state.model);
  }
  function stuckCard(m) {
    var st = m.stuck || [];
    if (!st.length) return "";
    state.lists["stuck"] = { title: "At the same station more than " + (state.rules || K.DEFAULT_RULES).stuckDays + " days", keys: st.map(function (x) { return x.key; }) };
    return '<section class="dcard" data-card="stuck"><div class="dhead"><h3>Not moving</h3><span class="dsub">' + st.length + " CDFs</span></div>" +
      '<p class="dnote">At the same station more than ' + (state.rules || K.DEFAULT_RULES).stuckDays + " days (counted since the add-in first saw them there; marked ≥ when it may be longer).</p>" +
      '<ol class="oldest">' + st.slice(0, 6).map(function (x) { return '<li class="tap" data-key="' + esc(x.key) + '" tabindex="0" role="button"><span class="ok">' + esc(x.key.replace(/^A330-1-/, "")) + '</span><span class="os">' + esc(D.stationShort(x.station)) + '</span><span class="od">' + (x.exact ? "" : "≥ ") + x.days + " d</span></li>"; }).join("") + "</ol>" +
      (st.length > 6 ? '<button class="btn ghost small" type="button" data-list="stuck">All ' + st.length + "</button>" : "") + "</section>";
  }
  // the dark display on top: the three numbers that matter, and the stations
  function renderGlance() {
    var m = state.model, g = $("glance");
    if (!m || m.empty) { g.innerHTML = '<div class="ecam-empty">No CDF sheet yet — <b>IMPORT</b> brings it in.</div>'; return; }
    var st = m.status, dw = S.deltaWeek(m), max = Math.max.apply(null, m.stations.map(function (s) { return s.n; }).concat([1]));
    var delta = function (n, good) { return n ? '<span class="d ' + (n * good > 0 ? "up" : "down") + '">' + (n > 0 ? "▲" : "▼") + Math.abs(n) + "</span>" : ""; };
    g.innerHTML = '<div class="ecam-top"><span>CDF STATUS ' + K.ddmm(m.today) + "</span><span class=\"since\">" + (dw ? "VS " + K.ddmm(dw.since) : "TODAY") + "</span></div>" +
      '<div class="gl-nums">' +
      '<div class="metric tap" data-list="status:open" tabindex="0" role="button"><span class="v amber">' + st.open + "</span>" + (dw ? delta(dw.open, -1) : "") + '<span class="k">OPEN</span></div>' +
      '<div class="metric tap" data-list="status:released" tabindex="0" role="button"><span class="v green">' + st.pct + '<small>%</small></span><span class="k">RELEASED · ' + st.released + "</span></div>" +
      '<div class="metric tap" data-list="week:closed" tabindex="0" role="button"><span class="v cyan">' + m.week.closed + '</span><span class="k">REL. 7 DAYS</span></div>' +
      '<div class="metric tap" data-list="week:opened" tabindex="0" role="button"><span class="v">' + m.week.opened + '</span><span class="k">NEW 7 DAYS</span></div></div>' +
      '<div class="ecam-sep"></div><div class="gl-bars">' + m.stations.map(function (s) {
        return '<div class="gb tap" data-list="station:' + esc(s.name) + '" tabindex="0" role="button"><span class="gn">' + esc(D.stationShort(s.name).toUpperCase().slice(0, 10)) + '</span><span class="gt"><i style="width:' + Math.max(2, Math.round(100 * s.n / max)) + '%"></i></span><span class="gv">' + s.n + "</span></div>";
      }).join("") + "</div>";
  }
  function renderSheetsLamp() {
    var d = state.data || {}, have = ["cdf", "gcc", "rev"].filter(function (k) { return d[k] && d[k].length > 1; }).length, m = state.model;
    var age = m && m.docs && m.docs.latest != null ? m.today - m.docs.latest : 0, old = age > 7;
    setLamp("sheetsLamp", have === 3 ? (old ? "REV OLD" : "SHEETS 3/3") : "SHEETS " + have + "/3", have === 3 && !old ? "ok" : have ? "warn" : "bad",
      ["CDF", "GCC", "MANTRAREV"].map(function (n, i) { return n + ": " + (d[["cdf", "gcc", "rev"][i]] && d[["cdf", "gcc", "rev"][i]].length > 1 ? "there" : "missing"); }).join(" · ") + (old ? " · the newest release in MANTRAREV is " + age + " days old" : ""));
  }
  function onDashClick(ev) {
    var t = ev.target.closest ? ev.target.closest("[data-list],[data-key]") : null;
    if (!t) return;
    ev.preventDefault();
    if (t.getAttribute("data-key")) { selectCdf(t.getAttribute("data-key")); return; }
    var l = state.lists[t.getAttribute("data-list")];
    if (l) openList(l.title, l.keys);
  }

  // the CDFs behind a number
  function openList(title, keys) {
    var cd = state.data && state.data.cdf ? S.cdfsOf(state.data.cdf) : [], by = {};
    cd.forEach(function (x) { by[x.key] = x; });
    $("listTitle").textContent = title; $("listCount").textContent = keys.length + " CDF" + (keys.length === 1 ? "" : "s");
    $("listBody").innerHTML = keys.length ? '<ul class="lst">' + keys.map(function (k) {
      var x = by[k] || {};
      return '<li><button class="lst-row" type="button" data-key="' + esc(k) + '"><span class="lk">' + esc(k) + '</span><span class="ls">' + esc(D.stationShort(x.station || "")) + "</span>" +
        '<span class="lr">' + esc(x.status || "") + "</span></button></li>";
    }).join("") + "</ul>" : '<p class="hint">None.</p>';
    Array.prototype.forEach.call($("listBody").querySelectorAll("[data-key]"), function (b) { b.addEventListener("click", function () { selectCdf(b.getAttribute("data-key")); }); });
    $("list").classList.remove("hidden");
  }
  function closeList() { $("list").classList.add("hidden"); }
  // select a CDF's row in sheet CDF
  function selectCdf(key) {
    if (!state.inExcel) return;
    var cd = state.data && state.data.cdf ? K.buildBlocks(state.data.cdf).filter(function (b) { return K.sameCdf(b.key, key); })[0] : null;
    if (!cd) return;
    Excel.run(function (ctx) {
      var ws = ctx.workbook.worksheets.getItem("CDF");
      ws.activate();
      ws.getRange("A" + (cd.anchor + 1) + ":AH" + (cd.rows[cd.rows.length - 1] + 1)).select();
      return ctx.sync();
    }).catch(function () { /* the sheet moved on */ });
  }

  // full screen, for a meeting: the same cards, larger, one at a time with the arrow keys
  function present() {
    if (!state.model) return;
    try { localStorage.setItem(PRESENT_KEY, JSON.stringify({ model: state.model, theme: document.documentElement.getAttribute("data-theme") })); } catch (e) { dashMsg("bad", "Couldn't hand the data to the present view (" + errText(e) + ")."); return; }
    var url = new URL("present.html?v=" + version(), location.href).href;
    if (window.Office && Office.context && Office.context.ui && Office.context.ui.displayDialogAsync) {
      Office.context.ui.displayDialogAsync(url, { height: 92, width: 92, displayInIframe: false }, function (r) {
        if (r.status !== Office.AsyncResultStatus.Succeeded) window.open(url, "_blank");
      });
    } else window.open(url, "_blank");
  }
  // a PowerPoint deck: PptxGenJS loads only when it's needed
  function loadScript(src) {
    return new Promise(function (ok, bad) { var s = document.createElement("script"); s.src = src; s.onload = ok; s.onerror = function () { bad(new Error("Couldn't load " + src + " — is lib/ uploaded to GitHub?")); }; document.head.appendChild(s); });
  }
  function makeDeck() {
    if (!state.model) return;
    var btn = $("dashDeck"); btn.disabled = true; setStatus("Building the PowerPoint deck…");
    (window.PptxGenJS ? Promise.resolve() : loadScript("lib/pptxgen.bundle.js?v=" + version())).then(function () {
      var pptx = window.CdfDeck.build(state.model, window.PptxGenJS);
      return pptx.write({ outputType: "blob" });
    }).then(function (blob) {
      saveBlob(blob, window.CdfDeck.fileName(state.model));
      dashMsg("ok", "PowerPoint deck saved: " + window.CdfDeck.fileName(state.model) + " (native charts — edit them in PowerPoint).");
    }).catch(function (e) { dashMsg("bad", errText(e)); }).then(function () { btn.disabled = false; clearStatus(); });
  }
  function copySummary() {
    if (!state.model) return;
    copyRich(D.summaryHtml(state.model), D.summaryText(state.model)).then(function (ok) { dashMsg(ok ? "ok" : "bad", ok ? "Copied — paste it in an email or Teams (Ctrl+V)." : "Copy failed here."); });
  }
  // a change in sheet CDF: the status follows a few seconds later
  function watchCdf() {
    if (!Office.context.requirements.isSetSupported("ExcelApi", "1.7")) return;
    Excel.run(function (ctx) {
      var ws = ctx.workbook.worksheets.getItemOrNullObject("CDF");
      ws.load("name");
      return ctx.sync().then(function () {
        if (ws.isNullObject || !ws.onChanged) return;
        ws.onChanged.add(function () {
          clearTimeout(state.watchTimer);
          state.watchTimer = setTimeout(function () { if (!state.busy) refreshDash(); }, 2500);
        });
        return ctx.sync();
      });
    }).catch(function () { /* Refresh still works */ });
  }

  // ================= the look of sheet CDF (110.6–110.7) and the looks from ChatGPT / Claude =================
  var LK = window.CdfLook;
  var BORDER_SIDES = ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideVertical", "InsideHorizontal"];
  // values: sheet CDF as it is now (in memory after an update, or read). opt.fresh: rows this update added (they
  // fit their text); opt.drop: formulas of looks just removed (their rules go too). 4 round trips + row heights.
  function lookPass(ctx, ws, values, opt) {
    opt = opt || {};
    var plan = LK.lookPlan(values, K);
    if (!plan) return Promise.resolve({ skipped: "Sheet CDF has nothing below the header row." });
    var canAreas = Office.context.requirements.isSetSupported("ExcelApi", "1.9");
    var looks = state.looks || [], lookFormulas = looks.map(function (l) { return l.formula; }).concat(opt.drop || []);
    var last = plan.last, rep = { last: last, cont: plan.cont.length, looks: looks.length, rules: 0 };
    var nf = ws.getRange(plan.body), cfs = ws.getRange().conditionalFormats, before = [], after = {};
    var rowJob = function (r, list) {
      var rg = ws.getRange(r + ":" + r);
      return { cells: 20, queue: function () { rg.load("rowHidden"); rg.format.load("rowHeight"); }, done: function () { list(r, rg); } };
    };
    var reads = [{ cells: (last - 1) * 34, queue: function () { nf.load("numberFormat"); cfs.load("items/type"); } }];
    for (var r = 2; r <= last; r++) reads.push(rowJob(r, function (row, rg) { before.push({ row: row, height: rg.format.rowHeight, hidden: !!rg.rowHidden }); }));
    var custom = [];
    return roundTrips(ctx, reads, READ_BUDGET).then(function () {
      cfs.items.forEach(function (cf) {
        if (cf.type !== "Custom") return;
        var x = { cf: cf, rule: cf.custom.rule, ranges: canAreas ? cf.getRanges() : null };
        x.rule.load("formula"); if (x.ranges) x.ranges.load("areaCount");
        custom.push(x);
      });
      return ctx.sync();
    }).then(function () {
      var found = custom.map(function (x) { return { formula: x.rule.formula, areas: x.ranges ? x.ranges.areaCount : 1 }; });
      rep.found = LK.describeExisting(found, lookFormulas);
      // 110.7 and the saved looks are written again from scratch: one clean rule each; the user's own rules stay
      custom.forEach(function (x) { if (LK.ruleKind(x.rule.formula, lookFormulas)) x.cf.delete(); });
      // 110.6
      var f = ws.getRange(plan.all).format;
      f.font.name = "Calibri"; f.font.size = 10; f.verticalAlignment = "Center"; f.wrapText = true;
      BORDER_SIDES.forEach(function (sd) { var b = f.borders.getItem(sd); b.style = "Continuous"; b.weight = "Thin"; b.color = LK.BLACK; });
      var p = ws.getRange(plan.purple).format; p.fill.color = LK.PURPLE; p.font.color = LK.WHITE;
      var hidden = LK.hiddenFormatCells(nf.numberFormat, 2);
      hidden.forEach(function (h) { ws.getRange(h.a1).numberFormat = [["General"]]; });
      rep.unhidden = hidden.length;
      // 110.7 + looks, top rule first
      var added = [];
      function addRule(addresses, formula, look) {
        addresses.forEach(function (addr) {
          var targets = addr.indexOf(",") < 0 ? [ws.getRange(addr)] : canAreas ? [ws.getRanges(addr)] : addr.split(",").map(function (a) { return ws.getRange(a.trim()); });
          targets.forEach(function (t) {
            var cf = t.conditionalFormats.add(Excel.ConditionalFormatType.custom), c = cf.custom;
            c.rule.formula = formula;
            if (look.fill) c.format.fill.color = look.fill;
            if (look.font) c.format.font.color = look.font;
            if (look.bold) c.format.font.bold = true;
            if (look.italic) c.format.font.italic = true;
            if (look.top) { var b = c.format.borders.getItem("EdgeTop"); b.style = "Continuous"; b.color = look.top; }
            added.push(cf);
          });
        });
      }
      plan.rules.forEach(function (ru) { addRule(ru.areas, ru.formula, ru); });
      looks.forEach(function (lk) { var a = LK.paintAreas(lk.spans, last); addRule([a.join(", ")], lk.formula, lk); });
      added.forEach(function (cf, i) { cf.priority = i; });
      rep.rules = added.length; rep.split = !canAreas;
      // heights: autofit the rows a filter doesn't hide, then read them again
      var runs = LK.visibleRuns(before);
      runs.forEach(function (run) { ws.getRange(run[0] + ":" + run[1]).format.autofitRows(); });
      rep.fitted = before.filter(function (b) { return !b.hidden; }).length; rep.hiddenRows = before.length - rep.fitted;
      var jobs = before.filter(function (b) { return !b.hidden; }).map(function (b) { return rowJob(b.row, function (row, rg) { after[row] = rg.format.rowHeight; }); });
      if (!jobs.length) return ctx.sync();
      return roundTrips(ctx, jobs, READ_BUDGET);
    }).then(function () {
      var back = LK.restoreHeights(before, after, opt.fresh);
      back.forEach(function (b) { ws.getRange(b.row + ":" + b.row).format.rowHeight = b.height; });
      rep.kept = back.length;
      rep.grew = before.filter(function (b) { return !b.hidden && typeof after[b.row] === "number" && after[b.row] > b.height + 0.01; }).length;
      return ctx.sync();
    }).then(function () { return rep; });
  }
  // the look again, from the sheet as it is now (LOOK page, after NEW ROWS, after an undo)
  function repaintIn(ctx, opt) {
    return readSheets(ctx, ["CDF"], {}).then(function (rd) {
      if (!rd.sheets.CDF) throw new Error("Sheet \"CDF\" was not found in this workbook");
      return lookPass(ctx, ctx.workbook.worksheets.getItem("CDF"), rd.sheets.CDF.values, opt).then(function (rep) { state.lookValues = rd.sheets.CDF.values; return rep; });
    });
  }
  function lookReportHtml(rep, title, brief) {
    if (!rep) return "";
    if (brief && !rep.error && !rep.skipped) return '<div class="note ok look-report"><b>' + esc(brief) + "</b> " + esc(title) + "</div>";
    if (rep.error) return '<div class="note bad look-report">' + esc(title || "Sheet look") + ": couldn't be repainted — " + esc(rep.error) + ". Try <b>Repaint CDF</b> on the LOOK page again.</div>";
    if (rep.skipped) return '<div class="note look-report">' + esc(rep.skipped) + "</div>";
    var fd = rep.found || { pieces: {}, broken: [] }, li = [];
    li.push("Rows 2–" + rep.last + ": Calibri 10, black lines around every cell, S:X purple, Middle Align, Wrap Text.");
    var fixed = (fd.broken || []).map(function (k) { return (k === "cont" ? "solution-row lines" : capFirst(k)) + " (" + fd[k] + " rule" + (fd[k] === 1 ? "" : "s") + ", " + fd.pieces[k] + " pieces)"; });
    li.push("Released green, Cancelled red and the line over " + plural(rep.cont, "solution row") + (fixed.length ? " — were split: " + esc(fixed.join(", ")) + "; now one rule each." : " — one rule each."));
    if (rep.looks) li.push(plural(rep.looks, "saved look") + " painted again.");
    if (rep.grew) li.push(plural(rep.grew, "row") + " grew to fit " + (rep.grew === 1 ? "its" : "their") + " text" + (rep.kept ? "; " + plural(rep.kept, "taller row") + " kept " + (rep.kept === 1 ? "its" : "their") + " height" : "") + ".");
    else if (rep.kept) li.push(plural(rep.kept, "taller row") + " kept " + (rep.kept === 1 ? "its" : "their") + " height (rows never shrink, so text in merged cells stays visible).");
    if (rep.hiddenRows) li.push(plural(rep.hiddenRows, "row") + " hidden by a filter kept " + (rep.hiddenRows === 1 ? "its" : "their") + " height.");
    if (rep.unhidden) li.push(plural(rep.unhidden, "cell") + " with the hidden format ;;; now " + (rep.unhidden === 1 ? "shows its" : "show their") + " value.");
    if (fd.ownRules) li.push("Kept " + plural(fd.ownRules, "conditional format") + " of your own.");
    if (rep.split) li.push("This Excel can't put one rule on several ranges, so the solution-row line is " + plural(rep.rules - 2 - rep.looks, "rule") + ".");
    return '<div class="note ok look-report"><b>' + esc(title || "Sheet look ✓") + "</b><ul>" + li.map(function (x) { return "<li>" + x + "</li>"; }).join("") + "</ul></div>";
  }
  // Check changes: what Apply will paint (nothing is read for it)
  function renderLookPage() {
    var list = state.looks || [];
    $("looksCount").textContent = list.length ? String(list.length) : "";
    var box = $("savedLooks");
    if (!list.length) { box.innerHTML = '<p class="hint">None yet. A look you add is kept in this workbook and painted again with every Repaint.</p>'; return; }
    var today = LK.todaySerial(new Date()), vals = state.lookValues;
    box.innerHTML = list.map(function (lk, i) {
      var n = vals ? LK.matchRows(lk, vals, today).length : null;
      return '<div class="look-item">' + swatchHtml(lk) + '<div><div class="ln">' + esc(lk.name) + '</div><div class="lw">' + esc(lk.say) + "</div></div>" +
        '<div class="lm">Cells ' + esc(lk.cells) + (n == null ? "" : " · <b>" + plural(n, "row") + "</b> now") + "</div>" +
        '<div class="acts"><button class="btn ghost small" type="button" data-unlook="' + i + '">Remove</button></div></div>';
    }).join("");
    Array.prototype.forEach.call(box.querySelectorAll("[data-unlook]"), function (b) { b.addEventListener("click", function () { removeLook(+b.getAttribute("data-unlook")); }); });
  }
  function swatchHtml(lk) {
    var st = "background:" + (lk.fill || "#FFFFFF") + ";color:" + (lk.font || "#000000") + (lk.bold ? ";font-weight:800" : ";font-weight:400") + (lk.italic ? ";font-style:italic" : "");
    return '<span class="swatch" style="' + st + '" aria-hidden="true">Aa</span>';
  }
  function lookMsg(id, cls, html) { var m = $(id); m.className = "msg " + (cls || ""); m.innerHTML = html; }
  // the columns as the prompt names them: header and kind (text / number / date) from the first 200 rows — no values
  function lookColumns() {
    if (state.lookCols) return Promise.resolve(state.lookCols);
    if (!state.inExcel) return Promise.resolve(null);
    return Excel.run(function (ctx) {
      var ws = ctx.workbook.worksheets.getItemOrNullObject("CDF");
      return ctx.sync().then(function () {
        if (ws.isNullObject) throw new Error("Sheet \"CDF\" was not found in this workbook");
        var rg = ws.getRange("A1:AH201"); rg.load("values,numberFormat");
        return ctx.sync().then(function () {
          var kinds = LK.columnKinds(rg.values.slice(1), rg.numberFormat.slice(1), isDateFormat);
          state.lookCols = rg.values[0].map(function (h, c) { return { letter: LK.letter(c), header: String(h == null ? "" : h), kind: kinds[c] }; });
          return state.lookCols;
        });
      });
    });
  }
  function lookHeaders() { return (state.lookCols || []).map(function (c) { return c.header; }); }
  function copyLookRequest() {
    var ask = $("lookAsk").value.trim();
    if (!ask) { lookMsg("lookCopyMsg", "bad", "Write first what should stand out."); $("lookAsk").focus(); return; }
    lookColumns().then(function (cols) {
      var prompt = LK.buildPrompt(ask, cols || []);
      return copyText(prompt).then(function (ok) {
        lookMsg("lookCopyMsg", ok ? "ok" : "bad", ok ? "Request copied ✓ — paste it in ChatGPT or Claude (" + (cols ? cols.length + " column names, no values" : "column names unknown outside Excel") + ")." : "Copy failed here — click the button again.");
      });
    }).catch(function (e) { lookMsg("lookCopyMsg", "bad", esc(errText(e))); });
  }
  function onLookPaste(ev) {
    var text = ev && ev.clipboardData && ev.clipboardData.getData("text/plain");
    if (text) { ev.preventDefault(); $("lookPaste").value = text; }
    setTimeout(readLookAnswer, 0);
  }
  function readLookAnswer() {
    var text = $("lookPaste").value, box = $("lookPreview");
    if (!text.trim()) { box.innerHTML = ""; state.lookDraft = null; return; }
    var cols = state.inExcel ? lookColumns().catch(function () { return null; }) : Promise.resolve(null);
    cols.then(function () {
      var res = LK.parseLooks(text, lookHeaders(), K);
      if (res.error) { state.lookDraft = null; box.innerHTML = '<div class="note bad">' + esc(res.error) + (res.kind === "invalid" || res.kind === "nojson" ? ' Ask in the same chat: "Answer again with only the json block, in the format I gave."' : "") + "</div>"; return; }
      state.lookDraft = res.looks;
      return lookValues().then(function (vals) {
        var today = LK.todaySerial(new Date()), A = K;
        box.innerHTML = res.looks.map(function (lk) {
          var rows = vals ? LK.matchRows(lk, vals, today) : null, keys = matchKeys(rows, vals);
          return '<div class="look-item">' + swatchHtml(lk) + '<div><div class="ln">' + esc(lk.name) + '</div><div class="lw">' + esc(lk.say) + "</div></div>" +
            '<div class="lm">Cells ' + esc(lk.cells) + (rows ? " · <b>" + plural(rows.length, "row") + "</b> now" + (keys.list.length ? ": " + esc(keys.list.join(", ")) + (keys.more ? " …" : "") : "") : "") + "</div>" +
            '<details><summary>Excel rule</summary><code>' + esc(lk.formula) + "</code></details></div>";
        }).join("") +
          '<div class="row" style="margin-top:10px"><button id="lookAdd" class="btn primary" type="button">' + (res.looks.length > 1 ? "Add " + res.looks.length + " looks" : "Add to the sheet") + '</button><button id="lookDiscard" class="btn ghost" type="button">Discard</button></div>' +
          (state.inExcel ? "" : '<p class="hint">Outside Excel: the rule is read, but rows can only be counted and painted inside Excel.</p>') + '<div id="lookAddMsg"></div>';
        $("lookAdd").disabled = !state.inExcel || state.busy;
        $("lookAdd").addEventListener("click", addLooks);
        $("lookDiscard").addEventListener("click", function () { $("lookPaste").value = ""; box.innerHTML = ""; state.lookDraft = null; });
      });
    });
  }
  // the CDFs behind matched rows, each once (a solution row counts as its CDF): "901, 902 …"
  function matchKeys(rows, vals) {
    if (!rows || !vals) return { list: [], more: false };
    var A = K, of = {}, seen = {}, list = [];
    A.buildBlocks(vals).forEach(function (b) { b.rows.forEach(function (r) { of[r + 1] = b.key; }); });
    rows.forEach(function (r) { var k = of[r] || "row " + r; if (!seen[k]) { seen[k] = 1; list.push(String(k).replace(/^A330-1-/, "").replace(/\.-$/, "")); } });
    return { list: list.slice(0, 5), more: list.length > 5 };
  }
  function lookValues() {
    if (!state.inExcel) return Promise.resolve(null);
    return Excel.run(function (ctx) { return readSheets(ctx, ["CDF"], {}).then(function (rd) { state.lookValues = rd.sheets.CDF ? rd.sheets.CDF.values : null; return state.lookValues; }); }).catch(function () { return null; });
  }
  function saveLooks() { return writeStore(NS_LOOKS, { looks: (state.looks || []).map(LK.toStored), savedAt: new Date().toISOString() }); }
  function repaintNow(opt, msgId, title, brief) {
    if (state.busy) return Promise.resolve();
    state.busy = true; refreshButtons(); setStatus("Painting sheet CDF…");
    return Excel.run(function (ctx) { return repaintIn(ctx, opt); }).then(function (rep) {
      $(msgId).innerHTML = lookReportHtml(rep, typeof brief === "function" ? brief() : title, typeof brief === "function" ? title : null); renderLookPage(); return rep;
    }).catch(function (e) { $(msgId).innerHTML = '<div class="note bad">' + esc(errText(e)) + "</div>"; })
      .then(function (rep) { state.busy = false; clearStatus(); refreshButtons(); return rep; });
  }
  function addLooks() {
    var draft = state.lookDraft;
    if (!draft || !draft.length) return;
    var stamp = new Date().toISOString();
    var prev = (state.looks || []).slice();
    state.looks = prev.concat(draft.map(function (l) { l.addedAt = stamp; return l; }));
    saveLooks().then(function () {
      return repaintNow({}, "lookAddMsg", (draft.length > 1 ? draft.length + " looks" : "Look") + " added ✓", function () {
        var today = LK.todaySerial(new Date()), n = state.lookValues ? draft.reduce(function (t, l) { return t + LK.matchRows(l, state.lookValues, today).length; }, 0) : null;
        return (n == null ? "" : plural(n, "row") + " painted now. ") + "Kept in this workbook and painted again with every Repaint.";
      });
    }).then(function () {
      state.lookDraft = null; $("lookPaste").value = "";
      setTimeout(function () { var m = $("lookAddMsg"); $("lookPreview").innerHTML = ""; if (m) $("lookPreview").appendChild(m); }, 0);
    }).catch(function (e) { state.looks = prev; $("lookAddMsg").innerHTML = '<div class="note bad">Not added — ' + esc(errText(e)) + "</div>"; });
  }
  function removeLook(i) {
    var prev = (state.looks || []).slice(), gone = prev[i];
    if (!gone) return;
    state.looks = prev.filter(function (x, j) { return j !== i; });
    saveLooks().then(function () { return repaintNow({ drop: [gone.formula] }, "repaintMsg", "“" + gone.name + "” removed ✓", function () { return "Its rule is off the sheet."; }); })
      .catch(function (e) { state.looks = prev; $("repaintMsg").innerHTML = '<div class="note bad">Not removed — ' + esc(errText(e)) + "</div>"; renderLookPage(); });
  }
  function loadLooks(x) {
    state.looks = ((x && x.looks) || []).map(function (s) { return LK.fromStored(s, lookHeaders()); }).filter(Boolean);
    renderLookPage();
  }


  // ================= IMPORT: replace Mantra (650), GCC, MANTRAREV — each file goes to the sheet its headers belong to =================
  function imports() { return (state.rules || K.parseRules("")).imports; }
  // forTarget: the file was dropped on (or chosen in) that sheet's box — it goes there, whatever its headers say
  function takeFiles(list, forTarget) {
    var files = Array.prototype.slice.call(list || []);
    if (!files.length) return;
    if (forTarget) files = files.slice(0, 1);
    $("importConfirm").classList.add("hidden"); $("importMsg").innerHTML = "";
    setStatus("Reading the files…");
    Promise.all(files.map(detectFile)).then(function (found) {
      found.forEach(function (f) {
        if (forTarget && !f.error) {
          if (f.target === forTarget && !f.weak) f.why = "its headers, dropped on this box";
          else { f.why = f.target && f.target !== forTarget ? "your choice — its headers look like " + f.target + ", check it's the right file" : "your choice — its headers didn't match, check it's the right file"; f.weak = true; }
          f.target = forTarget;
        }
        state.files = state.files.filter(function (x) { return x.file.name !== f.file.name && (!f.target || x.target !== f.target); });
        state.files.push(f);
      });
      renderSlots(); refreshButtons();
    }).then(clearStatus, function (e) { clearStatus(); $("importMsg").innerHTML = '<div class="note bad">' + esc(errText(e)) + "</div>"; });
  }
  // the first rows of every sheet in the file → the IMPORT whose HEADERS are in one of them
  function detectFile(file) {
    var ext = (file.name.split(".").pop() || "").toLowerCase(), out = { file: file, target: null, sheet: null, rows: 0, why: "" };
    return file.arrayBuffer().then(function (buf) {
      out.buf = buf;
      if (ext === "csv") {
        var rows = K.parseDelimited(decodeText(buf).split(/\r?\n/).slice(0, 6).join("\n"));
        var imp = K.importFor(rows, imports());
        out.rows = decodeText(buf).split(/\r?\n/).filter(function (l) { return l.trim(); }).length;
        return finish(imp, null);
      }
      if (ext !== "xlsx" && ext !== "xlsm") { out.error = "Not an Excel or CSV file (." + ext + ")"; return out; }
      return xlsxHeads(buf).then(function (sheets) {
        var hit = null;
        sheets.forEach(function (s) { if (!hit) { var imp = K.importFor(s.rows, imports()); if (imp) hit = { imp: imp, s: s }; } });
        if (hit) { out.rows = hit.s.count; return finish(hit.imp, hit.s.name); }
        out.sheets = sheets.map(function (s) { return s.name; });
        out.rows = sheets.length ? sheets[0].count : 0;
        return finish(null, sheets.length ? sheets[0].name : null);
      });
    }).catch(function (e) { out.error = errText(e); return out; });
    function finish(imp, sheet) {
      out.sheet = sheet;
      if (imp) { out.target = imp.target; out.why = "its headers (" + imp.headers.slice(0, 2).join(", ") + ")"; return out; }
      var byName = K.fileNameFor(file.name, imports());
      if (byName) { out.target = byName.target; out.why = "its name — the headers didn't match, check it's the right file"; out.weak = true; }
      else out.why = "neither its headers nor its name match a sheet — choose one";
      return out;
    }
  }
  function xlsxHeads(buf) {
    var zip;
    function rid(el) { return el.getAttribute("r:id") || el.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id"); }
    return JSZip.loadAsync(buf).then(function (z) {
      zip = z;
      var wb = zip.file("xl/workbook.xml"), rels = zip.file("xl/_rels/workbook.xml.rels"), ss = zip.file("xl/sharedStrings.xml");
      if (!wb || !rels) throw new Error("This doesn't look like a valid Excel workbook");
      return Promise.all([wb.async("string"), rels.async("string"), ss ? ss.async("string") : Promise.resolve("")]);
    }).then(function (x) {
      var dp = new DOMParser(), sheets = dp.parseFromString(x[0], "application/xml").getElementsByTagNameNS("*", "sheet"), rl = dp.parseFromString(x[1], "application/xml").getElementsByTagNameNS("*", "Relationship");
      var strings = [];
      if (x[2]) { var sis = dp.parseFromString(x[2], "application/xml").getElementsByTagNameNS("*", "si"); for (var k = 0; k < sis.length; k++) strings.push(sis[k].textContent); }
      var list = [];
      for (var i = 0; i < sheets.length; i++) {
        var id = rid(sheets[i]), target = null;
        for (var j = 0; j < rl.length; j++) if (rl[j].getAttribute("Id") === id) target = rl[j].getAttribute("Target");
        if (!target) continue;
        var p = target.replace(/^\//, ""); if (p.indexOf("xl/") !== 0) p = "xl/" + p;
        list.push({ name: sheets[i].getAttribute("name"), path: p });
      }
      return Promise.all(list.map(function (s) {
        var f = zip.file(s.path);
        return (f ? f.async("string") : Promise.resolve("")).then(function (xml) {
          var rows = [], re = /<row\b[^>]*>([\s\S]*?)<\/row>/g, m;
          while (rows.length < 5 && (m = re.exec(xml))) {
            var row = [];
            (m[1].match(/<c\b[\s\S]*?(?:\/>|<\/c>)/g) || []).forEach(function (c) {
              var t = /\bt="(\w+)"/.exec(c), v = /<v>([\s\S]*?)<\/v>/.exec(c), is = /<t[^>]*>([\s\S]*?)<\/t>/.exec(c);
              row.push(t && t[1] === "s" && v ? strings[+v[1]] : t && t[1] === "inlineStr" && is ? is[1] : v ? v[1] : "");
            });
            rows.push(row);
          }
          return { name: s.name, rows: rows, count: (xml.match(/<row\b/g) || []).length };
        });
      }));
    });
  }
  function renderSlots() {
    var box = $("slots");
    if (!box) return;
    var d = state.data || {}, today = K.serialOf(new Date()), html = imports().map(function (imp) {
      var f = state.files.filter(function (x) { return x.target === imp.target; })[0], cur = d.all && d.all[imp.target], done = state.importDone[imp.target];
      var when = (state.imported || {})[imp.target.toUpperCase()] || (d.imported || {})[imp.target.toUpperCase()], newest = cur ? K.newestDate(cur, today) : null;
      var upd = cur ? '<div class="s-upd" title="Replaced: the last time the add-in replaced it (UPDATE LOG). Data to: the newest date in its Date columns.">' + ([when ? "Replaced <b>" + esc(when) + "</b>" : "", newest != null ? (when ? "data" : "Data") + " to <b>" + esc(K.ddmmyyyy(newest)) + "</b>" : ""].filter(Boolean).join(" · ") || "No dates in it") + "</div>" : "";
      return '<div class="slot2' + (f ? " has" : "") + '" data-slot="' + esc(imp.target) + '" tabindex="0" role="button" aria-label="Drop or choose the file for sheet ' + esc(imp.target) + '"><div class="s-head"><b>' + esc(imp.target) + '</b><span class="s-now">' + (cur ? "in the workbook: " + Math.max(0, cur.length - 1) + " rows" : "not in the workbook yet — will be created") + "</span></div>" + upd +
        (f ? '<div class="s-file">' + esc(f.file.name) + (f.sheet ? " · sheet “" + esc(f.sheet) + "”" : "") + " · " + f.rows + " rows <button class=\"slot-clear\" type=\"button\" data-clear=\"" + esc(f.file.name) + '" title="Remove">✕</button></div><div class="s-why' + (f.weak ? " warn-text" : "") + '">Matched by ' + esc(f.why) + "</div>"
          : '<div class="s-file dim">' + (done ? esc(done) : "Drop the file here, or click to choose") + "</div>") + "</div>";
    }).join("");
    var loose = state.files.filter(function (x) { return !x.target; });
    html += loose.map(function (f) {
      return '<div class="slot2 bad"><div class="s-head"><b>' + esc(f.file.name) + '</b><button class="slot-clear" type="button" data-clear="' + esc(f.file.name) + '" title="Remove">✕</button></div>' +
        (f.error ? '<div class="s-why warn-text">' + esc(f.error) + "</div>" : '<div class="s-why">' + esc(f.why) + '</div><select class="inp small" data-pick="' + esc(f.file.name) + '"><option value="">Replace which sheet?</option>' +
          imports().map(function (i) { return '<option value="' + esc(i.target) + '">' + esc(i.target) + "</option>"; }).join("") + "</select>") + "</div>";
    }).join("");
    box.innerHTML = html;
    Array.prototype.forEach.call(box.querySelectorAll("[data-slot]"), function (el) {
      var t = el.getAttribute("data-slot");
      el.addEventListener("dragover", function (ev) { ev.preventDefault(); ev.stopPropagation(); el.classList.add("drag"); });
      el.addEventListener("dragleave", function () { el.classList.remove("drag"); });
      el.addEventListener("drop", function (ev) { ev.preventDefault(); ev.stopPropagation(); el.classList.remove("drag"); takeFiles(ev.dataTransfer.files, t); });
      function pick(ev) { if (ev.target.closest && ev.target.closest("button,select")) return; state.slotTarget = t; var i = $("slotInput"); i.value = ""; i.click(); }
      el.addEventListener("click", pick);
      el.addEventListener("keydown", function (ev) { if ((ev.key === "Enter" || ev.key === " ") && ev.target === el) { ev.preventDefault(); pick(ev); } });
    });
    Array.prototype.forEach.call(box.querySelectorAll("[data-clear]"), function (b) { b.addEventListener("click", function () { state.files = state.files.filter(function (x) { return x.file.name !== b.getAttribute("data-clear"); }); renderSlots(); refreshButtons(); }); });
    Array.prototype.forEach.call(box.querySelectorAll("[data-pick]"), function (s) {
      s.addEventListener("change", function () {
        var f = state.files.filter(function (x) { return x.file.name === s.getAttribute("data-pick"); })[0];
        if (!f || !s.value) return;
        state.files = state.files.filter(function (x) { return x === f || x.target !== s.value; });
        f.target = s.value; f.why = "your choice"; f.weak = true; renderSlots(); refreshButtons();
      });
    });
  }
  function chosen() {
    var order = imports().map(function (i) { return i.target; });
    return state.files.filter(function (x) { return x.target && !x.error; }).sort(function (a, b) { return order.indexOf(a.target) - order.indexOf(b.target); });
  }
  function askImport() {
    var c = chosen();
    if (!c.length) return;
    var d = state.data || {}, rep = c.filter(function (f) { return d.all && d.all[f.target]; }).map(function (f) { return f.target; }), add = c.filter(function (f) { return !(d.all && d.all[f.target]); }).map(function (f) { return f.target; });
    var t = "";
    if (rep.length) t += "Sheet" + (rep.length > 1 ? "s" : "") + " <b>" + rep.map(esc).join(", ") + "</b> will be replaced by the file" + (rep.length > 1 ? "s" : "") + ". Formulas on other sheets that point to " + (rep.length > 1 ? "them" : "it") + " turn into #REF!.";
    if (add.length) t += (t ? "<br>" : "") + "New sheet" + (add.length > 1 ? "s" : "") + ": <b>" + add.map(esc).join(", ") + "</b>.";
    if (rep.indexOf("CDF") >= 0) t += '<br><span class="warn-text">Replacing CDF drops anything typed in the current CDF sheet since that file was made (rows added here, notes in E).</span>';
    var weak = c.filter(function (f) { return f.weak; });
    if (weak.length) t += '<br><span class="warn-text">⚠ ' + weak.map(function (f) { return esc(f.file.name) + " → " + esc(f.target) + " (" + esc(f.why) + ")"; }).join("<br>⚠ ") + "</span>";
    $("importConfirmText").innerHTML = t + "<br>Continue?";
    $("importConfirm").classList.remove("hidden"); $("importBtn").disabled = true;
  }
  function doImport() {
    $("importConfirm").classList.add("hidden");
    var c = chosen(), done = [], failed = [];
    state.busy = true; refreshButtons();
    c.reduce(function (p, f) {
      return p.then(function () {
        setStatus("Replacing sheet " + f.target + "…");
        return importOne(f).then(function () { done.push(f); state.importDone[f.target] = "Replaced ✓ " + hhmm(new Date()) + " — " + f.file.name; state.files = state.files.filter(function (x) { return x !== f; }); },
          function (e) { failed.push({ f: f, e: e }); });
      });
    }, Promise.resolve()).then(function () {
      // when each sheet was replaced: kept inside the workbook (no sheet for it)
      if (!done.length) return;
      state.imported = state.imported || {};
      done.forEach(function (f) { state.imported[f.target.toUpperCase()] = stamp(); });
      return writeStore(NS_IMPORTS, { map: state.imported }).catch(function () { /* shown for this session */ });
    }).then(function () {
      $("importMsg").innerHTML = (done.length ? '<div class="note ok">Replaced ✓ ' + done.map(function (f) { return esc(f.target); }).join(", ") + ".</div>" : "") +
        failed.map(function (x) { return '<div class="note bad">' + esc(x.f.target) + ": " + esc(errText(x.e)) + " — the sheet was left as it was.</div>"; }).join("");
      state.busy = false; clearStatus(); refreshButtons();
      return refreshDash();
    }).then(renderSlots);
  }
  function importOne(f) {
    var ext = (f.file.name.split(".").pop() || "").toLowerCase();
    if (ext === "csv") {
      var rows = K.parseDelimited(decodeText(f.buf));
      if (!rows.length) return Promise.reject(new Error("The CSV file is empty"));
      var width = rows.reduce(function (m, r) { return Math.max(m, r.length); }, 0);
      rows = rows.map(function (r) { while (r.length < width) r.push(""); return r; });
      return replaceSheet(f.target, function (ctx, rel) {
        var ws = ctx.workbook.worksheets.add("~new_" + rnd());
        if (rel) ws.position = rel.position + 1;
        for (var i = 0; i < rows.length; i += 2000) ws.getRangeByIndexes(i, 0, Math.min(2000, rows.length - i), width).values = rows.slice(i, i + 2000);
        ws.load("id");
        return ctx.sync().then(function () { return ws.id; });
      });
    }
    if (!state.canImport) return Promise.reject(new Error("This Excel version can't copy a sheet from a file (Microsoft 365 needed)"));
    var b64 = toBase64(f.buf);
    return replaceSheet(f.target, function (ctx, rel) {
      var opts = { sheetNamesToInsert: [f.sheet] };
      if (rel) { opts.positionType = Excel.WorksheetPositionType.after; opts.relativeTo = rel.tmpName; } else opts.positionType = Excel.WorksheetPositionType.end;
      var res = ctx.workbook.insertWorksheetsFromBase64(b64, opts);
      return ctx.sync().then(function () { if (!res.value || !res.value.length) throw new Error("Excel didn't add the sheet"); return res.value[0]; });
    });
  }
  // rename the old sheet → insert the new one next to it → delete the old → rename the new; back to the old one on failure
  function replaceSheet(target, insertFn) {
    return Excel.run(function (ctx) {
      var sheets = ctx.workbook.worksheets, old = sheets.getItemOrNullObject(target), tmp = null, newId = null;
      old.load("name,position");
      return ctx.sync().then(function () {
        if (old.isNullObject) return null;
        tmp = "~old_" + rnd(); old.name = tmp;
        return ctx.sync().then(function () { return { tmpName: tmp, position: old.position }; });
      }).then(function (rel) { return insertFn(ctx, rel); }).then(function (id) {
        newId = id;
        var ns = sheets.getItem(id);
        if (tmp) sheets.getItem(tmp).delete();
        ns.name = target;
        return ctx.sync();
      }).catch(function (err) {
        if (!tmp) throw err;
        var o = sheets.getItemOrNullObject(tmp), n = newId ? sheets.getItemOrNullObject(newId) : null;
        o.load("name"); if (n) n.load("name");
        return ctx.sync().then(function () { if (!o.isNullObject) { if (n && !n.isNullObject) n.delete(); o.name = target; } return ctx.sync(); })
          .catch(function () { /* best effort */ }).then(function () { throw err; });
      });
    });
  }

  // ================= ROWS: rows by hand, from TSV =================
  function rowsCfg() { return (state.rules || K.parseRules("")).rows; }
  function readRows() {
    var text = $("rowsPaste").value;
    if (!text.trim()) { state.rowsItems = null; $("rowsPreview").innerHTML = ""; refreshButtons(); return; }
    var p = R.parse(text);
    if (p.error) { state.rowsItems = null; $("rowsPreview").innerHTML = '<div class="note bad">' + esc(p.error) + "</div>"; refreshButtons(); return; }
    var need = state.inExcel && (!state.data || Date.now() - state.readAt > 30000) ? readAll() : Promise.resolve(state.data);
    need.then(function (d) {
      var sheet = (d && d.cdf) || [], cfg = rowsCfg(), mode = $("rowsLayout").value;
      var lay = R.layoutOf(p.rows, sheet[cfg.headerRow - 1] || [], mode === "auto" ? {} : { mode: mode, start: +$("rowsStart").value });
      state.rowsLayout = lay;
      state.rowsItems = p.rows.slice(lay.header ? 1 : 0).filter(function (r) { return r.some(function (x) { return K.trim(x); }); }).map(function (r) {
        var c = R.cellsOf(r, lay, cfg), s = R.suggest(c.cells, sheet, { startRow: cfg.startRow });
        return { cells: c.cells, dropped: c.dropped, mode: s.mode, row: s.row, why: s.why };
      });
      renderRowsPreview();
    }).catch(function (e) { $("rowsPreview").innerHTML = '<div class="note bad">' + esc(errText(e)) + "</div>"; });
  }
  // a block's inner rows are merged with its anchor: inserting between them would split the merge
  function insideBlock(sheet, r, mode) {
    var at = mode === "insert-above" ? r : r + 1;
    var hit = K.buildBlocks(sheet).filter(function (b) { return b.rows.length > 1 && at > b.anchor + 1 && at <= b.rows[b.rows.length - 1] + 1; })[0];
    return hit || null;
  }
  function renderRowsPreview() {
    var items = state.rowsItems, sheet = (state.data && state.data.cdf) || [], cfg = rowsCfg(), heads = sheet[cfg.headerRow - 1] || [];
    if (!items || !items.length) { $("rowsPreview").innerHTML = '<div class="note bad">No rows found in what was pasted.</div>'; refreshButtons(); return; }
    var pl = R.plan(items, sheet, { startRow: cfg.startRow }), problems = pl.errors.slice();
    items.forEach(function (it, i) {
      var b = it.mode !== "overwrite" ? insideBlock(sheet, +it.row, it.mode) : null;
      if (b) problems.push("Row " + (i + 1) + ": that place is inside the rows of " + b.key + " (merged cells) — insert above row " + (b.anchor + 1) + " or below row " + (b.rows[b.rows.length - 1] + 1));
      if (!it.cells.some(function (c) { return K.trim(c.value) !== ""; })) problems.push("Row " + (i + 1) + " has nothing to write");
      if (it.link && !/^https?:\/\/\S+$/i.test(it.link)) problems.push("Row " + (i + 1) + ": the GCC link isn't a web link");
      if (it.link && !R.keyOf(it.cells) && !K.trim((sheet[+it.row - 1] || [])[K.col("C")])) problems.push("Row " + (i + 1) + ": a GCC link needs a CDF number in C");
    });
    state.rowsPlan = problems.length ? null : pl;
    var head = '<div class="rows-why">' + esc(state.rowsLayout.why) + ". Empty cells in what you paste leave the sheet's cell as it is.</div>";
    $("rowsPreview").innerHTML = head + items.map(function (it, i) {
      var key = R.keyOf(it.cells) || "row " + (i + 1) + " (no CDF number)", r = +it.row, over = it.mode === "overwrite", cur = sheet[r - 1] || [];
      var cells = it.cells.filter(function (c) { return K.trim(c.value) !== ""; }).map(function (c) {
        var old = over ? K.trim(cur[c.col]) : "", nv = typeof c.value === "number" && (c.col === K.col("F") || c.col === K.col("Q")) ? K.ddmmyyyy(c.value) : String(c.value);
        var oldTxt = over && old && typeof cur[c.col] === "number" && (c.col === K.col("F") || c.col === K.col("Q")) ? K.ddmmyyyy(cur[c.col]) : old;
        return "<tr><td class=\"rc\">" + K.letter(c.col) + '</td><td class="rh">' + esc(String(heads[c.col] || "").replace(/\s+/g, " ").slice(0, 26)) + "</td><td>" + (over && oldTxt && oldTxt !== nv ? '<s class="old">' + esc(oldTxt.slice(0, 60)) + "</s> " : "") + esc(nv.slice(0, 120)) + "</td></tr>";
      }).join("");
      return '<div class="row-item"><div class="ri-head"><b>' + esc(key) + "</b>" + (pl.rows[i] ? '<span class="ri-final">→ row ' + pl.rows[i] + "</span>" : "") + "</div>" +
        '<div class="ri-ctl"><select class="inp small" data-mode="' + i + '"><option value="insert-below"' + (it.mode === "insert-below" ? " selected" : "") + '>Insert below row</option><option value="insert-above"' + (it.mode === "insert-above" ? " selected" : "") + '>Insert above row</option><option value="overwrite"' + (over ? " selected" : "") + '>Write over row</option></select>' +
        '<input class="inp small num" type="number" min="' + cfg.startRow + '" value="' + esc(it.row) + '" data-row="' + i + '" aria-label="Row number"></div>' +
        '<div class="ri-now">Row ' + r + " now: " + esc(R.rowSays(sheet, r)) + (it.why ? ' <span class="dim">· suggested: ' + esc(it.why) + "</span>" : "") + "</div>" +
        '<label class="ri-link"><span>GCC link</span><input class="inp small" type="url" data-link="' + i + '" value="' + esc(it.link || "") + '" placeholder="Paste the link to this CDF in GCC" spellcheck="false" autocomplete="off"></label>' +
        (it.link && !/^https?:\/\/\S+$/i.test(it.link) ? '<div class="ri-now warn-text">That isn\'t a web link (it should start with http:// or https://).</div>' : "") +
        (it.dropped ? '<div class="ri-now warn-text">' + it.dropped + " value" + (it.dropped === 1 ? "" : "s") + " fall outside " + esc(cfg.firstCol + ":" + cfg.lastCol) + " and are left out.</div>" : "") +
        '<details class="ri-cells"><summary>' + it.cells.filter(function (c) { return K.trim(c.value) !== ""; }).length + " cells" + (over ? " (old → new)" : "") + "</summary><table>" + cells + "</table></details></div>";
    }).join("") + (problems.length ? '<div class="note bad">' + problems.map(esc).join("<br>") + "</div>" : "");
    Array.prototype.forEach.call($("rowsPreview").querySelectorAll("[data-mode]"), function (s) { s.addEventListener("change", function () { items[+s.getAttribute("data-mode")].mode = s.value; items[+s.getAttribute("data-mode")].why = ""; renderRowsPreview(); }); });
    Array.prototype.forEach.call($("rowsPreview").querySelectorAll("[data-link]"), function (s) { s.addEventListener("change", function () { items[+s.getAttribute("data-link")].link = K.trim(s.value); renderRowsPreview(); }); });
    Array.prototype.forEach.call($("rowsPreview").querySelectorAll("[data-row]"), function (s) { s.addEventListener("change", function () { items[+s.getAttribute("data-row")].row = Math.max(cfg.startRow, Math.floor(+s.value || cfg.startRow)); items[+s.getAttribute("data-row")].why = ""; renderRowsPreview(); }); });
    $("rowsWrite").textContent = "Write " + items.length + " row" + (items.length === 1 ? "" : "s") + " to sheet " + cfg.sheet;
    refreshButtons();
  }
  function writeRows() {
    var items = state.rowsItems, cfg = rowsCfg();
    if (!items || !state.rowsPlan || state.busy) return;
    var said = items.map(function (it) { return R.rowSays(state.data.cdf, +it.row); });
    state.busy = true; refreshButtons(); setStatus("Writing rows…");
    Excel.run(function (ctx) {
      return readSheets(ctx, [cfg.sheet]).then(function (rd) {
        var sheet = rd.sheets[cfg.sheet] ? rd.sheets[cfg.sheet].values : null;
        if (!sheet) throw new Error("Sheet \"" + cfg.sheet + "\" was not found");
        if (items.some(function (it, i) { return R.rowSays(sheet, +it.row) !== said[i]; })) { state.data.cdf = sheet; throw new Error("The sheet changed since the preview — check the rows again (the preview is now up to date)"); }
        var pl = R.plan(items, sheet, { startRow: cfg.startRow }), ws = ctx.workbook.worksheets.getItem(cfg.sheet), undo = { inserted: [], overwritten: [], at: new Date() }, log = [];
        // 1. write-overs (the sheet's row numbers as they are now): keep what was there for the undo
        var olds = [];
        pl.ops.filter(function (o) { return o.kind === "overwrite"; }).forEach(function (o) {
          o.cells.forEach(function (c) { if (K.trim(c.value) === "") return; var rg = ws.getCell(o.row - 1, c.col); rg.load("formulas,numberFormat"); olds.push({ o: o, c: c, rg: rg }); });
        });
        return ctx.sync().then(function () {
          olds.forEach(function (x) {
            var old = x.rg.formulas[0][0];
            undo.overwritten.push({ row: x.o.row, col: x.c.col, old: old });
            if (typeof x.c.value === "number" && /^(General|@)?$/i.test(String(x.rg.numberFormat[0][0] || ""))) x.rg.numberFormat = [["dd/mm/yyyy"]];
            x.rg.values = [[x.c.value]];
            log.push([stamp(), "ROW (write over)", cfg.sheet, String(pl.rows[x.o.item]), R.keyOf(x.o.cells), K.letter(x.c.col), String(old), String(x.c.value), ""]);
          });
          // 2. inserts, from the bottom up; each new row takes the look of the row above it
          pl.ops.filter(function (o) { return o.kind === "insert"; }).forEach(function (o) {
            var nr = ws.getRange(o.at + ":" + (o.at + o.count - 1));
            nr.insert(Excel.InsertShiftDirection.down);
            var src = o.at - 1 >= cfg.startRow ? o.at - 1 : o.at + o.count;
            ws.getRange(o.at + ":" + (o.at + o.count - 1)).copyFrom(ws.getRange(src + ":" + src), Excel.RangeCopyType.formats);
            o.rows.forEach(function (cells, j) {
              cells.forEach(function (c) {
                if (K.trim(c.value) === "") return;
                var rg = ws.getCell(o.at + j - 1, c.col);
                if (typeof c.value === "number") rg.numberFormat = [["dd/mm/yyyy"]];
                rg.values = [[c.value]];
              });
              log.push([stamp(), "ROW (insert)", cfg.sheet, String(pl.rows[o.items[j]]), R.keyOf(cells), "", "", cells.filter(function (c) { return K.trim(c.value) !== ""; }).length + " cells", ""]);
            });
          });
          // 3. the GCC link on the CDF number (C) of each row that has one
          undo.links = [];
          pl.rows.forEach(function (r, i) {
            var url = K.trim(items[i].link);
            if (!url) return;
            var rg = ws.getCell(r - 1, K.col("C")), key = R.keyOf(items[i].cells) || K.trim((sheet[+items[i].row - 1] || [])[K.col("C")]);
            rg.hyperlink = { address: url, textToDisplay: key, screenTip: "Open " + key + " in GCC" };
            if (items[i].mode === "overwrite") undo.links.push(r);
            log.push([stamp(), "GCC LINK", cfg.sheet, String(r), key, "C", "", url, ""]);
          });
          pl.rows.forEach(function (r, i) { if (items[i].mode !== "overwrite") undo.inserted.push(r); });
          undo.inserted.sort(function (a, b) { return b - a; });
          return ctx.sync();
        }).then(function () { return appendLog(ctx, log).catch(function () { /* the rows are written; the log is a convenience */ }); })
          .then(function () { state.rowsUndo = undo; return pl; });
      });
    }).then(function (pl) {
      $("rowsMsg").innerHTML = '<div class="note ok">Written ✓ — ' + pl.rows.map(function (r, i) { return esc(R.keyOf(items[i].cells) || "row") + " on row " + r; }).join(", ") + '.<div class="row" style="margin-top:6px"><button id="rowsGo" class="btn small" type="button">Show in the sheet</button><button id="rowsPaint" class="btn small" type="button">Repaint CDF look</button></div></div>';
      $("rowsGo").addEventListener("click", function () { goRow(cfg.sheet, Math.min.apply(null, pl.rows)); });
      $("rowsPaint").addEventListener("click", function () { repaintNow({ fresh: pl.rows }, "rowsMsg", "Sheet look ✓"); });
      $("rowsPaste").value = ""; state.rowsItems = null; $("rowsPreview").innerHTML = "";
      $("rowsUndoInfo").textContent = "Last write " + hhmm(state.rowsUndo.at) + ": " + state.rowsUndo.inserted.length + " inserted, " + state.rowsUndo.overwritten.length + " cells written over.";
    }).catch(function (e) { $("rowsMsg").innerHTML = '<div class="note bad">' + esc(errText(e)) + "</div>"; if (state.rowsItems) renderRowsPreview(); })
      .then(function () { state.busy = false; clearStatus(); refreshButtons(); return refreshDash(); });
  }
  function undoRows() {
    var u = state.rowsUndo, cfg = rowsCfg();
    if (!u || state.busy) return;
    state.busy = true; refreshButtons(); setStatus("Undoing…");
    Excel.run(function (ctx) {
      var ws = ctx.workbook.worksheets.getItem(cfg.sheet);
      u.inserted.forEach(function (r) { ws.getRange(r + ":" + r).delete(Excel.DeleteShiftDirection.up); });
      (u.links || []).forEach(function (r) { ws.getCell(r - 1, K.col("C")).clear(Excel.ClearApplyTo.hyperlinks); });
      u.overwritten.forEach(function (x) { ws.getCell(x.row - 1, x.col).formulas = [[x.old]]; });
      return ctx.sync().then(function () { return appendLog(ctx, [[stamp(), "UNDO ROWS", cfg.sheet, "", "", "", "", "", u.inserted.length + " rows removed, " + u.overwritten.length + " cells put back"]]).catch(function () { /* fine */ }); });
    }).then(function () { $("rowsUndoMsg").innerHTML = '<div class="note ok">Undone ✓</div>'; state.rowsUndo = null; $("rowsUndoInfo").textContent = "Nothing to undo."; },
      function (e) { $("rowsUndoMsg").innerHTML = '<div class="note bad">' + esc(errText(e)) + "</div>"; })
      .then(function () { state.busy = false; clearStatus(); refreshButtons(); return refreshDash(); });
  }
  function goRow(sheet, r) { Excel.run(function (ctx) { var ws = ctx.workbook.worksheets.getItem(sheet); ws.activate(); ws.getRange("A" + r + ":AH" + r).select(); return ctx.sync(); }).catch(function () { /* fine */ }); }
  // UPDATE LOG: one line per change — only when the workbook already has that sheet; the add-in never adds it
  function appendLog(ctx, rows) {
    var name = (state.rules || K.DEFAULT_RULES).log;
    if (!name || !rows.length) return Promise.resolve();
    var sheets = ctx.workbook.worksheets, ws = sheets.getItemOrNullObject(name);
    ws.load("name");
    return ctx.sync().then(function () {
      if (ws.isNullObject) return null;
      var fresh = false, used = ws.getUsedRangeOrNullObject(true);
      if (used) used.load("rowIndex,rowCount");
      return ctx.sync().then(function () {
        if (!fresh && !used.isNullObject) return used.rowIndex + used.rowCount;
        var hdr = ws.getRange("A1:I1");
        hdr.values = [["Time", "Action", "Sheet", "Row", "Key (CDF)", "Column", "Old value", "New value", "Note"]];
        hdr.format.font.bold = true; hdr.format.fill.color = "#13294B"; hdr.format.font.color = "#FFFFFF";
        return 1;
      });
    }).then(function (startRow) {
      if (startRow == null) return;
      var rg = ws.getRangeByIndexes(startRow, 0, rows.length, 9);
      rg.numberFormat = rows.map(function (r) { return r.map(function () { return "@"; }); });
      rg.values = rows;
      return ctx.sync();
    });
  }
  function stamp() { var d = new Date(); return K.ddmmyyyy(K.serialOf(d)) + " " + hhmm(d); }

  // ================= SUMMARY: sheet CDF SUMMARY as cards =================
  var SUMMARY_SHEET = "CDF SUMMARY";
  function refreshSummary(manual) {
    if (!state.inExcel) return Promise.resolve();
    $("sumMeta").textContent = "Reading sheet " + SUMMARY_SHEET + "…";
    return Excel.run(function (ctx) {
      return readSheets(ctx, [SUMMARY_SHEET]).then(function (rd) { return rd.sheets[SUMMARY_SHEET]; });
    }).then(function (sh) {
      return (state.data && state.data.cdf ? Promise.resolve(state.data) : readAll()).then(function (d) {
        state.sumAt = Date.now();
        if (!sh) { state.sum = null; $("sumMeta").textContent = ""; $("sumView").innerHTML = '<div class="dempty"><b>This workbook has no sheet ' + esc(SUMMARY_SHEET) + ".</b> The algorithm writes it; once it's there, refresh.</div>"; return; }
        state.sum = V.parse(sh.values);
        renderSummary(d);
        if (manual) $("sumMeta").insertAdjacentHTML("beforeend", " · read " + esc(hhmm(new Date())));
      });
    }).catch(function (e) { $("sumMeta").textContent = ""; $("sumView").innerHTML = '<div class="note bad">' + esc(errText(e)) + "</div>"; });
  }
  function renderSummary(d) {
    var p = state.sum, r = state.rules || K.parseRules(""), out = V.render(p, { cdf: d && d.cdf, exclude: r.exclude });
    state.sumLists = out.lists;
    var n = p ? p.tables.length : 0;
    $("sumMeta").innerHTML = n ? "<b>" + n + "</b> tables in sheet " + esc(SUMMARY_SHEET) + (p.asOf ? " · as of <b>" + esc(p.asOf) + "</b>" : "") : "";
    $("sumView").innerHTML = out.html;
    $("sumView").classList.add("sv-view");
    filterSummary();
  }
  function onSumClick(ev) {
    var t = ev.target.closest ? ev.target.closest("[data-key],[data-list],[data-goto],[data-all]") : null;
    if (!t) return;
    ev.preventDefault();
    if (t.getAttribute("data-key")) { selectCdf(t.getAttribute("data-key")); return; }
    if (t.getAttribute("data-goto")) { var c = $("sv-" + t.getAttribute("data-goto")); if (c) c.scrollIntoView({ behavior: "smooth", block: "start" }); return; }
    if (t.getAttribute("data-all")) { var a = $("sv-" + t.getAttribute("data-all")); if (a) a.classList.add("all"); return; }
    var l = state.sumLists[t.getAttribute("data-list")];
    if (l) openList(l.title, l.keys);
  }
  // the filter: rows that hold the text stay, cards with none go
  function filterSummary() {
    var q = K.trim($("sumFilter").value).toLowerCase(), view = $("sumView");
    view.classList.toggle("filtering", !!q);
    Array.prototype.forEach.call(view.querySelectorAll(".dcard.sv"), function (card) {
      var any = false;
      Array.prototype.forEach.call(card.querySelectorAll("[data-text]"), function (x) {
        var hit = !q || x.getAttribute("data-text").indexOf(q) >= 0;
        x.classList.toggle("hit-no", !hit); if (hit) any = true;
      });
      card.classList.toggle("hit-none", !!q && !any && card.querySelector(".dhead h3").textContent.toLowerCase().indexOf(q) < 0);
      if (q && !any && !card.classList.contains("hit-none")) Array.prototype.forEach.call(card.querySelectorAll(".hit-no"), function (x) { x.classList.remove("hit-no"); });
    });
  }

  // ================= find a CDF =================
  function openLookup() { $("lookup").classList.remove("hidden"); var i = $("lookupInput"); i.value = ""; i.focus(); }
  function closeLookup() { $("lookup").classList.add("hidden"); }
  function doLookup(q) {
    var p = K.parseQuery(q), box = $("lookupResult");
    if (!p) { box.innerHTML = '<p class="hint">Type a CDF number: 856, 740.A or A330-1-856.A</p>'; return; }
    (state.data ? Promise.resolve(state.data) : readAll()).then(function (d) {
      var cdf = d.cdf || [], gi = S.gccIndex(d.gcc), ri = S.revIndex(d.rev), hits = K.buildBlocks(cdf).filter(function (b) { var k = K.keyParts(b.key); return k && k.num === p.num && (!p.rev || k.rev === p.rev); });
      if (!hits.length) { box.innerHTML = '<p class="hint">No CDF ' + esc(q) + " in sheet CDF.</p>"; return; }
      var C = function (L) { return K.col(L); };
      box.innerHTML = hits.map(function (b) {
        var a = cdf[b.anchor] || [], g = gi[K.keyBase(b.key)] || {}, h = cdf[0] || [];
        var f = function (L) { var v = a[C(L)]; return typeof v === "number" && (L === "F" || L === "Q") ? K.ddmmyyyy(v) : K.trim(v); };
        var docs = b.rows.map(function (r) {
          var x = cdf[r] || [], doc = K.trim(x[C("AA")]);
          if (!doc) return "";
          var real = doc !== "-" && !/answer at gcc|no eng/i.test(doc), ds = real && ri && !ri.error ? S.docState(ri, doc, x[C("AB")]) : null;
          var tag = ds ? (ds.state === "released" ? '<span class="lk-rs rel">released</span>' : '<span class="lk-rs work">in work' + (ds.released ? " · rev " + esc(ds.released) + " released" : "") + "</span>") : "";
          return "<li>" + esc(doc + (K.trim(x[C("AB")]) ? " rev " + K.trim(x[C("AB")]) : "")) + " " + tag + ' <span class="dim">' + esc(K.trim(x[C("U")])) + "</span></li>";
        }).join("");
        var notes = L && state.myNotes ? L.notesFor(b.key, state.myNotes.notes) : [];
        return '<div class="lk-card"><div class="lk-top"><b>' + esc(b.key) + '</b><span class="lk-st">' + esc(f("R")) + "</span></div>" +
          '<div class="lk-grid">' + [["Station", f("S")], ["Signature", f("T")], ["Supplier", f("P")], ["Received", f("F")], ["Closed", f("Q")], ["GCC open", g.open != null ? K.ddmmyyyy(g.open) : ""], ["Priority", f("G")], ["DWG", f("J")]]
            .filter(function (x) { return x[1] && x[1] !== "-"; }).map(function (x) { return "<span>" + esc(x[0]) + "</span><b>" + esc(x[1]) + "</b>"; }).join("") + "</div>" +
          (f("H") ? '<p class="lk-desc" dir="auto">' + esc(f("H").slice(0, 400)) + "</p>" : "") +
          (docs ? '<div class="lk-h">Solution documents</div><ul class="lk-docs">' + docs + "</ul>" : "") +
          (K.trim(a[C("E")]) ? '<div class="lk-h">Notes (column E)</div><p class="lk-desc" dir="auto">' + esc(K.trim(a[C("E")])).replace(/\n/g, "<br>") + "</p>" : "") +
          (notes.length ? '<div class="lk-h">In My Notes</div><ul class="lk-docs">' + notes.slice(0, 8).map(function (n) { return '<li dir="auto">' + esc((n.state && n.state !== "Open" ? n.state + ": " : "") + n.text) + "</li>"; }).join("") + "</ul>" : "") +
          '<button class="btn small" type="button" data-key="' + esc(b.key) + '">Go to row ' + (b.anchor + 1) + "</button></div>";
      }).join("");
      Array.prototype.forEach.call(box.querySelectorAll("[data-key]"), function (x) { x.addEventListener("click", function () { selectCdf(x.getAttribute("data-key")); closeLookup(); }); });
    }).catch(function (e) { box.innerHTML = '<div class="note bad">' + esc(errText(e)) + "</div>"; });
  }

  // ================= My Notes (Outlook) =================
  function refreshStatusPack() {
    if (!L || !state.data || !state.data.cdf) return null;
    var pack = N.bridgeStatus(state.data.cdf, new Date(), "CDF Updater " + version(), state.stations), text = L.encode(L.KIND_STATUS, pack);
    state.statusPack = { at: pack.at, count: pack.count, text: text };
    L.writeShared(window.localStorage, L.LS_STATUS, text);
    return state.statusPack;
  }
  function copyStatusForMyNotes(btn) {
    var p = refreshStatusPack(), label = "Copy CDF status for My Notes";
    if (!p) { linkMsg("bad", "Sheet CDF isn't read yet — open STATUS first."); return; }
    copyText(p.text).then(function (ok) {
      if (ok) { btn.textContent = "Copied ✓ — paste it in My Notes"; linkMsg("ok", "Copied the status of " + p.count + " CDFs. In Outlook: My Notes → Reports → CDF status → paste."); }
      else { var box = $("linkCopyBox"); box.value = p.text; box.classList.remove("hidden"); box.select(); linkMsg("", "Press Ctrl+C to copy the selected text, then paste it in My Notes."); }
      setTimeout(function () { btn.textContent = label; }, 4000);
    });
  }
  function linkMsg(cls, text) { var m = $("linkMsg"); m.className = "msg " + (cls || ""); m.textContent = text; }
  // notes pasted from My Notes → saved in the workbook → written into column E at once (the paste is the go-ahead)
  function takeMyNotes(text) {
    var r = L.decode(text, L.KIND_NOTES);
    if (!r.ok) { linkMsg("bad", r.error); return; }
    $("linkPaste").value = "";
    saveMyNotes(r.data);
    writeNotes(true);
  }
  function saveMyNotes(data) { state.myNotes = data; if (storesSupported()) writeStore(NS_NOTES, { data: data, pastedAt: new Date().toISOString() }).catch(function () { /* this session */ }); renderLink(); }
  // My Notes left newer notes in the storage the two add-ins share: say so, one click writes them
  function adoptSharedNotes() {
    if (!L) return;
    var shared = L.readShared(window.localStorage, L.LS_NOTES, L.KIND_NOTES);
    if (!shared || (state.myNotes && L.newer(state.myNotes, shared) !== shared) || (state.myNotes && shared.at === state.myNotes.at)) return;
    saveMyNotes(shared);
    writeNotes(false);
  }
  function writeNotes(go) {
    if (!state.myNotes || !state.inExcel) return;
    (state.data ? Promise.resolve(state.data) : readAll()).then(function (d) {
      var plan = N.notesPlan(d.cdf || [], state.myNotes.notes);
      state.notesPlan = plan;
      if (!plan.length) { linkMsg("ok", "Notes from My Notes (" + (state.myNotes.notes || []).length + ") — column E is already up to date."); $("linkPlan").innerHTML = ""; return; }
      if (!go) {
        $("linkPlan").innerHTML = '<div class="note">' + plan.length + " CDF" + (plan.length === 1 ? "" : "s") + ' have new notes from Outlook. <button id="notesGo" class="btn small primary" type="button">Write them into column E</button></div>';
        $("notesGo").addEventListener("click", function () { writeNotes(true); });
        showBanner("New notes from My Notes for " + plan.length + " CDFs — SETUP › Link with My Notes writes them into column E.", false);
        return;
      }
      setStatus("Writing notes into column E…");
      return Excel.run(function (ctx) {
        var ws = ctx.workbook.worksheets.getItem("CDF");
        plan.forEach(function (p) { var c = ws.getCell(p.row - 1, K.col("E")); c.values = [[p.value]]; c.format.wrapText = true; });
        return ctx.sync().then(function () {
          return appendLog(ctx, plan.map(function (p) { return [stamp(), "NOTES (My Notes)", "CDF", String(p.row), p.key, "E", p.old, p.value, p.n + " note" + (p.n === 1 ? "" : "s")]; })).catch(function () { /* fine */ });
        });
      }).then(function () {
        linkMsg("ok", "Wrote the Outlook notes of " + plan.length + " CDF" + (plan.length === 1 ? "" : "s") + " into column E (lines that start with ✉). Your own text in E stayed.");
        $("linkPlan").innerHTML = ""; showBanner("", false);
        return refreshDash();
      }).then(clearStatus, function (e) { clearStatus(); linkMsg("bad", errText(e)); });
    });
  }
  function renderLink() {
    var tag = $("linkTag");
    if (!tag) return;
    var n = state.myNotes;
    tag.textContent = n ? (n.notes || []).length + " notes from Outlook" : "";
    $("linkOutInfo").textContent = state.statusPack ? "Ready: " + state.statusPack.count + " CDFs, with the Notes of column E." : "The status of every CDF in this workbook, with the Notes of column E.";
  }

  // ================= small things =================
  function showPage(name) {
    state.page = name;
    Array.prototype.forEach.call(document.querySelectorAll(".ecp .key"), function (k) {
      var on = k.getAttribute("data-page") === name;
      k.classList.toggle("on", on); k.setAttribute("aria-selected", on ? "true" : "false"); k.tabIndex = on ? 0 : -1;
    });
    Array.prototype.forEach.call(document.querySelectorAll("main > .page"), function (p) { p.classList.toggle("on", p.id === "page-" + name); });
    if (name === "look") { renderLookPage(); if (state.inExcel && (state.looks || []).length && !state.lookValues) lookValues().then(renderLookPage); }
    if (name === "import") renderSlots();
    if (name === "summary" && state.inExcel && Date.now() - state.sumAt > 60000) refreshSummary();
    if (name === "status" && state.inExcel && Date.now() - state.readAt > 60000) refreshDash();
    try { window.scrollTo(0, 0); } catch (e) { /* ignore */ }
  }
  function refreshButtons() {
    var b = state.busy, x = state.inExcel;
    $("importBtn").disabled = b || !x || !chosen().length;
    $("importBtn").textContent = chosen().length ? "Replace " + chosen().map(function (f) { return f.target; }).join(", ") : "Replace sheets";
    $("rowsWrite").classList.toggle("hidden", !state.rowsItems || !state.rowsItems.length);
    $("rowsWrite").disabled = b || !x || !state.rowsPlan;
    $("rowsUndo").disabled = b || !x || !state.rowsUndo;
    $("repaintBtn").disabled = b || !x;
    ["dashRefresh", "dashPresent", "dashDeck", "dashCopy"].forEach(function (id) { $(id).disabled = b || !x || (id !== "dashRefresh" && !state.model); });
    var la = $("lookAdd"); if (la) la.disabled = b || !x;
  }
  function version() { var v = $("appVersion"); return v ? v.textContent.replace(/^v/, "") : ""; }
  function setStatus(t) { $("status").classList.remove("hidden"); $("statusText").textContent = t; }
  function clearStatus() { $("status").classList.add("hidden"); }
  function setLamp(id, text, cls, title) { var l = $(id); l.textContent = text; l.className = "lamp " + (cls || ""); if (title) l.title = title; }
  function setPill(text, cls) { setLamp("rulesPill", text, cls); }
  function showBanner(t, bad) { var b = $("banner"); b.textContent = t; b.className = "banner" + (bad ? " bad" : "") + (t ? "" : " hidden"); }
  function errText(e) {
    if (!e) return "Something went wrong";
    var msg = e.message || String(e);
    if (e.code === "UnsupportedFeature") msg += " (the file has something Excel can't copy here, e.g. a PivotTable or a chart in Excel on the web)";
    if (e.debugInfo && e.debugInfo.errorLocation) msg += " (" + e.debugInfo.errorLocation + ")";
    return msg;
  }
  function esc(s) { return String(s == null ? "" : s).replace(/[&<>"']/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]; }); }
  function plural(n, w) { return n + " " + w + (n === 1 ? "" : "s"); }
  function capFirst(s) { return s.charAt(0).toUpperCase() + s.slice(1); }
  function hhmm(d) { return ("0" + d.getHours()).slice(-2) + ":" + ("0" + d.getMinutes()).slice(-2); }
  function rnd() { return Math.random().toString(36).slice(2, 8); }
  function isDateFormat(f) { f = String(f || "").replace(/"[^"]*"|\[[^\]]*\]|\\./g, ""); return /[dmy]/i.test(f) && !/^@$|^general$/i.test(f); }
  function decodeText(buf) { try { return new TextDecoder("utf-8", { fatal: true }).decode(buf); } catch (e) { return new TextDecoder("windows-1255").decode(buf); } }
  function toBase64(buf) { var b = new Uint8Array(buf), s = "", CH = 0x8000; for (var i = 0; i < b.length; i += CH) s += String.fromCharCode.apply(null, b.subarray(i, i + CH)); return btoa(s); }
  function openUrl(url) { try { if (window.Office && Office.context && Office.context.ui && Office.context.ui.openBrowserWindow) { Office.context.ui.openBrowserWindow(url); return; } } catch (e) { /* fall back */ } window.open(url, "_blank"); }
  function saveBlob(blob, name) {
    var a = document.createElement("a"), url = URL.createObjectURL(blob);
    a.href = url; a.download = name; document.body.appendChild(a); a.click();
    setTimeout(function () { document.body.removeChild(a); URL.revokeObjectURL(url); }, 1500);
  }
  function copyText(text) {
    function legacy() { var ta = document.createElement("textarea"); ta.value = text; ta.style.cssText = "position:fixed;left:-9999px;top:0;opacity:0"; document.body.appendChild(ta); ta.select(); var ok = false; try { ok = document.execCommand("copy"); } catch (e) { ok = false; } document.body.removeChild(ta); return ok; }
    try { if (navigator.clipboard && navigator.clipboard.writeText) return navigator.clipboard.writeText(text).then(function () { return true; }, function () { return legacy(); }); } catch (e) { /* fall through */ }
    return Promise.resolve(legacy());
  }
  function copyRich(html, text) {
    function legacy() {
      var div = document.createElement("div"); div.contentEditable = "true"; div.style.cssText = "position:fixed;left:-9999px;top:0;opacity:0"; div.innerHTML = html; document.body.appendChild(div);
      var range = document.createRange(); range.selectNodeContents(div); var sel = window.getSelection(); sel.removeAllRanges(); sel.addRange(range);
      var ok = false; try { ok = document.execCommand("copy"); } catch (e) { ok = false; } sel.removeAllRanges(); document.body.removeChild(div); return ok;
    }
    try {
      if (navigator.clipboard && window.ClipboardItem) return navigator.clipboard.write([new ClipboardItem({ "text/html": new Blob([html], { type: "text/html" }), "text/plain": new Blob([text], { type: "text/plain" }) })]).then(function () { return true; }, function () { return legacy(); });
    } catch (e) { /* fall through */ }
    return Promise.resolve(legacy());
  }
  function themeChoice() { try { return localStorage.getItem(THEME_KEY) || "auto"; } catch (e) { return "auto"; } }
  function officeIsDark() {
    try { var t = window.Office && Office.context && Office.context.officeTheme; if (t && t.bodyBackgroundColor) { var h = t.bodyBackgroundColor.replace("#", ""), rgb = [0, 2, 4].map(function (i) { return parseInt(h.substr(i, 2), 16) / 255; }); return 0.2126 * rgb[0] + 0.7152 * rgb[1] + 0.0722 * rgb[2] < 0.4; } } catch (e) { /* ignore */ }
    try { return !!(window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches); } catch (e) { return false; }
  }
  function applyTheme() {
    var c = themeChoice(), dark = c === "dark" || (c === "auto" && officeIsDark());
    document.documentElement.setAttribute("data-theme", dark ? "dark" : "light");
    Array.prototype.forEach.call(document.querySelectorAll("[data-theme-choice]"), function (b) { b.classList.toggle("on", b.getAttribute("data-theme-choice") === c); });
  }
})();
