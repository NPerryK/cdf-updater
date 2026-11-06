/* CDF Updater - task pane (Office.js glue + UI) */
(function () {
  "use strict";

  var E = window.CdfEngine;
  var LS_KEY = "cdfUpdater.localRules.v1";
  var MAX_ROWS_SHOWN = 300;

  var state = {
    inExcel: false,
    canImport: false,
    rules: null,
    rulesText: "",
    source: "github",       // "github" | "local"
    localName: null,
    loadedAt: null,
    files: {},              // slot -> File
    slotStatus: {},         // slot -> {kind:"done"|"failed", msg}
    busy: false,
    lastPreview: null,
    draft: null,            // new-row draft waiting for approval
    queue: [],              // CDF rows waiting for review (paste can hold several)
    queueTotal: 0,
    written: [],
    batch: null,            // several CDFs from one reply: {items, plan, on[]}
    prompt: null,           // {text, version} built from ai-prompt.txt
    promptWarn: null,       // the reply was made with other instructions
    mail: null              // last built email {subject, text, html}
  };
  var GROUPS_KEY = "cdfUpdater.mailGroups.v1";

  var $ = function (id) { return document.getElementById(id); };

  // ---------------- startup ----------------
  function start(info) {
    state.inExcel = !!(info && window.Office && info.host === Office.HostType.Excel);
    setLamp("excelLamp", state.inExcel ? "EXCEL" : "NO EXCEL", state.inExcel ? "ok" : "warn",
      state.inExcel ? "Connected to this workbook" : "Opened outside Excel - workbook actions are off");
    if (!state.inExcel) {
      showBanner("Opened outside Excel: you can check the rules file, but workbook actions only work inside Excel.", false);
    } else {
      state.canImport = Office.context.requirements.isSetSupported("ExcelApi", "1.13");
      if (!state.canImport) {
        showBanner("This Excel version can't replace a sheet from a file (Microsoft 365 required). Running the algorithm still works.", true);
      }
    }
    try { localStorage.removeItem("cdfUpdater.apiKey.v1"); localStorage.removeItem("cdfUpdater.aiMode.v1"); } catch (e) { /* ignore */ }
    wire();
    restoreLocalRules();
    loadRules().then(function () { return fetchPrompt(); });
  }

  if (window.Office && Office.onReady) Office.onReady(start);
  else document.addEventListener("DOMContentLoaded", function () { start(null); });

  function wire() {
    $("importBtn").addEventListener("click", askImport);
    $("importYes").addEventListener("click", doImport);
    $("importNo").addEventListener("click", function () { $("importConfirm").classList.add("hidden"); refreshButtons(); });
    $("previewBtn").addEventListener("click", preview);
    $("reloadRules").addEventListener("click", function () { loadRules(); });
    $("localRules").addEventListener("change", onLocalRules);
    $("useGithub").addEventListener("click", function () {
      state.source = "github"; state.localName = null;
      try { localStorage.removeItem(LS_KEY); } catch (e) { /* ignore */ }
      loadRules();
    });
    // new row
    $("copyPrompt").addEventListener("click", copyPrompt);
    $("openClaude").addEventListener("click", openClaude);
    $("openProject").addEventListener("click", openProject);
    $("copyInstructions").addEventListener("click", function () { copyInstructions($("copyInstructions")); });
    $("saveProject").addEventListener("click", saveProject);
    $("projectUrl").addEventListener("keydown", function (ev) { if (ev.key === "Enter") saveProject(); });
    $("pasteBox").addEventListener("paste", onPaste);
    $("pasteBox").addEventListener("input", function () { $("pasteBox").classList.remove("read"); $("pasteMsg").textContent = ""; refreshButtons(); });
    $("parsePaste").addEventListener("click", parsePasted);
    // email
    $("mailKind").addEventListener("change", function () { $("mailSubject").value = defaultSubject(); $("mailPreview").innerHTML = ""; });
    $("mailBuild").addEventListener("click", buildMail);
    $("saveGroup").addEventListener("click", function () { $("saveGroupBox").classList.toggle("hidden"); $("groupName").focus(); });
    $("saveGroupYes").addEventListener("click", saveLocalGroup);
    ["mailTo", "mailCc"].forEach(function (id) {
      $(id).addEventListener("input", function () { if (state.mail) renderMail(); });
    });
    $("mailSubject").addEventListener("input", function () {
      if (state.mail) { state.mail.subject = $("mailSubject").value; renderMail(); }
    });
  }

  function renderExtras() {
    renderMailKinds(); renderGroups(); renderProject();
  }

  // checklist state of a step: done + a short note next to its title
  function markStep(n, text) {
    var el = $("step" + n) || (n === 3 ? $("newRowCard") : n === 4 ? $("mailCard") : null);
    if (!el) return;
    el.classList.add("done");
    $("state" + n).textContent = text + " " + hhmm(new Date());
  }

  function renderLamps() {
    var r = state.rules;
    if (!r) { setLamp("algoLamp", "ALGO", "", "No rules loaded"); return; }
    if (r.cdfAlgorithm) {
      var want = r.cdfAlgorithm.version, have = window.CdfAlgo.ALGORITHM_VERSION;
      if (want && want !== have) setLamp("algoLamp", "ALGO CHECK", "warn", "rules.txt expects algorithm " + want + ", the code implements " + have);
      else setLamp("algoLamp", "ALGO " + have, "ok", "CDF algorithm " + have + ", stage " + window.CdfAlgo.STAGE);
    } else if (r.algorithms.length) setLamp("algoLamp", "ALGO RULES", "ok", "Algorithm from the rules file");
    else setLamp("algoLamp", "ALGO", "", "No algorithm in the rules file");
  }

  // ---------------- rules ----------------
  function restoreLocalRules() {
    try {
      var saved = JSON.parse(localStorage.getItem(LS_KEY) || "null");
      if (saved && saved.text) { state.source = "local"; state.localName = saved.name; state.rulesText = saved.text; }
    } catch (e) { /* storage unavailable */ }
  }

  function loadRules() {
    setPill("RULES", "");
    var p = state.source === "local"
      ? Promise.resolve(state.rulesText)
      : fetch("rules.txt?v=" + Date.now(), { cache: "no-store" }).then(function (r) {
          if (!r.ok) throw new Error("Couldn't read rules.txt (HTTP " + r.status + ")");
          return r.text();
        });
    return p.then(applyRulesText).catch(function (e) {
      state.rules = null;
      setPill("NO RULES", "bad"); renderLamps();
      $("rulesErrors").innerHTML = '<div class="err">' + esc(e.message) + "</div>";
      renderSlots(); renderAlgos(); renderExtras(); refreshButtons();
    });
  }

  function applyRulesText(text) {
    state.rulesText = text;
    state.rules = E.parseRules(text);
    state.loadedAt = new Date();
    state.lastPreview = null;
    $("results").innerHTML = "";
    var errs = state.rules.errors.length;
    setPill(errs ? "RULES " + errs + " ERR" : "RULES OK", errs ? "bad" : "ok");
    renderLamps();
    renderRules(); renderSlots(); renderAlgos(); renderExtras(); refreshButtons();
  }

  function onLocalRules(ev) {
    var f = ev.target.files[0];
    ev.target.value = "";
    if (!f) return;
    f.arrayBuffer().then(function (buf) {
      var text = decodeText(buf);
      state.source = "local"; state.localName = f.name;
      try { localStorage.setItem(LS_KEY, JSON.stringify({ name: f.name, text: text })); } catch (e) { /* ignore */ }
      applyRulesText(text);
    });
  }

  function renderRules() {
    var src = state.source === "local"
      ? "Local file: <b>" + esc(state.localName || "") + "</b> (kept on this computer until you switch back to GitHub)."
      : "<b>rules.txt</b> from GitHub — reloaded every time you click \"Check changes\".";
    if (state.loadedAt) src += " Loaded at " + hhmm(state.loadedAt) + ".";
    $("rulesSource").innerHTML = src;
    $("useGithub").classList.toggle("hidden", state.source !== "local");
    $("rulesText").textContent = state.rulesText;
    var r = state.rules;
    $("rulesErrors").innerHTML = r && r.errors.length
      ? r.errors.map(function (e) {
          return '<div class="err">Line ' + e.line + ": " + esc(e.message) + (e.text ? '<br><code>' + esc(e.text) + "</code>" : "") + "</div>";
        }).join("")
      : "";
  }

  function renderAlgos() {
    var r = state.rules;
    if (r && r.cdfAlgorithm) { $("algos").innerHTML = cdfAlgoCard(r); return; }
    if (!r || !r.algorithms.length) { $("algos").innerHTML = '<p class="hint">No algorithm is defined in the rules file.</p>'; return; }
    $("algos").innerHTML = r.algorithms.map(function (a) {
      var steps = a.actions.map(function (x) { return "<li>" + esc(E.describeAction(x)) + "</li>"; }).join("");
      var oc = a.onChange.map(function (x) { return '<li class="onchange">If the row changed — ' + esc(E.describeAction(x)) + "</li>"; }).join("");
      return '<div class="algo">' +
        '<div class="algo-name">' + esc(a.name) + "</div>" +
        '<div class="algo-meta">Base: <b>' + (a.baseFirst ? "first sheet" : esc(a.base)) + "</b> · Update: <b>" + esc(a.update || "?") + "</b><br>" +
        (a.match ? "Match: <b>UPDATE." + a.match.update.col + " = BASE." + a.match.base.col + "</b> · " : "") +
        "Data from row " + a.baseStartRow + " / " + a.updateStartRow + "</div>" +
        "<ol>" + steps + oc + "</ol></div>";
    }).join("");
  }

  // ---------------- import ----------------
  function renderSlots() {
    var imports = state.rules ? state.rules.imports : [];
    var valid = {};
    imports.forEach(function (s) { valid[s.slot] = true; });
    Object.keys(state.files).forEach(function (k) { if (!valid[k]) delete state.files[k]; });

    if (!imports.length) { $("slots").innerHTML = '<p class="hint">No IMPORT lines in the rules file.</p>'; return; }

    $("slots").innerHTML = "";
    imports.forEach(function (s) {
      var f = state.files[s.slot], st = state.slotStatus[s.slot];
      var el = document.createElement("label");
      el.className = "slot" + (f ? " has-file" : "") + (st && !f ? " " + st.kind : "");
      el.innerHTML =
        '<div class="slot-title"><span>' + (s.file ? '<span class="slot-expect">' + esc(s.file) + "</span>" : "File " + s.slot) + "</span>" +
        (f ? '<button type="button" class="slot-clear" title="Remove">✕</button>' : "") + "</div>" +
        '<div class="slot-target">Replaces sheet <b>' + esc(s.target) + "</b>" + (s.from ? " (from \"" + esc(s.from) + "\")" : "") + "</div>" +
        '<div class="slot-file">' + (f ? esc(f.name) + " (" + kb(f.size) + ")" : "Drop the file here or click to choose") + "</div>" +
        (f && s.file && !E.importFileMatches(s.file, f.name) ? '<div class="slot-warn">Expected a file named ' + esc(s.file) + "</div>" : "") +
        (st && !f ? '<div class="slot-msg">' + esc(st.msg) + "</div>" : "") +
        '<input type="file" accept=".xlsx,.xlsm,.csv" multiple hidden>';
      var input = el.querySelector("input");
      input.addEventListener("change", function () { takeImportFiles(input.files, s.slot); input.value = ""; });
      var clr = el.querySelector(".slot-clear");
      if (clr) clr.addEventListener("click", function (ev) { ev.preventDefault(); ev.stopPropagation(); delete state.files[s.slot]; renderSlots(); refreshButtons(); });
      el.addEventListener("dragover", function (ev) { ev.preventDefault(); el.classList.add("drag"); });
      el.addEventListener("dragleave", function () { el.classList.remove("drag"); });
      el.addEventListener("drop", function (ev) {
        ev.preventDefault(); el.classList.remove("drag");
        takeImportFiles(ev.dataTransfer.files, s.slot);
      });
      $("slots").appendChild(el);
    });
  }

  // Files go to the slot whose FILE name they match, wherever they are dropped.
  // A file that matches no slot goes to the slot it was dropped on (or the next empty one).
  function takeImportFiles(list, droppedOn) {
    var files = Array.prototype.slice.call(list || []);
    if (!files.length) return;
    var imports = state.rules ? state.rules.imports : [];
    var unmatched = [];
    files.forEach(function (f) {
      var slot = E.matchImportSlot(imports, f.name);
      if (slot != null) setSlotFile(slot, f); else unmatched.push(f);
    });
    unmatched.forEach(function (f, i) {
      var target = droppedOn;
      if (i > 0 || (files.length > 1 && state.files[droppedOn] && files.indexOf(state.files[droppedOn]) >= 0)) {
        var free = imports.filter(function (s) { return !state.files[s.slot]; })[0];
        if (free) target = free.slot;
      }
      setSlotFile(target, f);
    });
    $("importConfirm").classList.add("hidden");
    renderSlots(); refreshButtons();
  }

  function setSlotFile(slot, file) {
    state.files[slot] = file;
    delete state.slotStatus[slot];
  }

  function mismatchWarnings(slots) {
    var w = slots.filter(function (s) { return s.file && !E.importFileMatches(s.file, state.files[s.slot].name); });
    if (!w.length) return "";
    return '<br><span class="warn-text">⚠ ' + w.map(function (s) {
      return "File " + s.slot + " is “" + esc(state.files[s.slot].name) + "” — expected a file named " + esc(s.file) + ".";
    }).join("<br>⚠ ") + "</span>";
  }

  function chosenSlots() {
    return (state.rules ? state.rules.imports : []).filter(function (s) { return state.files[s.slot]; });
  }

  function askImport() {
    var slots = chosenSlots();
    if (!slots.length) return;
    sheetNames().then(function (existing) {
      var replace = slots.filter(function (s) { return existing.indexOf(s.target) >= 0; }).map(function (s) { return s.target; });
      var create = slots.filter(function (s) { return existing.indexOf(s.target) < 0; }).map(function (s) { return s.target; });
      var html = "";
      if (replace.length) html += "Sheet" + (replace.length > 1 ? "s" : "") + " <b>" + replace.map(esc).join(", ") + "</b> will be deleted and replaced with the file contents. Formulas on other sheets that point to " + (replace.length > 1 ? "them" : "it") + " will turn into #REF!.";
      if (create.length) html += (html ? "<br>" : "") + "New sheet" + (create.length > 1 ? "s" : "") + " will be created: <b>" + create.map(esc).join(", ") + "</b>.";
      $("importConfirmText").innerHTML = html + mismatchWarnings(slots) + "<br>Continue?";
      $("importConfirm").classList.remove("hidden");
      $("importBtn").disabled = true;
    }).catch(showError);
  }

  function doImport() {
    $("importConfirm").classList.add("hidden");
    var slots = chosenSlots();
    state.busy = true; refreshButtons();
    var chain = Promise.resolve();
    slots.forEach(function (s) {
      chain = chain.then(function () {
        var f = state.files[s.slot];
        setStatus("Replacing sheet " + s.target + "…");
        return importOne(s, f).then(function (info) {
          delete state.files[s.slot];
          state.slotStatus[s.slot] = { kind: "done", msg: "Replaced ✓ " + hhmm(new Date()) + ", " + f.name + (info ? ", " + info : "") };
        }).catch(function (e) {
          state.slotStatus[s.slot] = { kind: "failed", msg: errText(e) };
        }).then(renderSlots);
      });
    });
    chain.then(function () {
      state.busy = false; clearStatus(); refreshButtons();
      var ok = slots.filter(function (s) { return state.slotStatus[s.slot] && state.slotStatus[s.slot].kind === "done"; });
      if (ok.length) markStep(1, ok.length + (ok.length > 1 ? " sheets" : " sheet") + " replaced");
    });
  }

  function importOne(slot, file) {
    var ext = (file.name.split(".").pop() || "").toLowerCase();
    return file.arrayBuffer().then(function (buf) {
      if (ext === "csv") return importCsv(slot, buf);
      if (ext !== "xlsx" && ext !== "xlsm") throw new Error("Unsupported file type (." + ext + "). Save it as .xlsx and try again.");
      return workbookSheetNames(buf).then(function (names) {
        if (!names.length) throw new Error("No sheets found in the file");
        var src = slot.from || names[0];
        if (names.indexOf(src) < 0) throw new Error("The file has no sheet named \"" + src + "\". Sheets: " + names.join(", "));
        var b64 = toBase64(buf);
        return replaceSheet(slot.target, function (ctx, rel) {
          var opts = { sheetNamesToInsert: [src] };
          if (rel) { opts.positionType = Excel.WorksheetPositionType.after; opts.relativeTo = rel.tmpName; }
          else opts.positionType = Excel.WorksheetPositionType.end;
          var res = ctx.workbook.insertWorksheetsFromBase64(b64, opts);
          return ctx.sync().then(function () {
            if (!res.value || !res.value.length) throw new Error("Excel didn't add the sheet");
            return res.value[0];
          });
        }).then(function () { return names.length > 1 ? "sheet \"" + src + "\"" : ""; });
      });
    });
  }

  function importCsv(slot, buf) {
    var rows = parseCsv(decodeText(buf));
    if (!rows.length) throw new Error("The CSV file is empty");
    var width = rows.reduce(function (m, r) { return Math.max(m, r.length); }, 0);
    rows = rows.map(function (r) { while (r.length < width) r.push(""); return r; });
    return replaceSheet(slot.target, function (ctx, rel) {
      var ws = ctx.workbook.worksheets.add("~new_" + rnd());
      if (rel) ws.position = rel.position + 1;
      for (var i = 0; i < rows.length; i += 2000) {
        var chunk = rows.slice(i, i + 2000);
        ws.getRangeByIndexes(i, 0, chunk.length, width).values = chunk;
      }
      ws.load("id");
      return ctx.sync().then(function () { return ws.id; });
    }).then(function () { return rows.length + " rows"; });
  }

  // Rename old -> insert new next to it -> delete old -> rename new. Rolls back on failure.
  function replaceSheet(target, insertFn) {
    return Excel.run(function (ctx) {
      var sheets = ctx.workbook.worksheets;
      var old = sheets.getItemOrNullObject(target);
      old.load("name,position");
      var tmpName = null, newId = null;
      return ctx.sync().then(function () {
        if (old.isNullObject) return null;
        tmpName = "~old_" + rnd();
        old.name = tmpName;
        return ctx.sync().then(function () { return { tmpName: tmpName, position: old.position }; });
      }).then(function (rel) {
        return insertFn(ctx, rel);
      }).then(function (id) {
        newId = id;
        var ns = sheets.getItem(id);
        if (tmpName) sheets.getItem(tmpName).delete();
        ns.name = target;
        return ctx.sync();
      }).catch(function (err) {
        // roll back: keep the original sheet if it is still there
        if (!tmpName) throw err;
        var o = sheets.getItemOrNullObject(tmpName);
        var n = newId ? sheets.getItemOrNullObject(newId) : null;
        o.load("name"); if (n) n.load("name");
        return ctx.sync().then(function () {
          if (!o.isNullObject) {
            if (n && !n.isNullObject) n.delete();
            o.name = target;
          }
          return ctx.sync();
        }).catch(function () { /* best effort */ }).then(function () { throw err; });
      });
    });
  }

  function workbookSheetNames(buf) {
    return JSZip.loadAsync(buf).then(function (zip) {
      var wb = zip.file("xl/workbook.xml");
      if (!wb) throw new Error("This doesn't look like a valid Excel workbook");
      return wb.async("string");
    }).then(function (xml) {
      var doc = new DOMParser().parseFromString(xml, "application/xml");
      var els = doc.getElementsByTagNameNS("*", "sheet");
      var names = [];
      for (var i = 0; i < els.length; i++) names.push(els[i].getAttribute("name"));
      return names;
    });
  }

  function sheetNames() {
    if (!state.inExcel) return Promise.reject(new Error("This only works inside Excel"));
    return Excel.run(function (ctx) {
      var ws = ctx.workbook.worksheets; ws.load("items/name");
      return ctx.sync().then(function () { return ws.items.map(function (w) { return w.name; }); });
    });
  }

  // ---------------- algorithm ----------------
  function preview() {
    if (state.busy) return;
    state.busy = true; refreshButtons();
    $("results").innerHTML = "";
    setStatus("Reading rules…");
    var reload = state.source === "github" ? loadRules() : Promise.resolve();
    reload.then(function () {
      if (!state.rules || state.rules.errors.length) throw new Error("The rules file has errors — fix them and try again.");
      if (!state.inExcel) throw new Error("This only works inside Excel");
      setStatus("Comparing data…");
      if (state.rules.cdfAlgorithm) return runCdf(false).then(function (res) { renderCdfResults(res, false); });
      return runAlgorithms(false).then(function (results) {
        state.lastPreview = results;
        renderResults(results, false);
      });
    }).catch(showError).then(function () { state.busy = false; clearStatus(); refreshButtons(); });
  }

  function applyChanges() {
    if (state.busy) return;
    state.busy = true; refreshButtons();
    setStatus("Writing changes…");
    var job = state.rules.cdfAlgorithm
      ? runCdf(true).then(function (res) { renderCdfResults(res, true); })
      : runAlgorithms(true).then(function (results) { state.lastPreview = null; renderResults(results, true); });
    job.catch(showError).then(function () { state.busy = false; clearStatus(); refreshButtons(); });
  }

  // ================= CDF algorithm (code) =================
  function cdfAlgoCard(r) {
    var A = window.CdfAlgo, want = r.cdfAlgorithm.version, have = A.ALGORITHM_VERSION;
    var html = '<div class="algo"><div class="algo-name">CDF algorithm</div>' +
      '<div class="algo-meta">Algorithm version in the code: <b>' + esc(have) + "</b>" +
      (want ? "<br>The rules file expects: <b>" + esc(want) + "</b>" : "") + "</div>";
    if (want && want !== have) html += '<div class="note warn" style="margin-top:6px">The rules file expects algorithm version ' + esc(want) +
      " but the code implements " + esc(have) + ". The code needs an update before the results can be trusted.</div>";
    html += "<ol>" + A.STAGES.map(function (s) {
      return s.done ? '<li class="done"><span class="tickmark">✓</span> ' + esc(s.name) + "</li>"
        : '<li class="onchange"><span class="later">Later:</span> ' + esc(s.name) + "</li>";
    }).join("") + "</ol></div>";
    return html;
  }

  function readOptionalSheet(ctx, name) {
    var ws = ctx.workbook.worksheets.getItemOrNullObject(name);
    ws.load("name");
    return ctx.sync().then(function () { return ws.isNullObject ? null : readSheet(ctx, ws); });
  }

  function runCdf(apply) {
    var rules = state.rules, now = new Date();
    return Excel.run(function (ctx) {
      var data = {};
      return readOptionalSheet(ctx, "CDF").then(function (d) {
        if (!d) throw new Error("Sheet \"CDF\" was not found in this workbook");
        data.cdf = d;
        return readOptionalSheet(ctx, "GCC");
      }).then(function (d) {
        if (!d) throw new Error("Sheet \"GCC\" was not found — import CDF_List_Export first");
        data.gcc = d;
        return readOptionalSheet(ctx, "CDF SUMMARY");
      }).then(function (d) {
        data.summary = d;
        var res = window.CdfAlgo.run({ cdf: data.cdf.values, gcc: data.gcc.values, now: now }, rules.cdfConfig);
        res.compare = data.summary ? window.CdfAlgo.compareWithSheet(res.summary, data.summary.values) : null;
        if (!apply) return res;
        if (res.errors.some(function (e) { return e.stage === "00.1"; })) throw new Error("Headers don't match the algorithm — nothing was written.");
        var cdfWs = ctx.workbook.worksheets.getItem("CDF"), gccWs = ctx.workbook.worksheets.getItem("GCC");
        setStatus("Writing " + (res.cdfChanges.length + res.gccChanges.length) + " cells…");
        return writeChanges(ctx, gccWs, res.gccChanges).then(function () {
          return writeChanges(ctx, cdfWs, res.cdfChanges);
        }).then(function () {
          if (!rules.log) return;
          var stamp = stampOf(now), name = "CDF algorithm " + window.CdfAlgo.ALGORITHM_VERSION + " · stage " + window.CdfAlgo.STAGE;
          var rows = res.gccChanges.concat(res.cdfChanges).map(function (c) {
            return [stamp, name, c.sheet, String(c.row), c.key, c.colLetter, show(c.oldValue, null), show(c.newValue, null), c.rule === "20.2" ? "" : "Rule " + c.rule + " — " + c.why];
          });
          res.info.onlyGcc.forEach(function (k) { rows.push([stamp, name, "GCC", "", k, "", "", "", E.LOG_NOTES.notFound]); });
          if (!rows.length) rows.push([stamp, name, "", "", "", "", "", "", E.LOG_NOTES.none]);
          setStatus("Writing log…");
          return appendLog(ctx, rules.log, rows);
        }).then(function () { return res; });
      });
    });
  }

  function renderCdfResults(res, applied) {
    var nChanges = res.cdfChanges.length + res.gccChanges.length;
    var html = '<div class="res-block"><div class="res-title">CDF algorithm · stage ' + res.stage + "</div>";
    if (res.errors.length) html += '<div class="note bad"><b>' + res.errors.length + " error" + (res.errors.length > 1 ? "s" : "") + "</b> — these cells are not changed:<ul>" +
      res.errors.map(function (e) { return "<li>" + esc((e.key ? e.key + ": " : "") + e.message) + " <span class=\"rule\">(" + esc(e.stage) + ")</span></li>"; }).join("") + "</ul></div>";
    if (!res.summary) { $("results").innerHTML = html + "</div>"; return; }

    html += '<div class="tiles">' +
      tile(res.info.gccRows, "CDFs in GCC") +
      tile(res.info.matched, "Matched in CDF") +
      tile(res.info.onlyGcc.length, "Only in GCC", res.info.onlyGcc.length ? "warnt" : "") +
      tile(countKeys(res.cdfChanges), applied ? "CDFs updated" : "CDFs to change", "hl") +
      tile(nChanges, applied ? "Cells updated" : "Cells to change", "hl") +
      tile(res.warnings.length, "Warnings", res.warnings.length ? "warnt" : "") + "</div>";

    if (applied && nChanges) html += '<div class="note ok">Update written ✓' + (state.rules.log ? " — every change and its previous value is in sheet <b>" + esc(state.rules.log) + "</b>." : "") + "</div>";
    if (!nChanges) html += '<div class="note ok">Nothing to update — the CDF sheet already matches GCC.</div>';
    if (applied && nChanges) markStep(2, nChanges + " cells written");
    else if (!nChanges && !res.errors.length) markStep(2, "Up to date");
    if (res.info.onlyGcc.length) html += '<div class="note warn">In GCC but not in the CDF sheet (add them as new rows):' + chips(res.info.onlyGcc) + "</div>";
    if (res.info.onlyCdf.length) html += '<div class="note warn">In the CDF sheet but not in GCC (left as they are):' + chips(res.info.onlyCdf) + "</div>";
    if (res.warnings.length) html += '<div class="note warn"><ul>' + res.warnings.map(function (w) { return "<li>" + esc(w.key + ": " + w.message) + "</li>"; }).join("") + "</ul></div>";

    html += crossCheckHtml(res, applied);
    if (nChanges) html += cdfChangesTable(res.gccChanges.concat(res.cdfChanges));
    html += "</div>";

    if (!applied && nChanges) {
      html += '<div class="confirm go"><p><b>' + nChanges + "</b> cell" + (nChanges > 1 ? "s" : "") + " will be written to sheets CDF" +
        (res.gccChanges.length ? " and GCC" : "") + ". Save the workbook first — Ctrl+Z may not undo this" + (state.rules.log ? ", but the log keeps the previous values" : "") + ".</p>" +
        '<div class="row"><button id="applyBtn" class="btn ok">Apply ' + nChanges + " change" + (nChanges > 1 ? "s" : "") + "</button>" +
        '<button id="cancelApply" class="btn ghost">Cancel</button></div></div>';
    }
    $("results").innerHTML = html;
    var ab = $("applyBtn");
    if (ab) ab.addEventListener("click", applyChanges);
    var cb = $("cancelApply");
    if (cb) cb.addEventListener("click", function () { $("results").innerHTML = ""; });
    $("results").scrollIntoView({ behavior: "smooth", block: "start" });
  }

  function countKeys(changes) {
    var k = {}; changes.forEach(function (c) { k[c.key] = true; }); return Object.keys(k).length;
  }

  // the cross-check: control totals (130.6) + every computed table against CDF SUMMARY in this workbook
  function crossCheckHtml(res, applied) {
    var ok = res.checks.filter(function (c) { return c.ok; }).length;
    var html = '<div class="xcheck"><div class="xcheck-title">Cross-check</div>' +
      '<div class="xsub">Control totals (' + ok + "/" + res.checks.length + " OK)</div><ul class=\"xlist\">" +
      res.checks.map(function (c) {
        return '<li class="' + (c.ok ? "ok" : "bad") + '"><span class="mark">' + (c.ok ? "✓" : "✗") + "</span>" + esc(c.name) +
          '<div class="xdetail">' + esc(c.detail) + "</div></li>";
      }).join("") + "</ul>";
    if (res.compare) {
      var same = res.compare.filter(function (t) { return t.found && !t.diffs.length; }).length;
      html += '<div class="xsub">CDF SUMMARY tables: result vs. what the workbook shows now (' + same + "/" + res.compare.length + " identical)</div>" +
        '<p class="hint" style="margin:2px 0 6px">' + (applied ? "Writing these tables to CDF SUMMARY comes with stage 3."
          : "A difference here is expected when GCC changed since the last run — it shows what the table will be after the update.") + "</p><ul class=\"xlist\">" +
        res.compare.map(function (t) {
          if (!t.found) return '<li class="warn"><span class="mark">?</span>' + esc(t.table) + '<div class="xdetail">Not found in CDF SUMMARY</div></li>';
          if (!t.diffs.length) return '<li class="ok"><span class="mark">=</span>' + esc(t.table) + ' <span class="xdetail-inline">(' + esc(t.at) + ")</span></li>";
          return '<li class="diff"><span class="mark">≠</span>' + esc(t.table) + ' <span class="xdetail-inline">(' + esc(t.at) + ")</span>" +
            '<table class="xdiff"><tr><th></th><th>Now</th><th>Result</th></tr>' +
            t.diffs.map(function (d) { return "<tr><td>" + esc(d.label) + "</td><td>" + esc(d.current) + "</td><td>" + esc(d.computed) + "</td></tr>"; }).join("") +
            "</table></li>";
        }).join("") + "</ul>";
    }
    html += '<details class="xtables"><summary>Show the computed tables</summary>' +
      res.summary.tables.map(function (t) {
        return '<div class="xsub">' + esc(t.name) + '</div><table class="xdiff"><tr>' + t.headers.map(function (h) { return "<th>" + esc(h) + "</th>"; }).join("") + "</tr>" +
          t.rows.map(function (r) { return "<tr>" + r.map(function (v) { return "<td>" + esc(v).replace(/\n/g, "<br>") + "</td>"; }).join("") + "</tr>"; }).join("") + "</table>";
      }).join("") + "</details></div>";
    return html;
  }

  function cdfChangesTable(changes) {
    var rows = changes.slice(0, MAX_ROWS_SHOWN).map(function (c) {
      return '<tr title="' + esc("Rule " + c.rule + (c.why ? " — " + c.why : "") + (c.source ? " · " + c.source : "")) + '"><td>' + esc(c.sheet === "GCC" ? "GCC " + c.row : c.row) +
        '</td><td class="v k">' + esc(c.key) + "</td><td>" + c.colLetter +
        '</td><td class="v old">' + esc(show(c.oldValue, null) || "(empty)") + '</td><td class="v new">' + esc(show(c.newValue, null) || "(empty)") +
        '</td><td class="rule">' + esc(c.rule) + "</td></tr>";
    }).join("");
    var more = changes.length > MAX_ROWS_SHOWN ? '<div class="more">and ' + (changes.length - MAX_ROWS_SHOWN) + " more changes…</div>" : "";
    return '<div class="tbl-wrap"><table><thead><tr><th>Row</th><th>Key</th><th>Col</th><th>Before</th><th>After</th><th>Rule</th></tr></thead><tbody>' +
      rows + "</tbody></table>" + more + "</div>";
  }

  function runAlgorithms(apply) {
    var rules = state.rules, now = new Date(), out = [];
    return Excel.run(function (ctx) {
      var chain = Promise.resolve();
      rules.algorithms.forEach(function (algo) {
        chain = chain.then(function () {
          var baseWs = algo.baseFirst ? ctx.workbook.worksheets.getFirst(true) : ctx.workbook.worksheets.getItemOrNullObject(algo.base);
          var updWs = ctx.workbook.worksheets.getItemOrNullObject(algo.update);
          baseWs.load("name"); updWs.load("name");
          return ctx.sync().then(function () {
            if (!algo.baseFirst && baseWs.isNullObject) throw new Error("Base sheet \"" + algo.base + "\" was not found in the workbook");
            if (updWs.isNullObject) throw new Error("Update sheet \"" + algo.update + "\" was not found — did you import the file?");
            if (baseWs.name === updWs.name) throw new Error("The base sheet and the update sheet are the same sheet (" + baseWs.name + ")");
            var data = [];
            return readSheet(ctx, baseWs).then(function (b) { data.push(b); return readSheet(ctx, updWs); })
              .then(function (u) { data.push(u); return data; });
          }).then(function (data) {
            var res = E.computeChanges(algo, data[0], data[1], { now: now });
            res.algo = algo; res.baseName = baseWs.name; res.updateName = updWs.name;
            out.push(res);
            if (apply && res.changes.length) {
              setStatus("Writing " + res.changes.length + " cells to sheet " + baseWs.name + "…");
              return writeChanges(ctx, baseWs, res.changes);
            }
          });
        });
      });
      return chain.then(function () {
        if (apply && rules.log) { setStatus("Writing log…"); return writeLog(ctx, rules.log, out, now); }
      }).then(function () { return out; });
    });
  }

  function readSheet(ctx, ws) {
    var used = ws.getUsedRangeOrNullObject(true);
    used.load("rowIndex,rowCount,columnIndex,columnCount");
    return ctx.sync().then(function () {
      if (used.isNullObject) return { values: [], formats: [] };
      var rows = used.rowIndex + used.rowCount, cols = used.columnIndex + used.columnCount;
      var values = [], formats = [], CH = 2000, chain = Promise.resolve();
      for (var r = 0; r < rows; r += CH) {
        (function (start) {
          chain = chain.then(function () {
            var rg = ws.getRangeByIndexes(start, 0, Math.min(CH, rows - start), cols);
            rg.load("values,numberFormat");
            return ctx.sync().then(function () {
              Array.prototype.push.apply(values, rg.values);
              Array.prototype.push.apply(formats, rg.numberFormat);
            });
          });
        })(r);
      }
      return chain.then(function () { return { values: values, formats: formats }; });
    });
  }

  function writeChanges(ctx, ws, changes) {
    var canSuspend = Office.context.requirements.isSetSupported("ExcelApi", "1.6");
    if (canSuspend) ctx.application.suspendApiCalculationUntilNextSync();
    var chain = Promise.resolve();
    for (var i = 0; i < changes.length; i += 400) {
      (function (batch) {
        chain = chain.then(function () {
          if (canSuspend) ctx.application.suspendApiCalculationUntilNextSync();
          batch.forEach(function (c) {
            var rg = ws.getRangeByIndexes(c.row - 1, c.col, 1, 1);
            if (c.numberFormat) rg.numberFormat = [[c.numberFormat]];
            rg.values = [[safeValue(c.newValue)]];
          });
          return ctx.sync();
        });
      })(changes.slice(i, i + 400));
    }
    return chain;
  }

  function stampOf(now) { return dateStr(now) + " " + hhmm(now); }

  function appendLog(ctx, name, rows) {
    var sheets = ctx.workbook.worksheets;
    var ws = sheets.getItemOrNullObject(name);
    ws.load("name");
    return ctx.sync().then(function () {
      var fresh = ws.isNullObject;
      if (fresh) ws = sheets.add(name);
      var used = fresh ? null : ws.getUsedRangeOrNullObject(true);
      if (used) used.load("rowIndex,rowCount");
      return ctx.sync().then(function () {
        if (!fresh && !used.isNullObject) return used.rowIndex + used.rowCount;
        var hdr = ws.getRange("A1:I1");
        hdr.values = [["Run time", "Algorithm", "Sheet", "Row", "Key (CDF)", "Column", "Old value", "New value", "Note"]];
        hdr.format.font.bold = true;
        hdr.format.fill.color = "#13294B";
        hdr.format.font.color = "#FFFFFF";
        ws.freezePanes.freezeRows(1);
        return 1;
      });
    }).then(function (startRow) {
      for (var i = 0; i < rows.length; i += 2000) {
        var chunk = rows.slice(i, i + 2000);
        var rg = ws.getRangeByIndexes(startRow + i, 0, chunk.length, 9);
        rg.numberFormat = chunk.map(function (r) { return r.map(function () { return "@"; }); });
        rg.values = chunk;
      }
      ws.getRangeByIndexes(0, 0, startRow + rows.length, 9).format.autofitColumns();
      return ctx.sync();
    });
  }

  function writeLog(ctx, name, results, now) {
    var stamp = stampOf(now);
    return Promise.resolve().then(function () {
      var rows = [];
      results.forEach(function (res) {
        res.changes.forEach(function (c) {
          rows.push([stamp, res.algo.name, res.baseName, String(c.row), c.key, c.colLetter,
            show(c.oldValue, c.oldFormat), show(c.newValue, c.numberFormat || c.oldFormat), c.onChange ? E.LOG_NOTES.auto : ""]);
        });
        res.stats.unmatched.forEach(function (u) {
          rows.push([stamp, res.algo.name, res.updateName, String(u.updateRow), u.key, "", "", "", E.LOG_NOTES.notFound]);
        });
      });
      if (!rows.length) rows.push([stamp, "", "", "", "", "", "", "", E.LOG_NOTES.none]);
      return appendLog(ctx, name, rows);
    });
  }

  // ---------------- results UI ----------------
  function renderResults(results, applied) {
    var total = 0, html = "";
    results.forEach(function (res) {
      var s = res.stats;
      total += res.changes.length;
      html += '<div class="res-block">';
      html += '<div class="res-title">' + esc(res.algo.name) + "</div>";
      html += '<p class="hint">Base: <b>' + esc(res.baseName) + "</b> · Update: <b>" + esc(res.updateName) + "</b></p>";
      html += '<div class="tiles">' +
        tile(s.uniqueUpdateKeys, "IDs in update") +
        tile(s.matchedKeys, "Found in base") +
        tile(s.unmatched.length, "Not found", s.unmatched.length ? "warnt" : "") +
        tile(s.changedRows, applied ? "Rows updated" : "Rows to change", "hl") +
        tile(s.changedCells, applied ? "Cells updated" : "Cells to change", "hl") +
        tile(s.matchedRows - s.changedRows, "Unchanged") +
        "</div>";

      if (applied && res.changes.length) { html += '<div class="note ok">Update written ✓' + (state.rules.log ? " — every change and its previous value is saved in sheet <b>" + esc(state.rules.log) + "</b>." : "") + "</div>"; markStep(2, "Written"); }
      if (!res.changes.length) html += '<div class="note ok">Nothing to update — the base sheet already matches the update file.</div>';

      if (s.duplicateUpdateKeys.length) {
        html += '<div class="note warn">IDs that appear more than once in the update file (the last row was used):' + chips(s.duplicateUpdateKeys) + "</div>";
      }
      if (s.unmatched.length) {
        html += '<div class="note warn">' + s.unmatched.length + " ID" + (s.unmatched.length > 1 ? "s" : "") + " from the update file " + (s.unmatched.length > 1 ? "were" : "was") + " not found in the base sheet:" +
          chips(s.unmatched.map(function (u) { return u.key; })) + "</div>";
      }
      if (res.changes.length) html += changesTable(res.changes);
      html += "</div>";
    });

    if (!applied && total > 0) {
      html += '<div class="confirm go"><p><b>' + total + "</b> cell" + (total > 1 ? "s" : "") + " will be written. Save the workbook first — Ctrl+Z may not undo this" +
        (state.rules.log ? ", but the log keeps the previous values" : "") + ".</p>" +
        '<div class="row"><button id="applyBtn" class="btn ok">Apply ' + total + " change" + (total > 1 ? "s" : "") + "</button>" +
        '<button id="cancelApply" class="btn ghost">Cancel</button></div></div>';
    }
    $("results").innerHTML = html;
    var ab = $("applyBtn");
    if (ab) ab.addEventListener("click", applyChanges);
    var cb = $("cancelApply");
    if (cb) cb.addEventListener("click", function () { $("results").innerHTML = ""; state.lastPreview = null; });
    $("results").scrollIntoView({ behavior: "smooth", block: "start" });
  }

  function changesTable(changes) {
    var rows = changes.slice(0, MAX_ROWS_SHOWN).map(function (c) {
      return '<tr class="' + (c.onChange ? "auto" : "") + '"><td>' + c.row + '</td><td class="v k">' + esc(c.key) + "</td><td>" + c.colLetter +
        '</td><td class="v old">' + esc(show(c.oldValue, c.oldFormat) || "(empty)") + '</td><td class="v new">' + esc(show(c.newValue, c.numberFormat || c.oldFormat) || "(empty)") + "</td></tr>";
    }).join("");
    var more = changes.length > MAX_ROWS_SHOWN ? '<div class="more">and ' + (changes.length - MAX_ROWS_SHOWN) + " more changes…</div>" : "";
    return '<div class="tbl-wrap"><table><thead><tr><th>Row</th><th>Key</th><th>Col</th><th>Before</th><th>After</th></tr></thead><tbody>' +
      rows + "</tbody></table>" + more + "</div>";
  }

  function tile(n, label, cls) { return '<div class="tile ' + (cls || "") + '"><div class="n">' + n + '</div><div class="l">' + label + "</div></div>"; }
  function chips(list) {
    var shown = list.slice(0, 60).map(function (k) { return '<span class="chip">' + esc(k) + "</span>"; }).join("");
    return '<div class="chips">' + shown + (list.length > 60 ? '<span class="chip">+' + (list.length - 60) + "</span>" : "") + "</div>";
  }

  // ================= new row (via claude.ai) =================
  var PROJECT_KEY = "cdfUpdater.claudeProject.v1";

  function showNewRowNote(msg, kind) {
    $("newRowForm").innerHTML = '<div class="note ' + (kind || "bad") + '">' + esc(msg) + "</div>";
  }

  function fetchPrompt() {
    return fetch("ai-prompt.txt?v=" + Date.now(), { cache: "no-store" })
      .then(function (r) { return r.ok ? r.text() : null; })
      .catch(function () { return null; })
      .then(function (t) {
        state.promptBase = t && t.trim() ? t : CdfPrompt.DEFAULT_PROMPT;
        state.prompt = CdfPrompt.build(state.promptBase);
        renderProject();
        return state.prompt;
      });
  }

  // ---- Claude Project: link + which instructions version was copied into it ----
  function loadProject() {
    var p = null;
    try { p = JSON.parse(localStorage.getItem(PROJECT_KEY) || "null"); } catch (e) { p = null; }
    p = p || {};
    if (!p.url && state.rules && state.rules.claudeProject) { p.url = state.rules.claudeProject; p.fromRules = true; }
    return p;
  }
  function saveProjectRecord(p) {
    try { localStorage.setItem(PROJECT_KEY, JSON.stringify({ url: p.url || null, version: p.version || null, at: p.at || null })); } catch (e) { /* ignore */ }
  }

  function projectUrlOk(u) { return /^https:\/\/claude\.ai\/project\/[A-Za-z0-9-]+/.test(String(u || "").trim()); }

  function renderProject() {
    var p = loadProject(), cur = state.prompt && state.prompt.version;
    var has = !!p.url;
    $("nrProject").classList.toggle("hidden", !has);
    $("nrPrompt").classList.toggle("hidden", has);
    $("projectSetupTitle").textContent = has ? "Claude Project settings" : "Skip the prompt: set up a Claude Project";
    // without a project the set-up link sits right under the prompt step; with one it moves to the end
    var setup = $("projectSetup"), body = setup.parentNode;
    if (has && setup.nextElementSibling) body.appendChild(setup);
    if (!has && setup.previousElementSibling !== $("nrPrompt")) body.insertBefore(setup, $("nrPrompt").nextElementSibling);
    if (has && document.activeElement !== $("projectUrl")) $("projectUrl").value = p.url;

    var info = "";
    if (has) {
      info += '<dl class="kv"><dt>Project</dt><dd>' + esc(p.url.replace(/^https:\/\/claude\.ai\//, "")) + (p.fromRules ? " (from rules.txt)" : "") + "</dd>" +
        "<dt>Instructions</dt><dd>" + (p.version ? 'version <span class="code">' + esc(p.version) + "</span>" + (p.at ? ", copied " + esc(p.at) : "") : "not copied from this computer") + "</dd>" +
        (cur ? '<dt>Current</dt><dd>version <span class="code">' + esc(cur) + "</span></dd>" : "") + "</dl>" +
        (p.fromRules ? "" : '<div class="row"><button id="forgetProject" class="btn ghost small" type="button">Stop using the project</button></div>');
    }
    $("projectInfo").innerHTML = info;
    var fp = $("forgetProject");
    if (fp) fp.addEventListener("click", function () {
      var rec = loadProject(); rec.url = null; saveProjectRecord(rec);
      $("projectUrl").value = ""; $("projectUrlMsg").textContent = ""; renderProject();
    });

    // the prompt changed after the instructions were copied into the project
    var stale = has && cur && p.version && p.version !== cur;
    $("projectNote").innerHTML = stale
      ? '<div class="note warn"><b>The instructions changed</b> (ai-prompt.txt was edited). Replace the project\'s instructions so new replies follow them.' +
        '<div><button id="copyInstr2" class="btn small" type="button"><svg class="ico"><use href="#i-copy"/></svg>Copy project instructions</button></div></div>'
      : "";
    var c2 = $("copyInstr2");
    if (c2) c2.addEventListener("click", function () { copyInstructions(c2); });
  }

  function copyInstructions(btn) {
    fetchPrompt().then(function (pr) {
      return copyText(pr.text).then(function (ok) {
        var msg = $("copyInstrMsg");
        if (!ok) { msg.className = "msg bad"; msg.textContent = "Copy failed here — click the button again."; return; }
        var rec = loadProject();
        rec.version = pr.version; rec.at = dateStr(new Date()) + " " + hhmm(new Date());
        saveProjectRecord(rec);
        msg.className = "msg ok";
        msg.textContent = "Copied ✓ (version " + pr.version + "). In the project: Instructions → paste → Save.";
        if (btn && btn.id === "copyInstr2") { $("projectSetup").open = true; }
        renderProject();
      });
    });
  }

  function saveProject() {
    var u = $("projectUrl").value.trim(), msg = $("projectUrlMsg");
    if (!projectUrlOk(u)) {
      msg.className = "msg bad";
      msg.textContent = "That isn't a project address. Open the project in Claude and copy the address from the browser — it starts with https://claude.ai/project/";
      return;
    }
    var rec = loadProject(); rec.url = u.split(/[?#]/)[0]; saveProjectRecord(rec);
    msg.className = "msg ok";
    msg.textContent = rec.version ? "Saved ✓ — from now on: Open the Claude project, drag in the PDFs, send." : "Saved ✓ — also copy the instructions into the project (step 2 above).";
    renderProject();
  }

  function openUrl(url) {
    try {
      if (window.Office && Office.context && Office.context.ui && Office.context.ui.openBrowserWindow) { Office.context.ui.openBrowserWindow(url); return; }
    } catch (e) { /* fall back */ }
    window.open(url, "_blank");
  }
  function openClaude() { openUrl("https://claude.ai/new"); }
  function openProject() { var p = loadProject(); if (p.url) openUrl(p.url); }

  function newRowSheet(ctx) {
    var s = state.rules.newRow.sheet;
    var ws = s.first ? ctx.workbook.worksheets.getFirst(true) : ctx.workbook.worksheets.getItemOrNullObject(s.name);
    ws.load("name");
    return ctx.sync().then(function () {
      if (!s.first && ws.isNullObject) throw new Error("Sheet \"" + s.name + "\" (NEWROW_SHEET) was not found");
      return ws;
    });
  }

  function copyPrompt() {
    fetchPrompt().then(function (pr) {
      return copyText(pr.text).then(function (ok) {
        var m = $("copyPromptMsg");
        m.className = ok ? "msg ok" : "msg bad";
        m.textContent = ok ? "Prompt copied ✓ — paste it into Claude together with the files." : "Copy failed here — click the button again.";
      });
    });
  }

  // ---- reading the reply ----
  function onPaste(ev) {
    var text = ev.clipboardData && ev.clipboardData.getData("text/plain");
    if (!text) { setTimeout(function () { refreshButtons(); if ($("pasteBox").value.trim()) parsePasted(); }, 0); return; }
    ev.preventDefault();
    $("pasteBox").value = text;   // a paste is always a new reply: it replaces what was in the box
    refreshButtons();
    parsePasted();
  }

  function parsePasted() {
    if (state.busy) return;
    var res = E.parseAiPaste($("pasteBox").value);
    if (res.error) {
      $("pasteBox").classList.remove("read");
      $("pasteMsg").textContent = "";
      if (res.kind === "empty" || res.kind === "cut") showNewRowNote(res.error, "bad");
      else showFormatHelp(res);
      return;
    }
    state.draft = null; state.queue = []; state.written = [];
    var items = res.items.map(function (it) { return { fields: it, main: it.main_file || null, att: it.att_file || null }; });
    (state.prompt ? Promise.resolve(state.prompt) : fetchPrompt()).then(function () {
      state.promptWarn = versionWarning(items);
      $("pasteBox").classList.add("read");
      $("pasteMsg").className = "msg ok";
      $("pasteMsg").textContent = "Read ✓ " + items.length + " CDF" + (items.length > 1 ? "s" : "");
      if (items.length === 1) { state.queue = items; state.queueTotal = 1; nextInQueue(); }
      else planBatch(items);
    });
  }

  // Claude answered in another format (a table row, plain text, broken JSON): say why, and offer
  // a message that makes it redo the answer in the same chat + the instructions for next time.
  function showFormatHelp(res) {
    var p = loadProject(), inProject = !!p.url;
    var html = '<div class="note bad fix"><b>' + esc(res.error) + "</b>" +
      '<p class="fix-why">' + (inProject
        ? "Usually the project's instructions are missing or are older ones, or the chat was started outside the project."
        : "Usually the prompt wasn't pasted into the chat, or only part of it was.") + "</p>" +
      '<div class="fix-step"><span class="fix-n">Now</span><div>' +
        '<button id="copyFixup" class="btn small" type="button"><svg class="ico"><use href="#i-copy"/></svg>Copy a fix-up message</button>' +
        '<p>Paste it in the same chat. Claude redoes the answer from the PDFs already there — no need to upload them again.</p></div></div>' +
      (inProject
        ? '<div class="fix-step"><span class="fix-n">Next time</span><div>' +
            '<button id="copyFixInstr" class="btn small" type="button"><svg class="ico"><use href="#i-copy"/></svg>Copy project instructions</button>' +
            "<p>In the project: Instructions → select all → paste → Save. Replace everything that was there, then start a new chat in the project.</p></div></div>"
        : "") +
      '<div id="fixMsg" class="msg"></div></div>';
    $("newRowForm").innerHTML = html;
    $("copyFixup").addEventListener("click", function () {
      fetchPrompt().then(function () {
        return copyText(CdfPrompt.fixupMessage(state.promptBase));
      }).then(function (ok) {
        var m = $("fixMsg"); m.className = ok ? "msg ok" : "msg bad";
        m.textContent = ok ? "Copied ✓ — paste it in the chat, send, then copy the new code block here." : "Copy failed here — click the button again.";
      });
    });
    var ci = $("copyFixInstr");
    if (ci) ci.addEventListener("click", function () {
      $("projectSetup").open = true;
      copyInstructions(ci);
      var m = $("fixMsg"); m.className = "msg ok";
      m.textContent = "Instructions copied ✓ — replace the project's instructions with them.";
    });
  }

  function versionWarning(items) {
    var cur = state.prompt && state.prompt.version;
    if (!cur) return null;
    var seen = {};
    items.forEach(function (it) { seen[it.fields.prompt_version || ""] = true; });
    var vs = Object.keys(seen);
    if (vs.length === 1 && vs[0] === cur) return null;
    var old = vs.filter(function (v) { return v && v !== cur; });
    if (old.length) return "This reply was made with older instructions (version " + old.join(", ") + "; current " + cur + "). " +
      "Check the fields carefully, and update the project's instructions so the next reply follows the current ones.";
    return "This reply has no instructions version, so it may follow older instructions. If it came from the Claude Project, copy the project instructions again.";
  }

  function versionNoteHtml() {
    if (!state.promptWarn) return "";
    return '<div class="note warn">' + esc(state.promptWarn) +
      '<div><button class="btn small js-copy-instr" type="button"><svg class="ico"><use href="#i-copy"/></svg>Copy project instructions</button></div></div>';
  }
  function wireVersionNote(root) {
    Array.prototype.forEach.call(root.querySelectorAll(".js-copy-instr"), function (b) {
      b.addEventListener("click", function () { $("projectSetup").open = true; copyInstructions(b); b.textContent = "Copied ✓ — paste it into the project"; });
    });
  }

  // ---- where each CDF goes: one read of the sheet for the whole batch ----
  function planEntries(ctx, items) {
    var nr = state.rules.newRow, algo = state.rules.algorithms[0], now = new Date();
    var builts = items.map(function (item) {
      return E.buildNewRowCells(nr, {
        ai: item.fields, mainName: item.main, attName: item.att,
        now: now, dateFormat: algo ? algo.dateFormat : "dd/mm/yyyy"
      });
    });
    return newRowSheet(ctx).then(function (ws) {
      var cols = nr.cells.map(function (c) { return c.idx; }).concat([nr.key.idx]);
      if (nr.match) cols.push(nr.match.idx);
      var width = Math.max.apply(null, cols) + 1;
      var used = ws.getUsedRangeOrNullObject(true);
      used.load("rowIndex,rowCount");
      var hdr = ws.getRangeByIndexes(nr.headerRow - 1, 0, 1, width);
      hdr.load("values");
      return ctx.sync().then(function () {
        var rows = used.isNullObject ? 0 : used.rowIndex + used.rowCount;
        var values = [], formats = [], chain = Promise.resolve();
        for (var r = 0; r < rows; r += 2000) {
          (function (start) {
            chain = chain.then(function () {
              var rg = ws.getRangeByIndexes(start, 0, Math.min(2000, rows - start), width);
              rg.load("values,numberFormat");
              return ctx.sync().then(function () {
                Array.prototype.push.apply(values, rg.values);
                Array.prototype.push.apply(formats, rg.numberFormat);
              });
            });
          })(r);
        }
        return chain.then(function () {
          var targets = E.planNewRowTargets({
            keys: values.map(function (v) { return v[nr.key.idx]; }),
            match: nr.match ? values.map(function (v) { return v[nr.match.idx]; }) : null,
            rows: values, formats: formats
          }, nr.startRow, builts.map(function (b) { return b.cdfFull || b.cdf; }));
          return {
            ws: ws, wsName: ws.name, headers: hdr.values[0], hasMatch: !!nr.match,
            entries: items.map(function (item, i) {
              var t = targets[i];
              return { item: item, built: builts[i], target: t, existing: t.existingValues, existingFmt: t.existingFormats };
            })
          };
        });
      });
    });
  }

  // where the CDF goes, in words (short for the list, long for the form)
  function targetText(t, short) {
    if (t.existing) return short ? "Updates row " + t.row : "already exists in row " + t.row + " — that row will be updated";
    if (t.prepared) return short ? "Row " + t.row + " (prepared)" : "goes to row " + t.row + ", the row prepared for it (its PDF links are already there)";
    if (t.insert) return short ? "New row " + t.row + ", after " + t.after : "goes to a new row " + t.row + ", inserted right after " + t.after + " and its solution rows";
    return short ? "New row " + t.row : "goes to a new row " + t.row + " at the end";
  }
  function targetFlags(t, hasMatch) {
    var f = [];
    if (!t.existing && !t.prepared && !t.insert && hasMatch) f.push("No prepared row with this CDF number in column B — it is added after the last row.");
    if (!t.existing && !t.insert && t.after) f.push("Revision of " + t.after + ", which is the last CDF in the sheet — added at the end.");
    return f;
  }

  // write one entry's cells (no sync) and return its log rows
  function writeEntryCells(ws, wsName, en, stamp) {
    var log = [], t = en.target, at = t.writeRow || t.row;
    if (t.insert) ws.getRangeByIndexes(at - 1, 0, 1, 1).getEntireRow().insert("Down");
    en.built.cells.forEach(function (c) {
      if (E.isEmpty(c.value)) return;
      var rg = ws.getRangeByIndexes(at - 1, c.idx, 1, 1);
      if (c.link) {
        rg.hyperlink = { address: c.link.address, textToDisplay: c.link.textToDisplay };
      } else {
        if (c.format) rg.numberFormat = [[c.format]];
        rg.values = [[safeValue(c.value)]];
      }
      log.push([stamp, E.LOG_NOTES.newRow, wsName, String(t.row), en.built.cdfFull || en.built.cdf, c.col,
        show(en.existing[c.idx], en.existingFmt[c.idx]), c.display, t.insert ? E.LOG_NOTES.newRow + " (inserted)" : E.LOG_NOTES.newRow]);
    });
    return log;
  }

  // ---- several CDFs: overview first, then write all or check one by one ----
  function planBatch(items) {
    state.busy = true; refreshButtons();
    setStatus("Finding the place of each CDF in the workbook…");
    Excel.run(function (ctx) { return planEntries(ctx, items); })
      .then(function (plan) { state.batch = { items: items, plan: plan, on: items.map(function () { return true; }) }; renderBatch(); })
      .catch(function (e) { showNewRowNote(errText(e), "bad"); })
      .then(function () { state.busy = false; clearStatus(); refreshButtons(); });
  }

  function renderBatch() {
    var b = state.batch, plan = b.plan;
    var nOn = b.on.filter(Boolean).length;
    var html = versionNoteHtml() +
      '<div class="batch"><div class="batch-head"><b>' + plan.entries.length + " CDFs in the reply</b><span>sheet " + esc(plan.wsName) + "</span></div>" +
      '<ul class="batch-list">' + plan.entries.map(function (en, i) {
        var t = en.target, flags = targetFlags(t, plan.hasMatch).concat(en.item.fields.flags || [], en.built.warnings);
        if (t.dupOf != null) flags = ["Same CDF as number " + (t.dupOf + 1) + " in this list — both write to row " + t.row].concat(flags);
        return '<li class="' + (b.on[i] ? "" : "off") + '"><label class="b-row">' +
          '<input type="checkbox" data-b="' + i + '"' + (b.on[i] ? " checked" : "") + ">" +
          '<span class="b-cdf">' + esc(en.built.cdfFull || en.built.cdf || "(no CDF number)") + "</span>" +
          '<span class="b-target ' + (t.existing ? "upd" : t.insert ? "ins" : "new") + '">' + esc(targetText(t, true)) + "</span>" +
          '<span class="b-desc">' + esc(en.item.fields.description || "") + "</span>" +
          (flags.length ? '<ul class="b-flags">' + flags.map(function (f) { return "<li>" + esc(f) + "</li>"; }).join("") + "</ul>" : "") +
          "</label></li>";
      }).join("") + "</ul>" +
      '<div class="confirm go"><p>Empty fields are not written and don\'t erase existing values. Every cell goes to ' + esc(state.rules.log || "the log") + ".</p>" +
      '<div class="row wrap"><button id="writeAllBtn" class="btn ok"' + (nOn ? "" : " disabled") + ">Write " + nOn + " row" + (nOn === 1 ? "" : "s") + "</button>" +
      '<button id="reviewBtn" class="btn"' + (nOn ? "" : " disabled") + ">Check one by one</button>" +
      '<button id="cancelBatch" class="btn ghost">Cancel</button></div></div></div>';
    $("newRowForm").innerHTML = html;
    wireVersionNote($("newRowForm"));
    Array.prototype.forEach.call($("newRowForm").querySelectorAll("[data-b]"), function (el) {
      el.addEventListener("change", function () { b.on[+el.getAttribute("data-b")] = el.checked; renderBatch(); });
    });
    $("writeAllBtn").addEventListener("click", function () { writeBatch(selectedBatchItems()); });
    $("reviewBtn").addEventListener("click", function () {
      state.queue = selectedBatchItems(); state.queueTotal = state.queue.length; state.batch = null; nextInQueue();
    });
    $("cancelBatch").addEventListener("click", cancelNewRows);
  }

  function selectedBatchItems() {
    var b = state.batch;
    return b.items.filter(function (_, i) { return b.on[i]; });
  }

  function cancelNewRows() {
    state.draft = null; state.queue = []; state.batch = null; state.written = [];
    $("newRowForm").innerHTML = "";
    $("pasteBox").value = ""; $("pasteBox").classList.remove("read"); $("pasteMsg").textContent = "";
    refreshButtons();
  }

  function writeBatch(items) {
    if (state.busy || !items.length) return;
    state.busy = true; refreshButtons();
    setStatus("Writing " + items.length + " rows…");
    var stamp = stampOf(new Date());
    Excel.run(function (ctx) {
      return planEntries(ctx, items).then(function (plan) {   // re-plan: the sheet may have changed meanwhile
        var log = [], done = [];
        plan.entries.forEach(function (en) {
          var rows = writeEntryCells(plan.ws, plan.wsName, en, stamp);
          log = log.concat(rows);
          done.push({ cdf: en.built.cdfFull || en.built.cdf, row: en.target.row, cells: rows.length, existing: en.target.existing });
        });
        return ctx.sync().then(function () {
          return state.rules.log && log.length ? appendLog(ctx, state.rules.log, log) : null;
        }).then(function () { return done; });
      });
    }).then(function (done) { state.batch = null; showWritten(done); })
      .catch(function (e) { showNewRowNote(errText(e), "bad"); })
      .then(function () { state.busy = false; clearStatus(); refreshButtons(); });
  }

  function showWritten(done) {
    $("pasteBox").value = ""; $("pasteBox").classList.remove("read"); $("pasteMsg").textContent = "";
    var one = done.length === 1 ? done[0] : null;
    $("newRowForm").innerHTML = '<div class="note ok">' + (one
      ? one.cells + " cells written in row " + one.row + " (CDF " + esc(one.cdf) + ")"
      : done.length + " CDFs written: " + done.map(function (w) { return esc(w.cdf) + " (row " + w.row + ")"; }).join(", ")) + " ✓" +
      '<div><button id="mailNewRow" class="btn small">Send by email</button></div></div>';
    markStep(3, done.length + (done.length === 1 ? " row" : " rows") + " written");
    $("mailNewRow").addEventListener("click", function () {
      $("mailKind").value = "newrows"; $("mailSubject").value = defaultSubject();
      $("mailCard").open = true; $("mailCard").scrollIntoView({ behavior: "smooth" });
    });
  }

  // ---- one draft at a time from the queue ----
  function nextInQueue() {
    if (!state.queue || !state.queue.length) return;
    var item = state.queue.shift();
    state.busy = true; refreshButtons();
    setStatus("Finding the place in the workbook…");
    prepareDraft(item).then(function () { /* rendered */ })
      .catch(function (e) { showNewRowNote(errText(e), "bad"); })
      .then(function () { state.busy = false; clearStatus(); refreshButtons(); });
  }

  function prepareDraft(item) {
    return Excel.run(function (ctx) {
      return planEntries(ctx, [item]).then(function (plan) {
        var en = plan.entries[0];
        return { ws: plan.wsName, headers: plan.headers, hasMatch: plan.hasMatch, target: en.target, existing: en.existing, existingFmt: en.existingFmt, built: en.built };
      });
    }).then(function (place) {
      state.draft = { built: place.built, place: place, item: item };
      renderDraft();
    });
  }

  function queuePos() {
    var total = state.queueTotal || 1, left = state.queue ? state.queue.length : 0;
    return { n: total - left, total: total };
  }

  function renderDraft() {
    var d = state.draft, p = d.place, item = d.item, q = queuePos();
    var flags = targetFlags(p.target, p.hasMatch).concat(item.fields.flags || [], d.built.warnings);
    var html = versionNoteHtml();
    if (q.total > 1) html += '<div class="queue"><span>CDF ' + q.n + " of " + q.total + "</span><b>" + esc(d.built.cdf) + "</b></div>" +
      '<div class="qbar"><span style="width:' + Math.round(100 * q.n / q.total) + '%"></span></div>';
    html += '<div class="nr-target' + (p.target.existing ? " upd" : "") + '">CDF ' + esc(d.built.cdfFull || d.built.cdf) + " " +
      esc(targetText(p.target, false)) + " (sheet " + esc(p.ws) + ")</div>";
    if (flags.length) html += '<div class="note warn">Check:<ul>' + flags.map(function (f) { return "<li>" + esc(f) + "</li>"; }).join("") + "</ul></div>";

    d.built.cells.forEach(function (c, i) {
      var header = E.cleanHeader(p.headers[c.idx]) || "no header";
      var old = p.existing[c.idx];
      var oldShow = show(old, p.existingFmt[c.idx]);
      html += '<div class="nr-field"><div class="nr-label"><span class="col">' + c.col + '</span><span class="hdr">' + esc(header) +
        '</span><span class="src">' + esc(c.label) + "</span></div>";
      if (c.kind === "ai" && c.label.indexOf("DESCRIPTION") >= 0) {
        html += '<textarea class="inp" dir="ltr" data-i="' + i + '">' + esc(c.display) + "</textarea>";
      } else if (c.display.indexOf("\n") >= 0) {   // name + phone on two lines
        html += '<textarea class="inp short" dir="ltr" rows="2" data-i="' + i + '">' + esc(c.display) + "</textarea>";
      } else {
        html += '<input class="inp" dir="ltr" data-i="' + i + '" value="' + esc(c.display) + '"' +
          (c.kind === "link" ? ' placeholder="File name, e.g. A330-1-NNN.-.pdf"' : "") + ">";
        if (c.kind === "link" && c.prefix) html += '<div class="nr-link">Link: ' + esc(c.prefix) + "&lt;file name&gt;</div>";
      }
      if (!E.isEmpty(old) && oldShow !== c.display) html += '<div class="nr-old">Currently in the cell: ' + esc(oldShow) + " — will be replaced</div>";
      html += "</div>";
    });

    var more = state.queue && state.queue.length;
    html += '<div class="confirm go" style="margin-top:12px"><p>Empty fields are not written and don\'t erase existing values. Every cell goes to ' + esc(state.rules.log || "the log") + ".</p>" +
      '<div class="row wrap"><button id="writeRowBtn" class="btn ok">' + (p.target.existing ? "Update row " : p.target.insert ? "Insert row " : "Add row ") + p.target.row + "</button>" +
      (more ? '<button id="skipRowBtn" class="btn">Skip</button>' : "") +
      '<button id="cancelRowBtn" class="btn ghost">' + (more ? "Cancel all" : "Cancel") + "</button></div></div>";
    $("newRowForm").innerHTML = html;
    wireVersionNote($("newRowForm"));
    $("writeRowBtn").addEventListener("click", writeDraft);
    var sk = $("skipRowBtn");
    if (sk) sk.addEventListener("click", function () { state.draft = null; nextInQueue(); });
    $("cancelRowBtn").addEventListener("click", cancelNewRows);
  }

  function writeDraft() {
    var d = state.draft;
    if (!d || state.busy) return;
    // take the edited values from the form
    var inputs = $("newRowForm").querySelectorAll("[data-i]");
    Array.prototype.forEach.call(inputs, function (el) {
      var c = d.built.cells[+el.getAttribute("data-i")];
      var v = el.value.replace(/\r/g, "").trim();
      if (c.kind === "today" || c.kind === "now") {
        var serial = E.parseDateString(v);
        if (serial != null && v !== c.display) { c.value = serial; c.display = v; }
        else if (serial == null) { c.value = v; c.format = null; c.display = v; }
      } else if (c.kind === "link") {
        c.value = v; c.display = v;
        c.link = v ? { address: (c.prefix || "") + v, textToDisplay: v } : null;
      } else if (v !== c.display) { c.value = v; c.display = v; }
    });

    state.busy = true; refreshButtons();
    setStatus("Writing to the workbook…");
    var stamp = stampOf(new Date()), row = d.place.target.row;
    var entry = { built: d.built, target: d.place.target, existing: d.place.existing, existingFmt: d.place.existingFmt };
    Excel.run(function (ctx) {
      return newRowSheet(ctx).then(function (ws) {
        var log = writeEntryCells(ws, ws.name, entry, stamp);
        return ctx.sync().then(function () {
          return state.rules.log && log.length ? appendLog(ctx, state.rules.log, log) : null;
        }).then(function () { return log.length; });
      });
    }).then(function (n) {
      state.draft = null;
      state.written = (state.written || []).concat([{ cdf: d.built.cdfFull || d.built.cdf, row: row, cells: n }]);
      if (state.queue && state.queue.length) {
        state.busy = false;
        nextInQueue();
        return true; // next draft is loading
      }
      var done = state.written; state.written = [];
      showWritten(done);
      return false;
    }).catch(function (e) { showNewRowNote(errText(e), "bad"); return false; })
      .then(function (loadingNext) { if (!loadingNext) { state.busy = false; clearStatus(); refreshButtons(); } });
  }

  // ================= email =================
  var MAIL_KINDS = [
    { id: "changes", label: "Changes from the last run", subject: "CDF status update" },
    { id: "notfound", label: "CDFs not found in the last run", subject: "CDFs not found in tracker" },
    { id: "newrows", label: "New rows added today", subject: "New CDF entries" }
  ];

  function renderMailKinds() {
    var sel = $("mailKind"), keep = sel.value;
    var opts = MAIL_KINDS.map(function (k) { return '<option value="' + k.id + '">' + esc(k.label) + "</option>"; });
    (state.rules ? state.rules.reports : []).forEach(function (r, i) {
      opts.push('<option value="report:' + i + '">Report: ' + esc(r.name) + "</option>");
    });
    sel.innerHTML = opts.join("");
    if (keep && sel.querySelector('option[value="' + keep + '"]')) sel.value = keep;
    if (!$("mailSubject").value || !keep) $("mailSubject").value = defaultSubject();
  }

  function defaultSubject() {
    var v = $("mailKind").value, today = dateStr(new Date());
    if (v.indexOf("report:") === 0) {
      var r = state.rules && state.rules.reports[+v.split(":")[1]];
      return (r ? r.name : "Report") + " – " + today;
    }
    var k = MAIL_KINDS.filter(function (x) { return x.id === v; })[0];
    return (k ? k.subject : "CDF") + " – " + today;
  }

  function loadGroups() {
    var local = [];
    try { local = JSON.parse(localStorage.getItem(GROUPS_KEY) || "[]"); } catch (e) { local = []; }
    var fromRules = (state.rules ? state.rules.emailGroups : []).map(function (g) { return { name: g.name, addresses: g.addresses, rules: true }; });
    return fromRules.concat(local);
  }

  function renderGroups() {
    var groups = loadGroups(), box = $("groupChips");
    if (!groups.length) { box.innerHTML = '<span class="hint" style="margin:0">No saved groups yet.</span>'; return; }
    box.innerHTML = groups.map(function (g, i) {
      return '<span class="chip" data-g="' + i + '" title="' + esc(g.addresses.join("; ")) + '">' + esc(g.name) +
        (g.rules ? "" : '<span class="x" data-del="' + i + '">✕</span>') + "</span>";
    }).join("");
    Array.prototype.forEach.call(box.querySelectorAll(".chip"), function (el) {
      el.addEventListener("click", function (ev) {
        var g = groups[+el.getAttribute("data-g")];
        if (ev.target.hasAttribute("data-del")) { deleteLocalGroup(g.name); return; }
        var cur = splitAddrs($("mailTo").value);
        g.addresses.forEach(function (a) { if (cur.indexOf(a) < 0) cur.push(a); });
        $("mailTo").value = cur.join("; ");
        if (state.mail) renderMail();
      });
    });
  }

  function splitAddrs(s) { return String(s || "").split(/[;,\s]+/).filter(Boolean); }

  function saveLocalGroup() {
    var name = $("groupName").value.trim(), addrs = splitAddrs($("mailTo").value);
    if (!name || !addrs.length) { $("groupName").placeholder = "Needs a name and recipients in To"; return; }
    var local = loadGroups().filter(function (g) { return !g.rules && g.name !== name; });
    local.push({ name: name, addresses: addrs });
    try { localStorage.setItem(GROUPS_KEY, JSON.stringify(local)); } catch (e) { /* ignore */ }
    $("groupName").value = ""; $("saveGroupBox").classList.add("hidden");
    renderGroups();
  }

  function deleteLocalGroup(name) {
    var local = loadGroups().filter(function (g) { return !g.rules && g.name !== name; });
    try { localStorage.setItem(GROUPS_KEY, JSON.stringify(local)); } catch (e) { /* ignore */ }
    renderGroups();
  }

  function readHeaderMap(ctx, ws, headerRow, width) {
    var rg = ws.getRangeByIndexes(headerRow - 1, 0, 1, width);
    rg.load("values");
    return ctx.sync().then(function () {
      return function (colLetter) {
        var h = E.cleanHeader(rg.values[0][E.colToIndex(colLetter)]);
        return h || colLetter;
      };
    });
  }

  function buildMail() {
    if (state.busy) return;
    var kind = $("mailKind").value, title = $("mailSubject").value.trim() || defaultSubject();
    state.busy = true; refreshButtons();
    $("mailPreview").innerHTML = "";
    setStatus("Preparing email…");
    Excel.run(function (ctx) {
      var headerRow = state.rules.newRow.headerRow || 1;
      if (kind.indexOf("report:") === 0) {
        var rep = state.rules.reports[+kind.split(":")[1]];
        var ws = rep.sheet.first ? ctx.workbook.worksheets.getFirst(true) : ctx.workbook.worksheets.getItemOrNullObject(rep.sheet.name);
        ws.load("name");
        return ctx.sync().then(function () {
          if (!rep.sheet.first && ws.isNullObject) throw new Error("Sheet \"" + rep.sheet.name + "\" was not found");
          return readSheet(ctx, ws);
        }).then(function (data) {
          var width = Math.max.apply(null, rep.columns.map(function (c) { return c.idx; })) + 1;
          return readHeaderMap(ctx, ws, headerRow, width).then(function (hname) {
            var out = E.runReport(rep, data.values);
            var fmts = data.formats;
            return {
              title: title,
              intro: out.rows.length + " item(s).",
              sections: [{
                heading: "",
                headers: rep.columns.map(function (c) { return hname(c.col); }),
                rows: out.rows.map(function (r) {
                  return r.cells.map(function (v, j) { return show(v, (fmts[r.row - 1] || [])[rep.columns[j].idx]); });
                })
              }],
              empty: !out.rows.length
            };
          });
        });
      }

      if (!state.rules.log) throw new Error("There is no log (LOG NONE) — nothing to build the email from");
      var logWs = ctx.workbook.worksheets.getItemOrNullObject(state.rules.log);
      logWs.load("name");
      return ctx.sync().then(function () {
        if (logWs.isNullObject) throw new Error("There is no " + state.rules.log + " sheet yet — run an update or add a row first");
        return readSheet(ctx, logWs);
      }).then(function (logData) {
        var run = E.lastRunFromLog(logData.values, dateStr(new Date()));
        return newRowSheet(ctx).then(function (base) {
          return readHeaderMap(ctx, base, headerRow, 200).then(function (hname) {
            if (kind === "changes") {
              return {
                title: title, intro: run.stamp ? "Update run: " + run.stamp + " · " + run.changes.length + " change(s)." : "",
                sections: [{ heading: "", headers: ["CDF", "Field", "Before", "After"],
                  rows: run.changes.map(function (c) { return [c.key, hname(c.col), c.oldValue, c.newValue]; }) }],
                empty: !run.changes.length
              };
            }
            if (kind === "notfound") {
              return {
                title: title, intro: run.stamp ? "Update run: " + run.stamp + " · these CDFs are in the update file but not in the tracker." : "",
                sections: [{ heading: "", headers: ["CDF", "Row in update file"], rows: run.notFound.map(function (c) { return [c.key, c.row]; }) }],
                empty: !run.notFound.length
              };
            }
            var byCdf = {}, order = [];
            run.newRows.forEach(function (c) {
              if (!byCdf[c.key]) { byCdf[c.key] = []; order.push(c.key); }
              byCdf[c.key].push([hname(c.col), c.newValue]);
            });
            return {
              title: title, intro: order.length + " new / updated CDF row(s) today.",
              sections: order.map(function (k) { return { heading: k, headers: ["Field", "Value"], rows: byCdf[k] }; }),
              empty: !order.length
            };
          });
        });
      });
    }).then(function (doc) {
      if (doc.empty) { $("mailPreview").innerHTML = '<div class="note warn">Nothing to send for this output.</div>'; state.mail = null; return; }
      state.mail = E.buildEmail(doc);
      renderMail();
      markStep(4, "Draft ready");
    }).catch(function (e) { $("mailPreview").innerHTML = '<div class="note bad">' + esc(errText(e)) + "</div>"; })
      .then(function () { state.busy = false; clearStatus(); refreshButtons(); });
  }

  var MAILTO_LIMIT = 1900;

  function renderMail() {
    var m = state.mail, to = splitAddrs($("mailTo").value), cc = splitAddrs($("mailCc").value);
    var full = E.mailtoUrl(to, cc, m.subject, m.text);
    var tooLong = full.length > MAILTO_LIMIT;
    var url = tooLong ? E.mailtoUrl(to, cc, m.subject, "") : full;
    $("mailPreview").innerHTML =
      '<div class="mail-box"><div class="mail-head"><b>To:</b> ' + esc(to.join("; ") || "—") +
      (cc.length ? "<br><b>Cc:</b> " + esc(cc.join("; ")) : "") + "<br><b>Subject:</b> " + esc(m.subject) + "</div>" +
      '<div class="mail-body">' + m.html + "</div></div>" +
      (tooLong ? '<div class="note warn" style="margin-top:8px">The email is too long to open directly — the table will be copied; just paste it into the draft (Ctrl+V).</div>' : "") +
      '<div class="mail-actions">' +
      '<a id="openMail" class="btn primary" href="' + esc(url) + '">Open draft in Outlook</a>' +
      '<button id="copyMail" class="btn" type="button">Copy as table (to paste in an email)</button></div>' +
      '<div id="copyMsg" class="nr-meta"></div>';
    if (!to.length) $("copyMsg").textContent = "No recipients — you can add them in Outlook.";
    $("openMail").addEventListener("click", function () {
      if (tooLong) copyRich(m.html, m.text).then(function (ok) { $("copyMsg").textContent = ok ? "Table copied — paste it into the email body." : "Copy failed — use the copy button."; });
    });
    $("copyMail").addEventListener("click", function () {
      copyRich(m.html, m.text).then(function (ok) { $("copyMsg").textContent = ok ? "Copied ✓ — paste it into the email body (Ctrl+V)." : "Copy failed here."; });
    });
  }

  function copyText(text) {
    function legacy() {
      var ta = document.createElement("textarea");
      ta.value = text; ta.setAttribute("readonly", "");
      ta.style.cssText = "position:fixed;left:-9999px;top:0;opacity:0";
      document.body.appendChild(ta); ta.select();
      var ok = false;
      try { ok = document.execCommand("copy"); } catch (e) { ok = false; }
      document.body.removeChild(ta);
      return ok;
    }
    try {
      if (navigator.clipboard && navigator.clipboard.writeText) {
        return navigator.clipboard.writeText(text).then(function () { return true; }, function () { return legacy(); });
      }
    } catch (e) { /* fall through */ }
    return Promise.resolve(legacy());
  }

  // copy as rich text (HTML) with plain-text fallback
  function copyRich(html, text) {
    function legacy() {
      var div = document.createElement("div");
      div.contentEditable = "true";
      div.style.cssText = "position:fixed;left:-9999px;top:0;opacity:0";
      div.innerHTML = html;
      document.body.appendChild(div);
      var range = document.createRange(); range.selectNodeContents(div);
      var sel = window.getSelection(); sel.removeAllRanges(); sel.addRange(range);
      var ok = false;
      try { ok = document.execCommand("copy"); } catch (e) { ok = false; }
      sel.removeAllRanges(); document.body.removeChild(div);
      return ok;
    }
    try {
      if (navigator.clipboard && window.ClipboardItem) {
        var item = new ClipboardItem({
          "text/html": new Blob([html], { type: "text/html" }),
          "text/plain": new Blob([text], { type: "text/plain" })
        });
        return navigator.clipboard.write([item]).then(function () { return true; }, function () { return legacy(); });
      }
    } catch (e) { /* fall through */ }
    return Promise.resolve(legacy());
  }

  // ---------------- small helpers ----------------
  function refreshButtons() {
    var ok = !!state.rules && !state.rules.errors.length;
    $("importBtn").disabled = state.busy || !ok || !state.inExcel || !state.canImport || !chosenSlots().length;
    var n = chosenSlots().length;
    $("importBtn").textContent = n > 1 ? "Replace " + n + " sheets" : n === 1 ? "Replace sheet" : "Replace sheets";
    $("previewBtn").disabled = state.busy || !ok || !state.inExcel;
    var ab = $("applyBtn"); if (ab) ab.disabled = state.busy;
    var hasRowSpec = ok && state.rules.newRow.cells.length > 0;
    $("parsePaste").disabled = state.busy || !hasRowSpec || !state.inExcel || !$("pasteBox").value.trim();
    $("mailBuild").disabled = state.busy || !ok || !state.inExcel;
    var wb = $("writeRowBtn"); if (wb) wb.disabled = state.busy;
  }

  function setStatus(t) { $("statusText").textContent = t; $("status").classList.remove("hidden"); }
  function clearStatus() { $("status").classList.add("hidden"); }
  function setLamp(id, t, cls, title) { var p = $(id); p.textContent = t; p.className = "lamp " + (cls || ""); if (title) p.title = title; }
  function setPill(t, cls) {
    setLamp("rulesPill", t, cls, cls === "ok" ? "rules.txt loaded with no errors" : cls === "bad" ? "The rules file has errors - see Rules file at the bottom" : "Loading rules.txt");
  }
  function showBanner(t, bad) { var b = $("banner"); b.textContent = t; b.className = "banner" + (bad ? " bad" : ""); }
  function showError(e) { $("results").innerHTML = '<div class="note bad">' + esc(errText(e)) + "</div>"; }

  function errText(e) {
    if (!e) return "Unknown error";
    var msg = e.message || String(e);
    if (e.code === "AccessDenied" || /protect/i.test(msg)) msg += " (the workbook or sheet may be protected)";
    if (e.code === "UnsupportedFeature") msg += " (the file contains something that can't be copied, e.g. a PivotTable or chart in Excel on the web)";
    return msg;
  }

  function safeValue(v) {
    if (v === null || v === undefined) return "";
    if (typeof v === "string" && v.charAt(0) === "=") return "'" + v; // never turn text into a formula
    return v;
  }

  function isDateFormat(f) {
    if (typeof f !== "string" || /^general$/i.test(f) || f === "@") return false;
    var core = f.replace(/"[^"]*"/g, "").replace(/\[[^\]]*\]/g, "");
    return /y/i.test(core) && /d|m/i.test(core);
  }

  function serialToDate(serial) { return new Date(Date.UTC(1899, 11, 30) + Math.round(serial * 86400000)); }

  function show(v, fmt) {
    if (v === null || v === undefined) return "";
    if (typeof v === "number" && isDateFormat(fmt)) {
      var d = serialToDate(v);
      var s = pad(d.getUTCDate()) + "/" + pad(d.getUTCMonth() + 1) + "/" + d.getUTCFullYear();
      if (v % 1 && /h/i.test(fmt)) s += " " + pad(d.getUTCHours()) + ":" + pad(d.getUTCMinutes());
      return s;
    }
    return String(v);
  }

  function pad(n) { return (n < 10 ? "0" : "") + n; }
  function hhmm(d) { return pad(d.getHours()) + ":" + pad(d.getMinutes()); }
  function dateStr(d) { return pad(d.getDate()) + "/" + pad(d.getMonth() + 1) + "/" + d.getFullYear(); }
  function kb(n) { return n > 1048576 ? (n / 1048576).toFixed(1) + " MB" : Math.max(1, Math.round(n / 1024)) + " KB"; }
  function rnd() { return Math.random().toString(36).slice(2, 8); }
  function esc(s) {
    return String(s == null ? "" : s).replace(/[&<>"']/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c];
    });
  }

  function toBase64(buf) {
    var bytes = new Uint8Array(buf), s = "", CH = 0x8000;
    for (var i = 0; i < bytes.length; i += CH) s += String.fromCharCode.apply(null, bytes.subarray(i, i + CH));
    return btoa(s);
  }

  // UTF-8 first, Hebrew Windows encoding as fallback
  function decodeText(buf) {
    try { return new TextDecoder("utf-8", { fatal: true }).decode(buf).replace(/^﻿/, ""); }
    catch (e) { return new TextDecoder("windows-1255").decode(buf); }
  }

  function parseCsv(text) {
    var first = text.split(/\r?\n/)[0] || "";
    var delim = [",", ";", "\t"].map(function (d) { return [d, first.split(d).length]; }).sort(function (a, b) { return b[1] - a[1]; })[0][0];
    var rows = [], row = [], field = "", q = false;
    for (var i = 0; i < text.length; i++) {
      var c = text[i];
      if (q) {
        if (c === '"' && text[i + 1] === '"') { field += '"'; i++; }
        else if (c === '"') q = false;
        else field += c;
      } else if (c === '"') q = true;
      else if (c === delim) { row.push(field); field = ""; }
      else if (c === "\n" || c === "\r") {
        if (c === "\r" && text[i + 1] === "\n") i++;
        row.push(field); rows.push(row); row = []; field = "";
      } else field += c;
    }
    if (field !== "" || row.length) { row.push(field); rows.push(row); }
    return rows;
  }
})();
