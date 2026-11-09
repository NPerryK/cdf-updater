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
    mail: null,             // last built email {subject, text, html}
    page: "today",
    privText: "",           // settings kept in the workbook
    history: [], snap: null, prevSnap: null,
    precheck: {}, precheckFor: {},
    morning: { steps: {} },
    filter: {}, lastRes: null, lookupCache: null, undo: null, abTarget: null,
    myNotes: null, myNotesAt: null, myNotesAuto: false, statusPack: null, stations: null, stuck: null
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
    applyTheme();
    wire();
    watchActionBar();
    restoreLocalRules();
    loadStores().then(function () { return loadRules(); }).then(function () { return fetchPrompt(); });
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
    $("repaintBtn").addEventListener("click", function () { repaintNow({}, "repaintMsg", "Sheet look ✓"); });
    $("lookCopy").addEventListener("click", copyLookRequest);
    $("openChatGpt").addEventListener("click", function () { openUrl("https://chatgpt.com/"); });
    $("openClaudeLook").addEventListener("click", function () { openUrl("https://claude.ai/new"); });
    $("lookPaste").addEventListener("paste", onLookPaste);
    $("lookPaste").addEventListener("input", function () { clearTimeout(state.lookTimer); state.lookTimer = setTimeout(readLookAnswer, 400); });
    $("rulesPill").addEventListener("click", function () { if (state.rules && state.rules.errors.length) showPage("setup"); });   // the red light leads to its reason
    // email
    $("mailKind").addEventListener("change", function () { $("mailSubject").value = defaultSubject(); $("mailPreview").innerHTML = ""; state.mail = null; renderKindHint(); });
    $("mailBuild").addEventListener("click", buildMail);
    $("saveGroup").addEventListener("click", function () { $("saveGroupBox").classList.toggle("hidden"); $("groupName").focus(); });
    $("saveGroupYes").addEventListener("click", saveLocalGroup);
    ["mailTo", "mailCc"].forEach(function (id) {
      $(id).addEventListener("input", function () { if (state.mail) renderMail(); });
    });
    $("mailSubject").addEventListener("input", function () {
      if (state.mail) { state.mail.subject = $("mailSubject").value; renderMail(); }
    });
    // v1.4
    var ecpKeys = Array.prototype.slice.call(document.querySelectorAll(".ecp .key"));
    ecpKeys.forEach(function (k, i) {
      k.addEventListener("click", function () { showPage(k.getAttribute("data-page")); });
      k.addEventListener("keydown", function (ev) {   // arrow keys move between pages, like a tab list
        var j = ev.key === "ArrowRight" ? i + 1 : ev.key === "ArrowLeft" ? i - 1 : ev.key === "Home" ? 0 : ev.key === "End" ? ecpKeys.length - 1 : null;
        if (j == null) return;
        ev.preventDefault();
        var t = ecpKeys[(j + ecpKeys.length) % ecpKeys.length];
        showPage(t.getAttribute("data-page")); t.focus();
      });
    });
    $("morningBtn").addEventListener("click", morningRun);
    document.addEventListener("paste", onPanePaste);
    $("glance").addEventListener("click", function (ev) {
      var b = ev.target.closest && ev.target.closest("[data-list]");
      if (b) openList(b.getAttribute("data-list"));
    });
    $("listClose").addEventListener("click", closeList);
    $("list").addEventListener("click", function (ev) { if (ev.target === $("list")) closeList(); });
    $("lookupInput").addEventListener("paste", onLookupPaste);
    $("linkCopy").addEventListener("click", function () { copyStatusForMyNotes($("linkCopy")); });
    $("linkPaste").addEventListener("paste", function (ev) {
      var t = ev.clipboardData && ev.clipboardData.getData("text/plain");
      if (t) { ev.preventDefault(); $("linkPaste").value = t; takeMyNotes(t); }
    });
    $("linkPaste").addEventListener("change", function () { if ($("linkPaste").value.trim()) takeMyNotes($("linkPaste").value); });
    $("undoBtn").addEventListener("click", undoCheckUi);
    $("privSave").addEventListener("click", savePrivate);
    $("privMove").addEventListener("click", movePrivate);
    $("lookupOpen").addEventListener("click", openLookup);
    $("lookupClose").addEventListener("click", closeLookup);
    $("lookupInput").addEventListener("keydown", function (ev) { if (ev.key === "Enter") doLookup(); if (ev.key === "Escape") closeLookup(); });
    $("lookup").addEventListener("click", function (ev) { if (ev.target === $("lookup")) closeLookup(); });
    $("bigPreviewClose").addEventListener("click", function () { $("bigPreview").classList.add("hidden"); });
    $("bigPreview").addEventListener("click", function (ev) { if (ev.target === $("bigPreview")) $("bigPreview").classList.add("hidden"); });
    document.addEventListener("keydown", function (ev) {
      if ((ev.ctrlKey || ev.metaKey) && (ev.key === "f" || ev.key === "F")) { ev.preventDefault(); openLookup(); }
      if (ev.key === "Escape") { closeLookup(); closeList(); $("bigPreview").classList.add("hidden"); }
    });
    Array.prototype.forEach.call(document.querySelectorAll("[data-theme-choice]"), function (b) {
      b.addEventListener("click", function () {
        try { localStorage.setItem(THEME_KEY, b.getAttribute("data-theme-choice")); } catch (e) { /* ignore */ }
        applyTheme();
      });
    });
  }

  function renderExtras() {
    renderMailKinds(); renderGroups(); renderProject(); renderMorningFiles();
    var n = state.rules ? state.rules.privateLines || 0 : 0;
    $("privTag").textContent = n ? n + " line" + (n > 1 ? "s" : "") + " in use" : "";
  }

  // checklist state of a step: done + a short note next to its title
  function markStep(n, text) {
    var el = $("step" + n) || (n === 3 ? $("newRowCard") : n === 4 ? $("mailCard") : null);
    if (!el) return;
    el.classList.add("done");
    $("state" + n).textContent = text + " " + hhmm(new Date());
    if (n === 3) keyDot("newrows", "done");
    if (n === 4) keyDot("email", "done");
  }

  function renderLamps() {
    var r = state.rules;
    if (!r) { setLamp("algoLamp", "ALGO", "", "No rules loaded"); return; }
    if (r.cdfAlgorithm) {
      var want = r.cdfAlgorithm.version, have = window.CdfAlgo.ALGORITHM_VERSION;
      if (want && want !== have) setLamp("algoLamp", "ALGO CHECK", "warn", "rules.txt expects algorithm " + want + ", the code implements " + have);
      else setLamp("algoLamp", "ALGO " + have, "ok", "CDF algorithm " + have + ", stages 1–" + (window.CdfAlgo.STAGES_DONE || window.CdfAlgo.STAGE) + " of 4");
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
    state.rulesPublic = E.parseRules(text);
    state.rules = F.mergeRules(state.rulesPublic, state.privText || "");
    // outside Excel there's no workbook to hold the people settings: checking the hosting shouldn't show their absence as an error
    if (!state.inExcel) state.rules.errors = state.rules.errors.filter(function (e) { return !e.setting; });
    else {
      var se = state.rules.errors.filter(function (e) { return e.setting; })[0];
      if (se) showBanner("One step before the update: " + se.message, true);   // said where the buttons are, not only in SETUP
    }
    state.loadedAt = new Date();
    state.lastPreview = null;
    $("results").innerHTML = "";
    var errs = state.rules.errors.length, onlySettings = errs && state.rules.errors.every(function (e) { return e.setting; });
    setPill(onlySettings ? "SETTINGS MISSING" : errs ? "RULES " + errs + " ERR" : "RULES OK", errs ? "bad" : "ok");
    if (errs) $("rulesPill").title = onlySettings ? "Three people lines are missing from the workbook settings — click to open SETUP" : "Click to see the errors in SETUP";
    if (!errs && state.rules.privateLines) $("rulesPill").title = "rules.txt + " + state.rules.privateLines + " lines of workbook settings, no errors";
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
          return '<div class="err">' + (e.setting ? "Workbook settings" : "Line " + e.line) + ": " + esc(e.message) + (e.text ? '<br><code>' + esc(e.text) + "</code>" : "") + "</div>";
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
        (f ? precheckHtml(s.slot) : "") +
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
    renderSlots(); renderMorningFiles(); refreshButtons();
    precheckNew();
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
      var pcBad = [];
      slots.forEach(function (s) { (state.precheck[s.slot] || []).forEach(function (x) { if (x.level === "bad") pcBad.push(s.target + ": " + x.text); }); });
      $("importConfirmText").innerHTML = html + mismatchWarnings(slots) +
        (pcBad.length ? '<br><span class="warn-text">⚠ ' + pcBad.map(esc).join("<br>⚠ ") + "</span>" : "") + "<br>Continue?";
      $("importConfirm").classList.remove("hidden");
      $("importBtn").disabled = true;
    }).catch(showError);
  }

  function doImport() {
    $("importConfirm").classList.add("hidden");
    state.busy = true; refreshButtons();
    importSlots(chosenSlots()).then(function () { state.busy = false; clearStatus(); refreshButtons(); });
  }

  // replace the sheets of these slots, one after the other; resolves with the slots that failed
  function importSlots(slots) {
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
    return chain.then(function () {
      var ok = slots.filter(function (s) { return state.slotStatus[s.slot] && state.slotStatus[s.slot].kind === "done"; });
      if (ok.length) markStep(1, ok.length + (ok.length > 1 ? " sheets" : " sheet") + " replaced");
      state.lookupCache = null;
      slots.forEach(function (s) { delete state.precheck[s.slot]; delete state.precheckFor[s.slot]; });
      renderSlots(); renderMorningFiles();
      return slots.filter(function (s) { return !state.slotStatus[s.slot] || state.slotStatus[s.slot].kind !== "done"; });
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
      if (state.rules.cdfAlgorithm) return runCdf(false).then(function (res) {
        renderCdfResults(res, false);
        var n = res.cdfChanges.length + res.gccChanges.length + (res.rowOps || []).length;
        return showSnapshot(res, !n && !res.errors.length);
      });
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
      ? runCdf(true).then(function (res) { state.lookupCache = null; renderCdfResults(res, true); return showSnapshot(res, true); })
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
    return runCdfIn(apply).then(function (res) {
      if (apply) state.trimConfirmed = false;                     // a yes to the first trim is for one Apply only
      var sig = res && res.history && res.history.keepSignature;
      if (!sig || !storesSupported()) return res;
      // EMAIL DRAFT was cleared for NOT CURRENT: the signature waits in the workbook settings for the next good run
      state.signature = sig;
      return writeStore(NS_SETTINGS, { privateRules: state.privText || "", signature: sig, savedAt: new Date().toISOString() }).catch(function () { /* the draft can be signed by hand */ }).then(function () { return res; });
    });
  }
  function runCdfIn(apply) {
    var rules = state.rules, now = new Date();
    return Excel.run(function (ctx) {
      var data = {};
      return readSheets(ctx, ["CDF", "GCC", "CDF SUMMARY", "Mantra", "MANTRAREV", "EMAIL DRAFT"], { backup: true, formats: ["CDF"] }).then(function (rd) {
        var sh = rd.sheets;
        if (!sh.CDF) throw new Error("Sheet \"CDF\" was not found in this workbook");
        if (!sh.GCC) throw new Error("Sheet \"GCC\" was not found — import CDF_List_Export first");
        data.cdf = sh.CDF; data.gcc = sh.GCC; data.summary = sh["CDF SUMMARY"]; data.mantra = sh.Mantra;
        data.mantraRev = sh.Mantra ? sh.MANTRAREV : null; data.email = sh["EMAIL DRAFT"];
        data.backup = window.CdfHistory ? rd.backup : null;
        var res = window.CdfAlgo.run({ cdf: data.cdf.values, gcc: data.gcc.values, now: now }, rules.cdfConfig);
        res.compare = data.summary ? window.CdfAlgo.compareWithSheet(res.summary, data.summary.values) : null;
        mantraStage(res, data, rules.cdfConfig);
        historyStage(res, data, rules.cdfConfig, now);
        reportsStage(res, data, rules.cdfConfig, now);
        res.lookPlan = window.CdfLook ? window.CdfLook.lookPlan(finalCdfOf(res, data), window.CdfAlgo) : null;
        if (!apply) return res;
        if (res.errors.some(function (e) { return e.stage === "00.1"; })) throw new Error("Headers don't match the algorithm — nothing was written.");
        var cdfWs = ctx.workbook.worksheets.getItem("CDF"), gccWs = ctx.workbook.worksheets.getItem("GCC");
        setStatus("Writing " + (res.cdfChanges.length + res.gccChanges.length) + " cells…");
        // check the row operations; then the log (so a write that stops halfway can still be undone); then the
        // row operations (bottom-up) and every CDF cell in the row numbers they leave
        return checkRowOps(ctx, cdfWs, res.rowOps).then(function () {
          if (!rules.log) return;
          var stamp = stampOf(now), name = "CDF algorithm " + window.CdfAlgo.ALGORITHM_VERSION + " · stage " + (res.history && res.history.ran ? 4 : res.stage2 && res.stage2.ran ? 2 : 1);
          var rows = res.gccChanges.concat(res.cdfChanges).map(function (c) {
            return [stamp, name, c.sheet, String(c.row), c.key, c.colLetter, show(c.oldValue, c.numberFormat || null), show(c.newValue, c.numberFormat || null), c.rule === "20.2" ? "" : "Rule " + c.rule + " — " + c.why];
          });
          if (res.rowOps.length) rows = rows.concat(window.CdfMantra.rowOpLogRows(res.rowOps, stamp, name, function (v) { return show(v, null); }));
          res.info.onlyGcc.forEach(function (k) { rows.push([stamp, name, "GCC", "", k, "", "", "", E.LOG_NOTES.notFound]); });
          rows = rows.concat(reportLogRows(res.reports, stamp, name)).concat(historyLogRows(res.history, stamp, name));
          if (!rows.length) rows.push([stamp, name, "", "", "", "", "", "", E.LOG_NOTES.none]);
          setStatus("Writing log…");
          return appendLog(ctx, rules.log, rows);
        }).then(function () {
          return doRowOps(ctx, cdfWs, res.rowOps);
        }).then(function () {
          setStatus("Writing " + (res.cdfChanges.length + res.gccChanges.length) + " cells…");
          return writeChanges(ctx, gccWs, res.gccChanges);
        }).then(function () {
          return writeChanges(ctx, cdfWs, res.cdfChanges);
        }).then(function () {
          // 110.6–110.7 on the sheet as the update left it; a look that can't be painted is reported, not thrown
          setStatus("Painting sheet CDF…");
          return lookPass(ctx, cdfWs, finalCdfOf(res, data), { fresh: freshRows(res.rowOps) })
            .then(function (rep) { res.look = rep; state.lookValues = finalCdfOf(res, data); }, function (e) { res.look = { error: errText(e) }; });
        }).then(function () {
          return writeReports(ctx, res.reports);
        }).then(function () {
          // the update itself is written by now: a BACKUPDATA that can't be written (protected, full) is reported, not thrown
          return writeHistory(ctx, res.history, res.reports).catch(function (e) { res.history.writeError = errText(e); });
        }).then(function () { return res; });
      });
    });
  }

  // ---------------- stage 4 — BACKUPDATA, the release tables, sheet ERROR ----------------
  // BACKUPDATA grows by thousands of rows a run, so it is read by column in chunks: the RunID (and snapshot type)
  // of each area, then only the RUNS rows, the EVENTS rows and the baseline run's FINAL rows
  // ---------------- reading the workbook in a few round trips
  // Every Excel round trip (ctx.sync) costs time, more than the work inside it. The sheets an update needs are read
  // together: which exist (1 trip), their used ranges and BACKUPDATA's headers (1), every block of rows, grouped up to
  // READ_BUDGET cells a trip, and BACKUPDATA's runs and baseline (2). The values are exactly what a sheet-by-sheet read gives.
  var READ_BUDGET = 150000, WRITE_BUDGET = 60000, BK_CHUNK = 20000;   // cells a trip: a few MB, well inside what Excel takes
  // jobs: [{ cells, queue(), done() }] → queued and synced in groups of at most `budget` cells, in order
  function roundTrips(ctx, jobs, budget) {
    var groups = [], cur = [], n = 0;
    jobs.forEach(function (j) { if (cur.length && n + j.cells > budget) { groups.push(cur); cur = []; n = 0; } cur.push(j); n += j.cells; });
    if (cur.length) groups.push(cur);
    return groups.reduce(function (p, g) {
      return p.then(function () {
        g.forEach(function (j) { j.queue(); });
        return ctx.sync().then(function () { g.forEach(function (j) { if (j.done) j.done(); }); });
      });
    }, Promise.resolve());
  }
  function newBackup() {
    return { exists: false, last: { runs: 1, cdfs: 1, solutions: 1, events: 1, mantradocs: 1 }, runs: [], events: [], eventRows: [], rowIds: { cdfs: [], solutions: [], mantradocs: [] }, snap: {}, base: null, baseline: null, headerProblems: [] };
  }
  // names → { sheets: { name: { values, formats } | null }, backup: (opt.backup) what readBackup gives }
  // opt.formats: the sheets whose number formats are read (the others come with formats: [] — nothing reads them)
  function readSheets(ctx, names, opt) {
    opt = opt || {};
    var fmt = {}; (opt.formats || []).forEach(function (n) { fmt[n] = true; });
    var H = window.CdfHistory, wss = ctx.workbook.worksheets, sh = {}, used = {}, out = { sheets: {}, backup: null }, withBk = !!(opt.backup && H);
    var all = names.concat(withBk ? ["BACKUPDATA"] : []), bk = withBk ? newBackup() : null, head = null, bkWs = null, idReads = [];
    all.forEach(function (n) { sh[n] = wss.getItemOrNullObject(n); sh[n].load("name"); });
    return ctx.sync().then(function () {
      all.forEach(function (n) {
        if (sh[n].isNullObject) return;
        used[n] = sh[n].getUsedRangeOrNullObject(true);
        used[n].load(n === "BACKUPDATA" ? "rowCount,rowIndex" : "rowIndex,rowCount,columnIndex,columnCount");
      });
      if (withBk && !sh.BACKUPDATA.isNullObject) { bk.exists = true; bkWs = sh.BACKUPDATA; head = bkWs.getRange("A1:CG1"); head.load("values"); }
      return ctx.sync();
    }).then(function () {
      var jobs = [];
      names.forEach(function (n) {
        if (sh[n].isNullObject) { out.sheets[n] = null; return; }
        var u = used[n];
        if (u.isNullObject) { out.sheets[n] = { values: [], formats: [] }; return; }
        var rows = u.rowIndex + u.rowCount, cols = u.columnIndex + u.columnCount, d = out.sheets[n] = { values: [], formats: [] }, k = fmt[n] ? 2 : 1;
        var CH = Math.max(1, Math.min(rows, Math.floor(READ_BUDGET / k / Math.max(1, cols))));
        for (var r = 0; r < rows; r += CH) (function (start, count) {
          var rg;
          jobs.push({ cells: count * cols * k,
            queue: function () { rg = sh[n].getRangeByIndexes(start, 0, count, cols); rg.load(k === 2 ? "values,numberFormat" : "values"); },
            done: function () { Array.prototype.push.apply(d.values, rg.values); if (k === 2) Array.prototype.push.apply(d.formats, rg.numberFormat); } });
        })(r, Math.min(CH, rows - r));
      });
      if (bk && bk.exists) {
        bk.headerProblems = H.checkHeaders(head.values[0]);
        if (!bk.headerProblems.length) {
          var u = used.BACKUPDATA, end = u.isNullObject ? 1 : u.rowIndex + u.rowCount, cols = { runs: ["A", "A"], cdfs: ["S", "T"], solutions: ["AG", "AH"], events: ["BD", "BD"], mantradocs: ["BX", "BY"] };
          bk.end = end;
          Object.keys(cols).forEach(function (a) {
            for (var r = 2; r <= end; r += BK_CHUNK) (function (a, r) {
              var x = { a: a, r: r, rg: null }, to = Math.min(end, r + BK_CHUNK - 1);
              idReads.push(x);
              jobs.push({ cells: (to - r + 1) * (cols[a][0] === cols[a][1] ? 1 : 2), queue: function () { x.rg = bkWs.getRange(cols[a][0] + r + ":" + cols[a][1] + to); x.rg.load("values"); } });
            })(a, r);
          });
        }
      }
      return roundTrips(ctx, jobs, READ_BUDGET);
    }).then(function () {
      if (!bk) return;
      out.backup = bk;
      if (!bk.exists || bk.headerProblems.length) return;
      return backupFinish(ctx, bkWs, bk, idReads);
    }).then(function () { return out; });
  }
  // the RunID columns → each area's last row and its runs' spans; then RUNS, EVENTS and the baseline's FINAL rows (2 trips)
  function backupFinish(ctx, ws, out, reads) {
    var H = window.CdfHistory, spans = {};
    out.rowIds = { cdfs: [], solutions: [], mantradocs: [] };     // the RunID of every sheet row from row 2: what a trim removes
    reads.forEach(function (x) {
      x.rg.values.forEach(function (v, i) {
        var id = String(v[0] == null ? "" : v[0]).trim(), row = x.r + i;
        if (out.rowIds[x.a]) out.rowIds[x.a][row - 2] = id;
        if (!id) return;
        var k = x.a + "|" + id + "|" + String(v[v.length - 1] == null ? "" : v[v.length - 1]).trim();
        if (row > out.last[x.a]) out.last[x.a] = row;
        if (!spans[k]) spans[k] = [row, row]; else spans[k][1] = row;
      });
    });
    Object.keys(out.rowIds).forEach(function (a) { out.rowIds[a].length = Math.max(0, out.last[a] - 1); for (var i = 0; i < out.rowIds[a].length; i++) if (out.rowIds[a][i] == null) out.rowIds[a][i] = ""; });
    // a run can be the baseline only while its FINAL snapshot is here (older runs keep just their RUNS row)
    out.snap = {};
    Object.keys(spans).forEach(function (k) { var p = k.split("|"); if (p[0] === "cdfs" && p[2] === "FINAL" && spans["solutions|" + p[1] + "|FINAL"]) out.snap[p[1]] = true; });
    var rr = out.last.runs > 1 ? ws.getRange("A2:P" + out.last.runs) : null, er = out.last.events > 1 ? ws.getRange("BD2:BU" + out.last.events) : null;
    if (rr) rr.load("values"); if (er) er.load("values");
    return ctx.sync().then(function () {
      out.runs = rr ? rr.values.filter(function (r) { return String(r[0] || "").trim(); }) : [];
      out.eventRows = er ? er.values : [];
      out.events = out.eventRows.filter(function (r) { return String(r[0] || "").trim(); });
      out.baseline = H.baselineRun(out.runs, function (id) { return !!out.snap[id]; });
      if (!out.baseline) return out;
      var id = out.baseline.runId, part = {}, areas = { cdfs: ["S", "AD"], solutions: ["AG", "BA"], mantradocs: ["BX", "CG"] };
      Object.keys(areas).forEach(function (a) {
        var sp = spans[a + "|" + id + "|FINAL"];
        if (!sp) return;
        part[a] = ws.getRange(areas[a][0] + sp[0] + ":" + areas[a][1] + sp[1]); part[a].load("values");
      });
      return ctx.sync().then(function () {
        out.base = {};
        Object.keys(areas).forEach(function (a) {
          out.base[a] = part[a] ? part[a].values.filter(function (r) { return String(r[0]).trim() === id && String(r[1]).trim() === "FINAL"; }).map(function (r) { return r.map(function (v) { return v == null ? "" : v; }); }) : [];
        });
        return out;
      });
    });
  }
  function readBackup(ctx) { return readSheets(ctx, [], { backup: true }).then(function (r) { return r.backup; }); }
  // the CDF sheet after the whole update, in memory (stage 3 and stage 4 both read it)
  function finalCdfOf(res, data) {
    if (!res.finalCdf) res.finalCdf = window.CdfMantra ? window.CdfMantra.finalValues(data.cdf.values, res.rowOps || [], res.cdfChanges) : data.cdf.values;
    return res.finalCdf;
  }
  function historyStage(res, data, cfg, now) {
    res.history = null;
    var H = window.CdfHistory, bk = data.backup;
    if (!H || !res.summary) return;
    var s2 = res.stage2 || {};
    if (!data.mantra || !(s2.ran || s2.bad)) { res.history = { skipped: "BACKUPDATA and the release tables need sheets Mantra and MANTRAREV (stage 2) — not written this time.", missing: true }; return; }
    if (bk && bk.headerProblems.length) { res.history = { skipped: "Sheet BACKUPDATA doesn't have the BACKUPDATA-v3 headers (" + bk.headerProblems.slice(0, 2).join("; ") + ") — nothing is appended until it's fixed.", bad: true }; return; }
    var wants = {};
    (s2.blocks || []).forEach(function (b) { wants[b.key] = b.want; });
    var out = H.buildRun({
      now: now, started: now, algo: window.CdfAlgo, M: window.CdfMantra, runs: bk ? bk.runs : [], base: bk ? bk.base : null, history: bk ? bk.events : [],
      hasSnapshot: function (id) { return !!(bk && bk.snap[id]); },
      startCdf: data.cdf.values, finalCdf: finalCdfOf(res, data), mantra: data.mantra.values, mantraRev: data.mantraRev ? data.mantraRev.values : [], gcc: data.gcc.values,
      cat: s2.catalogs ? s2.catalogs.cat : null, rev: s2.catalogs ? s2.catalogs.rev : null,
      matchOf: function (key, i, AA) { var w = (wants[key] || [])[i]; return w && String(w.AA).trim() === AA ? w.match || null : null; },
      errors: { stage1Errors: res.errors, checks: res.checks, stage2Errors: s2.errors || [] }
    });
    // what BACKUPDATA keeps after this run: the last two committed runs and 8 days of events; the first (bulk) trim waits for a yes
    var trim = bk && bk.exists ? H.retention({ runs: bk.runs, snap: bk.snap, rows: bk.rowIds, eventRows: bk.eventRows, now: now.getTime(), current: { runId: out.runId, commits: out.commits } }) : null;
    if (trim) trim.allowed = !trim.bulk || !!state.trimConfirmed;
    var types = {}, committed = H.committedRuns(bk ? bk.runs : []), atMs = now.getTime(), ticks = [];
    out.events.forEach(function (e) { types[e[2]] = (types[e[2]] || 0) + 1; });
    // the last 7 days on one rail: every release event the release tables can show (committed runs, and this one when it commits)
    (bk ? bk.events : []).filter(function (e) { return committed[String(e[0]).trim()]; }).map(function (e) { return { e: e, mine: false }; })
      .concat(out.commits ? out.events.map(function (e) { return { e: e, mine: true }; }) : []).forEach(function (x) {
        var type = String(x.e[2]).trim(), t = H.parseIso(x.e[13]);
        if (t == null || t > atMs || t <= atMs - 168 * 3600000 || !/^(CDF_RELEASED|DOCUMENT_RELEASED|GCC_ANSWER_RECEIVED)$/.test(type)) return;
        ticks.push({ t: t, type: type, mine: x.mine, label: (type === "DOCUMENT_RELEASED" ? String(x.e[6]).trim() + " " + String(x.e[7]).trim() : String(x.e[4]).trim()) });
      });
    res.history = { ran: true, out: out, types: types, ticks: ticks, exists: !!(bk && bk.exists), last: bk ? bk.last : null, trim: trim };
  }
  function historyWork(h) { return !!(h && h.ran && h.out.events.length); }
  function historyText(h) { return historyWork(h) ? plural(h.out.events.length, "new event") + " to record in BACKUPDATA" : ""; }
  function historyLogRows(h, stamp, name) {
    if (!h || !h.ran) return [];
    h.logStamp = stamp; h.logName = name;
    var o = h.out, c = o.counts, row = (h.last ? h.last.runs : 1) + 1, tr = h.trim, cut = "";
    if (o.commits && tr && !tr.skipped) {
      var gone = tr.rows.cdfs + tr.rows.solutions + tr.rows.mantradocs + tr.rows.events;
      cut = tr.allowed ? (gone ? "; removed " + gone.toLocaleString() + " rows no run needs (" + (tr.drop.length ? (tr.drop.length === 1 ? "1 run's snapshot" : tr.drop.length + " runs' snapshots") : "") + (tr.drop.length && tr.rows.events ? ", " : "") + (tr.rows.events ? tr.rows.events + " old events" : "") + ") — not undoable" : "")
        : "; older history (" + tr.drop.length + " runs) kept: the trim wasn't confirmed";
    }
    var rows = [[stamp, name, "BACKUPDATA", String(row), o.runId, "RUN", "", o.commits ? "COMMITTED" : "FAILED",
      o.commits ? "Stage 4: +" + c.cdfs + " CDFS, +" + c.solutions + " SOLUTIONS, +" + c.events + " EVENTS, +" + c.mantradocs + " MANTRADOCS rows" + cut + "; ERROR COUNT " + o.errors.length
        : "Stage 4: FAILED — only the RUNS row is kept; ERROR COUNT " + o.errors.length],
      [stamp, name, "ERROR", "", "", "", "", "", "Rebuilt (stage 4): " + (o.errors.length ? o.errors.length + " error" + (o.errors.length > 1 ? "s" : "") : "no errors")]];
    if (!o.commits) rows.push([stamp, name, "EMAIL DRAFT", "", "", "", "", "", "NOT CURRENT – RUN FAILED " + o.runId + " (150.8.10)"]);
    return rows;
  }
  // append the run to BACKUPDATA (each area under its own last row), rebuild sheet ERROR
  function writeHistory(ctx, h, rep) {
    if (!h || !h.ran) return Promise.resolve();
    var H = window.CdfHistory, o = h.out;
    // a CDF SUMMARY that reads back other numbers is a failed check of this run too (100.12)
    var late = Promise.resolve();
    if (rep && rep.readBack && rep.readBack.length) {
      var wasCommitted = o.commits;
      o.errors = H.errorRows(H.collectErrors({ readBack: rep.readBack }), o.runId, o.at).concat(o.errors);
      o.commits = false; o.run[4] = "FAILED"; o.run[14] = String(o.errors.length); o.run[15] = "FAILED";
      // the log said COMMITTED before writing; say what it is now (and undo leaves the run alone)
      if (wasCommitted && state.rules && state.rules.log) late = appendLog(ctx, state.rules.log, [[h.logStamp || stampOf(new Date()), h.logName || "CDF algorithm", "BACKUPDATA", String((h.last ? h.last.runs : 1) + 1), o.runId, "RUN", "COMMITTED", "FAILED",
        "CDF SUMMARY read back other numbers (100.12) — the run is FAILED: only its RUNS row is kept, EMAIL DRAFT says NOT CURRENT; ERROR COUNT " + o.errors.length]]);
    }
    var ws = ctx.workbook.worksheets.getItemOrNullObject("BACKUPDATA"), tr = h.trim;
    ws.load("name");
    setStatus("Writing BACKUPDATA…");
    return late.then(function () { return ctx.sync(); }).then(function () {
      if (ws.isNullObject) {
        ws = ctx.workbook.worksheets.add("BACKUPDATA");
        var hr = H.headerRow();
        ws.getRange("A1:" + H.letter(hr.length - 1) + "1").values = [hr];
      }
      var last = h.last || { runs: 1, cdfs: 1, solutions: 1, events: 1, mantradocs: 1 }, jobs = [];
      // a committed run: its snapshot and events, under what's there; a failed one is nobody's baseline, so only its RUNS row
      if (o.commits) jobs = appendJobs(ws, "cdfs", o.cdfs, last.cdfs + 1).concat(appendJobs(ws, "solutions", o.solutions, last.solutions + 1),
        appendJobs(ws, "events", o.events, last.events + 1), appendJobs(ws, "mantradocs", o.mantradocs, last.mantradocs + 1));
      jobs = jobs.concat(appendJobs(ws, "runs", [o.run], last.runs + 1));   // the run row last: everything it counts is in place
      return roundTrips(ctx, jobs, WRITE_BUDGET);
    }).then(function () {
      // then the trim: what the next run doesn't need goes (the new rows are already in place below)
      if (!o.commits || !tr || tr.skipped || !tr.allowed) return;
      setStatus("Trimming BACKUPDATA…");
      return trimAreas(ctx, ws, tr).then(function () { tr.done = true; });
    }).then(function () {
      setStatus("Writing ERROR…");
      return writeErrorSheet(ctx, o);
    }).then(function () {
      if (!o.commits) return writeEmailNotCurrent(ctx, o.runId, h);
    }).then(function () {
      return finishWorkbook(ctx);
    }).then(function () { h.written = true; });
  }
  // what this Excel can't do (an older ExcelApi): kept for the pane instead of failing the update
  function noteUnsupported(what) {
    state.unsupported = state.unsupported || {};
    state.unsupported[what] = true;
    try { console.warn("CDF Updater: not supported here — " + what); } catch (e) { /* no console */ }
  }
  // the rows of each area that go, bottom-up, cells below moving up in the area's own columns only (the other areas stay);
  // all in one round trip — Excel runs them in the order given
  function trimAreas(ctx, ws, tr) {
    var H = window.CdfHistory;
    ["cdfs", "solutions", "mantradocs", "events"].forEach(function (a) {
      var c1 = H.col(H.AREAS[a]), c2 = c1 + H.HEADERS[a].length - 1;
      (tr.segments[a] || []).forEach(function (g) { ws.getRange(H.letter(c1) + g[0] + ":" + H.letter(c2) + g[1]).delete("Up"); });
    });
    return ctx.sync();
  }
  // 150.8.10: a failed run leaves EMAIL DRAFT with one line; the signature is kept in the workbook settings for the next good run
  function writeEmailNotCurrent(ctx, runId, h) {
    var ws = ctx.workbook.worksheets.getItemOrNullObject("EMAIL DRAFT");
    ws.load("name");
    return ctx.sync().then(function () {
      if (ws.isNullObject) ws = ctx.workbook.worksheets.add("EMAIL DRAFT");
      var used = ws.getUsedRangeOrNullObject();
      used.load("rowCount,columnCount,rowIndex,columnIndex,values");
      return ctx.sync().then(function () {
        if (!used.isNullObject) {
          var lines = window.CdfSummary ? window.CdfSummary.readSignature(used.values) : [];
          if (lines.length) h.keepSignature = lines;
          var all = ws.getRangeByIndexes(0, 0, used.rowIndex + used.rowCount, Math.max(used.columnIndex + used.columnCount, 1));
          all.unmerge(); all.clear("All");
        }
        var a1 = ws.getRange("A1");
        a1.numberFormat = [["@"]];
        a1.values = [["NOT CURRENT – RUN FAILED " + runId]];
        a1.format.font.bold = true; a1.format.font.color = "#B10001";
        h.emailNotCurrent = true;
        return ctx.sync();
      });
    });
  }
  // 100.0 / 150.8.2: CDF SUMMARY, ERROR and EMAIL DRAFT visible, EMAIL DRAFT right after ERROR, and CDF SUMMARY active at A1
  function finishWorkbook(ctx) {
    var wss = ctx.workbook.worksheets, sum = wss.getItemOrNullObject("CDF SUMMARY"), err = wss.getItemOrNullObject("ERROR"), mail = wss.getItemOrNullObject("EMAIL DRAFT");
    [sum, err, mail].forEach(function (w) { w.load("name,position,visibility"); });
    return ctx.sync().then(function () {
      [sum, err, mail].forEach(function (w) { if (!w.isNullObject && w.visibility !== "Visible") try { w.visibility = "Visible"; } catch (e) { noteUnsupported("sheet visibility"); } });
      if (!err.isNullObject && !mail.isNullObject && mail.position !== err.position + 1) {
        try { mail.position = mail.position > err.position ? err.position + 1 : err.position; } catch (e) { noteUnsupported("sheet order"); }
      }
      if (!sum.isNullObject) { sum.activate(); sum.getRange("A1").select(); }
      return ctx.sync();
    }).catch(function (e) { noteUnsupported("finishing view (" + errText(e) + ")"); });
  }
  // one area's new rows as write jobs of up to 2,000 rows (roundTrips groups them, WRITE_BUDGET cells a trip)
  function appendJobs(ws, area, rows, at) {
    var H = window.CdfHistory, c1 = H.col(H.AREAS[area]), w = H.HEADERS[area].length, CH = 2000, jobs = [];
    for (var i = 0; i < rows.length; i += CH) (function (i) {
      var part = rows.slice(i, i + CH);
      jobs.push({ cells: part.length * w * 2, queue: function () {
        var rg = ws.getRange(H.letter(c1) + (at + i) + ":" + H.letter(c1 + w - 1) + (at + i + part.length - 1));
        rg.numberFormat = part.map(function (r) { return r.map(function () { return "@"; }); });   // ids and UTC times stay text (06.9)
        rg.values = part.map(function (r) { return r.map(function (v) { return v == null ? "" : String(v); }); });
      } });
    })(i);
    return jobs;
  }
  // 95: ERROR after CDF SUMMARY; only this run's failures; RUN ID, AS OF and ERROR COUNT under them
  function writeErrorSheet(ctx, o) {
    var H = window.CdfHistory, S = window.CdfSummary, ws = ctx.workbook.worksheets.getItemOrNullObject("ERROR"), sum = ctx.workbook.worksheets.getItemOrNullObject("CDF SUMMARY");
    ws.load("name"); sum.load("position");
    return ctx.sync().then(function () {
      if (ws.isNullObject) { ws = ctx.workbook.worksheets.add("ERROR"); if (!sum.isNullObject) try { ws.position = sum.position + 1; } catch (e) { /* ExcelApi 1.7 */ } }
      var used = ws.getUsedRangeOrNullObject(true);
      used.load("rowCount,rowIndex");
      return ctx.sync().then(function () {
        var lay = H.errorLayout(o.errors, o.runId, (S ? S.jlm(new Date(H.parseIso(o.at))).date + " " + S.jlm(new Date(H.parseIso(o.at))).time : o.at) + " (" + H.TZ + ")");
        var end = Math.max(used.isNullObject ? 0 : used.rowIndex + used.rowCount, lay.end, 2), all = ws.getRange("A1:O" + end);
        all.unmerge(); all.clear("All");
        var grid = [];
        for (var r = 1; r <= lay.end; r++) { var row = []; for (var c = 0; c < 15; c++) row.push(""); grid.push(row); }
        lay.cells.forEach(function (x) { grid[x.r - 1][x.c] = x.v; });
        var rect = ws.getRange("A1:O" + lay.end);
        rect.numberFormat = grid.map(function (r, i) { return r.map(function (v, j) { return i === lay.info.r1 + 1 && j === 1 ? "#,##0" : "@"; }); });   // ERROR COUNT is a number
        rect.values = grid;
        ws.getRange("A1:O1").merge(false);
        var head = ws.getRange("A1:O2").format;
        head.fill.color = "#B10001"; head.font.color = "#FFFFFF"; head.font.bold = true; head.horizontalAlignment = "Center";
        var fmt = function (rg) { var f = rg.format; f.font.name = "Calibri"; f.font.size = 10; f.verticalAlignment = "Center"; f.wrapText = true;
          ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideVertical", "InsideHorizontal"].forEach(function (sd) { var b = f.borders.getItem(sd); b.style = "Continuous"; b.weight = "Thin"; b.color = "#BCBCBC"; }); };
        fmt(ws.getRange("A1:O" + (lay.records ? lay.records.r2 : 2)));
        if (lay.records) { var body = ws.getRange("A3:O" + lay.records.r2).format; body.fill.color = "#FFFFFF"; body.font.color = "#000000"; body.horizontalAlignment = "Left"; }
        var info = ws.getRange("A" + lay.info.r1 + ":B" + lay.info.r2), inf = info.format;
        inf.fill.color = "#FFFFFF"; inf.font.color = "#000000"; inf.font.name = "Calibri"; inf.font.size = 10; inf.horizontalAlignment = "Left";
        [["A", 21], ["B", 28.6], ["C", 17.1], ["F", 17.1], ["G", 17.1], ["I", 17.1], ["L", 17.1]].forEach(function (x) { ws.getRange(x[0] + ":" + x[0]).format.columnWidth = excelWidth(x[1]); });
        try { ws.showGridlines = false; } catch (e) { /* ExcelApi 1.8 */ }
        return ctx.sync();
      });
    });
  }

  // ---------------- stage 3 — CDF SUMMARY and EMAIL DRAFT ----------------
  // built from the CDF sheet as the whole update leaves it (in memory), compared with the two sheets now
  function reportsStage(res, data, cfg, now) {
    res.reports = null;
    if (!window.CdfSummary) { res.reports = { skipped: "The reports part of the add-in (cdf-summary.js) didn't load — reload the add-in." }; return; }
    if (!res.summary) return;
    // the release tables come from BACKUPDATA's history; a workbook without one yet keeps the tables it has (there's no history to compute them from)
    var h = res.history;
    res.reports = buildReports(finalCdfOf(res, data), data, cfg, now, h && h.ran ? (h.exists ? h.out.release : null) : undefined);
    if (h && h.ran && !h.out.commits && res.reports.ran) res.reports.emailFailed = h.out.runId;
  }
  // the release tables from BACKUPDATA events of committed runs, with the CDF sheet as it is (no run of its own)
  function releaseNow(cdfValues, data, now) {
    var H = window.CdfHistory, M = window.CdfMantra, bk = data.backup;
    if (!H || !M || !bk || !bk.exists || bk.headerProblems.length || !data.mantra) return null;
    var committed = H.committedRuns(bk.runs), cat = M.readMantra(data.mantra.values), sols = H.solutionSnapshot(cdfValues, window.CdfAlgo, { cat: cat, M: M }, "-", "FINAL", "");
    return H.releaseTables(bk.events.filter(function (e) { return committed[String(e[0]).trim()]; }), sols, cat, M, now.getTime());
  }
  function buildReports(cdfValues, data, cfg, now, release) {
    var S = window.CdfSummary, A = window.CdfAlgo, M = window.CdfMantra;
    var rf = A.run({ cdf: cdfValues, gcc: data.gcc.values, now: now }, cfg);
    if (!rf.summary) return { skipped: "Stage 1 stopped, so CDF SUMMARY and EMAIL DRAFT weren't built." };
    var current = data.summary ? data.summary.values : [], emailNow = data.email ? data.email.values : [];
    if (release === undefined) release = releaseNow(cdfValues, data, now);
    var model = S.build({ cdf: cdfValues, summary: rf.summary, cfg: cfg, algo: A, mantra: M, current: current, release: release || null });
    var lay = S.layoutSummary(model, { aliases: S.supplierAliases(A.settings(cfg)) });
    var mail = S.mailParts(now, signatureOf(emailNow));
    var le = S.layoutEmail(model, mail);
    var cmp = S.compare(lay, current), ecmp = data.email ? S.compare(le, emailNow) : [];
    var subjectNow = data.email ? String(((emailNow[0] || [])[0]) == null ? "" : emailNow[0][0]).trim() : "";
    var changedTables = cmp.filter(function (c) { return S.differs(c) || c.moved; });
    return {
      ran: true, model: model, lay: lay, le: le, mail: mail, cmp: cmp, ecmp: ecmp, explain: rf.summary.explain,
      changedTables: changedTables,
      summaryDiffers: !data.summary || changedTables.length > 0,
      emailDiffers: !data.email || subjectNow !== mail.subject || ecmp.some(function (c) { return S.differs(c) || c.moved; }),
      emailSubjectNow: subjectNow, hadSummary: !!data.summary, hadEmail: !!data.email
    };
  }
  function reportsWork(rep) { return !!(rep && rep.ran && (rep.summaryDiffers || rep.emailDiffers)); }
  function reportLogRows(rep, stamp, name) {
    if (!reportsWork(rep)) return [];
    var rows = [];
    if (rep.summaryDiffers) rows.push([stamp, name, "CDF SUMMARY", "", "", "", "", "", "Rebuilt (stage 3): " + rep.lay.tables.length + " tables, rows 1–" + rep.lay.end + (rep.changedTables.length ? "; changed: " + rep.changedTables.map(function (c) { return c.name; }).join(", ") : "")]);
    if (rep.emailDiffers && !rep.emailFailed) rows.push([stamp, name, "EMAIL DRAFT", "", "", "", "", "", "Rebuilt (stage 3): " + rep.mail.subject]);
    return rows;
  }
  function writeReports(ctx, rep) {
    if (!reportsWork(rep)) return Promise.resolve();
    var p = Promise.resolve();
    if (rep.summaryDiffers) p = p.then(function () {
      setStatus("Writing CDF SUMMARY…");
      return writeReportSheet(ctx, "CDF SUMMARY", rep.lay, false).then(function (back) { rep.readBack = back; rep.wroteSummary = true; });
    });
    if (rep.emailDiffers && !rep.emailFailed) p = p.then(function () {       // a failing run: stage 4 writes NOT CURRENT instead (150.8.10)
      setStatus("Writing EMAIL DRAFT…");
      return writeReportSheet(ctx, "EMAIL DRAFT", rep.le, true).then(function () { rep.wroteEmail = true; });
    });
    return p;
  }
  // 100.1: Calibri 10, red title and header rows, white body, a thin #BCBCBC border on every cell; widths 100.1.3
  function excelWidth(chars) { return Math.round((chars * 7 + 5) * 0.75 * 4) / 4; }
  function writeReportSheet(ctx, name, lay, isEmail) {
    var S = window.CdfSummary, ws = ctx.workbook.worksheets.getItemOrNullObject(name);
    ws.load("name");
    return ctx.sync().then(function () {
      if (ws.isNullObject) ws = ctx.workbook.worksheets.add(name);
      var used = ws.getUsedRangeOrNullObject(true);
      used.load("rowCount,rowIndex");
      return ctx.sync().then(function () {
        var oldEnd = used.isNullObject ? 0 : used.rowIndex + used.rowCount, end = Math.max(oldEnd, lay.end, 1);
        var all = ws.getRange("A1:H" + end);
        all.unmerge(); all.clear("All");
        Object.keys(S.WIDTHS).forEach(function (L) { ws.getRange(L + ":" + L).format.columnWidth = excelWidth(S.WIDTHS[L]); });
        lay.tables.forEach(function (p) { writeReportTable(ws, p); });
        lay.merges.forEach(function (m) { ws.getRange(S.letter(m.c1) + m.r1 + ":" + S.letter(m.c2) + m.r2).merge(false); });
        lay.tables.forEach(function (p) { reportBorders(ws, p); });
        lay.cells.filter(function (c) { return c.role === "subject" || c.role === "line"; }).forEach(function (c) {
          var rg = ws.getRange(S.letter(c.c) + c.r);
          rg.numberFormat = [["@"]]; rg.values = [[c.v]];
          rg.format.font.name = "Calibri"; rg.format.font.size = c.role === "subject" ? 12 : 10; rg.format.font.bold = c.role === "subject"; rg.format.font.color = "#000000";
          rg.format.horizontalAlignment = "Left"; rg.format.verticalAlignment = "Center";
        });
        try { ws.showGridlines = false; } catch (e) { /* ExcelApi 1.8 */ }
        try { ws.getRange("A1:H" + lay.end).format.autofitRows(); } catch (e) { /* not everywhere */ }
        return ctx.sync();
      }).then(function () {
        if (isEmail) return null;
        // 100.12: every number the written tables show must be the number computed here
        var rg = ws.getRange("A1:H" + lay.end);
        rg.load("values");
        return ctx.sync().then(function () {
          return S.readBack(lay, function (r, c) { return (rg.values[r - 1] || [])[c]; });
        });
      });
    });
  }
  function writeReportTable(ws, p) {
    var S = window.CdfSummary, w = p.c2 - p.c1 + 1, h = p.r2 - p.r1 + 1, grid = {};
    p.cells.forEach(function (c) { grid[c.r + ":" + c.c] = c; });
    var v = [], nf = [];
    for (var r = p.r1; r <= p.r2; r++) {
      var vr = [], fr = [];
      for (var c = p.c1; c <= p.c2; c++) {
        var x = grid[r + ":" + c];
        vr.push(x ? x.v : "");
        fr.push(x && x.role === "num" ? "#,##0" : x && x.role === "time" ? "@" : "General");
      }
      v.push(vr); nf.push(fr);
    }
    var rect = ws.getRange(S.letter(p.c1) + p.r1 + ":" + S.letter(p.c2) + p.r2);
    rect.numberFormat = nf;                     // text stays text (times like 05/11/2026 06:47 aren't turned into dates)
    rect.formulas = v;
    var f = rect.format;
    f.fill.color = "#FFFFFF"; f.font.name = "Calibri"; f.font.size = 10; f.font.color = "#000000"; f.font.bold = false;
    f.horizontalAlignment = "Left"; f.verticalAlignment = "Center"; f.wrapText = true;
    var head = ws.getRange(S.letter(p.c1) + p.r1 + ":" + S.letter(p.c2) + (p.r1 + 1)).format;
    head.fill.color = "#B10001"; head.font.color = "#FFFFFF"; head.font.bold = true; head.horizontalAlignment = "Center";
    // numbers and MAIN GROUP centred: vertical runs of the same role in one column
    for (var cc = p.c1; cc <= p.c2; cc++) {
      var start = null;
      for (var rr = p.r1 + 2; rr <= p.r2 + 1; rr++) {
        var x2 = grid[rr + ":" + cc], mid = rr <= p.r2 && x2 && (x2.role === "num" || x2.role === "center");
        if (mid && start === null) start = rr;
        if (!mid && start !== null) { ws.getRange(S.letter(cc) + start + ":" + S.letter(cc) + (rr - 1)).format.horizontalAlignment = "Center"; start = null; }
      }
    }
    return { w: w, h: h };
  }
  function reportBorders(ws, p) {
    var S = window.CdfSummary, rg = ws.getRange(S.letter(p.c1) + p.r1 + ":" + S.letter(p.c2) + p.r2);
    var sides = ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"];
    if (p.c2 > p.c1) sides.push("InsideVertical");
    if (p.r2 > p.r1) sides.push("InsideHorizontal");
    sides.forEach(function (sd) { var b = rg.format.borders.getItem(sd); b.style = "Continuous"; b.weight = "Thin"; b.color = "#BCBCBC"; });
  }
  // after an undo: CDF SUMMARY and EMAIL DRAFT again from the CDF sheet as it is now
  function rebuildReports(ctx, now) {
    var data = {}, cfg = state.rules.cdfConfig;
    return readSheets(ctx, ["CDF", "GCC", "CDF SUMMARY", "EMAIL DRAFT", "Mantra"], { backup: true, formats: ["CDF"] }).then(function (rd) {
      var sh = rd.sheets;
      data.cdf = sh.CDF; data.gcc = sh.GCC; data.summary = sh["CDF SUMMARY"]; data.email = sh["EMAIL DRAFT"];
      data.mantra = window.CdfHistory ? sh.Mantra : null;
      data.backup = window.CdfHistory && data.mantra ? rd.backup : null;   // the release tables from the committed runs that are left
      if (!data.cdf || !data.gcc || !window.CdfSummary) return null;
      var rep = buildReports(data.cdf.values, data, cfg, now);
      return writeReports(ctx, rep).then(function () { return rep; });
    });
  }

  // stage 2 — Mantra → CDF (AA, AB, U:X, closing date Q, solution rows). It reads the CDF sheet as stage 1
  // leaves it; the two become one update with every CDF cell in the row numbers after the row operations.
  function mantraStage(res, data, cfg) {
    var M = window.CdfMantra;
    res.rowOps = [];
    res.stage1Cells = res.cdfChanges.length + res.gccChanges.length;
    if (!M) { res.stage2 = { skipped: "The Mantra part of the add-in (cdf-mantra.js) didn't load — reload the add-in." }; return; }
    if (!res.summary) { res.stage2 = { skipped: "Stage 1 stopped, so Mantra wasn't checked." }; return; }
    if (!data.mantra) { res.stage2 = { skipped: "There's no sheet \"Mantra\" — AA, AB, U:X and the closing date weren't checked.", missing: true }; return; }
    if (!data.mantraRev) { res.stage2 = { skipped: "There's no sheet \"MANTRAREV\" — without it, documents released before Mantra look missing. AA, AB, U:X weren't checked.", missing: true }; return; }
    var s2 = M.run({ cdf: M.applyInMemory(data.cdf.values, res.cdfChanges), gcc: data.gcc.values, mantra: data.mantra.values, mantraRev: data.mantraRev.values }, cfg, window.CdfAlgo);
    if (s2.errors.some(function (e) { return e.stage === "00.1.2"; })) {
      res.stage2 = { skipped: "Sheet Mantra can't be read: " + s2.errors.map(function (e) { return e.message; }).join("; "), bad: true, errors: s2.errors };
      return;
    }
    var plan = M.layout(s2, data.cdf.values);
    res.cdfChanges = M.combine(res.cdfChanges, plan);
    // Q is a date: keep the cell's own date format (so the log and undo read it as the sheet shows it)
    res.cdfChanges.forEach(function (c) {
      if (c.colLetter !== "Q" || typeof c.newValue !== "number") return;
      var f = ((data.cdf.formats || [])[(c.sourceRow || c.row) - 1] || [])[c.col];
      c.numberFormat = isDateFormat(f) ? f : "dd/mm/yyyy";
    });
    res.rowOps = plan.ops;
    var kinds = {};
    s2.blocks.forEach(function (b) { kinds[b.kind] = (kinds[b.kind] || 0) + 1; });
    res.stage2 = {
      ran: true, info: s2.info, errors: s2.errors, cells: plan.changes.length, kinds: kinds, blocks: s2.blocks,   // blocks: stage 4 keeps their match
      catalogs: s2.catalogs || null,                                                                           // stage 4 reads Mantra from here
      rowsAdded: plan.rowsAdded, rowsRemoved: plan.rowsRemoved,
      noPeople: !(cfg.personGroups && cfg.personGroups.length)
    };
  }

  // Office.js for the row operations of stage 2. Everything is checked before anything moves: the row numbers
  // of the whole update depend on every operation happening, so one failed check stops the write.
  function checkRowOps(ctx, ws, ops) {
    if (!ops || !ops.length) return Promise.resolve();
    var M = window.CdfMantra, A = window.CdfAlgo, req = Office.context.requirements;
    var canAreas = req.isSetSupported("ExcelApi", "1.13"), canCopy = req.isSetSupported("ExcelApi", "1.9");
    setStatus("Checking " + ops.length + " block" + (ops.length > 1 ? "s" : "") + " that change rows…");
    // 110.4.1: rows can't be merged across rows a filter hides — the filter goes first
    var filt = canCopy && ws.autoFilter ? ws.autoFilter : null;
    if (filt) filt.load("isDataFiltered");
    var probes = ops.map(function (o) {
      var p = { o: o };
      if (o.kind === "remove") { p.f = ws.getRange("A" + o.first + ":AH" + o.last); p.f.load("formulas"); return p; }
      if (o.own) { p.f = ws.getRange("A" + o.own + ":AH" + o.own); p.f.load("formulas"); }
      if (canAreas) { p.m = ws.getRange("A" + o.anchor + ":AH" + o.at).getMergedAreasOrNullObject(); p.m.load("address"); }
      return p;
    });
    return ctx.sync().then(function () {
      if (filt && filt.isDataFiltered) throw new Error("Sheet CDF is filtered. Solution rows are added or removed in this update, and that can't be done safely while rows are hidden — clear the filter (Data → Clear) and apply again. Nothing was written.");
      probes.forEach(function (p) {
        var o = p.o;
        if (o.kind === "remove") {
          p.f.formulas.forEach(function (row, i) {
            if (M.contAnchor(row[A.col("R")]) !== o.anchor || String(row[A.col("C")] || "").trim())
              throw new Error(o.key + ": row " + (o.first + i) + " isn't one of its solution rows, so it can't be removed — nothing was written. Fix the block by hand and check again.");
          });
          return;
        }
        var areas = p.m && !p.m.isNullObject ? parseAreas(p.m.address) : null;
        if (areas && areas.some(function (a) { return a.r1 <= o.last && a.r2 >= o.at; }))
          throw new Error(o.key + ": merged cells reach below its last solution row (row " + o.last + "), so rows can't be added safely — nothing was written. Fix the block by hand and check again.");
        // which columns join the anchor: the block's own merges and formulas, else the usual block
        var own = p.f ? p.f.formulas[0] : null;
        o.formulas = {};
        if (own) own.forEach(function (f, c) { if (M.contAnchor(f) === o.anchor) o.formulas[A.letter(c)] = f; });
        else M.REF_COLS.forEach(function (L) { o.formulas[L] = M.contFormula(L, o.anchor); });
        o.mergeCols = areas && o.have > 1
          ? areas.filter(function (a) { return a.r1 === o.anchor && a.r2 === o.last && a.c1 === a.c2; }).map(function (a) { return A.letter(a.c1); })
          : M.MERGE_COLS.filter(function (L) { return !o.formulas[L]; });
        o.canCopy = canCopy;
      });
    });
  }
  function doRowOps(ctx, ws, ops) {
    if (!ops || !ops.length) return Promise.resolve();
    setStatus("Adding and removing solution rows…");
    ops.forEach(function (o) {
        if (o.kind === "remove") { ws.getRange(o.first + ":" + o.last).delete("Up"); return; }
        ws.getRange(o.at + ":" + (o.at + o.count - 1)).insert("Down");
        addContinuation(ws, { anchor: o.anchor, at: o.at, end: o.last + o.count, mergeCols: o.mergeCols, formulas: o.formulas, template: o.canCopy ? o.template : null });
    });
    return ctx.sync();
  }
  // the inserted rows (at..end) join the block at `anchor`: merges, the =IF($R$a…) references, the continuation look
  function addContinuation(ws, b) {
    var n = b.end - b.at + 1;
    b.mergeCols.forEach(function (L) { var rg = ws.getRange(L + b.anchor + ":" + L + b.end); rg.unmerge(); rg.merge(false); });
    Object.keys(b.formulas).forEach(function (L) {
      var f = []; for (var i = 0; i < n; i++) f.push([b.formulas[L]]);
      ws.getRange(L + b.at + ":" + L + b.end).formulas = f;
    });
    if (b.template) [["Q", "X"], ["AA", "AH"]].forEach(function (s) {
      ws.getRange(s[0] + b.at + ":" + s[1] + b.end).copyFrom(s[0] + b.template + ":" + s[1] + b.template, "Formats");
    });
  }
  // "CDF!A7:A8,CDF!B7:B8" → [{ c1, r1, c2, r2 }] (0-based columns, 1-based rows)
  function parseAreas(address) {
    return String(address || "").split(",").map(function (a) {
      var m = /([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(a.replace(/\$/g, "").trim());
      if (!m) return null;
      return { c1: window.CdfAlgo.col(m[1]), r1: +m[2], c2: window.CdfAlgo.col(m[3] || m[1]), r2: +(m[4] || m[2]) };
    }).filter(Boolean);
  }

  // ================= the look of sheet CDF (110.6–110.7) and the looks from ChatGPT / Claude =================
  var LK = window.CdfLook, NS_LOOKS = "urn:cdf-updater:looks:v1";
  var BORDER_SIDES = ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideVertical", "InsideHorizontal"];
  // values: sheet CDF as it is now (in memory after an update, or read). opt.fresh: rows this update added (they
  // fit their text); opt.drop: formulas of looks just removed (their rules go too). 4 round trips + row heights.
  function lookPass(ctx, ws, values, opt) {
    opt = opt || {};
    var plan = LK.lookPlan(values, window.CdfAlgo);
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
        var x = { cf: cf, rule: cf.custom.rule, ranges: canAreas ? cf.getRangesOrNullObject() : null };
        x.rule.load("formula"); if (x.ranges) x.ranges.load("areaCount");
        custom.push(x);
      });
      return ctx.sync();
    }).then(function () {
      var found = custom.map(function (x) { return { formula: x.rule.formula, areas: x.ranges && !x.ranges.isNullObject ? x.ranges.areaCount : 1 }; });
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
  function freshRows(ops) {
    var out = [];
    (ops || []).forEach(function (o) { if (o.kind === "insert") for (var i = 0; i < o.count; i++) out.push(o.finalFirst + i); });
    return out;
  }
  function lookReportHtml(rep, title, brief) {
    if (!rep) return "";
    if (brief && !rep.error && !rep.skipped) return '<div class="note ok look-report"><b>' + esc(brief) + "</b> " + esc(title) + "</div>";
    if (rep.error) return '<div class="note warn look-report">' + esc(title || "Sheet look") + ": couldn't be repainted — " + esc(rep.error) + ". The update itself is written; try <b>Repaint CDF</b> on the LOOK page.</div>";
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
  function lookPlanHtml(res, applied) {
    if (applied) return lookReportHtml(res.look, "Sheet look ✓");
    var p = res.lookPlan;
    if (!p) return "";
    var n = (state.looks || []).length;
    return '<div class="note look-report">At Apply, the look of rows 2–' + p.last + " (110.6–110.7): Released green, Cancelled red, a line over " + plural(p.cont.length, "solution row") + (n ? ", " + plural(n, "saved look") : "") + ".</div>";
  }

  // ---- LOOK page
  function renderLookPage() {
    var list = state.looks || [];
    $("looksCount").textContent = list.length ? String(list.length) : "";
    var box = $("savedLooks");
    if (!list.length) { box.innerHTML = '<p class="hint">None yet. A look you add is kept in this workbook and painted again after every Apply.</p>'; return; }
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
      var res = LK.parseLooks(text, lookHeaders(), E);
      if (res.error) { state.lookDraft = null; box.innerHTML = '<div class="note bad">' + esc(res.error) + (res.kind === "invalid" || res.kind === "nojson" ? ' Ask in the same chat: "Answer again with only the json block, in the format I gave."' : "") + "</div>"; return; }
      state.lookDraft = res.looks;
      return lookValues().then(function (vals) {
        var today = LK.todaySerial(new Date()), A = window.CdfAlgo;
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
    var A = window.CdfAlgo, of = {}, seen = {}, list = [];
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
        return (n == null ? "" : plural(n, "row") + " painted now. ") + "Kept in this workbook and painted again after every Apply.";
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

  function renderCdfResults(res, applied) {
    state.lastRes = res; state.lastResAt = Date.now(); state.lastApplied = applied; state.filter = {};
    var all = res.gccChanges.concat(res.cdfChanges), nChanges = all.length, nCdf = countKeys(all.concat(res.rowOps || []));
    var failed = res.checks.filter(function (c) { return !c.ok; }).length, nErr = res.errors.length;
    var s2 = res.stage2 || {}, ops = res.rowOps || [], nOps = ops.length, rep = res.reports, repWork = reportsWork(rep), hisWork = historyWork(res.history);
    var work = nChanges + nOps + (repWork ? 1 : 0) + (hisWork ? 1 : 0), cellWork = nChanges + nOps;
    var rowsTxt = rowOpsText(s2), repTxt = reportsText(rep, applied), v;
    if (!res.summary) v = ["bad", "!", "The algorithm stopped", nErr + " error" + (nErr > 1 ? "s" : "") + " — nothing can be written"];
    else if (nErr || failed) v = ["bad", "!", failed ? failed + " control total" + (failed > 1 ? "s" : "") + " failed" : nErr + " error" + (nErr > 1 ? "s" : "") + " found",
      nChanges ? nChanges + " cells could change. Check the details before applying." : "Nothing to update — but check the details below."];
    else if (applied && work) v = ["ok", "✓", "Update written", cellWork ? nChanges + " cell" + (nChanges === 1 ? "" : "s") + (rowsTxt ? ", " + rowsTxt : "") + " in " + nCdf + " CDF" + (nCdf > 1 ? "s" : "") + (repTxt ? "; " + repTxt : "") + ", all control totals OK" : capFirst(repTxt) + ". All control totals OK."];
    else if (!work) v = ["ok", "✓", "Nothing to update", "The CDF sheet already matches GCC" + (s2.ran ? " and Mantra" : "") + (rep && rep.ran ? ", and CDF SUMMARY and EMAIL DRAFT are up to date" : "") + ". All control totals OK."];
    else if (!cellWork && !repWork) v = ["go", "Δ", "New events to record", "The CDF sheet already matches GCC and Mantra. " + capFirst(historyText(res.history)) + "."];
    else if (!cellWork) v = ["go", "Δ", "CDF SUMMARY and EMAIL DRAFT to rebuild", "The CDF sheet already matches GCC" + (s2.ran ? " and Mantra" : "") + ". " + capFirst(repTxt) + (hisWork ? "; " + historyText(res.history) : "") + "."];
    else v = ["go", "Δ", nChanges + " cell" + (nChanges === 1 ? "" : "s") + " to update in " + nCdf + " CDF" + (nCdf > 1 ? "s" : ""), (rowsTxt ? "Also " + rowOpsText(s2, true) + ". " : "") + "All control totals OK — review the changes and apply"];
    var html = '<div class="res-block"><div class="verdict ' + v[0] + '"><span class="vi">' + v[1] + '</span><div><div class="vt">' + esc(v[2]) + '</div><div class="vs">' + esc(v[3]) + "</div></div></div>";
    if (nErr) html += '<div class="note bad"><b>' + nErr + " error" + (nErr > 1 ? "s" : "") + "</b> — these cells are not changed:<ul>" +
      res.errors.map(function (e) { return "<li>" + esc((e.key ? e.key + ": " : "") + e.message) + " <span class=\"rule\">(" + esc(e.stage) + ")</span></li>"; }).join("") + "</ul></div>";
    if (!res.summary) { $("results").innerHTML = html + "</div>"; return; }

    html += '<div class="tiles">' +
      tile(res.info.gccRows, "CDFs in GCC") +
      tile(res.info.matched, "Matched in CDF") +
      tile(res.info.onlyGcc.length, "Only in GCC", res.info.onlyGcc.length ? "warnt" : "") +
      tile(nCdf, applied ? "CDFs updated" : "CDFs to change", "hl") +
      tile(nChanges, applied ? "Cells updated" : "Cells to change", "hl") +
      tile(res.warnings.length, "Warnings", res.warnings.length ? "warnt" : "") + "</div>";

    html += mantraHtml(res, applied);
    html += lookPlanHtml(res, applied);
    html += historyHtml(res, applied);
    html += reportsHtml(res, applied);
    if (applied && work) html += '<div class="note ok">Update written ✓' + (state.rules.log ? " — every change and its previous value is in sheet <b>" + esc(state.rules.log) + "</b>." : "") +
      (state.rules.log ? '<div><button id="undoNow" class="btn small" type="button"><svg class="ico"><use href="#i-undo"/></svg>Undo this update</button></div>' : "") + "</div>";
    if (applied && work) markStep(2, (cellWork ? nChanges + " cells written" + (rowsTxt ? ", " + rowsTxt : "") : "") + (repTxt ? (cellWork ? "; " : "") + repTxt : ""));
    else if (!work && !nErr) markStep(2, "Up to date");
    if (res.info.onlyGcc.length) html += '<div class="note warn">In GCC but not in the CDF sheet (add them as new rows):' + chips(res.info.onlyGcc) + "</div>";
    if (res.info.onlyCdf.length) html += '<div class="note warn">In the CDF sheet but not in GCC (left as they are):' + chips(res.info.onlyCdf) + "</div>";
    if (res.warnings.length) html += '<div class="note warn"><ul>' + res.warnings.map(function (w) { return "<li>" + esc(w.key + ": " + w.message) + "</li>"; }).join("") + "</ul></div>";

    if (nChanges) html += '<div id="changesBox"></div>';
    html += crossCheckHtml(res, applied);
    html += "</div>";

    if (!applied && work) {
      var repLine = repWork ? (rep.summaryDiffers && rep.emailDiffers ? "CDF SUMMARY and EMAIL DRAFT" : rep.summaryDiffers ? "CDF SUMMARY" : "EMAIL DRAFT") + " will be rebuilt from it" : "";
      if (res.history && res.history.ran) repLine += (repLine ? ", and " : "") + "the run is appended to BACKUPDATA" + (repLine ? "" : " from it");
      html += '<div class="confirm go"><p>' + (cellWork ? "<b>" + nChanges + "</b> cell" + (nChanges === 1 ? "" : "s") + " will be written to " + (res.gccChanges.length ? "sheets CDF and GCC" : "sheet CDF") + (rowsTxt ? ", and " + rowsTxt + " in sheet CDF" : "") + (repLine ? "; " + repLine : "") + ". Save the workbook first — Ctrl+Z may not undo this" + (state.rules.log ? ", but the log keeps the previous values and SETUP can undo it" : "") + "."
        : esc(capFirst(repLine.replace(" from it", " from the CDF sheet as it is"))) + ". Nothing in sheet CDF changes.") + "</p>" +
        '<div class="row"><button id="applyBtn" class="btn ok">' + (cellWork ? "Apply " + (nChanges + nOps) + " change" + (nChanges + nOps === 1 ? "" : "s") : !repWork ? "Record " + plural(res.history.out.events.length, "event") : "Rebuild " + (rep.summaryDiffers && rep.emailDiffers ? "both sheets" : rep.summaryDiffers ? "CDF SUMMARY" : "EMAIL DRAFT")) + "</button>" +
        '<button id="cancelApply" class="btn ghost">Cancel</button></div></div>';
    }
    $("results").innerHTML = html;
    if (nChanges) renderChanges();
    var ab = $("applyBtn");
    if (ab) ab.addEventListener("click", applyChanges);
    var cb = $("cancelApply");
    if (cb) cb.addEventListener("click", function () { $("results").innerHTML = ""; });
    var un = $("undoNow");
    if (un) un.addEventListener("click", function () { showPage("setup"); undoCheckUi(); });
    var rm = $("repMail");
    if (rm) rm.addEventListener("click", function () { showPage("email"); $("mailKind").value = "daily"; $("mailSubject").value = defaultSubject(); buildMail(); });
    var bk = $("bkDownload");
    if (bk) bk.addEventListener("click", downloadBackup);
    var tk = $("trimOk");
    if (tk) tk.addEventListener("change", function () { state.trimConfirmed = tk.checked; });
    if (!state.morningRunning) $("results").scrollIntoView({ behavior: "smooth", block: "start" });
  }

  function plural(n, w) { return n + " " + w + (n === 1 ? "" : "s"); }
  function capFirst(s) { s = String(s || ""); return s.charAt(0).toUpperCase() + s.slice(1); }
  // "CDF SUMMARY: 6 tables change, EMAIL DRAFT for 06/11/2026" — or after applying, "CDF SUMMARY and EMAIL DRAFT rebuilt"
  function reportsText(rep, applied) {
    if (!reportsWork(rep)) return "";
    if (rep.emailFailed) {                                       // a failing run: EMAIL DRAFT says NOT CURRENT (150.8.10)
      if (applied) return (rep.summaryDiffers ? "CDF SUMMARY rebuilt, " : "") + "EMAIL DRAFT: NOT CURRENT";
      return (rep.summaryDiffers ? "CDF SUMMARY: " + (rep.hadSummary ? plural(rep.changedTables.length, "table") + " change" + (rep.changedTables.length === 1 ? "s" : "") : "new") + ", " : "") + "EMAIL DRAFT: NOT CURRENT";
    }
    var both = rep.summaryDiffers && rep.emailDiffers;
    if (applied) return (both ? "CDF SUMMARY and EMAIL DRAFT" : rep.summaryDiffers ? "CDF SUMMARY" : "EMAIL DRAFT") + " rebuilt";
    var out = [];
    if (rep.summaryDiffers) out.push("CDF SUMMARY: " + (rep.hadSummary ? plural(rep.changedTables.length, "table") + " change" + (rep.changedTables.length === 1 ? "s" : "") : "new"));
    if (rep.emailDiffers) out.push("EMAIL DRAFT for " + rep.mail.date);
    return out.join(", ");
  }
  // stage 4 at a glance: the last 7 days of releases on one rail, this run's events, what BACKUPDATA gets, the ERROR count
  var EV_KINDS = [["CDF_RELEASED", "CDF released", "CDFs released", "cdf"], ["DOCUMENT_RELEASED", "document released", "documents released", "doc"],
    ["GCC_ANSWER_RECEIVED", "GCC answer", "GCC answers", "gcc"], ["REVISION_CHANGE", "revision change", "revision changes", "rev"]];
  function historyHtml(res, applied) {
    var h = res.history, H = window.CdfHistory;
    if (!h) return "";
    if (!h.ran) return h.missing ? "" : '<div class="note ' + (h.bad ? "bad" : "warn") + ' hist-note"><b>History (stage 4)</b> — ' + esc(h.skipped) + "</div>";   // no Mantra: stage 2's note says it
    var o = h.out, c = o.counts, base = o.base, nErr = o.errors.length, atMs = H.parseIso(o.at);
    var sub = base ? "since " + base.runId + " · " + H.jlmShort(H.parseIso(base.at)) : h.exists ? "no committed run with a snapshot — history starts again now" : "BACKUPDATA is new — history starts with this run";
    var html = '<div class="hist' + (nErr ? " failing" : "") + '"><div class="hist-head"><span class="hist-t">History</span><span class="hist-s">' + esc(sub) + "</span></div>";
    // the rail: 7 days, the last 24 hours lit
    var span = 168 * 3600000, rel = { cdf: 0, doc: 0, gcc: 0 };
    var marks = h.ticks.slice().sort(function (a, b) { return a.t - b.t; }).map(function (k) {
      var kind = k.type === "CDF_RELEASED" ? "cdf" : k.type === "DOCUMENT_RELEASED" ? "doc" : "gcc", x = Math.max(0, Math.min(100, 100 - (atMs - k.t) / span * 100));
      rel[kind]++;
      return '<i class="tk ' + kind + (k.mine ? " mine" : "") + '" style="left:' + x.toFixed(2) + '%" title="' + esc(k.label + " · " + H.jlmShort(k.t)) + '"></i>';
    }).join("");
    var bx = base ? 100 - (atMs - H.parseIso(base.at)) / span * 100 : null;
    html += '<div class="rail" role="img" aria-label="' + esc("Last 7 days: " + rel.cdf + " CDFs released, " + rel.doc + " documents released, " + rel.gcc + " GCC answers") + '">' +
      '<div class="rail-track"><span class="rail-24"></span>' + (bx != null && bx > 0 ? '<span class="rail-base" style="left:' + bx.toFixed(2) + '%" title="' + esc("Baseline " + base.runId) + '"></span>' : "") + marks + "</div>" +
      '<div class="rail-scale"><span>−7 d</span><span class="r24">−24 h</span><span>now</span></div>' +
      '<div class="rail-key"><span class="cdf"><i></i>' + rel.cdf + " CDF" + (rel.cdf === 1 ? "" : "s") + '</span><span class="doc"><i></i>' + rel.doc + " document" + (rel.doc === 1 ? "" : "s") + '</span><span class="gcc"><i></i>' + rel.gcc + " GCC answer" + (rel.gcc === 1 ? "" : "s") + "</span></div></div>";
    // this run's events
    var evs = EV_KINDS.filter(function (k) { return h.types[k[0]]; });
    html += '<div class="hist-ev">' + (evs.length ? (applied ? "Recorded: " : "New since the baseline: ") + evs.map(function (k) { var n = h.types[k[0]]; return '<span class="he ' + k[3] + '"><b>' + n + "</b> " + (n === 1 ? k[1] : k[2]) + "</span>"; }).join("")
      : base ? "No new events since the baseline." : "The first run records the sheets as they are — events start with the next run.") + "</div>";
    // what BACKUPDATA gets
    var led = [["RUNS", "run", 1], ["CDFS", "CDFs", c.cdfs], ["SOLUTIONS", "solutions", c.solutions], ["EVENTS", "events", c.events], ["MANTRADOCS", "Mantra", c.mantradocs]];
    html += '<div class="ledger" role="list" aria-label="Rows BACKUPDATA ' + (applied && h.written ? "got" : "gets") + '">' + led.map(function (x, i) {
      if (i && !o.commits) return '<span role="listitem" class="off" title="' + esc("A failed run keeps no " + x[0] + " rows") + '"><b>—</b>' + esc(x[1]) + "</span>";
      return '<span role="listitem" title="' + esc("+" + x[2].toLocaleString() + " row" + (x[2] === 1 ? "" : "s") + " in " + x[0]) + '"><b>+' + x[2].toLocaleString() + "</b>" + esc(x[1]) + "</span>"; }).join("") + "</div>";
    html += trimHtml(h, applied);
    if (applied && h.writeError) html += '<div class="note bad hist-in"><b>BACKUPDATA wasn\'t written</b> — ' + esc(h.writeError) + ". The rest of the update is in place; the next Apply records the run.</div>";
    html += nErr
      ? '<div class="note bad hist-in"><b>' + nErr + " error" + (nErr > 1 ? "s" : "") + "</b> — the run " + (applied ? "is" : "will be") + " written as FAILED: BACKUPDATA keeps only its RUNS row, its events don't count, and EMAIL DRAFT says NOT CURRENT. Sheet ERROR lists " + (nErr > 1 ? "them" : "it") + ":<ul>" +
        o.errors.slice(0, 6).map(function (e) { return "<li>" + esc((e[5] ? e[5] + ": " : "") + e[8] + " — " + e[11]) + ' <span class="rule">(' + esc(e[2]) + ")</span></li>"; }).join("") + (nErr > 6 ? "<li>and " + (nErr - 6) + " more…</li>" : "") + "</ul></div>"
      : '<div class="hint hist-hint ok">' + (applied && h.written ? "Appended to BACKUPDATA as <b class=\"rid\">" + esc(o.runId) + "</b> (COMMITTED) · sheet ERROR: no errors." : "No errors — the run commits as <b class=\"rid\">" + esc(o.runId) + "</b>.") + "</div>";
    if (!h.exists) html += '<div class="hint hist-hint">The release tables stay as they are until BACKUPDATA has history to compute them from.</div>';
    return html + "</div>";
  }
  // what BACKUPDATA lets go: the snapshots no run needs any more and events older than 8 days. The first (bulk) trim asks first.
  function trimHtml(h, applied) {
    var o = h.out, tr = h.trim;
    if (!o.commits) return '<div class="hint hist-hint">A failed run is nobody\'s baseline, so BACKUPDATA keeps only its RUNS row.</div>';
    if (!tr || tr.skipped) return "";
    var gone = tr.rows.cdfs + tr.rows.solutions + tr.rows.mantradocs, ev = tr.rows.events, keep = tr.keep.length ? tr.keep[0] : "";
    var what = (tr.drop.length ? (tr.drop.length === 1 ? "1 older run's snapshot" : tr.drop.length + " older runs' snapshots") + " (" + gone.toLocaleString() + " rows)" : "") + (tr.drop.length && ev ? " and " : "") + (ev ? plural(ev, "event") + " older than 8 days" : "");
    if (!what) return "";
    if (!tr.bulk) return '<div class="hint hist-hint trim-roll">' + (applied && tr.done ? "Removed " : "Lets go of ") + esc(what) + " — BACKUPDATA keeps " + (keep ? esc(keep) + " and this run" : "this run") + ", and 8 days of events.</div>";
    if (applied) return tr.done
      ? '<div class="note ok hist-in trim"><b>Older history removed</b> — ' + esc(what) + ". From now on BACKUPDATA keeps the last two committed runs and 8 days of events.</div>"
      : '<div class="note warn hist-in trim"><b>Older history still there</b> — ' + esc(what) + ". It goes with the next Apply you confirm.</div>";
    return '<div class="trim-box" role="group" aria-labelledby="trimT"><div class="trim-t" id="trimT">Older history</div>' +
      "<p>BACKUPDATA holds " + esc(what) + ". From now on it keeps the last two committed runs and 8 days of events — what the next comparison and the release tables use. Removing the older rows can't be undone.</p>" +
      '<button id="bkDownload" class="btn" type="button"><svg class="ico"><use href="#i-download"/></svg>Download a backup (.zip)</button>' +
      '<label class="trim-ok"><input type="checkbox" id="trimOk"' + (state.trimConfirmed ? " checked" : "") + "> I have a backup — remove the older history with this Apply</label>" +
      '<div class="hint" id="bkMsg">No downloads in this Excel? Save a copy of the workbook instead (File › Save a Copy). Without the tick, this Apply only adds the run.</div></div>';
  }
  // the whole BACKUPDATA, one CSV per area, in a zip — read in chunks so a big sheet doesn't stall Excel
  function downloadBackup() {
    var H = window.CdfHistory, msg = $("bkMsg"), btn = $("bkDownload");
    if (!window.JSZip) { if (msg) msg.textContent = "The zip part of the add-in didn't load — save a copy of the workbook instead."; return; }
    if (btn) btn.disabled = true;
    if (msg) msg.textContent = "Reading BACKUPDATA…";
    var files = {};
    Excel.run(function (ctx) {
      var ws = ctx.workbook.worksheets.getItem("BACKUPDATA"), used = ws.getUsedRangeOrNullObject(true);
      used.load("rowCount,rowIndex");
      return ctx.sync().then(function () {
        var end = used.isNullObject ? 1 : used.rowIndex + used.rowCount, chain = Promise.resolve(), CH = 5000;
        Object.keys(H.AREAS).forEach(function (a) {
          var c1 = H.col(H.AREAS[a]), c2 = c1 + H.HEADERS[a].length - 1, lines = [H.HEADERS[a].map(csvCell).join(",")];
          files[a] = lines;
          for (var r = 2; r <= end; r += CH) (function (r) {
            chain = chain.then(function () {
              var rg = ws.getRange(H.letter(c1) + r + ":" + H.letter(c2) + Math.min(end, r + CH - 1));
              rg.load("values");
              return ctx.sync().then(function () {
                rg.values.forEach(function (row) { if (String(row[0] == null ? "" : row[0]).trim()) lines.push(row.map(csvCell).join(",")); });
                if (msg) msg.textContent = "Reading BACKUPDATA… " + a + " row " + Math.min(end, r + CH - 1).toLocaleString();
              });
            });
          })(r);
        });
        return chain;
      });
    }).then(function () {
      var zip = new JSZip(), stamp = fileStamp(new Date());
      Object.keys(files).forEach(function (a) { zip.file("BACKUPDATA_" + a.toUpperCase() + ".csv", "\ufeff" + files[a].join("\r\n")); });
      return zip.generateAsync({ type: "blob", compression: "DEFLATE" }).then(function (blob) {
        var name = "BACKUPDATA_backup_" + stamp + ".zip";
        saveBlob(name, blob, msg, Object.keys(files).reduce(function (n, a) { return n + files[a].length - 1; }, 0).toLocaleString() + " rows");
      });
    }).catch(function (e) { if (msg) msg.textContent = "The backup couldn't be made (" + errText(e) + ") — save a copy of the workbook instead."; })
      .then(function () { if (btn) btn.disabled = false; });
  }
  function csvCell(v) { v = v == null ? "" : String(v); return /[",\r\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v; }
  function saveBlob(name, blob, msg, what) {
    try {
      var url = URL.createObjectURL(blob), a = document.createElement("a");
      a.href = url; a.download = name; document.body.appendChild(a); a.click(); document.body.removeChild(a);
      setTimeout(function () { URL.revokeObjectURL(url); }, 8000);
      if (msg) msg.textContent = "Saved " + name + " (" + what + ") — look in Downloads. If nothing arrived, this Excel doesn't allow downloads from add-ins: save a copy of the workbook instead.";
      state.lastSaved = name;
    } catch (e) { if (msg) msg.textContent = "This Excel doesn't allow downloads from add-ins — save a copy of the workbook instead (File › Save a Copy)."; }
  }
  // stage 3 at a glance: the tables of CDF SUMMARY (changed or the same) and the email
  function reportsHtml(res, applied) {
    var rep = res.reports;
    if (!rep) return "";
    if (!rep.ran) return '<div class="note warn rep-note"><b>CDF SUMMARY (stage 3)</b> — ' + esc(rep.skipped) + "</div>";
    var S = window.CdfSummary, changed = {}, car = rep.model.tables.filter(function (t) { return t.carried; });
    rep.changedTables.forEach(function (c) { changed[c.id] = c; });
    var html = '<div class="rep"><div class="rep-head"><span class="rep-t">Reports</span><span class="rep-s">' +
      esc(applied ? (rep.wroteSummary || rep.wroteEmail || rep.emailFailed ? reportsText(rep, true) : "up to date") : reportsText(rep, false) || "CDF SUMMARY and EMAIL DRAFT are up to date") + "</span></div>";
    if (rep.emailFailed) html += '<div class="note bad rep-in">The run ' + (applied ? "failed" : "will fail") + ', so EMAIL DRAFT ' + (applied ? "shows" : "will show") + " only “NOT CURRENT – RUN FAILED " + esc(rep.emailFailed) + "” — no status goes out from a failed run (150.8.10). The next good run rebuilds it, signature included.</div>";
    html += '<div class="rep-grid" role="list">' + rep.lay.tables.map(function (p) {
      var c = changed[p.id], t = rep.model.tables.filter(function (x) { return x.id === p.id; })[0];
      var rowsDiff = c && c.found ? c.added.length + c.removed.length + c.changed.length : 0, onlyMoves = c && c.found && !rowsDiff;
      var n = c ? (c.found ? (onlyMoves ? "" : rowsDiff) : "new") : "";
      var icon = !c ? "=" : applied && rep.wroteSummary ? "✓" : onlyMoves ? "↓" : "Δ", cls = !c ? "same" : applied && rep.wroteSummary ? "done" : onlyMoves ? "mov" : "chg";
      var short = p.name.replace(/^PERSONS – GROUP – SIGNATURE – CDF – PENDING MANTRA DOCUMENTS$/, "PERSONS").replace(/^DOCUMENTS RELEASED \/ GCC ANSWERS — /, "DOCS ").replace(/^CDFS RELEASED — /, "CDFS ").replace(/LAST (24 HOURS|7 DAYS)/, function (m, x) { return x === "24 HOURS" ? "24 h" : "7 d"; });
      var tip = p.name + " · " + S.letter(p.c1) + p.r1 + (!c ? " · the same" : !c.found ? " · not in the sheet yet" : onlyMoves ? " · the same rows, " + (applied ? "moved" : "moves") + " to row " + p.r1 + " (was " + c.at + ")" : " · " + c.added.length + " row(s) in, " + c.removed.length + " out, " + c.changed.length + " changed") + (t && t.carried ? " · kept as it is (no BACKUPDATA history to compute it from)" : /^(cdf|doc)(24|7)$/.test(p.id) ? " · from BACKUPDATA events" : "");
      return '<span role="listitem" class="rt ' + cls + (t && t.carried ? " car" : "") + '" title="' + esc(tip) + '">' +
        '<i aria-hidden="true">' + icon + "</i>" + esc(short) + (n !== "" ? "<b>" + esc(String(n)) + "</b>" : "") + "</span>";
    }).join("") + "</div>";
    if (!applied && rep.changedTables.length) html += '<div class="rep-legend"><span><i>Δ</i>rows change</span><span><i>↓</i>only moves</span><span><i>=</i>the same</span><span class="car">kept as it is</span></div>';
    if (car.length) html += '<div class="hint rep-hint">The four release tables (CDFS / DOCUMENTS RELEASED) are kept as they are' + (car[0].title.match(/\(AS OF [^)]*\)/) ? " " + esc(car[0].title.match(/\(AS OF [^)]*\)/)[0]) : "") + " — they are computed from BACKUPDATA when the workbook has sheet Mantra.</div>";
    else if (res.history && res.history.ran) html += '<div class="hint rep-hint">The four release tables are computed from BACKUPDATA — the events of committed runs' + (res.history.out.commits ? ", this run included" : " (this run fails, so its events don't count)") + ".</div>";
    if (rep.model.missing.length) html += '<div class="hint rep-hint">Not in CDF SUMMARY now, so not written: ' + esc(rep.model.missing.join(", ")) + ".</div>";
    if (rep.model.unmapped.length) html += '<div class="note warn rep-in">Not in the people list (PERSON_GROUP), shown as UNMAPPED: ' + esc(rep.model.unmapped.join(", ")) + "</div>";
    if (applied && rep.readBack) html += rep.readBack.length
      ? '<div class="note bad rep-in"><b>CDF SUMMARY shows ' + rep.readBack.length + " number" + (rep.readBack.length > 1 ? "s" : "") + " other than computed</b> — is calculation set to Manual? " + esc(rep.readBack.slice(0, 8).map(function (b) { return b.cell + " " + b.got + " (computed " + b.want + ")"; }).join(", ")) + "</div>"
      : '<div class="hint rep-hint ok">Read back after writing: every number CDF SUMMARY shows is the number computed here.</div>';
    if (rep.emailDiffers || (applied && rep.wroteEmail)) html += '<div class="rep-mail"><span><b>EMAIL DRAFT</b> ' + esc(applied && rep.wroteEmail ? "for " + rep.mail.date + ", AS OF " + rep.mail.time : rep.hadEmail ? "now " + (rep.emailSubjectNow.replace(/^.*— /, "") || "empty") + " → " + rep.mail.date : "new, for " + rep.mail.date) + '</span><button id="repMail" class="btn small" type="button">Open the email</button></div>';

    return html + "</div>";
  }
  // "5 solution rows added, 1 removed" — or, before applying, "5 solution rows to add, 1 to remove"
  function rowOpsText(s2, before) {
    var a = s2 && s2.rowsAdded, r = s2 && s2.rowsRemoved, out = [];
    if (a) out.push(plural(a, "solution row") + (before ? " to add" : " added"));
    if (r) out.push(r + (before ? " to remove" : " removed"));
    return out.join(", ");
  }
  // stage 2 at a glance: what Mantra was read, which blocks change rows, what needs a look
  function mantraHtml(res, applied) {
    var s2 = res.stage2;
    if (!s2) return "";
    if (!s2.ran) return '<div class="note ' + (s2.bad ? "bad" : "warn") + ' mantra-note"><b>Mantra (stage 2)</b> — ' + esc(s2.skipped) + "</div>";
    var ops = res.rowOps || [], k = s2.kinds || {}, cat = s2.info && s2.info.catalog;
    var html = '<div class="mantra"><div class="mantra-head"><span class="mantra-t">Mantra</span><span class="mantra-s">' +
      (cat ? esc(cat.docs.toLocaleString() + " documents · " + cat.mantraRev.toLocaleString() + " in MANTRAREV") : "") + "</span></div>" +
      '<div class="mantra-kinds">' + [["docs", "with documents"], ["answer", "answered in GCC"], ["empty", "no answer yet"], ["cancelled", "cancelled"]].map(function (x) {
        return k[x[0]] ? '<span class="mk"><b>' + k[x[0]] + "</b> " + x[1] + "</span>" : "";
      }).join("") + "</div>";
    // one square per solution row: kept, new (green outline) or removed (struck)
    function squares(o) {
      var h = '', n = Math.min(o.kind === "insert" ? o.need : o.have, 9);
      for (var i = 0; i < n; i++) h += '<i class="' + (o.kind === "insert" ? (i < o.have ? "k" : "n") : (i < o.need ? "k" : "x")) + '"></i>';
      return '<span class="ro-sq" aria-hidden="true">' + h + "</span>";
    }
    if (ops.length) html += '<ul class="rowops">' + ops.slice().reverse().map(function (o) {
      var add = o.kind === "insert", f = o.finalFirst, l = o.finalFirst + o.count - 1;
      var what = add ? "solution " + (o.count > 1 ? "rows " + (o.have + 1) + "–" + o.need : "row " + o.need) + " · " + (applied ? "" : "new ") + "row" + (o.count > 1 ? "s " + f + "–" + l : " " + f)
        : plural(o.count, "extra row") + (applied ? " removed" : " to remove · row" + (o.count > 1 ? "s " + o.first + "–" + o.last : " " + o.first) + " now") + " · " + plural(o.need, "solution") + " left";
      return '<li class="' + (add ? "add" : "del") + '"><span class="ro-n">' + (add ? "+" : "−") + o.count + "</span>" + squares(o) +
        '<span class="ro-t"><b>' + esc(o.key) + "</b><span>" + esc(what) + "</span></span></li>";
    }).join("") + "</ul>";
    if (s2.errors.length) html += '<div class="note warn"><b>' + s2.errors.length + " to check</b> — written as CHECK SOLUTION or left as they are:<ul>" +
      s2.errors.slice(0, 30).map(function (e) { return "<li>" + esc((e.key ? e.key + ": " : "") + e.message) + ' <span class="rule">(' + esc(e.stage) + ")</span></li>"; }).join("") +
      (s2.errors.length > 30 ? "<li>and " + (s2.errors.length - 30) + " more…</li>" : "") + "</ul></div>";
    if (s2.noPeople) html += '<div class="hint mantra-hint">Column V (the group of the person in X) is left as it is: the workbook settings have no PERSON_GROUP lines yet.</div>';
    return html + "</div>";
  }

  // the changes, one card per CDF, with filters by column and rule
  function renderChanges() {
    var res = state.lastRes, box = $("changesBox");
    if (!res || !box) return;
    var all = res.gccChanges.concat(res.cdfChanges), f = state.filter || {};
    var shown = F.filterChanges(all, f), groups = F.groupChangesByCdf(shown), facets = F.changeFacets(all);
    function facet(kind, id, n, label) {
      var on = f[kind] === id;
      return '<button type="button" class="facet' + (on ? " on" : "") + '" data-k="' + kind + '" data-v="' + esc(id) + '" title="' + esc(label) + '">' + esc(kind === "rule" ? "rule " + id : id) + "<b>" + n + "</b></button>";
    }
    var html = '<div class="changes-head"><b>' + (state.lastApplied ? "Written" : "Changes") + " by CDF</b><span class=\"hint\" style=\"margin:0\">" + shown.length + " of " + all.length + " cells</span></div>";
    if (facets.columns.length > 1 || facets.rules.length > 1) {
      html += '<div class="facets">' + facets.columns.map(function (c) { return facet("col", c.id, c.n, "Only column " + c.id); }).join("") +
        '<span class="facet-sep"></span>' + facets.rules.map(function (r) { return facet("rule", r.id, r.n, "Only rule " + r.id); }).join("") + "</div>";
    }
    var limit = 120;
    html += '<div class="cgroups">' + groups.slice(0, limit).map(function (g) {
      return '<div class="cgroup"><div class="cg-head"><b>' + esc(g.key) + "</b><span>" + (g.sheet === "GCC" ? "GCC row " : "row ") + g.row + "</span>" +
        (g.sheet === "GCC" ? '<span class="gcc">GCC SHEET</span>' : "") + "</div>" +
        g.changes.map(function (c) {
          return '<div class="cg-line" title="' + esc("Rule " + c.rule + (c.why ? " — " + c.why : "") + (c.source ? " (" + c.source + ")" : "")) + '"><span class="c">' + c.colLetter +
            '</span><span class="ch"><span class="o">' + esc(show(c.oldValue, c.numberFormat || null) || "(empty)") + '</span><span class="a">→</span><span class="n">' + esc(show(c.newValue, c.numberFormat || null) || "(empty)") +
            '</span></span><span class="r">' + esc(c.rule) + "</span></div>";
        }).join("") + "</div>";
    }).join("") + "</div>" + (groups.length > limit ? '<div class="more">and ' + (groups.length - limit) + " more CDFs…</div>" : "");
    box.innerHTML = html;
    Array.prototype.forEach.call(box.querySelectorAll(".facet"), function (b) {
      b.addEventListener("click", function () {
        var k = b.getAttribute("data-k"), val = b.getAttribute("data-v");
        state.filter[k] = state.filter[k] === val ? null : val;
        renderChanges();
      });
    });
  }

  function countKeys(changes) {
    var k = {}; changes.forEach(function (c) { k[c.key] = true; }); return Object.keys(k).length;
  }

  // the cross-check: control totals (130.6) + every computed table against CDF SUMMARY in this workbook
  function crossCheckHtml(res, applied) {
    var ok = res.checks.filter(function (c) { return c.ok; }).length, rep = res.reports && res.reports.ran ? res.reports : null;
    var cmp = rep ? rep.cmp : null, nT = cmp ? cmp.length : res.compare ? res.compare.length : 0;
    var same = cmp ? cmp.filter(function (c) { return !window.CdfSummary.differs(c) && !c.moved; }).length : res.compare ? res.compare.filter(function (t) { return t.found && !t.diffs.length; }).length : 0;
    var line = "Control totals " + ok + "/" + res.checks.length + " OK" + (nT ? ", CDF SUMMARY " + same + "/" + nT + " the same" : "");
    var openIt = ok < res.checks.length;
    var html = '<details class="xcheck"' + (openIt ? " open" : "") + '><summary><span><span class="xs-t">Cross-check</span><span class="xs-s">' + esc(line) +
      (nT && same < nT ? " — open to see what changes and why" : "") + '</span></span><svg class="chev"><use href="#i-chev"/></svg></summary><div class="xcheck-body">' +
      '<div class="xsub">Control totals (' + ok + "/" + res.checks.length + " OK)</div><ul class=\"xlist\">" +
      res.checks.map(function (c) {
        return '<li class="' + (c.ok ? "ok" : "bad") + '"><span class="mark">' + (c.ok ? "✓" : "✗") + "</span>" + esc(c.name) +
          '<div class="xdetail">' + esc(c.detail) + "</div></li>";
      }).join("") + "</ul>";
    if (cmp) {
      html += '<div class="xsub">CDF SUMMARY now vs. ' + (applied ? "written" : "after the update") + " (" + same + "/" + nT + " the same)</div>" +
        '<p class="hint" style="margin:2px 0 6px"><b>Sheet</b> is what CDF SUMMARY showed before; <b>' + (applied ? "Written" : "After") + "</b> is built from the CDF sheet as the update leaves it. A CDF row in or out of MOR / IGOR says why.</p><ul class=\"xlist\">" +
        cmp.map(function (t) {
          if (!t.found) return '<li class="warn"><span class="mark">+</span>' + esc(t.name) + '<div class="xdetail">Not in CDF SUMMARY yet — ' + (applied ? "written" : "it will be written") + "</div></li>";
          var d = t.added.length + t.removed.length + t.changed.length;
          if (!d) return '<li class="ok"><span class="mark">=</span>' + esc(t.name) + ' <span class="xdetail-inline">(' + esc(t.at) + (t.moved ? ", moves" : "") + ")</span></li>";
          var why = function (row) { return /^pending/.test(t.id) && rep.explain && row[0] !== "TOTAL" ? rep.explain(t.name, row[0]) : ""; };
          var rowTxt = function (r) { return r.filter(function (x) { return x !== ""; }).join(" · "); };
          var lines = t.changed.map(function (c) { return "<tr><td>" + esc(c.label) + "</td><td>" + esc(c.now).replace(/\n/g, "<br>") + "</td><td>" + esc(c.will).replace(/\n/g, "<br>") + "</td></tr>"; })
            .concat(t.removed.slice(0, 25).map(function (r) { var w = why(r); return '<tr class="out"><td>−</td><td>' + esc(rowTxt(r)).replace(/\n/g, "<br>") + "</td><td></td></tr>" + (w ? '<tr class="why"><td colspan="3">' + esc(w) + "</td></tr>" : ""); }))
            .concat(t.added.slice(0, 25).map(function (r) { var w = why(r); return '<tr class="in"><td>+</td><td></td><td>' + esc(rowTxt(r)).replace(/\n/g, "<br>") + "</td></tr>" + (w ? '<tr class="why"><td colspan="3">' + esc(w) + "</td></tr>" : ""); }));
          var more = Math.max(0, t.removed.length - 25) + Math.max(0, t.added.length - 25);
          return '<li class="diff"><span class="mark">Δ</span>' + esc(t.name) + ' <span class="xdetail-inline">(' + esc(t.at) + ")</span>" +
            '<table class="xdiff"><tr><th></th><th>Sheet</th><th>' + (applied ? "Written" : "After") + "</th></tr>" + lines.join("") + (more ? '<tr><td colspan="3">and ' + more + " more row(s)…</td></tr>" : "") + "</table></li>";
        }).join("") + "</ul>";
    } else if (res.compare) {
      html += '<div class="xsub">CDF SUMMARY sheet vs. today (' + same + "/" + res.compare.length + " the same)</div><ul class=\"xlist\">" +
        res.compare.map(function (t) {
          if (!t.found) return '<li class="warn"><span class="mark">?</span>' + esc(t.table) + '<div class="xdetail">Not found in CDF SUMMARY</div></li>';
          if (!t.diffs.length) return '<li class="ok"><span class="mark">=</span>' + esc(t.table) + ' <span class="xdetail-inline">(' + esc(t.at) + ")</span></li>";
          return '<li class="diff"><span class="mark">Δ</span>' + esc(t.table) + ' <span class="xdetail-inline">(' + esc(t.at) + ")</span>" +
            '<table class="xdiff"><tr><th></th><th>Sheet</th><th>Today</th></tr>' +
            t.diffs.map(function (d) {
              return "<tr><td>" + esc(d.label) + "</td><td>" + esc(d.current).replace(/\n/g, "<br>") + "</td><td>" + esc(d.computed).replace(/\n/g, "<br>") + "</td></tr>" +
                (d.why ? '<tr class="why"><td colspan="3">' + esc(d.why) + "</td></tr>" : "");
            }).join("") + "</table></li>";
        }).join("") + "</ul>";
    }
    var shownTables = rep ? rep.model.tables.map(function (t) { return { name: t.title, headers: t.headers, rows: t.rows }; }) : res.summary.tables;
    html += '<details class="xtables"><summary>Show the computed tables</summary>' +
      shownTables.map(function (t) {
        return '<div class="xsub">' + esc(t.name) + '</div><table class="xdiff"><tr>' + t.headers.map(function (h) { return "<th>" + esc(h) + "</th>"; }).join("") + "</tr>" +
          t.rows.map(function (r) { return "<tr>" + r.map(function (v) { return "<td>" + esc(v).replace(/\n/g, "<br>") + "</td>"; }).join("") + "</tr>"; }).join("") + "</table>";
      }).join("") + "</details></div></details>";
    return html;
  }

  function cdfChangesTable(changes) {
    var rows = changes.slice(0, MAX_ROWS_SHOWN).map(function (c) {
      return '<tr title="' + esc("Rule " + c.rule + (c.why ? " — " + c.why : "") + (c.source ? " · " + c.source : "")) + '"><td>' + esc(c.sheet === "GCC" ? "GCC " + c.row : c.row) +
        '</td><td class="v k">' + esc(c.key) + "</td><td>" + c.colLetter +
        '</td><td class="v old">' + esc(show(c.oldValue, c.numberFormat || null) || "(empty)") + '</td><td class="v new">' + esc(show(c.newValue, c.numberFormat || null) || "(empty)") +
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
    if (vs.length === 1 && vs[0] === "") return null;            // a reply from any AI or from Excel, without the add-in's prompt: nothing to compare
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
        }).then(function () {
          if (plan.wsName !== "CDF") return null;
          return repaintIn(ctx, { fresh: done.map(function (d) { return d.row; }) }).catch(function () { return null; });   // the look waits for the next Apply
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
      renderKindHint(); showPage("email");
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
        }).then(function () {
          return ws.name === "CDF" ? repaintIn(ctx, { fresh: [row] }).catch(function () { return null; }) : null;
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

  var DAILY = { id: "daily", label: "Daily CDF status (like EMAIL DRAFT)", subject: "A330-300 P2F — Daily CDF Status" };
  function mailKinds() { return (state.rules && state.rules.cdfAlgorithm ? [DAILY] : []).concat(MAIL_KINDS); }

  function renderMailKinds() {
    var sel = $("mailKind"), keep = sel.value;
    function opt(id, label) { return '<option value="' + esc(id) + '">' + esc(label) + "</option>"; }
    var html = "";
    if (state.rules && state.rules.cdfAlgorithm) html += '<optgroup label="Daily">' + opt(DAILY.id, DAILY.label) + "</optgroup>";
    var focus = focusMailKinds();
    if (focus.length) html += '<optgroup label="For a supplier or a person">' + focus.map(function (k) { return opt(k.id, k.label); }).join("") + "</optgroup>";
    html += '<optgroup label="From the log">' + MAIL_KINDS.map(function (k) { return opt(k.id, k.label); }).join("") + "</optgroup>";
    var reps = state.rules ? state.rules.reports : [];
    if (reps.length) html += '<optgroup label="Reports">' + reps.map(function (r, i) { return opt("report:" + i, "Report: " + r.name); }).join("") + "</optgroup>";
    sel.innerHTML = html;
    if (keep && sel.querySelector('option[value="' + keep + '"]')) sel.value = keep;
    if (!$("mailSubject").value || !keep) $("mailSubject").value = defaultSubject();
    renderKindHint();
  }
  var KIND_HINTS = {
    daily: "The daily status in the format of sheet EMAIL DRAFT: six tables, red headers, Hebrew opening and signature.",
    changes: "Every cell the last update changed, with its previous value.",
    notfound: "CDFs from the last update that are not in the tracker.",
    newrows: "The CDF rows added or updated today."
  };
  function renderKindHint() {
    var v = $("mailKind").value, k = focusKind(v);
    $("mailKindHint").textContent = k ? (k.type === "supplier" ? "English email for " + k.supplier + ": its open CDFs at Conversion Engineer, with the drawing and the problem description."
      : "Hebrew email to " + k.person + ": the CDFs at Final Approval with everything released — ready to sign.") : KIND_HINTS[v] || "";
  }

  function defaultSubject() {
    var v = $("mailKind").value, today = dateStr(new Date());
    if (v.indexOf("report:") === 0) {
      var r = state.rules && state.rules.reports[+v.split(":")[1]];
      return (r ? r.name : "Report") + " – " + today;
    }
    if (v === "daily") return DAILY.subject + " — " + today;
    var fk = focusKind(v);
    if (fk) return fk.type === "supplier" ? "A330-300 P2F — Open CDFs at Conversion Engineer — " + fk.supplier + " — " + today : "A330-300 P2F — CDFs ready for your signature — " + fk.person + " — " + today;
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

  function buildMail(force) {
    if (state.busy && force !== true) return Promise.resolve(false);
    var kind = $("mailKind").value, title = $("mailSubject").value.trim() || defaultSubject();
    var wasBusy = state.busy;
    state.busy = true; refreshButtons();
    $("mailPreview").innerHTML = "";
    setStatus("Preparing email…");
    return Excel.run(function (ctx) {
      var headerRow = state.rules.newRow.headerRow || 1;
      if (kind === "daily") return buildDaily(ctx, title);
      var fk = focusKind(kind);
      if (fk) return buildFocus(ctx, fk, title);
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
      if (doc.empty) { $("mailPreview").innerHTML = '<div class="note warn">Nothing to send for this output.</div>'; state.mail = null; return false; }
      state.mail = doc.daily ? E.buildDailyEmail(doc) : E.buildEmail(doc);
      state.mail.notes = doc.notes || [];
      state.mail.eml = !!doc.eml;
      state.mail.runId = doc.runId || "";
      renderMail();
      markStep(4, "Draft ready");
      return true;
    }).catch(function (e) { $("mailPreview").innerHTML = '<div class="note bad">' + esc(errText(e)) + "</div>"; return false; })
      .then(function (ok) { state.busy = wasBusy && force === true; if (!state.busy) clearStatus(); refreshButtons(); return ok; });
  }

  // ---- daily status email (algorithm 150): the six tables of EMAIL DRAFT, built from the sheets as they are now ----
  function buildDaily(ctx, subject) {
    var data = {};
    return readSheets(ctx, ["CDF", "GCC", "CDF SUMMARY", "EMAIL DRAFT", "Mantra"], { backup: true, formats: ["CDF"] }).then(function (rd) {
      var sh = rd.sheets;
      if (!sh.CDF) throw new Error("Sheet \"CDF\" was not found in this workbook");
      if (!sh.GCC) throw new Error("Sheet \"GCC\" was not found — import CDF_List_Export first");
      data.cdf = sh.CDF; data.gcc = sh.GCC; data.summary = sh["CDF SUMMARY"]; data.email = sh["EMAIL DRAFT"];
      data.mantra = window.CdfHistory ? sh.Mantra : null;
      data.backup = window.CdfHistory && data.mantra ? rd.backup : null;
      if (!window.CdfSummary) throw new Error("The reports part of the add-in (cdf-summary.js) didn't load — reload the add-in.");
      var rep = buildReports(data.cdf.values, data, state.rules.cdfConfig, new Date());
      if (!rep.ran) throw new Error(rep.skipped);
      var doc = window.CdfSummary.emailDoc(rep.model, rep.mail);
      if (subject && subject !== rep.mail.subject) doc.subject = subject;
      var car = rep.model.tables.filter(function (t) { return t.carried; });
      doc.notes = ["Built from the CDF sheet as it is now — the same tables the update writes to EMAIL DRAFT."];
      if (car.length) doc.notes.push("The release tables are copied from CDF SUMMARY " + ((car[0].title.match(/\(AS OF [^)]*\)/) || [""])[0]) + " — the add-in computes them from BACKUPDATA once the workbook has it and sheet Mantra.");
      else if (data.backup && data.backup.exists) doc.notes.push("The release tables are computed now from BACKUPDATA (committed runs only).");
      if (rep.model.missing.filter(function (n) { return /^CDFS/.test(n); }).length) doc.notes.push("CDF SUMMARY has no CDFS RELEASED tables, so the email has 4 tables, not 6.");
      var a1 = data.email ? String(((data.email.values[0] || [])[0]) == null ? "" : data.email.values[0][0]).trim() : "";
      if (/^NOT CURRENT – RUN FAILED/.test(a1)) doc.notes.unshift("The last update failed (" + a1.replace(/^NOT CURRENT – RUN FAILED\s*/, "") + ") — sheet EMAIL DRAFT says NOT CURRENT. This email shows the sheets as they are, not a checked status: fix sheet ERROR's errors and run the update first.");
      else if (rep.emailDiffers && data.email) doc.notes.push("Sheet EMAIL DRAFT is older (" + (rep.emailSubjectNow || "empty") + ") — Check changes → Apply rebuilds it.");
      doc.runId = data.backup && data.backup.baseline ? data.backup.baseline.runId : "";   // the last committed run: in the file names (150.8.7)
      doc.eml = true;
      return doc;
    });
  }

  var MAILTO_LIMIT = 1900;
  function fileStamp(d) { function p(n) { return (n < 10 ? "0" : "") + n; } return d.getFullYear() + p(d.getMonth() + 1) + p(d.getDate()) + "_" + p(d.getHours()) + p(d.getMinutes()) + p(d.getSeconds()); }
  function saveFile(name, content, type) {
    try {
      var url = URL.createObjectURL(new Blob([content], { type: type + ";charset=utf-8" })), a = document.createElement("a");
      a.href = url; a.download = name; document.body.appendChild(a); a.click(); document.body.removeChild(a);
      setTimeout(function () { URL.revokeObjectURL(url); }, 4000);
      $("copyMsg").textContent = "Saved " + name + " — look in Downloads. If nothing was saved, this Excel doesn't allow downloads from add-ins: use Copy as table.";
      state.lastSaved = name;
    } catch (e) { $("copyMsg").textContent = "This Excel doesn't allow saving files from add-ins — use Copy as table."; }
  }

  function renderMail() {
    var m = state.mail, to = splitAddrs($("mailTo").value), cc = splitAddrs($("mailCc").value);
    var full = E.mailtoUrl(to, cc, m.subject, m.text);
    var tooLong = full.length > MAILTO_LIMIT;
    var url = tooLong ? E.mailtoUrl(to, cc, m.subject, "") : full;
    $("mailPreview").innerHTML =
      '<div class="mail-box"><div class="mail-head"><b>To:</b> ' + esc(to.join("; ") || "—") +
      (cc.length ? "<br><b>Cc:</b> " + esc(cc.join("; ")) : "") + "<br><b>Subject:</b> " + esc(m.subject) + "</div>" +
      '<div class="mail-body">' + m.html + "</div></div>" +
      (m.notes && m.notes.length ? '<div class="note warn" style="margin-top:8px"><ul>' + m.notes.map(function (n) { return "<li>" + esc(n) + "</li>"; }).join("") + "</ul>Notes for you — they are not part of the email.</div>" : "") +
      (tooLong ? '<div class="note warn" style="margin-top:8px">The email is too long to open directly — the table will be copied; just paste it into the draft (Ctrl+V).</div>' : "") +
      '<div class="mail-actions">' +
      '<a id="openMail" class="btn primary" href="' + esc(url) + '">Open draft in Outlook</a>' +
      '<div class="row"><button id="copyMail" class="btn" type="button" style="flex:1"><svg class="ico"><use href="#i-copy"/></svg>Copy as table</button>' +
      '<button id="bigPreviewBtn" class="btn" type="button" style="flex:1"><svg class="ico"><use href="#i-expand"/></svg>Large preview</button></div>' +
      (m.eml ? '<div class="row"><button id="saveEml" class="btn" type="button" style="flex:1"><svg class="ico"><use href="#i-download"/></svg>Save as .eml draft</button>' +
        '<button id="saveHtml" class="btn" type="button" style="flex:1"><svg class="ico"><use href="#i-download"/></svg>Save as .html</button></div>' : "") + "</div>" +
      '<div id="copyMsg" class="nr-meta"></div>';
    if (!to.length) $("copyMsg").textContent = "No recipients — you can add them in Outlook.";
    $("openMail").addEventListener("click", function () {
      if (tooLong) copyRich(m.html, m.text).then(function (ok) { $("copyMsg").textContent = ok ? "Table copied — paste it into the email body." : "Copy failed — use the copy button."; });
    });
    $("bigPreviewBtn").addEventListener("click", openBigPreview);
    if (m.eml) {
      // 150.6 / 150.8.7: the same email as an .eml draft (opens in Outlook) and as an .html page; file names with date, time and RunID
      var stamp = fileStamp(new Date()) + (m.runId ? "_" + m.runId : "");
      $("saveEml").addEventListener("click", function () {
        saveFile("CDF_Status_" + stamp + ".eml", window.CdfSummary.eml({ subject: m.subject }, m.html, m.text, to, cc), "message/rfc822");
      });
      $("saveHtml").addEventListener("click", function () {
        saveFile("CDF_Status_" + stamp + ".html", "<!doctype html><html><head><meta charset=\"utf-8\"><title>" + esc(m.subject) + "</title></head><body>" + m.html + "</body></html>", "text/html");
      });
    }
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

  // ======================================================================
  //  v1.4: pages, today at a glance, workbook stores, import pre-check,
  //        morning update, undo, CDF lookup, focused emails, big preview,
  //        theme and the action bar
  // ======================================================================
  var F = window.CdfFeatures;
  var NS_SETTINGS = "urn:cdf-updater:settings:v1", NS_HISTORY = "urn:cdf-updater:history:v1", NS_NOTES = "urn:cdf-updater:mynotes:v1", NS_STATIONS = "urn:cdf-updater:stations:v1";
  var L = window.CdfLink;
  var THEME_KEY = "cdfUpdater.theme.v1", PREVIEW_KEY = "cdfUpdater.preview.v1";
  function version() { var v = $("appVersion"); return v ? v.textContent.replace(/^v/, "") : ""; }

  // ---------------- pages (ECAM control-panel keys) ----------------
  function showPage(name) {
    state.page = name;
    Array.prototype.forEach.call(document.querySelectorAll(".ecp .key"), function (k) {
      var on = k.getAttribute("data-page") === name;
      k.classList.toggle("on", on); k.setAttribute("aria-selected", on ? "true" : "false"); k.tabIndex = on ? 0 : -1;
    });
    Array.prototype.forEach.call(document.querySelectorAll("main > .page"), function (p) { p.classList.toggle("on", p.id === "page-" + name); });
    if (name === "newrows" || name === "email") keyDot(name, null);
    if (name === "look") { renderLookPage(); if (state.inExcel && (state.looks || []).length && !state.lookValues) lookValues().then(renderLookPage); }
    updateActionBar();
    try { window.scrollTo(0, 0); } catch (e) { /* ignore */ }
  }
  function keyDot(page, kind) {
    var k = $("tab-" + page);
    if (!k) return;
    var d = k.querySelector(".dot");
    if (!kind) { if (d) k.removeChild(d); return; }
    if (state.page === page) return;
    if (!d) { d = document.createElement("span"); k.appendChild(d); }
    d.className = "dot" + (kind === "done" ? " done" : "");
  }

  // ---------------- workbook stores: custom XML parts (travel with the file, no extra sheets) ----------------
  function storesSupported() {
    try { return state.inExcel && Office.context.requirements.isSetSupported("ExcelApi", "1.5"); } catch (e) { return false; }
  }
  function readStore(ns) {
    if (!storesSupported()) return Promise.resolve(null);
    return Excel.run(function (ctx) {
      var part = ctx.workbook.customXmlParts.getByNamespace(ns).getOnlyItemOrNullObject();
      return ctx.sync().then(function () {
        if (part.isNullObject) return null;
        var x = part.getXml();
        return ctx.sync().then(function () { return F.decodeStore(x.value); });
      });
    }).catch(function () { return null; });
  }
  function writeStore(ns, obj) {
    if (!storesSupported()) return Promise.reject(new Error("This Excel version can't keep settings in the workbook (ExcelApi 1.5 needed)"));
    return Excel.run(function (ctx) {
      var parts = ctx.workbook.customXmlParts.getByNamespace(ns);
      parts.load("items");
      return ctx.sync().then(function () {
        var xml = F.encodeStore(ns, obj);
        if (parts.items.length) {
          parts.items[0].setXml(xml);
          for (var i = 1; i < parts.items.length; i++) parts.items[i].delete();
        } else ctx.workbook.customXmlParts.add(xml);
        return ctx.sync().then(function () { return true; });
      });
    });
  }
  function loadStores() {
    return readStore(NS_SETTINGS).then(function (s) {
      state.privText = (s && s.privateRules) || "";
      state.signature = (s && s.signature) || [];
      $("privText").value = state.privText;
      return readStore(NS_HISTORY);
    }).then(function (h) {
      state.history = (h && h.snapshots) || [];
      state.snap = state.history.length ? state.history[state.history.length - 1] : null;
      state.prevSnap = state.snap ? F.previousSnapshot(state.history, state.snap.at) : null;
      renderGlance(false); renderHistoryInfo();
      return readStore(NS_STATIONS);
    }).then(function (st) {
      state.stations = (st && st.map) || null;
      if (state.stations) { state.stuck = F.stuckCount(state.stations, new Date()); renderGlance(false); }
      return readStore(NS_NOTES);
    }).then(function (n) {
      state.myNotes = (n && n.data) || null;
      state.myNotesAt = (n && n.pastedAt) || null;
      adoptSharedNotes();
      renderLink();
      return readStore(NS_LOOKS);
    }).then(function (lk) { loadLooks(lk); });
  }

  // ---------------- today at a glance (upper ECAM display) ----------------
  var GL_LABEL = { open: "OPEN", conv: "CONV ENG", fa: "FINAL APPR", pending: "TO SIGN", released: "RELEASED" };
  var GL_SHORT = { "1": "INIT", "2": "POC", "3": "CONV", "4": "AVNC", "9": "F/A" };
  function showSnapshot(res, save) {
    if (!res || !res.summary) return Promise.resolve();
    var snap = F.snapshot(res.summary, res.info, new Date());
    state.prevSnap = F.previousSnapshot(state.history, snap.at);
    state.snap = snap;
    if (save) state.history = F.pushSnapshot(state.history, snap, 60);
    renderGlance(true); renderHistoryInfo();
    if (save) refreshStatusPack().catch(function () { /* the copy button builds it again */ });
    return save ? writeStore(NS_HISTORY, { snapshots: state.history }).catch(function () { /* the numbers still show */ }) : Promise.resolve();
  }
  function renderGlance(anim) {
    var s = state.snap, el = $("glance");
    if (!s) { el.innerHTML = '<div class="ecam-empty">Run <b>Check changes</b> to see today\'s numbers.</div>'; return; }
    var p = state.prevSnap, m = F.glance(s, p), at = new Date(s.at);
    var html = '<div class="ecam-top"><span>CDF STATUS ' + dateStr(at) + " " + hhmm(at) + "</span>" +
      '<span class="since">' + (p ? "SINCE " + dateStr(new Date(p.at)).slice(0, 5) + " " + hhmm(new Date(p.at)) : "FIRST RUN") + "</span></div>" +
      '<div class="metrics">' + m.map(function (x, i) {
        var d = x.delta == null ? '<div class="d flat"></div>' : x.delta === 0 ? '<div class="d flat">=</div>'
          : '<div class="d">' + (x.delta > 0 ? "▲" : "▼") + Math.abs(x.delta) + "</div>";
        return '<button type="button" class="metric' + (anim ? " anim" : "") + '" data-list="' + x.id + '" style="animation-delay:' + (i * 60) + 'ms" title="' + esc(x.label) + ' — show the list">' +
          '<div class="v">' + x.value + '</div><div class="k">' + GL_LABEL[x.id] + "</div>" + d + "</button>";
      }).join("") + "</div>";
    var bars = F.stationBars(s);
    if (bars.length) {
      html += '<div class="ecam-sep"></div><div class="ecam-cols"><div class="bars">' + bars.map(function (b, i) {
        return '<button type="button" class="bar" data-list="station:' + esc(b.key) + '" title="' + esc(b.key) + ' — show the list"><span class="bl">' + esc(GL_SHORT[b.code] || b.label.slice(0, 4).toUpperCase()) +
          '</span><span class="bt"><span class="bf" style="width:' + Math.max(b.pct, b.value ? 2 : 0) + "%;" + (anim ? "animation-delay:" + (120 + i * 60) + "ms" : "animation:none") +
          '"></span></span><span class="bv">' + b.value + "</span></button>";
      }).join("") + "</div>" + trendSvg(F.trendSeries(state.history.concat(state.history.length && state.history[state.history.length - 1].at === s.at ? [] : [s]), "open", 14)) + "</div>";
    }
    html += stuckLine();
    el.innerHTML = html;
  }
  // days at station: a lamp under the display once CDF Updater has tracked the stations
  function stuckLine() {
    var k = state.stuck;
    if (!state.stations || !k) return "";
    var first = Object.keys(state.stations).reduce(function (m, key) { var d = state.stations[key][3] || state.stations[key][1]; return !m || d < m ? d : m; }, "");
    if (!k.total) return '<div class="ecam-foot"><span class="dim">DAYS AT STATION — TRACKING SINCE ' + esc(first.split("-").reverse().slice(0, 2).join("/")) + "</span></div>";
    return '<div class="ecam-foot"><button type="button" class="stuck-lamp" data-list="stuck" title="Open CDFs ' + F.STUCK + ' or more days at the same station — show the list">' +
      '<span class="sl-v">' + k.total + "</span> CDF" + (k.total === 1 ? "" : "S") + " " + F.STUCK + "+ DAYS AT STATION</button></div>";
  }
  function trendSvg(series) {
    if (series.length < 2) return '<div class="trend"><div class="tl">OPEN TREND</div><div class="tn">appears after the 2nd run</div></div>';
    var w = 104, h = 34, vals = series.map(function (x) { return x.value; });
    var min = Math.min.apply(null, vals), max = Math.max.apply(null, vals), span = max - min || 1;
    var pts = series.map(function (x, i) {
      return [Math.round(2 + i * (w - 4) / (series.length - 1)), Math.round(h - 3 - (x.value - min) * (h - 8) / span) + 2];
    });
    var d = pts.map(function (p, i) { return (i ? "L" : "M") + p[0] + " " + p[1]; }).join(" ");
    var last = pts[pts.length - 1];
    return '<div class="trend"><div class="tl">OPEN, ' + series.length + ' RUNS</div><svg viewBox="0 0 ' + w + " " + (h + 4) + '" aria-hidden="true">' +
      '<path d="' + d + ' L' + last[0] + " " + (h + 2) + " L2 " + (h + 2) + 'Z" fill="rgba(69,212,131,.10)"/>' +
      '<path d="' + d + '" fill="none" stroke="#45d483" stroke-width="1.6" stroke-linejoin="round" stroke-linecap="round"/>' +
      '<circle cx="' + last[0] + '" cy="' + last[1] + '" r="2.6" fill="#45d483"/></svg>' +
      '<div class="tn">' + min + "–" + max + "</div></div>";
  }
  function renderHistoryInfo() {
    var n = (state.history || []).length;
    $("historyInfo").textContent = n ? "The workbook keeps today's numbers after every update: " + n + " run" + (n > 1 ? "s" : "") + " so far (for the arrows and the trend)." : "";
  }

  // ---------------- import pre-check ----------------
  function xlsxSheetInfo(buf, wanted) {
    var zip;
    function attr(el, name) { return el.getAttribute(name) || el.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id"); }
    return JSZip.loadAsync(buf).then(function (z) {
      zip = z;
      return Promise.all([zip.file("xl/workbook.xml").async("string"), zip.file("xl/_rels/workbook.xml.rels").async("string")]);
    }).then(function (x) {
      var wb = new DOMParser().parseFromString(x[0], "application/xml"), sheets = wb.getElementsByTagNameNS("*", "sheet"), rid = null;
      for (var i = 0; i < sheets.length; i++) if (i === 0 || sheets[i].getAttribute("name") === wanted) { rid = attr(sheets[i], "r:id"); if (sheets[i].getAttribute("name") === wanted) break; }
      var rels = new DOMParser().parseFromString(x[1], "application/xml").getElementsByTagNameNS("*", "Relationship"), target = null;
      for (var j = 0; j < rels.length; j++) if (rels[j].getAttribute("Id") === rid) target = rels[j].getAttribute("Target");
      if (!target) throw new Error("sheet not found in the file");
      var p = target.replace(/^\//, ""); if (p.indexOf("xl/") !== 0) p = "xl/" + p;
      var ss = zip.file("xl/sharedStrings.xml");
      return Promise.all([zip.file(p).async("string"), ss ? ss.async("string") : Promise.resolve("")]);
    }).then(function (x) {
      var xml = x[0], strings = [];
      if (x[1]) {
        var sis = new DOMParser().parseFromString(x[1], "application/xml").getElementsByTagNameNS("*", "si");
        for (var k = 0; k < sis.length; k++) strings.push(sis[k].textContent);
      }
      var first = /<row\b[^>]*>([\s\S]*?)<\/row>/.exec(xml), headers = [];
      if (first) (first[1].match(/<c\b[\s\S]*?(?:\/>|<\/c>)/g) || []).forEach(function (c) {
        var t = /\bt="(\w+)"/.exec(c), v = /<v>([\s\S]*?)<\/v>/.exec(c), is = /<t[^>]*>([\s\S]*?)<\/t>/.exec(c);
        headers.push(t && t[1] === "s" && v ? strings[+v[1]] : t && t[1] === "inlineStr" && is ? is[1] : v ? v[1] : "");
      });
      return { rows: (xml.match(/<row\b/g) || []).length, headers: headers };
    });
  }
  function sheetInfo(target) {
    if (!state.inExcel) return Promise.resolve(null);
    return Excel.run(function (ctx) {
      var ws = ctx.workbook.worksheets.getItemOrNullObject(target);
      ws.load("name");
      return ctx.sync().then(function () {
        if (ws.isNullObject) return null;
        var used = ws.getUsedRangeOrNullObject(true);
        used.load("rowIndex,rowCount,columnIndex,columnCount");
        return ctx.sync().then(function () {
          if (used.isNullObject) return { rows: 0, headers: [] };
          var hdr = ws.getRangeByIndexes(0, 0, 1, Math.max(1, (used.columnIndex || 0) + used.columnCount));
          hdr.load("values");
          return ctx.sync().then(function () { return { rows: used.rowIndex + used.rowCount, headers: hdr.values[0] }; });
        });
      });
    }).catch(function () { return null; });
  }
  function precheckSlot(slot) {
    var f = state.files[slot.slot];
    if (!f) return Promise.resolve();
    state.precheck[slot.slot] = [{ level: "ok", text: "Checking the file…", pending: true }];
    state.precheckFor[slot.slot] = f;
    renderSlots(); renderMorningFiles();
    var ext = (f.name.split(".").pop() || "").toLowerCase();
    return f.arrayBuffer().then(function (buf) {
      if (ext === "csv") { var rows = parseCsv(decodeText(buf)); return { rows: rows.length, headers: rows[0] || [] }; }
      return xlsxSheetInfo(buf, slot.from);
    }).then(function (info) {
      return sheetInfo(slot.target).then(function (cur) {
        if (state.files[slot.slot] !== f) return;
        state.precheck[slot.slot] = F.precheckImport({ expect: slot.file, fileName: f.name, lastModified: f.lastModified, now: Date.now(), sheet: info, current: cur });
      });
    }).catch(function (e) {
      state.precheck[slot.slot] = [{ level: "warn", text: "Couldn't check the file (" + errText(e) + ")." }];
    }).then(function () { renderSlots(); renderMorningFiles(); updateActionBar(); });
  }
  function precheckNew() {
    var imports = state.rules ? state.rules.imports : [];
    return Promise.all(imports.map(function (s) {
      if (!state.files[s.slot]) { delete state.precheck[s.slot]; delete state.precheckFor[s.slot]; return null; }
      return state.precheckFor[s.slot] === state.files[s.slot] ? null : precheckSlot(s);
    }));
  }
  function precheckHtml(slot) {
    var pc = state.precheck[slot];
    if (!pc || !pc.length) return "";
    return '<div class="slot-check">' + pc.map(function (x) { return '<div class="' + x.level + '">' + esc(x.text) + "</div>"; }).join("") + "</div>";
  }

  // ---------------- morning update ----------------
  var MSTEPS = [["files", "Check the files"], ["import", "Replace the sheets"], ["check", "Check changes"], ["apply", "Apply"], ["mail", "Prepare the daily email"]];
  function renderMorningFiles() {
    var imports = state.rules ? state.rules.imports : [], box = $("morningFiles");
    box.innerHTML = imports.map(function (s) {
      var f = state.files[s.slot], pc = state.precheck[s.slot] || [], st = state.slotStatus[s.slot];
      var lvl = f ? (pc.length && !pc[0].pending ? F.worstLevel(pc) : "ready") : "";
      var shown = pc.filter(function (x) { return x.level === lvl; })[0] || pc[0];
      var txt = f ? (shown ? shown.text : f.name) : st && st.kind === "done" ? "Replaced " + (st.msg.match(/\d\d:\d\d/) || [""])[0] : "Drop " + (s.file ? "the " + s.file + " file" : "the file") + " here";
      return '<label class="mfile ' + (lvl === "ok" ? "ready" : lvl) + '" title="' + esc(f ? f.name : "Replaces sheet " + s.target) + '"><b>' + esc(s.file || s.target) +
        ' → ' + esc(s.target) + "</b><span>" + esc(txt) + '</span><input type="file" accept=".xlsx,.xlsm,.csv" multiple hidden></label>';
    }).join("");
    Array.prototype.forEach.call(box.querySelectorAll(".mfile"), function (el, i) {
      var s = imports[i], input = el.querySelector("input");
      input.addEventListener("change", function () { takeImportFiles(input.files, s.slot); input.value = ""; });
      el.addEventListener("dragover", function (ev) { ev.preventDefault(); el.classList.add("drag"); });
      el.addEventListener("dragleave", function () { el.classList.remove("drag"); });
      el.addEventListener("drop", function (ev) { ev.preventDefault(); el.classList.remove("drag"); takeImportFiles(ev.dataTransfer.files, s.slot); });
    });
    var ok = !!state.rules && !state.rules.errors.length && !!state.rules.cdfAlgorithm;
    $("morning").classList.toggle("hidden", !(state.rules && state.rules.cdfAlgorithm));
    $("morningBtn").disabled = state.busy || !ok || !state.inExcel;
    $("morningBtn").textContent = chosenSlots().length ? "Run morning update" : "Run morning update (no new files)";
  }
  function setM(id, st, sub) {
    state.morning.steps[id] = { st: st, sub: sub || "" };
    $("morningSteps").innerHTML = MSTEPS.map(function (x) {
      var s = state.morning.steps[x[0]] || { st: "wait", sub: "" };
      return '<li class="' + s.st + '"><span><span class="t">' + esc(x[1]) + "</span>" + (s.sub ? '<span class="s">' + esc(s.sub) + "</span>" : "") + "</span></li>";
    }).join("");
  }
  function mStop(step, reasons, level, next) {
    setM(step, level === "bad" && !next ? "fail" : "stop", level === "bad" && !next ? "Stopped" : "Waiting for you");
    $("morningStop").innerHTML = '<div class="mstop ' + (level === "bad" ? "bad" : "") + '"><p>' + (next ? "Stopped — check this first:" : "Stopped:") + "</p><ul>" +
      reasons.map(function (r) { return "<li>" + esc(r.text) + "</li>"; }).join("") + "</ul>" +
      '<div class="row wrap">' + (next ? '<button id="mContinue" class="btn small" type="button">Continue anyway</button>' : "") +
      '<button id="mCancel" class="btn ghost small" type="button">' + (next ? "Stop here" : "Close") + "</button></div></div>";
    state.busy = false; clearStatus(); refreshButtons();
    return new Promise(function (resolve) {
      var c = $("mContinue");
      if (c) c.addEventListener("click", function () { $("morningStop").innerHTML = ""; state.busy = true; refreshButtons(); resolve(next()); });
      $("mCancel").addEventListener("click", function () { $("morningStop").innerHTML = ""; setM(step, "fail", "Stopped by you"); resolve(); });
    });
  }
  function morningRun() {
    if (state.busy) return Promise.resolve();
    state.morning = { steps: {} };
    $("morningStop").innerHTML = "";
    MSTEPS.forEach(function (x) { state.morning.steps[x[0]] = { st: "wait" }; });
    state.busy = true; state.morningRunning = true; refreshButtons();
    var slots = chosenSlots();
    setM("files", "run");
    return precheckNew().then(function () {
      var pc = [];
      slots.forEach(function (s) { (state.precheck[s.slot] || []).forEach(function (x) { pc.push({ level: x.level, text: (s.file || s.target) + ": " + x.text }); }); });
      state.morning.precheck = pc;
      if (!slots.length) { setM("files", "skip", "No new files — the sheets stay as they are"); setM("import", "skip"); return mCheck(); }
      var bad = pc.filter(function (x) { return x.level === "bad"; }), warn = pc.filter(function (x) { return x.level === "warn"; });
      if (bad.length) return mStop("files", bad.concat(warn), "bad", function () { setM("files", "done", "Continued after a warning"); return mImport(slots); });
      setM("files", "done", warn.length ? warn.length + " note" + (warn.length > 1 ? "s" : "") + " — " + warn[0].text : slots.length + " file" + (slots.length > 1 ? "s" : "") + " OK");
      return mImport(slots);
    }).catch(function (e) { mFail(e); }).then(function () { state.busy = false; state.morningRunning = false; clearStatus(); refreshButtons(); });
  }
  function mFail(e) {
    var cur = MSTEPS.filter(function (x) { return (state.morning.steps[x[0]] || {}).st === "run"; })[0];
    if (cur) setM(cur[0], "fail", errText(e));
    $("morningStop").innerHTML = '<div class="mstop bad"><p>Stopped:</p><ul><li>' + esc(errText(e)) + "</li></ul></div>";
  }
  function mImport(slots) {
    setM("import", "run", "Replacing " + slots.map(function (s) { return s.target; }).join(", ") + "…");
    return importSlots(slots).then(function (failed) {
      if (failed.length) throw new Error("Couldn't replace " + failed.map(function (s) { return s.target + " (" + state.slotStatus[s.slot].msg + ")"; }).join(", "));
      setM("import", "done", slots.map(function (s) { return s.target; }).join(" and ") + " replaced");
      return mCheck();
    });
  }
  function mCheck() {
    setM("check", "run");
    setStatus("Checking changes…");
    var reload = state.source === "github" ? loadRules() : Promise.resolve();
    return reload.then(function () {
      if (!state.rules || state.rules.errors.length) throw new Error("The rules file has errors");
      return runCdf(false);
    }).then(function (res) {
      renderCdfResults(res, false);
      showSnapshot(res, false);
      var lastSaved = state.history.length ? state.history[state.history.length - 1] : null;
      var dec = F.morningDecision(res, { prev: lastSaved });
      if (!res.summary) throw new Error(res.errors.map(function (e) { return e.message; }).join("; "));
      if (dec.stop) {
        setM("check", "done", dec.cells + " cell" + (dec.cells === 1 ? "" : "s") + " to change");
        var lvl = dec.reasons.some(function (r) { return r.level === "bad"; }) ? "bad" : "warn";
        return mStop("apply", dec.reasons, lvl, function () { return mApply(dec.cells); });
      }
      setM("check", "done", dec.cells ? dec.cells + " cell" + (dec.cells === 1 ? "" : "s") + " to change, all control totals OK" : "Nothing to update, all control totals OK");
      if (!dec.cells && !dec.rows && !dec.reports && !dec.history) { setM("apply", "skip", "Nothing to apply"); return showSnapshot(res, true).then(mMail); }
      return mApply(dec.cells);
    });
  }
  function mApply(cells) {
    setM("apply", "run", "Writing " + cells + " cells…");
    setStatus("Writing changes…");
    return runCdf(true).then(function (res) {
      renderCdfResults(res, true);
      var rt = rowOpsText(res.stage2);
      var rpt = reportsText(res.reports, true), h = res.history, hs = h && h.written ? "run " + (h.out.commits ? "recorded" : "recorded as FAILED") + " in BACKUPDATA" + (h.out.events.length ? " (" + plural(h.out.events.length, "event") + ")" : "") : "";
      setM("apply", "done", (res.cdfChanges.length + res.gccChanges.length) + " cells written" + (rt ? ", " + rt : "") + (rpt ? ", " + rpt : "") + (hs ? ", " + hs : "") + ", logged");
      return showSnapshot(res, true);
    }).then(mMail);
  }
  function mMail() {
    setM("mail", "run");
    $("mailKind").value = "daily"; $("mailSubject").value = defaultSubject();
    return buildMail(true).then(function (ok) {
      if (!ok) throw new Error($("mailPreview").innerText || "The email could not be prepared");
      setM("mail", "done", "Ready on the EMAIL page");
      keyDot("email", "attn");
      $("morningStop").innerHTML = '<div class="note ok">Morning update done ✓ <div class="row wrap"><button id="mOpenMail" class="btn small ok" type="button">Open the daily email</button>' +
        '<button id="mCopyNotes" class="btn small" type="button">Copy for My Notes</button></div>' + dueNotesHtml() + "</div>";
      $("mOpenMail").addEventListener("click", function () { showPage("email"); });
      $("mCopyNotes").addEventListener("click", function () { copyStatusForMyNotes($("mCopyNotes")); });
      wireDueNotes($("morningStop"));
      myNotesExtras($("morningStop"));
      try { $("morningStop").scrollIntoView({ behavior: "smooth", block: "center" }); } catch (e) { /* ignore */ }
    });
  }

  // ---------------- undo the last update ----------------
  function undoCheckUi() {
    if (state.busy) return;
    if (!state.rules || !state.rules.log) { $("undoBox").innerHTML = '<div class="note warn">There is no log (LOG NONE), so there is nothing to undo from.</div>'; return; }
    state.busy = true; refreshButtons(); setStatus("Reading the log…");
    var logName = state.rules.log;
    Excel.run(function (ctx) {
      return readOptionalSheet(ctx, logName).then(function (log) {
        if (!log) return { plan: null };
        var plan = F.undoPlan(log.values);
        if (!plan) return { plan: null };
        var rgs = plan.entries.map(function (e) {
          var rg = ctx.workbook.worksheets.getItem(e.sheet).getRange(e.col + e.row);
          rg.load("values,numberFormat");
          return rg;
        });
        // rows the update added must still be continuation rows of their anchor; rows it removed go back under the same CDF
        var probes = plan.rowOps.map(function (o) {
          var ws = ctx.workbook.worksheets.getItem(o.sheet), rg = o.kind === "insert" ? ws.getRange("R" + o.row + ":R" + (o.row + o.count - 1)) : ws.getRange("C" + o.anchor);
          rg.load(o.kind === "insert" ? "formulas" : "values");
          return rg;
        });
        var bkWs = plan.runOps && plan.runOps.length ? ctx.workbook.worksheets.getItemOrNullObject(plan.runOps[0].sheet) : null;
        if (bkWs) bkWs.load("name");
        return ctx.sync().then(function () {
          // BACKUPDATA removed or renamed since: the cells and rows can still be put back, the run is reported as left
          var runProbes = (plan.runOps || []).map(function (o) {
            if (!bkWs || bkWs.isNullObject) return null;
            var rg = bkWs.getRange("A" + o.row + ":E" + o.row); rg.load("values"); return rg;
          });
          return ctx.sync().then(function () {
            var cur = rgs.map(function (rg) { return show(rg.values[0][0], rg.numberFormat[0][0]); });
            var rowsNow = probes.map(function (rg, i) { return plan.rowOps[i].kind === "insert" ? rg.formulas.map(function (r) { return r[0]; }) : rg.values[0][0]; });
            var runsNow = runProbes.map(function (rg) { if (!rg) return []; var v = rg.values[0]; return [v[0], v[4]]; });
            return { plan: plan, check: F.undoCheck(plan.entries, cur), rows: F.undoRowCheck(plan.rowOps, rowsNow), runs: F.undoRunCheck(plan.runOps || [], runsNow) };
          });
        });
      });
    }).then(function (x) {
      state.undo = x;
      if (!x.plan) { $("undoBox").innerHTML = '<div class="note ok" style="margin-top:8px">Nothing to undo — the log has no update that wasn\'t undone already.</div>'; return; }
      var c = x.check, rw = x.rows, ru = x.runs || { ready: [], conflicts: [] }, nr = rw.ready.length, bad = c.conflicts.length + rw.conflicts.length + ru.conflicts.length;
      var addBack = rw.ready.filter(function (o) { return o.kind === "insert"; }), putBack = rw.ready.filter(function (o) { return o.kind === "remove"; });
      var rowsLine = (addBack.length ? "<br>" + addBack.reduce(function (n, o) { return n + o.count; }, 0) + " solution row(s) the update added will be removed (" + esc(addBack.map(function (o) { return o.key; }).join(", ")) + ")." : "") +
        (putBack.length ? "<br>" + putBack.reduce(function (n, o) { return n + o.count; }, 0) + " solution row(s) it removed will come back (" + esc(putBack.map(function (o) { return o.key; }).join(", ")) + ")." : "") +
        (ru.ready.length ? "<br>Its run in BACKUPDATA (" + esc(ru.ready.map(function (o) { return o.runId; }).join(", ")) + ") will be marked FAILED, so its events stop counting in the release tables." : "");
      $("undoBox").innerHTML = '<div class="confirm ' + (bad ? "" : "go") + '"><p>Last update: <b>' + esc(x.plan.stamp) + "</b> (" + esc(x.plan.name) + ").<br>" +
        c.ready.length + " cell" + (c.ready.length === 1 ? "" : "s") + " can be put back" +
        (c.conflicts.length ? ", " + c.conflicts.length + " changed again since then and will be left as they are:" : ".") + rowsLine + "</p>" +
        (bad ? '<ul class="undo-list">' + c.conflicts.slice(0, 20).map(function (k) { return "<li>" + esc(k.sheet + "!" + k.col + k.row + " " + k.key + ": now “" + k.now + "”, the update wrote “" + k.after + "”") + "</li>"; }).join("") +
          rw.conflicts.map(function (o) { return "<li>" + esc(o.key + ": its rows changed since the update (row " + o.row + "), so they are left as they are") + "</li>"; }).join("") +
          ru.conflicts.map(function (o) { return "<li>" + esc("BACKUPDATA row " + o.row + " no longer holds " + o.runId + " as COMMITTED, so the run is left as it is") + "</li>"; }).join("") + "</ul>" : "") +
        (c.ready.length || nr || ru.ready.length ? '<div class="row"><button id="undoGo" class="btn danger small" type="button"><svg class="ico"><use href="#i-undo"/></svg>Undo ' +
          (c.ready.length ? c.ready.length + " cell" + (c.ready.length === 1 ? "" : "s") : "") + (c.ready.length && nr ? " and " : "") + (nr ? nr + " row change" + (nr === 1 ? "" : "s") : "") + (!c.ready.length && !nr ? "the run" : "") + "</button></div>" : "") + "</div>";
      var go = $("undoGo");
      if (go) go.addEventListener("click", undoGo);
    }).catch(function (e) { $("undoBox").innerHTML = '<div class="note bad">' + esc(errText(e)) + "</div>"; })
      .then(function () { state.busy = false; clearStatus(); refreshButtons(); });
  }
  function undoGo() {
    var x = state.undo;
    if (!x || !x.plan || state.busy) return;
    var rowsReady = (x.rows && x.rows.ready) || [], runsReady = (x.runs && x.runs.ready) || [];
    state.busy = true; refreshButtons(); setStatus("Putting back " + x.check.ready.length + " cells…");
    var stamp = stampOf(new Date()), M = window.CdfMantra;
    Excel.run(function (ctx) {
      // 1. cells, in the row numbers the update left; 2. rows, bottom-up, so the ones above keep their numbers
      x.check.ready.forEach(function (e) {
        ctx.workbook.worksheets.getItem(e.sheet).getRange(e.col + e.row).values = [[safeValue(F.valueFromShown(e.before))]];
      });
      rowsReady.forEach(function (o) {
        var ws = ctx.workbook.worksheets.getItem(o.sheet), end = o.row + o.count - 1;
        if (o.kind === "insert") { ws.getRange(o.row + ":" + end).delete("Up"); return; }
        ws.getRange(o.row + ":" + end).insert("Down");
        var formulas = {};
        M.REF_COLS.forEach(function (L) { formulas[L] = M.contFormula(L, o.anchor); });
        addContinuation(ws, { anchor: o.anchor, at: o.row, end: end, mergeCols: M.MERGE_COLS, formulas: formulas, template: null });
        o.values.forEach(function (e) { ws.getRange(e.col + e.row).values = [[safeValue(F.valueFromShown(e.before))]]; });
      });
      runsReady.forEach(function (o) { ctx.workbook.worksheets.getItem(o.sheet).getRange("E" + o.row).values = [["FAILED"]]; });
      return ctx.sync().then(function () {
        return appendLog(ctx, state.rules.log, F.undoLogRows(x.plan, x.check.ready, stamp).concat(F.undoRowLogRows(x.plan, rowsReady, stamp)).concat(F.undoRunLogRows(x.plan, runsReady, stamp)));
      }).then(function () {
        if (!rowsReady.length) return null;
        setStatus("Painting sheet CDF…");
        return repaintIn(ctx, {}).catch(function () { return null; });   // rows moved back: the look follows them
      }).then(function () {
        setStatus("Rebuilding CDF SUMMARY and EMAIL DRAFT…");
        return rebuildReports(ctx, new Date()).catch(function (e) { return { failed: errText(e) }; });
      });
    }).then(function (rep) {
      state.lookupCache = null;
      var rl = rep && rep.failed ? " CDF SUMMARY couldn't be rebuilt (" + esc(rep.failed) + ") — run Check changes." : rep && (rep.wroteSummary || rep.wroteEmail) ? " " + esc(reportsText(rep, true)) + " from the sheet as it is now." : "";
      $("undoBox").innerHTML = '<div class="note ok" style="margin-top:8px">Undone ✓ — ' + x.check.ready.length + " cells are back to their values before " + esc(x.plan.stamp) +
        (rowsReady.length ? ", and the solution rows are as they were" : "") + (runsReady.length ? ". The run is marked FAILED in BACKUPDATA" : "") + ". The undo is in the log too." + rl + "</div>";
      $("results").innerHTML = "";
    }).catch(function (e) { $("undoBox").innerHTML = '<div class="note bad">' + esc(errText(e)) + "</div>"; })
      .then(function () { state.busy = false; clearStatus(); refreshButtons(); });
  }

  // ---------------- CDF lookup ----------------
  function openLookup() { $("lookup").classList.remove("hidden"); $("lookupInput").focus(); $("lookupInput").select(); }
  function closeLookup() { $("lookup").classList.add("hidden"); }
  function lookupData() {
    var c = state.lookupCache;
    if (c && Date.now() - c.at < 120000) return Promise.resolve(c.data);
    return Excel.run(function (ctx) {
      var data = {}, log = state.rules && state.rules.log;
      return readSheets(ctx, ["CDF", "GCC"].concat(log ? [log] : [])).then(function (rd) {
        var sh = rd.sheets;
        if (!sh.CDF) throw new Error("Sheet \"CDF\" was not found in this workbook");
        data.cdf = sh.CDF.values; data.gcc = sh.GCC ? sh.GCC.values : []; data.log = log && sh[log] ? sh[log].values : [];
        state.lookupCache = { at: Date.now(), data: data }; return data;
      });
    });
  }
  // opts.go: select the row right away (a CDF number pasted into the pane); opts.also: more numbers that were pasted
  function doLookup(q, opts) {
    opts = opts || {};
    if (q != null) $("lookupInput").value = q;
    var query = $("lookupInput").value;
    if (!state.inExcel) { $("lookupResult").innerHTML = '<div class="note warn">This only works inside Excel.</div>'; return Promise.resolve(); }
    $("lookupResult").innerHTML = '<div class="hint">Looking…</div>';
    return lookupData().then(function (data) {
      var r = F.lookupCdf(data, query);
      renderLookup(r, opts);
      if (r.found && opts.go) return selectCdfRow(r.row).then(function () {
        var n = $("lkSelected");
        if (n) { n.textContent = "Row " + r.row + " is selected in sheet CDF ✓"; n.classList.remove("hidden"); }
      }).catch(function () { /* the button still works */ });
    }).catch(function (e) { $("lookupResult").innerHTML = '<div class="note bad">' + esc(errText(e)) + "</div>"; });
  }
  function selectCdfRow(row) {
    return Excel.run(function (ctx) { ctx.workbook.worksheets.getItem("CDF").getRange("A" + row + ":AB" + row).select(); return ctx.sync(); });
  }
  // CDF numbers in pasted text, as lookup queries ("961", "740.A", "961.-")
  function pastedCdfQueries(text) {
    text = String(text || "").trim();
    if (!text || text.length > 20000) return [];
    var q = L ? L.refs(text).map(function (r) { return r.n + (r.rev ? "." + r.rev : ""); }) : [];
    if (!q.length && /^0*\d{1,4}(\.([A-Z]{1,2}|-))?$/i.test(text)) q = [text.replace(/^0+(?=\d)/, "")];
    return q.slice(0, 8);
  }
  // paste a CDF number anywhere in the pane (not in a text box) → the lookup opens on it and selects its row
  function onPanePaste(ev) {
    var t = ev.target, tag = t && t.tagName;
    if (tag === "INPUT" || tag === "TEXTAREA" || tag === "SELECT" || (t && t.isContentEditable)) return;
    var text = ev.clipboardData && ev.clipboardData.getData("text/plain");
    if (L && L.kindOf(text) === L.KIND_NOTES) {   // notes from My Notes, pasted anywhere
      ev.preventDefault(); showPage("setup"); takeMyNotes(text); return;
    }
    if (L && L.kindOf(text) === L.KIND_STATUS) {
      ev.preventDefault(); showPage("setup"); linkMsg("bad", L.decode(text, L.KIND_NOTES).error); return;
    }
    var q = pastedCdfQueries(text);
    if (!q.length) return;
    ev.preventDefault();
    openLookup();
    doLookup(q[0], { go: true, also: q.slice(1) });
  }
  function onLookupPaste(ev) {
    var q = pastedCdfQueries(ev.clipboardData && ev.clipboardData.getData("text/plain"));
    if (!q.length) return;
    ev.preventDefault();
    doLookup(q[0], { go: true, also: q.slice(1) });
  }
  function renderLookup(r, opts) {
    opts = opts || {};
    var box = $("lookupResult");
    var also = (opts.also || []).length ? '<div class="lk-also"><span class="hint">Also pasted:</span>' + opts.also.map(function (s) {
      return '<button class="facet" type="button" data-q="' + esc(s) + '">' + esc(s) + "</button>";
    }).join("") + "</div>" : "";
    if (!r.found) {
      box.innerHTML = '<div class="note warn">' + esc(r.error) + "</div>" + (r.suggestions && r.suggestions.length
        ? '<div class="lk-sugg">' + r.suggestions.map(function (s) { return '<button class="facet" type="button" data-q="' + esc(s) + '">' + esc(s) + "</button>"; }).join("") + "</div>" : "") + also;
      wireAlso(box, opts);
      return;
    }
    var f = r.fields;
    function pillCls(s) { s = (s || "").toUpperCase(); return s === "OPEN" ? "open" : s === "RELEASED" ? "released" : s === "CANCELLED" ? "cancelled" : ""; }
    function kv(label, v) { return v === "" || v == null || String(v).trim() === "-" ? "" : "<dt>" + esc(label) + "</dt><dd>" + esc(v) + "</dd>"; }
    var received = typeof f.received === "number" ? E.serialToDateString(f.received, "dd/mm/yyyy") : f.received;
    box.innerHTML =
      '<div class="lk-title"><b>' + esc(r.key) + "</b><span>row " + r.row + " in sheet CDF</span></div>" +
      '<div id="lkSelected" class="lk-selected hidden"></div>' + also +
      (r.others.length ? '<div class="hint">Also: ' + esc(r.others.join(", ")) + "</div>" : "") +
      '<div class="lk-pills">' + (f.status ? '<span class="pill ' + pillCls(f.status) + '">' + esc(f.status) + "</span>" : "") +
      (f.station ? '<span class="pill">' + esc(f.station) + "</span>" : "") + (f.signature && f.signature !== "-" ? '<span class="pill">✎ ' + esc(f.signature) + "</span>" : "") + "</div>" +
      pathHtml(f.status, f.station, state.stations ? F.daysAt(state.stations[r.key], new Date()) : null) +
      (f.description ? '<div class="lk-desc">' + esc(f.description) + "</div>" : "") +
      '<dl class="lk-grid">' + kv("Received", received) + kv("Priority", f.priority) + kv("Tags", f.tags) + kv("DWG.", f.dwg + (f.dwgRev && f.dwgRev !== "-" ? " rev " + f.dwgRev : "")) +
      kv("DWG. name", f.dwgName) + kv("NRC", f.nrc) + kv("Author", f.author) + kv("Technical adviser", f.adviser) + kv("Send to", f.sendTo) +
      kv("Deviation reason", f.reason) + kv("Minor / major", f.classification) + "</dl>" +
      '<div class="lk-h">Solutions</div>' + (r.solutions.length ? r.solutions.map(function (s) {
        var rel = /RELEASED$/i.test(s.mantra);
        return '<div class="lk-sol"><div><b>' + esc(s.doc || "(no document)") + "</b>" + (s.doc && s.doc !== "-" ? " Rev " + esc(s.rev || "-") : "") +
          '<div class="rw">row ' + s.row + (s.w && s.w !== "-" ? ", " + esc(s.w) : "") + (s.x && s.x !== "-" ? ", " + esc(s.x) : "") + "</div></div>" +
          (s.mantra ? '<span class="m ' + (rel ? "rel" : "wip") + '">' + esc(s.mantra) + "</span>" : "<span></span>") + "</div>";
      }).join("") : '<div class="hint">No solution rows yet.</div>') +
      '<div class="lk-h">GCC</div>' + (r.gcc ? '<dl class="lk-grid">' + kv("Row", r.gcc.row) + kv("Status", r.gcc.status) + kv("Station", r.gcc.station) +
        kv("Assignee", r.gcc.assignee) + kv("Final approval", r.gcc.finalApproval) + "</dl>" : '<div class="hint">Not in the GCC sheet.</div>') +
      lookupNotesHtml(r.key) +
      '<div class="lk-h">History</div>' + (r.history.length ? '<ul class="lk-hist">' + r.history.slice(0, 30).map(function (h) {
        return '<li><span class="w">' + esc(h.when) + '</span><span class="c">' + esc(h.col) + "</span><span>" + esc(h.before || "(empty)") + " → <b>" + esc(h.after || "(empty)") + "</b>" + (h.by === F.UNDO ? " (undo)" : "") + "</span></li>";
      }).join("") + "</ul>" : '<div class="hint">No changes in the log yet.</div>') +
      '<div class="row" style="margin-top:12px"><button id="lkGo" class="btn small" type="button">Go to row ' + r.row + "</button></div>";
    $("lkGo").addEventListener("click", function () {
      selectCdfRow(r.row).then(closeLookup).catch(function () { /* ignore */ });
    });
    wireAlso(box, opts);
  }
  // the other numbers from the same paste keep their "select the row" behaviour
  function wireAlso(box, opts) {
    Array.prototype.forEach.call(box.querySelectorAll("[data-q]"), function (b) {
      b.addEventListener("click", function () {
        var q = b.getAttribute("data-q"), rest = (opts.also || []).filter(function (x) { return x !== q; });
        var pasted = (opts.also || []).indexOf(q) >= 0;
        doLookup(q, pasted ? { go: true, also: [$("lookupInput").value].concat(rest) } : {});
      });
    });
  }

  // ---------------- the CDFs behind a number on the display ----------------
  function listData() {
    return lookupData().then(function (data) {
      var res = state.lastRes && state.lastRes.summary && Date.now() - (state.lastResAt || 0) < 10 * 60000 ? state.lastRes : null;
      var summary = res ? res.summary : null;
      if (!summary && data.gcc && data.gcc.length) {
        try { summary = window.CdfAlgo.run({ cdf: data.cdf, gcc: data.gcc, now: new Date() }, state.rules.cdfConfig).summary; } catch (e) { summary = null; }
      }
      var gccKeys = null;
      if (data.gcc && data.gcc.length) { gccKeys = {}; data.gcc.forEach(function (r, i) { if (i) gccKeys[E.loose(E.normalizeCdf(r[0]))] = true; }); }
      return { cdf: data.cdf, summary: summary, stations: state.stations || {}, gccKeys: gccKeys };
    });
  }
  function daysText(r) {
    if (r.days == null) return "";
    if (r.exact) return r.days + " d at station";
    return r.min >= F.STUCK ? r.min + "+ d at station" : "open " + r.days + " d";
  }
  function openList(kind) {
    if (!state.inExcel) return;
    $("list").classList.remove("hidden");
    $("listTitle").textContent = "Loading…"; $("listCount").textContent = ""; $("listBody").innerHTML = "";
    listData().then(function (data) {
      var l = F.glanceList(kind, data, new Date()), max = 400;
      $("listTitle").textContent = l.title;
      $("listCount").textContent = l.rows.length + " CDF" + (l.rows.length === 1 ? "" : "s");
      var html = (l.notInGcc.length ? '<div class="note warn">The number on the display counts CDFs that are in GCC. ' + l.notInGcc.length + " more " + (l.notInGcc.length === 1 ? "is" : "are") +
        " only in sheet CDF: " + esc(l.notInGcc.slice(0, 6).join(", ")) + (l.notInGcc.length > 6 ? " …" : "") + "</div>" : "") +
        (kind !== "released" && l.rows.some(function (r) { return r.days != null && !r.exact; }) ? '<p class="hint">"open N d": the day it reached its station isn\'t known yet — CDF Updater learns it from now on. "N+ d": at least that long.</p>' : "") +
        '<ul class="cl-list">' + l.rows.slice(0, max).map(function (r) {
          var ph = L ? L.phaseOf(r.status, r.station) : { index: -1 };
          return '<li><button type="button" class="cl-row" data-key="' + esc(r.key) + '" data-row="' + r.row + '">' +
            '<span class="cl-key">' + esc(r.key) + "</span>" +
            '<span class="cl-days' + (r.stuck ? " stuck" : "") + (r.exact ? "" : " approx") + '">' + esc(daysText(r)) + "</span>" +
            '<span class="cl-st">' + (ph.index >= 0 && L ? '<i class="ph ph-' + ph.index + '"></i>' : "") + esc(r.station || r.status) + "</span>" +
            '<span class="cl-who">' + esc(r.ready ? "Ready for " + r.ready : r.signature || r.supplier) + "</span></button></li>";
        }).join("") + "</ul>" + (l.rows.length > max ? '<p class="hint">and ' + (l.rows.length - max) + " more</p>" : "") +
        (l.rows.length ? "" : '<p class="hint">' + (kind === "stuck" ? "None yet. CDF Updater counts the days from the first update it sees each CDF at its station." : "No CDFs.") + "</p>");
      $("listBody").innerHTML = html;
      Array.prototype.forEach.call($("listBody").querySelectorAll(".cl-row"), function (b) {
        b.addEventListener("click", function () {
          closeList(); openLookup();
          doLookup(lookupQuery(b.getAttribute("data-key")), { go: true });
        });
      });
    }).catch(function (e) { $("listTitle").textContent = "Couldn't list"; $("listBody").innerHTML = '<div class="note bad">' + esc(errText(e)) + "</div>"; });
  }
  function closeList() { $("list").classList.add("hidden"); }
  // "A330-1-961.-" → "961.-" (what the lookup understands, with the revision)
  function lookupQuery(key) { var p = L && L.keyParts(key); return p ? p.n + (p.rev ? "." + p.rev : "") + (p.extra ? "" : "") : key; }

  // ---------------- the path of a CDF (same picture as in My Notes) ----------------
  function pathHtml(status, station, at) {
    if (!L) return "";
    var ph = L.phaseOf(status, station), cls = ph.cancelled ? " cancelled" : ph.index === 3 ? " fa" : ph.index === 4 ? " rel" : "";
    var segs = L.PHASES.map(function (p, i) {
      var st = ph.cancelled ? "" : i < ph.index ? " done" : i === ph.index ? " cur" : "";
      return '<span class="fseg' + st + '" title="' + esc(p.label) + '"><i></i><b>' + esc(p.id) + "</b></span>";
    }).join("");
    var days = "";
    if (at && /^open$/i.test(String(status).trim())) {
      days = at.exact ? at.days + " days at " + esc(String(station).replace(/^\d+\s*-\s*/, ""))
        : at.min >= F.STUCK ? "at least " + at.min + " days at " + esc(String(station).replace(/^\d+\s*-\s*/, "")) : "open " + at.days + " days";
      days = '<div class="fdays' + (at.stuck ? " stuck" : "") + '">' + days + "</div>";
    }
    return '<div class="fpath' + cls + '" role="img" aria-label="' + esc(ph.cancelled ? "Cancelled" : (L.PHASES[ph.index] || { label: status }).label) + '">' + segs + "</div>" + days;
  }

  // ---------------- My Notes lines at the end of the morning update ----------------
  function myNotesExtras(box) {
    if (!L || !state.myNotes) return;
    var pack = state.statusPack;
    (pack ? Promise.resolve(pack) : refreshStatusPack()).then(function (p) {
      if (!p) return;
      var data = L.decode(p.text, L.KIND_STATUS).data;
      var closable = L.closable(state.myNotes.notes, data);
      if (!closable.length || !box.isConnected) return;
      var div = document.createElement("div");
      div.className = "due-notes rel-notes";
      div.innerHTML = "<b>" + closable.length + " open note" + (closable.length === 1 ? " is" : "s are") + " on CDFs that were released — close " + (closable.length === 1 ? "it" : "them") + " in My Notes:</b>" +
        closable.slice(0, 6).map(function (n) { var k = n.refs[0]; return '<button type="button" class="facet" data-lk="' + esc(lookupQuery(k)) + '" title="' + esc(n.text) + '">' + esc(k.replace(/^A330-1-/, "")) + "</button>"; }).join("");
      box.querySelector(".note").appendChild(div);
      wireDueNotes(div);
    }).catch(function () { /* the line is optional */ });
  }
  function unknownRefsHtml() {
    if (!L || !state.myNotes || !state.statusPack) return "";
    var u = L.unknownRefs(state.myNotes.notes, L.decode(state.statusPack.text, L.KIND_STATUS).data);
    if (!u.length) return "";
    return '<div class="note warn">Notes in My Notes mention CDF numbers that are not in sheet CDF — a typo? ' +
      u.slice(0, 8).map(function (x) { return "<b>" + esc(x.ref) + "</b>" + (x.notes > 1 ? " (" + x.notes + " notes)" : ""); }).join(", ") + "</div>";
  }

  // ---------------- big email preview ----------------
  function openBigPreview() {
    var m = state.mail;
    if (!m) return;
    var payload = JSON.stringify({ subject: m.subject, html: m.html, text: m.text });
    try { localStorage.setItem(PREVIEW_KEY, payload); } catch (e) { /* ignore */ }
    function overlay() {
      $("bigPreviewTitle").textContent = m.subject;
      $("bigPreviewBody").innerHTML = m.html;
      $("bigPreview").classList.remove("hidden");
    }
    try {
      if (window.Office && Office.context && Office.context.ui && Office.context.ui.displayDialogAsync) {
        var url = location.href.replace(/taskpane\.html[^#]*.*$/, "preview.html?v=" + version());
        Office.context.ui.displayDialogAsync(url, { height: 80, width: 55, displayInIframe: true }, function (r) {
          if (r.status !== "succeeded") { overlay(); return; }
          var dlg = r.value;
          dlg.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
            if (arg.message === "ready" && dlg.messageChild) { try { dlg.messageChild(payload); } catch (e) { /* ignore */ } }
            if (arg.message === "close") dlg.close();
          });
        });
        return;
      }
    } catch (e) { /* fall back */ }
    overlay();
  }

  // ---------------- workbook settings ----------------
  function savePrivate() {
    var text = $("privText").value, bad = F.privateLines(text), msg = $("privMsg");
    if (bad.length) { msg.className = "msg bad"; msg.textContent = "Line " + bad[0].line + ": " + bad[0].message; return; }
    var errs = E.parseRules(text).errors;
    if (errs.length) { msg.className = "msg bad"; msg.textContent = "Line " + errs[0].line + ": " + errs[0].message; return; }
    writeStore(NS_SETTINGS, { privateRules: text, signature: state.signature || [], savedAt: new Date().toISOString() }).then(function () {
      state.privText = text;
      applyRulesText(state.rulesText);
      msg.className = "msg ok";
      msg.textContent = "Saved in this workbook ✓ (" + state.rules.privateLines + " lines). Save the workbook (Ctrl+S) to keep them, then remove these lines from rules.txt on GitHub.";
    }).catch(function (e) { msg.className = "msg bad"; msg.textContent = errText(e); });
  }
  function movePrivate() {
    var x = F.extractPrivate(state.rulesText), msg = $("privMsg");
    if (!x.moved) { msg.className = "msg"; msg.textContent = "rules.txt has no settings lines left to move."; return; }
    var cur = $("privText").value.trim();
    $("privText").value = (cur ? cur + "\n" : "") + x.moved;
    msg.className = "msg ok";
    msg.textContent = x.moved.split("\n").length + " lines copied here. Click Save to workbook.";
  }

  // ---------------- theme ----------------
  function themeChoice() { try { return localStorage.getItem(THEME_KEY) || "auto"; } catch (e) { return "auto"; } }
  function officeIsDark() {
    try {
      var t = window.Office && Office.context && Office.context.officeTheme;
      if (t && t.bodyBackgroundColor) {
        var h = t.bodyBackgroundColor.replace("#", ""), rgb = [0, 2, 4].map(function (i) { return parseInt(h.substr(i, 2), 16) / 255; });
        return 0.2126 * rgb[0] + 0.7152 * rgb[1] + 0.0722 * rgb[2] < 0.4;
      }
    } catch (e) { /* ignore */ }
    try { return !!(window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches); } catch (e) { return false; }
  }
  function applyTheme() {
    var c = themeChoice(), dark = c === "dark" || (c === "auto" && officeIsDark());
    document.documentElement.setAttribute("data-theme", dark ? "dark" : "light");
    Array.prototype.forEach.call(document.querySelectorAll("[data-theme-choice]"), function (b) { b.classList.toggle("on", b.getAttribute("data-theme-choice") === c); });
  }

  // ---------------- action bar: the next thing to do on this page ----------------
  var AB_TARGETS = {
    today: ["#applyBtn", "#previewBtn"],
    newrows: ["#writeAllBtn", "#writeRowBtn"],
    email: ["#openMail", "#mailBuild"],
    setup: ["#undoGo"]
  };
  function updateActionBar() {
    var ab = $("abPrimary"), list = AB_TARGETS[state.page || "today"] || [], target = null;
    for (var i = 0; i < list.length && !target; i++) {
      var el = document.querySelector("#page-" + (state.page || "today") + " " + list[i]);
      if (el && el.offsetParent !== null) target = el;
    }
    if (target !== state.abTarget) watchTarget(target);
    state.abTarget = target;
    if (!target) { ab.classList.add("hidden"); return; }
    ab.textContent = target.textContent.trim();
    ab.disabled = !!target.disabled || state.busy;
    ab.className = "btn ab-btn " + (/\bok\b|\bdanger\b/.test(target.className) ? "ok" : "primary") + (state.abOnscreen ? " onscreen" : "");
  }
  // The bar only offers the button while the real one is scrolled out of view.
  var abIo = null;
  function watchTarget(el) {
    state.abOnscreen = false;
    if (!window.IntersectionObserver) return;
    if (!abIo) abIo = new IntersectionObserver(function (entries) {
      entries.forEach(function (en) {
        if (en.target !== state.abTarget) return;
        state.abOnscreen = en.isIntersecting;
        $("abPrimary").classList.toggle("onscreen", en.isIntersecting);
      });
    }, { rootMargin: "0px 0px -56px 0px" });
    abIo.disconnect();
    if (el) abIo.observe(el);
  }
  function watchActionBar() {
    var t = null;
    var mo = new MutationObserver(function () { clearTimeout(t); t = setTimeout(updateActionBar, 30); });
    mo.observe(document.querySelector("main"), { childList: true, subtree: true, attributes: true, attributeFilter: ["disabled", "class"] });
    $("abPrimary").addEventListener("click", function () { if (state.abTarget && !state.abTarget.disabled) state.abTarget.click(); });
  }

  // ---------------- focused emails ----------------
  function focusMailKinds() {
    var cfg = state.rules && state.rules.cdfConfig, out = [];
    if (!state.rules || !state.rules.cdfAlgorithm || !cfg) return out;
    var seen = {};
    Object.keys(cfg.suppliers || {}).forEach(function (k) {
      var s = cfg.suppliers[k];
      if (!seen[s]) { seen[s] = true; out.push({ id: "sup:" + s, type: "supplier", supplier: s, label: "Open at conversion eng. — " + s }); }
    });
    (cfg.pendingTables || []).forEach(function (p) { out.push({ id: "pend:" + p.name, type: "pending", table: p.name, person: p.person, label: "Ready to sign — " + p.person }); });
    return out;
  }
  function focusKind(id) { return focusMailKinds().filter(function (k) { return k.id === id; })[0]; }
  function buildFocus(ctx, kind, subject) {
    var data = {};
    return readSheets(ctx, ["CDF", "GCC", "EMAIL DRAFT"]).then(function (rd) {
      var sh = rd.sheets, draft = sh["EMAIL DRAFT"];
      if (!sh.CDF) throw new Error("Sheet \"CDF\" was not found in this workbook");
      if (!sh.GCC) throw new Error("Sheet \"GCC\" was not found — import CDF_List_Export first");
      data.cdf = sh.CDF; data.gcc = sh.GCC;
      var res = window.CdfAlgo.run({ cdf: data.cdf.values, gcc: data.gcc.values, now: new Date() }, state.rules.cdfConfig);
      if (!res.summary) throw new Error("The CDF algorithm stopped: " + res.errors.map(function (e) { return e.message; }).join("; "));
      var name = draftSignerName(draft);
      var doc = F.buildFocusDoc(kind, res.summary, new Date(), name ? [kind.type === "supplier" ? "Best regards," : "בברכה,", name] : null);
      doc.daily = true;
      if (subject) doc.subject = subject;
      return doc;
    });
  }

  // ---------------- link with My Notes (Outlook) ----------------
  function shortDate(iso) { var d = new Date(iso); return isNaN(d) ? "" : dateStr(d) + " " + hhmm(d); }
  // the CDF status of every CDF, packed for My Notes; also left in the shared storage
  var LINK_MISSING = "cdf-link.js didn't load — upload it to GitHub with the other files, then reopen the add-in.";
  function refreshStatusPack() {
    if (!L) return Promise.reject(new Error(LINK_MISSING));
    if (!state.inExcel || !state.rules || !state.rules.cdfAlgorithm) return Promise.resolve(null);
    return Excel.run(function (ctx) {
      var data = {}, logName = state.rules.log;
      return readSheets(ctx, ["CDF", "GCC"].concat(logName ? [logName] : [])).then(function (rd) {
        var sh = rd.sheets, lg = logName ? sh[logName] : null;
        if (!sh.CDF) throw new Error("Sheet \"CDF\" was not found in this workbook");
        data.cdf = sh.CDF.values; data.gcc = sh.GCC;
        var summary = null, g = data.gcc, now = new Date();
        if (g) { try { var res = window.CdfAlgo.run({ cdf: data.cdf, gcc: g.values, now: now }, state.rules.cdfConfig); summary = res.summary; } catch (e) { /* status without "ready to sign" */ } }
        var tr = F.trackStations(state.stations, data.cdf, lg ? lg.values : [], now);
        state.stations = tr.map; state.stuck = F.stuckCount(tr.map, now);
        if (tr.changed && storesSupported()) writeStore(NS_STATIONS, { map: tr.map, at: now.toISOString() }).catch(function () { /* kept for this session */ });
        renderGlance(false);
        var pack = F.bridgeStatus(data.cdf, summary, now, "CDF Updater " + version(), tr.map);
        var text = L.encode(L.KIND_STATUS, pack);
        state.statusPack = { at: pack.at, count: pack.count, text: text };
        L.writeShared(window.localStorage, L.LS_STATUS, text);
        renderLink();
        return state.statusPack;
      });
    });
  }
  function copyStatusForMyNotes(btn) {
    var fresh = state.statusPack && Date.now() - Date.parse(state.statusPack.at) < 10 * 60000;
    var label = btn.textContent;
    btn.disabled = true; btn.textContent = "Preparing…";
    (fresh ? Promise.resolve(state.statusPack) : refreshStatusPack()).then(function (p) {
      if (!p) throw new Error("The CDF algorithm isn't on in rules.txt");
      return copyText(p.text).then(function (ok) {
        var box = $("linkCopyBox");
        if (ok) {
          box.classList.add("hidden");
          btn.textContent = "Copied ✓ — paste it in My Notes";
          linkMsg("ok", "Copied the status of " + p.count + " CDFs. In Outlook: My Notes → Reports → CDF status → paste.");
        } else {
          box.value = p.text; box.classList.remove("hidden"); box.focus(); box.select();
          btn.textContent = label;
          linkMsg("", "Press Ctrl+C to copy the selected text, then paste it in My Notes.");
          showPage("setup");
        }
        setTimeout(function () { btn.textContent = label; }, 4000);
      });
    }).catch(function (e) { btn.textContent = label; linkMsg("bad", errText(e)); })
      .then(function () { btn.disabled = false; });
  }
  function linkMsg(cls, text) { var m = $("linkMsg"); m.className = "msg " + (cls || ""); m.textContent = text; }
  function takeMyNotes(text) {
    if (!L) { linkMsg("bad", LINK_MISSING); return Promise.resolve(false); }
    var r = L.decode(text, L.KIND_NOTES);
    if (!r.ok) { linkMsg("bad", r.error); return Promise.resolve(false); }
    return saveMyNotes(r.data, false).then(function () {
      $("linkPaste").value = "";
      var i = F.notesInfo(r.data);
      linkMsg("ok", "Read ✓ — " + i.notes + " note" + (i.notes === 1 ? "" : "s") + " on " + i.cdfs + " CDF" + (i.cdfs === 1 ? "" : "s") +
        " (copied " + shortDate(r.data.at) + "). They show in the CDF lookup." + (storesSupported() ? " Saved in this workbook." : ""));
      return true;
    });
  }
  function saveMyNotes(data, auto) {
    state.myNotes = data; state.myNotesAt = new Date().toISOString(); state.myNotesAuto = !!auto;
    state.lookupCache = null;
    renderLink();
    if (!storesSupported()) return Promise.resolve();
    return writeStore(NS_NOTES, { data: data, pastedAt: state.myNotesAt }).catch(function () { /* still used in this session */ });
  }
  // when Excel and Outlook share the storage on this computer, the notes arrive without pasting
  function adoptSharedNotes() {
    if (!L) return;
    var shared = L.readShared(window.localStorage, L.LS_NOTES, L.KIND_NOTES);
    if (shared && L.newer(state.myNotes, shared) === shared && shared !== state.myNotes && (!state.myNotes || shared.at !== state.myNotes.at)) saveMyNotes(shared, true);
  }
  function renderLink() {
    var tag = $("linkTag"), out = $("linkOutInfo");
    if (!tag) return;
    var n = state.myNotes, i = F.notesInfo(n);
    tag.textContent = n ? i.notes + " note" + (i.notes === 1 ? "" : "s") + " from Outlook" : "";
    out.textContent = state.statusPack ? "Ready: " + state.statusPack.count + " CDFs, as of " + shortDate(state.statusPack.at) + "." : "The status of every CDF in this workbook.";
    var uk = $("linkUnknown");
    if (uk) uk.innerHTML = unknownRefsHtml();
    if (n && !$("linkMsg").textContent) {
      linkMsg("", "Notes from My Notes: " + i.notes + " on " + i.cdfs + " CDFs, copied " + shortDate(n.at) + (state.myNotesAuto ? " — arrived automatically." : "."));
    }
  }
  function stateCls(st) { return st === "Overdue" ? "late" : st === "Due today" ? "today" : st === "Done" ? "done" : st === "Info" ? "info" : "open"; }
  function lookupNotesHtml(key) {
    if (!L || !state.myNotes) return "";
    var list = L.notesFor(key, state.myNotes.notes);
    return '<div class="lk-h">Notes in Outlook <span class="lk-sub">My Notes, copied ' + esc(shortDate(state.myNotes.at)) + "</span></div>" + (list.length ? '<ul class="lk-notes">' + list.slice(0, 20).map(function (n) {
      return '<li class="' + stateCls(n.state) + '"><div class="ln-top"><span class="ln-st">' + esc(n.state || "Open") + "</span>" + (n.due ? '<span class="ln-due">Due ' + esc(n.due.split("-").reverse().join("/")) + "</span>" : "") +
        (n.cat ? '<span class="ln-cat">' + esc(n.cat) + "</span>" : "") + '</div><div class="ln-text" dir="auto">' + esc(n.text) + "</div>" +
        '<div class="ln-mail" dir="auto">' + esc((n.subject || "") + (n.from ? " — " + n.from : "")) + "</div></li>";
    }).join("") + "</ul>" : '<div class="hint">No notes on this CDF in My Notes.</div>');
  }
  function dueNotesHtml() {
    if (!L) return "";
    var due = L.dueNotes(state.myNotes && state.myNotes.notes, new Date());
    if (!due.length) return "";
    return '<div class="due-notes"><b>' + due.length + " follow-up" + (due.length === 1 ? "" : "s") + " in My Notes " + (due.length === 1 ? "is" : "are") + " due:</b>" +
      due.slice(0, 6).map(function (n) {
        var k = (n.refs || [])[0] || "";
        return '<button type="button" class="facet" data-lk="' + esc(k) + '" title="' + esc(n.text) + '">' + esc(k.replace(/^A330-1-/, "")) + "<b>" + esc(n.due.split("-").reverse().slice(0, 2).join("/")) + "</b></button>";
      }).join("") + (due.length > 6 ? '<span class="hint"> and ' + (due.length - 6) + " more</span>" : "") + "</div>";
  }
  function wireDueNotes(box) {
    Array.prototype.forEach.call(box.querySelectorAll("[data-lk]"), function (b) {
      b.addEventListener("click", function () { openLookup(); doLookup(b.getAttribute("data-lk")); });
    });
  }

  // The sender's name from the end of sheet EMAIL DRAFT (the line after "בברכה,").
  // the signature lines of EMAIL DRAFT; after a failed run (NOT CURRENT) the ones kept in the workbook settings
  function signatureOf(values) {
    var lines = window.CdfSummary ? window.CdfSummary.readSignature(values || []) : [];
    return lines.length ? lines : (state.signature || []);
  }
  function draftSignerName(draft) {
    var lines = signatureOf(draft ? draft.values : []);
    lines = lines.filter(function (l) { return !/^(בברכה|best regards|regards|thanks)[,.]?$/i.test(l); });
    return lines.length ? lines[lines.length - 1] : "";
  }

  // ---------------- small helpers ----------------
  function refreshButtons() {
    var ok = !!state.rules && !state.rules.errors.length;
    // missing people settings stop the CDF algorithm only: importing the exports doesn't need them
    var okImport = !!state.rules && !state.rules.errors.some(function (e) { return !e.setting; });
    $("importBtn").disabled = state.busy || !okImport || !state.inExcel || !state.canImport || !chosenSlots().length;
    var n = chosenSlots().length;
    $("importBtn").textContent = n > 1 ? "Replace " + n + " sheets" : n === 1 ? "Replace sheet" : "Replace sheets";
    $("previewBtn").disabled = state.busy || !ok || !state.inExcel;
    var ab = $("applyBtn"); if (ab) ab.disabled = state.busy;
    var hasRowSpec = ok && state.rules.newRow.cells.length > 0;
    $("parsePaste").disabled = state.busy || !hasRowSpec || !state.inExcel || !$("pasteBox").value.trim();
    $("mailBuild").disabled = state.busy || !ok || !state.inExcel;
    var wb = $("writeRowBtn"); if (wb) wb.disabled = state.busy;
    var mb = $("morningBtn"); if (mb) mb.disabled = state.busy || !ok || !state.inExcel || !state.rules.cdfAlgorithm;
    $("undoBtn").disabled = state.busy || !state.inExcel;
    $("repaintBtn").disabled = state.busy || !state.inExcel;
    var la = $("lookAdd"); if (la) la.disabled = state.busy || !state.inExcel;
    $("privSave").disabled = state.busy || !state.inExcel;
    updateActionBar();
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
