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
    mail: null              // last built email {subject, text, html}
  };
  var GROUPS_KEY = "cdfUpdater.mailGroups.v1";

  var $ = function (id) { return document.getElementById(id); };

  // ---------------- startup ----------------
  function start(info) {
    state.inExcel = !!(info && window.Office && info.host === Office.HostType.Excel);
    if (!state.inExcel) {
      showBanner("התוסף פתוח מחוץ לאקסל: אפשר לבדוק את קובץ הכללים, אבל פעולות על החוברת זמינות רק בתוך Excel.", false);
    } else {
      state.canImport = Office.context.requirements.isSetSupported("ExcelApi", "1.13");
      if (!state.canImport) {
        showBanner("גרסת ה-Excel הזו לא תומכת בהחלפת גיליון מקובץ (נדרש Microsoft 365). הרצת האלגוריתם עדיין עובדת.", true);
      }
    }
    try { localStorage.removeItem("cdfUpdater.apiKey.v1"); localStorage.removeItem("cdfUpdater.aiMode.v1"); } catch (e) { /* ignore */ }
    wire();
    restoreLocalRules();
    loadRules();
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
    $("pasteBox").addEventListener("input", refreshButtons);
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
    renderMailKinds(); renderGroups();
  }

  // ---------------- rules ----------------
  function restoreLocalRules() {
    try {
      var saved = JSON.parse(localStorage.getItem(LS_KEY) || "null");
      if (saved && saved.text) { state.source = "local"; state.localName = saved.name; state.rulesText = saved.text; }
    } catch (e) { /* storage unavailable */ }
  }

  function loadRules() {
    setPill("טוען כללים…", "");
    var p = state.source === "local"
      ? Promise.resolve(state.rulesText)
      : fetch("rules.txt?v=" + Date.now(), { cache: "no-store" }).then(function (r) {
          if (!r.ok) throw new Error("לא הצלחתי לקרוא את rules.txt (HTTP " + r.status + ")");
          return r.text();
        });
    return p.then(applyRulesText).catch(function (e) {
      state.rules = null;
      setPill("אין כללים", "bad");
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
    setPill(errs ? errs + " שגיאות בכללים" : "כללים תקינים", errs ? "bad" : "ok");
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
      ? "קובץ מקומי: <b dir=\"ltr\">" + esc(state.localName || "") + "</b> (נשמר במחשב הזה עד שתחזור לקובץ מ-GitHub)."
      : "<b dir=\"ltr\">rules.txt</b> מתוך GitHub — נקרא מחדש בכל לחיצה על \"בדוק שינויים\".";
    if (state.loadedAt) src += " נטען ב-" + hhmm(state.loadedAt) + ".";
    $("rulesSource").innerHTML = src;
    $("useGithub").classList.toggle("hidden", state.source !== "local");
    $("rulesText").textContent = state.rulesText;
    var r = state.rules;
    $("rulesErrors").innerHTML = r && r.errors.length
      ? r.errors.map(function (e) {
          return '<div class="err">שורה ' + e.line + ": " + esc(e.message) + (e.text ? '<br><code>' + esc(e.text) + "</code>" : "") + "</div>";
        }).join("")
      : "";
  }

  function renderAlgos() {
    var r = state.rules;
    if (!r || !r.algorithms.length) { $("algos").innerHTML = '<p class="hint">לא הוגדר אלגוריתם בקובץ הכללים.</p>'; return; }
    $("algos").innerHTML = r.algorithms.map(function (a) {
      var steps = a.actions.map(function (x) { return "<li>" + esc(E.describeAction(x)) + "</li>"; }).join("");
      var oc = a.onChange.map(function (x) { return '<li class="onchange">אם השורה השתנתה — ' + esc(E.describeAction(x)) + "</li>"; }).join("");
      return '<div class="algo">' +
        '<div class="algo-name">' + esc(a.name) + "</div>" +
        '<div class="algo-meta">בסיס: <b>' + (a.baseFirst ? "הגיליון הראשון" : esc(a.base)) + "</b> · עדכון: <b>" + esc(a.update || "?") + "</b><br>" +
        (a.match ? "התאמה: <b dir=\"ltr\">UPDATE." + a.match.update.col + " = BASE." + a.match.base.col + "</b> · " : "") +
        "נתונים משורה " + a.baseStartRow + " / " + a.updateStartRow + "</div>" +
        "<ol>" + steps + oc + "</ol></div>";
    }).join("");
  }

  // ---------------- import ----------------
  function renderSlots() {
    var imports = state.rules ? state.rules.imports : [];
    var valid = {};
    imports.forEach(function (s) { valid[s.slot] = true; });
    Object.keys(state.files).forEach(function (k) { if (!valid[k]) delete state.files[k]; });

    if (!imports.length) { $("slots").innerHTML = '<p class="hint">אין שורות IMPORT בקובץ הכללים.</p>'; return; }

    $("slots").innerHTML = "";
    imports.forEach(function (s) {
      var f = state.files[s.slot], st = state.slotStatus[s.slot];
      var el = document.createElement("label");
      el.className = "slot" + (f ? " has-file" : "") + (st && !f ? " " + st.kind : "");
      el.innerHTML =
        '<div class="slot-title"><span>קובץ ' + s.slot + "</span>" +
        (f ? '<button type="button" class="slot-clear" title="הסר">✕</button>' : "") + "</div>" +
        '<div class="slot-target">→ גיליון <b>' + esc(s.target) + "</b>" + (s.from ? " (מתוך \"" + esc(s.from) + "\")" : "") + "</div>" +
        '<div class="slot-file">' + (f ? esc(f.name) + " · " + kb(f.size) : "לחץ לבחירה או גרור קובץ לכאן (xlsx / xlsm / csv)") + "</div>" +
        (st && !f ? '<div class="slot-msg">' + esc(st.msg) + "</div>" : "") +
        '<input type="file" accept=".xlsx,.xlsm,.csv" hidden>';
      var input = el.querySelector("input");
      input.addEventListener("change", function () { if (input.files[0]) setFile(s.slot, input.files[0]); });
      var clr = el.querySelector(".slot-clear");
      if (clr) clr.addEventListener("click", function (ev) { ev.preventDefault(); ev.stopPropagation(); delete state.files[s.slot]; renderSlots(); refreshButtons(); });
      el.addEventListener("dragover", function (ev) { ev.preventDefault(); el.classList.add("drag"); });
      el.addEventListener("dragleave", function () { el.classList.remove("drag"); });
      el.addEventListener("drop", function (ev) {
        ev.preventDefault(); el.classList.remove("drag");
        if (ev.dataTransfer.files[0]) setFile(s.slot, ev.dataTransfer.files[0]);
      });
      $("slots").appendChild(el);
    });
  }

  function setFile(slot, file) {
    state.files[slot] = file;
    delete state.slotStatus[slot];
    $("importConfirm").classList.add("hidden");
    renderSlots(); refreshButtons();
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
      if (replace.length) html += "הגיליונות <b>" + replace.map(esc).join(", ") + "</b> יימחקו ויוחלפו בתוכן הקבצים. נוסחאות בגיליונות אחרים שמפנות אליהם יהפכו ל-#REF!.";
      if (create.length) html += (html ? "<br>" : "") + "ייווצרו גיליונות חדשים: <b>" + create.map(esc).join(", ") + "</b>.";
      $("importConfirmText").innerHTML = html + "<br>להמשיך?";
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
        setStatus("מחליף את הגיליון " + s.target + "…");
        return importOne(s, f).then(function (info) {
          delete state.files[s.slot];
          state.slotStatus[s.slot] = { kind: "done", msg: "הוחלף ✓ " + hhmm(new Date()) + " · " + f.name + (info ? " · " + info : "") };
        }).catch(function (e) {
          state.slotStatus[s.slot] = { kind: "failed", msg: errText(e) };
        }).then(renderSlots);
      });
    });
    chain.then(function () { state.busy = false; clearStatus(); refreshButtons(); });
  }

  function importOne(slot, file) {
    var ext = (file.name.split(".").pop() || "").toLowerCase();
    return file.arrayBuffer().then(function (buf) {
      if (ext === "csv") return importCsv(slot, buf);
      if (ext !== "xlsx" && ext !== "xlsm") throw new Error("סוג קובץ לא נתמך (." + ext + "). שמור אותו כ-xlsx ונסה שוב.");
      return workbookSheetNames(buf).then(function (names) {
        if (!names.length) throw new Error("לא נמצאו גיליונות בקובץ");
        var src = slot.from || names[0];
        if (names.indexOf(src) < 0) throw new Error("בקובץ אין גיליון בשם \"" + src + "\". קיימים: " + names.join(", "));
        var b64 = toBase64(buf);
        return replaceSheet(slot.target, function (ctx, rel) {
          var opts = { sheetNamesToInsert: [src] };
          if (rel) { opts.positionType = Excel.WorksheetPositionType.after; opts.relativeTo = rel.tmpName; }
          else opts.positionType = Excel.WorksheetPositionType.end;
          var res = ctx.workbook.insertWorksheetsFromBase64(b64, opts);
          return ctx.sync().then(function () {
            if (!res.value || !res.value.length) throw new Error("Excel לא הוסיף את הגיליון");
            return res.value[0];
          });
        }).then(function () { return names.length > 1 ? "גיליון \"" + src + "\"" : ""; });
      });
    });
  }

  function importCsv(slot, buf) {
    var rows = parseCsv(decodeText(buf));
    if (!rows.length) throw new Error("קובץ ה-CSV ריק");
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
    }).then(function () { return rows.length + " שורות"; });
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
      if (!wb) throw new Error("הקובץ לא נראה כמו חוברת Excel תקינה");
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
    if (!state.inExcel) return Promise.reject(new Error("פעולה זו זמינה רק בתוך Excel"));
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
    setStatus("קורא כללים…");
    var reload = state.source === "github" ? loadRules() : Promise.resolve();
    reload.then(function () {
      if (!state.rules || state.rules.errors.length) throw new Error("יש שגיאות בקובץ הכללים — תקן ונסה שוב.");
      if (!state.inExcel) throw new Error("פעולה זו זמינה רק בתוך Excel");
      setStatus("משווה נתונים…");
      return runAlgorithms(false);
    }).then(function (results) {
      state.lastPreview = results;
      renderResults(results, false);
    }).catch(showError).then(function () { state.busy = false; clearStatus(); refreshButtons(); });
  }

  function applyChanges() {
    if (state.busy) return;
    state.busy = true; refreshButtons();
    setStatus("כותב שינויים…");
    runAlgorithms(true).then(function (results) {
      state.lastPreview = null;
      renderResults(results, true);
    }).catch(showError).then(function () { state.busy = false; clearStatus(); refreshButtons(); });
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
            if (!algo.baseFirst && baseWs.isNullObject) throw new Error("גיליון הבסיס \"" + algo.base + "\" לא נמצא בחוברת");
            if (updWs.isNullObject) throw new Error("גיליון העדכון \"" + algo.update + "\" לא נמצא — ייבאת את הקובץ?");
            if (baseWs.name === updWs.name) throw new Error("גיליון הבסיס וגיליון העדכון הם אותו גיליון (" + baseWs.name + ")");
            var data = [];
            return readSheet(ctx, baseWs).then(function (b) { data.push(b); return readSheet(ctx, updWs); })
              .then(function (u) { data.push(u); return data; });
          }).then(function (data) {
            var res = E.computeChanges(algo, data[0], data[1], { now: now });
            res.algo = algo; res.baseName = baseWs.name; res.updateName = updWs.name;
            out.push(res);
            if (apply && res.changes.length) {
              setStatus("כותב " + res.changes.length + " תאים לגיליון " + baseWs.name + "…");
              return writeChanges(ctx, baseWs, res.changes);
            }
          });
        });
      });
      return chain.then(function () {
        if (apply && rules.log) { setStatus("כותב יומן…"); return writeLog(ctx, rules.log, out, now); }
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
        hdr.values = [["זמן הרצה", "אלגוריתם", "גיליון", "שורה", "מפתח (CDF)", "עמודה", "ערך קודם", "ערך חדש", "הערה"]];
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
      html += '<p class="hint">בסיס: <b>' + esc(res.baseName) + "</b> · עדכון: <b>" + esc(res.updateName) + "</b></p>";
      html += '<div class="tiles">' +
        tile(s.uniqueUpdateKeys, "מזהים בעדכון") +
        tile(s.matchedKeys, "נמצאו בבסיס") +
        tile(s.unmatched.length, "לא נמצאו", s.unmatched.length ? "warnt" : "") +
        tile(s.changedRows, applied ? "שורות עודכנו" : "שורות ישתנו", "hl") +
        tile(s.changedCells, applied ? "תאים עודכנו" : "תאים ישתנו", "hl") +
        tile(s.matchedRows - s.changedRows, "ללא שינוי") +
        "</div>";

      if (applied && res.changes.length) html += '<div class="note ok">העדכון נכתב ✓' + (state.rules.log ? " — כל השינויים והערכים הקודמים נשמרו בגיליון <b>" + esc(state.rules.log) + "</b>." : "") + "</div>";
      if (!res.changes.length) html += '<div class="note ok">אין מה לעדכן — גיליון הבסיס כבר תואם לקובץ העדכון.</div>';

      if (s.duplicateUpdateKeys.length) {
        html += '<div class="note warn">מזהים שמופיעים יותר מפעם אחת בקובץ העדכון (נלקחה השורה האחרונה):' + chips(s.duplicateUpdateKeys) + "</div>";
      }
      if (s.unmatched.length) {
        html += '<div class="note warn">' + s.unmatched.length + " מזהים מקובץ העדכון לא נמצאו בגיליון הבסיס:" +
          chips(s.unmatched.map(function (u) { return u.key; })) + "</div>";
      }
      if (res.changes.length) html += changesTable(res.changes);
      html += "</div>";
    });

    if (!applied && total > 0) {
      html += '<div class="confirm go"><p>ייכתבו <b>' + total + "</b> תאים. מומלץ לשמור את הקובץ לפני — ייתכן שלא ניתן יהיה לבטל עם Ctrl+Z" +
        (state.rules.log ? ", אבל היומן שומר את הערכים הקודמים" : "") + ".</p>" +
        '<div class="row"><button id="applyBtn" class="btn ok">החל ' + total + " שינויים</button>" +
        '<button id="cancelApply" class="btn ghost">ביטול</button></div></div>';
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
        '</td><td class="v old">' + esc(show(c.oldValue, c.oldFormat) || "(ריק)") + '</td><td class="v new">' + esc(show(c.newValue, c.numberFormat || c.oldFormat) || "(ריק)") + "</td></tr>";
    }).join("");
    var more = changes.length > MAX_ROWS_SHOWN ? '<div class="more">ועוד ' + (changes.length - MAX_ROWS_SHOWN) + " שינויים…</div>" : "";
    return '<div class="tbl-wrap"><table><thead><tr><th>שורה</th><th>מפתח</th><th>עמ׳</th><th>לפני</th><th>אחרי</th></tr></thead><tbody>' +
      rows + "</tbody></table>" + more + "</div>";
  }

  function tile(n, label, cls) { return '<div class="tile ' + (cls || "") + '"><div class="n">' + n + '</div><div class="l">' + label + "</div></div>"; }
  function chips(list) {
    var shown = list.slice(0, 60).map(function (k) { return '<span class="chip">' + esc(k) + "</span>"; }).join("");
    return '<div class="chips">' + shown + (list.length > 60 ? '<span class="chip">+' + (list.length - 60) + "</span>" : "") + "</div>";
  }

  // ================= new row (via claude.ai) =================
  function showNewRowNote(msg, kind) {
    $("newRowForm").innerHTML = '<div class="note ' + (kind || "bad") + '">' + esc(msg) + "</div>";
  }

  function fetchPrompt() {
    return fetch("ai-prompt.txt?v=" + Date.now(), { cache: "no-store" })
      .then(function (r) { return r.ok ? r.text() : null; })
      .catch(function () { return null; })
      .then(function (t) { return t && t.trim() ? t : CdfPrompt.DEFAULT_PROMPT; });
  }

  function newRowSheet(ctx) {
    var s = state.rules.newRow.sheet;
    var ws = s.first ? ctx.workbook.worksheets.getFirst(true) : ctx.workbook.worksheets.getItemOrNullObject(s.name);
    ws.load("name");
    return ctx.sync().then(function () {
      if (!s.first && ws.isNullObject) throw new Error("הגיליון \"" + s.name + "\" (NEWROW_SHEET) לא נמצא");
      return ws;
    });
  }

  function copyPrompt() {
    fetchPrompt().then(function (base) {
      var text = CdfPrompt.pastePrompt(base);
      return copyText(text).then(function (ok) {
        $("copyPromptMsg").textContent = ok
          ? "ההנחיה הועתקה ✓ — הדבק אותה ב-Claude יחד עם הקבצים."
          : "ההעתקה נכשלה בסביבה הזו — פתח \"הצג את תוכן הקובץ\" למטה והעתק ידנית.";
      });
    });
  }

  function openClaude() {
    var url = "https://claude.ai/new";
    try {
      if (window.Office && Office.context && Office.context.ui && Office.context.ui.openBrowserWindow) { Office.context.ui.openBrowserWindow(url); return; }
    } catch (e) { /* fall back */ }
    window.open(url, "_blank");
  }

  function parsePasted() {
    if (state.busy) return;
    var res = E.parseAiPaste($("pasteBox").value);
    if (res.error) { showNewRowNote(res.error, "bad"); return; }
    state.queue = res.items.map(function (it) {
      return { fields: it, main: it.main_file || null, att: it.att_file || null };
    });
    state.queueTotal = state.queue.length;
    nextInQueue();
  }

  // ---- one draft at a time from the queue ----
  function nextInQueue() {
    if (!state.queue || !state.queue.length) return;
    var item = state.queue.shift();
    state.busy = true; refreshButtons();
    setStatus("מחפש את המקום בקובץ…");
    prepareDraft(item).then(function () { /* rendered */ })
      .catch(function (e) { showNewRowNote(errText(e), "bad"); })
      .then(function () { state.busy = false; clearStatus(); refreshButtons(); });
  }

  function prepareDraft(item) {
    var nr = state.rules.newRow, algo = state.rules.algorithms[0];
    var built = E.buildNewRowCells(nr, {
      ai: item.fields, mainName: item.main, attName: item.att,
      now: new Date(), dateFormat: algo ? algo.dateFormat : "dd/mm/yyyy"
    });
    return Excel.run(function (ctx) {
      return newRowSheet(ctx).then(function (ws) {
        var maxIdx = Math.max.apply(null, nr.cells.map(function (c) { return c.idx; }).concat([nr.key.idx]));
        var used = ws.getUsedRangeOrNullObject(true);
        used.load("rowIndex,rowCount");
        var hdr = ws.getRangeByIndexes(nr.headerRow - 1, 0, 1, maxIdx + 1);
        hdr.load("values");
        return ctx.sync().then(function () {
          var rows = used.isNullObject ? 0 : used.rowIndex + used.rowCount;
          var keyRg = rows ? ws.getRangeByIndexes(0, nr.key.idx, rows, 1) : null;
          if (keyRg) keyRg.load("values");
          return ctx.sync().then(function () {
            var keyVals = keyRg ? keyRg.values.map(function (r) { return r[0]; }) : [];
            var target = E.findNewRowTarget(keyVals, nr.startRow, built.cdf);
            var rowRg = ws.getRangeByIndexes(target.row - 1, 0, 1, maxIdx + 1);
            rowRg.load("values,numberFormat");
            return ctx.sync().then(function () {
              return { ws: ws.name, headers: hdr.values[0], target: target, existing: rowRg.values[0], existingFmt: rowRg.numberFormat[0] };
            });
          });
        });
      });
    }).then(function (place) {
      state.draft = { built: built, place: place, item: item };
      renderDraft();
    });
  }

  function queuePos() {
    var total = state.queueTotal || 1, left = state.queue ? state.queue.length : 0;
    return { n: total - left, total: total };
  }

  function renderDraft() {
    var d = state.draft, p = d.place, item = d.item, q = queuePos();
    var flags = (item.fields.flags || []).concat(d.built.warnings);
    var html = q.total > 1 ? '<div class="queue"><span>חריג ' + q.n + " מתוך " + q.total + "</span><span dir=\"ltr\">" + esc(d.built.cdf) + "</span></div>" : "";
    html += '<div class="nr-target">' + (p.target.existing
      ? "CDF " + esc(d.built.cdf) + " כבר קיים בשורה " + p.target.row + " בגיליון " + esc(p.ws) + " — השורה תתעדכן"
      : "שורה חדשה " + p.target.row + " בגיליון " + esc(p.ws)) + "</div>";
    if (flags.length) html += '<div class="note warn">לבדוק:<ul>' + flags.map(function (f) { return "<li>" + esc(f) + "</li>"; }).join("") + "</ul></div>";

    d.built.cells.forEach(function (c, i) {
      var header = E.cleanHeader(p.headers[c.idx]) || "ללא כותרת";
      var old = p.existing[c.idx];
      var oldShow = show(old, p.existingFmt[c.idx]);
      html += '<div class="nr-field"><div class="nr-label"><span class="col">' + c.col + '</span><span class="hdr">' + esc(header) +
        '</span><span class="src">' + esc(c.label) + "</span></div>";
      if (c.kind === "ai" && c.label.indexOf("DESCRIPTION") >= 0) {
        html += '<textarea class="inp" dir="ltr" data-i="' + i + '">' + esc(c.display) + "</textarea>";
      } else {
        html += '<input class="inp" dir="ltr" data-i="' + i + '" value="' + esc(c.display) + '"' +
          (c.kind === "link" ? ' placeholder="שם הקובץ, למשל A330-1-NNN.-.pdf"' : "") + ">";
        if (c.kind === "link" && c.prefix) html += '<div class="nr-link">קישור: ' + esc(c.prefix) + "&lt;שם הקובץ&gt;</div>";
      }
      if (!E.isEmpty(old) && oldShow !== c.display) html += '<div class="nr-old">קיים בתא: ' + esc(oldShow) + " — יוחלף</div>";
      html += "</div>";
    });

    var more = state.queue && state.queue.length;
    html += '<div class="confirm go" style="margin-top:12px"><p>שדה ריק לא נכתב (ולא מוחק ערך קיים). הכל יירשם ב-' + esc(state.rules.log || "יומן") + ".</p>" +
      '<div class="row wrap"><button id="writeRowBtn" class="btn ok">' + (p.target.existing ? "עדכן שורה " : "הוסף שורה ") + p.target.row + "</button>" +
      (more ? '<button id="skipRowBtn" class="btn ghost">דלג</button>' : "") +
      '<button id="cancelRowBtn" class="btn ghost">' + (more ? "בטל הכל" : "ביטול") + "</button></div></div>";
    $("newRowForm").innerHTML = html;
    $("writeRowBtn").addEventListener("click", writeDraft);
    var sk = $("skipRowBtn");
    if (sk) sk.addEventListener("click", function () { state.draft = null; nextInQueue(); });
    $("cancelRowBtn").addEventListener("click", function () { state.draft = null; state.queue = []; $("newRowForm").innerHTML = ""; });
  }

  function writeDraft() {
    var d = state.draft;
    if (!d || state.busy) return;
    // take the edited values from the form
    var inputs = $("newRowForm").querySelectorAll("[data-i]");
    Array.prototype.forEach.call(inputs, function (el) {
      var c = d.built.cells[+el.getAttribute("data-i")];
      var v = el.value.trim();
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
    setStatus("כותב לקובץ…");
    var now = new Date(), stamp = stampOf(now), row = d.place.target.row;
    Excel.run(function (ctx) {
      return newRowSheet(ctx).then(function (ws) {
        var log = [];
        d.built.cells.forEach(function (c) {
          if (E.isEmpty(c.value)) return;
          var rg = ws.getRangeByIndexes(row - 1, c.idx, 1, 1);
          if (c.link) {
            rg.hyperlink = { address: c.link.address, textToDisplay: c.link.textToDisplay };
          } else {
            if (c.format) rg.numberFormat = [[c.format]];
            rg.values = [[safeValue(c.value)]];
          }
          var old = d.place.existing[c.idx];
          log.push([stamp, E.LOG_NOTES.newRow, ws.name, String(row), d.built.cdf, c.col,
            show(old, d.place.existingFmt[c.idx]), c.display, E.LOG_NOTES.newRow]);
        });
        return ctx.sync().then(function () {
          return state.rules.log ? appendLog(ctx, state.rules.log, log) : null;
        }).then(function () { return log.length; });
      });
    }).then(function (n) {
      var cdf = d.built.cdf;
      state.draft = null;
      state.written = (state.written || []).concat([{ cdf: cdf, row: row, cells: n }]);
      if (state.queue && state.queue.length) {
        state.busy = false;
        nextInQueue();
        return true; // next draft is loading
      }
      var done = state.written; state.written = [];
      $("pasteBox").value = "";
      $("newRowForm").innerHTML = '<div class="note ok">' + (done.length > 1
        ? "נכתבו " + done.length + " חריגים: " + done.map(function (w) { return esc(w.cdf) + " (שורה " + w.row + ")"; }).join(", ")
        : "נכתבו " + n + " תאים בשורה " + row + " (CDF " + esc(cdf) + ")") + " ✓" +
        '<div class="row" style="margin-top:6px"><button id="mailNewRow" class="btn small">שלח במייל</button></div></div>';
      $("mailNewRow").addEventListener("click", function () {
        $("mailKind").value = "newrows"; $("mailSubject").value = defaultSubject();
        $("mailCard").open = true; $("mailCard").scrollIntoView({ behavior: "smooth" });
      });
      return false;
    }).catch(function (e) { showNewRowNote(errText(e), "bad"); return false; })
      .then(function (loadingNext) { if (!loadingNext) { state.busy = false; clearStatus(); refreshButtons(); } });
  }

  // ================= email =================
  var MAIL_KINDS = [
    { id: "changes", label: "שינויים מההרצה האחרונה", subject: "CDF status update" },
    { id: "notfound", label: "CDF שלא נמצאו בהרצה האחרונה", subject: "CDFs not found in tracker" },
    { id: "newrows", label: "שורות חדשות שנוספו היום", subject: "New CDF entries" }
  ];

  function renderMailKinds() {
    var sel = $("mailKind"), keep = sel.value;
    var opts = MAIL_KINDS.map(function (k) { return '<option value="' + k.id + '">' + esc(k.label) + "</option>"; });
    (state.rules ? state.rules.reports : []).forEach(function (r, i) {
      opts.push('<option value="report:' + i + '">דוח: ' + esc(r.name) + "</option>");
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
    if (!groups.length) { box.innerHTML = '<span class="hint" style="margin:0">אין קבוצות שמורות עדיין.</span>'; return; }
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
    if (!name || !addrs.length) { $("groupName").placeholder = "צריך שם ונמענים בשדה To"; return; }
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
    setStatus("מכין מייל…");
    Excel.run(function (ctx) {
      var headerRow = state.rules.newRow.headerRow || 1;
      if (kind.indexOf("report:") === 0) {
        var rep = state.rules.reports[+kind.split(":")[1]];
        var ws = rep.sheet.first ? ctx.workbook.worksheets.getFirst(true) : ctx.workbook.worksheets.getItemOrNullObject(rep.sheet.name);
        ws.load("name");
        return ctx.sync().then(function () {
          if (!rep.sheet.first && ws.isNullObject) throw new Error("הגיליון \"" + rep.sheet.name + "\" לא נמצא");
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

      if (!state.rules.log) throw new Error("אין יומן (LOG NONE) — אין ממה לבנות את המייל");
      var logWs = ctx.workbook.worksheets.getItemOrNullObject(state.rules.log);
      logWs.load("name");
      return ctx.sync().then(function () {
        if (logWs.isNullObject) throw new Error("עדיין אין גיליון " + state.rules.log + " — הרץ עדכון או הוסף שורה קודם");
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
      if (doc.empty) { $("mailPreview").innerHTML = '<div class="note warn">אין נתונים לשליחה בפלט הזה.</div>'; state.mail = null; return; }
      state.mail = E.buildEmail(doc);
      renderMail();
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
      (tooLong ? '<div class="note warn" style="margin-top:8px">המייל ארוך מדי לפתיחה ישירה — הטבלה תועתק, ובטיוטה רק להדביק (Ctrl+V).</div>' : "") +
      '<div class="mail-actions">' +
      '<a id="openMail" class="btn primary" href="' + esc(url) + '">פתח טיוטה ב-Outlook</a>' +
      '<button id="copyMail" class="btn" type="button">העתק כטבלה (להדבקה במייל)</button></div>' +
      '<div id="copyMsg" class="nr-meta"></div>';
    if (!to.length) $("copyMsg").textContent = "אין נמענים — אפשר להוסיף ב-Outlook.";
    $("openMail").addEventListener("click", function () {
      if (tooLong) copyRich(m.html, m.text).then(function (ok) { $("copyMsg").textContent = ok ? "הטבלה הועתקה — הדבק בגוף המייל." : "ההעתקה נכשלה — השתמש בכפתור ההעתקה."; });
    });
    $("copyMail").addEventListener("click", function () {
      copyRich(m.html, m.text).then(function (ok) { $("copyMsg").textContent = ok ? "הועתק ✓ — הדבק בגוף המייל (Ctrl+V)." : "ההעתקה נכשלה בסביבה הזו."; });
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
    $("importBtn").textContent = n > 1 ? "החלף " + n + " גיליונות" : "החלף גיליון";
    $("previewBtn").disabled = state.busy || !ok || !state.inExcel;
    var ab = $("applyBtn"); if (ab) ab.disabled = state.busy;
    var hasRowSpec = ok && state.rules.newRow.cells.length > 0;
    $("parsePaste").disabled = state.busy || !hasRowSpec || !state.inExcel || !$("pasteBox").value.trim();
    $("mailBuild").disabled = state.busy || !ok || !state.inExcel;
    var wb = $("writeRowBtn"); if (wb) wb.disabled = state.busy;
  }

  function setStatus(t) { $("statusText").textContent = t; $("status").classList.remove("hidden"); }
  function clearStatus() { $("status").classList.add("hidden"); }
  function setPill(t, cls) { var p = $("rulesPill"); p.textContent = t; p.className = "pill " + (cls || ""); }
  function showBanner(t, bad) { var b = $("banner"); b.textContent = t; b.className = "banner" + (bad ? " bad" : ""); }
  function showError(e) { $("results").innerHTML = '<div class="note bad">' + esc(errText(e)) + "</div>"; }

  function errText(e) {
    if (!e) return "שגיאה לא ידועה";
    var msg = e.message || String(e);
    if (e.code === "AccessDenied" || /protect/i.test(msg)) msg += " (ייתכן שהחוברת או הגיליון מוגנים)";
    if (e.code === "UnsupportedFeature") msg += " (הקובץ מכיל רכיב שלא נתמך בהעתקה, למשל PivotTable או גרף ב-Excel באינטרנט)";
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
