/*
 * CDF Updater - rules engine
 * Pure logic (no Office.js): parses rules.txt and computes the cell changes.
 * Works in the browser (window.CdfEngine) and in Node (module.exports) for tests.
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.CdfEngine = factory();
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  // ---------- helpers ----------
  function colToIndex(col) {
    var s = String(col).toUpperCase();
    if (!/^[A-Z]{1,3}$/.test(s)) throw new Error("Invalid column: " + col);
    var n = 0;
    for (var i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64);
    return n - 1;
  }

  function indexToCol(idx) {
    var n = idx + 1, s = "";
    while (n > 0) {
      var m = (n - 1) % 26;
      s = String.fromCharCode(65 + m) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }

  function isEmpty(v) {
    return v === null || v === undefined || (typeof v === "string" && v.trim() === "");
  }

  // value equality used to decide whether a cell really changes
  function sameValue(a, b) {
    var na = isEmpty(a) ? "" : (typeof a === "string" ? a.trim() : String(a));
    var nb = isEmpty(b) ? "" : (typeof b === "string" ? b.trim() : String(b));
    return na === nb;
  }

  // loose text used for MATCH keys and IF comparisons
  function loose(v) {
    if (isEmpty(v)) return "";
    return String(v).trim().replace(/\s+/g, " ").toUpperCase();
  }

  function dateToSerial(d, withTime) {
    var utc = Date.UTC(d.getFullYear(), d.getMonth(), d.getDate());
    var serial = (utc - Date.UTC(1899, 11, 30)) / 86400000;
    if (withTime) serial += (d.getHours() * 3600 + d.getMinutes() * 60 + d.getSeconds()) / 86400;
    return serial;
  }

  // ---------- tokenizer ----------
  function tokenize(line) {
    var tokens = [], i = 0, n = line.length;
    while (i < n) {
      var ch = line[i];
      if (ch === " " || ch === "\t" || ch === "\r") { i++; continue; }
      if (ch === "#") break; // comment
      if (ch === '"' || ch === "“" || ch === "”") {
        var j = i + 1, buf = "";
        while (j < n) {
          var c = line[j];
          if (c === '"' && line[j + 1] === '"') { buf += '"'; j += 2; continue; }
          if (c === '"' || c === "“" || c === "”") break;
          buf += c; j++;
        }
        if (j >= n) throw new Error("Missing closing quote");
        tokens.push({ t: "str", v: buf });
        i = j + 1; continue;
      }
      if (line.substr(i, 2) === "->") { tokens.push({ t: "op", v: "->" }); i += 2; continue; }
      if (line.substr(i, 2) === "<>" || line.substr(i, 2) === "!=") { tokens.push({ t: "op", v: "<>" }); i += 2; continue; }
      if (ch === "=") { tokens.push({ t: "op", v: "=" }); i++; continue; }
      if (ch === ",") { tokens.push({ t: "op", v: "," }); i++; continue; }
      var m =/^(-?\d+(\.\d+)?(?![A-Za-z_.])|[A-Za-z0-9_.]+)/.exec(line.slice(i));
      if (m) {
        var w = m[0];
        if (/^-?\d+(\.\d+)?$/.test(w)) tokens.push({ t: "num", v: parseFloat(w) });
        else tokens.push({ t: "word", v: w });
        i += w.length; continue;
      }
      throw new Error("Unexpected character: " + ch);
    }
    return tokens;
  }

  // ---------- parser ----------
  function Cursor(tokens) { this.tk = tokens; this.p = 0; }
  Cursor.prototype.peek = function () { return this.tk[this.p]; };
  Cursor.prototype.next = function () { return this.tk[this.p++]; };
  Cursor.prototype.done = function () { return this.p >= this.tk.length; };
  Cursor.prototype.word = function (expected) {
    var t = this.next();
    if (!t || t.t !== "word") throw new Error("Expected " + (expected || "a word"));
    if (expected && t.v.toUpperCase() !== expected) throw new Error("Expected " + expected + " but found " + t.v);
    return t.v.toUpperCase();
  };
  Cursor.prototype.isWord = function (w) {
    var t = this.peek();
    return t && t.t === "word" && t.v.toUpperCase() === w;
  };
  Cursor.prototype.op = function (v) {
    var t = this.next();
    if (!t || t.t !== "op" || t.v !== v) throw new Error("Expected " + v);
  };
  Cursor.prototype.str = function () {
    var t = this.next();
    if (!t || t.t !== "str") throw new Error("Expected text in quotes");
    return t.v;
  };
  Cursor.prototype.int = function () {
    var t = this.next();
    if (!t || t.t !== "num" || t.v < 1 || Math.floor(t.v) !== t.v) throw new Error("Expected a positive whole number");
    return t.v;
  };
  Cursor.prototype.end = function () {
    if (!this.done()) throw new Error("Unexpected text at the end of the line: " + this.peek().v);
  };

  // BASE.T / UPDATE.M
  function parseRef(cur, allowed) {
    var t = cur.next();
    if (!t || t.t !== "word" || t.v.indexOf(".") < 0) throw new Error("Expected a reference like BASE.T or UPDATE.M");
    var parts = t.v.split(".");
    var side = parts[0].toUpperCase(), col = parts[1].toUpperCase();
    if (side !== "BASE" && side !== "UPDATE") throw new Error("A reference must start with BASE or UPDATE: " + t.v);
    if (allowed && allowed.indexOf(side) < 0) throw new Error("Only " + allowed.join("/") + " is allowed here: " + t.v);
    return { side: side, col: col, idx: colToIndex(col) };
  }

  function parseValue(cur) {
    var t = cur.peek();
    if (!t) throw new Error("Missing value");
    if (t.t === "str") { cur.next(); return { kind: "text", v: t.v }; }
    if (t.t === "num") { cur.next(); return { kind: "num", v: t.v }; }
    if (t.t === "word") {
      var w = t.v.toUpperCase();
      if (w === "TODAY" || w === "NOW" || w === "EMPTY") { cur.next(); return { kind: w.toLowerCase() }; }
      if (w.indexOf(".") > 0) return { kind: "ref", ref: parseRef(cur) };
    }
    throw new Error("Invalid value: " + t.v);
  }

  function parseCondition(cur) {
    var left = parseRef(cur);
    var t = cur.next(), op;
    if (t && t.t === "op" && (t.v === "=" || t.v === "<>")) op = t.v;
    else if (t && t.t === "word" && t.v.toUpperCase() === "CONTAINS") op = "CONTAINS";
    else if (t && t.t === "word" && t.v.toUpperCase() === "NOT" && cur.isWord("CONTAINS")) { cur.next(); op = "NOTCONTAINS"; }
    else throw new Error("Invalid operator (allowed: =  <>  CONTAINS  NOT CONTAINS)");
    var right = parseValue(cur);
    return { left: left, op: op, right: right };
  }

  function parseAction(cur) {
    if (cur.isWord("COPY")) {
      cur.next();
      var from = parseRef(cur, ["UPDATE", "BASE"]);
      cur.op("->");
      var to = parseRef(cur, ["BASE"]);
      return { type: "copy", from: from, to: to };
    }
    if (cur.isWord("SET")) {
      cur.next();
      var target = parseRef(cur, ["BASE"]);
      cur.op("=");
      return { type: "set", to: target, value: parseValue(cur) };
    }
    throw new Error("Expected COPY or SET");
  }

  function newAlgorithm(name) {
    return {
      name: name, base: null, baseFirst: true, update: null,
      baseStartRow: 2, updateStartRow: 2, match: null,
      skipEmpty: false, dateFormat: "dd/mm/yyyy",
      actions: [], onChange: [], line: 0
    };
  }

  // fields the AI returns for a new CDF row (NEWROW X = AI.<FIELD>)
  var AI_FIELDS = {
    CDF: "cdf_number", DATE: "form_date", PRIORITY: "priority", DESCRIPTION: "description",
    OEM: "oem_category", DRAWING: "drawing_number", INSTALLATION: "installation",
    CODE: "code", INITIATOR: "initiator"
  };

  function parseColumnLetter(cur) {
    var t = cur.next();
    if (!t || t.t !== "word" || !/^[A-Za-z]{1,3}$/.test(t.v)) throw new Error("Expected a column letter (e.g. C or AA)");
    var col = t.v.toUpperCase();
    return { col: col, idx: colToIndex(col) };
  }

  function parseNewRowSource(cur) {
    var t = cur.peek();
    if (!t) throw new Error("Missing source for the field");
    if (t.t === "str") { cur.next(); return { kind: "text", v: t.v }; }
    if (t.t === "num") { cur.next(); return { kind: "num", v: t.v }; }
    var w = cur.word();
    if (w === "CDF") return { kind: "cdf" };
    if (w === "TODAY" || w === "NOW") return { kind: w.toLowerCase() };
    if (w === "LINK") {
      var which = cur.word();
      if (which !== "MAIN" && which !== "ATT") throw new Error("LINK must be followed by MAIN or ATT");
      var prefix = "";
      if (cur.peek() && cur.peek().t === "str") prefix = cur.str();
      return { kind: "link", which: which, prefix: prefix };
    }
    if (w.indexOf("AI.") === 0) {
      var f = w.slice(3);
      if (!AI_FIELDS[f]) throw new Error("Unknown AI field: " + f + ". Available: " + Object.keys(AI_FIELDS).join(", "));
      return { kind: "ai", field: f, key: AI_FIELDS[f] };
    }
    throw new Error("Unknown source: " + w + " (allowed: CDF, TODAY, NOW, \"text\", AI.<FIELD>, LINK MAIN/ATT)");
  }

  function parseSheetSpec(cur) {
    if (cur.isWord("FIRST")) { cur.next(); return { first: true, name: null }; }
    return { first: false, name: cur.str() };
  }

  function parseRules(text) {
    var result = {
      imports: [], algorithms: [], log: "UPDATE LOG", errors: [],
      newRow: { sheet: { first: true, name: null }, startRow: null, headerRow: 1, key: null, cells: [] },
      emailGroups: [], reports: [],
      cdfAlgorithm: null, claudeProject: null,
      cdfConfig: { stations: {}, suppliers: {}, finalApprovalSigner: null, pendingTables: [], duplicates: [] }
    };
    var lines = String(text || "").replace(/^﻿/, "").split(/\n/);
    var algo = null;

    function current(lineNo) {
      if (!algo) { algo = newAlgorithm("Algorithm"); algo.line = lineNo; result.algorithms.push(algo); }
      return algo;
    }

    lines.forEach(function (raw, i) {
      var lineNo = i + 1;
      try {
        var tokens = tokenize(raw);
        if (!tokens.length) return;
        var cur = new Cursor(tokens);
        var kw = cur.word();
        switch (kw) {
          case "IMPORT": {
            var slot = cur.int(); var target = cur.str(); var from = null, file = null;
            while (!cur.done()) {
              if (cur.isWord("FROM")) { cur.next(); from = cur.str(); }
              else if (cur.isWord("FILE")) { cur.next(); file = cur.str(); }
              else throw new Error("IMPORT accepts: FILE \"<file name>\" and FROM \"<sheet in the file>\"");
            }
            if (result.imports.some(function (x) { return x.slot === slot; })) throw new Error("IMPORT " + slot + " is defined twice");
            if (target.length > 31 || /[\\\/\?\*\[\]:]/.test(target)) throw new Error("Invalid Excel sheet name: " + target);
            if (file !== null && !fileKey(file)) throw new Error("FILE needs a file name (letters or digits)");
            result.imports.push({ slot: slot, target: target, from: from, file: file, line: lineNo });
            break;
          }
          case "ALGORITHM": {
            var nm = cur.str(); cur.end();
            algo = newAlgorithm(nm); algo.line = lineNo; result.algorithms.push(algo);
            break;
          }
          case "BASE": {
            var a = current(lineNo);
            if (cur.isWord("FIRST")) { cur.next(); a.baseFirst = true; a.base = null; }
            else { a.base = cur.str(); a.baseFirst = false; }
            cur.end(); break;
          }
          case "UPDATE": { var a2 = current(lineNo); a2.update = cur.str(); cur.end(); break; }
          case "BASE_START_ROW": { current(lineNo).baseStartRow = cur.int(); cur.end(); break; }
          case "UPDATE_START_ROW": { current(lineNo).updateStartRow = cur.int(); cur.end(); break; }
          case "MATCH": {
            var l = parseRef(cur); cur.op("="); var r = parseRef(cur); cur.end();
            if (l.side === r.side) throw new Error("MATCH needs one UPDATE side and one BASE side");
            current(lineNo).match = l.side === "UPDATE" ? { update: l, base: r } : { update: r, base: l };
            break;
          }
          case "SKIP_EMPTY": {
            var yn = cur.word().toUpperCase(); cur.end();
            if (yn !== "YES" && yn !== "NO") throw new Error("SKIP_EMPTY takes YES or NO");
            current(lineNo).skipEmpty = yn === "YES"; break;
          }
          case "DATE_FORMAT": { current(lineNo).dateFormat = cur.str(); cur.end(); break; }
          case "COPY": case "SET": {
            cur.p = 0; var act = parseAction(cur); cur.end();
            act.line = lineNo; act.conds = [];
            current(lineNo).actions.push(act); break;
          }
          case "IF": {
            var conds = [parseCondition(cur)];
            while (cur.isWord("AND")) { cur.next(); conds.push(parseCondition(cur)); }
            if (cur.isWord("THEN")) cur.next();
            var act2 = parseAction(cur); cur.end();
            act2.line = lineNo; act2.conds = conds;
            current(lineNo).actions.push(act2); break;
          }
          case "ON_CHANGE": {
            var act3 = parseAction(cur); cur.end();
            act3.line = lineNo; act3.conds = [];
            current(lineNo).onChange.push(act3); break;
          }
          case "LOG": {
            if (cur.isWord("NONE")) { cur.next(); result.log = null; }
            else result.log = cur.str();
            cur.end(); break;
          }
          // ----- new CDF row -----
          case "NEWROW_SHEET": { result.newRow.sheet = parseSheetSpec(cur); cur.end(); break; }
          case "NEWROW_START_ROW": { result.newRow.startRow = cur.int(); cur.end(); break; }
          case "NEWROW_HEADER_ROW": { result.newRow.headerRow = cur.int(); cur.end(); break; }
          case "NEWROW_KEY": { result.newRow.key = parseColumnLetter(cur); cur.end(); break; }
          case "NEWROW": {
            var c = parseColumnLetter(cur); cur.op("="); var src = parseNewRowSource(cur); cur.end();
            if (result.newRow.cells.some(function (x) { return x.idx === c.idx; })) throw new Error("Column " + c.col + " is already defined for the new row");
            result.newRow.cells.push({ col: c.col, idx: c.idx, source: src, line: lineNo });
            break;
          }
          case "AI_MODEL": { cur.str(); cur.end(); break; } // old setting, ignored
          case "CLAUDE_PROJECT": {
            var purl = cur.str().trim(); cur.end();
            if (!/^https:\/\/claude\.ai\/project\/[A-Za-z0-9-]+/.test(purl)) throw new Error("CLAUDE_PROJECT needs the project's address, e.g. \"https://claude.ai/project/...\"");
            result.claudeProject = purl; break;
          }
          // ----- CDF algorithm (code) and its settings -----
          case "CDF_ALGORITHM": { result.cdfAlgorithm = { version: cur.done() ? "" : cur.str(), line: lineNo }; cur.end(); break; }
          case "STATION": case "SUPPLIER": {
            var from1 = cur.str(); cur.op("="); var to1 = cur.str(); cur.end();
            result.cdfConfig[kw === "STATION" ? "stations" : "suppliers"][from1] = to1; break;
          }
          case "FINAL_APPROVAL_RELEASE": { result.cdfConfig.finalApprovalSigner = cur.str(); cur.end(); break; }
          case "PENDING_TABLE": { var tn = cur.str(), who = cur.str(); cur.end(); result.cdfConfig.pendingTables.push({ name: tn.toUpperCase(), person: who }); break; }
          case "DUPLICATE_CDF": { result.cdfConfig.duplicates.push(cur.str()); cur.end(); break; }
          // ----- email -----
          case "EMAIL_GROUP": {
            var gname = cur.str(), addrs = cur.str(); cur.end();
            var list = addrs.split(/[;,\s]+/).filter(Boolean);
            list.forEach(function (a) { if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(a)) throw new Error("Invalid email address: " + a); });
            result.emailGroups.push({ name: gname, addresses: list, line: lineNo });
            break;
          }
          case "REPORT": {
            var rep = { name: cur.str(), sheet: { first: true, name: null }, fromRow: null, conds: [], columns: [], line: lineNo };
            while (!cur.done()) {
              if (cur.isWord("SHEET")) { cur.next(); rep.sheet = parseSheetSpec(cur); }
              else if (cur.isWord("FROM_ROW")) { cur.next(); rep.fromRow = cur.int(); }
              else if (cur.isWord("WHERE") || cur.isWord("AND")) {
                cur.next();
                var cond = parseCondition(cur);
                if (cond.left.side !== "BASE" || (cond.right.kind === "ref" && cond.right.ref.side !== "BASE")) throw new Error("Reports can only use BASE.<column>");
                rep.conds.push(cond);
              }
              else if (cur.isWord("COLUMNS")) {
                cur.next();
                do {
                  if (cur.peek() && cur.peek().t === "op" && cur.peek().v === ",") cur.next();
                  rep.columns.push(parseColumnLetter(cur));
                } while (cur.peek() && cur.peek().t === "op" && cur.peek().v === ",");
              }
              else throw new Error("REPORT accepts: SHEET, FROM_ROW, WHERE, AND, COLUMNS");
            }
            if (!rep.columns.length) throw new Error("REPORT is missing COLUMNS");
            result.reports.push(rep);
            break;
          }
          default:
            throw new Error("Unknown command: " + kw);
        }
      } catch (e) {
        result.errors.push({ line: lineNo, message: e.message, text: raw.trim() });
      }
    });

    result.algorithms.forEach(function (a) {
      if (!a.update) result.errors.push({ line: a.line, message: "Algorithm \"" + a.name + "\" is missing UPDATE" });
      if (!a.match) result.errors.push({ line: a.line, message: "Algorithm \"" + a.name + "\" is missing MATCH" });
      if (!a.actions.length) result.errors.push({ line: a.line, message: "Algorithm \"" + a.name + "\" has no COPY/SET actions" });
    });
    if (result.cdfAlgorithm && result.algorithms.length) {
      result.errors.push({ line: result.cdfAlgorithm.line, message: "Use either CDF_ALGORITHM or ALGORITHM blocks, not both" });
    }
    var nr = result.newRow;
    if (nr.cells.length && !nr.key) {
      var cdfCell = nr.cells.filter(function (x) { return x.source.kind === "cdf"; })[0];
      if (cdfCell) nr.key = { col: cdfCell.col, idx: cdfCell.idx };
      else result.errors.push({ line: nr.cells[0].line, message: "The new row is missing NEWROW_KEY (or a column set to CDF)" });
    }
    if (nr.startRow == null) nr.startRow = result.algorithms.length ? result.algorithms[0].baseStartRow : 2;
    result.reports.forEach(function (r) {
      if (r.fromRow == null) r.fromRow = result.algorithms.length ? result.algorithms[0].baseStartRow : 2;
    });
    result.imports.sort(function (x, y) { return x.slot - y.slot; });
    return result;
  }

  // ---------- algorithm ----------
  /*
   * base / update: { values: 2D array starting at A1, formats: 2D array or null }
   * opts: { now: Date }
   * returns { changes:[{row, col, colLetter, key, oldValue, newValue, numberFormat, line, onChange}], stats }
   */
  function computeChanges(algo, base, update, opts) {
    opts = opts || {};
    var now = opts.now || new Date();
    var bv = base.values || [], uv = update.values || [];
    var bf = base.formats || null, uf = update.formats || null;
    var mU = algo.match.update.idx, mB = algo.match.base.idx;

    function cell(arr, r, c) { var row = arr[r]; return row && c < row.length ? row[c] : ""; }

    // index update rows by key (last occurrence wins)
    var updIndex = {}, order = [], dupKeys = {}, updateRows = 0;
    for (var r = algo.updateStartRow - 1; r < uv.length; r++) {
      var k = loose(cell(uv, r, mU));
      if (!k) continue;
      updateRows++;
      if (updIndex.hasOwnProperty(k)) dupKeys[k] = true; else order.push(k);
      updIndex[k] = r;
    }

    // index base rows by key (may repeat)
    var baseIndex = {};
    for (var b = algo.baseStartRow - 1; b < bv.length; b++) {
      var kb = loose(cell(bv, b, mB));
      if (!kb) continue;
      (baseIndex[kb] = baseIndex[kb] || []).push(b);
    }

    var changes = [], unmatched = [], matchedKeys = 0, matchedRows = 0, changedRows = 0;

    order.forEach(function (key) {
      var ur = updIndex[key];
      var rows = baseIndex[key];
      if (!rows) { unmatched.push({ key: String(cell(uv, ur, mU)).trim(), updateRow: ur + 1 }); return; }
      matchedKeys++;

      rows.forEach(function (br) {
        matchedRows++;
        var working = {};   // colIdx -> current value
        var rowChanges = {}; // colIdx -> change

        function get(ref) {
          if (ref.side === "UPDATE") return cell(uv, ur, ref.idx);
          return working.hasOwnProperty(ref.idx) ? working[ref.idx] : cell(bv, br, ref.idx);
        }
        function fmtOf(ref) {
          if (ref.side === "UPDATE") return uf ? cell(uf, ur, ref.idx) : null;
          return bf ? cell(bf, br, ref.idx) : null;
        }
        function evalValue(v) {
          switch (v.kind) {
            case "text": return { value: v.v, format: null };
            case "num": return { value: v.v, format: null };
            case "empty": return { value: "", format: null };
            case "today": return { value: dateToSerial(now, false), format: algo.dateFormat };
            case "now": return { value: dateToSerial(now, true), format: algo.dateFormat + " hh:mm" };
            case "ref": return { value: get(v.ref), format: fmtOf(v.ref) };
          }
        }
        function test(c) {
          var left = get(c.left);
          if (c.right.kind === "empty") {
            return c.op === "=" ? isEmpty(left) : c.op === "<>" ? !isEmpty(left) : false;
          }
          var rv = evalValue(c.right).value;
          var L = loose(left), R = loose(rv);
          if (c.op === "=") return L === R;
          if (c.op === "<>") return L !== R;
          if (c.op === "CONTAINS") return R !== "" && L.indexOf(R) >= 0;
          if (c.op === "NOTCONTAINS") return R === "" || L.indexOf(R) < 0;
          return false;
        }
        function apply(act, isOnChange) {
          for (var i = 0; i < act.conds.length; i++) if (!test(act.conds[i])) return false;
          var nv;
          if (act.type === "copy") nv = { value: get(act.from), format: fmtOf(act.from) };
          else nv = evalValue(act.value);
          if (act.type === "copy" && algo.skipEmpty && isEmpty(nv.value)) return false;
          var idx = act.to.idx;
          var currentVal = get(act.to);
          if (sameValue(currentVal, nv.value)) return false;
          var original = cell(bv, br, idx);
          working[idx] = nv.value;
          var fmt = nv.format && nv.format !== "General" ? nv.format : null;
          if (sameValue(original, nv.value)) { delete rowChanges[idx]; return false; }
          rowChanges[idx] = {
            row: br + 1, col: idx, colLetter: indexToCol(idx), key: String(cell(bv, br, mB)).trim(),
            oldValue: original, newValue: nv.value, numberFormat: fmt,
            oldFormat: bf ? cell(bf, br, idx) : null,
            line: act.line, onChange: !!isOnChange
          };
          return true;
        }

        algo.actions.forEach(function (a) { apply(a, false); });
        if (Object.keys(rowChanges).length) {
          algo.onChange.forEach(function (a) { apply(a, true); });
        }
        var list = Object.keys(rowChanges).map(function (k) { return rowChanges[k]; });
        if (list.some(function (c) { return !c.onChange; })) {
          changedRows++;
          list.sort(function (x, y) { return (x.onChange - y.onChange) || (x.col - y.col); });
          Array.prototype.push.apply(changes, list);
        }
      });
    });

    return {
      changes: changes,
      stats: {
        updateRows: updateRows,
        uniqueUpdateKeys: order.length,
        matchedKeys: matchedKeys,
        matchedRows: matchedRows,
        changedRows: changedRows,
        changedCells: changes.length,
        unmatched: unmatched,
        duplicateUpdateKeys: Object.keys(dupKeys)
      }
    };
  }

  // ---------- import file names ----------
  // "CDF_List_Export (3).xlsx" -> "cdflistexport3" ; compare ignoring case, spaces and separators
  function fileKey(name) {
    return String(name || "").replace(/\.(xlsx|xlsm|xls|csv)$/i, "").toLowerCase().replace(/[^a-z0-9\u0590-\u05ff]/g, "");
  }
  function importFileMatches(pattern, fileName) {
    var k = fileKey(pattern);
    return !!k && fileKey(fileName).indexOf(k) >= 0;
  }
  // slot whose FILE name appears in the file name; the longest (most specific) name wins
  function matchImportSlot(imports, fileName) {
    var best = null;
    (imports || []).forEach(function (s) {
      if (s.file && importFileMatches(s.file, fileName) && (!best || fileKey(s.file).length > fileKey(best.file).length)) best = s;
    });
    return best ? best.slot : null;
  }

  // ---------- new CDF row ----------
  var LOG_NOTES = {
    auto: "Automatic (ON_CHANGE)",
    notFound: "Not found in base sheet",
    none: "No changes",
    newRow: "New row"
  };
  // notes written by the first (Hebrew) version, still recognised when reading an old log
  var LEGACY_NOTES = { notFound: "\u05DC\u05D0 \u05E0\u05DE\u05E6\u05D0 \u05D1\u05D2\u05D9\u05DC\u05D9\u05D5\u05DF \u05D4\u05D1\u05E1\u05D9\u05E1", newRow: "\u05E9\u05D5\u05E8\u05D4 \u05D7\u05D3\u05E9\u05D4" };
  function isNote(v, kind) { return v === LOG_NOTES[kind] || v === LEGACY_NOTES[kind]; }

  // "A330-1-961.-.pdf" -> A330-1-961 ; "A330-1-682.A-ATT.pdf" -> A330-1-682.A (isAtt)
  function parseCdfFileName(name) {
    var base = String(name || "").replace(/\.pdf$/i, "").trim();
    var isAtt = /ATT/i.test(base);
    var core = base.replace(/[-_ ]?ATT\b.*$/i, "").replace(/[-_ ]?ATT$/i, "");
    var m = /([A-Z]\d{3}-\d+-\d+)(?:\.([A-Z]{1,2}|-)?)?$/i.exec(core) || /([A-Z]\d{3}-\d+-\d+)(?:\.([A-Z](?![A-Z])|-))?/i.exec(core);
    if (!m) return { cdf: null, isAtt: isAtt };
    var rev = m[2] && m[2] !== "-" ? m[2].toUpperCase() : null;
    return { cdf: m[1].toUpperCase() + (rev ? "." + rev : ""), isAtt: isAtt };
  }

  function normalizeCdf(s) {
    return String(s || "").trim().toUpperCase().replace(/\.-?$/, "");
  }

  function serialToDateString(serial) {
    var d = new Date(Date.UTC(1899, 11, 30) + Math.round(serial * 86400000));
    var p = function (n) { return (n < 10 ? "0" : "") + n; };
    return p(d.getUTCDate()) + "/" + p(d.getUTCMonth() + 1) + "/" + d.getUTCFullYear();
  }

  function parseDateString(s) {
    var m = /^\s*(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{4})\s*$/.exec(String(s || ""));
    if (!m) return null;
    var d = +m[1], mo = +m[2], y = +m[3];
    if (mo < 1 || mo > 12 || d < 1 || d > 31) return null;
    return (Date.UTC(y, mo - 1, d) - Date.UTC(1899, 11, 30)) / 86400000;
  }

  /*
   * opts: { ai: {cdf_number, ...}, mainName, attName, now, dateFormat }
   * returns { cdf, cells:[{col, idx, label, value, display, format, link}], warnings:[] }
   */
  function buildNewRowCells(newRow, opts) {
    var ai = opts.ai || {}, warnings = [];
    var fromFile = parseCdfFileName(opts.mainName).cdf;
    var fromAi = ai.cdf_number ? normalizeCdf(ai.cdf_number) : null;
    var cdf = fromFile || fromAi || "";
    if (fromFile && fromAi && normalizeCdf(fromFile) !== fromAi) {
      warnings.push("The CDF number in the file name (" + fromFile + ") differs from the one on the form (" + fromAi + ")");
    }
    if (!cdf) warnings.push("No CDF number found — fill it in manually");
    if (opts.attName) {
      var attCdf = parseCdfFileName(opts.attName).cdf;
      if (attCdf && cdf && normalizeCdf(attCdf) !== normalizeCdf(cdf)) warnings.push("The ATT file (" + attCdf + ") does not belong to " + cdf);
    }
    var fmt = opts.dateFormat || "dd/mm/yyyy";
    var cells = newRow.cells.map(function (c) {
      var s = c.source, out = { col: c.col, idx: c.idx, kind: s.kind, value: "", display: "", format: null, link: null, label: "" };
      switch (s.kind) {
        case "cdf": out.value = cdf; out.label = "CDF"; break;
        case "text": out.value = s.v; out.label = "Fixed"; break;
        case "num": out.value = s.v; out.label = "Fixed"; break;
        case "today": out.value = dateToSerial(opts.now || new Date(), false); out.format = fmt; out.label = "Today's date"; break;
        case "now": out.value = dateToSerial(opts.now || new Date(), true); out.format = fmt + " hh:mm"; out.label = "Now"; break;
        case "ai": out.value = ai[s.key] == null ? "" : String(ai[s.key]).trim(); out.label = "AI · " + s.field; break;
        case "link":
          var fname = s.which === "ATT" ? opts.attName : opts.mainName;
          out.label = s.which === "ATT" ? "ATT link" : "CDF link";
          out.prefix = s.prefix;
          if (fname) { out.value = fname; out.link = { address: s.prefix + fname, textToDisplay: fname }; }
          break;
      }
      out.display = out.format ? serialToDateString(out.value) : String(out.value);
      return out;
    });
    return { cdf: cdf, cells: cells, warnings: warnings };
  }

  // keyValues: values of the key column, index 0 = row 1
  function findNewRowTarget(keyValues, startRow, cdf) {
    var want = loose(normalizeCdf(cdf)), last = startRow - 2;
    for (var i = startRow - 1; i < keyValues.length; i++) {
      var v = keyValues[i];
      if (isEmpty(v)) continue;
      last = i;
      if (want && loose(normalizeCdf(v)) === want) return { row: i + 1, existing: true };
    }
    return { row: Math.max(last + 2, startRow), existing: false };
  }

  // Targets for several CDFs written together: each new CDF takes the next free row,
  // and a CDF that appears twice in the batch goes to the same row (dupOf = index of the first).
  function assignNewRowTargets(keyValues, startRow, cdfs) {
    var keys = (keyValues || []).slice(), seen = {};
    return cdfs.map(function (cdf, i) {
      var id = loose(normalizeCdf(cdf));
      var t = findNewRowTarget(keys, startRow, cdf);
      if (id && seen[id] != null) t.dupOf = seen[id];
      if (id && seen[id] == null) seen[id] = i;
      if (!t.existing) {
        while (keys.length < t.row) keys.push("");
        keys[t.row - 1] = cdf || "(new)";
      }
      return t;
    });
  }

  // ---------- answer pasted from claude.ai ----------
  var PASTE_KEYS = ["main_file", "att_file", "cdf_number", "form_date", "priority", "description",
    "oem_category", "drawing_number", "installation", "code", "initiator"];
  var PASTE_ALIASES = {
    cdf: "cdf_number", cdf_num: "cdf_number", date: "form_date", oem: "oem_category",
    drawing: "drawing_number", main: "main_file", att: "att_file", attachment: "att_file", owner: "initiator",
    version: "prompt_version", promptversion: "prompt_version"
  };

  // returns { items: [{main_file, att_file, cdf_number, ..., flags:[]}], error }
  function parseAiPaste(text) {
    var s = String(text || "").replace(/^﻿/, "").trim();
    if (!s) return { items: [], error: "Nothing was pasted" };
    var fence = /```(?:json|JSON)?\s*([\s\S]*?)```/.exec(s);
    if (fence) s = fence[1].trim();
    var a = s.indexOf("["), o = s.indexOf("{");
    var start = a < 0 ? o : o < 0 ? a : Math.min(a, o);
    if (start < 0) return { items: [], error: "No JSON found. Copy the code block from Claude's reply (the Copy button in the corner of the block)." };
    var close = s.charAt(start) === "[" ? "]" : "}";
    var end = s.lastIndexOf(close);
    if (end < start) return { items: [], error: "The JSON is cut off — make sure you copied the whole block" };
    var body = s.slice(start, end + 1), data = null;
    var attempts = [
      body,
      body.replace(/[“”„‟″]/g, '"').replace(/[‘’]/g, "'"),
    ];
    attempts.push(attempts[1].replace(/,\s*([\]}])/g, "$1"));
    for (var i = 0; i < attempts.length && data === null; i++) {
      try { data = JSON.parse(attempts[i]); } catch (e) { data = null; }
    }
    if (data === null) return { items: [], error: "The JSON is not valid. Ask Claude: \"Return only the JSON block again\"." };
    var list = Array.isArray(data) ? data : (data && Array.isArray(data.rows) ? data.rows : [data]);
    var items = [];
    list.forEach(function (raw) {
      if (!raw || typeof raw !== "object") return;
      var it = { flags: [] };
      Object.keys(raw).forEach(function (k) {
        var key = k.trim().toLowerCase().replace(/[\s-]+/g, "_");
        key = PASTE_ALIASES[key] || key;
        if (key === "flags") {
          it.flags = Array.isArray(raw[k]) ? raw[k].map(String).filter(Boolean) : (raw[k] ? [String(raw[k])] : []);
        } else if (PASTE_KEYS.indexOf(key) >= 0 || key === "prompt_version") {
          it[key] = raw[k] == null ? "" : String(raw[k]).trim();
        }
      });
      if (it.prompt_version == null) it.prompt_version = "";
      PASTE_KEYS.forEach(function (k) { if (it[k] == null) it[k] = ""; });
      if (it.cdf_number || it.main_file || it.description) items.push(it);
    });
    if (!items.length) return { items: [], error: "No CDF rows found in the reply" };
    return { items: items, error: null };
  }

  // ---------- reports ----------
  function runReport(report, values) {
    function get(row, idx) { var r = values[row]; return r && idx < r.length ? r[idx] : ""; }
    function test(c, r) {
      var left = get(r, c.left.idx);
      if (c.right.kind === "empty") return c.op === "=" ? isEmpty(left) : c.op === "<>" ? !isEmpty(left) : false;
      var rv = c.right.kind === "ref" ? get(r, c.right.ref.idx) : c.right.v;
      var L = loose(left), R = loose(rv);
      if (c.op === "=") return L === R;
      if (c.op === "<>") return L !== R;
      if (c.op === "CONTAINS") return R !== "" && L.indexOf(R) >= 0;
      if (c.op === "NOTCONTAINS") return R === "" || L.indexOf(R) < 0;
      return false;
    }
    var rows = [];
    for (var r = report.fromRow - 1; r < values.length; r++) {
      var row = values[r] || [];
      if (!row.some(function (v) { return !isEmpty(v); })) continue;
      if (report.conds.every(function (c) { return test(c, r); })) {
        rows.push({ row: r + 1, cells: report.columns.map(function (c) { return get(r, c.idx); }) });
      }
    }
    return { rows: rows };
  }

  // ---------- reading the log back ----------
  // log columns: stamp | algorithm | sheet | row | key | column | old | new | note
  function lastRunFromLog(values, todayStr) {
    var rows = (values || []).slice(1).filter(function (r) { return r && !isEmpty(r[0]); });
    var runRows = rows.filter(function (r) { return !isNote(r[8], "newRow"); });
    var stamp = runRows.length ? String(runRows[runRows.length - 1][0]) : null;
    var last = runRows.filter(function (r) { return String(r[0]) === stamp; });
    function entry(r) { return { stamp: String(r[0]), sheet: r[2], row: r[3], key: String(r[4]), col: String(r[5]), oldValue: r[6], newValue: r[7], note: r[8] }; }
    return {
      stamp: stamp,
      changes: last.filter(function (r) { return !r[8]; }).map(entry),
      notFound: last.filter(function (r) { return isNote(r[8], "notFound"); }).map(entry),
      newRows: rows.filter(function (r) { return isNote(r[8], "newRow") && (!todayStr || String(r[0]).indexOf(todayStr) === 0); }).map(entry)
    };
  }

  // ---------- email ----------
  function htmlEsc(s) {
    return String(s == null ? "" : s).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; });
  }

  // doc: { title, intro, sections:[{heading, headers:[], rows:[[...]]}] }
  function buildEmail(doc) {
    var text = [doc.title, ""];
    if (doc.intro) text.push(doc.intro, "");
    var cell = "padding:4px 8px;border:1px solid #c8d0db;text-align:left;vertical-align:top;";
    var html = '<div style="font-family:Calibri,Arial,sans-serif;font-size:11pt;direction:ltr;text-align:left">' +
      '<p style="margin:0 0 8px"><b>' + htmlEsc(doc.title) + "</b></p>" +
      (doc.intro ? '<p style="margin:0 0 8px">' + htmlEsc(doc.intro) + "</p>" : "");
    doc.sections.forEach(function (s) {
      if (s.heading) text.push(s.heading);
      text.push(s.headers.join(" | "));
      s.rows.forEach(function (r) { text.push(r.map(function (v) { return String(v == null ? "" : v); }).join(" | ")); });
      text.push("");
      html += (s.heading ? '<p style="margin:12px 0 4px"><b>' + htmlEsc(s.heading) + "</b></p>" : "") +
        '<table style="border-collapse:collapse;font-size:10pt">' +
        "<tr>" + s.headers.map(function (h) { return '<th style="' + cell + 'background:#13294B;color:#ffffff">' + htmlEsc(h) + "</th>"; }).join("") + "</tr>" +
        s.rows.map(function (r) { return "<tr>" + r.map(function (v) { return '<td style="' + cell + '">' + htmlEsc(v) + "</td>"; }).join("") + "</tr>"; }).join("") +
        "</table>";
    });
    html += "</div>";
    return { subject: doc.title, text: text.join("\r\n").replace(/\s+$/, ""), html: html };
  }

  function mailtoUrl(to, cc, subject, body) {
    var q = [];
    if (cc && cc.length) q.push("cc=" + encodeURIComponent(cc.join(";")));
    q.push("subject=" + encodeURIComponent(subject || ""));
    if (body) q.push("body=" + encodeURIComponent(body));
    return "mailto:" + to.map(encodeURIComponent).join(";").replace(/%40/g, "@") + "?" + q.join("&");
  }

  function cleanHeader(v) { return isEmpty(v) ? "" : String(v).replace(/\s+/g, " ").trim(); }

  // readable version of a rule for the UI
  function describeAction(a) {
    function iso(s) { return s; }
    function ref(r) { return iso(r.side + "." + r.col); }
    function val(v) {
      if (v.kind === "text") return iso('"' + v.v + '"');
      if (v.kind === "num") return iso(String(v.v));
      if (v.kind === "ref") return ref(v.ref);
      return iso(v.kind.toUpperCase());
    }
    var cond = a.conds && a.conds.length
      ? "If " + a.conds.map(function (c) {
          var op = { "=": "=", "<>": "≠", "CONTAINS": "contains", "NOTCONTAINS": "does not contain" }[c.op];
          return ref(c.left) + " " + op + " " + val(c.right);
        }).join(" and ") + " — "
      : "";
    if (a.type === "copy") return cond + "Copy " + ref(a.from) + " → " + ref(a.to);
    return cond + "Set " + ref(a.to) + " = " + val(a.value);
  }

  return {
    AI_FIELDS: AI_FIELDS,
    LOG_NOTES: LOG_NOTES,
    parseCdfFileName: parseCdfFileName,
    importFileMatches: importFileMatches,
    matchImportSlot: matchImportSlot,
    parseAiPaste: parseAiPaste,
    assignNewRowTargets: assignNewRowTargets,
    normalizeCdf: normalizeCdf,
    buildNewRowCells: buildNewRowCells,
    findNewRowTarget: findNewRowTarget,
    runReport: runReport,
    lastRunFromLog: lastRunFromLog,
    buildEmail: buildEmail,
    mailtoUrl: mailtoUrl,
    cleanHeader: cleanHeader,
    serialToDateString: serialToDateString,
    parseDateString: parseDateString,
    parseRules: parseRules,
    computeChanges: computeChanges,
    describeAction: describeAction,
    colToIndex: colToIndex,
    indexToCol: indexToCol,
    dateToSerial: dateToSerial,
    isEmpty: isEmpty,
    loose: loose
  };
});
