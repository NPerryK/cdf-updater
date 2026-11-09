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
    CODE: "code", DWG_REV: "drawing_rev", AUTHOR: "author", ADVISER: "technical_adviser",
    INITIATOR: "author"   // old name of AUTHOR
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
    if (w === "CDF") {
      if (cur.isWord("FULL")) { cur.next(); return { kind: "cdf_full" }; }   // as printed: A330-1-973.- / A330-1-682.A
      if (cur.isWord("REV")) { cur.next(); return { kind: "cdf_rev" }; }     // revision letter: - / A
      return { kind: "cdf" };                                                 // without ".-": A330-1-973
    }
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
    throw new Error("Unknown source: " + w + " (allowed: CDF, CDF FULL, CDF REV, TODAY, NOW, \"text\", AI.<FIELD>, LINK MAIN/ATT)");
  }

  function parseSheetSpec(cur) {
    if (cur.isWord("FIRST")) { cur.next(); return { first: true, name: null }; }
    return { first: false, name: cur.str() };
  }

  function parseRules(text) {
    var result = {
      imports: [], algorithms: [], log: "UPDATE LOG", errors: [],
      newRow: { sheet: { first: true, name: null }, startRow: null, headerRow: 1, key: null, match: null, cells: [] },
      emailGroups: [], reports: [],
      cdfAlgorithm: null, claudeProject: null,
      cdfConfig: { stations: {}, suppliers: {}, finalApprovalSigner: null, pendingTables: [], duplicates: [], personGroups: [], personAliases: [] }
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
          case "NEWROW_MATCH": { result.newRow.match = parseColumnLetter(cur); cur.end(); break; }
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
          // 120: people → group, for column V (PERSON_GROUP "group" "name; name; …", PERSON_ALIAS "as written" = "as in the group")
          case "PERSON_GROUP": {
            var pg = cur.str().trim(), names = cur.str(); cur.end();
            if (!pg) throw new Error("PERSON_GROUP needs a group name");
            var plist = names.split(/[;\n]+/).map(function (x) { return x.trim(); }).filter(Boolean);
            if (!plist.length) throw new Error("PERSON_GROUP \"" + pg + "\" has no names — separate them with ;");
            var have = result.cdfConfig.personGroups.filter(function (g) { return g.group === pg; })[0];
            if (have) Array.prototype.push.apply(have.people, plist); else result.cdfConfig.personGroups.push({ group: pg, people: plist });
            break;
          }
          case "PERSON_ALIAS": { var af = cur.str(); cur.op("="); var at = cur.str(); cur.end(); result.cdfConfig.personAliases.push({ from: af, to: at }); break; }
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
    if (/\d\.$/.test(core)) core += "-";   // "A330-1-961.-ATT": the "-" before ATT is the revision
    var m = /([A-Z]\d{3}-\d+-\d+)(?:\.([A-Z]{1,2}|-)?)?$/i.exec(core) || /([A-Z]\d{3}-\d+-\d+)(?:\.([A-Z](?![A-Z])|-))?/i.exec(core);
    if (!m) return { cdf: null, printed: null, isAtt: isAtt };
    var rev = m[2] && m[2] !== "-" ? m[2].toUpperCase() : null;
    return { cdf: m[1].toUpperCase() + (rev ? "." + rev : ""), printed: m[1].toUpperCase() + (m[2] ? "." + m[2].toUpperCase() : ""), isAtt: isAtt };
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
  // "A330-1-973.-" -> "-" ; "A330-1-682.A" -> "A"
  function cdfRevision(printed) {
    var m = /\.([A-Z]{1,2})$/.exec(normalizeCdf(printed));
    return m ? m[1] : "-";
  }

  function buildNewRowCells(newRow, opts) {
    var ai = opts.ai || {}, warnings = [];
    var file = parseCdfFileName(opts.mainName), fromFile = file.cdf;
    var fromAi = ai.cdf_number ? normalizeCdf(ai.cdf_number) : null;
    var cdf = fromFile || fromAi || "";
    if (fromFile && fromAi && normalizeCdf(fromFile) !== fromAi) {
      warnings.push("The CDF number in the file name (" + fromFile + ") differs from the one on the form (" + fromAi + ")");
    }
    if (!cdf) warnings.push("No CDF number found — fill it in manually");
    var printedAi = ai.cdf_number ? String(ai.cdf_number).trim().toUpperCase().replace(/\s+/g, "") : "";
    var cdfFull = printedAi && /\.([A-Z]{1,2}|-)$/.test(printedAi) ? printedAi
      : file.printed && /\.([A-Z]{1,2}|-)$/.test(file.printed) ? file.printed
      : cdf ? (/\.[A-Z]{1,2}$/.test(cdf) ? cdf : cdf + ".-") : "";
    if (opts.attName) {
      var attCdf = parseCdfFileName(opts.attName).cdf;
      if (attCdf && cdf && normalizeCdf(attCdf) !== normalizeCdf(cdf)) warnings.push("The ATT file (" + attCdf + ") does not belong to " + cdf);
    }
    var fmt = opts.dateFormat || "dd/mm/yyyy";
    var cells = newRow.cells.map(function (c) {
      var s = c.source, out = { col: c.col, idx: c.idx, kind: s.kind, value: "", display: "", format: null, link: null, label: "" };
      switch (s.kind) {
        case "cdf": out.value = cdf; out.label = "CDF"; break;
        case "cdf_full": out.value = cdfFull; out.label = "CDF as printed"; break;
        case "cdf_rev": out.value = cdfFull ? cdfRevision(cdfFull) : ""; out.label = "CDF revision"; break;
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
    return { cdf: cdf, cdfFull: cdfFull, cells: cells, warnings: warnings };
  }

  // Where a new CDF goes. keyValues / matchValues: the key column (C) and the column of the
  // prepared rows (B), index 0 = row 1. Returns {row, existing, prepared, insert, after}:
  //   existing - the CDF is already in the key column: that row is updated
  //   prepared - an empty row prepared for this CDF (B already holds its number and PDF links)
  //   insert   - a revision (.A, .B): a new row inserted right after the original CDF's block
  //   otherwise - a new row after the last used row
  function findNewRowTarget(keyValues, startRow, cdf, matchValues) {
    var want = loose(normalizeCdf(cdf)), match = matchValues || [], i;
    var last = startRow - 2;
    for (i = startRow - 1; i < keyValues.length; i++) {
      var v = keyValues[i];
      if (isEmpty(v)) continue;
      last = i;
      if (want && loose(normalizeCdf(v)) === want) return { row: i + 1, existing: true };
    }
    if (want) {
      for (i = startRow - 1; i < match.length; i++) {
        if (!isEmpty(match[i]) && isEmpty(keyValues[i]) && loose(normalizeCdf(match[i])) === want) return { row: i + 1, existing: false, prepared: true };
      }
    }
    var rev = /^(.*)\.([A-Z]{1,2})$/.exec(normalizeCdf(cdf));
    if (rev) {
      var base = loose(rev[1]), anchor = -1;
      for (i = startRow - 1; i < keyValues.length; i++) {
        if (!isEmpty(keyValues[i]) && loose(normalizeCdf(keyValues[i]).replace(/\.[A-Z]{1,2}$/, "")) === base) anchor = i;
      }
      if (anchor >= 0) {
        var n = Math.max(keyValues.length, match.length), j = anchor + 1;
        while (j < n && isEmpty(keyValues[j]) && isEmpty(match[j])) j++;
        if (j < n) return { row: j + 1, existing: false, insert: true, after: String(keyValues[anchor]).trim() };
        return { row: n + 1, existing: false, after: String(keyValues[anchor]).trim() };
      }
    }
    for (i = startRow - 1; i < match.length; i++) if (!isEmpty(match[i]) && i > last) last = i;
    return { row: Math.max(last + 2, startRow), existing: false };
  }

  // Several CDFs written together, in this order. Each target is found on the sheet as it will be
  // after the earlier ones were written (a new CDF takes its row; an inserted row pushes the rows
  // below it down). writeRow = where to write when the earlier ones are done; row = final row.
  // A CDF that appears twice goes to the same row (dupOf = index of the first).
  // sheet: { keys, match, rows, formats } - rows/formats (whole rows) are optional
  function planNewRowTargets(sheet, startRow, cdfs) {
    var keys = (sheet.keys || []).slice(), match = (sheet.match || []).slice();
    var rows = sheet.rows ? sheet.rows.slice() : null, fmts = sheet.formats ? sheet.formats.slice() : null;
    var out = [], seen = {};
    cdfs.forEach(function (cdf, i) {
      var id = loose(normalizeCdf(cdf));
      var t = findNewRowTarget(keys, startRow, cdf, match);
      if (id && seen[id] != null) t.dupOf = seen[id];
      if (id && seen[id] == null) seen[id] = i;
      t.writeRow = t.row;
      t.existingValues = t.insert || !rows ? [] : (rows[t.row - 1] || []).slice();
      t.existingFormats = t.insert || !fmts ? [] : (fmts[t.row - 1] || []).slice();
      if (t.insert) {
        out.forEach(function (o) { if (o.row >= t.row) o.row++; });
        keys.splice(t.row - 1, 0, ""); match.splice(t.row - 1, 0, "");
        if (rows) rows.splice(t.row - 1, 0, []); if (fmts) fmts.splice(t.row - 1, 0, []);
      }
      if (!t.existing) {
        while (keys.length < t.row) keys.push("");
        keys[t.row - 1] = cdf || "(new)";
      }
      out.push(t);
    });
    return out;
  }

  function assignNewRowTargets(keyValues, startRow, cdfs, matchValues) {
    return planNewRowTargets({ keys: keyValues, match: matchValues }, startRow, cdfs);
  }

  // ---------- answer pasted from claude.ai ----------
  var PASTE_KEYS = ["main_file", "att_file", "cdf_number", "form_date", "priority", "description",
    "oem_category", "drawing_number", "installation", "code", "drawing_rev", "author", "technical_adviser"];
  var PASTE_ALIASES = {
    cdf: "cdf_number", cdf_num: "cdf_number", date: "form_date", oem: "oem_category",
    drawing: "drawing_number", main: "main_file", att: "att_file", attachment: "att_file", owner: "author",
    initiator: "author", adviser: "technical_adviser", technical_advisor: "technical_adviser",
    dwg_rev: "drawing_rev", drawing_revision: "drawing_rev",
    // the keys of the older "CDF New Rows" project format
    nrc: "code", jc: "drawing_number", dwgver: "drawing_rev", dwg_ver: "drawing_rev", title: "installation",
    summary: "description", open: "form_date", open_date: "form_date", advisor: "technical_adviser",
    notes: "flags", note: "flags",
    version: "prompt_version", promptversion: "prompt_version"
  };

  // Rows given as lists of values, without field names — from any AI (Claude, ChatGPT…) or copied from Excel.
  // The order is told by where the CDF number stands and by how many values a row has:
  var PASTE_ORDERS = [
    // CDF first, then the two files: what ChatGPT gives when asked for "rows"
    { cdf: 0, n: 13, keys: ["cdf_number", "main_file", "att_file", "form_date", "priority", "description", "oem_category", "drawing_number", "installation", "code", "drawing_rev", "author", "technical_adviser"] },
    // the add-in's own field order
    { cdf: 2, n: 13, keys: PASTE_KEYS.slice() },
    // without the files
    { cdf: 0, n: 11, keys: ["cdf_number", "form_date", "priority", "description", "oem_category", "drawing_number", "installation", "code", "drawing_rev", "author", "technical_adviser"] },
    // the CDF sheet's own columns A–O (CDF, link file, link, revision, notes, date, priority, …)
    { cdf: 0, n: 15, sheet: true, keys: ["cdf_number", "att_file", "main_file", "cdf_rev", "", "form_date", "priority", "description", "oem_category", "drawing_number", "installation", "code", "drawing_rev", "author", "technical_adviser"] }
  ];
  // header names (the add-in's, older ones, or the CDF sheet's own headers) → the add-in's fields
  var HEADER_ALIASES = {
    cdf_num_link_to_dev: "cdf_number", cdf_number: "cdf_number", link_to_gcc: "main_file", link_to_gcc_file: "att_file",
    cdf_revision: "cdf_rev", received_date: "form_date", problem_description: "description", tags: "oem_category",
    dwg: "drawing_number", "dwg._name": "installation", dwg_name: "installation", cdf_dwg_rev: "drawing_rev",
    technical_adviser: "technical_adviser"
  };
  var CDF_RX = /^A\d{3}-\d+-\d+/i;
  function fieldOf(name) {
    var k = String(name == null ? "" : name).trim().toLowerCase().replace(/[\s\-()\/]+/g, "_").replace(/_+/g, "_").replace(/^_|_$/g, "");
    k = HEADER_ALIASES[k] || PASTE_ALIASES[k] || k;
    return PASTE_KEYS.indexOf(k) >= 0 || k === "cdf_rev" ? k : null;
  }
  // lists of values → objects with field names (or null when the order can't be told)
  function rowsToObjects(rows) {
    rows = rows.filter(function (r) { return Array.isArray(r) && r.some(function (v) { return v != null && String(v).trim() !== ""; }); });
    if (!rows.length) return null;
    var head = rows[0].map(fieldOf), named = head.filter(Boolean).length;
    if (named >= 3 && named >= rows[0].length / 2 && !rows[0].some(function (v) { return CDF_RX.test(String(v).trim()); })) {
      return { order: "the header row", list: rows.slice(1).map(function (r) {
        var o = {}; head.forEach(function (k, i) { if (k && o[k] == null) o[k] = r[i]; }); return o; }) };
    }
    var at = {};
    rows.forEach(function (r) { r.forEach(function (v, i) { if (CDF_RX.test(String(v == null ? "" : v).trim())) at[i] = (at[i] || 0) + 1; }); });
    var idx = Object.keys(at).sort(function (a, b) { return at[b] - at[a]; })[0];
    if (idx == null) return null;
    idx = +idx;
    var n = rows.reduce(function (m, r) { return Math.max(m, r.length); }, 0);
    // exactly that many values, and the date column holds dates: another format with the CDF first isn't misread
    var DATE_RX = /^(\d{1,2}[\/.\-]\d{1,2}[\/.\-]\d{2,4}|\d{4}-\d{2}-\d{2}|\d{1,2} [A-Za-z]{3} \d{4})?$/;
    var ord = PASTE_ORDERS.filter(function (o) {
      if (o.cdf !== idx || o.n !== n) return false;
      var d = o.keys.indexOf("form_date");
      return rows.every(function (r) { return DATE_RX.test(String(r[d] == null ? "" : r[d]).trim()); });
    })[0];
    if (!ord) return null;
    return { order: ord.keys.filter(Boolean).join(", "), list: rows.map(function (r) {
      var o = {};
      ord.keys.forEach(function (k, i) { if (k) o[k] = r[i] == null ? "" : r[i]; });
      // the sheet's columns carry the revision apart: put it back on the number (A330-1-973 + A → A330-1-973.A)
      if (o.cdf_rev != null) { var rv = String(o.cdf_rev).trim(); if (rv && !/\.[^.]*$/.test(String(o.cdf_number).trim())) o.cdf_number = String(o.cdf_number).trim() + "." + rv; delete o.cdf_rev; }
      return o;
    }) };
  }
  function splitTable(s) {
    var lines = s.split(/\r?\n/).filter(function (l) { return l.trim(); });
    var sep = lines.some(function (l) { return l.split("\t").length >= 4; }) ? "\t" : null;
    if (!sep) return null;
    return lines.map(function (l) { return l.split(sep).map(function (v) { return v.trim().replace(/^"(.*)"$/, "$1"); }); });
  }

  // A table row (tab-separated) instead of JSON: Claude answered in another format,
  // usually because it did not get the add-in's instructions.
  function looksLikeTable(s) {
    if (/\{\s*"/.test(s)) return false;
    return s.split(/\r?\n/).some(function (line) { return line.split("\t").length >= 4; });
  }

  // The first JSON value in a reply, read leniently: smart quotes, a comma before ] or }, raw line breaks,
  // // comment lines, Python-style 'quotes' / None / True. → { data } or { kind: nojson | cut | invalid, error }
  function looseJson(text) {
    var s = String(text || "");
    var a = s.indexOf("["), o = s.indexOf("{");
    var start = a < 0 ? o : o < 0 ? a : Math.min(a, o);
    if (start < 0) return { kind: "nojson", error: "No JSON found. Copy the code block from the AI's reply (the Copy button in the corner of the block)." };
    var close = s.charAt(start) === "[" ? "]" : "}";
    var end = s.lastIndexOf(close);
    if (end < start) return { kind: "cut", error: "The JSON is cut off — make sure you copied the whole block" };
    var body = s.slice(start, end + 1), data = null;
    var attempts = [
      body,
      body.replace(/[“”„‟″]/g, '"').replace(/[‘’]/g, "'"),
    ];
    attempts.push(attempts[1].replace(/,\s*([\]}])/g, "$1"));                       // a comma before ] or }
    attempts.push(attempts[2].replace(/[\t\r\n]+/g, " "));                           // raw tabs / line breaks inside strings
    attempts.push(attempts[2].replace(/^\s*\/\/.*$/gm, "").replace(/[\t\r\n]+/g, " "));   // // comment lines
    if (attempts[1].indexOf('"') < 0)                                                 // Python-style: 'single quotes', None, True
      attempts.push(attempts[2].replace(/'((?:[^'\\]|\\.)*)'/g, function (m, x) { return JSON.stringify(x.replace(/\\'/g, "'")); })
        .replace(/\bNone\b/g, "null").replace(/\bTrue\b/g, "true").replace(/\bFalse\b/g, "false").replace(/[\t\r\n]+/g, " "));
    for (var i = 0; i < attempts.length && data === null; i++) {
      try { data = JSON.parse(attempts[i]); } catch (e) { data = null; }
    }
    if (data === null) return { kind: "invalid", error: "The JSON is not valid — the reply is not in a format the add-in reads." };
    return { data: data };
  }

  // returns { items: [{main_file, att_file, cdf_number, ..., flags:[]}], error, kind }
  //   kind (when there is an error): empty | table | nojson | cut | invalid | norows
  function parseAiPaste(text) {
    var s = String(text || "").replace(/^\uFEFF/, "").trim();
    if (!s) return { items: [], error: "Nothing was pasted", kind: "empty" };
    s = s.replace(/\u00a0/g, " ");                               // non-breaking spaces from web pages
    var fence = /```[a-zA-Z]*[ \t]*\r?\n?([\s\S]*?)```/.exec(s);       // ```json, ```javascript, ``` … — any code block
    if (fence) s = fence[1].trim();
    var order = null, list = null;
    // rows copied from Excel (tab-separated): read them when the CDF number tells the column order
    if (looksLikeTable(s)) {
      var tab = splitTable(s), conv = tab && rowsToObjects(tab);
      if (!conv) return { items: [], kind: "table",
        error: "This is a table (columns separated by tabs) whose column order the add-in can't tell. Paste the JSON block from the AI's reply, or a table with a header row." };
      order = conv.order; list = conv.list;
    }
    if (!list) {
      var lj = looseJson(s);
      if (lj.error) return { items: [], kind: lj.kind, error: lj.error };
      var data = lj.data;
      list = Array.isArray(data) ? data : (data && Array.isArray(data.rows) ? data.rows : [data]);
      // a list of lists (values without field names): the order is told by where the CDF number stands
      if (list.length && list.every(function (r) { return Array.isArray(r); })) {
        var conv2 = rowsToObjects(list);
        if (!conv2) return { items: [], kind: "invalid", error: "These rows have values without field names, in an order the add-in can't tell (it looks for the CDF number, e.g. A330-1-978, and the number of values in a row)." };
        order = conv2.order; list = conv2.list;
      } else if (Array.isArray(data) && data.length === 1 && Array.isArray(data[0]) && !data[0].some(Array.isArray)) {
        var conv3 = rowsToObjects([data[0]]);
        if (conv3) { order = conv3.order; list = conv3.list; }
      }
    }
    var items = [];
    list.forEach(function (raw) {
      if (!raw || typeof raw !== "object") return;
      var it = { flags: [] }, unused = [];
      Object.keys(raw).forEach(function (k) {
        var key = k.trim().toLowerCase().replace(/[\s-]+/g, "_");
        key = PASTE_ALIASES[key] || key;
        if (key === "flags") {
          it.flags = it.flags.concat(Array.isArray(raw[k]) ? raw[k].map(String).filter(Boolean) : (raw[k] ? [String(raw[k])] : []));
        } else if (PASTE_KEYS.indexOf(key) >= 0 || key === "prompt_version") {
          var val = raw[k] == null ? "" : String(raw[k]).trim();
          if (key === "form_date" && val === "-") val = "";
          if (!it[key]) it[key] = val;
        } else if (!/^dwg$/.test(key)) unused.push(k);
      });
      if (unused.length) it.flags.push("Not used (not in the add-in's format): " + unused.join(", "));
      if (order) it.flags.push("Read from rows without field names, in this order: " + order + " — check the fields.");
      if (it.prompt_version == null) it.prompt_version = "";
      PASTE_KEYS.forEach(function (k) { if (it[k] == null) it[k] = ""; });
      if (it.cdf_number || it.main_file || it.description) items.push(it);
    });
    if (!items.length) return { items: [], kind: "norows", error: "No CDF rows found in the reply" };
    return { items: items, error: null, kind: null };
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

  // ---------- daily status email in the EMAIL DRAFT format (algorithm 150) ----------
  function dashKey(v) { return String(v == null ? "" : v).toUpperCase().replace(/[\u2010-\u2015\-]/g, "-").replace(/\s+/g, " ").trim(); }

  // a table read from a sheet: rows of cells; a row whose first cell is empty continues the row
  // above, and its empty cells are covered by the cells above (rowspan) - like the merged cells
  function spanRows(rows) {
    var out = [], owner = [];
    rows.forEach(function (row, i) {
      var cont = i > 0 && row[0] === "" && row.some(function (v) { return v !== ""; });
      out.push(row.map(function (v, j) {
        if (cont && v === "" && owner[j]) { owner[j].rowspan++; return null; }
        var c = { v: v, rowspan: 1 };
        owner[j] = c;
        return c;
      }));
    });
    return out;
  }

  // finds tables in sheet values by the start of their title (first column); returns
  // [{title, headers, rows (cells with rowspan, null = covered), startRow, endRow}] (rows 0-based)
  function readSheetTables(values, prefixes) {
    function v(r, c) { var row = values[r]; var x = row && c < row.length ? row[c] : ""; return x == null ? "" : String(x).replace(/\r\n?/g, "\n").trim(); }
    var out = [];
    prefixes.forEach(function (pfx) {
      var want = dashKey(pfx);
      for (var r = 0; r < values.length; r++) {
        if (dashKey(v(r, 0)).indexOf(want) !== 0) continue;
        var width = (values[r + 1] || []).length, n = 0;
        for (var c = 0; c < width; c++) if (v(r + 1, c)) n = c + 1;
        if (!n) continue;
        var headers = [], rows = [], end = r + 1;
        for (c = 0; c < n; c++) headers.push(v(r + 1, c));
        for (var i = r + 2; i < values.length; i++) {
          var row = [];
          for (c = 0; c < n; c++) row.push(v(i, c));
          if (row.every(function (x) { return x === ""; })) break;
          rows.push(row); end = i;
          if (row[0].toUpperCase() === "TOTAL") break;
        }
        out.push({ title: v(r, 0), headers: headers, rows: spanRows(rows), startRow: r, endRow: end });
        break;
      }
    });
    return out;
  }

  // doc: { subject, opening, tables:[{title, headers, rows:[[{v, rowspan}|null]]}], signature:[lines] }
  function buildDailyEmail(doc) {
    var dir = doc.dir === "ltr" ? "ltr" : "rtl", side = dir === "rtl" ? "right" : "left";
    var font = "font-family:Calibri,Arial,sans-serif;font-size:10pt;";
    var border = "border:1px solid #BCBCBC;";
    var th = font + border + "background:#B10001;color:#FFFFFF;font-weight:bold;padding:2px 6px;text-align:center;vertical-align:middle;";
    var td = font + border + "background:#FFFFFF;color:#000000;padding:2px 6px;text-align:left;vertical-align:middle;";
    var html = '<div style="' + font + 'color:#000000">' +
      '<p dir="' + dir + '" style="' + font + 'margin:0 0 12px;text-align:' + side + '">' +
        htmlEsc(doc.opening).replace(/(AS OF [^<]*?\(Asia\/Jerusalem\))/, '<span dir="ltr">$1</span>') + "</p>";
    var text = [doc.opening, ""];
    doc.tables.forEach(function (t, k) {
      var n = t.headers.length;
      html += '<table dir="ltr" cellspacing="0" cellpadding="0" style="border-collapse:collapse;' + font + '">' +
        '<tr><th colspan="' + n + '" style="' + th + '">' + htmlEsc(t.title) + "</th></tr>" +
        "<tr>" + t.headers.map(function (h, i) {
          var w = t.widths && t.widths[i];   // optional column widths (px), so long text doesn't make very tall rows
          return "<th" + (w ? ' width="' + w + '"' : "") + ' style="' + th + (w ? "width:" + w + "px;min-width:" + w + "px;" : "") + '">' + htmlEsc(h) + "</th>";
        }).join("") + "</tr>" +
        t.rows.map(function (row) {
          return "<tr>" + row.map(function (c) {
            if (!c) return "";
            var num = /^\d+$/.test(c.v);
            return "<td" + (c.rowspan > 1 ? ' rowspan="' + c.rowspan + '"' : "") + ' style="' + td + (num ? "text-align:center;" : "") + '">' +
              htmlEsc(c.v).replace(/\n/g, "<br>") + "</td>";
          }).join("") + "</tr>";
        }).join("") + "</table>" +
        (k < doc.tables.length - 1 ? "<br><br>" : "");
      text.push(t.title, t.headers.join("\t"));
      t.rows.forEach(function (row) {
        text.push(row.map(function (c) { return c ? c.v.replace(/\n/g, " / ") : ""; }).join("\t"));
      });
      text.push("", "");
    });
    var sig = (doc.signature && doc.signature.length) ? doc.signature : ["בברכה,"];
    html += '<p dir="' + dir + '" style="' + font + 'margin:16px 0 0;text-align:' + side + '">' + sig.map(htmlEsc).join("<br>") + "</p></div>";
    text = text.concat(sig);
    return { subject: doc.subject, text: text.join("\r\n"), html: html };
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
    parseAiPaste: parseAiPaste, looseJson: looseJson,
    readSheetTables: readSheetTables,
    buildDailyEmail: buildDailyEmail,
    assignNewRowTargets: assignNewRowTargets,
    planNewRowTargets: planNewRowTargets,
    cdfRevision: cdfRevision,
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
