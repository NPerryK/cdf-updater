// In-memory mock of the Office.js / Excel API subset used by the task pane.
(function () {
  var wb = { sheets: [], nextId: 1, calls: [] };
  window.__wb = wb;

  function newSheet(name, values, formats, f, merges) {
    return { id: "{S" + (wb.nextId++) + "}", name: name, values: values || [], formats: formats || [], f: f || [], merges: merges || [] };
  }
  wb.lost = [];       // values a merge threw away (the add-in should never cause one)
  // =IF($R$7="","",$R$7) style formulas are worked out; any other formula keeps its stored value
  function evalCell(sh, r, c, depth) {
    var fr = sh.f[r], f = fr && fr[c];
    var v = (sh.values[r] || [])[c];
    if (!f || depth > 8) return v === undefined ? "" : v;
    var m = /^=IF\(\$([A-Z]+)\$(\d+)="","",\$\1\$\2\)$/.exec(f);
    if (m) return evalCell(sh, +m[2] - 1, colIdx(m[1]), depth + 1);
    try { var x = formula(sh, f.slice(1), depth); return Array.isArray(x) ? x[0] : x; } catch (e) { return v === undefined ? "" : v; }
  }
  wb.evalCellForTest = function (sh, r, c) { return evalCell(sh, r, c, 0); };
  // a small formula engine for what the add-in writes: COUNTIFS, COUNTIF, SUM, COUNTA, SUMPRODUCT, + - * / &, refs, ranges, {arrays}
  function formula(sh0, src, depth) {
    var toks = [], i = 0, re = /\s*(?:("(?:[^"]|"")*")|(\{)|(\})|((?:'[^']+'|[A-Za-z][\w ]*?)!)?(\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?|\$?[A-Z]{1,3}:\$?[A-Z]{1,3})|([A-Z]+)\(|(\d+(?:\.\d+)?)|(<>|<=|>=|[-+*\/&(),;=<>]))/y;
    while (i < src.length) {
      re.lastIndex = i; var m = re.exec(src);
      if (!m) throw new Error("formula? " + src.slice(i));
      i = re.lastIndex;
      if (m[1]) toks.push({ t: "s", v: m[1].slice(1, -1).replace(/""/g, '"') });
      else if (m[2]) toks.push({ t: "{" }); else if (m[3]) toks.push({ t: "}" });
      else if (m[5]) toks.push({ t: "ref", sheet: m[4] ? m[4].slice(0, -1).replace(/^'|'$/g, "") : null, a: m[5].replace(/\$/g, "") });
      else if (m[6]) toks.push({ t: "fn", v: m[6] });
      else if (m[7]) toks.push({ t: "n", v: +m[7] });
      else toks.push({ t: "op", v: m[8] });
    }
    var k = 0;
    function peek() { return toks[k]; }
    function eat(v) { var x = toks[k++]; if (v && (!x || x.v !== v && x.t !== v)) throw new Error("expected " + v); return x; }
    function sheetOf(name) { return name ? find(name) : sh0; }
    function rangeVals(tk) {
      var sh = sheetOf(tk.sheet), a = tk.a, out = [];
      var colOnly = /^([A-Z]+):([A-Z]+)$/.exec(a);
      if (colOnly) { var cc = colIdx(colOnly[1]); for (var r = 0; r < sh.values.length; r++) out.push(evalCell(sh, r, cc, depth + 1)); return out; }
      var p = a.split(":"), m1 = /([A-Z]+)(\d+)/.exec(p[0]), m2 = /([A-Z]+)(\d+)/.exec(p[1] || p[0]);
      for (var rr = +m1[2] - 1; rr <= +m2[2] - 1; rr++) for (var c2 = colIdx(m1[1]); c2 <= colIdx(m2[1]); c2++) out.push(evalCell(sh, rr, c2, depth + 1));
      return out;
    }
    function scalar(x) { return Array.isArray(x) ? x[0] : x; }
    function num(x) { x = scalar(x); return typeof x === "number" ? x : x === "" || x == null ? 0 : isNaN(+x) ? 0 : +x; }
    function matches(v, crit) {
      var s = String(v == null ? "" : v).trim().toUpperCase(), c = String(crit == null ? "" : crit).toUpperCase();
      if (c === "<>") return s !== "";
      if (c.indexOf("<>") === 0) return s !== c.slice(2);
      if (/[*?]/.test(c)) return new RegExp("^" + c.replace(/[.+^${}()|[\]\\]/g, "\\$&").replace(/\*/g, ".*").replace(/\?/g, ".") + "$").test(s);
      if (typeof crit === "number") return typeof v === "number" ? v === crit : s === String(crit);
      return s === c;
    }
    function countifs(args) {
      // a criterion may be an array (constant or range) → an array of counts
      var n = 1; args.forEach(function (a, j) { if (j % 2 && Array.isArray(a.v)) n = Math.max(n, a.v.length); });
      var out = [];
      for (var x = 0; x < n; x++) {
        var cnt = 0, len = args[0].v.length;
        for (var r = 0; r < len; r++) {
          var ok = true;
          for (var j = 0; j < args.length; j += 2) { var cr = args[j + 1].v; if (!matches(args[j].v[r], Array.isArray(cr) ? cr[x] : cr)) { ok = false; break; } }
          if (ok) cnt++;
        }
        out.push(cnt);
      }
      return out.length === 1 && !args.some(function (a, j) { return j % 2 && Array.isArray(a.v); }) ? out[0] : out;
    }
    function call(name) {
      var args = [];
      if (peek().v !== ")") { args.push(expr()); while (peek() && (peek().v === "," || peek().v === ";")) { eat(); args.push(expr()); } }
      eat(")");
      var A = function (j) { return args[j]; };
      if (name === "COUNTIFS" || name === "COUNTIF") return countifs(args.map(function (a) { return { v: a }; }));
      if (name === "SUM" || name === "SUMPRODUCT") return args.reduce(function (s, a) { return s + (Array.isArray(a) ? a.reduce(function (t, x) { return t + num(x); }, 0) : num(a)); }, 0);
      if (name === "COUNTA") return A(0).filter(function (x) { return x !== "" && x != null; }).length;
      throw new Error("fn " + name);
    }
    function atom() {
      var tk = eat();
      if (tk.t === "n") return tk.v;
      if (tk.t === "s") return tk.v;
      if (tk.t === "{") { var arr = [atom()]; while (peek().v === "," || peek().v === ";") { eat(); arr.push(atom()); } eat("}"); return arr; }
      if (tk.t === "ref") { var v = rangeVals(tk); return /:/.test(tk.a) ? v : v[0]; }
      if (tk.t === "fn") return call(tk.v);
      if (tk.t === "op" && tk.v === "(") { var e = expr(); eat(")"); return e; }
      if (tk.t === "op" && tk.v === "-") return -num(atom());
      throw new Error("atom " + JSON.stringify(tk));
    }
    function bin(a, b, op) {
      var f = op === "+" ? function (x, y) { return num(x) + num(y); } : op === "-" ? function (x, y) { return num(x) - num(y); } : op === "*" ? function (x, y) { return num(x) * num(y); } : op === "/" ? function (x, y) { return num(x) / num(y); } : function (x, y) { return String(x == null ? "" : x) + String(y == null ? "" : y); };
      if (Array.isArray(a) || Array.isArray(b)) { var n = Math.max(Array.isArray(a) ? a.length : 1, Array.isArray(b) ? b.length : 1), o = []; for (var x = 0; x < n; x++) o.push(f(Array.isArray(a) ? a[x] : a, Array.isArray(b) ? b[x] : b)); return o; }
      return f(a, b);
    }
    function term() { var a = atom(); while (peek() && (peek().v === "*" || peek().v === "/")) { var op = eat().v; a = bin(a, atom(), op); } return a; }
    function arith() { var a = term(); while (peek() && (peek().v === "+" || peek().v === "-")) { var op = eat().v; a = bin(a, term(), op); } return a; }
    function expr() { var a = arith(); while (peek() && peek().v === "&") { eat(); a = bin(a, arith(), "&"); } return a; }
    var out = expr();
    if (k < toks.length) throw new Error("left: " + JSON.stringify(toks.slice(k)));
    return out;
  }
  function shiftRefs(sh, at, n) {           // rows from `at` (0-based) move by n: absolute refs follow, as in Excel
    sh.f.forEach(function (row) {
      if (!row) return;
      row.forEach(function (f, c) {
        if (!f) return;
        row[c] = f.replace(/\$([A-Z]+)\$(\d+)/g, function (all, L, R) { var r0 = +R - 1; if (r0 >= at) { if (n < 0 && r0 < at - n) return "#REF!"; return "$" + L + "$" + (r0 + n + 1); } return all; });
      });
    });
  }
  wb.addSheet = function (name, values, formats, f, merges, filtered) { var s = newSheet(name, values, formats, f, merges); s.filtered = !!filtered; wb.sheets.push(s); return s; };
  wb.get = function (name) { return wb.sheets.filter(function (s) { return s.name === name; })[0]; };

  function ensure(sh, r, c) {
    while (sh.values.length <= r) sh.values.push([]);
    while (sh.formats.length <= r) sh.formats.push([]);
    var row = sh.values[r], frow = sh.formats[r];
    while (row.length <= c) row.push("");
    while (frow.length <= c) frow.push("General");
  }

  function Range(sh, r, c, nr, nc) {
    this._sh = sh; this.rowIndex = r; this.columnIndex = c; this.rowCount = nr; this.columnCount = nc;
    this.isNullObject = false;
    var self = this;
    // formats are recorded per range (wb.fmt) so a test can look at them
    var fmt = { font: {}, fill: {}, autofitColumns: function () {}, autofitRows: function () {}, borders: { getItem: function (side) { var b = { side: side }; (wb.borders = wb.borders || []).push({ sheet: sh.name, r: r, c: c, nr: nr, nc: nc, b: b }); return b; } } };
    (wb.fmt = wb.fmt || []).push({ sheet: sh.name, r: r, c: c, nr: nr, nc: nc, f: fmt });
    this.format = fmt;
    Object.defineProperty(this, "values", {
      get: function () {
        var out = [];
        for (var i = 0; i < nr; i++) {
          var row = [];
          for (var j = 0; j < nc; j++) row.push(evalCell(sh, r + i, c + j, 0));
          out.push(row);
        }
        return out;
      },
      set: function (v) {
        if (v.length !== nr || v[0].length !== nc) throw new Error("dimension mismatch " + v.length + "x" + v[0].length + " vs " + nr + "x" + nc);
        for (var i = 0; i < nr; i++) for (var j = 0; j < nc; j++) { ensure(sh, r + i, c + j); sh.values[r + i][c + j] = v[i][j]; if (sh.f[r + i]) sh.f[r + i][c + j] = undefined; }
      }
    });
    Object.defineProperty(this, "formulas", {
      get: function () {
        var out = [];
        for (var i = 0; i < nr; i++) {
          var row = [];
          for (var j = 0; j < nc; j++) { var f = (sh.f[r + i] || [])[c + j]; row.push(f || evalCell(sh, r + i, c + j, 0)); }
          out.push(row);
        }
        return out;
      },
      set: function (v) {
        if (v.length !== nr || v[0].length !== nc) throw new Error("formulas dimension mismatch");
        for (var i = 0; i < nr; i++) for (var j = 0; j < nc; j++) {
          ensure(sh, r + i, c + j); sh.f[r + i] = sh.f[r + i] || [];
          var x = v[i][j];
          if (typeof x === "string" && x.charAt(0) === "=") { sh.f[r + i][c + j] = x; sh.values[r + i][c + j] = ""; }
          else { sh.f[r + i][c + j] = undefined; sh.values[r + i][c + j] = x; }
        }
      }
    });
    Object.defineProperty(this, "numberFormat", {
      get: function () {
        var out = [];
        for (var i = 0; i < nr; i++) {
          var row = [];
          for (var j = 0; j < nc; j++) { var f = (sh.formats[r + i] || [])[c + j]; row.push(f || "General"); }
          out.push(row);
        }
        return out;
      },
      set: function (v) {
        for (var i = 0; i < nr; i++) for (var j = 0; j < nc; j++) { ensure(sh, r + i, c + j); sh.formats[r + i][c + j] = v[i][j]; }
      }
    });
  }
  Range.prototype.load = function () { return this; };
  Range.prototype.clear = function () {
    var sh = this._sh;
    for (var i = 0; i < this.rowCount; i++) for (var j = 0; j < this.columnCount; j++) {
      var r = this.rowIndex + i, c = this.columnIndex + j;
      if (sh.values[r] && c < sh.values[r].length) sh.values[r][c] = "";
      if (sh.formats[r] && c < sh.formats[r].length) sh.formats[r][c] = "General";
      if (sh.f[r]) sh.f[r][c] = undefined;
    }
    wb.calls.push("clear " + sh.name + " " + (this.rowIndex + 1) + "+" + this.rowCount);
  };
  function overlaps(a, b) { return a.r1 <= b.r2 && b.r1 <= a.r2 && a.c1 <= b.c2 && b.c1 <= a.c2; }
  Range.prototype._box = function () { return { r1: this.rowIndex, c1: this.columnIndex, r2: this.rowIndex + this.rowCount - 1, c2: this.columnIndex + this.columnCount - 1 }; };
  Range.prototype.insert = function (dir) {
    var sh = this._sh, at = this.rowIndex, n = this.rowCount;
    if (!this._rows || dir !== "Down") throw new Error("mock: only whole rows, Down (got " + dir + ")");
    wb.calls.push("insertRows " + (at + 1) + "+" + n);
    while (sh.values.length < at) sh.values.push([]);
    var blank = []; for (var i = 0; i < n; i++) blank.push([]);
    shiftRefs(sh, at, n);
    [sh.values, sh.formats, sh.f].forEach(function (a) { while (a.length < at) a.push([]); Array.prototype.splice.apply(a, [at, 0].concat(blank.map(function () { return []; }))); });
    sh.merges.forEach(function (m) {
      if (m.r1 >= at) { m.r1 += n; m.r2 += n; }
      else if (m.r2 >= at) m.r2 += n;               // Excel grows a merge the rows go into
    });
    var links = {};
    Object.keys(sh.links || {}).forEach(function (k) { var p = k.split(":"), row = +p[0]; links[(row > at ? row + n : row) + ":" + p[1]] = sh.links[k]; });
    sh.links = links;
  };
  Range.prototype.delete = function (dir) {
    var sh = this._sh, at = this.rowIndex, n = this.rowCount;
    if (!this._rows && dir === "Up") {
      // cells below move up in these columns only (BACKUPDATA's areas trim this way)
      var c1 = this.columnIndex, c2 = c1 + this.columnCount - 1, len = sh.values.length;
      wb.calls.push("deleteUp " + sh.name + "!" + colName(c1) + (at + 1) + ":" + colName(c2) + (at + n));
      [sh.values, sh.formats, sh.f].forEach(function (a) {
        for (var r = at; r < len; r++) {
          var src = a[r + n];
          for (var c = c1; c <= c2; c++) {
            var v = src ? src[c] : undefined;
            if (!a[r]) { if (v === undefined) continue; a[r] = []; }
            a[r][c] = v === undefined ? (a === sh.values ? "" : undefined) : v;
          }
        }
      });
      return;
    }
    if (!this._rows || dir !== "Up") throw new Error("mock: only whole rows, Up (got " + dir + ")");
    wb.calls.push("deleteRows " + (at + 1) + "+" + n);
    [sh.values, sh.formats, sh.f].forEach(function (a) { a.splice(at, n); });
    shiftRefs(sh, at + n, -n);
    sh.merges = sh.merges.filter(function (m) {
      var lo = at, hi = at + n - 1;
      if (m.r1 > hi) { m.r1 -= n; m.r2 -= n; return true; }
      if (m.r2 < lo) return true;
      var keepTop = Math.max(0, lo - m.r1), keepBottom = Math.max(0, m.r2 - hi);
      if (keepTop + keepBottom === 0) return false;
      m.r1 = m.r1 < lo ? m.r1 : lo; m.r2 = m.r1 + keepTop + keepBottom - 1;
      return m.r2 > m.r1 || m.c2 > m.c1;
    });
  };
  Range.prototype.merge = function (across) {
    var sh = this._sh, b = this._box();
    if (across) throw new Error("mock: merge across not used");
    if (sh.merges.some(function (m) { return overlaps(m, b); })) throw new Error("mock: merge over an existing merge at " + sh.name + "!" + colName(b.c1) + (b.r1 + 1) + ":" + colName(b.c2) + (b.r2 + 1) + " (existing: " + sh.merges.filter(function (m) { return overlaps(m, b); }).map(function (m) { return colName(m.c1) + (m.r1 + 1) + ":" + colName(m.c2) + (m.r2 + 1); }).join(",") + ")");
    for (var r = b.r1; r <= b.r2; r++) for (var c = b.c1; c <= b.c2; c++) {
      if (r === b.r1 && c === b.c1) continue;
      var v = evalCell(sh, r, c, 0);
      if (v !== "" && v !== null && v !== undefined) { wb.lost.push(sh.name + "!" + c + ":" + (r + 1) + "=" + v); ensure(sh, r, c); sh.values[r][c] = ""; if (sh.f[r]) sh.f[r][c] = undefined; }
    }
    if (b.r2 > b.r1 || b.c2 > b.c1) sh.merges.push(b);
    if (window.__mergeTrace && b.c2 > 7 && sh.name === "CDF SUMMARY") (wb.mergeLog = wb.mergeLog || []).push(sh.name + "!" + colName(b.c1) + (b.r1 + 1) + ":" + colName(b.c2) + (b.r2 + 1) + " ← " + (new Error().stack || "").split("\n").filter(function (l) { return /addin\.test/.test(l); }).slice(0, 4).map(function (l) { return l.trim().replace(/https:\/\/addin\.test\//, ""); }).join(" < "));
  };
  Range.prototype.unmerge = function () { var b = this._box(); this._sh.merges = this._sh.merges.filter(function (m) { return !overlaps(m, b); }); };
  Range.prototype.getMergedAreasOrNullObject = function () {
    var b = this._box(), sh = this._sh, hit = sh.merges.filter(function (m) { return overlaps(m, b); });
    if (!hit.length) return { isNullObject: true, load: function () { return this; } };
    return { isNullObject: false, load: function () { return this; }, address: hit.map(function (m) { return sh.name + "!" + colName(m.c1) + (m.r1 + 1) + ":" + colName(m.c2) + (m.r2 + 1); }).join(",") };
  };
  Range.prototype.copyFrom = function (src, type) { wb.calls.push("copyFrom " + src + " " + type + " -> " + (this.rowIndex + 1)); };
  Range.prototype.select = function () { wb.calls.push("select " + this._sh.name + "!" + (this.rowIndex + 1)); };
  Range.prototype.getEntireRow = function () {
    var rr = new Range(this._sh, this.rowIndex, 0, 1, 16384); rr._rows = true;
    return { insert: function (dir) { wb.calls.push("insertRow " + (rr.rowIndex + 1)); return rr.insert(dir); } };
  };
  Object.defineProperty(Range.prototype, "hyperlink", {
    set: function (h) {
      var sh = this._sh; sh.links = sh.links || {};
      sh.links[(this.rowIndex + 1) + ":" + this.columnIndex] = h;
      ensure(sh, this.rowIndex, this.columnIndex);
      sh.values[this.rowIndex][this.columnIndex] = h.textToDisplay;
    }
  });

  function colIdx(s) { var n = 0; for (var i = 0; i < s.length; i++) n = n * 26 + s.charCodeAt(i) - 64; return n - 1; }
  function colName(i) { var n = i + 1, s = ""; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }

  function Worksheet(sh) {
    this._sh = sh; this.isNullObject = false;
    this.autoFilter = { load: function () { return this; }, get isDataFiltered() { return !!sh.filtered; } };
    this.freezePanes = { freezeRows: function () {} };
  }
  Object.defineProperty(Worksheet.prototype, "name", {
    get: function () { return this._sh.name; },
    set: function (n) {
      if (wb.sheets.some(function (s) { return s !== this._sh && s.name.toLowerCase() === n.toLowerCase(); }, this)) throw Object.assign(new Error("name conflict: " + n), { code: "InvalidArgument" });
      this._sh.name = n;
    }
  });
  Object.defineProperty(Worksheet.prototype, "id", { get: function () { return this._sh.id; } });
  Object.defineProperty(Worksheet.prototype, "position", {
    get: function () { return wb.sheets.indexOf(this._sh); },
    set: function (p) { var i = wb.sheets.indexOf(this._sh); wb.sheets.splice(i, 1); wb.sheets.splice(p, 0, this._sh); }
  });
  Worksheet.prototype.load = function () { return this; };
  Worksheet.prototype.activate = function () { wb.active = this._sh.name; wb.calls.push("activate " + this._sh.name); };
  Object.defineProperty(Worksheet.prototype, "visibility", {
    get: function () { return this._sh.visibility || "Visible"; },
    set: function (v) { this._sh.visibility = v; wb.calls.push("visibility " + this._sh.name + " " + v); }
  });
  Worksheet.prototype.delete = function () {
    var i = wb.sheets.indexOf(this._sh); if (i < 0) throw new Error("already deleted"); wb.sheets.splice(i, 1); wb.calls.push("delete " + this._sh.name);
  };
  Worksheet.prototype.getUsedRangeOrNullObject = function () {
    var sh = this._sh, lastR = -1, lastC = -1;
    sh.values.forEach(function (row, r) { row.forEach(function (v, c) { if (v !== "" && v !== null && v !== undefined) { if (r > lastR) lastR = r; if (c > lastC) lastC = c; } }); });
    if (lastR < 0) return { isNullObject: true, load: function () {} };
    return new Range(sh, 0, 0, lastR + 1, lastC + 1);
  };
  Worksheet.prototype.getRangeByIndexes = function (r, c, nr, nc) { return new Range(this._sh, r, c, nr, nc); };
  Worksheet.prototype.getRange = function (addr) {
    var cols = /^([A-Z]+):([A-Z]+)$/.exec(addr);
    if (cols) return new Range(this._sh, 0, colIdx(cols[1]), 1, colIdx(cols[2]) - colIdx(cols[1]) + 1);
    var rows = /^(\d+):(\d+)$/.exec(addr);
    if (rows) { var rr = new Range(this._sh, +rows[1] - 1, 0, +rows[2] - +rows[1] + 1, 16384); rr._rows = true; return rr; }
    if (/^[A-Z]+\d+$/.test(addr)) addr = addr + ":" + addr;
    var m = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(addr);
    return new Range(this._sh, +m[2] - 1, colIdx(m[1]), +m[4] - +m[2] + 1, colIdx(m[3]) - colIdx(m[1]) + 1);
  };

  var NULL = { isNullObject: true, load: function () { return this; } };
  function find(key) { return wb.sheets.filter(function (s) { return s.name === key || s.id === key; })[0]; }

  var worksheets = {
    load: function () { return this; },
    get items() { return wb.sheets.map(function (s) { return new Worksheet(s); }); },
    getItemOrNullObject: function (k) { var s = find(k); return s ? new Worksheet(s) : NULL; },
    getItem: function (k) { var s = find(k); if (!s) throw Object.assign(new Error("ItemNotFound " + k), { code: "ItemNotFound" }); return new Worksheet(s); },
    getFirst: function () { return new Worksheet(wb.sheets[0]); },
    add: function (name) { return new Worksheet(wb.addSheet(name)); }
  };

  var ctx = {
    workbook: {
      worksheets: worksheets,
      customXmlParts: (function () {
        var parts = window.__xmlParts = window.__xmlParts || [];
        function nsOf(xml) { var m = /xmlns="([^"]+)"/.exec(xml); return m ? m[1] : ""; }
        function wrap(p) {
          return { isNullObject: false, id: p.id, load: function () { return this; },
            getXml: function () { return { value: p.xml }; },
            setXml: function (x) { p.xml = x; wb.calls.push("setXml " + nsOf(x)); },
            delete: function () { parts.splice(parts.indexOf(p), 1); } };
        }
        return {
          add: function (xml) { var p = { id: "x" + (parts.length + 1), xml: xml }; parts.push(p); wb.calls.push("addXml " + nsOf(xml)); return wrap(p); },
          getByNamespace: function (ns) {
            var list = parts.filter(function (p) { return nsOf(p.xml) === ns; });
            return { load: function () { return this; }, get items() { return list.map(wrap); },
              getOnlyItemOrNullObject: function () { return list.length === 1 ? wrap(list[0]) : NULL; } };
          }
        };
      })(),
      insertWorksheetsFromBase64: function (b64, opts) {
        wb.calls.push("insert " + JSON.stringify(opts));
        var name = opts.sheetNamesToInsert[0];
        var data = window.__fixtures[name];
        if (!data) throw new Error("fixture missing for " + name);
        var nm = name; while (find(nm)) nm += " (2)";
        var s = newSheet(nm, JSON.parse(JSON.stringify(data)));
        if (opts.positionType === "After") wb.sheets.splice(wb.sheets.indexOf(find(opts.relativeTo)) + 1, 0, s);
        else wb.sheets.push(s);
        if (!b64 || b64.length < 100) throw new Error("bad base64");
        return { value: [s.id] };
      }
    },
    application: { suspendApiCalculationUntilNextSync: function () {} },
    sync: function () { wb.syncs = (wb.syncs || 0) + 1; if (window.__syncTrace) { var st = (new Error().stack || "").split("\n").filter(function (l) { return /taskpane\.js/.test(l); })[0] || "?"; (wb.syncAt = wb.syncAt || []).push(st.replace(/^\s*at\s*/, "").replace(/\(?https:\/\/addin\.test\/taskpane\.js[^:]*:(\d+):\d+\)?/, "@$1")); } return new Promise(function (r) { setTimeout(r, 1); }); }
  };

  (window.__initialSheets || []).forEach(function (s) { wb.addSheet(s.name, s.values, s.formats, s.f, s.merges, s.filtered); });

  window.Excel = {
    run: function (fn) { return Promise.resolve().then(function () { return fn(ctx); }); },
    WorksheetPositionType: { after: "After", end: "End", before: "Before", beginning: "Beginning" }
  };
  window.Office = {
    HostType: { Excel: "Excel" },
    context: { requirements: { isSetSupported: function () { return true; } } },
    onReady: function (cb) { setTimeout(function () { cb({ host: "Excel" }); }, 0); }
  };
})();
