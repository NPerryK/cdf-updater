/*
 * CDF Updater 2 — the dashboard cards, from the model of cdf-stats.js. One renderer for the pane and for the full-screen
 * present view: render(model, { mode: "pane" | "present" }) → { html, lists, cards }. lists maps a data-list id to
 * { title, keys } so a click on a number or a bar opens the CDFs behind it.
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory(require("./cdf-core.js"), require("./cdf-charts.js"), require("./cdf-stats.js"));
  else root.CdfDash = factory(root.CdfCore, root.CdfCharts, root.CdfStats);
})(typeof self !== "undefined" ? self : this, function (K, G, S) {
  "use strict";
  var esc = G.esc;

  function plural(n, w) { return n + " " + w + (n === 1 ? "" : "s"); }
  function sign(n) { return n > 0 ? "+" + n : String(n); }
  function stationShort(s) { return String(s).replace(/^\d+\s*-\s*/, ""); }

  // the sentence about pace: are the open CDFs going down, and when do they run out
  function paceText(m) {
    var p = m.pace;
    if (!p || !p.weeks) return "";
    var base = p.closed + " closed and " + p.opened + " opened a week (last " + p.weeks + " full weeks)";
    if (p.zeroOn) return base + " — at this pace the " + m.status.open + " open CDFs are done in about " + plural(p.weeksToZero, "week") + ", around " + K.ddmmyyyy(p.zeroOn) + ".";
    if (p.net < 0) return base + " — more come in than go out: the open count grows by about " + Math.abs(p.net) + " a week.";
    return base + " — the open count holds steady.";
  }
  function freshText(m, today) {
    var out = [];
    if (m.fresh && m.fresh.gcc != null) out.push({ name: "GCC", day: m.fresh.gcc, old: today - m.fresh.gcc > 3 });
    if (m.docs && m.docs.latest != null) out.push({ name: "MANTRAREV", day: m.docs.latest, old: today - m.docs.latest > 7 });
    return out;
  }

  function render(m, opt) {
    opt = opt || {};
    var present = opt.mode === "present", lists = {}, cards = [], W = present ? 520 : 320, LW = present ? 170 : 118;
    function list(id, title, keys) { lists[id] = { title: title, keys: keys || [] }; return id; }
    function tap(id, n) { return n ? ' data-list="' + esc(id) + '" tabindex="0" role="button"' : ""; }
    function card(id, title, body, sub) { cards.push({ id: id, title: title }); return '<section class="dcard" data-card="' + id + '"><div class="dhead"><h3>' + esc(title) + "</h3>" + (sub ? '<span class="dsub">' + sub + "</span>" : "") + "</div>" + body + "</section>"; }

    if (m.empty) {
      var e = '<div class="dempty"><b>No CDF sheet to read yet.</b> Import the CDF file (or open the workbook that has it) and the dashboard fills in.</div>';
      return { html: e, lists: lists, cards: cards };
    }
    var st = m.status, dw = S.deltaWeek(m), html = "";

    // ---- the headline: status, % released, this week
    var kpis = [
      { id: "status:open", label: "Open", n: st.open, keys: st.keys.open, cls: "c-open", d: dw ? dw.open : null, good: -1 },
      { id: "status:released", label: "Released", n: st.released, keys: st.keys.released, cls: "c-rel", d: dw ? dw.released : null, good: 1 },
      { id: "status:cancelled", label: "Cancelled", n: st.cancelled, keys: st.keys.cancelled, cls: "c-can" }
    ];
    var head = '<div class="kpis">' + kpis.map(function (k) {
      list(k.id, k.label + " CDFs", k.keys);
      var d = k.d == null || !k.d ? "" : '<span class="kd ' + (k.d * k.good > 0 ? "up" : "down") + '" title="since ' + K.ddmm(dw.since) + '">' + sign(k.d) + "</span>";
      return '<div class="kpi ' + k.cls + (k.n ? " tap" : "") + '"' + tap(k.id, k.n) + '><span class="kn">' + k.n + '</span><span class="kl">' + k.label + d + "</span></div>";
    }).join("") + '<div class="kpi ringbox">' + G.ring(st.pct, { label: "released", size: present ? 112 : 62 }) + "</div></div>" +
      G.strip([{ n: st.released, cls: "c-rel", label: "Released" }, { n: st.open, cls: "c-open", label: "Open" }, { n: st.cancelled, cls: "c-can", label: "Cancelled" }], { width: W });
    var wk = m.week;
    list("week:closed", "Released in the last 7 days", wk.closedKeys); list("week:opened", "Opened in the last 7 days", wk.openedKeys);
    head += '<div class="wkline"><span class="wk tap"' + tap("week:closed", wk.closed) + "><b>" + wk.closed + "</b> released</span><span class=\"wk tap\"" + tap("week:opened", wk.opened) + "><b>" + wk.opened + "</b> opened</span>" +
      '<span class="wk ' + (wk.closed - wk.opened >= 0 ? "good" : "bad") + '"><b>' + sign(wk.closed - wk.opened) + "</b> net</span><span class=\"wk-sub\">last 7 days</span></div>";
    var fr = freshText(m, m.today);
    if (fr.length) head += '<div class="fresh">' + fr.map(function (f) { return '<span class="' + (f.old ? "old" : "") + '" title="The newest date in the sheet">' + f.name + " to " + K.ddmm(f.day) + (f.old ? " · " + (m.today - f.day) + " days old" : "") + "</span>"; }).join("") + "</div>";
    html += card("status", "CDF status", head, esc(K.ddmmyyyy(m.today)));

    // ---- stations
    var sItems = m.stations.map(function (s) { return { label: stationShort(s.name), n: s.n, list: list("station:" + s.name, "Open at " + s.name, s.keys), cls: /final/i.test(s.name) ? "c-sign" : /conversion/i.test(s.name) ? "c-open" : "c-mid" }; });
    html += card("stations", "Where the open CDFs are", sItems.length ? G.hbars(sItems, { width: W, labelWidth: LW, title: "Open CDFs by station" }) : '<div class="chart-empty">No open CDFs.</div>', st.open + " open");

    // ---- weekly flow + pace
    var wkItems = m.weekly.map(function (x) { return { week: x.week, a: x.opened, b: x.closed, aList: list("wo:" + x.week, "Opened in the week of " + K.ddmm(x.week), x.openedKeys), bList: list("wc:" + x.week, "Released in the week of " + K.ddmm(x.week), x.closedKeys) }; });
    html += card("flow", "Opened vs released, by week",
      '<div class="legend"><span class="lg c-open">opened</span><span class="lg c-closed">released</span><span class="lg avgl">8-week average</span></div>' +
      G.columns2(wkItems, { width: W, height: present ? 190 : 132, aName: "Opened", bName: "Released", avg: [{ v: m.pace.opened, cls: "c-open", label: "Opened a week: " + m.pace.opened }, { v: m.pace.closed, cls: "c-closed", label: "Released a week: " + m.pace.closed }] }) +
      '<p class="dnote">' + esc(paceText(m)) + "</p>");

    // ---- aging
    var ag = m.aging, aItems = ag.buckets.map(function (b, i) { return { label: b.label, n: b.n, list: list("age:" + i, "Open " + b.label, b.keys), cls: i >= 3 ? "c-warn" : i === 2 ? "c-mid" : "c-open" }; });
    var oldest = ag.oldest.slice(0, present ? 8 : 5).map(function (o) { return '<li data-key="' + esc(o.key) + '" class="tap" tabindex="0" role="button"><span class="ok">' + esc(o.key.replace(/^A330-1-/, "")) + '</span><span class="os">' + esc(stationShort(o.station)) + '</span><span class="od">' + o.days + " d</span></li>"; }).join("");
    html += card("aging", "Age of the open CDFs", G.hbars(aItems, { width: W, labelWidth: LW, title: "Open CDFs by age" }) +
      (oldest ? '<div class="sublabel">Oldest open</div><ol class="oldest">' + oldest + "</ol>" : "") +
      (ag.unknown ? '<p class="dnote">' + plural(ag.unknown, "open CDF") + " without an open date.</p>" : ""), ag.median != null ? "median " + ag.median + " d" : "");

    // ---- suppliers at Conversion Engineer
    var sp = m.suppliers, spItems = sp.conv.map(function (s) { return { label: s.name, n: s.n, list: list("sup:" + s.name, "At Conversion Engineer — " + s.name, s.keys), cls: s.name === "Unassigned" ? "c-mid" : "c-open" }; });
    html += card("suppliers", "Waiting for the supplier (Conversion Engineer)", spItems.length ? G.hbars(spItems, { width: W, labelWidth: LW, title: "Open at Conversion Engineer by supplier" }) : '<div class="chart-empty">None at Conversion Engineer.</div>', plural(sp.convTotal, "CDF"));

    // ---- attention: no engineering document, not in GCC
    list("nodoc", "Open, waiting for an engineering document", m.noDoc.keys);
    list("notgcc", "In sheet CDF but not in GCC", m.notInGcc);
    var att = '<div class="attn"><div class="at tap"' + tap("nodoc", m.noDoc.n) + '><span class="an">' + m.noDoc.n + '</span><span class="al">open with no engineering document yet</span></div>' +
      '<div class="at tap"' + tap("notgcc", m.notInGcc.length) + '><span class="an">' + m.notInGcc.length + '</span><span class="al">in CDF, not in GCC' + (m.notInGcc.length ? ": " + esc(m.notInGcc.slice(0, 3).map(function (k) { return k.replace(/^A330-1-/, ""); }).join(", ")) + (m.notInGcc.length > 3 ? " …" : "") : "") + "</span></div></div>";
    html += card("attention", "Needs a look", att);

    // ---- trend
    var tg = m.targets && m.targets.OPEN ? m.targets.OPEN : null;
    html += card("trend", "Open CDFs over time", G.line(m.trend.map(function (t) { return { d: t.d, v: t.open }; }), { width: W, height: present ? 170 : 96, target: tg, title: "Open CDFs over time",
      empty: "The line starts tomorrow: the add-in keeps one point a day in this workbook." }) + (tg ? '<p class="dnote">Target: ' + tg.value + " open" + (tg.by ? " by " + K.ddmmyyyy(tg.by) : "") + ".</p>" : ""), m.trend.length > 1 ? plural(m.trend.length, "day") : "");

    // ---- solution documents: released or in work (MANTRAREV)
    var d = m.docs, T = "Solution documents";
    if (d && !d.error) {
      list("doc:released", "Open CDFs with a released solution revision", d.keys.released);
      list("doc:work", "Open CDFs with a solution revision in work", d.keys.work);
      var wl = d.work.slice(0, present ? 10 : 5).map(function (x) {
        return '<li data-key="' + esc(x.keys[0]) + '" class="tap" tabindex="0" role="button"><span class="ok">' + esc(x.doc + " " + x.rev) + '</span><span class="os">' + esc(x.released ? "rev " + x.released + " released" : "no revision released") + '</span><span class="od">' + esc(x.keys[0].replace(/^A330-1-/, "")) + (x.keys.length > 1 ? " +" + (x.keys.length - 1) : "") + "</span></li>";
      }).join("");
      html += card("docs", T,
        '<div class="attn"><div class="at tap"' + tap("doc:released", d.released.length) + '><span class="an c-rel-t">' + d.released.length + '</span><span class="al">revisions released</span></div>' +
        '<div class="at tap"' + tap("doc:work", d.work.length) + '><span class="an c-work-t">' + d.work.length + '</span><span class="al">revisions in work</span></div></div>' +
        (wl ? '<div class="sublabel">In work</div><ol class="oldest docs">' + wl + "</ol>" : "") +
        (d.latest != null && m.today - d.latest > 7 ? '<p class="dnote warn">MANTRAREV ends on ' + K.ddmmyyyy(d.latest) + " — a revision released since then still shows as in work. Import a fresh export.</p>" : ""), "open CDFs · MANTRAREV");
    } else if (d && d.error) html += card("docs", T, '<p class="dnote warn">' + esc(d.error) + "</p>");
    else html += card("docs", T, '<div class="chart-empty">No MANTRAREV sheet yet — import it to tell released revisions from revisions in work.</div>');

    return { html: html, lists: lists, cards: cards };
  }

  // the status as plain text and a small HTML table (copy into an email or Teams)
  function summaryText(m) {
    if (m.empty) return "";
    var st = m.status, lines = ["A330-300 P2F — CDF status " + K.ddmmyyyy(m.today),
      "Open " + st.open + " · Released " + st.released + " (" + st.pct + "%) · Cancelled " + st.cancelled,
      "Last 7 days: " + m.week.closed + " released, " + m.week.opened + " opened",
      "Open by station: " + m.stations.map(function (s) { return stationShort(s.name) + " " + s.n; }).join(", ")];
    if (m.suppliers.conv.length) lines.push("At Conversion Engineer: " + m.suppliers.conv.map(function (s) { return s.name + " " + s.n; }).join(", "));
    var p = paceText(m); if (p) lines.push(p);
    return lines.join("\n");
  }
  function summaryHtml(m) {
    if (m.empty) return "";
    var st = m.status, td = 'style="border:1px solid #9aa7b4;padding:3px 8px"', th = 'style="border:1px solid #9aa7b4;padding:3px 8px;background:#14324f;color:#fff;text-align:left"';
    var rows = [["Open", st.open], ["Released", st.released + " (" + st.pct + "%)"], ["Cancelled", st.cancelled], ["Released, last 7 days", m.week.closed], ["Opened, last 7 days", m.week.opened]]
      .concat(m.stations.map(function (s) { return ["  at " + s.name, s.n]; }));
    return '<p style="font-family:Calibri,Arial;font-size:11pt"><b>A330-300 P2F — CDF status ' + K.ddmmyyyy(m.today) + '</b></p><table style="border-collapse:collapse;font-family:Calibri,Arial;font-size:10pt"><tr><th ' + th + ">CDF</th><th " + th + ">#</th></tr>" +
      rows.map(function (r) { return "<tr><td " + td + ">" + esc(r[0]) + "</td><td " + td + ">" + esc(r[1]) + "</td></tr>"; }).join("") + '</table><p style="font-family:Calibri,Arial;font-size:10pt">' + esc(paceText(m)) + "</p>";
  }

  return { render: render, paceText: paceText, summaryText: summaryText, summaryHtml: summaryHtml, stationShort: stationShort };
});
