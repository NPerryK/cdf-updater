/*
 * CDF Updater 2 — a PowerPoint deck of the dashboard (PptxGenJS). Native PowerPoint charts, so the numbers stay
 * editable in PowerPoint. build(model, PptxGenJS) → the presentation object; the pane calls writeFile on it.
 *   1 Status · 2 Where the open CDFs are · 3 Opened vs released by week · 4 Age of the open CDFs
 *   5 Open over time (once there are 2 days) · 6 Documents released (MANTRAREV) · 7 Needs a look
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory(require("./cdf-core.js"), require("./cdf-stats.js"), require("./cdf-dash.js"));
  else root.CdfDeck = factory(root.CdfCore, root.CdfStats, root.CdfDash);
})(typeof self !== "undefined" ? self : this, function (K, S, D) {
  "use strict";

  var INK = "13202B", MUTED = "5B6B7A", NAVY = "0E2A47", LINE = "D5DCE3", BG = "FFFFFF";
  var OPEN = "E3A21A", REL = "1F9D8B", CAN = "A3AEB9", WARN = "C8463D", MID = "6C8EB3", SIGN = "7B5EA7", DOC = "3D7DCA";
  var FONT = "Segoe UI";

  function title(pptx, slide, text, sub) {
    slide.background = { color: BG };
    slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 13.333, h: 0.9, fill: { color: NAVY }, line: { color: NAVY } });
    slide.addText(text, { x: 0.5, y: 0.12, w: 9.5, h: 0.66, fontFace: FONT, fontSize: 24, bold: true, color: "FFFFFF", margin: 0 });
    if (sub) slide.addText(sub, { x: 9.2, y: 0.12, w: 3.65, h: 0.66, fontFace: FONT, fontSize: 13, color: "C9D6E2", align: "right", margin: 0 });
  }
  function foot(slide, m, n) {
    slide.addText("A330-300 P2F · CDF status · " + K.ddmmyyyy(m.today), { x: 0.5, y: 7.05, w: 8, h: 0.3, fontFace: FONT, fontSize: 9, color: MUTED, margin: 0 });
    slide.addText(String(n), { x: 12.33, y: 7.05, w: 0.5, h: 0.3, fontFace: FONT, fontSize: 9, color: MUTED, align: "right", margin: 0 });
  }
  function note(slide, text, y) { if (text) slide.addText(text, { x: 0.5, y: y || 6.35, w: 12.3, h: 0.55, fontFace: FONT, fontSize: 13, color: INK, margin: 0, valign: "top" }); }
  function axis() { return { catAxisLabelFontFace: FONT, valAxisLabelFontFace: FONT, catAxisLabelColor: MUTED, valAxisLabelColor: MUTED, catAxisLabelFontSize: 11, valAxisLabelFontSize: 10,
    valGridLine: { color: "E6EBF0", size: 0.75 }, catGridLine: { style: "none" }, showLegend: false, dataLabelFontFace: FONT, dataLabelFontSize: 11, dataLabelColor: INK }; }
  function merge(a, b) { var o = {}; Object.keys(a).forEach(function (k) { o[k] = a[k]; }); Object.keys(b).forEach(function (k) { o[k] = b[k]; }); return o; }
  function kpi(slide, x, y, w, n, label, color) {
    slide.addShape("rect", { x: x, y: y, w: w, h: 1.55, fill: { color: "F5F7F9" }, line: { color: LINE, width: 0.75 } });
    slide.addShape("rect", { x: x, y: y, w: 0.09, h: 1.55, fill: { color: color }, line: { color: color } });
    slide.addText(String(n), { x: x + 0.3, y: y + 0.12, w: w - 0.4, h: 0.85, fontFace: FONT, fontSize: 40, bold: true, color: INK, margin: 0 });
    slide.addText(label, { x: x + 0.3, y: y + 0.98, w: w - 0.4, h: 0.45, fontFace: FONT, fontSize: 13, color: MUTED, margin: 0 });
  }

  function build(m, PptxGenJS, opt) {
    opt = opt || {};
    var pptx = new PptxGenJS(), n = 0;
    pptx.layout = "LAYOUT_WIDE";
    pptx.author = opt.author || "CDF Updater";
    pptx.title = "A330-300 P2F — CDF status " + K.ddmmyyyy(m.today);
    if (m.empty) { var e = pptx.addSlide(); title(pptx, e, "CDF status", K.ddmmyyyy(m.today)); note(e, "Sheet CDF was empty when this deck was made.", 2); return pptx; }
    var st = m.status;

    // 1 — status
    var s = pptx.addSlide(); n++;
    title(pptx, s, "A330-300 P2F — CDF status", K.ddmmyyyy(m.today));
    var dw = S.deltaWeek(m);
    kpi(s, 0.5, 1.3, 2.9, st.open, "Open" + (dw && dw.open ? "  (" + (dw.open > 0 ? "+" : "") + dw.open + " since " + K.ddmm(dw.since) + ")" : ""), OPEN);
    kpi(s, 3.6, 1.3, 2.9, st.released, "Released · " + st.pct + "%", REL);
    kpi(s, 6.7, 1.3, 2.9, st.cancelled, "Cancelled", CAN);
    kpi(s, 9.8, 1.3, 3.0, m.week.closed + " / " + m.week.opened, "Released / opened, last 7 days", MID);
    s.addChart(pptx.ChartType.doughnut, [{ name: "CDFs", labels: ["Released", "Open", "Cancelled"], values: [st.released, st.open, st.cancelled] }],
      { x: 0.5, y: 3.1, w: 4.3, h: 3.2, holeSize: 62, chartColors: [REL, OPEN, CAN], showLegend: true, legendPos: "r", legendFontFace: FONT, legendFontSize: 12, dataLabelColor: "FFFFFF", showPercent: true, showValue: false, dataLabelFontSize: 11 });
    s.addChart(pptx.ChartType.bar, [{ name: "Open", labels: m.stations.map(function (x) { return D.stationShort(x.name); }), values: m.stations.map(function (x) { return x.n; }) }],
      merge(axis(), { x: 5.1, y: 3.1, w: 7.7, h: 3.2, barDir: "bar", chartColors: [OPEN], showValue: true, catAxisOrientation: "maxMin", barGapWidthPct: 45, showTitle: true, title: "Open CDFs by station", titleFontFace: FONT, titleFontSize: 13, titleColor: INK }));
    note(s, D.paceText(m), 6.45); foot(s, m, n);

    // 2 — where
    s = pptx.addSlide(); n++;
    title(pptx, s, "Where the open CDFs are", st.open + " open");
    s.addChart(pptx.ChartType.bar, [{ name: "Open", labels: m.stations.map(function (x) { return x.name; }), values: m.stations.map(function (x) { return x.n; }) }],
      merge(axis(), { x: 0.5, y: 1.2, w: 6.4, h: 5, barDir: "bar", chartColors: [OPEN], showValue: true, catAxisOrientation: "maxMin", barGapWidthPct: 40, showTitle: true, title: "By station", titleFontFace: FONT, titleFontSize: 14, titleColor: INK }));
    if (m.suppliers.conv.length) s.addChart(pptx.ChartType.bar, [{ name: "At Conversion Engineer", labels: m.suppliers.conv.map(function (x) { return x.name; }), values: m.suppliers.conv.map(function (x) { return x.n; }) }],
      merge(axis(), { x: 7.2, y: 1.2, w: 5.6, h: 5, barDir: "bar", chartColors: [MID], showValue: true, catAxisOrientation: "maxMin", barGapWidthPct: 40, showTitle: true, title: "Waiting for the supplier (Conversion Engineer, " + m.suppliers.convTotal + ")", titleFontFace: FONT, titleFontSize: 14, titleColor: INK }));
    foot(s, m, n);

    // 3 — flow
    s = pptx.addSlide(); n++;
    title(pptx, s, "Opened vs released, by week", "last " + m.weekly.length + " weeks");
    var labels = m.weekly.map(function (w) { return K.ddmm(w.week); });
    s.addChart(pptx.ChartType.bar, [{ name: "Opened", labels: labels, values: m.weekly.map(function (w) { return w.opened; }) }, { name: "Released", labels: labels, values: m.weekly.map(function (w) { return w.closed; }) }],
      merge(axis(), { x: 0.5, y: 1.2, w: 12.3, h: 4.9, barDir: "col", barGrouping: "clustered", chartColors: [OPEN, REL], showValue: true, showLegend: true, legendPos: "t", legendFontFace: FONT, legendFontSize: 12, barGapWidthPct: 60 }));
    note(s, D.paceText(m), 6.3); foot(s, m, n);

    // 4 — aging
    s = pptx.addSlide(); n++;
    title(pptx, s, "How long the open CDFs have been open", m.aging.median != null ? "median " + m.aging.median + " days" : "");
    s.addChart(pptx.ChartType.bar, [{ name: "Open", labels: m.aging.buckets.map(function (b) { return b.label; }), values: m.aging.buckets.map(function (b) { return b.n; }) }],
      merge(axis(), { x: 0.5, y: 1.2, w: 6.6, h: 5, barDir: "col", chartColors: [OPEN, OPEN, MID, WARN, WARN], showValue: true, barGapWidthPct: 45 }));
    var rows = [[{ text: "Oldest open", options: { bold: true, color: "FFFFFF", fill: { color: NAVY } } }, { text: "Station", options: { bold: true, color: "FFFFFF", fill: { color: NAVY } } }, { text: "Days", options: { bold: true, color: "FFFFFF", fill: { color: NAVY }, align: "right" } }]]
      .concat(m.aging.oldest.slice(0, 8).map(function (o) { return [o.key, D.stationShort(o.station), { text: String(o.days), options: { align: "right" } }]; }));
    s.addTable(rows, { x: 7.5, y: 1.3, w: 5.3, colW: [2.2, 2.3, 0.8], fontFace: FONT, fontSize: 12, color: INK, border: { type: "solid", color: LINE, pt: 0.75 }, rowH: 0.42 });
    foot(s, m, n);

    // 5 — trend
    if (m.trend.length > 1) {
      s = pptx.addSlide(); n++;
      title(pptx, s, "Open CDFs over time", m.trend.length + " days");
      var tl = m.trend.map(function (t) { return K.ddmm(t.d); }), series = [{ name: "Open", labels: tl, values: m.trend.map(function (t) { return t.open; }) }];
      var tg = m.targets && m.targets.OPEN;
      if (tg) series.push({ name: "Target", labels: tl, values: m.trend.map(function () { return tg.value; }) });
      s.addChart(pptx.ChartType.line, series, merge(axis(), { x: 0.5, y: 1.2, w: 12.3, h: 5, chartColors: [OPEN, WARN], lineSize: 3, lineDataSymbol: "none", showLegend: !!tg, legendPos: "t" }));
      if (tg) note(s, "Target: " + tg.value + " open" + (tg.by ? " by " + K.ddmmyyyy(tg.by) : "") + ".", 6.3);
      foot(s, m, n);
    }

    // 6 — documents
    var d = m.docs;
    if (d && !d.error) {
      s = pptx.addSlide(); n++;
      title(pptx, s, "Documents released (MANTRAREV)", d.total + " releases in the sheet");
      kpi(s, 0.5, 1.3, 2.9, d.last7.length, "Released, last 7 days", DOC);
      kpi(s, 0.5, 3.05, 2.9, d.last30, "Released, last 30 days", DOC);
      if (d.medianDays != null) kpi(s, 0.5, 4.8, 2.9, d.medianDays, "Days to release (median, 30 days)", MID);
      s.addChart(pptx.ChartType.bar, [{ name: "Released", labels: d.weekly.map(function (w) { return K.ddmm(w.week); }), values: d.weekly.map(function (w) { return w.n; }) }],
        merge(axis(), { x: 3.7, y: 1.2, w: 9.1, h: 5.1, barDir: "col", chartColors: [DOC], showValue: true, barGapWidthPct: 50, showTitle: true, title: "Per week", titleFontFace: FONT, titleFontSize: 13, titleColor: INK }));
      if (d.staleDays > 7) note(s, "The newest release in MANTRAREV is from " + K.ddmmyyyy(d.latest) + ": the export is " + d.staleDays + " days old.", 6.45);
      foot(s, m, n);
    }

    // 7 — attention
    s = pptx.addSlide(); n++;
    title(pptx, s, "Needs a look", "");
    kpi(s, 0.5, 1.3, 4, m.noDoc.n, "Open with no engineering document yet", WARN);
    kpi(s, 4.8, 1.3, 4, m.notInGcc.length, "In sheet CDF but not in GCC", MID);
    kpi(s, 9.1, 1.3, 3.7, m.aging.buckets.slice(3).reduce(function (t, b) { return t + b.n; }, 0), "Open more than 3 months", WARN);
    var list = m.aging.oldest.slice(0, 10).map(function (o) { return { text: o.key + " — " + D.stationShort(o.station) + ", " + o.days + " days", options: { bullet: true } }; });
    if (list.length) s.addText([{ text: "Open the longest", options: { bold: true, breakLine: true } }].concat(list), { x: 0.5, y: 3.15, w: 12.3, h: 3.6, fontFace: FONT, fontSize: 14, color: INK, valign: "top", paraSpaceAfter: 4 });
    foot(s, m, n);
    return pptx;
  }
  function fileName(m) { return "CDF_Status_" + K.ddmmyyyy(m.today).split("/").reverse().join("") + ".pptx"; }

  return { build: build, fileName: fileName };
});
