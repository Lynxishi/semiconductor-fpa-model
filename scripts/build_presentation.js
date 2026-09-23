// Build outputs/<TICKER>/presentation/<Division>_<FY>_Executive_Review.pptx from outputs/<TICKER>/presentation/deck_data.json.
// Every figure and every sentence comes from deck_data.json (see scripts/prepare_deck_data.py).
// Run: node scripts/build_presentation.js --company TICKER   (default MCHP, or the FPA_COMPANY environment variable)
const path = require("path");
const pptxgen = require("pptxgenjs");
const ROOT = path.resolve(__dirname, "..");
const ai = process.argv.indexOf("--company");
const TICKER = (ai > 0 ? process.argv[ai + 1] : process.env.FPA_COMPANY || "MCHP").toUpperCase();
const OUTDIR = path.join(ROOT, "outputs", TICKER, "presentation");
const D = require(path.join(OUTDIR, "deck_data.json"));
const N = D.narrative, CO = D.company;
const BG_T = path.join(ROOT, "presentation", "assets", "bg_title.png");
const BG_C = path.join(ROOT, "presentation", "assets", "bg_content.png");

// ---------- palette (matches the dashboard's emerald-glass dark theme) ----------
const C = { card: "121A17", card2: "18231F", line: "26332E", ink: "EEF5F1", ink2: "A9BAB2", muted: "7C8F86",
  emerald: "43C99A", violet: "A497FF", lime: "CDF56E", coral: "F07B6C", sage: "9FB3AA", gold: "B08C19",
  p1: "8676F2", p2: "22A072", p3: "B08C19", good: "43C99A", crit: "EF6B6B", warn: "F2C14E", dark: "0B120F" };
const FONT = "Calibri";
const W = 13.333, M = 0.6;

// ---------- number formatting ----------
const f1 = v => Math.abs(v).toLocaleString("en-US", { minimumFractionDigits: 1, maximumFractionDigits: 1 });
const money = v => (v < 0 ? "−$" : "$") + f1(v) + "m";
const smoney = v => (v >= 0 ? "+$" : "−$") + f1(v) + "m";
const pct = (v, d = 1) => (v * 100).toFixed(d) + "%";
const spct = v => (v >= 0 ? "+" : "−") + Math.abs(v * 100).toFixed(1) + "%";

const B = D.scenarios.Base, K = B.kpis, S = D.static;
const cap = Object.fromEntries(S.capex_inputs.map(x => [x.key, x.value]));
const cs = Object.fromEntries(S.capex_summary.map(x => [x.scenario, x]));
const vreq = Object.fromEntries(S.variance_required.map(x => [x.variance_id, x]));
const vpnl = Object.fromEntries(S.variance_pnl.map(x => [x.line_id, x]));
const vk = (p, k) => S.variance_kpis.find(x => x.product_id === p && x.kpi_id === k);
const pf = id => B.product_fy.find(x => x.product_id === id);
const hist = S.history;
const PRODS = CO.products.map((p, i) => [p.id, p.short, [C.p1, C.p2, C.p3][i]]);   // [id, name, colour]
const niceStep = x => { const e = Math.pow(10, Math.floor(Math.log10(x))); return [5, 2, 1].map(k => k * e).find(v => v <= x) || e; };

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.author = "FP&A portfolio project";
pres.title = N.deck_title;

// ---------- helpers ----------
function base(slide, n, eyebrow, title) {
  slide.background = { path: BG_C };
  slide.addText(eyebrow.toUpperCase(), { x: M, y: 0.38, w: 9, h: 0.3, fontFace: FONT, fontSize: 11, bold: true, color: C.emerald,
    charSpacing: 2, margin: 0, isTextBox: true });
  slide.addText(title, { x: M, y: 0.68, w: W - 2 * M, h: 0.9, fontFace: FONT, fontSize: 30, bold: true, color: C.ink, margin: 0,
    valign: "top", isTextBox: true });
  slide.addText(D.disclaimer,
    { x: M, y: 7.05, w: 10.5, h: 0.25, fontFace: FONT, fontSize: 9, color: C.muted, margin: 0, isTextBox: true });
  slide.addText(String(n), { x: W - M - 0.6, y: 7.05, w: 0.6, h: 0.25, fontFace: FONT, fontSize: 9, color: C.muted, align: "right", margin: 0, isTextBox: true });
}
function card(slide, x, y, w, h, fill = C.card) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.16, fill: { color: fill }, line: { color: C.line, width: 0.75 } });
}
function dot(slide, x, y, color = C.emerald, d = 0.16) {
  slide.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color }, line: { color, width: 0 } });
}
function pill(slide, x, y, w, text, color, textColor = C.dark) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h: 0.34, rectRadius: 0.17, fill: { color }, line: { color, width: 0 } });
  slide.addText(text, { x, y, w, h: 0.34, fontFace: FONT, fontSize: 11, bold: true, color: textColor, align: "center", valign: "middle", margin: 0, isTextBox: true });
}
function stat(slide, x, y, w, h, label, value, sub, color = C.ink, fill) {
  card(slide, x, y, w, h, fill);
  slide.addText(label, { x: x + 0.25, y: y + 0.2, w: w - 0.5, h: 0.3, fontFace: FONT, fontSize: 12, color: fill ? "DDF7EC" : C.ink2, margin: 0, isTextBox: true });
  slide.addText(value, { x: x + 0.25, y: y + 0.5, w: w - 0.5, h: 0.7, fontFace: FONT, fontSize: 34, bold: true, color, margin: 0, isTextBox: true });
  slide.addText(sub, { x: x + 0.25, y: y + 1.2, w: w - 0.5, h: 0.3, fontFace: FONT, fontSize: 11.5, color: fill ? "DDF7EC" : C.muted, margin: 0, isTextBox: true });
}
const axis = { catAxisLabelColor: C.ink2, valAxisLabelColor: C.muted, catAxisLabelFontFace: FONT, valAxisLabelFontFace: FONT,
  catAxisLabelFontSize: 10, valAxisLabelFontSize: 10, valGridLine: { color: "22302A", size: 0.5 }, catGridLine: { style: "none" },
  catAxisLineShow: false, valAxisLineShow: false, legendFontFace: FONT, legendColor: C.ink2, legendFontSize: 10.5,
  plotArea: { fill: { color: C.card } } };

// ===== 1. Title =====
{
  const s = pres.addSlide(); s.background = { path: BG_T };
  s.addShape(pres.shapes.OVAL, { x: M, y: 1.35, w: 0.62, h: 0.62, fill: { color: C.emerald }, line: { color: C.emerald, width: 0 } });
  s.addShape(pres.shapes.OVAL, { x: M + 0.17, y: 1.52, w: 0.28, h: 0.28, fill: { color: C.dark }, line: { color: C.dark, width: 0 } });
  s.addText(N.eyebrow_date, { x: M, y: 2.35, w: 10, h: 0.35, fontFace: FONT, fontSize: 13, bold: true,
    color: C.emerald, charSpacing: 3, margin: 0, isTextBox: true });
  s.addText(CO.division, { x: M, y: 2.75, w: 11.5, h: 1.1, fontFace: FONT, fontSize: 54, bold: true, color: C.ink, margin: 0, isTextBox: true });
  s.addText(N.title_sub,
    { x: M, y: 3.85, w: 11.5, h: 0.6, fontFace: FONT, fontSize: 24, color: C.ink2, margin: 0, isTextBox: true });
  pill(s, M, 4.85, 2.9, `Excel model · ${D.excel_checks_pass}/${D.excel_checks_pass + D.excel_checks_fail} checks pass`, C.emerald);
  pill(s, M + 3.05, 4.85, 3.05, "Web dashboard · verified vs Excel", C.violet, "120B2E");
  pill(s, M + 6.25, 4.85, 2.55, "Downside · Base · Upside", C.lime);
  s.addText(N.title_footer + D.disclaimer, { x: M, y: 6.35, w: 11.8, h: 0.6, fontFace: FONT, fontSize: 10.5, color: C.muted, margin: 0, isTextBox: true });
  s.addNotes(N.title_notes);
}

// ===== 2. Executive summary =====
{
  const s = pres.addSlide();
  base(s, 2, "Executive summary", N.s2_title);
  const w = (W - 2 * M - 0.9) / 4;
  stat(s, M, 1.75, w, 1.65, `Revenue ${CO.forecast_year} (Base)`, money(K.revenue), `${spct(K.revenue_growth)} vs ${CO.base_year} base year`, "FFFFFF", "1F7A5A");
  stat(s, M + (w + 0.3), 1.75, w, 1.65, "Operating profit", money(K.operating_profit), `${pct(K.operating_margin)} margin`, C.ink);
  stat(s, M + 2 * (w + 0.3), 1.75, w, 1.65, "Free cash flow", money(K.fcf), `${pct(K.fcf_margin)} of revenue`, C.ink);
  stat(s, M + 3 * (w + 0.3), 1.75, w, 1.65, "YTD op. profit vs budget", smoney(K.ytd_op_variance),
    `${spct(K.ytd_op_variance / K.ytd_budget_op)} · ${N.act_range} actuals`, C.lime);
  const msgs = [
    [N.s2_m1_head, `YTD revenue ${smoney(K.ytd_revenue_variance)} (${spct(K.ytd_revenue_variance / K.ytd_budget_revenue)}). ${N.s2_mix}`, C.emerald],
    [N.s2_m2_head, N.s2_m2, C.warn],
    ["Capital request", N.s2_m3, C.violet],
  ];
  msgs.forEach(([h, t, col], i) => {
    const y = 3.75 + i * 1.0;
    card(s, M, y, W - 2 * M, 0.85);
    dot(s, M + 0.3, y + 0.34, col, 0.18);
    s.addText(h, { x: M + 0.7, y: y + 0.12, w: 2.9, h: 0.6, fontFace: FONT, fontSize: 16, bold: true, color: C.ink, valign: "middle", margin: 0, isTextBox: true });
    s.addText(t, { x: M + 3.6, y: y + 0.08, w: W - 2 * M - 3.85, h: 0.7, fontFace: FONT, fontSize: 13.5, color: C.ink2, valign: "middle", margin: 0, isTextBox: true });
  });
  s.addNotes(N.s2_notes);
}

// ===== 3. Approach =====
{
  const s = pres.addSlide();
  base(s, 3, "How the model works", "One linked model: engineering drivers flow to cash and decisions");
  const steps = [["Drivers", "Demand, selling price, yield, utilization, material cost, opex, capex, tax"],
    ["Operating model", "Units produced → good units → units sold; cost per good unit; inventory"],
    ["P&L", "Revenue, cost of revenue, gross profit, R&D, SG&A, operating profit"],
    ["Cash flow", "Depreciation, working capital, capex → free cash flow"],
    ["Decisions", "Budget variance, scenario range, capex NPV / IRR / payback"]];
  const bw = 2.12, gap = 0.34;
  steps.forEach(([h, t], i) => {
    const x = M + i * (bw + gap), y = 1.85;
    card(s, x, y, bw, 2.0, i === 0 ? C.card2 : C.card);
    s.addText(String(i + 1).padStart(2, "0"), { x: x + 0.22, y: y + 0.2, w: 0.8, h: 0.3, fontFace: FONT, fontSize: 12, bold: true, color: C.emerald, margin: 0, isTextBox: true });
    s.addText(h, { x: x + 0.22, y: y + 0.5, w: bw - 0.44, h: 0.4, fontFace: FONT, fontSize: 17, bold: true, color: C.ink, margin: 0, isTextBox: true });
    s.addText(t, { x: x + 0.22, y: y + 0.95, w: bw - 0.44, h: 0.95, fontFace: FONT, fontSize: 12, color: C.ink2, valign: "top", margin: 0, isTextBox: true });
    if (i < steps.length - 1) s.addShape(pres.shapes.RIGHT_ARROW, { x: x + bw + 0.06, y: y + 0.85, w: 0.22, h: 0.3, fill: { color: C.emerald }, line: { color: C.emerald, width: 0 } });
  });
  const cols = [["Reported public data", N.s3_reported, C.emerald],
    ["Simulated operational data", N.s3_sim, C.warn],
    ["Independent validation", `${D.excel_checks_pass} Excel checks, a separate Python rebuild of every phase, and a browser test matching ${D.dashboard_test.compared} dashboard values to Excel.`, C.violet]];
  const cw = (W - 2 * M - 0.6) / 3;
  cols.forEach(([h, t, col], i) => {
    const x = M + i * (cw + 0.3), y = 4.2;
    card(s, x, y, cw, 2.45);
    dot(s, x + 0.25, y + 0.3, col, 0.18);
    s.addText(h, { x: x + 0.55, y: y + 0.2, w: cw - 0.8, h: 0.4, fontFace: FONT, fontSize: 15, bold: true, color: C.ink, margin: 0, isTextBox: true });
    s.addText(t, { x: x + 0.25, y: y + 0.75, w: cw - 0.5, h: 1.5, fontFace: FONT, fontSize: 13, color: C.ink2, valign: "top", margin: 0, isTextBox: true });
  });
  s.addNotes("Every step is a visible row in Excel, so any output can be traced back to an assumption. One scenario switch drives the forecast; the budget is locked to the Base case.");
}

// ===== 4. Context: history =====
{
  const s = pres.addSlide();
  base(s, 4, N.s4_eyebrow, N.s4_title);
  card(s, M, 1.75, 7.9, 4.95);
  s.addChart(pres.charts.BAR, [{ name: "Net sales ($m)", labels: hist.map(h => h.fiscal_year), values: hist.map(h => Math.round(h.net_sales * 10) / 10) }],
    { x: M + 0.2, y: 1.9, w: 7.5, h: 4.65, barDir: "col", chartColors: [C.emerald], barGapWidthPct: 55, showValue: true, dataLabelPosition: "outEnd",
      dataLabelColor: C.ink, dataLabelFontFace: FONT, dataLabelFontSize: 11, dataLabelFormatCode: "$#,##0", valAxisLabelFormatCode: "$#,##0",
      showTitle: true, title: N.s4_chart, titleColor: C.ink2, titleFontFace: FONT, titleFontSize: 12,
      showLegend: false, ...axis });
  const rx = M + 8.2, rw = W - M - rx;
  stat(s, rx, 1.75, rw, 1.55, N.s4_stat1_label, N.s4_stat1_value, N.s4_stat1_sub, C.coral);
  stat(s, rx, 3.45, rw, 1.55, N.s4_stat2_label, N.s4_stat2_value, N.s4_stat2_sub, C.lime);
  stat(s, rx, 5.15, rw, 1.55, N.s4_stat3_label, N.s4_stat3_value, N.s4_stat3_sub, C.ink);
  s.addNotes(N.s4_notes);
}

// ===== 5. Forecast year =====
{
  const s = pres.addSlide();
  base(s, 5, `${CO.forecast_year} forecast · Base scenario`, `Revenue grows ${spct(K.revenue_growth)} to ${money(K.revenue)}; operating margin ${pct(K.operating_margin)}`);
  card(s, M, 1.75, 8.3, 4.95);
  const m = B.monthly, labs = m.map(x => new Date(x.month + "T00:00:00").toLocaleString("en-US", { month: "short" }));
  const rv_ = m.flatMap(x => [x.budget_revenue, x.outlook_revenue]), rs = niceStep(Math.max(1e-6, (Math.max(...rv_) - Math.min(...rv_)) / 2.5));
  const rmin = Math.max(0, Math.floor((Math.min(...rv_) - rs) / rs) * rs);
  s.addChart(pres.charts.LINE, [
    { name: "Budget", labels: labs, values: m.map(x => +x.budget_revenue.toFixed(2)) },
    { name: N.outlook_series, labels: labs, values: m.map(x => +x.outlook_revenue.toFixed(2)) }],
    { x: M + 0.2, y: 1.9, w: 7.9, h: 4.65, chartColors: [C.sage, C.emerald], lineSize: 2.5, lineDataSymbol: "circle", lineDataSymbolSize: 6,
      lineDash: ["dash", "solid"], valAxisLabelFormatCode: "$#,##0", valAxisMinVal: rmin, showTitle: true, title: "Monthly revenue, $m",
      titleColor: C.ink2, titleFontFace: FONT, titleFontSize: 12, showLegend: true, legendPos: "b", ...axis });
  const rx = M + 8.6, rw = W - M - rx;
  PRODS.forEach(([id, name, col], i) => {
    const p = pf(id), y = 1.75 + i * 1.3;
    card(s, rx, y, rw, 1.15);
    dot(s, rx + 0.25, y + 0.28, col, 0.18);
    s.addText(name, { x: rx + 0.55, y: y + 0.17, w: rw - 0.8, h: 0.4, fontFace: FONT, fontSize: 14, bold: true, color: C.ink, margin: 0, isTextBox: true });
    s.addText(`${money(p.revenue)}  ·  ${pct(p.mix)} of revenue  ·  ${pct(p.gross_margin)} gross margin`,
      { x: rx + 0.25, y: y + 0.6, w: rw - 0.5, h: 0.4, fontFace: FONT, fontSize: 12, color: C.ink2, margin: 0, isTextBox: true });
  });
  card(s, rx, 5.65, rw, 1.05, "1F7A5A");
  s.addText(`Free cash flow ${money(K.fcf)} (${pct(K.fcf_margin)} of revenue); FY outlook ${smoney(K.outlook_op_vs_budget)} vs budget`,
    { x: rx + 0.25, y: 5.72, w: rw - 0.5, h: 0.9, fontFace: FONT, fontSize: 13, bold: true, color: "FFFFFF", valign: "middle", margin: 0, isTextBox: true });
  s.addNotes(N.s5_notes);
}

// ===== 6. Budget vs actual waterfall =====
{
  const s = pres.addSlide();
  base(s, 6, N.s6_eyebrow, N.s6_title);
  const x0 = M, y0 = 1.75, cw = 8.3, ch = 4.95;
  card(s, x0, y0, cw, ch);
  s.addText("Operating-profit bridge, $m (favorable = positive)", { x: x0 + 0.25, y: y0 + 0.18, w: 6, h: 0.3, fontFace: FONT, fontSize: 12, color: C.ink2, margin: 0, isTextBox: true });
  const br = S.op_bridge, n = br.length;
  let run = 0; const bars = br.map(b => {
    if (b.kind !== "delta") { run = b.value; return { lo: 0, hi: b.value, b }; }
    const r = { lo: Math.min(run, run + b.value), hi: Math.max(run, run + b.value), b }; run += b.value; return r;
  });
  const lvMin = Math.min(...bars.filter(z => z.b.kind === "delta").map(z => z.lo), ...bars.filter(z => z.b.kind !== "delta").map(z => z.hi));
  const lvMax = Math.max(...bars.map(z => z.hi)), stp = niceStep(Math.max(1e-6, lvMax - lvMin));
  const vmin = lvMin > 0 ? Math.max(0, Math.floor((lvMin - 1.5 * (lvMax - lvMin)) / stp) * stp) : Math.floor(Math.min(0, ...bars.map(z => z.lo)) / stp) * stp;
  const vmax = Math.ceil(lvMax / stp) * stp + stp / 2;
  const px0 = x0 + 0.55, px1 = x0 + cw - 0.25, py0 = y0 + 0.75, py1 = y0 + ch - 0.95;
  const Y = v => py1 - (Math.max(v, vmin) - vmin) / (vmax - vmin) * (py1 - py0);
  const band = (px1 - px0) / n, bw = band * 0.62;
  s.addText(`axis starts at $${vmin}m`, { x: px0, y: py1 + 0.62, w: 2.5, h: 0.22, fontFace: FONT, fontSize: 9, italic: true, color: C.muted, margin: 0, isTextBox: true });
  bars.forEach((z, i) => {
    const x = px0 + i * band + (band - bw) / 2, tot = z.b.kind !== "delta";
    const col = tot ? C.violet : (z.b.value >= 0 ? C.good : C.crit);
    const top = Y(z.hi), h = Math.max(0.03, Y(z.lo) - top);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: top, w: bw, h, rectRadius: Math.min(0.06, h / 2), fill: { color: col }, line: { color: col, width: 0 } });
    const lab = tot ? "$" + f1(z.b.value) : (z.b.value >= 0 ? "+" : "−") + f1(z.b.value);
    s.addText(lab, { x: x - 0.25, y: top - 0.3, w: bw + 0.5, h: 0.26, fontFace: FONT, fontSize: 10.5, bold: tot, color: C.ink, align: "center", margin: 0, isTextBox: true });
    const SHORT = { start: "Budget", b_price: "Price", b_vol: "Volume", b_mix: "Mix", b_mat: "Material", b_yld: "Yield",
      b_fix: "Fixed mfg", b_abs: "Vol. & inv.", b_rd: "R&D", b_sga: "SG&A", end: "Actual" };
    const name = SHORT[z.b.step_id] || z.b.label;
    s.addText(name, { x: x - 0.15, y: py1 + 0.08, w: bw + 0.3, h: 0.3, fontFace: FONT, fontSize: 10.5, color: C.ink2, align: "center", valign: "top", margin: 0, isTextBox: true });
  });
  const rx = x0 + cw + 0.3, rw = W - M - rx;
  card(s, rx, y0, rw, ch);
  s.addText("Required variances (division, $m)", { x: rx + 0.25, y: y0 + 0.18, w: rw - 0.5, h: 0.3, fontFace: FONT, fontSize: 12, color: C.ink2, margin: 0, isTextBox: true });
  const rows = ["s_rev", "s_price", "s_vol", "s_mix", "s_mat", "s_yld", "s_opex", "s_op"].map(k => vreq[k]);
  s.addTable(rows.map(r => [
    { text: r.label, options: { color: C.ink, bold: r.variance_id === "s_rev" || r.variance_id === "s_op" } },
    { text: smoney(r.value).replace("m", ""), options: { align: "right", color: r.value >= 0 ? C.good : C.crit, bold: true } }]),
    { x: rx + 0.2, y: y0 + 0.6, w: rw - 0.4, colW: [rw - 1.6, 1.2], fontFace: FONT, fontSize: 12, rowH: 0.36,
      border: { type: "solid", color: C.line, pt: 0.5 }, fill: { color: C.card } });
  s.addText(`Revenue decomposition: ${vreq.reconciliation_revenue.assessment}  ·  Profit bridge: ${vreq.reconciliation_bridge.assessment}`,
    { x: rx + 0.25, y: y0 + ch - 0.7, w: rw - 0.5, h: 0.5, fontFace: FONT, fontSize: 11, bold: true, color: C.emerald, margin: 0, isTextBox: true });
  s.addNotes(N.s6_notes);
}

// ===== 7. Operations risk =====
{
  const s = pres.addSlide();
  base(s, 7, "Semiconductor operations · final test", N.s7_title);
  const pm = B.product_monthly, labs = [...new Set(pm.map(x => x.month))].map(m => new Date(m + "T00:00:00").toLocaleString("en-US", { month: "short" }));
  const x0 = M, y0 = 1.75, cw = W - 2 * M, ch = 2.85;
  card(s, x0, y0, cw, ch);
  s.addText("Demand ÷ good-unit capacity at full load, Base scenario (TIGHT ≥ 0.85, SHORT ≥ 1.00)", { x: x0 + 0.25, y: y0 + 0.15, w: 10, h: 0.3, fontFace: FONT, fontSize: 12, color: C.ink2, margin: 0, isTextBox: true });
  const lw = 1.9, gx = x0 + 0.25 + lw, cellW = (cw - 0.5 - lw) / 12, cellH = 0.56;
  const heat = r => r < 0.6 ? "10201A" : r < 0.7 ? "17402F" : r < 0.8 ? "1F6A4D" : r < 0.9 ? "2F9D72" : "8EE6BB";
  labs.forEach((l, i) => s.addText(l, { x: gx + i * cellW, y: y0 + 0.5, w: cellW, h: 0.25, fontFace: FONT, fontSize: 10, color: C.muted, align: "center", margin: 0, isTextBox: true }));
  PRODS.forEach(([id, name, col], ri) => {
    const y = y0 + 0.82 + ri * (cellH + 0.1);
    dot(s, x0 + 0.25, y + 0.2, col, 0.15);
    s.addText(name, { x: x0 + 0.5, y, w: lw - 0.3, h: cellH, fontFace: FONT, fontSize: 12.5, bold: true, color: C.ink, valign: "middle", margin: 0, isTextBox: true });
    pm.filter(x => x.product_id === id).forEach((x, i) => {
      const r = x.demand_capacity_ratio, bin = r >= 1 ? "SHORT" : r >= 0.85 ? "TIGHT" : "OK";
      const lineC = bin === "SHORT" ? C.crit : bin === "TIGHT" ? C.warn : heat(r);
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: gx + i * cellW + 0.03, y, w: cellW - 0.06, h: cellH, rectRadius: 0.08, fill: { color: heat(r) }, line: { color: lineC, width: bin === "OK" ? 0 : 1.25 } });
      s.addText([{ text: r.toFixed(2), options: { fontSize: 11, bold: true, breakLine: true } }, { text: bin, options: { fontSize: 7.5 } }],
        { x: gx + i * cellW, y, w: cellW, h: cellH, fontFace: FONT, color: r >= 0.9 ? "06140E" : C.ink, align: "center", valign: "middle", margin: 0, isTextBox: true });
    });
  });
  const tiles = N.tiles.map(([l, v, sub, col]) => [l, v, sub, C[col]]);
  const tw = (W - 2 * M - 0.9) / 4;
  tiles.forEach(([l, v, sub, col], i) => stat(s, M + i * (tw + 0.3), 4.85, tw, 1.85, l, v, sub, col));
  s.addNotes(N.s7_notes);
}

// ===== 8. Scenarios =====
{
  const s = pres.addSlide();
  const SC = ["Downside", "Base", "Upside"], Kx = sc => D.scenarios[sc].kpis;
  base(s, 8, "Scenario planning · one model, three cases", `Operating profit ranges from ${money(Kx("Downside").operating_profit)} to ${money(Kx("Upside").operating_profit)}`);
  card(s, M, 1.75, 7.6, 4.95);
  const labels = ["Revenue", "Gross profit", "Operating profit", "Free cash flow"];
  s.addChart(pres.charts.BAR, SC.map(sc => ({ name: sc, labels, values: [Kx(sc).revenue, Kx(sc).gross_profit, Kx(sc).operating_profit, Kx(sc).fcf].map(v => Math.round(v * 10) / 10) })),
    { x: M + 0.2, y: 1.9, w: 7.2, h: 4.65, barDir: "col", barGrouping: "clustered", chartColors: [C.coral, C.sage, C.lime], barGapWidthPct: 60, barOverlapPct: -8,
      showValue: true, dataLabelPosition: "outEnd", dataLabelColor: C.ink, dataLabelFontFace: FONT, dataLabelFontSize: 9.5, dataLabelFormatCode: "$#,##0",
      valAxisLabelFormatCode: "$#,##0", showTitle: true, title: `${CO.forecast_year}, $m`, titleColor: C.ink2, titleFontFace: FONT, titleFontSize: 12,
      showLegend: true, legendPos: "b", ...axis });
  const rx = M + 7.9, rw = W - M - rx;
  card(s, rx, 1.75, rw, 4.95);
  s.addText(`Key drivers vs ${CO.base_year} base year`, { x: rx + 0.25, y: 1.93, w: rw - 0.5, h: 0.3, fontFace: FONT, fontSize: 12, color: C.ink2, margin: 0, isTextBox: true });
  const dv = Object.fromEntries(S.scenario_drivers.map(x => [x.driver_id, x]));
  const fmt = (id, v) => id === "yld_adj" ? (v >= 0 ? "+" : "−") + Math.abs(v * 100).toFixed(1) + " pts" : ["util", "tax", "capex_pct"].includes(id) ? pct(v, 0) : spct(v);
  const hdr = ["Driver", ...SC].map((t, i) => ({ text: t, options: { bold: true, color: i ? [C.coral, C.sage, C.lime][i - 1] : C.muted, align: i ? "right" : "left" } }));
  const names = { dem_g: "Demand growth", asp_chg: "Selling price", yld_adj: "Yield change", util: "Utilization", mat_inf: "Material inflation", rd_g: "R&D growth", sga_g: "SG&A growth", capex_pct: "Capex % revenue", tax: "Tax rate" };
  s.addTable([hdr, ...Object.keys(names).map(id => [{ text: names[id], options: { color: C.ink } },
    ...SC.map(sc => ({ text: fmt(id, dv[id][sc]), options: { align: "right", color: sc === "Base" ? C.ink : C.ink2 } }))])],
    { x: rx + 0.2, y: 2.3, w: rw - 0.4, colW: [rw - 0.4 - 3 * 0.95, 0.95, 0.95, 0.95], fontFace: FONT, fontSize: 11.5, rowH: 0.4,
      border: { type: "solid", color: C.line, pt: 0.5 }, fill: { color: C.card } });
  s.addNotes(N.s8_notes);
}

// ===== 9. Capital decision =====
{
  const s = pres.addSlide();
  const SC = ["Downside", "Base", "Upside"];
  base(s, 9, N.s9_eyebrow, N.s9_title);
  const x0 = M, cw = 6.3;
  card(s, x0, 1.75, cw, 4.95);
  s.addText(`$${f1(cap.price)}m equipment + $${f1(cap.install)}m install · +${cap.add_cap.toFixed(2)}m units/month · +${(cap.yld_gain * 100).toFixed(0)} pts yield · ${cap.life}-year life`,
    { x: x0 + 0.25, y: 1.93, w: cw - 0.5, h: 0.5, fontFace: FONT, fontSize: 12, color: C.ink2, margin: 0, isTextBox: true });
  const pf_ = t => ({ text: t, options: { align: "center", bold: true, color: t === "PASS" || t === "YES" ? "06140E" : "FFFFFF", fill: { color: t === "PASS" || t === "YES" ? C.good : C.crit } } });
  const rows = [
    [{ text: "Criterion", options: { bold: true, color: C.muted } }, ...SC.map((t, i) => ({ text: t, options: { bold: true, align: "center", color: [C.coral, C.sage, C.lime][i] } }))],
    [{ text: `NPV at ${pct(cap.rate, 0)}`, options: { color: C.ink } }, ...SC.map(sc => ({ text: money(cs[sc].npv), options: { align: "center", color: C.ink } }))],
    [{ text: "NPV > 0", options: { color: C.ink2 } }, ...SC.map(sc => pf_(cs[sc].npv_test))],
    [{ text: "IRR", options: { color: C.ink } }, ...SC.map(sc => ({ text: typeof cs[sc].irr === "number" ? pct(cs[sc].irr) : "n.m.", options: { align: "center", color: C.ink } }))],
    [{ text: `IRR ≥ ${pct(cap.hurdle, 0)} hurdle`, options: { color: C.ink2 } }, ...SC.map(sc => pf_(cs[sc].irr_test))],
    [{ text: "Payback", options: { color: C.ink } }, ...SC.map(sc => ({ text: typeof cs[sc].payback_years === "number" ? cs[sc].payback_years.toFixed(2) + " yrs" : "none", options: { align: "center", color: C.ink } }))],
    [{ text: `Payback ≤ ${cap.max_pb.toFixed(0)} years`, options: { color: C.ink2 } }, ...SC.map(sc => pf_(cs[sc].payback_test))],
    [{ text: "Meets all criteria", options: { bold: true, color: C.ink } }, ...SC.map(sc => pf_(cs[sc].meets_all))],
  ];
  s.addTable(rows, { x: x0 + 0.2, y: 2.5, w: cw - 0.4, colW: [cw - 0.4 - 3 * 1.3, 1.3, 1.3, 1.3], fontFace: FONT, fontSize: 12, rowH: 0.47,
    border: { type: "solid", color: C.line, pt: 0.5 }, fill: { color: C.card } });
  const rx = x0 + cw + 0.3, rw = W - M - rx;
  card(s, rx, 1.75, rw, 4.95);
  const cfs = S.capex_cashflows;
  s.addChart(pres.charts.LINE, SC.map(sc => ({ name: sc, labels: [0, 1, 2, 3, 4, 5, 6].map(y => "Y" + y), values: cfs.filter(x => x.scenario === sc).map(x => Math.round(x.cumulative_fcf * 10) / 10) })),
    { x: rx + 0.2, y: 1.9, w: rw - 0.4, h: 4.65, chartColors: [C.coral, C.sage, C.lime], lineSize: 2.5, lineDataSymbol: "circle", lineDataSymbolSize: 6,
      valAxisLabelFormatCode: "$#,##0", showTitle: true, title: "Cumulative incremental free cash flow, $m (payback where it crosses zero)",
      titleColor: C.ink2, titleFontFace: FONT, titleFontSize: 11.5, showLegend: true, legendPos: "b", ...axis });
  s.addNotes(N.s9_notes);
}

// ===== 10. Recommendation =====
{
  const s = pres.addSlide(); s.background = { path: BG_T };
  s.addText("RECOMMENDATION", { x: M, y: 0.55, w: 8, h: 0.3, fontFace: FONT, fontSize: 12, bold: true, color: C.emerald, charSpacing: 3, margin: 0, isTextBox: true });
  s.addText(N.s10_title, { x: M, y: 0.9, w: 11.8, h: 1.0, fontFace: FONT, fontSize: 34, bold: true, color: C.ink, margin: 0, isTextBox: true });
  card(s, M, 2.15, W - 2 * M, 1.35, "1F7A5A");
  s.addText(cap.checkpoint.replace("Checkpoint: ", "").replace(/^./, c => c.toUpperCase()),
    { x: M + 0.35, y: 2.2, w: W - 2 * M - 0.7, h: 1.25, fontFace: FONT, fontSize: 18, bold: true, color: "FFFFFF", valign: "middle", margin: 0, isTextBox: true });
  const steps = N.s10_steps;
  const cw = (W - 2 * M - 0.3) / 2;
  steps.forEach(([h, t], i) => {
    const x = M + (i % 2) * (cw + 0.3), y = 3.8 + Math.floor(i / 2) * 1.35;
    card(s, x, y, cw, 1.15);
    dot(s, x + 0.28, y + 0.3, C.lime, 0.18);
    s.addText(h, { x: x + 0.65, y: y + 0.16, w: cw - 0.9, h: 0.4, fontFace: FONT, fontSize: 15, bold: true, color: C.ink, margin: 0, isTextBox: true });
    s.addText(t, { x: x + 0.65, y: y + 0.55, w: cw - 0.9, h: 0.5, fontFace: FONT, fontSize: 12, color: C.ink2, margin: 0, isTextBox: true });
  });
  s.addText(D.disclaimer,
    { x: M, y: 7.05, w: 10.5, h: 0.25, fontFace: FONT, fontSize: 9, color: C.muted, margin: 0, isTextBox: true });
  s.addText("10", { x: W - M - 0.6, y: 7.05, w: 0.6, h: 0.25, fontFace: FONT, fontSize: 9, color: C.muted, align: "right", margin: 0, isTextBox: true });
  s.addNotes(N.s10_notes);
}

const out = path.join(OUTDIR, D.deck_file);
pres.writeFile({ fileName: out }).then(f => console.log("Wrote", path.relative(ROOT, f)));
