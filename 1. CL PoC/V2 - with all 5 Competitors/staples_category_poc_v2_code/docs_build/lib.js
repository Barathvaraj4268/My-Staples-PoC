// Shared helpers for the two Word documents. All numbers come from outputs/final/summary.json
// and all figures from outputs/final/figures (both written by run_framework.py).
const fs = require("fs");
const path = require("path");
const d = require("docx");

const ROOT = path.resolve(__dirname, "..");
const FIG = path.join(ROOT, "outputs", "final", "figures");
const S = JSON.parse(fs.readFileSync(path.join(ROOT, "outputs", "final", "summary.json"), "utf8"));

const C = { ink: "1F1F1F", muted: "595959", accent: "B00000", rule: "D9D9D9", head: "F2F2F2",
  zone: { "CURATE": "008300", "VERTICAL EXTENSION": "2A78D6", "REVIEW": "B07800", "1P-CORE GAP": "E34948",
          "OFF-BRAND": "7F7F7F", "VERIFY": "4A3AA7" },
  tint: { "CURATE": "E6F2E6", "VERTICAL EXTENSION": "E8F1FB", "REVIEW": "FDF3DC", "1P-CORE GAP": "FBE7E7",
          "OFF-BRAND": "F0F0F0", "VERIFY": "ECEAF6" } };
const FONT = "Calibri";

function pngSize(file) {
  const b = fs.readFileSync(file);
  return { w: b.readUInt32BE(16), h: b.readUInt32BE(20) };
}

// --- inline markup: **bold**, *italic*
function runs(text, opts = {}) {
  const out = [];
  const re = /(\*\*[^*]+\*\*|\*[^*]+\*)/g;
  let last = 0, m;
  while ((m = re.exec(text)) !== null) {
    if (m.index > last) out.push(new d.TextRun({ text: text.slice(last, m.index), font: FONT, ...opts }));
    const t = m[0];
    if (t.startsWith("**")) out.push(new d.TextRun({ text: t.slice(2, -2), bold: true, font: FONT, ...opts }));
    else out.push(new d.TextRun({ text: t.slice(1, -1), italics: true, font: FONT, ...opts }));
    last = m.index + t.length;
  }
  if (last < text.length) out.push(new d.TextRun({ text: text.slice(last), font: FONT, ...opts }));
  return out;
}

const P = (text, o = {}) => new d.Paragraph({ children: runs(text, o.run || {}), spacing: { after: o.after ?? 120, line: o.line ?? 276 },
  alignment: o.align, keepNext: o.keepNext, indent: o.indent });
const H1 = (t, o = {}) => new d.Paragraph({ heading: d.HeadingLevel.HEADING_1, children: [new d.TextRun({ text: t, font: FONT })],
  pageBreakBefore: o.pageBreak ?? false, spacing: { before: 120, after: 160 } });
const H2 = (t) => new d.Paragraph({ heading: d.HeadingLevel.HEADING_2, children: [new d.TextRun({ text: t, font: FONT })],
  spacing: { before: 200, after: 100 }, keepNext: true });
const H3 = (t) => new d.Paragraph({ heading: d.HeadingLevel.HEADING_3, children: [new d.TextRun({ text: t, font: FONT })],
  spacing: { before: 160, after: 80 }, keepNext: true });
const B = (t, lvl = 0) => new d.Paragraph({ numbering: { reference: "bullets", level: lvl }, children: runs(t),
  spacing: { after: 70, line: 264 } });
const N = (t, ref = "numbers") => new d.Paragraph({ numbering: { reference: ref, level: 0 }, children: runs(t),
  spacing: { after: 80, line: 264 } });
const PB = () => new d.Paragraph({ children: [new d.PageBreak()] });

function caption(text, src) {
  return new d.Paragraph({ spacing: { before: 60, after: 200 }, children: [
    ...runs(text, { size: 18, color: C.muted, italics: false }),
    ...(src ? [new d.TextRun({ text: "  Source: " + src, size: 16, color: "8A8A8A", font: FONT })] : []) ] });
}

function figure(file, widthIn, cap, src) {
  const f = path.join(FIG, file);
  const { w, h } = pngSize(f);
  const wpx = Math.round(widthIn * 96);
  const hpx = Math.round(wpx * h / w);
  return [new d.Paragraph({ alignment: d.AlignmentType.CENTER, keepNext: true, spacing: { after: 40 },
      children: [new d.ImageRun({ type: "png", data: fs.readFileSync(f), transformation: { width: wpx, height: hpx },
        altText: { title: cap, description: cap, name: file } })] }),
    caption(cap, src || `run_framework.py → figures/${file}`)];
}

// --- tables (DXA widths; 1440 = 1 inch)
function cell(text, w, o = {}) {
  const paras = String(text).split("\n").map(t => new d.Paragraph({ spacing: { after: 20, line: 240 },
    alignment: o.align, children: runs(t, { size: o.size ?? 17, bold: o.bold, color: o.color }) }));
  return new d.TableCell({ width: { size: w, type: d.WidthType.DXA }, children: paras,
    shading: o.fill ? { type: d.ShadingType.CLEAR, color: "auto", fill: o.fill } : undefined,
    margins: { top: 50, bottom: 50, left: 80, right: 80 }, verticalAlign: d.VerticalAlign.CENTER });
}

function table(header, rows, widths, o = {}) {
  const total = widths.reduce((a, b) => a + b, 0);
  const border = { style: d.BorderStyle.SINGLE, size: 4, color: C.rule };
  const head = new d.TableRow({ tableHeader: true, children: header.map((h, i) =>
    cell(h, widths[i], { bold: true, fill: o.headFill || C.head, size: o.headSize ?? 17, align: o.align?.[i] })) });
  const body = rows.map((r, ri) => new d.TableRow({ cantSplit: true, children: r.map((v, i) =>
    cell(v ?? "", widths[i], { size: o.size ?? 17, align: o.align?.[i], fill: o.rowFill ? o.rowFill(ri, i) : undefined,
      bold: o.boldCol === i })) }));
  return new d.Table({ width: { size: total, type: d.WidthType.DXA }, columnWidths: widths, rows: [head, ...body],
    borders: { top: border, bottom: border, left: border, right: border, insideHorizontal: border, insideVertical: border } });
}

function callout(lines, fill = "F4F6F8", color = C.ink) {
  // shaded single-cell box for key messages
  const w = callout.width || 12960;
  return new d.Table({ width: { size: w, type: d.WidthType.DXA }, columnWidths: [w],
    borders: { top: { style: d.BorderStyle.NONE }, bottom: { style: d.BorderStyle.NONE }, left: { style: d.BorderStyle.NONE },
      right: { style: d.BorderStyle.NONE }, insideHorizontal: { style: d.BorderStyle.NONE }, insideVertical: { style: d.BorderStyle.NONE } },
    rows: [new d.TableRow({ children: [new d.TableCell({ width: { size: w, type: d.WidthType.DXA },
      shading: { type: d.ShadingType.CLEAR, color: "auto", fill },
      margins: { top: 120, bottom: 120, left: 180, right: 180 },
      children: lines.map(t => new d.Paragraph({ spacing: { after: 60, line: 264 }, children: runs(t, { color, size: 20 }) })) })] })] });
}

const spacer = (pt = 6) => new d.Paragraph({ spacing: { after: pt * 20 }, children: [] });

function styles() {
  return {
    default: { document: { run: { font: FONT, size: 20, color: C.ink } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, color: C.accent, font: FONT }, paragraph: { spacing: { before: 240, after: 160 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 25, bold: true, color: C.ink, font: FONT }, paragraph: { spacing: { before: 200, after: 100 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 21, bold: true, color: C.muted, font: FONT }, paragraph: { spacing: { before: 160, after: 80 }, outlineLevel: 2 } },
    ],
  };
}

function numbering() {
  return { config: [
    { reference: "bullets", levels: [
      { level: 0, format: d.LevelFormat.BULLET, text: "•", alignment: d.AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 360, hanging: 240 } } } },
      { level: 1, format: d.LevelFormat.BULLET, text: "–", alignment: d.AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 240 } } } } ] },
    ...["numbers", "numbers2", "numbers3", "numbers4"].map(ref => ({ reference: ref, levels: [
      { level: 0, format: d.LevelFormat.DECIMAL, text: "%1.", alignment: d.AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 360, hanging: 300 } } } } ] })) ] };
}

function footer(label) {
  return new d.Footer({ children: [new d.Paragraph({ alignment: d.AlignmentType.RIGHT, children: [
    new d.TextRun({ text: label + "   ·   page ", size: 16, color: "8A8A8A", font: FONT }),
    new d.TextRun({ children: [d.PageNumber.CURRENT], size: 16, color: "8A8A8A", font: FONT }) ] })] });
}

// --- number formatting
const pct = (x, dgt = 0) => (x == null ? "–" : (100 * x).toFixed(dgt) + "%");
const f0 = (x) => (x == null ? "–" : Math.round(x).toLocaleString("en-US"));
const f1 = (x) => (x == null ? "–" : Number(x).toFixed(1));
const f2 = (x) => (x == null ? "–" : Number(x).toFixed(2));
const SHORT = { OfficeDepot: "Office Depot", WestElm: "West Elm", Wayfair: "Wayfair", Amazon: "Amazon", Walmart: "Walmart" };
const clean = (s) => String(s).split(" (competitors")[0].split(" (e.g. ")[0];
const aisle = (s) => String(s).replace(/ > /g, " › ").replace("Competitor dept: ", "Competitors' ");

function dd(name) { return S.deepen_detail.find(x => clean(x.name) === name); }
function ddLine(name) {
  const x = dd(name);
  if (!x) return "";
  const parts = Object.entries(x.by_competitor).filter(([, v]) => v.ratio >= 2)
    .sort((a, b) => b[1].ratio - a[1].ratio).slice(0, 3)
    .map(([c, v]) => `${SHORT[c]} ${v.ratio.toFixed(1)}x` + (v.basis === "items" ? ` (${f0(v.competitor_size)} items)` : ` (${f0(v.competitor_size)} shelves)`));
  return `Staples ${f0(x.staples_items)} items; ${parts.join(", ")}`;
}

module.exports = { d, S, C, FIG, ROOT, P, H1, H2, H3, B, N, PB, caption, figure, table, cell, callout, spacer, styles,
  numbering, footer, pct, f0, f1, f2, SHORT, clean, aisle, dd, ddLine, runs, FONT };
