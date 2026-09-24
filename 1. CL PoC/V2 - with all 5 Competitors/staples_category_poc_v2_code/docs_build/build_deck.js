// End-to-end deck (pptxgenjs, 13.33" x 7.5"). Every chart is a Python figure; every number is from summary.json.
const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");

const ROOT = path.resolve(__dirname, "..");
const FIG = path.join(ROOT, "outputs", "final", "figures");
const S = JSON.parse(fs.readFileSync(path.join(ROOT, "outputs", "final", "summary.json"), "utf8"));
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.title = "Staples marketplace – category-level recommendations (PoC v2)";

const INK = "1F2328", MUTED = "5F6368", LIGHT = "F3F4F6", ACC = "B00000", DARK = "1F2328", WHITE = "FFFFFF";
const ZC = { "CURATE": "008300", "VERTICAL EXTENSION": "2A78D6", "REVIEW": "C98500", "1P-CORE GAP": "E34948", "OFF-BRAND": "7F7F7F", "VERIFY": "4A3AA7" };
const ZT = { "CURATE": "E6F2E6", "VERTICAL EXTENSION": "E8F1FB", "REVIEW": "FDF3DC", "1P-CORE GAP": "FBE7E7", "OFF-BRAND": "EFEFEF", "VERIFY": "ECEAF6" };
const F = "Calibri";
const FC = S.family_counts, RU = S.recommendable_units_by_zone, V = S.verification, cal = S.calibration;
const pct = (x) => Math.round(100 * x) + "%";
const f0 = (x) => Math.round(x).toLocaleString("en-US");
const SHORT = { OfficeDepot: "Office Depot", WestElm: "West Elm", Wayfair: "Wayfair", Amazon: "Amazon", Walmart: "Walmart" };

function png(file) { const b = fs.readFileSync(path.join(FIG, file)); return { w: b.readUInt32BE(16), h: b.readUInt32BE(20) }; }
function img(slide, file, x, y, w, h) {       // fit inside the box, keep aspect, centre
  const { w: pw, h: ph } = png(file);
  let iw = w, ih = w * ph / pw;
  if (ih > h) { ih = h; iw = h * pw / ph; }
  slide.addImage({ path: path.join(FIG, file), x: x + (w - iw) / 2, y: y + (h - ih) / 2, w: iw, h: ih });
  return { iw, ih };
}
function header(slide, kicker, title) {
  slide.addText(kicker.toUpperCase(), { x: 0.5, y: 0.28, w: 12.3, h: 0.3, fontFace: F, fontSize: 11, bold: true, color: ACC, charSpacing: 2, margin: 0, isTextBox: true });
  slide.addText(title, { x: 0.5, y: 0.55, w: 12.3, h: 0.75, fontFace: F, fontSize: 26, bold: true, color: INK, margin: 0, valign: "top", isTextBox: true });
}
function source(slide, txt) {
  slide.addText("Source: " + txt, { x: 0.5, y: 7.08, w: 12.3, h: 0.28, fontFace: F, fontSize: 9, color: "8A8A8A", margin: 0, isTextBox: true });
}
function bullets(slide, items, x, y, w, h, size = 13) {
  slide.addText(items.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < items.length - 1, paraSpaceAfter: 8 } })),
    { x, y, w, h, fontFace: F, fontSize: size, color: INK, valign: "top", margin: 2, isTextBox: true });
}
function chip(slide, lab, x, y, w = 2.3) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h: 0.36, rectRadius: 0.18, fill: { color: ZC[lab] }, line: { color: ZC[lab] } });
  slide.addText(lab, { x, y, w, h: 0.36, fontFace: F, fontSize: 11, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0, isTextBox: true });
}
function stat(slide, big, small, x, y, w, color = INK) {
  slide.addText(big, { x, y, w, h: 0.75, fontFace: F, fontSize: 36, bold: true, color, margin: 0, isTextBox: true });
  slide.addText(small, { x, y: y + 0.75, w, h: 0.6, fontFace: F, fontSize: 11, color: MUTED, margin: 0, valign: "top", isTextBox: true });
}
const newSlide = () => { const s = pres.addSlide(); s.background = { color: WHITE }; return s; };

// 1 ── title ──────────────────────────────────────────────────────────────────────────────────
{ const s = pres.addSlide(); s.background = { color: DARK };
  s.addText("STAPLES.COM MARKETPLACE · CATEGORY PoC v2", { x: 0.7, y: 1.4, w: 11.5, h: 0.4, fontFace: F, fontSize: 14, bold: true, color: "F28B82", charSpacing: 3, margin: 0, isTextBox: true });
  s.addText("Where Staples should open its marketplace", { x: 0.7, y: 1.9, w: 11.5, h: 1.3, fontFace: F, fontSize: 44, bold: true, color: WHITE, margin: 0, isTextBox: true });
  s.addText("Category-level recommendations from six navigation trees – what to open to sellers, what to keep in 1P, what to fix, what to pass", { x: 0.7, y: 3.25, w: 10.5, h: 0.9, fontFace: F, fontSize: 18, color: "D0D4DA", margin: 0, isTextBox: true });
  s.addText(`Prepared for Pat · 24 September 2026 · Evidence: Staples + Office Depot, West Elm, Wayfair, Amazon, Walmart · ${f0(S.competitor_shelves_compared)} competitor category pages · ${f0(S.labelled_rows.total)} labelled checks`, { x: 0.7, y: 6.3, w: 12, h: 0.5, fontFace: F, fontSize: 12, color: "AEB4BC", margin: 0, isTextBox: true });
}

// 2 ── the answer ─────────────────────────────────────────────────────────────────────────────
{ const s = newSlide(); header(s, "The answer", "Five calls for Staples, from one decision framework");
  const cards = [
    ["CURATE", "Open to curated sellers now", `${FC["CURATE"]} aisles · ${RU["CURATE"]} sub-categories`, "Lead with Furniture › Décor (bedding, bath & shower, candles, décor, vanities); then Sewing & Tailoring, Grounds & outdoor living, home & office tables"],
    ["VERTICAL EXTENSION", "Phase 2, via a served vertical", `${FC["VERTICAL EXTENSION"]} aisles · ${RU["VERTICAL EXTENSION"]} sub-categories`, "Garden & greenhouse, outdoor structures, bedroom/dorm furniture, vases & picnic ware, fitness, material handling"],
    ["1P-CORE GAP", "Keep in 1P – fix the range", `${FC["1P-CORE GAP"]} aisles · ${RU["1P-CORE GAP"]} sub-categories`, `Office chairs, breakroom appliances, storage, chair mats, binders, coffee. Plus ${FC["REVIEW"]} REVIEW aisles for merchants (lighting, tabletop)`],
    ["VERIFY", "Fix findability first", `${V.carried_under_other_name} of ${V.checked} "gaps" already sold`, "Candles & home fragrance, evaporative coolers, pull-out pantries, sneeze guards, camera lenses – rename and cross-list"],
    ["OFF-BRAND", "Pass", `${S.offbrand_groups.competitor_departments} competitor depts · ${RU["OFF-BRAND"]} sub-categories`, "Livestock & garden supplies, craft niches, individual sports, holiday-specific décor"],
  ];
  cards.forEach(([lab, h, st, body], i) => {
    const x = 0.5 + i * 2.5, y = 1.55;
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: 2.35, h: 5.3, fill: { color: ZT[lab] }, line: { color: ZT[lab] } });
    chip(s, lab === "VERTICAL EXTENSION" ? "VERTICAL EXTENSION" : lab, x + 0.12, y + 0.18, 2.11);
    s.addText(h, { x: x + 0.15, y: y + 0.7, w: 2.05, h: 0.8, fontFace: F, fontSize: 16, bold: true, color: INK, margin: 0, valign: "top", isTextBox: true });
    s.addText(st, { x: x + 0.15, y: y + 1.55, w: 2.05, h: 0.6, fontFace: F, fontSize: 12, bold: true, color: ZC[lab], margin: 0, valign: "top", isTextBox: true });
    s.addText(body, { x: x + 0.15, y: y + 2.25, w: 2.05, h: 2.9, fontFace: F, fontSize: 12, color: INK, margin: 0, valign: "top", isTextBox: true });
  });
  source(s, "run_framework.py → summary.json, sheet Aisle_Recommendations");
  s.addNotes("One matrix, six zones. Counts are Staples aisles (L1 › L2) and the sub-categories seen at 2+ competitors that sit under them.");
}

// 3 ── agenda ─────────────────────────────────────────────────────────────────────────────────
{ const s = newSlide(); header(s, "Contents", "How we got to the answer");
  const items = [["1", "Competitors finalised", "Who we compared against, and what each tree gives us"], ["2", "Approach and flow", "Two engines (vector DB and graph DB), calibration, verification"],
    ["3", "Framework", "Six zones on adjacency × cannibalisation; the scores"], ["4", "High-level recommendations", "The matrix at Staples-aisle level; strategic themes"],
    ["5", "Recommendations by zone", "CURATE · VERTICAL EXTENSION · REVIEW · 1P-CORE GAP · OFF-BRAND · VERIFY"], ["6", "Confidence and next steps", "Stress test, method agreement, what to do next"]];
  items.forEach(([n, h, b], i) => { const y = 1.6 + i * 0.88;
    s.addShape(pres.shapes.OVAL, { x: 0.6, y, w: 0.6, h: 0.6, fill: { color: i === 3 ? ACC : INK }, line: { color: i === 3 ? ACC : INK } });
    s.addText(n, { x: 0.6, y, w: 0.6, h: 0.6, fontFace: F, fontSize: 18, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0, isTextBox: true });
    s.addText(h, { x: 1.45, y: y - 0.02, w: 5, h: 0.4, fontFace: F, fontSize: 18, bold: true, color: INK, margin: 0, isTextBox: true });
    s.addText(b, { x: 1.45, y: y + 0.35, w: 10.5, h: 0.35, fontFace: F, fontSize: 13, color: MUTED, margin: 0, isTextBox: true }); });
}

// 4 ── competitors ───────────────────────────────────────────────────────────────────────────
{ const s = newSlide(); header(s, "1 · Competitors finalised", "Five competitors, three lenses – and very different data");
  img(s, "fig01_data_availability.png", 0.4, 1.4, 8.4, 3.2);
  const roles = [["Mirror", "Office Depot", "Same B2B catalogue; shows where Staples is shallow in its own core"], ["Lifestyle", "West Elm · Wayfair", "Design-led home and office; the 'white chair' lens"], ["Scale", "Amazon · Walmart", "Everything stores; confirm which gaps are mainstream"]];
  roles.forEach(([r, who, what], i) => { const x = 0.5 + i * 2.8, y = 4.8;
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: 2.6, h: 2.05, fill: { color: LIGHT }, line: { color: LIGHT } });
    s.addText(r, { x: x + 0.15, y: y + 0.12, w: 2.3, h: 0.4, fontFace: F, fontSize: 16, bold: true, color: ACC, margin: 0, isTextBox: true });
    s.addText(who, { x: x + 0.15, y: y + 0.52, w: 2.3, h: 0.35, fontFace: F, fontSize: 13, bold: true, color: INK, margin: 0, isTextBox: true });
    s.addText(what, { x: x + 0.15, y: y + 0.9, w: 2.3, h: 1.1, fontFace: F, fontSize: 12, color: MUTED, margin: 0, valign: "top", isTextBox: true }); });
  bullets(s, [`Item counts: full for Office Depot and West Elm; Walmart only at L1–L2; none for Wayfair and Amazon – so the method measures breadth (shelves) everywhere and items where both sides publish them`,
    `Depth varies from 2 to 5 levels. Embeddings use whatever path a node has, not fixed L1/L2/L3`,
    `One department universe for all six (media, apparel, grocery, toys, pets … out); merchandising pages and facets removed; menu headings kept in the tree but skipped as meaning`], 9.1, 1.5, 3.8, 5.4, 13);
  source(s, "figures/fig01_data_availability.png · sheets Data_Audit, Scope_Audit");
}

// 5 ── coverage ───────────────────────────────────────────────────────────────────────────────
{ const s = newSlide(); header(s, "1 · Competitors finalised", "Office Depot mirrors Staples; the whitespace is with the lifestyle and scale players");
  img(s, "fig04_staples_coverage.png", 0.4, 1.45, 8.8, 4.0);
  const st = [["OfficeDepot", "98%"], ["WestElm", ""], ["Wayfair", ""], ["Amazon", ""], ["Walmart", ""]];
  st.forEach(([c], i) => { const x = 0.5 + i * 1.75;
    stat(s, pct(1 - cal[c].base_rate_carried), `of ${SHORT[c]}'s shelves are missing at Staples`, x, 5.6, 1.65, i === 0 ? MUTED : ACC); });
  bullets(s, ["A random labelled sample of each competitor's shelves: how many exist at Staples at all", "Where Staples does have the shelf, it is often a token one – e.g. Bath & Shower Accessories: 8 items; Curtains, Blinds & Shades: 2", "So the lifestyle story is as much DEEPEN (too thin) as ENTER (missing)"], 9.4, 1.5, 3.5, 5.3, 13);
  source(s, "figures/fig04_staples_coverage.png · sheets Calibration, Node_Ensemble");
}

// 6 ── flow ─────────────────────────────────────────────────────────────────────────────────
{ const s = newSlide(); header(s, "2 · Approach and flow", "Seven steps from six shelf lists to one decision matrix");
  img(s, "fig02_flow.png", 0.4, 1.4, 12.5, 2.9);
  const boxes = [["Two independent engines", "Method V embeds every category with its path in a Qdrant vector DB. Method G builds one knowledge graph of all six trees. Scores are averaged, and disagreement lowers the evidence grade."],
    ["Calibrated, not guessed", `${f0(S.labelled_rows.calibration)} labelled rows set the 'same shelf' threshold per competitor. No hard-coded cosine; a new competitor runs on pooled calibration from day one.`],
    ["Verified before recommended", `${V.checked} model-flagged gaps were checked against Staples.com: ${V.carried_under_other_name} were already sold under another name. Those become findability actions, not assortment.`]];
  boxes.forEach(([h, b], i) => { const x = 0.5 + i * 4.15, y = 4.5;
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: 3.95, h: 2.35, fill: { color: LIGHT }, line: { color: LIGHT } });
    s.addText(h, { x: x + 0.2, y: y + 0.15, w: 3.6, h: 0.45, fontFace: F, fontSize: 16, bold: true, color: INK, margin: 0, isTextBox: true });
    s.addText(b, { x: x + 0.2, y: y + 0.65, w: 3.6, h: 1.6, fontFace: F, fontSize: 12.5, color: MUTED, margin: 0, valign: "top", isTextBox: true }); });
  source(s, "figures/fig02_flow.png · code: poc_common.py, method_v_vector.py, method_g_graph.py, run_framework.py");
}

// 7 ── two engines ─────────────────────────────────────────────────────────────────────────
{ const s = newSlide(); header(s, "2 · Approach and flow", "Vector DB and graph DB: two lenses on the same question");
  const col = (x, h, sub, items) => { s.addText(h, { x, y: 1.45, w: 4.1, h: 0.45, fontFace: F, fontSize: 18, bold: true, color: ACC, margin: 0, isTextBox: true });
    s.addText(sub, { x, y: 1.9, w: 4.1, h: 0.4, fontFace: F, fontSize: 12, italic: true, color: MUTED, margin: 0, isTextBox: true });
    bullets(s, items, x, 2.35, 4.1, 2.6, 12.5); };
  col(0.5, "Method V – vectors + Qdrant", "similarity: 'which Staples shelf means the same?'", ["Label + ancestors, weighted by distance (level-agnostic)", "Encoder chosen by bake-off: spaCy " + pct(S.encoder_bakeoff.spacy) + " top-1, WordLlama " + pct(S.encoder_bakeoff.wordllama) + ", TF-IDF " + pct(S.encoder_bakeoff.tfidf) + " (bge-small tested automatically in Colab)", "One Qdrant collection per retailer: fast hub-and-spoke queries, and peer search ('does Wayfair have this?')"]);
  col(4.75, "Method G – knowledge graph", "structure: 'where does it sit, and what surrounds it?'", ["All six trees + Staples cross-listings + SAME_AS links in one graph", "Similarity flooding: parents and children must match too (Office Depot top-1 " + pct(cal.OfficeDepot.top1_G_wording_only) + " → " + pct(cal.OfficeDepot.top1_G) + ")", "Personalised PageRank from Staples' shelves = adjacency; Neo4j export"]);
  img(s, "fig03_match_quality.png", 9.0, 1.45, 3.9, 3.4);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 5.2, w: 12.3, h: 1.6, fill: { color: LIGHT }, line: { color: LIGHT } });
  s.addText([{ text: "Why both? ", options: { bold: true } }, { text: `They fail differently. Vectors confuse shared words ("Drinking Glasses" vs "Drinking Fountains"); the graph is anchored by where a shelf sits. On new-category units the two place ${pct(S.method_agreement.same_zone_share)} in the same zone, and their adjacency scores correlate at ρ = ${S.method_agreement["Adjacency (AAS)"].toFixed(2)}. Where they disagree, the average places the unit and the evidence grade drops.` }],
    { x: 0.7, y: 5.3, w: 11.9, h: 1.4, fontFace: F, fontSize: 13, color: INK, margin: 0, valign: "middle", isTextBox: true });
  source(s, "figures/fig03_match_quality.png · outputs/V/encoder_bakeoff.csv · sheet Calibration");
}

// 8 ── verification ─────────────────────────────────────────────────────────────────────────
{ const s = newSlide(); header(s, "2 · Approach and flow", "Most 'missing' categories are not missing – they are hard to find");
  img(s, "fig05_verification_loop.png", 0.4, 1.5, 8.4, 3.6);
  stat(s, `${V.carried_under_other_name} of ${V.checked}`, "model-flagged gaps were already on Staples.com under another name", 9.2, 1.6, 3.7, ACC);
  stat(s, `${V.confirmed_gaps}`, "confirmed real gaps – these feed the recommendations", 9.2, 3.2, 3.7, INK);
  bullets(s, ["Patio furniture (1,116 items) lives under Gift Shop › Professional Gifts › 'Compasses'", "Sneeze guards are sold as 'Desktop Privacy Panels'", "Checks are stored as labels: they override both models on every re-run and become the VERIFY findability list"], 0.5, 5.25, 12.3, 1.7, 13);
  source(s, "figures/fig05_verification_loop.png · sheet Verification_Log");
}

// 9 ── framework ────────────────────────────────────────────────────────────────────────────
{ const s = newSlide(); header(s, "3 · Framework", "Two questions place every category in one of six zones");
  const Z = [["CURATE", "CRS < 40 · AAS ≥ 60", "Open to curated sellers now"], ["VERTICAL EXTENSION", "CRS < 40 · AAS 40–60", "Phase 2 through a served vertical"],
    ["OFF-BRAND", "CRS < 40 · AAS < 40", "Pass – real gap, wrong store"], ["REVIEW", "CRS 40–60", "Merchant decision"], ["1P-CORE GAP", "CRS ≥ 60", "Fix in 1P; keep out of marketplace"], ["VERIFY", "Flagged, may already exist", "Check first; fix findability"]];
  Z.forEach(([lab, rule, act], i) => { const x = 0.5 + (i % 3) * 2.75, y = 1.5 + Math.floor(i / 3) * 2.55;
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: 2.6, h: 2.35, fill: { color: ZT[lab] }, line: { color: ZT[lab] } });
    chip(s, lab, x + 0.12, y + 0.15, 2.36);
    s.addText(rule, { x: x + 0.15, y: y + 0.65, w: 2.3, h: 0.5, fontFace: F, fontSize: 13, bold: true, color: INK, margin: 0, isTextBox: true });
    s.addText(act, { x: x + 0.15, y: y + 1.2, w: 2.3, h: 1.0, fontFace: F, fontSize: 12.5, color: MUTED, margin: 0, valign: "top", isTextBox: true }); });
  s.addText("The scores", { x: 9.0, y: 1.5, w: 3.9, h: 0.4, fontFace: F, fontSize: 16, bold: true, color: ACC, margin: 0, isTextBox: true });
  bullets(s, ["AAS – adjacency (0–100): how embedded in what Staples sells; 100 = a typical Staples shelf", "CRS – cannibalisation risk (0–100): closeness to a Staples shelf × how core that shelf is to 1P", "PC – how many of 5 competitors show it · GS – how big the missing chunk is", "O = (0.35 PC + 0.35 GS + 0.30 AAS/100) × (1 − CRS/100)", "Grade A = 3+ competitors and both engines agree"], 9.0, 1.95, 3.9, 4.9, 12);
  source(s, "poc_common.label_row, opportunity · run_framework.label_and_rank");
}

// 10 ── matrix ──────────────────────────────────────────────────────────────────────────────
{ const s = newSlide(); header(s, "4 · High-level recommendations", "The decision matrix: Staples aisles by zone");
  img(s, "fig06_matrix_aisles.png", 0.3, 1.3, 9.6, 5.75);
  const rows = ["CURATE", "VERTICAL EXTENSION", "REVIEW", "1P-CORE GAP", "OFF-BRAND", "VERIFY"];
  rows.forEach((lab, i) => { const y = 1.45 + i * 0.9;
    chip(s, lab, 10.1, y, 2.8);
    s.addText(`${lab === "OFF-BRAND" ? S.offbrand_groups.competitor_departments + " depts + " + S.offbrand_groups.staples_aisles + " aisles" : FC[lab] + " aisles"} · ${RU[lab]} sub-categories`, { x: 10.1, y: y + 0.4, w: 2.9, h: 0.35, fontFace: F, fontSize: 11.5, color: INK, margin: 0, isTextBox: true }); });
  source(s, "figures/fig06_matrix_aisles.png · sheet Aisle_Recommendations (dot = Staples L1 › L2 aisle; size = ΣO)");
}

// 11 ── themes ──────────────────────────────────────────────────────────────────────────────
{ const s = newSlide(); header(s, "4 · High-level recommendations", "L1 view: home décor, creativity and outdoor living lead; the core stays 1P");
  img(s, "fig14_themes.png", 0.3, 1.4, 8.9, 4.2);
  const th = S.themes;
  bullets(s, [`Home décor & furnishings leads: ${th["Home décor & furnishings"].by_zone.CURATE} CURATE sub-categories, ΣO ${th["Home décor & furnishings"].sum_O_curate_ve.toFixed(1)} – the 'white chair' logic`,
    `Celebrations, gifting & creativity: ΣO ${th["Celebrations, gifting & creativity"].sum_O_curate_ve.toFixed(1)} – sewing, quilting, party décor`,
    `Outdoor living & grounds: ΣO ${th["Outdoor living & grounds"].sum_O_curate_ve.toFixed(1)} – planters, patio, garden, outdoor structures`, `Core office & business supplies: ${S.core_office_theme.one_p} of ${S.core_office_theme.total} sub-categories are 1P-CORE GAP`], 9.4, 1.5, 3.5, 5.3, 13);
  source(s, "figures/fig14_themes.png · sheet Themes");
}

// 19 ── graph view ─────────────────────────────────────────────────────────────────────────
{ const s = newSlide(); header(s, "4 · High-level recommendations", "Graph-DB view: where each recommendation plugs into Staples' aisles");
  img(s, "fig17_graph_dock.png", 0.3, 1.35, 8.6, 5.65);
  bullets(s, ["Path in the knowledge graph: competitor shelf → its aisle (CHILD_OF) → the Staples aisle it matches (SAME_AS) → Staples department", "The largest group of CURATE picks docks into Furniture › Décor – one aisle to open first, not a scatter of categories", "The same graph is exported to Neo4j with ready-made queries (consensus whitespace, aisle whitespace, cannibalisation)"], 9.2, 1.5, 3.7, 5.4, 13);
  source(s, "method_g_graph.py + run_framework.py → figures/fig17_graph_dock.png · outputs/G/neo4j");
}
// 19b ── vector view ─────────────────────────────────────────────────────────────────────────
{ const s = newSlide(); header(s, "4 · High-level recommendations", "Vector-DB view: CURATE sits at the edge of Staples' territory");
  img(s, "fig16_vector_map.png", 0.3, 1.35, 8.6, 5.65);
  bullets(s, ["Each dot is a category vector from the Qdrant collections, projected to 2-D (t-SNE)", "Grey: Staples' own aisles and shelves. Colour: recommendation zone", "CURATE and VERTICAL EXTENSION hug Staples' furniture, décor and kitchen clusters; OFF-BRAND forms separate islands (garden, sports, crafts)"], 9.2, 1.5, 3.7, 5.4, 13);
  source(s, "method_v_vector.py + run_framework.py → figures/fig16_vector_map.png");
}

// 12–17 ── zones ────────────────────────────────────────────────────────────────────────────
const zoneSlides = [
  ["CURATE", "fig08_zone_curate.png", "Open these aisles to curated marketplace sellers now",
    ["Furniture › Décor is the flagship: bedding (West Elm 4x, Wayfair 10x deeper), bath & shower (8 items at Staples vs 144 at West Elm), candles, shower curtains, décor accents, plus new shelves (vanities, daybeds, wreath hangers)",
     "Sewing & Tailoring: quilting, sewing machines, garment steamers (Amazon, Office Depot, Walmart)", "Grounds & outdoor living: planters, hedge trimmers, patio shelf ('Compasses': rename it)", "Home & office tables, commercial restaurant equipment, camera & TV accessories"]],
  ["VERTICAL EXTENSION", "fig09_zone_vertical_extension.png", "Phase 2: enter through a vertical Staples already serves",
    ["Garden & grounds: greenhouses, germination, tillers; outdoor structures (gazebos, railings, shutters) – via facilities and hospitality buyers", "Bedroom / dorm furniture: armoires, teen dressers, nightstands – via student living", "Vases, picnic baskets, candleholders, curtains, slipcovers: the next wave after Décor opens", "Fitness & sports sit at the low end of the zone: only via school or corporate wellness"]],
  ["REVIEW", "fig10_zone_review.png", "Merchant decisions: real gaps next to a 1P line",
    ["Lighting: ceiling and outdoor lighting deep at Wayfair/West Elm, but next to desk and facilities lighting", "Tabletop & bath textiles: split home vs foodservice", "Bathroom tiles & flooring: carried by all five competitors", "Heaters, air conditioners, bookcases, microwaves, TV stands: decide on margin"]],
  ["1P-CORE GAP", "fig11_zone_1pcore_gap.png", "Competitors are deeper in Staples' core – fix it in 1P",
    ["Office chairs (Office Depot 5.7x, Wayfair 5.8x Staples' relative depth), gaming and breakroom chairs, chair mats, partitions", "Breakroom appliances: blenders, toaster ovens, water filters (Office Depot 2–7x)", "Consumables: binders, notebooks, markers, coffee, trash bags", "A marketplace here would compete with Staples' own shelf"]],
  ["OFF-BRAND", "fig12_zone_offbrand.png", "Real gaps, wrong store: what we deliberately pass on",
    ["Craft niches: screen printing, purse making, embroidery", "Holiday-specific décor and seasonal novelties", "Individual sports: swimming, handball, airsoft", "Farm & garden supplies: hydroponics, beekeeping, poultry care, soils"]],
  ["VERIFY", "fig13_zone_verify.png", "Already sold, but hard to find: fix findability before buying assortment",
    ["Rename or cross-list 'Compasses' (patio furniture)", "Surface candles & home fragrance, evaporative coolers, pull-out pantries, quilts, sneeze guards under customers' words", "Health brand pages (watch list) map to thin OTC shelves (Ear & Eye Drops: 4 items)", "Checked, not guessed: " + V.checked + " flagged gaps reviewed; " + V.carried_under_other_name + " were already sold"]],
];
zoneSlides.forEach(([lab, fig, title, pts]) => {
  const s = newSlide(); header(s, "5 · Recommendations by zone", title);
  chip(s, lab, 0.5, 1.35, 2.8);
  s.addText(`${lab === "OFF-BRAND" ? S.offbrand_groups.competitor_departments + " competitor departments + " + S.offbrand_groups.staples_aisles + " Staples aisles" : FC[lab] + " Staples aisles"} · ${RU[lab]} sub-categories with 2+ competitor evidence`, { x: 3.45, y: 1.35, w: 9.3, h: 0.36, fontFace: F, fontSize: 13, color: MUTED, valign: "middle", margin: 0, isTextBox: true });
  const slug = { "CURATE": "curate", "VERTICAL EXTENSION": "vertical_extension", "REVIEW": "review", "1P-CORE GAP": "1pcore_gap", "OFF-BRAND": "offbrand", "VERIFY": "verify" }[lab];
  img(s, `slides/slide_zone_${slug}.png`, 0.3, 1.85, 8.9, 5.1);
  s.addShape(pres.shapes.RECTANGLE, { x: 9.4, y: 1.85, w: 3.5, h: 5.1, fill: { color: ZT[lab] }, line: { color: ZT[lab] } });
  s.addText("What it means for Staples", { x: 9.55, y: 1.95, w: 3.2, h: 0.4, fontFace: F, fontSize: 14, bold: true, color: ZC[lab], margin: 0, isTextBox: true });
  bullets(s, pts, 9.55, 2.4, 3.25, 4.45, 12);
  source(s, `figures/slides/slide_zone_${slug}.png (deck version of ${fig}) · sheets ${lab.replace(/ /g, "_").replace("-", "")}, Aisle_Recommendations`);
});

// 18 ── curate evidence ─────────────────────────────────────────────────────────────────────
{ const s = newSlide(); header(s, "5 · Recommendations by zone", "CURATE evidence: several competitors, and multiples of Staples' depth");
  img(s, "fig15_peer_consensus.png", 0.3, 1.35, 6.2, 5.65);
  img(s, "fig18_deepen_evidence.png", 6.7, 1.35, 6.3, 5.65);
  source(s, "figures/fig15_peer_consensus.png, fig18_deepen_evidence.png · sheets Concept_Members, Deepen_Evidence");
}

// 20 ── confidence ─────────────────────────────────────────────────────────────────────────
{ const s = newSlide(); header(s, "6 · Confidence", "Zones are robust; the order inside a zone is a shortlist");
  img(s, "fig19_sensitivity.png", 0.3, 1.35, 6.3, 4.4);
  img(s, "fig20_method_agreement.png", 6.8, 1.35, 6.2, 2.6);
  bullets(s, [`${S.sensitivity["top20_curate_p_approved_ge_0.8"]} of the top 20 CURATE sub-categories stay CURATE in ≥80% of 500 stress-test runs`, `Top-1 shelf match: Office Depot ${pct(cal.OfficeDepot.top1_V)}; lifestyle and scale competitors ${pct(cal.Wayfair.top1_V)}–${pct(cal.WestElm.top1_V)} – hence the verification loop`,
    "Not yet included: Staples sales and margin (CRS uses an editable 'coreness' assumption), demand signals, merchant-confirmed labels"], 6.8, 4.1, 6.1, 2.9, 12.5);
  source(s, "figures/fig19_sensitivity.png, fig20_method_agreement.png");
}

// 21 ── next steps ─────────────────────────────────────────────────────────────────────────
{ const s = pres.addSlide(); s.background = { color: DARK };
  s.addText("NEXT STEPS", { x: 0.7, y: 0.6, w: 11, h: 0.4, fontFace: F, fontSize: 14, bold: true, color: "F28B82", charSpacing: 3, margin: 0, isTextBox: true });
  s.addText("From recommendation to launch", { x: 0.7, y: 1.0, w: 11, h: 0.8, fontFace: F, fontSize: 34, bold: true, color: WHITE, margin: 0, isTextBox: true });
  const steps = [["Merchant session", "Walk CURATE and REVIEW with category merchants; their verdicts flow back into the model as labels"],
    ["Sales bands", "Replace the coreness assumption with 1P revenue and margin by aisle"], ["Demand", "Add search volume for the top 40 sub-categories (optional term in O)"],
    ["Archetype PoC", "Décor (bedding, bath, candles), Sewing & Tailoring, outdoor living: SKU and attribute comparison"], ["Findability fixes", "Rename 'Compasses', cross-list VERIFY items, add search synonyms – no new assortment needed"],
    ["Re-run quarterly", "About 6 minutes end to end; a new competitor = 1 adapter + about 120 labelled rows"]];
  steps.forEach(([h, b], i) => { const x = 0.7 + (i % 3) * 4.05, y = 2.3 + Math.floor(i / 3) * 2.3;
    s.addShape(pres.shapes.OVAL, { x, y, w: 0.55, h: 0.55, fill: { color: ACC }, line: { color: ACC } });
    s.addText(String(i + 1), { x, y, w: 0.55, h: 0.55, fontFace: F, fontSize: 16, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0, isTextBox: true });
    s.addText(h, { x: x + 0.7, y: y + 0.02, w: 3.1, h: 0.45, fontFace: F, fontSize: 17, bold: true, color: WHITE, margin: 0, isTextBox: true });
    s.addText(b, { x: x + 0.7, y: y + 0.5, w: 3.1, h: 1.4, fontFace: F, fontSize: 12.5, color: "D0D4DA", margin: 0, valign: "top", isTextBox: true }); });
}

// 22 ── appendix: unit matrix ───────────────────────────────────────────────────────────────
{ const s = newSlide(); header(s, "Appendix", "Detail: every sub-category unit on the matrix");
  img(s, "fig07_matrix_units.png", 0.3, 1.3, 12.7, 5.7);
  source(s, "figures/fig07_matrix_units.png · sheet All_Units");
}

const out = path.join(ROOT, "deliver", "Staples_Category_PoC_v2.pptx");
pres.writeFile({ fileName: out }).then(() => console.log("wrote", out));
