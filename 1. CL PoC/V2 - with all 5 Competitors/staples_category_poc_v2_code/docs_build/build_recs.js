// Recommendations document for Staples (Pat). Landscape, figure-led. Numbers: summary.json.
const fs = require("fs");
const path = require("path");
const L = require("./lib");
const { d, S, C, P, H1, H2, H3, B, N, PB, figure, table, callout, spacer, pct, f0, f1, f2, SHORT, clean, aisle, ddLine } = L;

const W = 13824;                 // content width (DXA) on landscape Letter with 0.7" margins
L.callout.width = W;
const FW = 9.4;                  // figure width (inches)
const Z = S.zones, FC = S.family_counts, RU = S.recommendable_units_by_zone;
const V = S.verification, U = S.units;
const cal = S.calibration;

const zoneFam = (lab) => S.families[lab] || [];
const fam = (lab, name) => zoneFam(lab).find(r => r.family_display === name) || {};
const famRows = (lab, n = 10) => zoneFam(lab).slice(0, n).map((r, i) => [
  String(i + 1), aisle(r.family_display),
  (r.examples_list || [r.examples]).join("; ") + (r.n_units > 5 ? ` (+${r.n_units - 5} more)` : ""),
  `${r.k_max}/5`, r.competitors, f0(r.AAS), f0(r.CRS), f2(r.O_sum), r.best_evidence]);
const famTable = (lab, n = 10) => table(
  ["#", "Staples aisle (L1 › L2)", "Sub-categories to add or deepen (L2–L3)", "Peers (best sub-category)", "All competitors showing gaps in this aisle", "AAS", "CRS", "ΣO", "Grade"],
  famRows(lab, n), [400, 2600, 4424, 1100, 2600, 600, 600, 700, 800],
  { align: [d.AlignmentType.CENTER, null, null, d.AlignmentType.CENTER, null, d.AlignmentType.CENTER, d.AlignmentType.CENTER, d.AlignmentType.CENTER, d.AlignmentType.CENTER],
    rowFill: (ri, ci) => (ci === 1 ? C.tint[lab] : undefined) });

function zoneHeader(num, lab, title, decision, stats) {
  return [H1(`${num}. ${title}`, { pageBreak: true }),
    callout([`**Decision:** ${decision}`, stats], C.tint[lab], C.ink), spacer(4)];
}

const cur = zoneFam("CURATE");
const top = (lab) => (S.top_units[lab] || []);

// ------------------------------------------------------------------------------------------------
const cover = [
  new d.Paragraph({ spacing: { before: 1400, after: 200 }, children: [new d.TextRun({ text: "Staples.com Marketplace", size: 30, color: C.muted, font: L.FONT })] }),
  new d.Paragraph({ spacing: { after: 200 }, children: [new d.TextRun({ text: "Where to open the marketplace: category-level recommendations", size: 52, bold: true, color: C.ink, font: L.FONT })] }),
  new d.Paragraph({ spacing: { after: 500 }, children: [new d.TextRun({ text: "Which categories to open to marketplace sellers, which to keep in 1P, which to fix, and which to pass", size: 28, color: C.muted, font: L.FONT })] }),
  P(`**Prepared for:** Pat, Staples   ·   **Scope:** category level (L1–L3); product archetypes are a separate PoC   ·   **Date:** 24 September 2026`),
  P(`**Evidence base:** the navigation trees of Staples and five competitors – Office Depot, West Elm, Wayfair, Amazon and Walmart. That covers ${f0(S.competitor_shelves_compared)} competitor category pages, matched to Staples by two independent methods and checked against ${f0(S.labelled_rows.total)} labelled examples. Every chart and number in this document is produced by the delivered Python code and can be traced to *Category_Recommendations_v2.xlsx*.`),
];

// ------------------------------------------------------------------------------------------------
const exec = [
  H1("1. The answer on one page", { pageBreak: true }),
  ...figure("fig06_matrix_aisles.png", 8.2, "Figure 6 – The decision matrix at category level. Each dot is a Staples aisle (L1 › L2) holding the sub-categories that fall in that zone; dot size = total opportunity.", "run_framework.py → figures/fig06_matrix_aisles.png · sheet Aisle_Recommendations"),
  PB(),
  H2("Five calls for Staples"),
  N(`**Open now to curated marketplace sellers (CURATE) – ${FC["CURATE"]} aisles, ${RU["CURATE"]} sub-categories.** Lead with **Furniture › Décor**: bedding, bath and shower accessories, candles, décor accents, vanities and daybeds. Four of the five competitors carry far more here. Staples has ${ddLine("Bath & Shower Accessories").split(";")[0].replace("Staples ", "")} on its Bath & Shower Accessories shelf; West Elm alone has 20x that, scaled to catalogue size. Next come **Sewing & Tailoring** (quilting, sewing machines, garment steamers), **Grounds Maintenance** (outdoor planters, hedge trimmers, parts), **Commercial Restaurant Equipment** (home brewing, food trailers) and **Home & Office Tables** (buffets and sideboards, kids' tables).`),
  N(`**Enter in phase 2, through a vertical Staples already serves (VERTICAL EXTENSION) – ${FC["VERTICAL EXTENSION"]} aisles, ${RU["VERTICAL EXTENSION"]} sub-categories.** This covers garden and greenhouse equipment, outdoor structures (gazebo canopies, railings, shutters), bedroom furniture (armoires, teen dressers, nightstands), vases and picnic ware, fitness and sports equipment, and material handling. Each is a real gap, but it sits one step away from Staples' current neighbourhood, so it needs an anchor audience: education, hospitality or facilities.`),
  N(`**Keep in 1P and fix the assortment there (1P-CORE GAP) – ${FC["1P-CORE GAP"]} aisles, ${RU["1P-CORE GAP"]} sub-categories.** Competitors are deeper in office chairs, breakroom appliances, storage and organisation, chair mats, partitions, file cabinets, binders and coffee. These are Staples' own core lines, and a marketplace there would compete with Staples' own shelf. ${FC["REVIEW"]} more aisles, including lighting, tabletop, bath mats and towels, and bookcases, sit in **REVIEW** for a merchant decision.`),
  N(`**Fix findability before buying assortment (VERIFY) – ${FC["VERIFY"]} aisles, ${RU["VERIFY"]} sub-categories.** In the verification check, **${V.carried_under_other_name} of the ${V.checked}** categories the models first flagged as "missing" turned out to be already on Staples.com under other names. Examples are candles and home fragrance, evaporative coolers, pull-out pantries, sneeze guards and camera lenses. Customers who search competitor vocabulary may not find them.`),
  N(`**Pass (OFF-BRAND) – ${S.offbrand_groups.competitor_departments} competitor departments (plus ${S.offbrand_groups.staples_aisles} Staples aisles), ${RU["OFF-BRAND"]} sub-categories.** These are real gaps in the wrong store for Staples: livestock and garden supplies, craft niches such as screen printing, purse making and embroidery, sports equipment such as swimming and handball, and holiday-specific décor.`),
  spacer(4),
  callout([`**Where the opportunity is, by theme (Figure 14):** Home décor & furnishings leads: it has the most CURATE sub-categories (${S.themes["Home décor & furnishings"].by_zone.CURATE}) and the highest summed opportunity across CURATE and VERTICAL EXTENSION (${f1(S.themes["Home décor & furnishings"].sum_O_curate_ve)}). Celebrations, gifting & creativity (${f1(S.themes["Celebrations, gifting & creativity"].sum_O_curate_ve)}) and Outdoor living & grounds (${f1(S.themes["Outdoor living & grounds"].sum_O_curate_ve)}) follow. This is the "white chair" logic in data: design and lifestyle versions of things Staples already sells to the same customer, at low risk to its 1P core.`,
    `**How sure are we?** Zone placements are stable: ${S.sensitivity.top20_curate_p_approved_ge_0_8 ?? S.sensitivity["top20_curate_p_approved_ge_0.8"]} of the top 20 CURATE sub-categories stay CURATE in at least 80% of 500 stress-test runs. The exact ranking inside the zone is less stable, so treat the top of each list as a shortlist, not a fixed order. Every recommendation also carries an evidence grade: A means 3 or more competitors and both methods agree.`], "F4F6F8"),
];

// ------------------------------------------------------------------------------------------------
const panel = [
  H1("2. What the five competitors tell us", { pageBreak: true }),
  P(`Each competitor plays a different role in the evidence. The same method is applied to all of them, and it adapts to what each navigation tree provides. Office Depot and West Elm publish item counts, so depth can be measured in items. Walmart publishes counts only at the top two levels. Wayfair and Amazon publish none, so for them we measure breadth: how many distinct shelves a competitor devotes to a category.`),
  ...figure("fig01_data_availability.png", 8.6, "Figure 1 – The competitor panel and what each navigation tree gives us.", "run_framework.py → figures/fig01_data_availability.png · sheet Data_Audit"),
  B(`**Office Depot is a mirror.** ${pct(cal.OfficeDepot.base_rate_carried)} of its shelves, drawn at random, already exist at Staples. It tells us where Staples is shallower in its own core (the 1P-CORE GAP list), but it offers almost no whitespace.`),
  B(`**West Elm and Wayfair are the lifestyle lens.** ${pct(1 - cal.WestElm.base_rate_carried)} of West Elm's shelves and ${pct(1 - cal.Wayfair.base_rate_carried)} of Wayfair's are missing at Staples. The bigger story is depth: where Staples does have the shelf, it is often a token one. Bath & Shower Accessories has 8 items and Curtains, Blinds & Shades has 2.`),
  B(`**Amazon and Walmart are the scale lens.** ${pct(1 - cal.Amazon.base_rate_carried)} of Amazon's and ${pct(1 - cal.Walmart.base_rate_carried)} of Walmart's in-scope shelves are missing at Staples. Their breadth confirms which lifestyle gaps are mainstream demand rather than one retailer's taste.`),
  ...figure("fig04_staples_coverage.png", 8.0, "Figure 4 – How much of each competitor's shelf list Staples already covers.", "run_framework.py → figures/fig04_staples_coverage.png · sheets Node_Ensemble, Calibration"),
];

// ------------------------------------------------------------------------------------------------
const zonesExplained = [
  H1("3. How to read the matrix", { pageBreak: true }),
  P(`Think of Staples.com as a store map. For every category a competitor sells, we ask three questions. **Is it missing at Staples, or too thin?** **Is it close to what Staples already sells** – the adjacency score, AAS? **Would a marketplace there take sales from Staples' own 1P shelf** – the cannibalisation risk score, CRS? The two scores place every category in one of six zones.`),
  table(["Zone", "Rule", "What it means for Staples", "Action"], [
    ["CURATE", "CRS < 40 and AAS ≥ 60", "A gap right next to what Staples sells, with little 1P at stake", "Open to curated marketplace sellers now"],
    ["VERTICAL EXTENSION", "CRS < 40 and AAS 40–60", "A real gap one step further out", "Phase 2: enter through a vertical Staples serves (education, hospitality, facilities)"],
    ["REVIEW", "CRS 40–60", "Sits next to a 1P line; could go either way", "Merchant decision with sales and margin data"],
    ["1P-CORE GAP", "CRS ≥ 60", "Competitors are deeper in a Staples core line", "Fix in 1P merchandising; keep out of the marketplace"],
    ["OFF-BRAND", "CRS < 40 and AAS < 40", "A real gap, far from Staples' customers", "Pass"],
    ["VERIFY", "Flagged as missing, but may exist under another name", "Often a findability problem, not an assortment gap", "Check first; fix naming and navigation"],
  ], [2300, 2600, 4800, 4124], { size: 16, rowFill: (ri, ci) => (ci === 0 ? Object.values(C.tint)[ri] : undefined), boldCol: 0 }),
  spacer(6),
  P(`**Two kinds of recommendation.** "New" means the category does not exist at Staples (ENTER). "Deepen" means Staples has the aisle, but at least two competitors carry at least twice as much of it, relative to their size (DEEPEN). Recommendations are shown at the level Staples would act on: a Staples aisle (L1 › L2), with the sub-categories (L2–L3) to add underneath it. Only sub-categories seen at two or more competitors drive an aisle-level call. The ${S.watchlist_single_competitor} single-competitor sub-categories are kept as a watch list in the workbook.`),
  H2("The scores behind every recommendation"),
  table(["Score", "Plain-language meaning", "How it is built (both methods averaged)"], [
    ["AAS – adjacency (0–100)", "How embedded the category is in what Staples already sells. 100 = as embedded as a typical Staples shelf", "Vector method: closeness to the 10 nearest Staples shelves. Graph method: network proximity to Staples' assortment, plus how much of the competitor's aisle Staples already carries"],
    ["CRS – cannibalisation risk (0–100)", "How much 1P revenue a marketplace here would put at risk", "Closeness to a Staples shelf × how core that shelf is to 1P (office supplies, paper and ink score high; décor and lifestyle score low)"],
    ["PC – peer consensus", "How many of the five competitors show the gap", "Share of competitors that carry the category (new) or are at least 2x deeper (deepen)"],
    ["GS – gap size", "How big the missing assortment is", "Percentile of the missing chunk within each competitor's catalogue: shelves everywhere, plus items where counts exist"],
    ["O – opportunity (0–1)", "The ranking inside a zone", "O = (0.35 PC + 0.35 GS + 0.30 AAS/100) × (1 − CRS/100)"],
    ["Grade A / B / C", "Strength of evidence", "A = 3+ competitors and both methods agree on the zone; B = 2+ competitors or the methods agree; C = neither"],
  ], [3000, 4600, 6224], { size: 16 }),
  spacer(4),
];

// ------------------------------------------------------------------------------------------------
const bb = (n) => ddLine(n);
const curate = [
  ...zoneHeader(4, "CURATE", "CURATE – open these to marketplace sellers now",
    "curate a marketplace assortment in these aisles. Each is next to what Staples already sells, and Staples' own 1P revenue there is small.",
    `${FC["CURATE"]} aisles · ${RU["CURATE"]} sub-categories seen at 2+ competitors · median AAS ${f0(Z["CURATE"].median_AAS)}, median CRS ${f0(Z["CURATE"].median_CRS)}`),
  ...figure("fig08_zone_curate.png", FW, "Figure 8 – CURATE aisles ranked by total opportunity, with the sub-categories behind each.", "run_framework.py → figures/fig08_zone_curate.png · sheet CURATE, Aisle_Recommendations"),
  famTable("CURATE", 10),
  H2("What this means for Staples"),
  B(`**Make Furniture › Décor the flagship curated aisle.** Staples already has the aisle, but most of its sub-shelves are token. Beds & Bedding: ${bb("Beds & Bedding")}. Bath & Shower Accessories: ${bb("Bath & Shower Accessories")}. Candles: ${bb("Candles")}. Shower Curtains: ${bb("Shower Curtains & Accessories")}. New shelves competitors carry and Staples does not: makeup vanities, daybeds, wreath hangers, teen tapestries and wall hangings.`),
  B(`**Sewing & Tailoring is a surprise candidate with strong evidence** (Amazon, Office Depot, Walmart). Quilting, sewing tools, garment steamers and irons, and sewing machines extend Staples' Retail Store Supplies › Sewing & Tailoring aisle to a home-crafting customer.`),
  B(`**Outdoor living:** outdoor planters, hedge trimmers and outdoor tool parts extend Facilities › Grounds Maintenance. Staples files its patio and outdoor shelves under Gift Shop › Professional Gifts › "Compasses" – 1,116 patio-furniture items sit on that path (${bb("Compasses")}). That label hides the category from shoppers, so rename it when opening it to sellers.`),
  B(`**Home & hospitality furniture:** buffet tables and sideboards (Walmart, Wayfair, West Elm) and classroom and kids' tables extend Furniture › Home & Office Tables. Home brewing and food trailers extend Commercial Restaurant Equipment.`),
  B(`**Technology accessories with low core risk:** camera and camcorder accessories, digital picture frames, TV accessories, DVD and Blu-ray players, and cell phones, where Office Depot and Walmart are several times deeper.`),
  PB(),
  H2("The peer and depth evidence behind the top picks"),
  ...figure("fig15_peer_consensus.png", 6.9, "Figure 15 – Peer evidence behind the top recommendations: which competitors carry each gap (circles) or are at least 2x deeper (squares).", "run_framework.py → figures/fig15_peer_consensus.png · sheets Concept_Members, Deepen_Evidence"),
  ...figure("fig18_deepen_evidence.png", 7.6, "Figure 18 – DEEPEN evidence: each competitor's depth in a Staples aisle relative to Staples (size-factor adjusted; ⁱ = item counts).", "run_framework.py → figures/fig18_deepen_evidence.png · sheet Deepen_Evidence"),
];

const ve = [
  ...zoneHeader(5, "VERTICAL EXTENSION", "VERTICAL EXTENSION – phase 2, through a vertical Staples already serves",
    "do not open these to general marketplace traffic yet. Enter them through an audience Staples already serves – education, hospitality, facilities or home office – once the CURATE aisles prove the seller model.",
    `${FC["VERTICAL EXTENSION"]} aisles · ${RU["VERTICAL EXTENSION"]} sub-categories · median AAS ${f0(Z["VERTICAL EXTENSION"].median_AAS)}, median CRS ${f0(Z["VERTICAL EXTENSION"].median_CRS)}`),
  ...figure("fig09_zone_vertical_extension.png", FW, "Figure 9 – VERTICAL EXTENSION aisles ranked by total opportunity.", "run_framework.py → figures/fig09_zone_vertical_extension.png · sheet VERTICAL_EXTENSION"),
  famTable("VERTICAL EXTENSION", 10),
  H2("What this means for Staples"),
  B(`**Garden, grounds and outdoor structures** – greenhouses, plant germination, soil monitoring and tillers extend Facilities › Grounds Maintenance; gazebo canopies, weathervanes, bird houses, porch railings and shutters sit next to Staples' patio shelves. Enter through facilities and hospitality buyers before consumers.`),
  B(`**Bedroom and dorm furniture** – armoires and wardrobes, teen dressers and nightstands (four competitors), plus bed canopies. It is a step beyond the home office; the natural route is a dorm or student-living vertical tied to School Supplies.`),
  B(`**Décor accessories** – vases, picnic baskets, candleholders, curtains and blinds, slipcovers and tapestries. Once Furniture › Décor is open, these become the next wave.`),
  B(`**Fitness and sports** – fitness equipment, golf, fishing and boating sit at the low end of this zone (aisle AAS ${f0(fam("VERTICAL EXTENSION", "Fitness › Fitness Equipment").AAS)}). Pursue them only through a school athletics or corporate-wellness programme.`),
  B(`**Material handling** – loading dock levelers and forklift booms fit a B2B facilities vertical, not a consumer marketplace.`),
];

const review = [
  ...zoneHeader(6, "REVIEW", "REVIEW – a merchant decision, next to a 1P line",
    "put these in front of the category merchants with sales and margin data. The gap is real, but each sits next to a Staples 1P shelf.",
    `${FC["REVIEW"]} aisles · ${RU["REVIEW"]} sub-categories · median AAS ${f0(Z["REVIEW"].median_AAS)}, median CRS ${f0(Z["REVIEW"].median_CRS)}`),
  ...figure("fig10_zone_review.png", FW, "Figure 10 – REVIEW aisles ranked by total opportunity.", "run_framework.py → figures/fig10_zone_review.png · sheet REVIEW"),
  famTable("REVIEW", 10),
  H2("What this means for Staples"),
  B(`**Lighting.** Competitors are much deeper in ceiling and outdoor lighting (Ceiling Lighting: ${bb("Ceiling Lighting")}). These shelves sit next to Staples' desk lamps and facilities lighting, which are 1P lines. The likely call is a curated décor-lighting assortment that stays clear of task and commercial lighting.`),
  B(`**Tabletop and bath textiles.** Dinnerware (${bb("Dinnerware")}), flatware, and bath mats and towels (${bb("Bath Mats & Towels")}) overlap Staples' breakroom and foodservice lines. Split them: foodservice stays 1P, and home tabletop goes to the marketplace.`),
  B(`**Bathroom tiles and flooring** is carried by all five competitors. With fountain accessories, it is one of the ${S.zones["REVIEW"].enter} new-category REVIEW items; it neighbours Staples' facilities plumbing and flooring lines.`),
  B(`**Heaters, air conditioners, bookcases, microwaves and TV stands** are seasonal or bulky, sit next to 1P, and Office Depot is deeper in them. Decide on margin, not on the gap.`),
];

const core = [
  ...zoneHeader(7, "1P-CORE GAP", "1P-CORE GAP – fix in 1P, keep out of the marketplace",
    "competitors are deeper in lines Staples considers core. Close the gap with 1P assortment and vendor programmes. A marketplace here would compete with Staples' own shelf.",
    `${FC["1P-CORE GAP"]} aisles · ${RU["1P-CORE GAP"]} sub-categories · median AAS ${f0(Z["1P-CORE GAP"].median_AAS)}, median CRS ${f0(Z["1P-CORE GAP"].median_CRS)}`),
  ...figure("fig11_zone_1pcore_gap.png", FW, "Figure 11 – 1P-CORE GAP aisles ranked by total opportunity.", "run_framework.py → figures/fig11_zone_1pcore_gap.png · sheet 1PCORE_GAP"),
  famTable("1P-CORE GAP", 10),
  H2("What this means for Staples"),
  B(`**Seating and furniture core.** Office Chairs: ${bb("Office Chairs")}. Also gaming chairs, breakroom and dining chairs, chair mats (${bb("Chair Mats")}), partitions, shelving and file cabinets. In these core lines competitors run deeper ranges – mostly Office Depot and Wayfair – and the fix is 1P range extension.`),
  B(`**Breakroom appliances** – toaster ovens, blenders and food processors, water filters and dispensers. Office Depot is several times deeper (Blenders & Food Processors: ${bb("Blenders & Food Processors")}).`),
  B(`**Office consumables** – binders, notebooks, journals, markers, coffee and trash bags (Coffee: ${bb("Coffee")}). Depth gaps here are a 1P merchandising brief, not a marketplace one.`),
];

const offb = [
  ...zoneHeader(8, "OFF-BRAND", "OFF-BRAND – real gaps, wrong store: pass",
    "do not pursue these. Competitors sell them, but they sit far from Staples' customers and neighbourhood (AAS below 40). They would dilute the Staples brand.",
    `${S.offbrand_groups.competitor_departments} competitor departments + ${S.offbrand_groups.staples_aisles} Staples aisles · ${RU["OFF-BRAND"]} sub-categories seen at 2+ competitors (${U.by_zone["OFF-BRAND"]} including single-competitor ones)`),
  ...figure("fig12_zone_offbrand.png", FW, "Figure 12 – OFF-BRAND sub-categories grouped by the competitor department they come from.", "run_framework.py → figures/fig12_zone_offbrand.png · sheet OFFBRAND"),
  table(["#", "Competitor department", "Examples", "Competitors", "AAS", "Sub-categories"],
    zoneFam("OFF-BRAND").slice(0, 10).map((r, i) => [String(i + 1), aisle(r.family_display), r.examples, r.competitors, f0(r.AAS), String(r.n_units)]),
    [400, 3000, 6200, 2424, 700, 1100], { rowFill: (ri, ci) => (ci === 1 ? C.tint["OFF-BRAND"] : undefined) }),
  spacer(4),
  P(`These are the categories we deliberately leave out, which matters as much as what we pick. The biggest clusters are craft niches (screen printing, purse making, embroidery), holiday-specific décor, individual sports (swimming, handball, airsoft), and farm and garden supplies (hydroponics, beekeeping, poultry care, soils).`),
];

const verify = [
  ...zoneHeader(9, "VERIFY", "VERIFY – already sold, but hard to find",
    "before buying or opening any assortment here, check the Staples shelf listed below. Most of these are findability problems – naming and navigation – not assortment gaps.",
    `${FC["VERIFY"]} Staples aisles · ${RU["VERIFY"]} sub-categories seen at 2+ competitors · ${V.carried_under_other_name} of ${V.checked} checked "gaps" were already on Staples.com`),
  ...figure("fig13_zone_verify.png", FW, "Figure 13 – VERIFY: Staples aisles where categories the models first flagged as missing were found.", "run_framework.py → figures/fig13_zone_verify.png · sheets VERIFY, Verification_Log"),
  table(["Category flagged as missing", "Seen at", "Where it already lives on Staples.com (or status)"],
    top("VERIFY").slice(0, 12).map(r => [clean(r.name), `${r.k}/5`, r.verify_outcome === "to check" ? "Not yet checked – check before acting" : aisle(r.verify_outcome.replace("already carried: ", ""))]),
    [4200, 1200, 8424]),
  spacer(4),
  ...figure("fig05_verification_loop.png", 7.0, "Figure 5 – The verification check: of the model-flagged gaps we checked, most were already on Staples.com.", "run_framework.py → figures/fig05_verification_loop.png · sheet Verification_Log"),
  B(`**Findability quick wins.** Rename or cross-list "Compasses", the shelf that holds patio furniture. Surface candles and home fragrance, evaporative coolers, pull-out pantries, quilts and coverlets, and sneeze guards (sold as "Desktop Privacy Panels") under the words customers use.`),
  B(`**Health and personal-care pages** – mostly Walmart brand pages, kept in the watch list – map onto Staples' OTC and personal-care shelves. Staples carries these product types, but the shelves are thin: Ear Drops & Eye Drops has 4 items.`),
];

// ------------------------------------------------------------------------------------------------
const themes = [
  H1("10. The L1 view: strategic themes", { pageBreak: true }),
  P(`Grouping the aisles into themes gives the L1 story. The themes group Staples departments, and some shelves are placed by keyword because Staples files a few lifestyle shelves in unexpected places.`),
  ...figure("fig14_themes.png", 9.0, "Figure 14 – Strategic themes × zones (sub-categories with 2+ competitor evidence).", "run_framework.py → figures/fig14_themes.png · sheet Themes"),
  table(["Theme", "CURATE", "VERTICAL EXT.", "REVIEW", "1P-CORE GAP", "OFF-BRAND", "VERIFY", "ΣO (CURATE + VE)"],
    Object.entries(S.themes).map(([t, v]) => [t, ...["CURATE", "VERTICAL EXTENSION", "REVIEW", "1P-CORE GAP", "OFF-BRAND", "VERIFY"].map(z => String(v.by_zone[z] || 0)), f1(v.sum_O_curate_ve)]),
    [3824, 1300, 1400, 1200, 1400, 1300, 1200, 2200], { align: [null, 1, 1, 1, 1, 1, 1, 1].map(x => x ? d.AlignmentType.CENTER : undefined) }),
  spacer(4),
  B(`**Home décor & furnishings** leads on both counts: ${S.themes["Home décor & furnishings"].by_zone.CURATE} CURATE sub-categories and ΣO ${f1(S.themes["Home décor & furnishings"].sum_O_curate_ve)}. It is led by Furniture › Décor – bedding, bath, candles, vanities, daybeds – with bedroom furniture, tables and sofas behind it. This is the "white chair" logic: lifestyle furnishings for the same customer who buys the home-office chair.`),
  B(`**Celebrations, gifting & creativity** (ΣO ${f1(S.themes["Celebrations, gifting & creativity"].sum_O_curate_ve)}) is led by sewing and quilting in CURATE and party décor in VERTICAL EXTENSION. Most of its OFF-BRAND volume is holiday-specific and craft-niche.`),
  B(`**Outdoor living & grounds** (ΣO ${f1(S.themes["Outdoor living & grounds"].sum_O_curate_ve)}): planters, hedge trimmers, chimineas and the patio shelf in CURATE; greenhouses, gazebos and garden equipment behind them in VERTICAL EXTENSION.`),
  B(`**Core office & business supplies**: ${S.core_office_theme.one_p} of its ${S.core_office_theme.total} sub-categories are 1P-CORE GAP – as it should be.`),
];

const views = [
  H1("11. How the recommendations connect to Staples' assortment", { pageBreak: true }),
  P(`Two views from the two analytical engines show the same thing from different angles. The graph view (Figure 17) follows each recommendation through the knowledge graph to the Staples aisle and department it plugs into. The vector view (Figure 16) maps the whole category space. Visually, CURATE and VERTICAL EXTENSION sit next to Staples' furniture, décor and kitchen clusters, while much of OFF-BRAND sits in separate clusters.`),
  ...figure("fig17_graph_dock.png", 7.4, "Figure 17 – Graph-DB view: where each top recommendation plugs into Staples' existing aisles and departments.", "method_g_graph.py + run_framework.py → figures/fig17_graph_dock.png"),
  PB(),
  ...figure("fig16_vector_map.png", 8.0, "Figure 16 – Vector-DB view: the category space (t-SNE of the Qdrant vectors). Grey = Staples' shelves; colour = recommendation zone.", "method_v_vector.py + run_framework.py → figures/fig16_vector_map.png"),
];

const trust = [
  H1("12. How much to trust this, and what it does not yet include", { pageBreak: true }),
  ...figure("fig19_sensitivity.png", 7.0, "Figure 19 – Stress test: share of 500 re-weighted runs in which each top CURATE sub-category keeps its zone and its top-20 place.", "run_framework.py → figures/fig19_sensitivity.png · sheet All_Units (p_approved, p_top_n)"),
  B(`**Zones are robust; the order inside a zone is indicative.** ${S.sensitivity["top20_curate_p_approved_ge_0.8"]} of the top 20 CURATE sub-categories stay CURATE in 80% or more of runs, while only ${S.sensitivity["top20_curate_p_top_n_ge_0.8"]} hold a fixed top-20 place. Many opportunities score close together.`),
  B(`**Every flagged gap behind an actionable call was checked.** ${S.verification_coverage.labelled} of ${S.verification_coverage.member_shelves} competitor shelves behind the consensus ENTER recommendations are verified labels; in total ${V.checked} flagged gaps were checked and ${V.carried_under_other_name} turned out to be already sold.`),
  B(`**Two independent methods.** The vector and graph engines agree on the zone for ${pct(S.method_agreement.same_zone_share)} of new-category sub-categories. Their adjacency scores correlate at ρ = ${f2(S.method_agreement["Adjacency (AAS)"])}. Where they disagree, the average places the unit, and the evidence grade drops to B or C (Figure 20 in the methodology document).`),
  B(`**Matching quality.** The vector method picks the right Staples shelf first time for ${pct(cal.OfficeDepot.top1_V)} of Office Depot shelves and ${pct(cal.WestElm.top1_V)} of West Elm's. It does so for ${pct(cal.Wayfair.top1_V)}–${pct(cal.Walmart.top1_V)} of Wayfair, Amazon and Walmart shelves, whose vocabularies are further from Staples'. That is why flagged gaps go through verification before they are recommended.`),
  B(`**Not yet included:** Staples sales and margin data. The "coreness" behind CRS is an explicit, editable assumption; replacing it with Staples' category sales bands is the single biggest upgrade. Demand signals such as search volume and social trends are not included either; the code accepts them as an optional input. Finally, the labelled checks were made by the analysis team with Claude and still need merchant confirmation.`),
];

const next = [
  H1("13. Recommended next steps", { pageBreak: false }),
  N(`**Merchant session (2 hours):** walk the CURATE and REVIEW tables with the category merchants. Confirm or override each zone and record the verdict in the workbook – the code reads it back as labels.`, "numbers2"),
  N(`**Replace the coreness assumption with sales bands:** one data pull of 1P revenue and margin by L2 aisle turns CRS from an assumption into a measurement.`, "numbers2"),
  N(`**Add demand:** search volume for the top 40 sub-categories, from keyword tools and on-site search, feeds the optional demand term in O.`, "numbers2"),
  N(`**Run the archetype PoC on the top CURATE aisles:** Furniture › Décor (bedding, bath, candles), Sewing & Tailoring, and Grounds Maintenance and outdoor living. Compare SKUs and attributes to decide what a curated seller assortment looks like.`, "numbers2"),
  N(`**Ship the findability fixes now:** rename the "Compasses" shelf, cross-list the VERIFY items under the words competitors use, and add them to on-site search synonyms. These need no new assortment.`, "numbers2"),
  N(`**Re-run quarterly:** the pipeline re-runs end to end on fresh navigation trees (about 6 minutes). A new competitor is one adapter entry plus about 120 labelled rows.`, "numbers2"),
];

const appendix = [
  H1("Appendix – figure and table index", { pageBreak: true }),
  P(`Every figure is written by the Python code to *outputs/final/figures/*. Every table in this document is a view of *outputs/final/Category_Recommendations_v2.xlsx*.`),
  table(["Figure", "What it shows", "Produced by", "Workbook sheet"], [
    ["1", "Competitor panel and data availability", "run_framework.py (framework_outputs.fig01_data)", "Data_Audit"],
    ["4", "How much of each competitor Staples covers", "fig04_coverage", "Node_Ensemble, Calibration"],
    ["5", "Verification check", "fig05_verification", "Verification_Log"],
    ["6", "Decision matrix at aisle level (L2)", "fig06_family_matrix", "Aisle_Recommendations"],
    ["7", "Decision matrix at sub-category level (in methodology)", "fig07_unit_matrix", "All_Units"],
    ["8–13", "One chart per zone", "fig_zone", "CURATE … VERIFY"],
    ["14", "Strategic themes × zones", "fig14_themes", "Themes"],
    ["15", "Peer evidence behind top picks", "fig15_peers", "Concept_Members, Deepen_Evidence"],
    ["16", "Vector-DB map of the category space", "fig16_vector_map (Qdrant vectors)", "–"],
    ["17", "Graph-DB view: where recommendations dock", "fig17_graph (knowledge graph)", "Aisle_Recommendations"],
    ["18", "DEEPEN evidence by competitor", "fig18_deepen", "Deepen_Evidence"],
    ["19", "Stress test (500 runs)", "fig19_sensitivity", "All_Units"],
  ], [1100, 5200, 4500, 3024]),
];

// ------------------------------------------------------------------------------------------------
const doc = new d.Document({
  creator: "Category PoC team", title: "Staples.com Marketplace – Category-Level Recommendations",
  description: "Category-level marketplace recommendations for Staples, from six navigation trees",
  styles: L.styles(), numbering: L.numbering(),
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840, orientation: d.PageOrientation.LANDSCAPE },
      margin: { top: 900, bottom: 900, left: 1008, right: 1008 } } },
    footers: { default: L.footer("Staples marketplace – category recommendations (PoC v2)") },
    children: [...cover, ...exec, ...panel, ...zonesExplained, ...curate, ...ve, ...review, ...core, ...offb, ...verify,
      ...themes, ...views, ...trust, ...next, ...appendix],
  }],
});
const out = path.join(L.ROOT, "deliver", "Staples_Category_Recommendations_v2.docx");
fs.mkdirSync(path.dirname(out), { recursive: true });
d.Packer.toBuffer(doc).then(buf => { fs.writeFileSync(out, buf); console.log("wrote", out); });
