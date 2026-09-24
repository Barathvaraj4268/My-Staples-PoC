// Methodology document (portrait text, landscape pages for wide figures). Numbers: summary.json.
const fs = require("fs");
const path = require("path");
const L = require("./lib");
const { d, S, C, P, H1, H2, H3, B, N, PB, figure, table, callout, spacer, pct, f0, f1, f2, SHORT } = L;

const PW = 9936;            // portrait content width (Letter, 1.1" margins)
const LW = 13680;           // landscape content width (0.75" margins)
const cal = S.calibration, V = S.verification, R = S.retailers;
const COMPS = ["OfficeDepot", "WestElm", "Wayfair", "Amazon", "Walmart"];
const ALL = ["Staples", ...COMPS];
const mode = { terminal_only: "Full – leaf pages", cumulative: "Full – all pages", partial: "Partial – L1–L2", none: "None" };

function sec(children, landscape = false) {
  return { properties: { page: landscape
      ? { size: { width: 12240, height: 15840, orientation: d.PageOrientation.LANDSCAPE }, margin: { top: 900, bottom: 900, left: 1080, right: 1080 } }
      : { size: { width: 12240, height: 15840 }, margin: { top: 1300, bottom: 1200, left: 1152, right: 1152 } } },
    footers: { default: L.footer("Category PoC v2 – methodology") }, children };
}
const co = (lines, fill) => { L.callout.width = PW; return callout(lines, fill); };
const coL = (lines, fill) => { L.callout.width = LW; return callout(lines, fill); };

const bake = S.encoder_bakeoff || {};
const grid = S.context_grid_best || [];
const sa = S.scope_audit;
const sumDrop = (r, key) => Object.entries(sa[r] || {}).filter(([k]) => k.startsWith(key)).reduce((a, [, v]) => a + v, 0);

// ================================================================================================
const s1 = [
  new d.Paragraph({ spacing: { before: 1800, after: 160 }, children: [new d.TextRun({ text: "Staples Marketplace PoC · Category level", size: 28, color: C.muted, font: L.FONT })] }),
  new d.Paragraph({ spacing: { after: 200 }, children: [new d.TextRun({ text: "Methodology v2: six retailers, two engines, one decision framework", size: 48, bold: true, font: L.FONT })] }),
  new d.Paragraph({ spacing: { after: 600 }, children: [new d.TextRun({ text: "Track A (category gaps) and Track C (adjacency and cannibalisation), rebuilt for Staples, Office Depot, West Elm, Wayfair, Amazon and Walmart", size: 26, color: C.muted, font: L.FONT })] }),
  P(`**Version:** 2.0, 24 September 2026. It supersedes v1, which used Office Depot only. **Companion documents:** *Staples_Category_Recommendations_v2.docx*, written for Pat; *Staples_Category_PoC_v2.pptx*; the code in */code*; and *Category_Recommendations_v2.xlsx*.`),
  H1("Summary", { pageBreak: true }),
  P(`The PoC answers one question: **which categories should Staples open to marketplace sellers, and why?** It compares Staples' navigation tree – its category "shelf list" – with five competitors'. From the comparison it finds categories Staples lacks (**ENTER**) or carries only thinly (**DEEPEN**). It then scores each one on two Track C questions. Is it close to what Staples already sells (**AAS**)? Would a marketplace there cannibalise Staples' 1P core (**CRS**)? The two scores place every category in one of six zones: CURATE, VERTICAL EXTENSION, REVIEW, 1P-CORE GAP, OFF-BRAND and VERIFY.`),
  P(`Two independent engines do the matching, so no single technique decides alone. **Method V** is vector-first: it embeds every category label with its path and stores the vectors in a **Qdrant** vector database. **Method G** is graph-first: it builds one **knowledge graph** of all six trees and matches on structure as well as wording. The final scores average the two, and the evidence grade records whether they agree.`),
  co([`**Headline numbers from this run.** ${f0(S.competitor_shelves_compared)} competitor category pages compared · ${f0(S.labelled_rows.calibration)} labelled rows for calibration, plus ${f0(S.labelled_rows.verification)} verification checks · ${f0(S.units.total)} category units scored · ${S.family_counts["CURATE"]} Staples aisles to CURATE (${S.recommendable_units_by_zone["CURATE"]} sub-categories) · ${V.carried_under_other_name} of ${V.checked} flagged gaps found already on Staples.com.`], "F4F6F8"),
  H2("What changed from v1, and why"),
  table(["v1 (Office Depot only)", "What broke with six retailers", "v2 fix"], [
    ["Item counts on every tree", "Wayfair and Amazon publish none; Walmart only at L1–L2", "Count-optional metrics: breadth (number of shelves) for everyone, items wherever both sides publish them"],
    ["Embedding = own 0.5 + parent 0.3 + L1 0.2", "Trees run 2 to 5 levels, with menu headings and facets", "Level-agnostic path embedding: own label plus whatever ancestors exist, decaying weight, headings skipped. Asymmetric weights: +" + f1(100 * ((S.context_grid_chosen || 0) - (S.context_grid_label_only || 0))) + " points top-1 accuracy over label-only"],
    ["One gold set (Office Depot, 215 rows)", "Other taxonomies score differently", "120 labelled rows per new competitor; per-competitor thresholds; pooled fallback for the next competitor"],
    ["Units = one competitor's gap nodes", "Five overlapping answers", "ENTER concepts merged across competitors; DEEPEN anchored on Staples aisles; both rolled up to Staples L2"],
    ["AAS anchored per competitor", "Vector and graph AAS disagreed (ρ 0.06 in the first v2 run)", "Pooled quantile anchoring on one ruler: ρ rises to " + f2(S.method_agreement["Adjacency (AAS)"])],
    ["Model's gap flag taken as-is", `${pct(V.carried_under_other_name / V.checked)} of flagged gaps were naming differences`, "Verification loop plus a VERIFY zone that becomes a findability action list"],
    ["One Qdrant collection, filtered by retailer", "Local-mode payload filters are slow at 15k points", "One collection per retailer (about 80x faster per query in a development benchmark); the same layout shards on a server"],
    ["Peer coverage from gap units only", "A competitor can carry a gap outside its own gap units", "Every concept is searched in every competitor's Qdrant collection"],
  ], [2700, 3300, 3936]),
];

// ================================================================================================
const s2 = [
  H1("1. Inputs and scope", { pageBreak: true }),
  P(`Each tree is loaded by an **adapter** – one configuration entry per retailer that names its file, columns, count semantics and scope rules. Adding a sixth competitor means adding one entry. Every node gets a standard record: path, depth, label, parent, leaf flag, item count (or unknown), count mode and scope status.`),
  table(["Retailer", "Role", "Rows in file", "In scope (shelves)", "Leaf shelves", "Max depth", "Item counts"],
    ALL.map(r => [SHORT[r] || r, (R[r] || {}).role || "", f0((S.rows_in_files || {})[r]), f0(R[r].shelves_in_scope), f0(R[r].leaf_shelves), "L" + R[r].max_depth, mode[R[r].count_mode]]),
    [1500, 1100, 1300, 1600, 1300, 1136, 2000]),
  spacer(4),
  H2("1.1 One department universe for every retailer"),
  P(`The PoC compares like with like. One explicit list of out-of-scope departments applies to all six trees: media, apparel, beauty, auto, baby consumables, toys, grocery, pets, musical instruments, collectibles, gift cards, services, and brand, character and curated hubs. Everything else stays in, including lifestyle departments far from office supplies, so that the framework – not a pre-filter – decides what is OFF-BRAND. Amazon's team-curated "PoC-relevant L1" flag is honoured, with Cell Phones & Accessories added back because Staples sells phone accessories. Pet, which was a HOLD in v1, is out of the universe, consistent with that flag.`),
  H2("1.2 Cleaning navigation noise"),
  B(`**Merchandising pages** – New, Sale, Featured items, sub-brands such as Pierce & Ward, named collections, "shop by brand" and team shops – are removed with their subtrees. Walmart: ${f0(sumDrop("Walmart", "merchandising"))} nodes; West Elm: ${f0(sumDrop("WestElm", "merchandising"))} merchandising pages plus ${f0(sumDrop("WestElm", "retailer scope"))} named-collection and programme pages.`),
  B(`**Facets** are filters, not categories – by colour, fabric, style, room, size, or West Elm's "Rugs By Color". Walmart: ${f0(sumDrop("Walmart", "facet"))}; West Elm: ${f0(sumDrop("WestElm", "facet"))}; Wayfair: ${f0(sumDrop("Wayfair", "facet"))}.`),
  B(`**Headings** such as "Shop By Category", "Featured", "More Rooms" and Walmart's unnamed ID nodes are kept in the tree. They are skipped as meaning, and never become a recommendation. "All X" aggregate pages are dropped when they duplicate their parent.`),
  B(`**Depth.** Staples goes to L4, so competitor levels below L4 fold into their L4 parent: ${f0(sumDrop("Walmart", "deeper"))} Walmart nodes. Recommendations are made at L1–L3.`),
  B(`**Counts.** Staples publishes counts on leaf pages only, and cross-lists heavily: ${f0(S.verification.staples_crosslist_edges)} extra placements after scoping (one shelf is listed under 50 parents in the raw sheet). Office Depot and West Elm publish cumulative counts. Walmart's are cumulative at L1–L2, and its "900,000+" cap is treated as unknown. Wayfair and Amazon have none, so they contribute breadth only.`),
];

const s2fig = [
  ...figure("fig01_data_availability.png", 8.8, "Figure 1 – Competitor panel and what each navigation tree gives us."),
  ...figure("fig04_staples_coverage.png", 8.6, "Figure 4 – How much of each competitor's shelf list Staples already covers (model view vs labelled random sample)."),
];

// ================================================================================================
const s3 = [
  H1("2. The v2 flow", { pageBreak: true }),
  P(`Figure 2 (next page) shows the seven stages. Each stage is a function in the delivered code. The table maps each stage to where it lives.`),
  table(["Stage", "What happens", "Code"], [
    ["1 Load & scope", "Six adapters → one node table; universe, merchandising, facets, headings, depth fold; counts harmonised by mode", "poc_common.load_all"],
    ["2 Represent", "Label + decaying ancestor context (vector); wording + thesaurus + structure (graph)", "method_v_vector.node_vectors · poc_common.Lexical"],
    ["3 Match", "V: Qdrant top-50 → hybrid rerank. G: similarity flooding over the knowledge graph", "method_v_vector · method_g_graph"],
    ["4 Calibrate & verify", "Per-competitor thresholds from labelled rows plus ablation negatives; ensemble status; verification-loop labels override the models", "poc_common.calibrate · run_framework.build_node_table"],
    ["5 Build units", "ENTER concepts across competitors (plus peer search in Qdrant); DEEPEN aisles with size factors; roll-up to Staples aisles", "run_framework.enter_units / concept_table / deepen_units / assign_families"],
    ["6 Score", "AAS, CRS (V and G averaged), PC, GS, O", "run_framework.label_and_rank · poc_common.opportunity"],
    ["7 Decide", "Six zones, VERIFY, evidence grade, 500-run stress test, tables and figures", "poc_common.sensitivity · framework_outputs"],
  ], [1900, 5136, 2900]),
];

// ================================================================================================
const s4 = [
  H1("3. Representation: level-agnostic path embeddings", { pageBreak: true }),
  P(`A shelf label alone is often ambiguous: "Covers", "Bedroom", "Accessories". v1 fixed the context to exactly the parent and the L1, which breaks when trees differ in depth. v2 uses **whatever path the retailer shows**, with weights that decay going up the path:`),
  P(`**v(node) = w · e(label) + (1 − w) · Σₖ δ^(k−1) · e(ancestorₖ) / Σₖ δ^(k−1)**`, { align: d.AlignmentType.CENTER }),
  P(`Here e() is the text encoder, the ancestors are ordered from parent upward, and navigation headings are skipped. A Wayfair rug (two levels) and a Walmart L4 shelf are built the same way. Generic labels ("Accessories", "Storage") are first expanded with their parent's words.`),
  P(`**Asymmetric weights, chosen on data.** A grid over the competitor weight and decay and the Staples weight (Staples δ fixed at 0.5), scored on top-1 shelf accuracy across all labelled rows, puts **competitor w = 0.40, δ = 0.35 and Staples w = 0.70** at ${pct(S.context_grid_chosen, 1)} top-1, against ${pct(S.context_grid_label_only, 1)} for the label alone. The best cell (competitor w = 0.30) is within one percentage point; we kept 0.40 as the less extreme setting. The asymmetry has a plain reason. Competitor labels are often generic and need their context. Staples' upper levels are noisy – 1,116 patio-furniture items sit under Gift Shop › Professional Gifts › Compasses – so Staples shelves lean on their own label.`),
  table(["Competitor w", "Competitor δ", "Staples w", "Top-1 (pooled)"],
    [...grid.map(g => [f2(g.w_self_competitor), f2(g.ctx_decay_competitor), f2(g.w_self_staples), pct(g.top1_hybrid_pooled, 1)]),
     ["1.00", "–", "1.00", pct(S.context_grid_label_only, 1) + " (label only)"]], [2400, 2400, 2400, 2736]),
  caption("Top of the context-weight grid (method_v_vector.bakeoff → outputs/V/context_weight_grid.csv)."),
];

function caption(t) { return L.caption(t); }

// ================================================================================================
const s5 = [
  H1("4. Method V – vectors and the Qdrant vector database", { pageBreak: true }),
  H2("4.1 Which embedding model, and why"),
  P(`The encoder is chosen by a **bake-off**, not by reputation. Every encoder that loads is scored on the labelled rows (top-1 shelf accuracy after the hybrid rerank, pooled across competitors):`),
  table(["Encoder", "What it is", "Top-1 (pooled)", "Status here"], [
    ["sentence-transformers (BAAI/bge-small-en-v1.5)", "Transformer sentence embeddings, 384-d", "–", (S.encoder_unavailable || []).includes("st") ? "Not reachable in this sandbox (HuggingFace blocked); runs in Colab" : "available"],
    ["spaCy en_core_web_lg", "300-d word vectors, IDF-weighted mean", pct(bake.spacy, 1), "Used for this run (default)"],
    ["WordLlama (l2_supercat 256-d)", "Token embeddings distilled from an LLM, weights ship in the pip wheel", pct(bake.wordllama, 1), "Available offline"],
    ["TF-IDF char n-grams + SVD", "No model at all", pct(bake.tfidf, 1), "Fallback"],
  ], [3100, 3100, 1500, 2236]),
  spacer(4),
  P(`spaCy wins on this run by ${f1(100 * (bake.spacy - bake.wordllama))} points over WordLlama and ${f1(100 * (bake.spacy - bake.tfidf))} over TF-IDF. Setting encoder = "auto" repeats the bake-off wherever the code runs. In Colab, with HuggingFace reachable, it will test bge-small too and keep whichever is best. Thresholds are re-calibrated for whichever encoder wins, so no cosine value is hard-coded.`),
  H2("4.2 Retrieve, then rerank"),
  P(`For each competitor shelf, Qdrant returns the 50 nearest Staples shelves by cosine similarity. They are then re-scored as **score = 0.65 × cosine + 0.35 × wording**. Wording is TF-IDF on stemmed words, bigrams and character 3–5-grams: characters catch "Pads" ≈ "Notepads", and words catch exact category vocabulary. Adjacency for Method V is the mean of the top-10 hybrid scores – how dense Staples' neighbourhood is around the shelf.`),
  H2("4.3 Why a vector database, and how it is organised"),
  P(`At category level – about 15,000 vectors – brute-force NumPy would do. Qdrant earns its place for three reasons. It is the store the SKU-level archetype PoC will need: millions of vectors, approximate-nearest-neighbour indexes and payload filters. It is persisted to disk (outputs/V/qdrant_db), so run_framework.py can query it later. And it answers the **peer question** directly: "does Wayfair have a shelf like this concept?" is a nearest-neighbour search in Wayfair's collection.`),
  P(`**One collection per retailer.** v1 used one collection with a retailer filter. In Qdrant's local mode that filter is evaluated in Python: in a development benchmark on this data, 1,000 filtered queries took 167 s, against about 2 s per 1,000 unfiltered queries – roughly 80 times slower. Hub-and-spoke queries always target one retailer, so per-retailer collections are both faster and the natural sharding for a Qdrant server. Each point's payload carries uid, path, depth and whether a count exists.`),
];

// ================================================================================================
const G = { nodes: 15064, CHILD_OF: 14638, SAME_AS: 9095, CROSS: 2228 };
try { const gi = JSON.parse(fs.readFileSync(path.join(L.ROOT, "outputs", "G", "run_info_G.json"))); G.nodes = gi.n_nodes; G.CHILD_OF = gi.edges_by_type.CHILD_OF; G.SAME_AS = gi.edges_by_type.SAME_AS; G.CROSS = gi.edges_by_type.CROSS_LISTED_UNDER; } catch (e) {}
const s6 = [
  H1("5. Method G – the category knowledge graph", { pageBreak: true }),
  P(`Method G treats the six trees as **one network**: ${f0(G.nodes)} category nodes; ${f0(G.CHILD_OF)} CHILD_OF edges (the trees); ${f0(G.CROSS)} CROSS_LISTED_UNDER edges (Staples shelves that sit under several parents – the ${f0(S.verification.staples_crosslist_edges)} placements minus pairs of shelves listed under each other, which are one edge in an undirected graph); and, after matching, ${f0(G.SAME_AS)} SAME_AS edges linking competitor shelves to Staples shelves. A graph can use two things a vector search cannot:`),
  B(`**Structure-aware matching (similarity flooding).** Two shelves are more likely the same when their parents and children also match. The wording-plus-thesaurus score is refined over three rounds: σ = (S₀ + λp·parent term + λc·child term) / (1 + λp + λc), with λp = 0.40 and λc = 0.35. On Office Depot this lifts top-1 accuracy from ${pct(cal.OfficeDepot.top1_G_wording_only)} (wording only) to ${pct(cal.OfficeDepot.top1_G)}.`),
  B(`**Adjacency as network proximity (Personalised PageRank).** Random walks start on Staples' own shelves, weighted by log item count, and move through CHILD_OF, CROSS_LISTED and SAME_AS edges. A shelf's adjacency is the lift of this "Staples gravity" over plain PageRank, blended 70/30 with sibling coverage – the share of its competitor aisle that Staples already carries. That is exactly the "white chair" test: a missing shelf in an aisle Staples otherwise sells.`),
  P(`**Why a graph as well as vectors.** Similarity is the vector engine's job; structure and consensus are the graph's. The two methods fail differently. Vectors can be fooled by shared words ("Drinking Glasses" ≈ "Drinking Fountains"); the graph is anchored by where shelves sit. So averaging them, and recording when they disagree, gives a sturdier answer than either alone. The graph is exported for **Neo4j** (outputs/G/neo4j: nodes.csv, edges.csv, load_graph.cypher and example_queries.cypher). The example queries cover consensus whitespace, aisle whitespace, cannibalisation checks and PageRank with GDS, plus GraphML for Gephi. Setting NEO4J_URI pushes the graph to a live database.`),
];

// ================================================================================================
const s7 = [
  H1("6. Calibration, ensemble and verification", { pageBreak: true }),
  H2("6.1 Labelled data"),
  P(`Thresholds come from labelled examples, never from a fixed cosine. Office Depot keeps v1's 215 rows. Each new competitor received **120 rows**: 10 from every decile of the Method V score, which makes an unbiased random sample, plus 20 from the bottom two deciles, where real gaps live. Each row records whether Staples carries the category and on which shelf. The labels were produced by the analysis team with Claude, searching the full Staples tree and, where unclear, checking the live Staples.com page. They are marked "review with team" in *gold_matches_v2.csv*.`),
  table(["Competitor", "Labelled", "Carried (random sample)", "V τ", "V AUC", "V top-1", "G τ", "G AUC", "G top-1", "G wording only"],
    COMPS.map(c => [SHORT[c], String(cal[c].n_labelled), pct(cal[c].base_rate_carried), f2(cal[c].tau_V), f2(cal[c].auc_V), pct(cal[c].top1_V), f2(cal[c].tau_G), f2(cal[c].auc_G), pct(cal[c].top1_G), pct(cal[c].top1_G_wording_only)]),
    [1400, 900, 1300, 700, 800, 900, 700, 800, 900, 1536], { size: 16 }),
  spacer(4),
  B(`**τ, the same-shelf threshold,** maximises a cost-weighted Youden's J on the labelled rows. Real misses are rare – Office Depot has 98% carried – so the negatives are topped up with **ablation negatives**: each carried row is re-scored with its Staples department hidden, which shows what a genuine miss looks like.`),
  B(`**The gap screen** is set by a false-alarm budget: the 10th percentile of scores of shelves Staples does carry. A shelf below it is "possibly missing".`),
  B(`**Pooled calibration** covers a competitor with fewer than 40 labelled rows. The next competitor therefore runs on day one, with its calibration flagged as provisional.`),
  H2("6.2 Combining the two methods per shelf"),
  table(["Method V", "Method G", "Ensemble status", "Used as"], [
    ["matched", "matched or likely", "carried", "Carried; feeds DEEPEN depth"],
    ["likely", "likely", "likely", "Carried, exact shelf not pinned"],
    ["gap", "gap", "gap (hard)", "Missing"],
    ["gap", "likely (or vice versa)", "gap (soft)", "Missing, but VERIFY if most of a concept is soft"],
    ["gap", "matched (or vice versa)", "disputed", "Not a gap"],
  ], [2200, 2600, 2200, 2936]),
  P(`A labelled row always overrides both methods: truth beats the model.`),
  H2("6.3 The verification loop"),
  P(`Matching accuracy on the lifestyle and scale retailers is well below Office Depot's, because their vocabularies are further from Staples'. So v2 does not trust a model-flagged gap on sight. Over three rounds, ${V.checked} model-flagged shelves were checked against the Staples tree; after the last round, all ${S.verification_coverage.member_shelves} competitor shelves behind the consensus (2+ competitor) new-category recommendations in CURATE, VERTICAL EXTENSION, REVIEW, 1P-CORE GAP and VERIFY are verified labels. **${V.carried_under_other_name} (${pct(V.carried_under_other_name / V.checked)}) were already on Staples.com under other names**, and ${V.confirmed_gaps} were real gaps. These checks are stored as labels, so they override the models on every re-run. They are excluded from threshold fitting, because they were chosen by the model and would bias τ. Concepts cleared this way become the VERIFY zone's findability list.`),
];
const s7fig = [
  ...figure("fig03_match_quality.png", 8.8, "Figure 3 – Matching quality by competitor: top-1 shelf accuracy and carried-vs-missing separation (AUC), Methods V and G."),
  ...figure("fig05_verification_loop.png", 7.6, "Figure 5 – Verification loop: model-flagged gaps checked, and how many were already on Staples.com."),
];

// ================================================================================================
const sf = S.size_factors || { items: {}, shelves: {} };
const s8 = [
  H1("7. Building recommendation units", { pageBreak: true }),
  H2("7.1 ENTER – categories Staples does not carry"),
  B(`A competitor node becomes an ENTER candidate when it is a gap and **80% or more of its leaves are missing** at Staples, weighted by items where every leaf has a count. The highest such node at L1–L3 is kept; with counts, it must have at least 10 items.`),
  B(`**Concepts across competitors.** "Floor Lamps" at Wayfair and "Floor Lamps" at Walmart are the same opportunity. Candidates are clustered by hybrid similarity (average linkage, merge threshold = median same-shelf τ = 0.67), and average linkage avoids chaining unrelated shelves. Each concept is then searched in **every** competitor's Qdrant collection, so a competitor that sells it inside a bigger aisle still counts. That gives peer coverage PC = k of 5.`),
  B(`**Single-competitor concepts** (${S.watchlist_single_competitor}) are kept as a watch list, but do not drive aisle-level calls.`),
  H2("7.2 DEEPEN – Staples aisles that are too thin"),
  P(`Each carried competitor leaf is assigned to the Staples shelf it matches, and through Staples' canonical and cross-listed placements to every Staples aisle above it. For each Staples L2 or L3 aisle and each competitor:`),
  P(`**depth gap = ln(1 + competitor size) − ln(1 + size factor × Staples size)**`, { align: d.AlignmentType.CENTER }),
  P(`Size is measured in **items** when both sides publish counts for at least 80% of the leaves (mainly Office Depot and West Elm, and Walmart where an L2 page is itself a leaf), and in **number of shelves** otherwise. The **size factor** is a median-of-ratios on confidently matched pairs, as in DESeq, so one giant aisle cannot distort it. Items: Office Depot ${f2(sf.items.OfficeDepot)}, West Elm ${f2(sf.items.WestElm)}. Shelves: Office Depot ${f2(sf.shelves.OfficeDepot)}, West Elm ${f2(sf.shelves.WestElm)}, Wayfair ${f2(sf.shelves.Wayfair)}, Amazon ${f2(sf.shelves.Amazon)}, Walmart ${f2(sf.shelves.Walmart)}. An aisle is DEEPEN when at least two competitors, and at least half of those carrying it, are **at least 2x deeper** (depth gap ≥ ln 2). L3 aisles are preferred; an L2 is kept only when none of its L3s qualifies. Count-less Staples "hub" pages that only list sibling shelves are not treated as aisles.`),
  H2("7.3 Rolling up to the level Staples acts on"),
  P(`Recommendations are shown per **Staples aisle (L1 › L2)**. A DEEPEN unit already is one. An ENTER concept is **docked** into the aisle where it would live, by a vote over three signals. Graph: the competitor shelf's parent aisle, which Staples carries, followed across its SAME_AS link to Staples. Vectors: the same path, using Method V's link. Neighbours: the Staples L1 of the concept's 10 nearest Staples shelves. OFF-BRAND new-category concepts have nowhere to dock, so they are grouped by the competitor department they come from (names normalised, so "Arts Crafts & Sewing" and "Arts, Crafts & Sewing" are one); the few DEEPEN items that land in OFF-BRAND stay with their Staples aisle. VERIFY items are grouped by the Staples aisle where they were found.`),
];

// ================================================================================================
const s9 = [
  H1("8. Scoring and the decision framework (Track C)", { pageBreak: true }),
  H2("8.1 AAS – adjacency affinity (0–100)"),
  P(`Each method gives a raw adjacency: V is neighbourhood density; G is PageRank lift plus sibling coverage. Anchoring each competitor separately (as v1 did, and as the first v2 run still did) stretched each retailer's scale differently: in that run the vector and graph AAS of new-category units correlated at only ρ = 0.06. v2 puts both methods on **one ruler** by pooled quantile anchoring:`),
  P(`**AAS = 100 × min(1, F(A) / F(median A of shelves Staples carries))**`, { align: d.AlignmentType.CENTER }),
  P(`F is the method's empirical distribution over all competitor leaves pooled. So 100 means "as embedded as a typical shelf Staples already carries", and 0 means "the least embedded shelf in the panel". Agreement rises to ρ = ${f2(S.method_agreement["Adjacency (AAS)"])}. Unit AAS is the leaf-weighted mean inside a unit, averaged over V and G.`),
  H2("8.2 CRS – cannibalisation risk (0–100)"),
  P(`CRS = 100 × max over nearby Staples shelves of [coreness(shelf) × closeness(shelf)]. Closeness rises from 0 at a typical non-match to 1 at the same-shelf threshold. **Coreness** is an explicit, editable proxy for how much 1P revenue sits on a shelf. It is 1.0 for Paper, Office Supplies and Shipping; 0.6 for Furniture, with office seating overridden to 1.0; and 0.1–0.2 for décor and lifestyle shelves – the "white chair" zone. It is scaled by the shelf's depth within its level. This is the one place the PoC relies on judgement rather than data. Replacing it with Staples' sales bands is the first recommended upgrade.`),
  H2("8.3 Opportunity and the six zones"),
  P(`**O = (0.35 · PC + 0.35 · GS + 0.30 · AAS/100) × (1 − CRS/100)^γ**, with γ = 1.`, { align: d.AlignmentType.CENTER }),
  B(`**PC – peer consensus:** k/5 competitors carrying it (ENTER) or at least 2x deeper (DEEPEN). It needs no counts.`),
  B(`**GS – gap size:** the percentile of the missing chunk among that competitor's L1–L3 categories. It uses shelves for every retailer, and the shelf and item percentiles are averaged where counts exist. Percentiles keep items-based and shelf-based evidence comparable – the answer to "counts for some retailers, not others".`),
  B(`An optional **demand** term (external_signals.csv, 0–100 per category) is blended in at 25% when supplied.`),
  table(["Zone", "Rule", "Action"], [
    ["CURATE", "CRS < 40, AAS ≥ 60", "Open to curated marketplace sellers"],
    ["VERTICAL EXTENSION", "CRS < 40, 40 ≤ AAS < 60", "Phase 2, through a vertical Staples serves"],
    ["REVIEW", "40 ≤ CRS < 60", "Merchant decision"],
    ["1P-CORE GAP", "CRS ≥ 60", "Fix in 1P, keep out of the marketplace"],
    ["OFF-BRAND", "CRS < 40, AAS < 40", "Pass"],
    ["VERIFY", "Actionable ENTER concept whose absence is unconfirmed (mostly soft gaps, or most peers' equivalent shelves match a Staples shelf); or confirmed as already carried", "Check first; fix findability"],
  ], [2400, 4336, 3200], { rowFill: (ri, ci) => (ci === 0 ? Object.values(C.tint)[ri] : undefined) }),
  spacer(4),
  B(`**Evidence grade:** A = 3 or more competitors and both methods place the unit in the same zone; B = 2 or more competitors, or the methods agree; C = neither.`),
];

const s9fig = [
  ...figure("fig07_matrix_units.png", 8.2, "Figure 7 – Every sub-category unit (L2–L3) on the decision matrix; Figure 6 in the recommendations document rolls these up to Staples aisles."),
];

// ================================================================================================
const s10 = [
  H1("9. Robustness and limitations", { pageBreak: true }),
  B(`**Stress test.** 500 runs re-draw the O weights (Dirichlet around 0.35/0.35/0.30), jitter the zone thresholds by ±5 and vary γ between 0.5 and 2. ${S.sensitivity["top20_curate_p_approved_ge_0.8"]} of the top 20 CURATE units stay CURATE in 80% or more of runs; only ${S.sensitivity["top20_curate_p_top_n_ge_0.8"]} keep a fixed top-20 place. Read zones as robust and within-zone order as indicative.`),
  B(`**Method agreement.** V and G place ${pct(S.method_agreement.same_zone_share)} of new-category units in the same zone (AAS ρ = ${f2(S.method_agreement["Adjacency (AAS)"])}, CRS ρ = ${f2(S.method_agreement["Cannibalisation (CRS)"])}). Disagreement lowers the evidence grade rather than being hidden.`),
  B(`**Matching limits.** Top-1 shelf accuracy is ${pct(cal.OfficeDepot.top1_V)} for Office Depot, but ${pct(cal.Wayfair.top1_V)}–${pct(cal.WestElm.top1_V)} for the others. Walmart labels come from URLs and lose their punctuation, and brand pages behave like categories. This is why the verification loop exists. A transformer encoder (bge-small) where HuggingFace is reachable, and merchant-confirmed labels, are the two biggest accuracy levers.`),
  B(`**Judgement inputs.** Coreness weights, the department universe and the zone thresholds (60/40) are explicit configuration in poc_common.py, and visible in the workbook. No Staples sales, margin or demand data were available to the PoC.`),
  B(`**Counts.** Depth in items is measured mainly against Office Depot and West Elm (plus a few Walmart L2 leaves); elsewhere it is shelf breadth, which is a proxy. Walmart's L1 counts are capped and ignored.`),
];
const s10fig = [
  ...figure("fig19_sensitivity.png", 7.0, "Figure 19 – Stress test for the top CURATE units."),
  ...figure("fig20_method_agreement.png", 8.4, "Figure 20 – Method V vs Method G: AAS and CRS of new-category units."),
];

// ================================================================================================
const s11 = [
  H1("10. Running the code", { pageBreak: true }),
  table(["File", "Role", "Runtime here"], [
    ["poc_common.py", "Shared library: configuration, retailer adapters, scope rules, loading, counts, calibration, framework", "–"],
    ["method_v_vector.py", "Method V: encoder bake-off, path embeddings, Qdrant DB, matching, node scores", "about 2–4 min (including the bake-off and grid)"],
    ["method_g_graph.py", "Method G: knowledge graph, similarity flooding, PageRank, Neo4j/GraphML export", "about 30 s"],
    ["run_framework.py + framework_outputs.py", "Ensemble, verification overrides, ENTER/DEEPEN units, scoring, zones, stress test, all figures, workbook, summary.json", "about 1 min"],
    ["label_sample.py", "Draws a 120-row labelling sample for a new competitor", "seconds"],
    ["run_all_colab.ipynb", "Colab notebook: install, upload, run all three, show the key figures", "–"],
    ["gold_matches_v2.csv", `${S.labelled_rows.total} labelled rows (${S.labelled_rows.calibration} for calibration + ${S.labelled_rows.verification} verification)`, "–"],
  ], [3100, 4836, 2000]),
  spacer(4),
  H2("Colab or local"),
  N(`Put the six .xlsx trees and gold_matches_v2.csv in one folder, together with the .py files.`, "numbers2"),
  N(`pip install pandas numpy scipy scikit-learn openpyxl matplotlib networkx qdrant-client spacy adjustText, then python -m spacy download en_core_web_lg.`, "numbers2"),
  N(`python method_v_vector.py; python method_g_graph.py; python run_framework.py. Set POC_DATA_DIR and POC_OUT_DIR to change folders.`, "numbers2"),
  N(`Outputs: outputs/final/Category_Recommendations_v2.xlsx, outputs/final/figures/fig01–fig20, outputs/final/summary.json, outputs/G/neo4j/*, outputs/V/qdrant_db.`, "numbers2"),
  H2("Adding a competitor"),
  N(`Add one adapter to RETAILERS in poc_common.py: file pattern, level columns, count column and mode, role, scope rules.`, "numbers3"),
  N(`Run method_v_vector.py. The new competitor runs on pooled calibration, flagged as provisional.`, "numbers3"),
  N(`Run label_sample.py <Name>, label the 120 rows, append them to gold_matches_v2.csv and re-run all three scripts.`, "numbers3"),
  N(`Review the new VERIFY items and add the checked verdicts as "verification loop" rows.`, "numbers3"),
  H2("Key parameters (poc_common.CONFIG)"),
  table(["Parameter", "Value", "Meaning"], [
    ["w_self / ctx_decay (competitor)", "0.40 / 0.35", "Label vs ancestor weight in competitor vectors"],
    ["w_self_focal / ctx_decay_focal", "0.70 / 0.50", "Same, for Staples"],
    ["k_retrieve · w_sem / w_lex", "50 · 0.65 / 0.35", "Qdrant recall and hybrid rerank"],
    ["lambda_parent / lambda_child / n_iter", "0.40 / 0.35 / 3", "Similarity flooding"],
    ["ppr_alpha · aas_ppr_weight", "0.85 · 0.7", "PageRank damping; PageRank vs sibling-coverage blend"],
    ["gap_screen_false_alarm", "0.10", "Gap screen = 10th percentile of carried scores"],
    ["enter_coverage_max · max_unit_depth", "0.20 · 3", "ENTER needs ≥ 80% of leaves missing; units at L1–L3"],
    ["deepen_min · deepen_min_carriers", "ln 2 · 2", "At least 2 competitors at least 2x deeper"],
    ["aas_hi/lo · crs_hi/lo", "60/40 · 60/40", "Zone thresholds"],
    ["rank_weights · gamma", "PC 0.35, GS 0.35, AAS 0.30 · 1", "Opportunity score"],
    ["n_sensitivity", "500", "Stress-test runs"],
  ], [3600, 2600, 3736]),
];

// ================================================================================================
const doc = new d.Document({
  creator: "Category PoC team", title: "Category-Level PoC – Methodology v2",
  styles: L.styles(), numbering: L.numbering(),
  sections: [sec([...s1, ...s2]), sec(s2fig, true), sec([...s3]), sec(figure("fig02_flow.png", 9.2, "Figure 2 – End-to-end flow (v2)."), true),
    sec([...s4, ...s5, ...s6, ...s7]), sec(s7fig, true), sec([...s8, ...s9]), sec(s9fig, true), sec(s10), sec(s10fig, true), sec(s11)],
});
const out = path.join(L.ROOT, "deliver", "Staples_Category_PoC_Methodology_v2.docx");
d.Packer.toBuffer(doc).then(buf => { fs.writeFileSync(out, buf); console.log("wrote", out); });
