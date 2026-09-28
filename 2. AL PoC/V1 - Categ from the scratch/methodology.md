# Staples Assortment PoC: Node-Level Assortment Gap and Recommendation Methodology

**Status:** Draft v1.1 for review · 2026-09-27 · Owner: Sai (LatentView). v1.1 adds Sai's §11 answers and makes scope data-driven (§1.3–1.4). v1.2 keeps VOS and TG as separate method scores with a shared gate and fused re-rank (§6.5.6, §7), and sets the output format to code + PNG figures + a static 2-tab HTML report (§9).
**Supersedes:** Track B/C in `Documents/Staples PoC - Approach & Methodolgy -Initial Exploration.docx`, and the 6-step method shared in chat.
**Scope of this document:** the methodology to review. Code comes after sign-off.

---

## 0. Summary

**Question:** Within a granular Staples navigation node (for example *Accent & Waiting Room Chairs* or *Office Desks → L-Shaped*), which product archetypes and attribute variants (colour, material, style, price tier, size, use-context) does a competitor carry that Staples does not? Which of those can Staples add through its marketplace without cannibalizing what it already sells?

**Answer shape:** for each analysis node we produce:
1. a gap profile at attribute level (which attribute values are under-represented),
2. a ranked list of archetypes with a decision label (CURATE, TRADE-UP, STYLE-EXTENSION, and so on),
3. 3–5 exemplar competitor SKUs per shortlisted archetype, shown next to the nearest Staples SKU, with brand and seller notes.

**What changed from the earlier versions.** Each change is traced to evidence in §2 and §3.

| # | Change | Why |
|---|---|---|
| 1 | Map at **product level**, not category-name level | Wayfair "categories" include room and shop-by-space pages (*Meeting Space Seating*, *Hospitality Seating*). 156 products sit on more than one page. The name VLOOKUP failed, and path-embedding mapping in CL PoC V2 reached only 43–55% top-1 for Wayfair |
| 2 | The unit of analysis is the **product family**, not the SKU row | Staples lists each colour as a separate SKU: 4,725 SKUs collapse to about 2,860 families, and 58% of SKUs sit in multi-colour families. Wayfair lists one family with a selected option. Raw counts would inflate Staples' depth |
| 3 | **Share-based gaps** replace count-based Depth Gap | Staples data is close to a census (85–100% of site counts per leaf). Wayfair data is a sample of unknown design with no site totals. `log1p(comp) − log1p(staples)` compares a sample with a census. Shares compare like with like and stay defined when Staples has 0 |
| 4 | Every metric's inputs are **measured the same way on both sides** | Staples descriptions contain a structured spec dict (colour 97%, material 94%, dimensions 91%). Wayfair has free text only. Extracting Staples from specs and Wayfair with an LLM would build measurement bias into every gap. See §6.2 |
| 5 | Embed a **canonical product card**, not raw descriptions | Staples raw text averages 3.8k characters of serialized spec dict. Wayfair averages 0.7k characters of prose plus an appended customer review. Raw embeddings would separate products by source rather than by product |
| 6 | CRS is **price-agnostic**. PPR carries price | In the doc, TRADE-UP needs CRS > 60 and PPR ≥ 1.5. PPR ≥ 1.5 forces the price-proximity term `po` to 0, which removes 25 of CRS's 100 points, so TRADE-UP was nearly unreachable |
| 7 | **Exhaustive decision tree** in place of a threshold grid | The doc's label grid leaves some AAS × CRS regions undefined (for example AAS 40–60 with CRS 40–60) |
| 8 | New **STYLE-EXTENSION** label | This is the user's "flavours" ask: functionally the same as a Staples product but in a colour, material or style Staples lacks. Under the old rules these fell into SUBSTITUTE and were rejected |
| 9 | Components without data are **dropped or replaced** | Nothing in the data supports the Personalised-PageRank complement graph, the "also-viewed" `sub` term, image-based SigLIP DFI, or review-count demand. §6 gives a replacement for each and §11 lists the data requests |
| 10 | Brand fragmentation becomes **recruitability** | Wayfair's `vendor` is mostly Wayfair house brands (Latitude Run, Ebern Designs, Inbox Zero, George Oliver: the top 15 by volume are all house labels). HHI on those measures Wayfair's labelling, not the seller market |
| 11 | Method 1 and Method 2 each keep a **final score (VOS, TG)**, share **one safety gate**, and are **re-ranked together** | Separate scores keep each method readable and make agreement visible. The shared gate stops a large attribute gap that is really a substitute from being recommended. The fused rank gives the final list (§7) |
| 12 | Thresholds are **anchored by calibration**, not by percentile alone | Pure percentile thresholds put a fixed fraction of candidates into each label whatever the reality. Before Pat's session, Staples-vs-Staples pairs give a weakly supervised anchor (§6.6) |

---

## 1. Objective, mandate and scope

### 1.1 Business mandate (from *Objective & Goal.docx*)
- Stakeholder: Pat, Head of Staples Marketplace. Target: Q1 2027 planning.
- **Core-adjacent ("White Chair") logic:** design and lifestyle extensions of core items Staples lacks, for example aesthetic-first office seating that competes with Wayfair without diluting core B2B.
- **1P protection:** do not cannibalize Staples' existing assortment.
- **Constraint:** external market data only (competitor and Staples.com crawls). No internal Staples data (sales, margin, traffic).
- **Levers:** quick wins, basket attach (1–2 marketplace items per checkout), a predictable 6–18 month commission roadmap.

### 1.2 Analytical objective
For every in-scope analysis node *n*:
- **Where are the gaps?** Which attribute values and archetypes are over-represented at the competitor relative to Staples (Method 2, plus Method 1's vector opportunity)?
- **Are they safe and on-brand?** Would each candidate substitute a Staples item (CRS), does it fit Staples' customer (AAS), and is it priced to trade up or to undercut (PPR)? (Method 1)
- **What should be added?** Name the archetypes and exemplar SKUs, with brand and seller notes and a confidence level.

### 1.3 Scope: decided by the data, not written into the method
The current samples are a starting point. They will be enriched with more categories, or with completely different ones, and more competitors may follow. **Nothing in the method or code may be hard-wired to chairs, seating or desks.** Scope is worked out on every run:

- **In-scope analysis nodes:** every Staples analysis node that meets the minimum size on *both* sides after mapping (§5.3). With today's data that is Furniture → Chairs & Seating and Desks, because that is where the Wayfair sample overlaps. Tomorrow it could be Labels or Lighting, with no code change.
- **Staples-only nodes** (Staples data but no competitor data, for example Labels and Shipping today) are listed in a coverage report with their family counts and not scored.
- **Unmapped competitor products** go to a **new-node backlog** (today: bean bags, theatre seating, patio dining, door mats, restaurant sets and similar). They are not scored here. That is a category-level (CL PoC) question, not an assortment question.
- **Competitors** are added through a retailer adapter (a column mapping plus page-to-tree metadata). Wayfair, consumer and Professional pages together, is the only one today.

§1.4 lists what is generated per node and what stays fixed.

**First vertical slice:** a run-time choice, not a code path. With today's data it is *Chairs & Seating → Accent & Waiting Room Chairs* (the White Chair thesis node: 199 Staples SKUs against Wayfair *Accent Chairs*, *Custom Accent Chairs*, *Waiting Room & Reception Chairs* and *Reception Seating*), with *Office Chairs* as the contrast node. Then all other in-scope nodes.

### 1.4 Scope-agnostic design: what is generated per node vs fixed
Anything category-specific is **generated from the data** (usually an LLM draft from a sample, then a quick human review) and stored as a versioned config file per node or per L2. The code only reads those files. Examples in this document (chair types, desk shapes, boucle) are illustrations, not rules.

| Component | Generic (same everywhere) | Generated per L2 / node from data |
|---|---|---|
| Ingest (S0) | Retailer adapter: column map, price parser, PII stripping | New retailer = new adapter entry |
| Families (S1) | Grouping rule (brand + series + name stem, colour segment removed) | Colour-word list is extended from the observed spec colour values |
| Analysis nodes (S2) | Pseudo-L5 split rule (size + coverage thresholds) | The split facet is chosen automatically: the Staples spec key matching `*Type*/*Shape*/*Design*` with the highest coverage × entropy |
| Mapping (S3) | Classifier + LLM adjudication cascade | Competitor page → candidate Staples leaves crosswalk (LLM draft from page and tree names, human review) |
| Schema (S4) | Tier 1 (universal) and Tier 3 (lifestyle) field *definitions* | Tier 2 functional attributes and all controlled vocabularies: seeded from that L2's Staples spec keys, extended by LLM schema induction on a competitor sample, then frozen |
| Extraction (S5) | Prompt template, validation against Staples specs | Spec-key → schema synonym table (LLM-proposed for new keys) |
| DFI | 1–5 scale and validation protocol | Rubric anchors per L2 ("design-forward" means something different for labels than for chairs) |
| Archetypes (S6) | Grid + back-off + HDBSCAN validation | Candidate grid facets per node (LLM proposes merchant-relevant facets; information criteria pick 3–5) |
| Scores (S7–S9) | All formulas, decision tree, weights | Thresholds and PPR bands, calibrated per L2 |
| Recruitability (S10) | `brand_on_staples` logic | Competitor house-brand list, per retailer |

Every generated artefact carries a `generated_by` (model and prompt version) and `reviewed_by` field. Categories with no review yet run with a **provisional** flag shown in all outputs.

---

## 2. Data audit (what is actually in the files)

Profiled on 2026-09-27. These findings drive most of the design choices.

### 2.1 Staples: `Staples product level dataset latest - Samples.xlsx`
- 31,979 rows, 26,861 unique `product_id`. 4,781 are exact duplicate rows and 141 products sit in more than one leaf.
- Columns: `L1–L4`, `Final level Category`, `product_id`, `product_name`, `price`, `product_link`, `model_number`, `reviews`, `description`.
- Chairs & Seating plus Desks: **4,725 unique SKUs → about 2,860 name-stem families**. Colour variants are separate SKUs (for example *Arozzi Arena Gaming Desk* has 10 colour SKUs).
- **Coverage against site counts** (navigation tree `Count`): Office Chairs 982/1,008, Accent 199/229, Counter Stools 602/617, Office Desks 1,266/1,501, Sit & Stand 396/576. Treat as a near-census. The thin leaves (Big & Tall 39/79, Hutches 40/77) need a re-crawl or a caveat.
- **`description` is a serialized Python dict** `{paragraph, bullets, specification[{name, value, grpName, dscr}]}`. It parses 100% with `ast.literal_eval`. Spec coverage in Chairs and Desks: True Color 97%, Furnishing Material 94%, W/D/H 91%, Assembly 87%, Warranty 78%, Weight 70%, Series/Collection 63%, Base Material 54%, Arm Type 52%, **Furnishing Style 47%**, Chair Type ~20%, Desk Shape 34%.
- `reviews` looks like `"4.57 stars ( reviews)"`: rating is present for 61% of rows and **the review count is always empty**.
- `price` is a `"$xx.xx"` string with 152 nulls. Chairs and Desks median is $355.
- Brands: Flash Furniture (1,214), Boss, Bush, Offices To Go, HON and others. **Staples private label is about 1.4%**, so "1P" here means *items Staples sells*, not *Staples-brand items*.
- The navigation tree has no L5 (L5 is empty for all 1,877 rows). The deepest level is L4.

### 2.2 Wayfair: `Wayfair product level dataset latest - Samples.xlsx`
- 12,255 rows, 11,468 unique `product_id`. 787 are duplicate rows and 156 products appear on more than one listing page.
- Columns: `category` (the listing page, with "&" mangled to 3 spaces), `product_id`, `product_name`, `price`, `url` (60% carry a `piid` variant id), `vendor`, `selected_choice` (the displayed variant: colour or material, 57% filled), `description`.
- **No structured specs, rating, review count, image URL or option list.**
- `description`: median 595 characters. **41% have one customer review appended after `" | "`**, including reviewer name, city and date. That is PII: strip it and never show it.
- `vendor`: 612 values, mostly **Wayfair house brands**, with mojibake (`Latitude RunÂ®`).
- `price`: `$` string. 12 rows are `"per item"` (pack pricing). Median $230.
- Listing pages come from three Wayfair trees (Consumer, Professional B2B, Professional Shop-by-Space). The Wayfair navigation file has **no item counts**, so Wayfair site totals are unknown.
- Pages with no Staples counterpart: Outdoor Door Mats (962), Bean Bag Chairs (864), Patio Dining Chairs (814), Restaurant Table & Chair Sets (697), Theater Seating (198), Chiavari (19), Office Cubicles (48).
- Quality: some descriptions do not match their title (an armchair described as "this office chair is ergonomic…"), and some products are on the wrong page (an armchair listed under *Office Chairs*).

### 2.3 Other inputs
- `1. Staples_Navigation_Tree_repaired.xlsx`: 1,877 nodes, a canonical path per node, a Cross-Listings sheet, and per-leaf counts.
- `3. Wayfair_Navigation_Tree_v3.xlsx`: 2,231 rows across Consumer, Pro and Shop-by-Space trees, plus a Unique Categories sheet (1,308).
- `Vlookup - Manual Trail … [Failed - DO NOT USE].xlsx`: name-based page-to-leaf mapping. 25 of 32 pages got no match. It is kept only as evidence for change #1.
- The CL PoC V2 code (`1. CL PoC/V2 …/code/`) has reusable parts: the encoder bake-off harness, local Qdrant, the calibration utilities and the figure style.

### 2.4 What the data cannot support yet (and the fallback used)
| Needed for | Missing | Fallback in this method | Data request (§11) |
|---|---|---|---|
| Demand weighting | Review counts on both sides, rank position | Supply-side proxy (Wayfair family depth) with a caveat. The D term is off by default | Wayfair review count and rating; listing rank position |
| Image-based DFI | Image URLs | Text-based DFI from an LLM rubric plus text anchors | Primary image URL on both sides, then an image DFI |
| Wayfair colour breadth | Option lists (only `selected_choice`) | Colours from title, choice and description as a lower bound | Crawl the option list per family |
| Contextual adjacency (PPR graph) | Also-viewed and complement widgets | A_ctx from extracted use-context overlap | Optional: "customers also viewed" crawl |
| Absolute depth | Wayfair site totals and sampling design (the sample is a convenience set of successful scrapes) | Share-based metrics with credibility only | None for now (answered in §11) |

---

## 3. Critique of the proposed methodologies

### 3.1 Chat version (steps 1–7)
| Step | Assessment | Enhancement |
|---|---|---|
| 1 Nav and SKU gathering | Done, but the Wayfair sampling design is undocumented and there are no demand signals | Record the sampling method and site totals. Add review count, rating, option list and image URL (§11) |
| 2 Map competitor navigation to Staples L1–L5 | Mapping nav to nav is the wrong grain. Wayfair pages are rooms and uses, many-to-many, and they mis-shelve products. There is no Staples L5 | Map **each Wayfair product family** to a Staples **analysis node**. Create pseudo-L5 nodes from Staples' own type facets (§5) |
| 3 Finalise attribute list | The right idea, but a flat list mixes universal, functional and lifestyle attributes, and it lacks controlled vocabularies | A 3-tier schema with closed vocabularies, seeded from Staples' own spec keys and frozen after induction (§6.1) |
| 4 Extract from descriptions | Extraction differs by source: Staples has specs, Wayfair has text | Same instrument for every attribute that feeds a gap. Staples specs become **free ground truth** for scoring the LLM extractor (§6.2) |
| 5.1 Archetype = combination of all attributes | With about 15 attributes the cross-product is millions of cells and almost all are empty, which makes the result generic, as you noted | A **facet grid of 3–5 high-information facets per node**. The remaining attributes describe the archetype rather than define it. Validate with HDBSCAN (§6.3) |
| 5.1 Embed archetypes | Embedding a synthetic archetype string collapses to generic text | Embed **SKUs as canonical cards**. The archetype vector is the centroid of its members (§6.4) |
| 5.1 "Vector opportunity score" | Not defined | **VOS** = Method 1's final score, combining vector whitespace (VW, a balanced k-NN two-sample density test), AAS, AD and a CRS penalty (§6.5.6) |
| 5.1 AAS / CRS / PPR | Sound skeleton. Carries the doc's issues (next table) | Fixed in §6.5 |
| 5.2 DG, CR | Count-based, and both break under sample-vs-census comparison and at Staples = 0 (CR is undefined, DG is dominated by sample size) | Smoothed **log share ratio** with a Bayesian credible filter (§6.7.1) |
| 5.2 PPG (Wasserstein) | Unsigned, so it cannot distinguish "Staples lacks premium" from "Staples lacks value" | Keep the magnitude, add a sign, add a price-band coverage gap. Direction goes to PPR (§6.7.2) |
| 5.2 CG (colour gap) | Only light and neutral share, so it misses other colours | Generalise to an **Attribute Distribution Gap** for every categorical attribute (JSD plus signed value gaps). Colour is one instance of it (§6.7.3) |
| 5.2 DFG | Depends on image DFI, and there are no images | Text DFI now, image DFI once URLs are crawled (§6.2.4) |
| 5.2 Brand fragmentation | Measures Wayfair's labelling | Reframed as recruitability at the SKU stage (§8.3) |
| 5.2 TG weights 0.35/0.25/0.20/0.20 | Arbitrary, and `pct()` has no stated population | Percentile within an L2 peer group across archetypes. Dirichlet weight-sensitivity with a rank-stability report (§6.7.5) |
| 6 Insights | Good | A per-node view in the HTML report, and a method-agreement tier on every recommendation (§7, §9) |
| 7 Recs and vendors | Good | Exemplars chosen by centroid proximity plus low SKU-CRS plus diversity (MMR), each shown with its nearest Staples SKU (§8) |

### 3.2 Document version (Track B/C)
**Keep:** facet-grid archetypes validated by clustering; nameability test; LLM two-pass extraction with confidence and evidence span; schema freeze; dedup; the core AAS/CRS/PPR idea (mean-pooling for adjacency, max-pooling for cannibalization); UNDERCUT hard reject; VERTICAL EXTENSION hold; Pat calibration session; "no raw cosine in deliverables"; re-fit on encoder change.

**Fix:**
1. **TRADE-UP unreachable.** The `po` term inside CRS cancels PPR ≥ 1.5, so CRS is now price-agnostic.
2. **Label grid not exhaustive.** Replaced with an ordered decision tree (§6.6).
3. **A_ctx (PageRank on complement graph) and `sub` (also-viewed).** No data for either. Dropped, and weights renormalised. A_ctx is replaced by use-context overlap.
4. **SigLIP DFI.** No images. Text DFI for now (§6.2.4).
5. **Stratified facet-enumeration sampling.** It had not been done: the sample is 32 listing pages. Share metrics remove the need for inclusion-probability reweighting. If rank position is available, rank-weighting is optional.
6. **Revenue-at-risk and net contribution (5.5).** They need 1P margin and demand, which we do not have. Moved to an optional scenario appendix, not part of the score.
7. **Percentile-only thresholds.** Anchored by calibration pairs (§6.6).
8. **Dedup via manufacturer.** Manufacturer is unobservable behind Wayfair house brands. Dedup uses a card-embedding, dimension and price fingerprint instead (§4.3).
9. **Retailer set of 5 with a style reference corpus.** Descoped to Wayfair for now. The adapter pattern keeps the pipeline open to more competitors.

---

## 4. Pipeline overview

```
 S0  Ingest & clean ─► S1  Family grouping & dedup ─► S2  Analysis-node definition (pseudo-L5)
                                                              │
 S3  Product-level mapping  (Wayfair family → Staples node | NONE → backlog)
                                                              │
 S4  Attribute schema (3 tiers, frozen) ─► S5  Extraction + normalisation + QA
                                                              │
 S6  Archetype construction (facet grid → cluster validation → naming)
                                                              │
        ┌─────────────────────────────┴──────────────────────────────┐
 S7  METHOD 1 · Vector space                         S8  METHOD 2 · Attribute gaps
     VW · AAS · CRS · AD · PPR → labels                   LSR · PPG · CG · MSG · DFG
     final score: VOS                                     final score: TG
        └─────────────────────────────┬──────────────────────────────┘
 S9  Integration: shared safety gate → fused re-rank of VOS & TG + agreement tier
                                      │
 S10 SKU stage: exemplars, nearest-Staples comparison, brand & recruitability
                                      │
 S11 Outputs: node scorecards · archetype table · attribute insights · SKU recs · report
```

Every stage writes a versioned intermediate table, so each can be re-run and inspected on its own.

---

## 5. Stages S0–S3: data preparation, analysis nodes and mapping

**What these stages are for.** They produce no recommendations. They make the Staples-vs-competitor comparison fair. Every later stage compares "share of X at Wayfair" with "share of X at Staples" *within the same shelf*, which needs three things settled first:

| Stage | Question it settles | What goes wrong without it |
|---|---|---|
| S1 Families | *What is one product?* | Staples lists each colour as its own SKU (one desk = 10 rows) and Wayfair lists one product with options. Staples' depth is inflated about 1.7× and colour-heavy lines look like whole archetypes |
| S2 Analysis nodes | *Which shelf are we comparing on?* | Big leaves such as *Office Chairs* (about 900 SKUs) mix executive, task and gaming chairs, so gaps average out. S2 splits them using Staples' own type facet |
| S3 Mapping | *Which shelf does each Wayfair product belong on?* | Wayfair pages are rooms and uses, not product types, so page-name mapping failed. S3 places each product on a Staples shelf, or on a NONE backlog |

Side outputs that are useful on their own: colourway counts per family (S1, used for STYLE-EXTENSION), identical products already carried (S1, prevents false whitespace), and the backlog of Wayfair products with no Staples shelf (S3, handed to the category-level workstream).

### 5.1 S0 Ingest and clean
- Parse prices to float. Flag `per item` as pack pricing and use the unit price.
- Parse the Staples `description` dict into `paragraph`, `bullets` and `specs` (a long table: `product_id, name, value, grpName`).
- Parse the Staples rating (`"4.57 stars"` → 4.57). The review count is null and is kept null.
- Wayfair: split `description` at the first `" | "`. The left side is the product description. The right side is review text, which is **dropped** (PII) and not used in any score.
- Fix mojibake (`ftfy`, or a latin-1 → utf-8 round-trip). Normalise the Wayfair page name (triple space → " & ").
- Assign each Staples SKU its canonical leaf from the navigation tree, keeping cross-listings as secondary leaves.

### 5.2 S1 Family grouping and dedup
- **Staples family:** group by `(brand, Series/Collection spec if present, name stem)`. The name stem is the product name with the trailing colour segment and `(MODEL)` removed. Resolve ambiguous merges with a card-embedding cosine > 0.95 plus the same W/D/H within 5%. Keep `n_colourways` and the colour list per family.
- **Wayfair family:** `product_id` (for example `W112121311`, `ABFR3973`). Collapse repeated rows and keep every observed `selected_choice` and `piid`.
- **Cross-retailer identical products** (the same product on both sites, often behind a house brand): block on `(node, W/D/H within 5%, price within 15%)`, then card-embedding cosine > 0.93. Audit 100 flagged pairs by hand. Identical pairs are flagged `ALREADY_CARRIED`, so they are never reported as whitespace.
- **Counting unit** for every share: the family. Colourway breadth is a separate attribute (§6.7.4).

### 5.3 S2 Analysis-node definition (the granular level)
- Start from the Staples **leaf** (L3 or L4).
- **Pseudo-L5 split:** if a leaf has ≥ 250 families and a Staples type facet with ≥ 80% coverage after extraction, split on that facet. Examples:
  - *Office Chairs* → Executive / Task / Computer & Desk / Manager / Conference (Chair Type).
  - *Office Desks* → Computer / Workstation / Executive / Writing / Reception (Desk Type) or by shape (Rectangular / L / U).
  - *Accent & Waiting Room Chairs* → Guest / Lounge / Accent / Reception.
- **Minimum size:** ≥ 30 Staples families and ≥ 30 Wayfair families for full analysis. Below that the node is **thin** and gets a descriptive profile only, rolled up to its parent for scoring.
- Output: `analysis_nodes.csv` (node_id, path, parent, split facet, Staples and Wayfair family counts).

### 5.4 S3 Product-level mapping (Wayfair family → Staples node)
A three-stage cascade:
1. **Page prior (crosswalk).** Hand-map the 32 Wayfair pages to *candidate* Staples leaves, many-to-many, with an LLM draft and human review. *Office Chairs* → {Office Chairs, Big & Tall, Gaming, Drafting}. *Meeting Space Seating* → {Accent & Waiting, Stacking & Folding, Benches, Breakroom}. *Outdoor Door Mats* → {∅}.
2. **Product classifier.** Train a classifier (k-NN or logistic on card embeddings plus key attributes) on **Staples SKUs labelled with their own leaf**: 4.7k free labels. Predict a probability over the candidate set for each Wayfair family. Watch for domain shift, since Staples cards train the model and Wayfair cards are scored.
3. **LLM adjudication** when the top-1 margin is below *m*, the classifier disagrees with the page prior, or the family is in a room or space page. Use a closed candidate list plus `NONE` and a one-line reason.
   - Model: Sonnet 5 for adjudication and Haiku 4.5 for bulk, both with enforced JSON output.
- **Pseudo-L5 assignment** uses the extracted type attribute (S5).
- **Gold set:** 200 Wayfair families, stratified by page, labelled by two people. Target ≥ 90% top-1 accuracy on in-scope families and ≥ 85% precision on `NONE`. Report the confusion matrix.
- `NONE` results go to `new_node_backlog.csv` (with the nearest Staples L2 for adjacency), which is handed to the CL workstream.

---

## 6. Stages S4–S8: attributes, archetypes and the two methods

### 6.1 S4 Attribute schema (3 tiers, closed vocabularies)
Seed the schema from **Staples' own spec keys**. The result stays in the Staples merchandising team's vocabulary, and attributes Staples does not track (such as style on 53% of SKUs) become insights in their own right. Extend it with lifestyle attributes. Run schema induction on about 300 Wayfair families per L2 (Sonnet 5 proposes values with frequencies), then review by hand and **freeze**. Every field allows `unknown`. Nothing is imputed.

**Tier 1: universal (all nodes)**
| Attribute | Type / vocabulary |
|---|---|
| price | float (unit price) → node-specific price band (pooled quantiles rounded to merchant breakpoints) |
| brand_display | string (Wayfair house-brand flag) |
| colour_raw → colour_family | about 14 families (black, white, grey, beige/cream, brown/wood-tone, navy, blue, green, red/pink, yellow/orange, purple, metallic, multi/pattern, clear) |
| colour_tone | light-neutral / dark-neutral / warm-neutral / bold-saturated / pattern |
| material_primary, material_class | node vocabulary (for example mesh, bonded leather, faux leather, genuine leather, fabric, velvet, boucle, linen, vinyl, molded plastic, solid wood, engineered wood, metal, rattan/wicker, glass) |
| finish / frame_material / base_material | controlled vocabulary |
| style_family | modern, contemporary, mid-century, industrial, traditional, farmhouse/rustic, glam, scandinavian, minimalist, transitional, boho/coastal, gaming, utilitarian/corporate |
| dims W/D/H (in), weight (lb), weight_capacity (lb) | float, then size_class (compact / standard / oversize) |
| assembly, warranty_years, pack_qty, commercial_grade | enum / float / int / bool |
| certifications | BIFMA, GREENGUARD, FSC, and similar (multi-label) |

**Tier 2: node-specific functional attributes (these define substitution)**
Generated per L2 (§1.4). The two lists below are what we expect induction to produce for today's data. They are illustrations, not hard-coded fields.
- *Chairs:* form_factor/chair_type, mechanism (fixed / tilt / synchro / multi), arm_type, base_type (casters / legs / sled / pedestal / 4-star), swivel, height_adjustable, back_height, lumbar, headrest, recline, stackable, folding, seat_material, upholstery_texture.
- *Desks:* desk_type, shape (rectangular / L / U / corner / bow), sit_stand (none / manual / electric), top_material, width_band, drawers, storage, keyboard_tray, cable_mgmt, frame_material.

**Tier 3: experiential and lifestyle attributes (LLM-inferred, the "vibe")**
| Attribute | Type |
|---|---|
| aesthetic_tags | multi-label from a closed list of about 25 (sculptural, cozy, organic/curved, tufted, channel-tufted, minimal, statement, retro, luxe, natural/warm, playful, …) |
| design_forward_index (DFI) | 1–5 rubric, then scaled to 0–1 (§6.2.4) |
| use_context | home office, corporate office, reception/lobby, breakroom/café, conference, education, healthcare, hospitality, gaming, kids/teen, outdoor (multi-label) |
| end_user_segment | big & tall, bariatric, petite, kids, teens, gamers, students, remote workers, facilities buyers |
| key_benefits | ergonomic, space-saving, easy-clean/antimicrobial, sustainable, customisable, portable, … |
| positioning | value / mid / premium (from language, not price) |

Tier 1 and Tier 2 feed archetypes and substitution. Tier 3 feeds the lifestyle gap and naming.

### 6.2 S5 Extraction, normalisation and QA

#### 6.2.1 Sources and precedence
- **Staples:** deterministic mapping of spec keys to the schema through a synonym table (`True Color`/`Furnishing Color`/`Color Family` → colour, and so on). The LLM fills only the remaining gaps and Tier 3.
- **Wayfair:** the LLM reads the title, `selected_choice` and the cleaned description. Precedence when sources conflict: **title > selected_choice > description**, and the conflict is logged. This handles the mismatched descriptions seen in the audit.

#### 6.2.2 The same instrument on both sides (critical)
- Tier 3 (style, aesthetic, DFI, use-context) and `style_family` are extracted **by the same LLM prompt from the same kind of input (title plus a text summary) on both sides**, even when Staples has a spec value. A Staples "Furnishing Style: Contemporary" and an LLM's "mid-century" come from different instruments, and the gap would reflect that difference rather than the assortment.
- For Tier 1 and Tier 2 physical facts (dimensions, material, colour, arm type), Staples specs are the truth and Wayfair values come from the LLM. We validate that the two agree (next step).

#### 6.2.3 Extractor validation: Staples specs as free ground truth
- Hide the specs and run the Wayfair-style text-only extractor on 300 Staples families, stratified by node.
- Score per-field accuracy against the specs (exact match for enums, ±5% for numbers).
- **Rule:** a field below 85% accuracy is fixed through the prompt or vocabulary, or else excluded from gap metrics and reported descriptively only.
- Tier 3 has no spec truth. Two people label 100 Wayfair and 50 Staples families on 6 key fields. Targets: Cohen's κ ≥ 0.6 between raters and LLM agreement ≥ 0.8 of the human agreement.

#### 6.2.4 Design-Forward Index (DFI)
- **Now (text):** Sonnet 5 applies a 1–5 rubric with anchored examples:
  - 1 = corporate or utilitarian (black mesh task chair)
  - 3 = transitional or home-office friendly
  - 5 = residential designer or sculptural statement piece

  The input is the title, the description summary and the Tier 3 tags. The output is a score and a rationale. We also compute an embedding contrast, `cos(card, A_design) − cos(card, A_utility)` with prompt anchors, as a cross-check (Spearman ≥ 0.6 expected).
- **Later (image):** once image URLs are crawled, add a SigLIP-2 contrastive score as the doc proposed. DFI becomes the mean of the text and image scores after z-normalisation. Validate against 200 human-rated items (κ ≥ 0.65).

#### 6.2.5 Normalisation
- Colour goes to a family and tone (dictionary plus LLM fallback).
- Materials go to a class. Units are converted to inches and pounds.
- Price bands are computed per node on pooled prices.
- All values carry `value, confidence, evidence_span, source(spec|llm)`. Fields with confidence below 0.7 go to a review queue, and only the 6–8 fields that feed archetypes are reviewed.

#### 6.2.6 Operational notes
- About 2.9k Staples and about 8k in-scope Wayfair families × about 1.5k tokens is roughly 15–20M input tokens. Use the Batch API, prompt caching for the schema, and Haiku 4.5 for bulk.
- Prefer JSON-schema-enforced (tool-use) output.
- Cache raw LLM responses to disk so that re-runs cost nothing.

### 6.3 S6 Archetype construction (grid first, clusters as validation)
1. **Choose grid facets per node (3–5).** Candidates come from Tier 1 and Tier 2 plus price band and style_family. Select by:
   - merchant relevance, from a pre-agreed list per node;
   - under 40% unknown on both sides;
   - information: normalised entropy and mutual information with the HDBSCAN clusters from step 4.

   Example for *Accent & Waiting Room Chairs*: `form_factor × material_class × colour_tone × style_family × price_band`.
2. **Assign families to cells.** Merge cells hierarchically: if a cell has fewer than 10 pooled families (or fewer than 5 Wayfair families), drop the least-informative facet for that branch and merge upward. Stop when every live cell meets support. The expected result is about 8–25 archetypes per node.
3. **Descriptors.** Non-grid attributes (aesthetic tags, use-context, DFI distribution, dimensions) are summarised per archetype. They describe it but do not define it.
4. **Validate** with UMAP (n_components = 15–20, for clustering only) on fused card embeddings, then HDBSCAN (`min_cluster_size` scaled to node size, `leaf` selection). Report AMI/ARI between the clusters and the grid cells.
   - If ARI is below 0.4, inspect the split clusters and add the missing facet (swivel, boucle, and so on).
   - Bootstrap 20 × 80% resamples; target mean ARI ≥ 0.6.
   - Cross-check with k-prototypes (Gower).
5. **Name** each archetype with Sonnet 5, using its facet values plus 5 medoid titles. Run the nameability test: a merchant-literate colleague should be able to describe each of 5 cards in under 30 seconds without seeing the facet table.

**Note on "all attributes combined":** the full attribute vector still enters the analysis through (a) the fused embedding in Method 1 and (b) the per-attribute distribution gaps in Method 2. The grid only decides how archetypes are *named and counted*.

### 6.4 Embedding space (shared by S3, S6 and Method 1)
- **Canonical card** (identical template on both sides; no price; no raw spec dump):
  `"{node} | {form_factor} | {title_clean} | colour: {colour_family}/{colour_tone} | material: {material_class} | style: {style_family} | features: {key T2 values} | vibe: {aesthetic_tags} | use: {use_context} | {≤60-word LLM summary}"`
- **Views:**
  - (i) a text embedding of the card;
  - (ii) an attribute vector (one-hot Tier 1 and Tier 2 categoricals plus robust-scaled dimensions);
  - (iii) later, an image embedding.

  Fused = `[α·t, β·a]`, L2-normalised. α and β are tuned on the mapping gold set.
- **Encoder:** pluggable, with a bake-off harness reused from CL PoC V2. Candidates are `bge-large-en-v1.5`, `gte-large`, `e5-large-v2` and an API embedding model. The selection metric is mapping top-1 plus duplicate-pair recall. All thresholds are re-fitted when the encoder changes.
- **Source-mixing diagnostic:** on the known cross-retailer identical pairs (S1), the partner must be in the top-3 neighbours at least 80% of the time. If it is not, the card still carries source style and needs fixing.
- **Vector store:** local Qdrant (reused from CL PoC), with one collection per retailer and node/attribute payload filters. At about 11k vectors this is a convenience, not a requirement; numpy or FAISS would be equivalent.

### 6.5 S7 METHOD 1: Vector space (VW, AAS, CRS, AD, PPR → VOS)
All components are computed per Wayfair family *c* in analysis node *n*, then aggregated to archetypes. They combine into Method 1's final score, **VOS** (§6.5.6).

#### 6.5.1 Vector Whitespace (VW): how empty Staples' region is
- Pool the Staples and Wayfair families of node *n*. Give Staples points a weight `w_S = N_W / N_S`, so that when the two distributions are identical the expected Staples share of any neighbourhood is 0.5.
- For each Wayfair family *c*, the local Staples coverage is `LSC(c) = Σ_{j∈kNN(c)} w_j·1[j∈S] / Σ_{j∈kNN(c)} w_j`, with k = 15 (sensitivity at 10 and 25).
- Whitespace: `ws(c) = clip(1 − 2·LSC(c), 0, 1)`. A value of 1 means no Staples item is nearby. A value of 0 means Staples is at least as dense as Wayfair here.
- **Archetype level:** `VW(a) = mean_{c∈a} ws(c)`, reported with a bootstrap CI. It is a classifier two-sample test in disguise, so it is interpretable and independent of the attribute grid.

#### 6.5.2 Adjacency Affinity Score (AAS): does it belong on Staples?
Within an existing node, adjacency is partly given. AAS now measures **fit with Staples' customer and catalog**:
- `A_sem(c)` = mean cosine of the top-50 Staples families across the **whole L2** (for example all of Chairs & Seating). Mean-pooling asks whether the candidate sits in Staples' region.
- `A_ctx(c)` = weighted Jaccard between c's `use_context ∪ end_user_segment` and the node's Staples tag distribution. A kids' nursery rocker scores low in *Accent & Waiting Room*, and a reception lounge chair scores high.
- `AAS = 100 × (0.6·pct_L2(A_sem) + 0.4·A_ctx)`.
- Dropped: the PageRank complement term (no co-view data) and the separate JTBD term (folded into A_ctx).

#### 6.5.3 Cannibalization Risk Score (CRS): would it take a Staples sale? (price-agnostic)
- `S_max(c)` = mean cosine to the **top-3** Staples families **in the same analysis node**, on a *functional* embedding view (a card built from Tier 2 plus form_factor, with no colour, style or aesthetic tags). Max-pooling asks whether one specific Staples product is replaced.
- `FI(c)` = functional identity: the share of node-specific Tier 2 core attributes that equal those of the nearest Staples family (continuous 0–1, not the doc's binary `ff`).
- `CRS = 100 × (0.6·cal(S_max) + 0.4·FI)`, where `cal()` maps cosine to a probability of substitution via the calibration of §6.6, not a raw percentile.
- **Excluded by design:** price (carried by PPR) and aesthetics (carried by AD below). This is what makes STYLE-EXTENSION and TRADE-UP reachable.

#### 6.5.4 Aesthetic Delta (AD): is it visibly different?
- `AD(c) = 1 − cos_aesthetic(c, nearest Staples family)`, on an aesthetic view (colour family, tone, material_class, style_family, aesthetic tags, DFI).
- Report also `ΔDFI = DFI(c) − DFI(nearest Staples)`.

#### 6.5.5 Price Position Ratio (PPR): which direction is the price?
- `PPR(c) = price(c) / median price(top-3 Staples neighbours on the functional view)`.
- Initial bands, fitted per L2 in calibration: **≥ 1.5 trade-up · 0.85–1.5 parity · < 0.85 undercut**.
- PPR is **not** part of any score. It is a direction, not a magnitude, and it acts only through the decision labels (§6.6).

#### 6.5.6 Vector Opportunity Score (VOS): Method 1's final score
One number per archetype that summarises the vector method: *is there empty space here, does it fit Staples, is it visibly different, and is it safe?*
```
VOS(a) = 100 × [ 0.45·pct(VW(a)) + 0.30·AAS(a)/100 + 0.25·pct(AD(a)) ] × (1 − CRS(a)/100)^γ
```
- VW is the opportunity signal, AAS the brand fit and AD the visible difference (median over members). CRS dampens the result multiplicatively: open space that substitutes a Staples item is not an opportunity. γ = 1 by default, with sensitivity at 0.5 and 2.
- `pct()` = percentile rank across archetypes in the same L2 peer group, the same population TG uses, so the two scores are on comparable 0–100 scales.
- Weights ship with the same Dirichlet sensitivity as TG (§6.7.6).

### 6.6 Decision tree (exhaustive, ordered; applied per candidate family)
```
0. mapped to NONE / thin node                       → BACKLOG (not scored)
0b. flagged ALREADY_CARRIED (identical product)      → EXCLUDE
1. AAS < T_A_low                                      → OFF-BRAND            (reject)
2. CRS ≥ T_C_high  (functionally substitutes a Staples item)
     2a. PPR < 0.85                                   → UNDERCUT             (hard reject)
     2b. PPR ≥ 1.5 and (ΔDFI ≥ δ_D or material upgrade) → TRADE-UP          (approve, margin note)
     2c. AD ≥ T_AD (new colour/material/style)         → STYLE-EXTENSION     (approve; the "flavours" case)
     2d. otherwise                                     → SUBSTITUTE          (reject)
3. T_C_low ≤ CRS < T_C_high                           → REVIEW               (goes to Pat calibration queue;
                                                                              AD ≥ T_AD & AAS ≥ T_A_high → REVIEW-lean-approve)
4. CRS < T_C_low
     4a. AAS ≥ T_A_high                               → CURATE               (approve; true whitespace)
     4b. T_A_low ≤ AAS < T_A_high                      → EDGE / VERTICAL EXT. (hold for phase 2)
```
Each branch is exclusive and every case gets exactly one label. The "(DATA ERROR)" case (low AAS with high CRS) is caught by rule 1 first. It is still counted as a diagnostic, and a count above 2% triggers a review of the mapping and embeddings.

**Threshold setting (two steps):**
1. **Weak supervision (before Pat).** Build positive and negative pairs from Staples alone:
   - *positives* = different colourways of the same Staples family, and near-identical cross-brand listings (manual 50);
   - *negatives* = Staples families from different analysis nodes, and same-node families with different form_factor.

   Fit a logistic `cal(S_max)`. Set the first cut of `T_C_high` and `T_C_low` at P(substitute) = 0.7 and 0.3. Set the AAS thresholds from the distribution of Staples' own families, whose AAS against the rest of Staples marks "on-brand" (for example the 10th percentile of Staples' own AAS gives `T_A_low`).
2. **Pat calibration session** (from the doc, kept). Use 60 side-by-side pairs, stratified across the CRS range with extra pairs in the REVIEW band, 20 of them stratified by PPR. Hide the scores. Ask "would this take sales from that?" and "would you rather carry it at this price?". Refit `cal()`, the PPR bands and T_AD, then show Pat the resulting label distribution.

**Hygiene rules (kept from the doc):** no raw cosine in any deliverable; all thresholds are re-fitted when the encoder changes; PPR bands are set per L2.

**Archetype roll-up:** an archetype's label mix is the share of member families per label. Its **Safe Share** is the share labelled CURATE, STYLE-EXTENSION or TRADE-UP. Archetype CRS, AAS and PPR are the member medians. An archetype is **shortlist-eligible** when Safe Share ≥ 0.5 and UNDERCUT + SUBSTITUTE < 0.3.

### 6.7 S8 METHOD 2: Attribute-level gap metrics
All metrics are computed per node, both per archetype and per attribute value, on **family shares**.

#### 6.7.1 Share gap (replaces Depth Gap and Coverage Ratio)
- For archetype *a* in node *n*: `p_W(a) = n_W(a)/N_W(n)` and `p_S(a) = n_S(a)/N_S(n)`.
- Beta posteriors with a Jeffreys prior: `p_X ~ Beta(n_X + 0.5, N_X − n_X + 0.5)`.
- **Log Share Ratio** `LSR(a) = log(E[p_W] / E[p_S])`. This stays finite when Staples has 0.
- **Credibility:** `Pr(p_W > p_S)` by Monte Carlo. The gap is *credible* when it is ≥ 0.9.
- **Coverage flags:** `ABSENT` when `n_S = 0` and `n_W ≥ 5`; `THIN` when `p_S < 0.25·p_W`.

#### 6.7.2 Price Position Gap (PPG)
- `PPG_mag(a) = W1(log price_W, log price_S)` within the archetype (0 if either side has fewer than 3 families; then the archetype-level band gap below is used). Direction: `PPG_dir = sign(median_W − median_S)`.
- **Node-level price-band coverage:** per price band *b*, the share gap `p_W(b) − p_S(b)` with credibility. This is where "Staples has no $120–250 accent chairs" shows up.

#### 6.7.3 Attribute Distribution Gap (ADG): colour gap generalised
For each categorical attribute *k* (colour_family, colour_tone, material_class, style_family, aesthetic_tags, use_context, size_class, and Tier 2 features):
- `ADG_k(n) = JSD(P_W^k ‖ P_S^k)`, on smoothed shares. This answers how different the node's mix is on this attribute.
- **Value-level signed gap** `δ_{k,v} = p_W(v) − p_S(v)`, with Beta credibility. These are the attribute-level insights (for example "boucle 11% vs 0%", "white/cream 24% vs 6%").
- The **Colour Gap** of the original formula becomes `CG(a)` = JSD on colour_tone within the archetype. Colourway breadth is reported separately (next).

#### 6.7.4 Colourway breadth (the "flavours" of existing items)
- Staples: `n_colourways` per family (from S1).
- Wayfair: observed colours per family, which is a lower bound until the option list is crawled.
- Metric: for colour families present at Wayfair but absent from Staples in the same archetype, list the missing colour families. This feeds STYLE-EXTENSION recommendations directly.

#### 6.7.5 Design-Forward Gap (DFG)
- `DFG(a) = median DFI_W(a) − median DFI_S(a)`, plus the node-level **design-forward share gap** (share of DFI ≥ 0.6).
- The headline chart: DFI densities for Staples and Wayfair per node.

#### 6.7.6 Total Gap (TG): archetype-level composite
```
TG(a) = 0.35·pct(LSR) + 0.20·pct(PPG_mag) + 0.15·pct(CG) + 0.15·pct(MSG) + 0.15·pct(DFG)
```
- `MSG` = mean JSD over material_class and style_family within the archetype (this was missing from the original formula).
- `pct()` = percentile rank across **all archetypes in the same L2 peer group** (for example all Chairs & Seating archetypes), so nodes are comparable.
- Only credible LSR (Pr ≥ 0.9) contributes fully. Non-credible LSR is shrunk by ×0.5.
- **Sensitivity:** draw 1,000 weight vectors from Dirichlet(α = 20·w). Report each archetype's probability of staying in the node top-5 and Kendall τ against the base ranking. Also report the user's original 4-term weighting (0.35/0.25/0.20/0.20) side by side for comparison.

**Attribute-level Total Gap:** for attribute value *v* of attribute *k* in node *n*, `TG_attr(k,v) = pct(δ_{k,v}) × credibility`. This ranks "which values to add" independently of archetypes and answers the attribute-level question directly.

---

## 7. S9 Integration: two scores kept separate, one shared gate, fused re-rank

The two methods keep their own final scores, and each is reported and explained on its own:
- **Method 1 → VOS** (vector view: whitespace, fit, visible difference, safety), §6.5.6.
- **Method 2 → TG** (attribute view: share, price, colour, material/style and design-forward gaps), §6.7.6.

They are not blended into one formula. Insights are written from both. The final recommendation list is a re-ranking that combines the two rankings.

**Step 1: shared safety gate.** An archetype enters the final ranking only if it is **shortlist-eligible** (§6.6 roll-up): Safe Share ≥ 0.5 and UNDERCUT + SUBSTITUTE < 0.3. The gate applies to both methods, because TG measures gaps only: without the gate, a large TG gap could be a cheaper copy of a Staples product. Archetypes that fail the gate keep both scores in every table, with the reason they were excluded.

**Step 2: fused re-rank.** Within each node, over the eligible archetypes:
```
Final(a) = w₁·pctrank_node(VOS(a)) + w₂·pctrank_node(TG(a))     (w₁ = w₂ = 0.5, configurable)
```
- Ranks are fused rather than raw scores, so neither method's scale dominates. Reciprocal-rank fusion (`Σ 1/(60 + rank)`) is computed as a robustness check. If the two orderings disagree on the node's top 3, it is flagged.
- **Agreement tier** (shown next to every recommendation):
  - **Strong:** top tercile in both methods.
  - **Gap-led:** top tercile in TG only. The gap is real, but the vector view sees it as crowded, off-brand or close to Staples items. Read it together with its label mix.
  - **Vector-led:** top tercile in VOS only. There is whitespace in embedding space that the attribute grid does not capture, which usually means a facet is missing from the archetype grid. It is fed back to S6.
  - **Weak:** otherwise.
- **Method agreement per node:** Spearman ρ between the VOS and TG rankings, reported in the node header.
- **Shortlist:** per node, the top N by Final (N = 3–5) with tier Strong, Gap-led or Vector-led, plus a separate **STYLE-EXTENSION list** (colour or material variants of existing Staples families, from the SKU labels).
- **Demand:** D is off until review counts or rank position exist. When it is on, it enters as a third ranking in the fusion (w₃ = 0.3, with the others rescaled), not inside either method's score.
- **Headline chart per node:** TG (x) vs VOS (y), one point per archetype, coloured by agreement tier. Gated-out archetypes are shown hollow.

---

## 8. S10 SKU stage: exemplars and sellers

### 8.1 Exemplar selection (3–5 per shortlisted archetype)
- Candidates are Wayfair families in the archetype labelled CURATE, STYLE-EXTENSION or TRADE-UP.
- Rank by `0.5·(1 − SKU-CRS) + 0.3·centroid proximity + 0.2·DFI` and pick with MMR diversity (λ = 0.7), so the exemplars are not five near-identical chairs.

### 8.2 Evidence card per exemplar
Title, price, URL, display brand, key attributes, DFI, label, and **the nearest Staples family side by side** (title, price, URL) with CRS, PPR and AD. This is the same format as Pat's calibration pairs, so the recommendations read the way Pat already judged them.

### 8.3 Brand and seller information (replaces brand fragmentation)
- `brand_on_staples`: the Wayfair display brand already sells on Staples.com (for example Flash Furniture, Boss). This is a **quick-win recruit**, because the supplier relationship already exists.
- `wayfair_house_brand`: flag from a maintained list (Latitude Run, Ebern Designs, Inbox Zero, George Oliver, Red Barrel Studio, 17 Stories, Wrought Studio, Mercer41, Corrigan Studio, Hokku Designs, Wade Logan, Orren Ellis, Trule, Winston Porter, Mercury Row, Zipcode Design, Andover Mills, Etta Avenue, …). The manufacturer is not observable, so the recruit path is **source the archetype, not the SKU**.
- **Brand fragmentation** (HHI of display brands in the archetype) is reported descriptively and caveated.

---

## 9. S11 Insights and outputs

### 9.1 Deliverables (current phase)
1. **Python code**: the pipeline stages S0–S11, config-driven (§1.4).
2. **PNG figures**: one set per analysis node plus the overview figures, in `outputs/figures/`.
3. **One self-contained static HTML report**: `outputs/report/Staples_Assortment_Report.html`.

No Excel workbook or deck for now. Intermediate tables (CSV or parquet) are pipeline artefacts for debugging and re-runs, not deliverables.

### 9.2 HTML report: format
- **A single file, fully static and shareable.** It works offline and opened from email or a shared drive, with no server, no CDN and no external fonts or libraries. Figures are the same PNGs embedded as base64. Tables are plain HTML. A small inline vanilla-JS script handles tabs, the node dropdown and table sorting. Target size under 25 MB; if it grows past that, PNG resolution is reduced.
- The page is generated from the pipeline outputs by a template (Jinja2), so a re-run on new data rebuilds it with no manual editing.
- Must work in light and dark mode and at laptop and phone widths. Must print cleanly (each node section prints on its own).

**Tab 1: Approach & Methodology.** How we approached it, readable by a non-specialist first, with detail below:
- The business question and Pat's mandate (one paragraph), and what "core-adjacent" and "cannibalization" mean here.
- A flow diagram of S0 → S11, with one line per stage on what it does and why.
- Data coverage: retailers, sample sizes, in-scope nodes, Staples-only nodes and the backlog, plus data caveats (convenience sample, no demand data).
- The two methods side by side: the components of VOS and of TG, the shared gate, the decision labels with plain-English meanings, and how the final re-rank works.
- QA gate results (§10), with pass/fail and the actual values.
- Assumptions, and anything flagged *provisional*.

**Shelf filter (report only, added 2026-09-28).** For the two L2s that are already dense (Chairs & Seating, Desks), the dropdown is capped at 5 and 3 shelves. R1 enough data (scored, ≥ 30 families per side) and R2 at least one recommended archetype are hard gates; R3 high-confidence mapping ≥ 85%, R4 Spearman VOS↔TG ≥ 0.30 (on the displayed 2-dp value) and R5 at least one Strong-tier recommendation are counted. Rank = R3–R5 passed, then number of recommendations, then Spearman. L2s without a cap (every future category) list all their shelves. All shelves are still scored, and the rule results are written to `shelf_selection.csv`. The caps and thresholds live in `config/pipeline.yaml` → `report.shelf_filter`.

**Tab 2: Gaps & Recommendations** (renamed from Node-Level Analysis). A **single-select dropdown** at the top lists every analysis node by full path (for example `Furniture → Chairs & Seating → Gaming Chairs`, or `Furniture → Chairs & Seating → Office Chairs → Task` for a pseudo-L5 node). Thin nodes are listed but marked. Choosing a node shows only that node's content:

| Section | Content |
|---|---|
| **Node header** | Path; Staples and Wayfair family counts; mapping confidence; VOS↔TG agreement (ρ); provisional or thin flags |
| **Key insights** | 3–6 plain-English findings generated from the numbers (for example "Wayfair's mix is 24% white/cream vs 6% at Staples; credible"), each tied to the table or chart that supports it |
| **Coverage charts** | Price-band coverage (Staples vs Wayfair shares), DFI density, attribute-gap heatmap (JSD per attribute) |
| **Attribute-level gaps** | Top value-level gaps per attribute with credibility (§6.7.3), plus colourway breadth (§6.7.4) |
| **Method 1: vector view** | Archetype table with VW, AAS, CRS, AD, PPR band, label mix and **VOS**; decision scatter (AAS vs CRS, coloured by label) |
| **Method 2: attribute view** | Archetype table with LSR (credibility), PPG, CG, MSG, DFG and **TG**; weight-sensitivity summary |
| **Final recommendations** | TG vs VOS chart; the fused ranking with agreement tier and gate status; shortlisted archetypes with name, definition and the reason they rank where they do |
| **SKU recommendations** | Evidence cards (§8.2): competitor product next to its nearest Staples product, with price, link, label, CRS, PPR and AD. Plus the STYLE-EXTENSION list |
| **Vendor / seller view** | Display brands in recommended archetypes: `brand_on_staples` (quick-win recruit), `wayfair_house_brand` (source the archetype, not the SKU), brand fragmentation |
| **Excluded and caveats** | Gated-out archetypes with reasons; backlog items that came from this node's pages; data caveats specific to the node |

### 9.3 Figures (PNG, per node unless noted)
Price-band coverage · DFI density · attribute-gap heatmap · top value-level gaps (diverging bar) · decision scatter (AAS vs CRS) · TG vs VOS agreement chart · weight-sensitivity (top-5 stability) · overview: node coverage and label mix across all nodes · QA: mapping confusion and extractor accuracy per field.

---

## 10. Validation and QA gates (the pipeline does not advance past a failed gate)

| Gate | Metric | Target |
|---|---|---|
| G1 Mapping | top-1 accuracy on the 200-family gold set; NONE precision | ≥ 90% / ≥ 85% |
| G2 Extraction (T1/T2) | per-field accuracy against Staples specs (text-only extractor) | ≥ 85% per field used in gaps |
| G3 Extraction (T3, DFI) | κ between raters; LLM vs humans | κ ≥ 0.6; ≥ 0.8 × human agreement |
| G4 Families and dedup | manual audit of 100 merges and 100 cross-retailer pairs | precision ≥ 95% |
| G5 Embedding | source-mixing: identical pair in the top-3 | ≥ 80% |
| G6 Archetypes | ARI grid vs HDBSCAN; bootstrap stability; nameability | ≥ 0.4 / ≥ 0.6 / 5 of 5 |
| G7 Scores | TG and VOS top-5 stability under Dirichlet weights; PPR and CRS calibration AUC | ≥ 70% / AUC ≥ 0.8 |
| G8 Sanity | DATA-ERROR share; spot-check of 20 recommendations by a merchant-literate reviewer | ≤ 2% / ≥ 16 of 20 judged sensible |

---

## 11. Assumptions, open questions and data requests

**Assumptions** (reviewed by Sai on 2026-09-27)
1. ✅ *Confirmed.* Every item on Staples.com counts as "1P / existing assortment" for cannibalization, including third-party brands Staples retails. We have no marketplace flag.
2. ⚠️ *Revised.* The competitor sample is a **convenience sample**: every product page that scraped successfully, with failures dropped. It is not designed to be representative, and the failures may not be random (for example, some page templates failing more often). Consequences:
   - All claims are **share-based and within-sample**. No absolute-depth claims.
   - Every node reports its sample size, and every gap its credibility, so a thin sample shows up as "not credible" rather than as a false gap.
   - Where the scraper logs it, scrape success rate per listing page is kept as a data-quality column.
   - Samples will grow. The pipeline is re-run end to end on the new data, and nothing is tuned to today's counts.
3. ✅ *Confirmed.* Wayfair consumer and Professional pages are pooled as one competitor. `scope` is kept as a column so they can be split if needed.
4. ✅ *Confirmed.* Price is the list price at crawl time. Promotions are ignored.

**Open questions: status**
1. *Answered.* The Wayfair sample is a random or convenience set for now. The goal is to get the framework, code, outputs and recommendations working first, then correct from the top as data is enriched. No rank position or page totals, so the demand term stays off.
2. **Open (keep).** Can Ayan re-crawl for: Wayfair review count and rating, option or colour lists, primary image URL; the Staples review count (currently empty); and Staples Big & Tall, Hutches and Table Lamps, which are under-covered?
3. *Answered.* Chairs & Seating and Desks for now, but scope must adapt to whatever categories are shared later. See §1.3–1.4.
4. *Answered.* Run HuggingFace encoders locally. Checked 2026-09-27: huggingface.co is reachable (HTTP 200). `torch` and `sentence-transformers` are not installed yet and there is no GPU (12 CPU cores). CPU is enough at this scale: about 15k short cards with a base-size model takes minutes. The large models are tried in the bake-off only if run time allows. The Claude API is assumed available for extraction and adjudication.
5. **Open.** Is Pat's calibration session realistic in the timeline? If not, the weakly supervised thresholds (§6.6) become final, with a caveat.

**Data requests, by priority** (all optional: the pipeline runs without them)
- **P1:** Wayfair review count and rating, and image URL (both sides).
- **P2:** Wayfair option lists; Staples review counts; listing rank position; "customers also viewed" (which would enable A_ctx through a graph); scrape success rate per page.

---

## 12. Build plan (after sign-off)

| Phase | Deliverable | Gate |
|---|---|---|
| P0 | Ingest, clean, family grouping, dedup, analysis nodes; **coverage report** (which nodes are in scope for the data supplied) | G4 |
| P1 | Page crosswalk, product-level mapping, gold set | G1 |
| P2 | Schema induction and freeze; extraction; extractor validation; DFI | G2, G3 |
| P3 | Embedding bake-off, canonical cards, vector store | G5 |
| P4 | Archetypes for the vertical slice (*Accent & Waiting Room Chairs* and *Office Chairs*) | G6 |
| P5 | Method 1 and Method 2 on the slice; decision tree; weak calibration | G7 |
| P6 | Integration, insights, SKU cards, figures and the **HTML report** for the slice → **review with Sai** | G8 |
| P7 | Roll out to every in-scope node the coverage report finds (today: the rest of Chairs & Seating, then Desks); Pat calibration; final HTML report | all |

Each phase runs on the vertical slice before any bulk LLM spend. The bulk extraction run happens only after the P2 prompts pass validation on the slice.

**Re-run on new data:** when enriched or new-category samples arrive, the same pipeline runs from S0. New L2s get their generated artefacts (§1.4) drafted automatically and flagged *provisional* until reviewed. Existing reviewed artefacts are reused.

---

## 13. Implementation notes (PoC build, 2026-09-28)

These decisions were made while building the pipeline. Each one either tightens a rule above or fills a gap that only showed up in the data. The code in `src/alpoc/` follows them.

| # | Where | Decision | Why (evidence from the run) |
|---|---|---|---|
| 1 | S5 | Extraction backend is **local rules + embedding zero-shot**. The Claude backend is not wired yet | No API key was available. All Tier-3 and DFI outputs are provisional (G3 pending) |
| 2 | S3 | The **NONE threshold is calibrated across sources**: Youden-J between known in-scope products (single-candidate pages) and known out-of-scope products (pages with no Staples shelf) | Calibrating on Staples-vs-Staples similarity rejected 5,006 in-scope Wayfair products, because the two sources' text is systematically less similar |
| 3 | S2 | Pseudo-L5 split keys exclude **component facets** (Arm, Base, Back, Seat…) | The first run split Office Chairs on "Arm Type". It now splits on "Chair Type" |
| 4 | S6–S7, DFI | All vector math (VW, AAS, CRS, AD, archetype validation, DFI) runs on a **source-neutral card** built only from *text-instrument* extracted fields. There is no free title text, and the Staples spec values are not used in cards | Title-bearing cards let a classifier tell the retailer apart with AUC 0.97, and the median Wayfair product fell at the 10th percentile of Staples' affinity. Spec-filled Staples cards added a missingness signature. Specs remain the truth for physical facts in Method 2 |
| 5 | S7 | **AAS semantic affinity uses the functional view** (type, features, size), not the full card | On the full card, the aesthetic difference that is the White Chair opportunity was scored as "off-brand". The look is rewarded through AD only |
| 6 | S7 | **Functional peers** = every Staples family within 0.02 of the best functional match (up to 15). PPR uses their median price; AD is taken against the best-looking peer (conservative) | Neutral cards create many ties, and an arbitrary top-3 made PPR and AD unstable |
| 7 | S5 | Features that text can only confirm (swivel, lumbar, headrest…) are measured as **"mentioned: yes/no" on both sides** | Mixing a spec "no" with text-only "yes" would create false gaps |
| 8 | S5 | Colour is read from the **last** title segment first | Titles put the colour variant last ("…, Walnut Trim, Black") |
| 9 | QA | **G1** = Staples leave-one-out leaf accuracy plus NONE balanced accuracy (proxy until the human gold set exists). **G5** = retailer predictability from the card (lower is better), reported against the title-bearing card | No human gold set, and too few same-brand pairs (2) for the original G5 |
| 10 | S11 | **Report shelf filter**: Chairs & Seating is capped at 5 shelves and Desks at 3 (R1 data and R2 recommendations are gates; R3 confidence, R4 agreement and R5 Strong backing are counted) | 23 shelves made the report long (11.5 MB) and would crowd out new categories. Filtering drops the size to 5 MB, and no scoring changes (§9.2) |

**What the first full run shows (for review, not for sign-off):**
- G1 **fails** on the NONE decision. Out-of-scope products such as patio chairs and bean bags are semantically close to seating, so similarity alone rejects only about 29% of them. The page crosswalk catches whole pages; mixed pages need LLM adjudication.
- G6 **fails**. The embedding clusters split on product sub-form (barrel, wingback, club), which the grid does not use. Candidate facet for the next iteration: `form_factor` with finer accent-chair values.
- G8 **fails**. The DATA-ERROR share is about 12%. These products sit close to a few Staples items but far from the wider Staples catalogue or its use contexts. They are to be reviewed in Pat's calibration.
- Recommendations skew to **premium price bands**, because cheaper look-alikes are labelled UNDERCUT (PPR < 0.85). This protects 1P as intended, but the band is the first thing Pat's calibration should test.
