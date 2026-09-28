# CLAUDE.md: Staples Assortment (AL) PoC, V1

Guidance for Claude working in this folder. Read `methodology.md` before any analytical or code work. It is the source of truth for method decisions. If this file and `methodology.md` disagree, `methodology.md` wins; flag the conflict.

## Project in one paragraph
LatentView PoC for **Staples Marketplace** (stakeholder: Pat, Head of Staples Marketplace; Q1 2027 planning). Owner: Sai. Team: Teja, Barath, Ayan. The question: within Staples' most granular navigation nodes (L3/L4, plus pseudo-L5 splits), which product archetypes and attribute variants (colour, material, style, price tier, size, use-context, "vibe") does a competitor (Wayfair for now) carry that Staples lacks, and which of them can Staples add through the marketplace **without cannibalizing its existing assortment**. The thesis is "White Chair" / core-adjacent: design-forward and lifestyle extensions of core items. **Constraint: external data only.** There is no internal Staples sales, margin or traffic data.

## Folder map
| Path | What | Rules |
|---|---|---|
| `Documents/*.docx` | Objective & Goal (Pat's mandate), Initial-Exploration methodology (Track B/C), Examples (Scheels/Michaels/Best Buy depth analogies) | Read-only |
| `Excels/1. Staples_Navigation_Tree_repaired.xlsx` | Staples tree (1,877 nodes, L1–L4, **no L5**), leaf counts, Cross-Listings sheet | Read-only |
| `Excels/3. Wayfair_Navigation_Tree_v3.xlsx` | Wayfair tree: Consumer + Professional (B2B) + Professional Shop-by-Space; no item counts | Read-only |
| `Excels/Staples product level dataset latest - Samples.xlsx` | 31,979 rows / 26,861 SKUs across many L2s (18 MB) | Read-only. Slow to load: cache to parquet/pickle under `data/interim/` |
| `Excels/Wayfair product level dataset latest - Samples.xlsx` | 12,255 rows / 11,468 products across 32 Wayfair listing pages (seating, desks, mats) | Read-only |
| `Excels/Vlookup - Manual Trail … [Failed - DO NOT USE].xlsx` | Failed name-based page→leaf mapping | **Do not use as input.** It is evidence only |
| `Excels/~$*.xlsx` | Excel lock files | Ignore |
| `methodology.md` | Finalised methodology (draft v1 under Sai's review) | Edit only when asked |
| `../../1. CL PoC/V2 - with all 5 Competitors/staples_category_poc_v2_code/code/` | Earlier category-level PoC: encoder bake-off, Qdrant, calibration, figure style | Reuse patterns; do not modify |

Built (2026-09-28): `run_pipeline.py` (`--from sX` / `--only sX`), `src/alpoc/` (one module per stage + `figures.py`, `cards.py`, `vocab.py`, `templates/`), `config/` (`pipeline.yaml` thresholds/weights, `retailers.yaml`, `schema_universal.yaml`, `l2/*.yaml`, `crosswalk/wayfair.yaml`, `house_brands.yaml`), `data/interim/` (stage parquet + `qa_*.json`, git-ignored), `data/cache/` (embeddings, git-ignored), `outputs/` (report, figures, tables). See README.md.

## Data quirks (verified 2026-09-27; don't rediscover)
- **Staples `description` is a serialized Python dict** `{paragraph, bullets, specification:[{name,value,grpName,dscr}]}`. Parse with `ast.literal_eval` (100% parse). Specs are rich: colour 97%, material 94%, W/D/H 91%, style 47%.
- Staples `reviews` = `"4.57 stars ( reviews)"`. Rating only; **the review count is always empty**.
- Staples lists **colour variants as separate SKUs**: Chairs+Desks is 4,725 SKUs but only about 2,860 families. Count **families**, not rows.
- Staples has 4,781 exact-duplicate rows. 141 SKUs sit in more than one leaf (use the canonical nav-tree path).
- Staples private label is only about 1.4% in furniture. "1P" means *items Staples sells* (Flash Furniture, Boss, Bush, HON…).
- Wayfair `category` = listing page, with "&" mangled to 3 spaces. Pages include rooms and spaces (*Meeting Space Seating*, *Hospitality Seating*). 156 products sit on more than one page, and some products are mis-shelved. **Map at product level, never by page name.**
- Wayfair `description`: 41% have a customer review appended after `" | "`, with **reviewer name, city and date (PII)**. Split on the first `" | "`, keep the left side, drop the rest, and never output review text.
- Wayfair `vendor` = display brand, mostly **Wayfair house brands** (Latitude Run, Ebern Designs, Inbox Zero, …), with mojibake (`Â®`). It is not the manufacturer or seller.
- Wayfair has **no specs, rating, review count, images or option lists**. `selected_choice` is a single displayed variant (57% filled). 12 prices are `"per item"`.
- Wayfair is a **convenience sample** (every page that scraped successfully; failures dropped) with no site totals, while Staples is close to a census (85–100% of site counts). **Compare shares, never raw counts.** Samples will be enriched later, so nothing may be tuned to today's counts.
- *Today's* Wayfair–Staples overlap is Furniture → **Chairs & Seating** and **Desks**. Door mats, bean bags, patio dining, theatre seating, restaurant sets and chiavari go to the new-node backlog. This is a property of the current data, not of the method (see guardrail 0).

## Method guardrails (non-negotiable unless Sai changes them)
- **Build-time decisions are in `methodology.md` §13** (cross-source NONE threshold, source-neutral cards, functional AAS, peer sets). Read it before changing S3/S5/S7.
0. **Scope-agnostic.** Never hard-code categories (chairs, desks, …), attribute lists, vocabularies, split facets, DFI anchors or house-brand lists in code. In-scope nodes are found from the data on each run. Category-specific artefacts are generated per L2 (LLM draft + human review), stored as versioned config, and flagged *provisional* until reviewed (`methodology.md` §1.3–1.4). New categories or competitors must run with no code change.
1. Unit of analysis = **product family**. Analysis node = Staples leaf, split into pseudo-L5 by Staples' own type facet when large (§5.3).
2. **Product-level mapping** (Wayfair family → Staples node | NONE) through page prior + classifier trained on Staples SKUs + LLM adjudication. Gold-set accuracy ≥ 90% before moving on.
3. **Same instrument on both sides** for any attribute that feeds a gap metric. Staples specs are ground truth for validating the text extractor.
4. Embed a **canonical product card** (same template on both sides, no price, no raw spec dump), never raw descriptions.
5. **CRS is price-agnostic and aesthetic-agnostic.** Price goes through PPR only; aesthetics through AD/ΔDFI only. Otherwise TRADE-UP and STYLE-EXTENSION become unreachable.
6. Gaps are **share-based** (Beta-smoothed log share ratio + credibility), defined even when Staples = 0. No `log1p(count)` depth gaps.
7. Decision labels come from the **exhaustive ordered tree** in §6.6: BACKLOG, EXCLUDE, OFF-BRAND, UNDERCUT, TRADE-UP, STYLE-EXTENSION, SUBSTITUTE, REVIEW, CURATE, EDGE.
8. **No raw cosine values in any deliverable.** Thresholds come from calibration (weak supervision from Staples-vs-Staples pairs, then Pat). Re-fit when the encoder changes. PPR bands are per L2.
9. **Two method scores stay separate:** Method 1 → **VOS** (VW + AAS + AD, damped by CRS) and Method 2 → **TG**. They are never blended into one formula. Final recommendations = **shared safety gate** (label mix), then a **fused re-rank** of the two percentile ranks, with an agreement tier (Strong / Gap-led / Vector-led / Weak). PPR is never inside a score. Weights always ship with a Dirichlet sensitivity analysis (top-5 stability).
10. Every recommendation shows the **nearest Staples family side by side** with CRS/PPR/AD.
11. Never claim demand. The data measures assortment supply. The D term stays off until review counts or rank exist.
12. Wayfair house brands ≠ recruitable sellers. Flag `brand_on_staples` (existing supplier = quick win) and `wayfair_house_brand`.

## Environment
- Windows 11. Shells: PowerShell (primary) and Git Bash. Use forward slashes in Bash and quote paths: the folders contain spaces and dots.
- Python 3.14, **pandas 3.0** (default string dtype: `astype(str)` keeps NaN as missing, so use `fillna('')` before joining strings), scikit-learn 1.9, openpyxl, `anthropic` SDK 0.109.
- **Installed 2026-09-28:** torch (CPU), sentence-transformers, umap-learn, matplotlib, jinja2, ftfy, pyarrow. HDBSCAN comes from scikit-learn. **Not installed:** `python-docx` (read .docx by unzipping `word/document.xml`), qdrant-client (numpy is enough at this scale). Ask before installing more.
- Printing unicode to the console: set `PYTHONIOENCODING=utf-8` (cp1252 console otherwise crashes).
- HuggingFace is reachable from this machine (checked 2026-09-27) and Sai approved local HF encoders. `torch` and `sentence-transformers` are still to be installed (at P3). No GPU (12 CPU cores), so prefer base-size encoders and try large ones only in the bake-off.

## LLM usage
- Bulk extraction: `claude-haiku-4-5-20251001` through the Batch API with JSON-schema/tool-use output and prompt caching for the frozen schema.
- Schema induction, adjudication, DFI rubric and archetype naming: `claude-sonnet-5`. Reserve `claude-opus-5-5` for hard adjudication or review only.
- Cache every raw LLM response to disk (keyed by prompt version + product id) so re-runs are free. Version the prompts.
- **Ask Sai before any bulk run** (more than about 500 items or roughly $20+). Pilot on the vertical slice first.

## Working conventions
- Build as a **vertical slice first**: *Accent & Waiting Room Chairs* (+ *Office Chairs* as contrast), then all of Chairs & Seating, then Desks.
- Each stage writes a versioned intermediate table. Stages must be re-runnable independently. Keep thresholds and weights in `config.yaml`, not in code.
- Respect the QA gates in `methodology.md` §10. Report a failed gate plainly rather than working around it.
- **Deliverables (current phase): Python code, PNG figures, and ONE self-contained static HTML report.** No Excel or deck unless Sai asks. The report is a single file with no CDN, fonts or server: PNGs embedded as base64, inline vanilla JS, built from a Jinja2 template, under 25 MB. Tab 1 = Approach & Methodology. Tab 2 = Node-Level Analysis with a single-select dropdown of full node paths, where choosing a node shows all its tables, charts, insights, scores, SKU and vendor recommendations (`methodology.md` §9.2).
- Deliverable style: clear tables, charts in the CL PoC V2 figure style, merchant-readable language (archetype names a merchant can say out loud).
- Git: commit only when Sai asks. Never commit the large Excels again or any file containing review text or PII.

## Status
- 2026-09-27: data audited; `methodology.md` v1 → v1.2 (Sai's §11 answers; separate VOS/TG + fused re-rank; outputs = code + PNG + static 2-tab HTML).
- 2026-09-28: **PoC built and run end to end** (about 15 min cold; embeddings cached). 18 scored shelves + 5 thin; 22 archetypes shortlisted; 88 exemplar SKUs; report `outputs/report/Staples_Assortment_Report.html` (8.3 MB). Gates: G7 pass; G2 partial; G1/G5/G6/G8 fail (explained in methodology §13); G3/G4 pending human work.
- Next candidates: wire the Claude extraction/adjudication backend (needs `ANTHROPIC_API_KEY`), human gold sets (G1/G3/G4), a finer accent-chair `form_factor` facet (G6), Pat's calibration (PPR bands, CRS/AAS thresholds).
- Environment now has torch (CPU), sentence-transformers, umap-learn, matplotlib, jinja2, ftfy, pyarrow installed.
