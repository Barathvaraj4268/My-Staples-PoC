# Staples Assortment Gap PoC (AL PoC, V1)

For each granular Staples shelf, this pipeline finds the product archetypes and attribute variants a competitor
(Wayfair today) carries that Staples lacks, and checks whether Staples could add them through its marketplace
without cannibalising its own assortment. The method is in [methodology.md](methodology.md).

## Run

```bash
pip install pandas numpy scipy scikit-learn openpyxl pyarrow pyyaml ftfy jinja2 matplotlib umap-learn
pip install torch --index-url https://download.pytorch.org/whl/cpu
pip install sentence-transformers

python run_pipeline.py              # everything (first run downloads the encoders and embeds every card)
python run_pipeline.py --from s6    # re-run from archetypes onward
python run_pipeline.py --only s11   # rebuild the report only
```

Embeddings are cached in `data/cache/`, so re-runs only encode new text.

## Outputs

| Path | What |
|---|---|
| `outputs/report/Staples_Assortment_Report.html` | The shareable report: one static file that works offline. Tab 1 covers approach and methodology; Tab 2 has a node dropdown |
| `outputs/figures/<node>/*.png`, `outputs/figures/overview/*.png` | Every chart in the report |
| `outputs/tables/*.csv` | Final archetypes, candidates, attribute gaps, SKU recommendations, vendors, backlog, audit samples |
| `data/interim/` | Stage outputs (parquet) and QA files (`qa_*.json`) |

## Pipeline

| Stage | Module | Output |
|---|---|---|
| S0 Ingest & clean | `s0_ingest.py` | `sku_<retailer>.parquet` (reviews/PII stripped) |
| S1 Families & dedup | `s1_families.py` | `families.parquet` |
| S2 Analysis nodes | `s2_nodes.py` | `nodes.parquet` (leaf or pseudo-L5; scored / thin / staples_only) |
| S3 Mapping | `s3_mapping.py` | `mapping.parquet`, encoder bake-off |
| S4–S5 Attributes | `s5_extract.py` | `attributes.parquet`, validation vs Staples specs |
| S6 Archetypes | `s6_archetypes.py` | `archetypes.parquet`, grid vs HDBSCAN check |
| S7 Method 1 (VOS) | `s7_vector.py` | `candidates.parquet` (labels), `archetype_m1.parquet` |
| S8 Method 2 (TG) | `s8_gaps.py` | `archetype_m2.parquet`, `attr_gaps.parquet` |
| S9 Integration | `s9_integrate.py` | `final_archetypes.parquet` (gate, fused rank, tiers, sensitivity) |
| S10 SKUs & sellers | `s10_skus.py` | `sku_recs.parquet`, `style_extensions.parquet`, `vendor_view.parquet` |
| S11 Report | `s11_report.py`, `figures.py`, `templates/` | HTML, PNG, CSV |

## Adding data

- **More products or new categories:** replace the Excel files and re-run. Scope is decided by the data: any Staples
  shelf with enough families on both sides is scored. A new L2 runs on universal attributes and is flagged
  *provisional* until someone adds `config/l2/<l1>__<l2>.yaml` (Tier-2 attributes, grid facets, DFI anchors).
- **A new competitor:** add an adapter to `config/retailers.yaml`, a page crosswalk in `config/crosswalk/`, and its
  house brands to `config/house_brands.yaml`.
- **All thresholds and weights** live in `config/pipeline.yaml`.

## Current limitations

- There is no Anthropic API key, so extraction uses the local `rules` backend (vocabularies plus embedding zero-shot).
  The Claude backend described in the methodology is not wired yet.
- The human gold sets (mapping, Tier-3 tags and DFI, audits) and Pat's calibration session are pending, so those
  gates show PENDING.
