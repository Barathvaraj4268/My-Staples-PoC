# Staples Marketplace PoC – Category-Level Recommendations (v2, six retailers)

This pipeline compares the Staples.com navigation tree with those of Office Depot, West Elm, Wayfair, Amazon and Walmart. From the comparison it recommends which categories Staples should open to marketplace sellers, and places each one in one of six zones: CURATE, VERTICAL EXTENSION, REVIEW, 1P-CORE GAP, OFF-BRAND and VERIFY.

## Files

| File | What it does |
|---|---|
| `poc_common.py` | Shared library: `CONFIG`, retailer adapters (`RETAILERS`), department scope, 1P coreness, loading, count handling, calibration, labels and opportunity, stress test |
| `method_v_vector.py` | **Method V**: encoder bake-off, level-agnostic path embeddings, Qdrant vector DB (one collection per retailer), retrieve-then-rerank, per-competitor calibration. Writes `outputs/V/…` |
| `method_g_graph.py` | **Method G**: six-retailer knowledge graph, similarity flooding, Personalised PageRank adjacency, Neo4j and GraphML export. Writes `outputs/G/…` |
| `run_framework.py` | Combines the two methods, applies labels (including the verification loop), builds ENTER concepts and DEEPEN aisles, scores AAS, CRS, PC, GS and O, assigns the six zones, runs the stress test and rolls results up to Staples aisles |
| `framework_outputs.py` | Writes every figure (`fig01`–`fig20`, plus deck versions), `Category_Recommendations_v2.xlsx` and `summary.json` |
| `label_sample.py` | Draws the 120-row labelling sample for a new competitor |
| `run_all_colab.ipynb` | Colab notebook that runs everything end to end |
| `gold_matches_v2.csv` | 914 labelled rows: 695 for calibration (Office Depot 215; the other four competitors 120 each) and 219 verification-loop checks |

## Run

```bash
pip install pandas numpy scipy scikit-learn openpyxl matplotlib networkx qdrant-client spacy adjustText
python -m spacy download en_core_web_lg
# put the six navigation-tree .xlsx files and gold_matches_v2.csv in ../data (or set POC_DATA_DIR)
python method_v_vector.py      # ~2-4 min (includes the encoder bake-off and weight grid)
python method_g_graph.py       # ~30 s
python run_framework.py        # ~1 min
```

Outputs are written to `../outputs` (or `POC_OUT_DIR`):

- `final/Category_Recommendations_v2.xlsx` – every table in the documents
- `final/figures/fig01…fig20.png` – every chart in the documents and the deck
- `final/summary.json` – every number quoted in the documents and the deck
- `G/neo4j/*` – nodes, edges, and a load script plus example queries in Cypher
- `V/qdrant_db` – the persisted vector DB

A clean re-run reproduces every zone count exactly; continuous scores agree to about 0.1 point, because of floating-point summation order.

## Encoders

`CONFIG["encoder"]` is set to `"spacy"`, which is the encoder that produced the documents. Setting it to `"auto"` makes the code bake off every encoder that loads and keep the most accurate one on the labelled rows: sentence-transformers `BAAI/bge-small-en-v1.5`, spaCy, WordLlama and TF-IDF. In Colab this includes bge-small. The sandbox used for this run cannot reach HuggingFace. Thresholds are re-calibrated automatically for whichever encoder wins.

## Adding a competitor

1. Add an adapter to `RETAILERS` in `poc_common.py`: file pattern, level columns, count column and `count_mode` (`terminal_only`, `cumulative`, `partial` or `none`), role, and scope rules.
2. Run `method_v_vector.py`. The new competitor starts on pooled calibration, flagged as provisional.
3. Run `python label_sample.py <Name>` and label the 120 rows. Append them to `gold_matches_v2.csv`, then re-run all three scripts.
