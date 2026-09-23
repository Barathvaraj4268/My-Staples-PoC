#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
=====================================================================================
 STAPLES MARKETPLACE PoC  -  CATEGORY-LEVEL RECOMMENDATION  (Track A + C, revised)
 METHOD V : VECTOR-FIRST   (embeddings + vector index, retrieve-then-rerank)
=====================================================================================

What this script does, in plain language
----------------------------------------
Think of every category page on every retailer's site as a "shelf label".

 1. LOAD + HARMONISE   Read each retailer's navigation tree into one standard table.
                       Fix item counts so they mean the same thing for every retailer
                       (Staples only counts leaf pages and cross-lists heavily; Office
                       Depot's parent pages already include their children).
 2. EMBED              Turn every shelf label into a vector (a list of numbers that
                       captures meaning) so "Sofas & Lounge Chairs" sits next to
                       "Sofas, Loveseats & Sectionals" even though the words differ.
 3. RETRIEVE + RERANK  For every competitor shelf, pull the 50 closest Staples shelves
                       from a vector index (Qdrant, or plain NumPy), then re-score them
                       with a hybrid of meaning (embedding) + wording (TF-IDF).
 4. CALIBRATE          Use a small hand-labelled "gold" set to decide how close is
                       "close enough to count as the same shelf". No hard-coded cosine.
 5. GAP                Roll matches up the competitor tree: where does the competitor
                       carry assortment that Staples does not (ENTER) or carries more
                       deeply relative to its catalogue size (DEEPEN)?
 6. TRACK C            For each gap ask two questions:
                         Adjacency (AAS)       - does it sit inside the neighbourhood
                                                 Staples already occupies?   (mean-pool)
                         Cannibalization (CRS) - is it basically the same thing as a
                                                 Staples CORE 1P category?     (max-pool)
 7. DECIDE + RANK      Label every gap (CURATE / VERTICAL EXTENSION / 1P-CORE GAP /
                       OFF-BRAND / REVIEW), rank the approved ones, and stress-test the
                       ranking with 500 random re-weightings.

How to run
----------
 Google Colab
   1) Upload this file + the input files to /content (or mount Drive and point
      CONFIG["data_dir"] at the folder).
   2) Uncomment and run the INSTALL lines in section [0], then:
          !python method_v_vector.py
      (or paste the file cell by cell - cells are marked with "# %%").
 Local (Python 3.9+)
      pip install pandas numpy scikit-learn openpyxl matplotlib qdrant-client spacy
      python -m spacy download en_core_web_lg
      python method_v_vector.py

Inputs  (anywhere under CONFIG["data_dir"]; names are matched by pattern)
   Staples_Navigation_Tree_v2.xlsx        sheets "Navigation Tree", "Cross-Listings"
   officedepot_tree.csv                   (or OfficeDepot_Navigation_Tree.xlsx)
   gold_matches.csv                       hand-labelled calibration set (shipped)
   external_signals.csv       [optional]  demand signal per competitor path (0-100)

Outputs (CONFIG["out_dir"])
   method_v_results_<Competitor>.xlsx   README, Recommendations, All_Gap_Units, Node_Matches,
                                        Review_Queue, Staples_Ahead, Calibration, Gold_Scored
   method_v_decision_surface_<Competitor>.png, method_v_top_opportunities_<Competitor>.png
   method_v_cross_competitor_gaps.xlsx  (from 2 competitors: gap concepts + peer coverage)
   node_scores_method_v.csv             (consumed by compare_methods.py)
"""

# %% [0] INSTALL  ---------------------------------------------------------------
# Uncomment in Colab (or run once locally):
# !pip install -q pandas numpy scikit-learn openpyxl matplotlib qdrant-client
#
# Pick ONE encoder (CONFIG["encoder"]):
#   "spacy" (default - reproduces the PoC numbers; works offline once downloaded, ~400 MB)
# !pip install -q spacy && python -m spacy download en_core_web_lg
#   "st"    (sentence-transformers - recommended upgrade; needs HuggingFace access)
# !pip install -q sentence-transformers
#   "tfidf" needs nothing extra (character n-grams + SVD) - weakest, last-resort fallback

# %% [1] IMPORTS + CONFIG  --------------------------------------------------------
import glob
import math
import os
import re
import sys
import warnings
from collections import defaultdict

import numpy as np
import pandas as pd

warnings.filterwarnings("ignore")
pd.set_option("display.width", 200)
pd.set_option("display.max_columns", 30)

CONFIG = {
    "data_dir": ".",                     # folder with the input files (e.g. "/content/drive/MyDrive/staples")
    "out_dir": "outputs_method_v",
    # --- encoder ---------------------------------------------------------------
    "encoder": "spacy",                  # "spacy" | "st" | "tfidf"  (falls back in that order)
    "spacy_model": "en_core_web_lg",
    "st_model": "BAAI/bge-small-en-v1.5",  # or "sentence-transformers/all-MiniLM-L6-v2"
    "name_weights": (0.50, 0.30, 0.20),  # embedding mix: own name, parent name, L1 name (grid-searched)
    # --- vector index ------------------------------------------------------------
    "vector_backend": "qdrant",          # "qdrant" (local, in-memory) | "numpy"
    "k_retrieve": 50,                    # stage-1 recall from the vector index
    "k_adjacency": 10,                   # neighbours averaged for adjacency (mean-pool)
    "w_sem": 0.65, "w_lex": 0.35,        # stage-2 hybrid rerank weights
    "multi_match_margin": 0.08,          # also accept other Staples shelves within this of the best
                                         # (compound shelves: "Greeting & Note Cards" = 2 Staples shelves)
    # --- gap logic ---------------------------------------------------------------
    "enter_coverage_max": 0.20,          # < 20% of competitor items matched -> ENTER (whitespace)
    "deepen_dg_min": 0.69,               # depth gap above ln(2): competitor >= 2x relative depth -> DEEPEN
    "min_unit_items": 10,                # ignore gap units smaller than this
    "max_unit_level": 3,                 # recommend at L1-L3 (category level); L4 leaves roll up
    # --- decision surface (0-100 scales) -----------------------------------------
    "aas_hi": 60, "aas_lo": 40,
    "crs_hi": 60, "crs_lo": 40,
    # --- ranking -------------------------------------------------------------------
    "rank_weights": {"dg": 0.40, "wm": 0.30, "aas": 0.30},   # + "demand" if a file is given
    "demand_weight": 0.25,
    "gamma": 1.0,                        # cannibalization penalty exponent
    "top_n": 20,
    # --- robustness ------------------------------------------------------------------
    "n_sensitivity": 500, "dirichlet_conc": 20, "threshold_jitter": 5.0,
    "review_queue_size": 100,
    "false_gap_cost": 1.0,               # weight on the carried class when picking tau (1 = plain Youden)
    "prior_source": "stratified sample", # gold rows drawn at random -> base rate of 'carried'
    "gap_screen_false_alarm": 0.10,      # gap screen: accept flagging 10% of shelves Staples DOES carry
    "seed": 42,
}

# ---- Retailer adapters: add the next competitors here (same keys) -------------------
# count_semantics: "terminal_only" -> only leaf pages show a count (Staples)
#                  "cumulative"    -> parent pages show a total incl. children (Office Depot)
# role: "focal" (Staples) | "mirror" (B2B peer) | "lifestyle" | "scale" | "specialist"
RETAILERS = [
    {
        "name": "Staples", "role": "focal",
        "file_glob": ["*Staples_Navigation_Tree*.xlsx"], "sheet": "Navigation Tree",
        "level_cols": ["L1", "L2", "L3", "L4", "L5"], "count_col": "Count",
        "count_semantics": "terminal_only", "id_col": "Category ID", "url_col": "URL",
        "crosslist": {"sheet": "Cross-Listings", "header_row": 3,
                      "name_col": "Category", "parents_col": "All parents",
                      "max_parents": 999},         # keep all placements for counting: hub shelves such
                                                   # as Disposable Cups sit under 41 parents
        "scope_exclude": [r"^Gift Cards", r"^Tech Services", r"^Warranties",
                          r"Furniture Assembly$"],
    },
    {
        "name": "OfficeDepot", "role": "mirror",
        "file_glob": ["*officedepot_tree*.csv", "*OfficeDepot_Navigation_Tree*.xlsx"],
        "sheet": "Category Tree",
        "level_cols": ["L1", "L2", "L3", "L4", "L5"], "count_col": "Count",
        "count_semantics": "cumulative", "id_col": "Node ID", "url_col": "URL",
        # Services / programme pages / custom-print storefront: Staples runs these outside
        # its merchandising tree (Ink finder excluded on both sides), so they would create
        # phantom gaps. Re-include "^Print & Copy" if you want custom print in scope.
        "scope_exclude": [r"^Ink & Toner", r"^GreenerOffice", r"^Services", r"^Security Solutions",
                          r"^Print & Copy", r"Warranties & Services", r"Protection Plans$"],
    },
]
FOCAL = "Staples"

# ---- 1P coreness: proxy for "how much Staples 1P revenue sits on this shelf" --------
# No sales data in the PoC. These are editable assumptions; replace with Staples'
# category sales / margin bands when available (a single email ask, per Track C 5.5).
CORE_L1_WEIGHT = {
    "Paper": 1.0, "Office Supplies": 1.0, "Shipping, Packing & Mailing Supplies": 1.0,
    "Cleaning Supplies": 0.9, "Coffee, Water & Snacks": 0.9, "Printers & Scanners": 0.9,
    "Shredders, Projectors & Office Machines": 0.8, "Computers & Accessories": 0.8,
    "Facilities": 0.7, "Safety Supplies": 0.7, "School Supplies": 0.7, "Batteries & Power": 0.7,
    "Furniture": 0.6, "Hard Drives & Data Storage": 0.6, "Security, Banking & Cash": 0.6,
    "Networking & WiFi": 0.5, "Healthcare Supplies": 0.5, "Retail Store Supplies": 0.5,
    "Party Supplies": 0.4, "Phones, Cameras & Electronics": 0.4, "Tablets & iPads": 0.4,
    "Bags, Backpacks & Luggage": 0.6, "Computer Software": 0.5, "Audio": 0.4,
    "Gift Shop": 0.3, "Arts & Crafts": 0.3,
}
CORE_DEFAULT = 0.3
CORE_OVERRIDES = [  # (regex on Staples path, coreness) - last match wins
    (r"^Furniture > Chairs & Seating > (Office|Big and Tall|Stacking)", 1.0),
    (r"^Furniture > Desks > (Office Desks|Sit & Stand)", 0.9),
    (r"^Furniture > (File Cabinets|Chair Mats|Standing Desk Mats|Boards & Easels)", 0.9),
    (r"^Furniture > (Storage Furniture|Cubicle & Panel Systems|Carts & Stands|School Furniture)", 0.7),
    (r"^Furniture > Lamps & Lighting > Desk Lamps", 0.6),
    # the "white chair" zone: design / lifestyle shelves carry little 1P revenue
    (r"^Furniture > Decor", 0.15),
    (r"^Furniture > Lamps & Lighting > (?!Desk Lamps)", 0.2),
    (r"Accent|Sofa|Ottoman|Recliner|Loveseat|Bench|Coffee Tables", 0.2),
    (r"^Decor|^Fitness|^Expanded Assortment", 0.1),
]
GENERIC_NAMES = {"accessories", "other", "parts", "supplies", "storage", "more", "misc",
                 "equipment", "systems", "solutions", "kits", "refills", "tools", "products",
                 "components", "essentials", "basics", "appliances"}
STOP = {"and", "the", "of", "for", "with", "a", "an", "in", "to", "by", "all", "more", "other", "s"}

rng = np.random.default_rng(CONFIG["seed"])
os.makedirs(CONFIG["out_dir"], exist_ok=True)


def log(msg):
    print(msg, flush=True)


# %% [2] SHARED: FILE DISCOVERY + LOADING  (identical in method_g_graph.py) ----------
def find_file(patterns, data_dir):
    """Look in data_dir and one folder below it (plus /content in Colab). Deliberately not a
    full recursive crawl - that would walk an entire mounted Google Drive."""
    roots = [data_dir, os.path.join(data_dir, "*"), "/content"]
    for root in roots:
        for pat in patterns:
            hits = sorted(glob.glob(os.path.join(root, pat)))
            if hits:
                return hits[0]
    return None


def colab_upload_if_missing(patterns):
    try:
        from google.colab import files  # noqa
    except ImportError:
        return None
    log(f"  Please upload a file matching {patterns}")
    up = files.upload()
    for name in up:
        return name
    return None


def read_table(path, sheet=None, header=0):
    if path.lower().endswith(".csv"):
        return pd.read_csv(path)
    return pd.read_excel(path, sheet_name=sheet, header=header)


def path_str(parts):
    return " > ".join(parts)


def load_retailer(rc):
    """Read one retailer tree into the standard node table."""
    f = find_file(rc["file_glob"], CONFIG["data_dir"]) or colab_upload_if_missing(rc["file_glob"])
    if f is None:
        raise FileNotFoundError(f"No file for {rc['name']} matching {rc['file_glob']}")
    raw = read_table(f, rc.get("sheet"))
    lv = [c for c in rc["level_cols"] if c in raw.columns]
    raw = raw.dropna(subset=[lv[0]])
    rows = []
    for _, r in raw.iterrows():
        parts = tuple(str(r[c]).strip() for c in lv if pd.notna(r[c]) and str(r[c]).strip())
        rows.append({
            "retailer": rc["name"], "path": parts, "count_raw": pd.to_numeric(r.get(rc["count_col"]), errors="coerce"),
            "native_id": str(r.get(rc["id_col"], "")), "url": r.get(rc["url_col"], ""),
        })
    df = pd.DataFrame(rows).drop_duplicates("path")
    # add implied parents (a breadcrumb parent with no page of its own)
    have = set(df["path"])
    extra = []
    for p in list(have):
        for i in range(1, len(p)):
            if p[:i] not in have:
                have.add(p[:i])
                extra.append({"retailer": rc["name"], "path": p[:i], "count_raw": np.nan,
                              "native_id": "", "url": "", "implied": True})
    if extra:
        df = pd.concat([df, pd.DataFrame(extra)], ignore_index=True)
    df["implied"] = df["implied"].fillna(False) if "implied" in df else False
    df["level"] = df["path"].map(len)
    df["name"] = df["path"].map(lambda p: p[-1])
    df["path_str"] = df["path"].map(path_str)
    df["parent_path"] = df["path"].map(lambda p: p[:-1] if len(p) > 1 else None)
    df["uid"] = rc["name"] + "::" + df["path_str"]
    df["parent_uid"] = df["parent_path"].map(lambda p: rc["name"] + "::" + path_str(p) if p else None)
    kids = df.groupby("parent_uid")["uid"].apply(list).to_dict()
    df["children"] = df["uid"].map(lambda u: kids.get(u, []))
    df["is_leaf"] = df["children"].map(len) == 0
    # scope
    pats = [re.compile(x) for x in rc.get("scope_exclude", [])]
    df["in_scope"] = ~df["path_str"].map(lambda s: any(p.search(s) for p in pats))
    df["source_file"] = os.path.basename(f)
    log(f"  {rc['name']:<12} {len(df):>5} nodes  ({df['is_leaf'].sum()} leaves, "
        f"{(~df['in_scope']).sum()} out of scope)  <- {os.path.basename(f)}")
    return df.reset_index(drop=True), f


def load_crosslist_edges(rc, file, nodes):
    """Staples cross-listings: extra (child -> parent) placements. The sheet names parents by
    NAME, and ~50 names are shared by two shelves (e.g. 'Cleaning Supplies' is an L1 and also a
    shelf inside Lab Supplies). Each name resolves to ONE shelf: a non-leaf in the child's own
    L1 if there is one, otherwise the shallowest non-leaf, otherwise the shallowest shelf."""
    cl = rc.get("crosslist")
    if not cl:
        return []
    x = read_table(file, cl["sheet"], header=cl["header_row"])
    by_name = defaultdict(list)
    for u, n, leaf, lvl, pth in zip(nodes["uid"], nodes["name"], nodes["is_leaf"], nodes["level"], nodes["path"]):
        by_name[n].append((u, leaf, lvl, pth[0]))
    canon_parent = dict(zip(nodes["uid"], nodes["parent_uid"]))
    l1_of = dict(zip(nodes["uid"], nodes["path"].map(lambda p: p[0])))

    def resolve(pname, child_uid):
        cands = by_name.get(pname, [])
        if not cands:
            return None
        same = [c for c in cands if c[3] == l1_of[child_uid] and not c[1]]
        pool = same or [c for c in cands if not c[1]] or cands
        return sorted(pool, key=lambda c: c[2])[0][0]

    edges = []
    for _, r in x.iterrows():
        child_name, parents = r.get(cl["name_col"]), str(r.get(cl["parents_col"], ""))
        if pd.isna(child_name):
            continue
        plist = [p.strip() for p in parents.split(";") if p.strip()]
        if len(plist) > cl.get("max_parents", 999):
            continue
        cands = by_name.get(child_name, [])
        # the cross-listed child is the shelf with that name; prefer a leaf / deepest one
        child_uids = [sorted(cands, key=lambda c: (-int(c[1]), -c[2]))[0][0]] if cands else []
        for cu in child_uids:
            for p in plist:
                pu = resolve(p, cu)
                if pu and pu != cu and pu != canon_parent.get(cu) and not pu.startswith(cu + " > ") \
                        and not cu.startswith(pu + " > "):
                    edges.append((cu, pu, 1.0 / len(plist)))
    return list({(c, p): (c, p, w) for c, p, w in edges}.values())


def harmonise_counts(nodes, semantics, xedges=None):
    """
    Give every node a comparable 'subtree_items' figure.
      terminal_only : sum of leaf counts under the node, following BOTH canonical and
                      cross-listed placements, de-duplicated by leaf (Staples).
      cumulative    : the stated count; if blank, the sum of children (Office Depot).
    """
    uid_idx = {u: i for i, u in enumerate(nodes["uid"])}
    counts = nodes["count_raw"].fillna(0).values
    if semantics == "terminal_only":
        kids = defaultdict(set)
        for u, p in zip(nodes["uid"], nodes["parent_uid"]):
            if p:
                kids[p].add(u)
        for c, p, _ in (xedges or []):
            kids[p].add(c)
        leaf_mask = nodes["count_raw"].notna().values
        memo = {}

        def leaves_under(u, stack=()):
            if u in memo:
                return memo[u]
            s = {u} if leaf_mask[uid_idx[u]] else set()
            for c in kids.get(u, ()):
                if c not in stack:
                    s |= leaves_under(c, stack + (u,))
            memo[u] = s
            return s

        sys.setrecursionlimit(10000)
        sets = [leaves_under(u) for u in nodes["uid"]]
        nodes["subtree_items"] = [float(sum(counts[uid_idx[v]] for v in s)) for s in sets]
        nodes["leaf_set"] = sets
        nodes["count_estimated"] = False
    else:
        sub = {}
        est = {}
        for u, stated, ch in sorted(zip(nodes["uid"], nodes["count_raw"], nodes["children"]),
                                    key=lambda t: -t[0].count(" > ")):
            if pd.notna(stated):
                sub[u], est[u] = float(stated), False
            else:
                sub[u], est[u] = float(sum(sub.get(c, 0.0) for c in ch)), True
        nodes["subtree_items"] = nodes["uid"].map(sub)
        nodes["count_estimated"] = nodes["uid"].map(est)
        nodes["leaf_set"] = None
    return nodes


def retailer_total(nodes):
    """Catalogue size used to normalise depth into 'items per 10k'.
    terminal_only: distinct in-scope leaf counts (cross-listed leaves counted once).
    cumulative   : sum of the in-scope L1 totals."""
    if nodes["leaf_set"].notna().any():
        return float(nodes.loc[nodes["in_scope"] & nodes["count_raw"].notna(), "count_raw"].sum())
    l1 = nodes[(nodes["level"] == 1) & nodes["in_scope"]]
    return float(l1["subtree_items"].sum())


# %% [3] SHARED: TEXT NORMALISATION -------------------------------------------------
def tokens(s):
    s = str(s).lower().replace("™", "").replace("®", "").replace("’", "'")
    return [t for t in re.findall(r"[a-z0-9]+", s) if t not in STOP]


def stem(t):
    if len(t) > 4 and t.endswith("ies"):
        return t[:-3] + "y"
    if len(t) > 4 and re.search(r"(ses|xes|ches|shes)$", t):
        return t[:-2]
    if len(t) > 3 and t.endswith("s") and not t.endswith("ss"):
        return t[:-1]
    return t


def expanded_name(path):
    """Generic leaf names ('Accessories', 'Storage') borrow their parent's words."""
    name = path[-1]
    toks = [stem(t) for t in tokens(name)]
    if len(path) > 1 and (not toks or all(t in {stem(g) for g in GENERIC_NAMES} for t in toks)):
        return f"{path[-2]} {name}"
    return name


# %% [4] ENCODERS -------------------------------------------------------------------
class Encoder:
    """Pluggable text encoder. All scores downstream are calibrated per encoder."""

    def __init__(self, kind, corpus):
        self.kind = None
        order = [kind] + [k for k in ("spacy", "st", "tfidf") if k != kind]
        for k in order:
            try:
                getattr(self, f"_init_{k}")(corpus)
                self.kind = k
                break
            except Exception as e:  # noqa
                log(f"  encoder '{k}' unavailable ({type(e).__name__}: {str(e)[:80]}) - trying next")
        log(f"  encoder in use: {self.kind}")

    # spaCy word vectors, IDF-weighted average (offline once the model is installed)
    def _init_spacy(self, corpus):
        import spacy
        self.nlp = spacy.load(CONFIG["spacy_model"], exclude=["tok2vec", "tagger", "parser", "ner",
                                                              "lemmatizer", "attribute_ruler", "senter"])
        if self.nlp.vocab.vectors.shape[0] < 50000:
            raise RuntimeError("model has too few vectors - install en_core_web_lg")
        df = defaultdict(int)
        for t in corpus:
            for w in set(tokens(t)):
                df[w] += 1
        n = len(corpus)
        self.idf = {w: math.log((1 + n) / (1 + c)) + 1 for w, c in df.items()}
        self.dim = self.nlp.vocab.vectors.shape[1]

    def _enc_spacy(self, texts):
        out = np.zeros((len(texts), self.dim), dtype=np.float32)
        for i, t in enumerate(texts):
            acc, wsum = np.zeros(self.dim, dtype=np.float32), 0.0
            for w in tokens(t):
                lex = self.nlp.vocab[w]
                if lex.has_vector:
                    wt = self.idf.get(w, 1.0)
                    acc += wt * lex.vector
                    wsum += wt
            out[i] = acc / wsum if wsum else 0
        return out

    # sentence-transformers (recommended when HuggingFace is reachable)
    def _init_st(self, corpus):
        from sentence_transformers import SentenceTransformer
        self.model = SentenceTransformer(CONFIG["st_model"])

    def _enc_st(self, texts):
        return self.model.encode(list(texts), batch_size=64, normalize_embeddings=True,
                                 show_progress_bar=False).astype(np.float32)

    # TF-IDF char n-grams + SVD (no downloads at all)
    def _init_tfidf(self, corpus):
        from sklearn.decomposition import TruncatedSVD
        from sklearn.feature_extraction.text import TfidfVectorizer
        self.vec = TfidfVectorizer(analyzer="char_wb", ngram_range=(3, 5), min_df=1, sublinear_tf=True)
        X = self.vec.fit_transform([t.lower() for t in corpus])
        self.svd = TruncatedSVD(n_components=min(256, X.shape[1] - 1), random_state=0).fit(X)

    def _enc_tfidf(self, texts):
        return self.svd.transform(self.vec.transform([t.lower() for t in texts])).astype(np.float32)

    def encode(self, texts):
        E = getattr(self, f"_enc_{self.kind}")(texts)
        n = np.linalg.norm(E, axis=1, keepdims=True)
        n[n == 0] = 1
        return E / n


def node_embeddings(nodes, enc):
    """Own name + parent + L1 context, mixed then re-normalised."""
    names = nodes["path"].map(expanded_name).tolist()
    parents = nodes["path"].map(lambda p: expanded_name(p[:-1]) if len(p) > 1 else expanded_name(p)).tolist()
    l1s = nodes["path"].map(lambda p: p[0]).tolist()
    uniq = sorted(set(names) | set(parents) | set(l1s))
    E = enc.encode(uniq)
    look = {t: E[i] for i, t in enumerate(uniq)}
    a, b, c = CONFIG["name_weights"]
    V = np.stack([a * look[n] + b * look[p] + c * look[l] for n, p, l in zip(names, parents, l1s)])
    V /= np.clip(np.linalg.norm(V, axis=1, keepdims=True), 1e-9, None)
    return V, names


# %% [5] VECTOR INDEX -----------------------------------------------------------------
class VectorIndex:
    """Thin wrapper: Qdrant local mode (payload filter by retailer) or NumPy brute force.
    The same collection is meant to hold Track B SKU vectors later - that is where a
    real vector DB earns its keep (millions of points, HNSW, filters)."""

    def __init__(self, backend, dim):
        self.backend, self.dim = backend, dim
        self.vecs, self.payload = [], []
        if backend == "qdrant":
            try:
                from qdrant_client import QdrantClient
                from qdrant_client.models import Distance, VectorParams
                self.client = QdrantClient(":memory:")
                self.client.create_collection("category_nodes",
                                              vectors_config=VectorParams(size=dim, distance=Distance.COSINE))
            except Exception as e:
                log(f"  qdrant unavailable ({e}); using numpy")
                self.backend = "numpy"
        log(f"  vector backend: {self.backend}")

    def add(self, vectors, payloads):
        start = len(self.payload)
        self.vecs.append(vectors)
        self.payload.extend(payloads)
        if self.backend == "qdrant":
            from qdrant_client.models import PointStruct
            pts = [PointStruct(id=start + i, vector=vectors[i].tolist(), payload=payloads[i])
                   for i in range(len(payloads))]
            for j in range(0, len(pts), 512):
                self.client.upsert("category_nodes", points=pts[j:j + 512])

    def query(self, Q, k, retailer):
        """Return (row indices into the retailer's slice order, scores), shape (len(Q), k)."""
        allv = np.vstack(self.vecs)
        mask = np.array([p["retailer"] == retailer for p in self.payload])
        ids_r = np.where(mask)[0]
        pos = {g: i for i, g in enumerate(ids_r)}
        k = min(k, len(ids_r))
        if self.backend == "qdrant":
            from qdrant_client.models import FieldCondition, Filter, MatchValue
            flt = Filter(must=[FieldCondition(key="retailer", match=MatchValue(value=retailer))])
            I = np.zeros((len(Q), k), dtype=int)
            S = np.zeros((len(Q), k), dtype=np.float32)
            for i, q in enumerate(Q):
                res = self.client.query_points("category_nodes", query=q.tolist(), query_filter=flt,
                                               limit=k).points
                I[i] = [pos[r.id] for r in res]
                S[i] = [r.score for r in res]
            return I, S
        M = Q @ allv[ids_r].T
        I = np.argsort(-M, axis=1)[:, :k]
        return I, np.take_along_axis(M, I, axis=1)


# %% [6] LEXICAL SIMILARITY -----------------------------------------------------------
def lexical_model(texts):
    """Wording similarity: stemmed word uni/bi-grams + character 3-5 grams, averaged.
    Char n-grams catch 'Pads' ~ 'Notepads'; words catch exact category vocabulary."""
    from sklearn.feature_extraction.text import TfidfVectorizer

    def words(s):
        t = [stem(x) for x in tokens(s)]
        return t + [a + "_" + b for a, b in zip(t, t[1:])]

    class Lex:
        def __init__(self):
            self.w = TfidfVectorizer(analyzer=words, sublinear_tf=True).fit(texts)
            self.c = TfidfVectorizer(analyzer="char_wb", ngram_range=(3, 5), sublinear_tf=True).fit(
                [t.lower() for t in texts])

        def transform(self, xs):
            from scipy.sparse import hstack
            xs = list(xs)
            return hstack([self.w.transform(xs) * np.sqrt(0.5),
                           self.c.transform([x.lower() for x in xs]) * np.sqrt(0.5)]).tocsr()

    return Lex()


# %% [7] CALIBRATION ------------------------------------------------------------------
def ablation_prefix(gold_path):
    """The Staples department (L2, or L1 if the gold shelf is an L1) that we hide."""
    first = str(gold_path).split(" > ")
    return " > ".join(first[:2])


def ablation_mask(focal, gold_path, comp_path):
    """True = keep. Hide every gold alternative's department plus any Staples shelf that
    carries the same name as the answer or the competitor shelf (cross-department twins)."""
    hide = np.zeros(len(focal), dtype=bool)
    names = {comp_path.split(" > ")[-1].lower()}
    for alt in str(gold_path).split("|"):
        pre = ablation_prefix(alt)
        hide |= (focal["path_str"].eq(pre) | focal["path_str"].str.startswith(pre + " > ")).values
        names.add(alt.split(" > ")[-1].lower())
    hide |= focal["name"].str.lower().isin(names).values
    return ~hide


def build_calibration_set(gold, comp_scores, ablation_scores):
    """Positives = gold rows Staples carries (their real best score).
    Negatives = gold rows Staples does NOT carry  +  'ablation negatives': the same carried
    rows re-scored with the Staples department that holds their answer hidden. Hiding the
    answer shows what a genuine miss looks like, which gives a balanced set even when the
    competitor has almost no real gaps (true for Office Depot)."""
    rows = []
    for _, r in gold.iterrows():
        sc = comp_scores.get(r["comp_path"])
        if sc is None:
            continue
        rows.append({"comp_path": r["comp_path"], "y": int(r["staples_carries"]), "score": sc, "kind": "gold"})
        if int(r["staples_carries"]) == 1 and r["comp_path"] in ablation_scores:
            rows.append({"comp_path": r["comp_path"], "y": 0, "score": ablation_scores[r["comp_path"]],
                         "kind": "ablation"})
    return pd.DataFrame(rows)


def calibrate_threshold(cal_df):
    """Pick the 'same shelf' threshold with the lowest expected cost:
         cost = false_gap_cost x (carried shelves called gaps) + (real gaps called carried)
    i.e. a cost-weighted Youden's J. Re-run whenever the encoder changes (threshold hygiene)."""
    y, s = cal_df["y"].values, cal_df["score"].values
    c = CONFIG["false_gap_cost"]
    best = (float(np.median(s)), -1e9)
    for t in np.unique(np.round(s, 4)):
        pred = s >= t
        j = c * pred[y == 1].mean() - pred[y == 0].mean()
        if j > best[1]:
            best = (float(t), j)
    tau = best[0]
    try:
        from sklearn.metrics import roc_auc_score
        auc = float(roc_auc_score(y, s))
    except Exception:
        auc = float("nan")
    g = cal_df[cal_df["kind"] == "gold"]
    pred_g = g["score"].values >= tau
    yg = g["y"].values
    fn = int(((~pred_g) & (yg == 1)).sum())
    fp = int((pred_g & (yg == 0)).sum())
    m = {"tau": tau, "youden_j": best[1], "auc_augmented": auc, "n_positive": int((y == 1).sum()),
         "n_negative_gold": int(((y == 0) & (cal_df["kind"] == "gold")).sum()),
         "n_negative_ablation": int((cal_df["kind"] == "ablation").sum()),
         "gold_accuracy": float((pred_g == (yg == 1)).mean()),
         "false_gap_rate": fn / max(int((yg == 1).sum()), 1),       # carried shelves we would call gaps
         "missed_gap_rate": fp / max(int((yg == 0).sum()), 1),      # real gaps we would call carried
         "ablation_neg_below_tau": float((cal_df.loc[cal_df.kind == "ablation", "score"] < tau).mean())}
    return tau, m


def gap_screen_threshold(cal_df, gold):
    """Score below which a competitor shelf is flagged 'possibly not carried'.
    Set by a false-alarm budget: the q-th percentile of scores of shelves we KNOW Staples
    carries (q = gap_screen_false_alarm). Also returns the base rate of 'carried' among the
    randomly drawn gold rows - how much real whitespace this competitor offers at all."""
    pos = cal_df[(cal_df.kind == "gold") & (cal_df.y == 1)]["score"]
    s_enter = float(pos.quantile(CONFIG["gap_screen_false_alarm"]))
    neg = cal_df[(cal_df.kind == "gold") & (cal_df.y == 0)]["score"]
    r = gold[gold.get("source", pd.Series("", index=gold.index)).eq(CONFIG["prior_source"])]
    base = float(r["staples_carries"].mean()) if len(r) else float("nan")
    return s_enter, {"gap_screen_score": s_enter,
                     "gap_screen_recall_on_real_gaps": float((neg < s_enter).mean()) if len(neg) else float("nan"),
                     "base_rate_carried_random_sample": base, "n_random_sample": int(len(r))}


MIN_GOLD_ROWS = 40
_CAL_MEMORY = {}   # calibration of the last well-labelled competitor (same encoder, same spine)


def top1_accuracy(gold, best_map):
    """Share of carried gold rows whose top-1 Staples shelf is the gold shelf, or its
    parent / child (one of the '|' alternatives)."""
    ok, n = 0, 0
    for _, r in gold[(gold.staples_carries == 1) & gold.gold_staples_path.notna()].iterrows():
        pred = best_map.get(r["comp_path"])
        if pred is None:
            continue
        n += 1
        alts = str(r["gold_staples_path"]).split("|")
        if any(pred == a or pred.startswith(a + " > ") or a.startswith(pred + " > ") for a in alts):
            ok += 1
    return ok / max(n, 1), n


# %% [8] MAIN PIPELINE ------------------------------------------------------------------
def main():
    log("\n[1] Loading retailer trees")
    all_nodes, files = [], {}
    for rc in RETAILERS:
        df, f = load_retailer(rc)
        files[rc["name"]] = f
        xe = load_crosslist_edges(rc, f, df) if rc.get("crosslist") else []
        if xe:
            log(f"      + {len(xe)} cross-listing placements used for counting")
        df = harmonise_counts(df, rc["count_semantics"], xe)
        df["role"] = rc["role"]
        all_nodes.append(df)
    nodes = pd.concat(all_nodes, ignore_index=True)
    totals = {r: retailer_total(nodes[nodes.retailer == r]) for r in nodes.retailer.unique()}
    for r, t in totals.items():
        log(f"      {r:<12} in-scope catalogue: {t:,.0f} items")

    log("\n[2] Embedding every category node")
    nodes["name_expanded"] = nodes["path"].map(expanded_name)
    corpus = sorted(set(nodes["name_expanded"]) | set(nodes["path"].map(lambda p: p[0])))
    enc = Encoder(CONFIG["encoder"], corpus)
    V, _ = node_embeddings(nodes, enc)
    lex = lexical_model(corpus)

    focal = nodes[(nodes.retailer == FOCAL) & nodes.in_scope].reset_index()
    focal_V = V[focal["index"].values]
    focal_L = lex.transform(focal["name_expanded"])

    # coreness of every Staples node
    depth_pct = focal.groupby("level")["subtree_items"].rank(pct=True).fillna(0).values
    base = focal["path"].map(lambda p: CORE_L1_WEIGHT.get(p[0], CORE_DEFAULT)).values.astype(float)
    for pat, w in CORE_OVERRIDES:
        hit = focal["path_str"].str.contains(pat, regex=True).values
        base[hit] = w
    focal["coreness"] = base * (0.6 + 0.4 * depth_pct)

    idx = VectorIndex(CONFIG["vector_backend"], V.shape[1])
    idx.add(focal_V, [{"retailer": FOCAL, "path": s} for s in focal["path_str"]])

    results = []
    for rc in RETAILERS:
        comp = rc["name"]
        if comp == FOCAL:
            continue
        log(f"\n[3] Retrieve-then-rerank: {comp} -> {FOCAL}")
        cn = nodes[(nodes.retailer == comp) & nodes.in_scope].reset_index(drop=False)
        Q = V[cn["index"].values]
        I, S_sem = idx.query(Q, CONFIG["k_retrieve"], FOCAL)
        CL = lex.transform(cn["name_expanded"])
        S_lex = np.zeros_like(S_sem)
        for i in range(len(cn)):
            S_lex[i] = (focal_L[I[i]] @ CL[i].T).toarray().ravel()
        S = CONFIG["w_sem"] * S_sem + CONFIG["w_lex"] * S_lex
        order = np.argsort(-S, axis=1)
        I, S, S_sem, S_lex = (np.take_along_axis(a, order, axis=1) for a in (I, S, S_sem, S_lex))
        cn["best_score"] = S[:, 0]
        cn["best_sem"], cn["best_lex"] = S_sem[:, 0], S_lex[:, 0]
        cn["best_staples"] = focal["path_str"].values[I[:, 0]]
        cn["best_staples_uid"] = focal["uid"].values[I[:, 0]]
        cn["alt_staples"] = [" | ".join(focal["path_str"].values[I[i, 1:3]]) for i in range(len(cn))]
        k = CONFIG["k_adjacency"]
        cn["A_raw"] = S[:, :k].mean(axis=1)                                   # mean-pool
        core_mat = focal["coreness"].values[I]                                # for max-pool CRS below

        # ---- calibration on the gold set (+ ablation negatives)
        gold_f = find_file(["gold_matches.csv"], CONFIG["data_dir"])
        gold = pd.read_csv(gold_f) if gold_f else pd.DataFrame(
            columns=["competitor", "comp_path", "staples_carries", "gold_staples_path", "source"])
        gold = gold[gold["competitor"] == comp]
        if len(gold) >= MIN_GOLD_ROWS:
            pos = gold[(gold.staples_carries == 1) & gold.gold_staples_path.notna()]
            row_of = dict(zip(cn["path_str"], range(len(cn))))
            abl = {}
            FL_all = focal_L
            for _, r in pos.iterrows():
                i = row_of.get(r["comp_path"])
                if i is None:
                    continue
                mask = ablation_mask(focal, r["gold_staples_path"], r["comp_path"])
                sem = focal_V[mask] @ Q[i]
                lx = (FL_all[mask] @ CL[i].T).toarray().ravel()
                abl[r["comp_path"]] = float((CONFIG["w_sem"] * sem + CONFIG["w_lex"] * lx).max())
            cal_df = build_calibration_set(gold, dict(zip(cn["path_str"], cn["best_score"])), abl)
            tau, cal = calibrate_threshold(cal_df)
            s_enter, sinfo = gap_screen_threshold(cal_df, gold)
            cal.update(sinfo)
            cal["neg_median_score"] = float(cal_df.loc[cal_df.y == 0, "score"].median())
            cal["top1_accuracy"], cal["n_top1"] = top1_accuracy(gold, dict(zip(cn["path_str"], cn["best_staples"])))
            gscored = cal_df
            log(f"      confident-pairing tau = {tau:.3f} | AUC {cal['auc_augmented']:.3f} | top-1 shelf accuracy "
                f"{cal['top1_accuracy']:.1%} (pos {cal['n_positive']}, neg {cal['n_negative_gold']} real + "
                f"{cal['n_negative_ablation']} ablation)")
            _CAL_MEMORY.update(tau=tau, s_enter=s_enter, neg=cal["neg_median_score"], source=comp)
        elif _CAL_MEMORY:
            # same scorer + same Staples spine -> another competitor's calibration is a usable
            # provisional default; replace it by labelling ~150 rows for this competitor
            tau, s_enter = _CAL_MEMORY["tau"], _CAL_MEMORY["s_enter"]
            cal = {"tau": tau, "neg_median_score": _CAL_MEMORY["neg"], "gap_screen_score": s_enter,
                   "note": f"PROVISIONAL - calibration borrowed from {_CAL_MEMORY['source']}; add ~150 labelled "
                           f"{comp} rows to gold_matches.csv (see the Review_Queue sheet for candidates)"}
            gscored = pd.DataFrame()
            log(f"      !! only {len(gold)} gold rows for {comp} - borrowing calibration from "
                f"{_CAL_MEMORY['source']} (tau={tau:.3f}); label this competitor before relying on it")
        else:
            tau = float(np.quantile(cn["best_score"], 0.10))
            cal = {"tau": tau, "neg_median_score": float(np.quantile(cn["best_score"], 0.02)),
                   "note": "NO GOLD ROWS - fallback 10th percentile; results indicative only"}
            s_enter = float(np.quantile(cn["best_score"], 0.05))
            gscored = pd.DataFrame()
            log(f"      !! no gold rows for {comp} - using fallback tau={tau:.3f}; results indicative only")
        cn["matched"] = cn["best_score"] >= tau
        cn["gap_flag"] = cn["best_score"] < s_enter              # possibly not carried at all
        cn["margin_to_screen"] = cn["best_score"] - s_enter      # review queue = smallest |margin|
        m = CONFIG["multi_match_margin"]
        cn["match_set"] = [[focal["uid"].values[I[i, j]] for j in range(I.shape[1])
                            if S[i, j] >= tau and S[i, j] >= S[i, 0] - m] for i in range(len(cn))]
        cn["n_shelves_matched"] = cn["match_set"].map(len)
        if "gap_screen_recall_on_real_gaps" in cal:
            log(f"      gap screen: score < {s_enter:.3f} | catches {cal['gap_screen_recall_on_real_gaps']:.0%} of real "
                f"gold gaps at a {CONFIG['gap_screen_false_alarm']:.0%} false-alarm budget | "
                f"{cal['base_rate_carried_random_sample']:.1%} of randomly drawn {comp} shelves exist at Staples")
        # CRS: 0 = a typical non-match, 100 = a certain match - to a fully-core shelf
        lo = cal["neg_median_score"]
        f = np.clip((S - lo) / max(tau - lo, 1e-6), 0, 1)
        cn["crs_raw"] = (f * core_mat).max(axis=1)
        cn["crs_driver"] = [focal["path_str"].values[I[i, np.argmax(f[i] * core_mat[i])]] for i in range(len(cn))]

        # ---- roll matches up the competitor tree
        cn = rollup_and_score(cn, focal, totals[comp], totals[FOCAL], tau)
        cn["competitor"] = comp
        results.append((comp, cn, cal, gscored, tau))

    # Staples-ahead view (reverse direction) for context
    staples_ahead = reverse_view(nodes, V, lex, focal, focal_V)

    node_scores, unit_sets = [], []
    for comp, cn, cal, gscored, tau in results:
        units = select_units(cn)
        units = decide_and_rank(units, comp)
        sens = sensitivity(units)
        units = units.merge(sens, on="uid", how="left")
        unit_sets.append((comp, units, cal, gscored, tau))
        write_outputs(comp, cn, units, cal, gscored, staples_ahead, enc.kind, idx.backend, totals)
        charts(units, comp)
        node_scores.append(cn[["competitor", "uid", "path_str", "level", "is_leaf", "comp_items", "best_staples",
                               "best_score", "matched", "coverage", "DG", "WM", "gap_type", "AAS", "CRS", "O_node"]]
                           .merge(units[["uid", "label", "O"]], on="uid", how="left"))
        log("\n[7] Top recommendations (approved labels only)")
        show = units[units.approved].sort_values("O", ascending=False).head(CONFIG["top_n"])
        log(show[["path_str", "label", "comp_items", "staples_items", "coverage", "DG", "AAS", "CRS", "O",
                  "p_top_n"]].round(2).to_string(index=False))
    pd.concat(node_scores).to_csv(os.path.join(CONFIG["out_dir"], "node_scores_method_v.csv"), index=False)

    # cross-competitor gap concepts (needs >= 2 competitors): link recommendation units
    # across competitors with the same hybrid score used for matching
    def sim_v(df):
        E = np.stack([V[r["index"]] for r in df["row"]])
        L = lex.transform([r["name_expanded"] for r in df["row"]])
        return CONFIG["w_sem"] * (E @ E.T) + CONFIG["w_lex"] * (L @ L.T).toarray()

    xw = cross_competitor_whitespace(unit_sets, sim_v, {r[0]: r[4] for r in unit_sets})
    if len(xw):
        xw.to_excel(os.path.join(CONFIG["out_dir"], "method_v_cross_competitor_gaps.xlsx"), index=False)
        log(f"\n[6b] Cross-competitor gap concepts: {len(xw)} concepts, "
            f"{(xw.n_competitors >= 2).sum()} carried by 2+ competitors")
    log(f"\nDone. Outputs in ./{CONFIG['out_dir']}/")


# %% [9] GAP ROLL-UP -------------------------------------------------------------------
def rollup_and_score(cn, focal, comp_total, focal_total, tau):
    """For every competitor node: expected coverage, Staples depth, depth gap, whitespace
    mass, and exposure-weighted adjacency / cannibalization over its leaves."""
    uid2row = {u: i for i, u in enumerate(cn["uid"])}
    focal_leafsets = dict(zip(focal["uid"], focal["leaf_set"]))
    focal_counts = dict(zip(focal["uid"], focal["count_raw"]))

    # AAS anchoring: 100 = as embedded in Staples' neighbourhood as a TYPICAL shelf Staples
    # already carries (median), 0 = the 5th percentile of all competitor leaves. Anchored to
    # reference groups, not to raw cosine values, so it survives an encoder swap.
    leaves = cn[cn["is_leaf"]]
    ref100 = float(leaves.loc[leaves["matched"], "A_raw"].median())
    ref0 = float(leaves["A_raw"].quantile(0.05))
    cn["AAS_leaf"] = (100 * ((cn["A_raw"] - ref0) / max(ref100 - ref0, 1e-9))).clip(0, 100)
    cn["CRS_leaf"] = (100 * cn["crs_raw"]).clip(0, 100)

    # leaves under each competitor node (paths are prefix-closed)
    leaf_uids = [u for u, lf in zip(cn["uid"], cn["is_leaf"]) if lf]
    desc_leaves = {u: [] for u in cn["uid"]}
    for lu in leaf_uids:
        parts = lu.split(" > ")
        for k in range(1, len(parts) + 1):
            anc = " > ".join(parts[:k])
            if anc in desc_leaves:
                desc_leaves[anc].append(lu)

    items = cn["subtree_items"].fillna(0).values
    matched = cn["matched"].values                      # confident pairing (score >= tau)
    flag = cn["gap_flag"].values                        # possibly not carried
    likely = ~flag & ~matched                           # carried, but exact shelf not pinned

    # Size factor (median-of-ratios, as in DESeq): competitor items per Staples item on the
    # SAME shelf, from confident leaf-to-leaf pairs. Robust to one giant aisle (Staples'
    # 10k wall-art feed) that would distort a total-catalogue ratio.
    st_leaf = dict(zip(focal["uid"], focal["count_raw"]))
    pairs = [(items[i], st_leaf.get(cn.at[i, "best_staples_uid"]))
             for i in range(len(cn)) if cn.at[i, "is_leaf"] and matched[i]]
    pairs = [(a, b) for a, b in pairs if a > 0 and b is not None and pd.notna(b) and b > 0]
    sf = float(np.exp(np.median([np.log(a) - np.log(b) for a, b in pairs]))) if len(pairs) >= 20 \
        else comp_total / focal_total

    out = {k: [] for k in ["leaf_items", "covered_items", "staples_items", "staples_items_scaled", "coverage",
                           "AAS", "CRS", "top_uncovered", "matched_to"]}
    for u in cn["uid"]:
        L = [uid2row[v] for v in desc_leaves[u]]
        w = np.maximum(items[L], 1.0)
        tot = items[L].sum()
        cov = float(items[L][~flag[L]].sum())                      # items on shelves Staples carries
        st_uids = set()
        for i in L:
            if matched[i]:
                st_uids.update(cn.at[i, "match_set"])
        own = uid2row[u]
        if matched[own]:
            st_uids.update(cn.at[own, "match_set"])
        st_items = union_items(st_uids, focal_leafsets, focal_counts, focal)
        parity_fill = float(items[[i for i in L if likely[i]]].sum()) if not matched[own] else 0.0
        out["leaf_items"].append(tot)
        out["covered_items"].append(cov)
        out["staples_items"].append(st_items)
        out["staples_items_scaled"].append(sf * st_items + parity_fill)
        out["coverage"].append(cov / tot if tot > 0 else float(not flag[own]))
        out["AAS"].append(float(np.average(cn["AAS_leaf"].values[L], weights=w)))
        out["CRS"].append(max(float(cn["CRS_leaf"].values[own]),
                              float(np.average(cn["CRS_leaf"].values[L], weights=w))))
        unc = sorted([(items[i], cn.at[i, "name"]) for i in L if flag[i]], reverse=True)[:4]
        out["top_uncovered"].append(", ".join(f"{n} ({int(c)})" for c, n in unc))
        out["matched_to"].append(" | ".join(sorted(st_uids))[:300].replace(FOCAL + "::", ""))
    for k, v in out.items():
        cn[k] = v
    cn["comp_items"] = cn["subtree_items"].fillna(0)
    cn.attrs["size_factor"] = sf
    cn["comp_per10k"] = 1e4 * cn["comp_items"] / comp_total
    cn["staples_per10k"] = 1e4 * cn["staples_items_scaled"] / comp_total
    cn["DG"] = np.log1p(cn["comp_items"]) - np.log1p(cn["staples_items_scaled"])
    cn["WM"] = 1e4 * (cn["leaf_items"] - cn["covered_items"]).clip(lower=0) / comp_total
    enter = cn["gap_flag"] & (cn["coverage"] < CONFIG["enter_coverage_max"])
    pm = dict(zip(cn["uid"], cn["matched"]))
    cn["parent_matched"] = cn["parent_uid"].map(lambda p: bool(pm.get(p, False)))
    cn["gap_type"] = np.where(enter, "ENTER", np.where(cn["DG"] >= CONFIG["deepen_dg_min"], "DEEPEN", "PARITY"))
    gaps = cn["gap_type"] != "PARITY"
    cn["O_node"] = np.nan
    if gaps.any():
        cn.loc[gaps, "O_node"] = opportunity(cn[gaps], CONFIG["rank_weights"], CONFIG["gamma"], False)
    log(f"      size factor (competitor items per Staples item, median-of-ratios): {sf:.3f}")
    return cn


_UNION_CACHE = {}


def union_items(st_uids, focal_leafsets, focal_counts, focal):
    """Items across a set of Staples shelves, de-duplicated at leaf level."""
    if not st_uids:
        return 0.0
    key = frozenset(st_uids)
    if key in _UNION_CACHE:
        return _UNION_CACHE[key]
    leaves = set()
    for u in st_uids:
        s = focal_leafsets.get(u)
        if s:
            leaves |= s
    val = float(sum(focal_counts.get(v, 0) for v in leaves if pd.notna(focal_counts.get(v, np.nan))))
    _UNION_CACHE[key] = val
    return val


# %% [9b] CROSS-COMPETITOR CONSENSUS (active from 2 competitors) -------------------------
ROLE_WEIGHT = {"mirror": 1.0, "scale": 1.0, "lifestyle": 1.0, "specialist": 0.5}


def cross_competitor_whitespace(results, sim_fn, taus):
    """Link recommendation units (gap shelves) ACROSS competitors (a Wayfair shelf and a Target shelf that Staples
    lacks are the same whitespace), then report each whitespace concept with its peer
    coverage: PC = role-weighted share of competitors carrying it, Laplace-smoothed
    (k + 1) / (N + 2) so one competitor never reads as 0% or 100% consensus. This restores
    Track A's Peer Coverage without pairwise matching of full trees: only gap shelves are
    compared, and every competitor is matched to the Staples spine independently."""
    comps = [r[0] for r in results]
    if len(comps) < 2:
        return pd.DataFrame()
    role = {rc["name"]: rc.get("role", "scale") for rc in RETAILERS}
    rows = []
    for comp, cn, *_ in results:
        g = cn[(cn["gap_type"] != "PARITY") & (cn["level"] <= CONFIG["max_unit_level"])]
        for _, r in g.iterrows():
            rows.append({"competitor": comp, "uid": r["uid"], "path_str": r["path_str"], "name": r["name"],
                         "items": r["comp_items"], "gap_type": r["gap_type"], "row": r})
    if not rows:
        return pd.DataFrame()
    df = pd.DataFrame(rows)
    S = sim_fn(df)
    tau = float(np.mean(list(taus.values())))
    parent = list(range(len(df)))

    def find(i):
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i

    comp_arr = df["competitor"].values
    for i in range(len(df)):
        for j in np.where((S[i] >= tau) & (comp_arr != comp_arr[i]))[0]:
            parent[find(i)] = find(j)
    df["concept"] = [find(i) for i in range(len(df))]
    N = sum(ROLE_WEIGHT.get(role[c], 1.0) for c in comps)
    out = []
    for cid, g in df.groupby("concept"):
        carriers = sorted(set(g["competitor"]))
        k = sum(ROLE_WEIGHT.get(role[c], 1.0) for c in carriers)
        out.append({"concept_id": cid, "example_shelf": g.sort_values("items", ascending=False)["path_str"].iloc[0],
                     "n_competitors": len(carriers), "carried_by": ", ".join(carriers),
                     "peer_coverage_PC": (k + 1) / (N + 2), "competitor_items_total": g["items"].sum(),
                     "gap_types": ", ".join(sorted(set(g["gap_type"]))),
                     "member_shelves": " | ".join(f"{c}: {p}" for c, p in zip(g["competitor"], g["path_str"]))[:500]})
    res = pd.DataFrame(out).sort_values(["n_competitors", "competitor_items_total"], ascending=False)
    multi = res["n_competitors"] >= 2
    res["band"] = np.select([multi & (res["peer_coverage_PC"] >= 0.6), multi],
                            ["validated (most peers)", "contested (some peers)"], "single-peer")
    return res


# %% [10] UNIT SELECTION + DECISION + RANKING ------------------------------------------
def select_units(cn):
    """A 'recommendation unit' is the highest node carrying a gap of a given type:
    whole departments when Staples is absent (ENTER), L2+ shelves when it is thin (DEEPEN)."""
    gt = dict(zip(cn["uid"], cn["gap_type"]))
    par = dict(zip(cn["uid"], cn["parent_uid"]))
    keep = []
    for _, r in cn.iterrows():
        if r["gap_type"] == "PARITY" or r["comp_items"] < CONFIG["min_unit_items"]:
            continue
        if r["level"] > CONFIG["max_unit_level"]:
            continue
        if r["gap_type"] == "DEEPEN" and r["level"] < 2:
            continue
        p = par.get(r["uid"])
        if p in gt and gt[p] == r["gap_type"] and not (r["gap_type"] == "DEEPEN" and cn.loc[cn.uid == p, "level"].iloc[0] < 2):
            continue
        keep.append(r["uid"])
    return cn[cn["uid"].isin(keep)].copy()


def label_row(aas, crs, c=CONFIG, aas_hi=None, aas_lo=None, crs_hi=None, crs_lo=None):
    aas_hi = c["aas_hi"] if aas_hi is None else aas_hi
    aas_lo = c["aas_lo"] if aas_lo is None else aas_lo
    crs_hi = c["crs_hi"] if crs_hi is None else crs_hi
    crs_lo = c["crs_lo"] if crs_lo is None else crs_lo
    if crs >= crs_hi:
        return "1P-CORE GAP"
    if crs >= crs_lo:
        return "REVIEW"
    if aas >= aas_hi:
        return "CURATE"
    if aas >= aas_lo:
        return "VERTICAL EXTENSION"
    return "OFF-BRAND"


ACTION = {
    "CURATE": "Approve - open to 3P marketplace sellers",
    "VERTICAL EXTENSION": "Hold for phase 2 - adjacent via a served vertical, not core-adjacent",
    "1P-CORE GAP": "Route to 1P merchandising - competitor deeper in a core line; 3P would cannibalise",
    "REVIEW": "Calibration zone - human review (merchant session)",
    "OFF-BRAND": "Reject - real gap, wrong store",
    "VERIFY": "Check presence - parent shelf is carried; likely a naming difference, not whitespace",
}


def pct(s):
    return s.rank(pct=True).fillna(0)


def attach_demand(units):
    f = find_file(["external_signals.csv"], CONFIG["data_dir"])
    if not f:
        units["demand"] = np.nan
        return units, False
    d = pd.read_csv(f)
    units["demand"] = units["path_str"].map(dict(zip(d["comp_path"], d["demand_index"])))
    return units, units["demand"].notna().any()


def opportunity(units, w, gamma, use_demand):
    base = w["dg"] * pct(units["DG"]) + w["wm"] * pct(units["WM"]) + w["aas"] * units["AAS"] / 100
    if use_demand:
        dw = CONFIG["demand_weight"]
        base = (1 - dw) * base + dw * units["demand"].fillna(units["demand"].median()) / 100
    return base * (1 - units["CRS"] / 100) ** gamma


def decide_and_rank(units, comp):
    units["label"] = [label_row(a, c) for a, c in zip(units["AAS"], units["CRS"])]
    # an ENTER gap sitting under a parent shelf Staples clearly carries is usually a naming
    # difference ("Composition Books" inside Staples' "Notebooks") - verify before acting
    verify = (units["gap_type"] == "ENTER") & units["parent_matched"] & (units["level"] > 1)
    units.loc[verify, "label"] = "VERIFY"
    units.loc[units["label"] == "CURATE", "label"] = "CURATE - " + units["gap_type"]
    units["action"] = units["label"].map(lambda lab: ACTION[lab.split(" - ")[0]])
    units["approved"] = units["label"].str.startswith("CURATE")
    units, use_d = attach_demand(units)
    units["O"] = opportunity(units, CONFIG["rank_weights"], CONFIG["gamma"], use_d)
    units["rank_approved"] = units["O"].where(units["approved"]).rank(ascending=False, method="first")
    return units.sort_values(["approved", "O"], ascending=[False, False])


def sensitivity(units):
    """Re-draw weights (Dirichlet around the base) and jitter the label thresholds.
    Reports how often each unit is approved and lands in the top-N."""
    n, conc, jit = CONFIG["n_sensitivity"], CONFIG["dirichlet_conc"], CONFIG["threshold_jitter"]
    keys = list(CONFIG["rank_weights"])
    base = np.array([CONFIG["rank_weights"][k] for k in keys])
    appr = np.zeros(len(units))
    topn = np.zeros(len(units))
    use_d = units["demand"].notna().any()
    for _ in range(n):
        w = dict(zip(keys, rng.dirichlet(conc * base)))
        th = {k: CONFIG[k] + rng.normal(0, jit) for k in ("aas_hi", "aas_lo", "crs_hi", "crs_lo")}
        lab = np.array([label_row(a, c, aas_hi=th["aas_hi"], aas_lo=th["aas_lo"],
                                  crs_hi=th["crs_hi"], crs_lo=th["crs_lo"])
                        for a, c in zip(units["AAS"], units["CRS"])])
        ok = lab == "CURATE"
        O = opportunity(units, w, CONFIG["gamma"] * rng.uniform(0.5, 2.0), use_d).values
        O = np.where(ok, O, -1)
        top = np.argsort(-O)[:CONFIG["top_n"]]
        top = top[O[top] > -1]
        appr += ok
        topn[top] += 1
    return pd.DataFrame({"uid": units["uid"].values, "p_approved": appr / n, "p_top_n": topn / n})


# %% [11] REVERSE VIEW: WHERE STAPLES IS AHEAD ------------------------------------------
def reverse_view(nodes, V, lex, focal, focal_V):
    rows = []
    for rc in RETAILERS:
        if rc["name"] == FOCAL:
            continue
        cn = nodes[(nodes.retailer == rc["name"]) & nodes.in_scope]
        CV = V[cn.index.values]
        CL = lex.transform(cn["name_expanded"])
        FL = lex.transform(focal["name_expanded"])
        S = CONFIG["w_sem"] * (focal_V @ CV.T) + CONFIG["w_lex"] * (FL @ CL.T).toarray()
        j = S.argmax(axis=1)
        f = focal[["path_str", "level", "subtree_items"]].copy()
        f["best_competitor_path"] = cn["path_str"].values[j]
        f["best_score"] = S[np.arange(len(f)), j]
        f["competitor"] = rc["name"]
        rows.append(f[f["level"] == 2])
    return pd.concat(rows) if rows else pd.DataFrame()


# %% [12] OUTPUTS ------------------------------------------------------------------------
def write_outputs(comp, cn, units, cal, gscored, staples_ahead, enc_kind, backend, totals):
    fp = os.path.join(CONFIG["out_dir"], f"method_v_results_{comp}.xlsx")
    readme = pd.DataFrame({
        "item": ["Method", "Competitor", "Encoder", "Vector backend", "Match threshold tau",
                 "Staples in-scope items", f"{comp} in-scope items", "Units scored", "Approved (CURATE)",
                 "How to read"],
        "value": ["V - vector-first (embeddings, retrieve-then-rerank)", comp, enc_kind, backend,
                  round(cal.get("tau", float("nan")), 4), f"{totals[FOCAL]:,.0f}", f"{totals[comp]:,.0f}",
                  len(units), int(units["approved"].sum()),
                  "Recommendations = approved units ranked by O. AAS/CRS are 0-100. DG = log share gap "
                  "(>0: competitor devotes a bigger share of its catalogue). p_top_n = share of 500 "
                  "re-weighted runs in which the unit stays in the top-N."]})
    cols = ["rank_approved", "path_str", "level", "gap_type", "label", "action", "comp_items", "staples_items",
            "comp_per10k", "staples_per10k", "coverage", "DG", "WM", "AAS", "CRS", "O", "p_approved",
            "p_top_n", "top_uncovered", "matched_to", "crs_driver"]
    rec = units[units["approved"]].sort_values("O", ascending=False)[cols]
    allu = units.sort_values(["label", "O"], ascending=[True, False])[cols]
    nm = cn[["path_str", "level", "is_leaf", "comp_items", "best_staples", "best_score", "best_sem", "best_lex",
             "alt_staples", "matched", "gap_flag", "margin_to_screen", "n_shelves_matched", "coverage", "gap_type",
             "DG", "AAS", "CRS"]]
    rq = cn.assign(abs_margin=cn["margin_to_screen"].abs()).sort_values("abs_margin").head(
        CONFIG["review_queue_size"])[["path_str", "comp_items", "best_staples", "best_score", "margin_to_screen",
                                      "gap_flag", "alt_staples"]]
    rq["reviewer_verdict (carried? Y/N)"] = ""
    calib = pd.DataFrame([cal])
    with pd.ExcelWriter(fp, engine="openpyxl") as xw:
        readme.to_excel(xw, sheet_name="README", index=False)
        rec.round(3).to_excel(xw, sheet_name="Recommendations", index=False)
        allu.round(3).to_excel(xw, sheet_name="All_Gap_Units", index=False)
        nm.round(3).to_excel(xw, sheet_name="Node_Matches", index=False)
        rq.round(3).to_excel(xw, sheet_name="Review_Queue", index=False)
        if len(staples_ahead):
            staples_ahead.sort_values("best_score").round(3).to_excel(xw, sheet_name="Staples_Ahead", index=False)
        calib.round(4).to_excel(xw, sheet_name="Calibration", index=False)
        if len(gscored):
            gscored.round(4).to_excel(xw, sheet_name="Gold_Scored", index=False)
        for ws in xw.book.worksheets:
            for col in ws.columns:
                width = min(60, max(10, max(len(str(c.value or "")) for c in col[:200]) + 2))
                ws.column_dimensions[col[0].column_letter].width = width
            ws.freeze_panes = "B2"
    log(f"      wrote {fp}")


def charts(units, comp):
    """Decision surface + top-N bar. Colour = the decision, grouped into three slots so it
    stays readable for colour-blind viewers (blue approve / orange hold-review / aqua route
    to 1P; off-brand in neutral grey), with direct labels on the approved points."""
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    ink, muted, grid = "#0b0b0b", "#52514e", "#e4e3df"
    group = {"CURATE - ENTER": ("Approve (CURATE)", "#2a78d6"), "CURATE - DEEPEN": ("Approve (CURATE)", "#2a78d6"),
             "VERTICAL EXTENSION": ("Hold / review / verify", "#eb6834"), "REVIEW": ("Hold / review / verify", "#eb6834"),
             "VERIFY": ("Hold / review / verify", "#eb6834"),
             "1P-CORE GAP": ("Route to 1P merchandising", "#1baf7a"), "OFF-BRAND": ("Reject (off-brand)", "#a9a8a3")}
    u = units.copy()
    u["grp"] = u["label"].map(lambda lab: group.get(lab, ("Other", "#a9a8a3"))[0])
    fig, ax = plt.subplots(figsize=(9, 6.5))
    ax.axvspan(CONFIG["crs_lo"], CONFIG["crs_hi"], color="#f3f2ef", zorder=0, lw=0)
    ax.axhline(CONFIG["aas_hi"], color=grid, lw=1, zorder=1)
    ax.axhline(CONFIG["aas_lo"], color=grid, lw=1, ls="--", zorder=1)
    for name in ["Route to 1P merchandising", "Hold / review / verify", "Reject (off-brand)", "Approve (CURATE)"]:
        g = u[u["grp"] == name]
        if len(g):
            col = next(c for n, c in group.values() if n == name)
            ax.scatter(g["CRS"], g["AAS"], s=14 + 4 * np.sqrt(g["comp_items"]), c=col, alpha=0.8,
                       edgecolor="white", linewidth=1.2, label=f"{name} ({len(g)})", zorder=3)
    placed = []
    for _, r in u[u["label"].str.startswith("CURATE")].nlargest(10, "O").sort_values("AAS", ascending=False).iterrows():
        x, y = r["CRS"] + 1.5, r["AAS"]
        while any(abs(y - py) < 3.2 and abs(x - px) < 22 for px, py in placed):
            y -= 3.4
        placed.append((x, y))
        ax.annotate(r["name"], (r["CRS"], r["AAS"]), xytext=(x, y), fontsize=7.5, color=ink, va="center",
                    arrowprops=dict(arrowstyle="-", color="#a9a8a3", lw=0.6) if abs(y - r["AAS"]) > 1 else None)
    ax.text(50, -1.5, "review zone", ha="center", fontsize=7.5, color=muted)
    ax.set_xlabel("Cannibalization risk CRS (0-100): closeness to a Staples core 1P shelf", color=muted)
    ax.set_ylabel("Adjacency AAS (0-100): inside Staples' neighbourhood", color=muted)
    ax.set_title(f"Decision surface - gaps vs {comp} (bubble size = competitor items)", loc="left",
                 fontsize=11, color=ink)
    ax.set_xlim(-3, 103)
    ax.set_ylim(-3, 105)
    ax.tick_params(colors=muted, labelsize=8)
    for sp in ["top", "right"]:
        ax.spines[sp].set_visible(False)
    for sp in ["left", "bottom"]:
        ax.spines[sp].set_color(grid)
    leg = ax.legend(frameon=False, fontsize=8, loc="lower left", title="Decision", title_fontsize=8)
    for h in leg.legend_handles:
        h.set_sizes([40])
    fig.tight_layout()
    fig.savefig(os.path.join(CONFIG["out_dir"], f"method_v_decision_surface_{comp}.png"), dpi=160)
    plt.close(fig)

    top = u[u["approved"]].nlargest(CONFIG["top_n"], "O").iloc[::-1]
    if len(top):
        fig, ax = plt.subplots(figsize=(9, 0.34 * len(top) + 1.3))
        ax.barh(top["path_str"].str.replace(" > ", " › "), top["O"], color="#2a78d6", height=0.6)
        for y, (o, pt) in enumerate(zip(top["O"], top["p_top_n"])):
            ax.text(o + 0.006, y, f"stays top-{CONFIG['top_n']} in {pt:.0%} of runs", va="center", fontsize=7.5,
                    color=muted)
        ax.set_xlabel("Opportunity score O", color=muted)
        ax.set_title(f"Method V - approved categories vs {comp}", loc="left", fontsize=11, color=ink)
        ax.set_xlim(0, top["O"].max() * 1.45)
        ax.tick_params(colors=muted, labelsize=8)
        ax.tick_params(axis="y", labelcolor=ink)
        for sp in ["top", "right"]:
            ax.spines[sp].set_visible(False)
        for sp in ["left", "bottom"]:
            ax.spines[sp].set_color(grid)
        fig.tight_layout()
        fig.savefig(os.path.join(CONFIG["out_dir"], f"method_v_top_opportunities_{comp}.png"), dpi=160)
        plt.close(fig)


# %% [13] RUN ------------------------------------------------------------------------------
if __name__ == "__main__":
    main()
