#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
=====================================================================================
 STAPLES MARKETPLACE PoC  -  CATEGORY-LEVEL RECOMMENDATION  (Track A + C, revised)
 METHOD G : GRAPH-FIRST   (knowledge graph + structure-aware matching + graph analytics)
=====================================================================================

What this script does, in plain language
----------------------------------------
Think of every retailer's navigation as a family tree of shelves, and Staples' as a
tree with extra "also sits under" links (cross-listings). Method G puts all retailers
into ONE graph and reasons over the connections, not over word-meaning vectors.

 1. LOAD + HARMONISE   Same as Method V (shared code): standard node table, comparable
                       item counts (Staples' cross-listings followed, Office Depot's
                       cumulative counts respected).
 2. BUILD THE GRAPH    Nodes = category pages of every retailer. Edges = CHILD_OF
                       (hierarchy), CROSS_LISTED_UNDER (Staples' extra placements) and,
                       after matching, SAME_AS (competitor shelf = Staples shelf).
 3. MATCH BY WORDING + STRUCTURE ("similarity flooding")
                       Start from wording similarity (plus a small thesaurus), then let
                       evidence flow along the tree: two shelves look more alike when
                       their parents match, and two parents look more alike when their
                       children match. So "Breakroom" finds "Coffee, Water & Snacks"
                       through the coffee / beverages / candy shelves underneath, which
                       no word-similarity measure can do.
 4. CALIBRATE          Same gold set + ablation negatives as Method V.
 5. GAP                Same roll-up as Method V (ENTER / DEEPEN).
 6. TRACK C ON THE GRAPH
                         Adjacency (AAS)       - Personalised PageRank: start random walks
                                                 from Staples' shelves; how much of that
                                                 "footfall" reaches the competitor shelf?
                                                 plus: how much of its sibling shelf does
                                                 Staples already carry?
                         Cannibalization (CRS) - direct SAME_AS link to a Staples CORE shelf.
 7. DECIDE + RANK      Same decision surface and robustness check as Method V.
 8. EXPORT             Neo4j-ready CSVs + Cypher (constraints, load, example queries),
                       GraphML for Gephi, and an optional live push to Neo4j.

How to run
----------
 Google Colab
   1) Upload this file + the input files (or mount Drive, set CONFIG["data_dir"]).
   2) Uncomment and run the INSTALL lines in section [0], then:
          !python method_g_graph.py
 Local (Python 3.9+)
      pip install pandas numpy scipy scikit-learn networkx openpyxl matplotlib
      python method_g_graph.py
 Optional Neo4j push: set NEO4J_URI, NEO4J_USER, NEO4J_PASSWORD and pip install neo4j.

Inputs  (anywhere under CONFIG["data_dir"]; names are matched by pattern)
   Staples_Navigation_Tree_v2.xlsx, officedepot_tree.csv (or the Office Depot .xlsx),
   gold_matches.csv (calibration), external_signals.csv [optional demand, 0-100]

Outputs (CONFIG["out_dir"])
   method_g_results_<Competitor>.xlsx, method_g_decision_surface_<Competitor>.png,
   method_g_top_opportunities_<Competitor>.png, node_scores_method_g.csv,
   method_g_cross_competitor_gaps.xlsx (from 2 competitors),
   neo4j/ (nodes.csv, edges.csv, load_graph.cypher, example_queries.cypher),
   category_graph.graphml
"""

# %% [0] INSTALL  ---------------------------------------------------------------
# Uncomment in Colab (or run once locally):
# !pip install -q pandas numpy scipy scikit-learn networkx openpyxl matplotlib
# Optional - live push into a Neo4j database (AuraDB free tier works):
# !pip install -q neo4j

# %% [1] IMPORTS + CONFIG  --------------------------------------------------------
import glob
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
    "out_dir": "outputs_method_g",
    # --- structure-aware matching (similarity flooding) ------------------------------
    "lambda_parent": 0.40,               # weight of "our parents match" evidence (grid-searched)
    "lambda_child": 0.35,                # weight of "our children match" evidence (grid-searched)
    "n_iter": 3,                         # propagation rounds (2-3 is enough on 4-level trees)
    # --- graph analytics ----------------------------------------------------------
    "ppr_alpha": 0.85,                   # PageRank damping (0.15 = chance of jumping back to Staples)
    "aas_ppr_weight": 0.70,              # AAS = 0.7 x PageRank reach + 0.3 x sibling coverage
    "xlist_weight_cap": 15,              # Staples placements beyond this get ~zero graph weight
    "export_neo4j": True, "export_graphml": True,
    "multi_match_margin": 0.05,          # also accept other Staples shelves within this of the best
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

# ---- Thesaurus: a tiny SKOS-style altLabel layer (the "knowledge" in knowledge graph).
# Each word on the left is rewritten to the word on the right before comparing wording.
THESAURUS = {
    "couch": "sofa", "sofas": "sofa", "loveseat": "sofa", "lamps": "lighting", "lamp": "lighting",
    "janitorial": "cleaning", "breakroom": "kitchen", "pantry": "grocery", "groceries": "grocery",
    "flatware": "cutlery", "dinnerware": "tableware", "serveware": "tableware", "carpet": "rug",
    "rugs": "rug", "bins": "container", "bin": "container", "containers": "container",
    "notepads": "pad", "notepad": "pad", "pads": "pad", "tissues": "tissue", "tv": "television",
    "pc": "computer", "laptop": "computer", "chromebooks": "laptop computer", "voip": "phone",
    "teleconferencing": "conference phone", "conferencing": "conference", "wellness": "health",
    "apparel": "clothing", "workwear": "clothing", "amenities": "personal care", "stationery": "paper",
}
rng = np.random.default_rng(CONFIG["seed"])
os.makedirs(CONFIG["out_dir"], exist_ok=True)


def log(msg):
    print(msg, flush=True)


# %% [2] SHARED: FILE DISCOVERY + LOADING  (identical in method_v_vector.py) ----------
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


# %% [4] GRAPH CONSTRUCTION ----------------------------------------------------------
def synonymise(text):
    return " ".join(THESAURUS.get(t, t) for t in tokens(text))


def build_hierarchy(nodes_r, xedges=None):
    """parents / children index lists for one retailer (canonical + cross-listed)."""
    pos = {u: i for i, u in enumerate(nodes_r["uid"])}
    parents = [[] for _ in range(len(nodes_r))]
    children = [[] for _ in range(len(nodes_r))]
    for i, p in enumerate(nodes_r["parent_uid"]):
        if p in pos:
            parents[i].append(pos[p])
            children[pos[p]].append(i)
    for c, p, _w in (xedges or []):
        if c in pos and p in pos and pos[p] not in parents[pos[c]]:
            parents[pos[c]].append(pos[p])
            children[pos[p]].append(pos[c])
    return parents, children


# %% [5] STRUCTURE-AWARE MATCHING (similarity flooding) ------------------------------
def similarity_flooding(S0, par_a, ch_a, par_b, ch_b):
    """sigma(a,b) starts as wording similarity S0 and absorbs structural evidence:
         parent term : how well a's parent matches (any of) b's parents
         child term  : for each child of a, its best match among b's children, averaged
    Terms are only used where both sides have them, so roots/leaves are not penalised.
    (Melnik, Garcia-Molina & Rahm 2002, adapted to taxonomies.)"""
    from scipy.sparse import csr_matrix
    na, nb = S0.shape
    lp, lc = CONFIG["lambda_parent"], CONFIG["lambda_child"]
    has_pa = np.array([len(p) > 0 for p in par_a])
    has_pb = np.array([len(p) > 0 for p in par_b])
    has_ca = np.array([len(c) > 0 for c in ch_a])
    has_cb = np.array([len(c) > 0 for c in ch_b])
    # child-averaging operator for side A
    rows, cols, vals = [], [], []
    for a, kids in enumerate(ch_a):
        for k in kids:
            rows.append(a); cols.append(k); vals.append(1.0 / len(kids))
    Aavg = csr_matrix((vals, (rows, cols)), shape=(na, na))
    pa_first = np.array([p[0] if p else 0 for p in par_a])
    S = S0.copy()
    for _ in range(CONFIG["n_iter"]):
        # parent term: S[parent(a), best parent of b]
        Pa = S[pa_first, :]                                   # rows: a's parent
        P = np.zeros_like(S)
        for b in np.where(has_pb)[0]:
            P[:, b] = Pa[:, par_b[b]].max(axis=1)
        # child term: mean over a's children of max over b's children
        R = np.zeros_like(S)
        for b in np.where(has_cb)[0]:
            R[:, b] = S[:, ch_b[b]].max(axis=1)
        C = Aavg @ R
        wp = lp * np.outer(has_pa, has_pb)
        wc = lc * np.outer(has_ca, has_cb)
        S = (S0 + wp * P + wc * C) / (1 + wp + wc)
    return S


# %% [5b] GRAPH ANALYTICS -------------------------------------------------------------
def personalised_pagerank(G, seeds):
    import networkx as nx
    pr_seed = nx.pagerank(G, alpha=CONFIG["ppr_alpha"], personalization=seeds, weight="weight", max_iter=200)
    pr_glob = nx.pagerank(G, alpha=CONFIG["ppr_alpha"], weight="weight", max_iter=200)
    return pr_seed, pr_glob


def concepts_from_same_as(same_as_edges, all_uids):
    """Canonical category concepts = connected components of the SAME_AS graph.
    With several competitors, a concept carried by 3 of 6 competitors and not by Staples
    is a consensus whitespace - peer coverage returns naturally, no pairwise matching."""
    import networkx as nx
    H = nx.Graph()
    H.add_nodes_from(all_uids)
    H.add_edges_from((a, b) for a, b, _ in same_as_edges)
    cid = {}
    for k, comp in enumerate(nx.connected_components(H)):
        for u in comp:
            cid[u] = k
    return cid


def export_graph(G, nodes, out_dir):
    """Neo4j bulk-load CSVs + Cypher, and GraphML."""
    import networkx as nx
    nd = os.path.join(out_dir, "neo4j")
    os.makedirs(nd, exist_ok=True)
    keep = ["uid", "retailer", "name", "path_str", "level", "subtree_items", "is_leaf", "in_scope"]
    n = nodes[keep].copy()
    n["coreness"] = nodes.get("coreness", pd.Series(np.nan, index=nodes.index))
    n.columns = ["uid", "retailer", "name", "path", "level", "items", "is_leaf", "in_scope", "coreness"]
    n.to_csv(os.path.join(nd, "nodes.csv"), index=False)
    e = pd.DataFrame([(u, v, d.get("rel", "LINK"), round(float(d.get("weight", 1.0)), 4))
                      for u, v, d in G.edges(data=True)], columns=["source", "target", "rel", "weight"])
    e.to_csv(os.path.join(nd, "edges.csv"), index=False)
    cypher = """// ---- Staples marketplace PoC: category knowledge graph ----
// Copy nodes.csv + edges.csv into Neo4j's import folder, then run:
CREATE CONSTRAINT cat_uid IF NOT EXISTS FOR (c:Category) REQUIRE c.uid IS UNIQUE;

LOAD CSV WITH HEADERS FROM 'file:///nodes.csv' AS r
MERGE (c:Category {uid: r.uid})
SET c.retailer = r.retailer, c.name = r.name, c.path = r.path, c.level = toInteger(r.level),
    c.items = toFloat(r.items), c.is_leaf = (r.is_leaf = 'True'), c.coreness = toFloat(r.coreness)
MERGE (rt:Retailer {name: r.retailer})
MERGE (c)-[:SOLD_BY]->(rt);

LOAD CSV WITH HEADERS FROM 'file:///edges.csv' AS r
MATCH (a:Category {uid: r.source}), (b:Category {uid: r.target})
CALL apoc.merge.relationship(a, r.rel, {}, {weight: toFloat(r.weight)}, b, {}) YIELD rel
RETURN count(rel);
// (without APOC: run one LOAD CSV per relationship type with a WHERE r.rel = '...' filter)
"""
    queries = """// ---- Example questions the graph answers directly ----

// 1. Whitespace: competitor shelves with NO equivalent at Staples, whose sibling shelves
//    mostly DO have one (adjacent by construction - "Staples already sells the rest of this aisle")
MATCH (c:Category {retailer:'OfficeDepot'})-[:CHILD_OF]->(p)<-[:CHILD_OF]-(sib)
WHERE NOT (c)-[:SAME_AS]->(:Category {retailer:'Staples'})
WITH c, p, count(sib) AS n_sib,
     sum(CASE WHEN (sib)-[:SAME_AS]->(:Category {retailer:'Staples'}) THEN 1 ELSE 0 END) AS carried
WHERE n_sib >= 3 AND toFloat(carried)/n_sib >= 0.7
RETURN p.path AS aisle, c.name AS missing_shelf, c.items AS competitor_items, carried, n_sib
ORDER BY competitor_items DESC LIMIT 25;

// 2. Consensus whitespace across competitors (once 6 retailers are loaded):
MATCH (c:Category)-[:SAME_AS]-(d:Category)
WHERE c.retailer <> 'Staples' AND d.retailer <> 'Staples'
  AND NOT (c)-[:SAME_AS]-(:Category {retailer:'Staples'})
RETURN c.name, collect(DISTINCT d.retailer) AS also_carried_by, size(collect(DISTINCT d.retailer)) AS peers
ORDER BY peers DESC LIMIT 25;

// 3. Cannibalization check: does a candidate shelf point at a Staples CORE shelf?
MATCH (c:Category {retailer:'OfficeDepot'})-[s:SAME_AS]->(t:Category {retailer:'Staples'})
WHERE t.coreness >= 0.7
RETURN c.path, t.path, s.weight, t.coreness ORDER BY t.coreness * s.weight DESC LIMIT 25;

// 4. Adjacency with Graph Data Science: PageRank seeded on Staples shelves
CALL gds.graph.project('cat', 'Category', {CHILD_OF:{orientation:'UNDIRECTED', properties:'weight'},
     CROSS_LISTED_UNDER:{orientation:'UNDIRECTED', properties:'weight'},
     SAME_AS:{orientation:'UNDIRECTED', properties:'weight'}});
MATCH (s:Category {retailer:'Staples', is_leaf:true}) WITH collect(s) AS seeds
CALL gds.pageRank.stream('cat', {sourceNodes: seeds, relationshipWeightProperty:'weight'})
YIELD nodeId, score
WITH gds.util.asNode(nodeId) AS n, score WHERE n.retailer = 'OfficeDepot'
RETURN n.path, score ORDER BY score DESC LIMIT 25;

// 5. Staples taxonomy health: shelves that only hold items through cross-listing
MATCH (p:Category {retailer:'Staples'})<-[:CROSS_LISTED_UNDER]-(c)
WHERE NOT (p)<-[:CHILD_OF]-()
RETURN p.path, count(c) AS cross_listed_children ORDER BY cross_listed_children DESC LIMIT 25;
"""
    open(os.path.join(nd, "load_graph.cypher"), "w").write(cypher)
    open(os.path.join(nd, "example_queries.cypher"), "w").write(queries)
    if CONFIG["export_graphml"]:
        H = nx.Graph()
        for u, d in G.nodes(data=True):
            H.add_node(u, **{k: (v if isinstance(v, (int, float, str)) else str(v)) for k, v in d.items()})
        for u, v, d in G.edges(data=True):
            H.add_edge(u, v, rel=d.get("rel", ""), weight=float(d.get("weight", 1.0)))
        nx.write_graphml(H, os.path.join(out_dir, "category_graph.graphml"))
    # optional live push
    uri = os.environ.get("NEO4J_URI")
    if uri:
        try:
            from neo4j import GraphDatabase
            drv = GraphDatabase.driver(uri, auth=(os.environ.get("NEO4J_USER", "neo4j"),
                                                  os.environ.get("NEO4J_PASSWORD", "")))
            with drv.session() as ses:
                ses.run("CREATE CONSTRAINT cat_uid IF NOT EXISTS FOR (c:Category) REQUIRE c.uid IS UNIQUE")
                ses.run("UNWIND $rows AS r MERGE (c:Category {uid:r.uid}) SET c += r",
                        rows=n.where(pd.notna(n), None).to_dict("records"))
                for rel, g in e.groupby("rel"):
                    ses.run(f"UNWIND $rows AS r MATCH (a:Category {{uid:r.source}}),(b:Category {{uid:r.target}}) "
                            f"MERGE (a)-[x:{rel}]->(b) SET x.weight = r.weight", rows=g.to_dict("records"))
            log(f"      pushed {len(n)} nodes / {len(e)} edges to Neo4j at {uri}")
        except Exception as ex:
            log(f"      Neo4j push skipped ({ex})")
    log(f"      wrote Neo4j import files to {nd}/")


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
    import networkx as nx
    log("\n[1] Loading retailer trees")
    all_nodes, xedges = [], {}
    for rc in RETAILERS:
        df, f = load_retailer(rc)
        xe = load_crosslist_edges(rc, f, df) if rc.get("crosslist") else []
        xedges[rc["name"]] = xe
        if xe:
            log(f"      + {len(xe)} cross-listing placements (edges in the graph)")
        df = harmonise_counts(df, rc["count_semantics"], xe)
        df["role"] = rc["role"]
        all_nodes.append(df)
    nodes = pd.concat(all_nodes, ignore_index=True)
    totals = {r: retailer_total(nodes[nodes.retailer == r]) for r in nodes.retailer.unique()}
    for r, t in totals.items():
        log(f"      {r:<12} in-scope catalogue: {t:,.0f} items")

    log("\n[2] Building the category graph")
    nodes["name_expanded"] = nodes["path"].map(expanded_name)
    nodes["name_syn"] = nodes["name_expanded"].map(synonymise)
    lex = lexical_model(sorted(set(nodes["name_syn"])))
    focal = nodes[(nodes.retailer == FOCAL) & nodes.in_scope].reset_index()
    depth_pct = focal.groupby("level")["subtree_items"].rank(pct=True).fillna(0).values
    base = focal["path"].map(lambda p: CORE_L1_WEIGHT.get(p[0], CORE_DEFAULT)).values.astype(float)
    for pat, w in CORE_OVERRIDES:
        base[focal["path_str"].str.contains(pat, regex=True).values] = w
    focal["coreness"] = base * (0.6 + 0.4 * depth_pct)
    nodes.loc[focal["index"].values, "coreness"] = focal["coreness"].values
    par_b, ch_b = build_hierarchy(focal, xedges.get(FOCAL))
    FL = lex.transform(focal["name_syn"])

    G = nx.Graph()
    for _, r in nodes[nodes.in_scope].iterrows():
        G.add_node(r["uid"], retailer=r["retailer"], name=r["name"], level=int(r["level"]),
                   items=float(r["subtree_items"] or 0))
    for _, r in nodes[nodes.in_scope & nodes.parent_uid.notna()].iterrows():
        if r["parent_uid"] in G:
            G.add_edge(r["uid"], r["parent_uid"], rel="CHILD_OF", weight=1.0)
    npl = defaultdict(int)
    for c, p, w in xedges.get(FOCAL, []):
        npl[c] += 1
    for c, p, w in xedges.get(FOCAL, []):
        if c in G and p in G and not G.has_edge(c, p):
            G.add_edge(c, p, rel="CROSS_LISTED_UNDER",
                       weight=(1.0 / npl[c]) if npl[c] <= CONFIG["xlist_weight_cap"] else 0.001)
    log(f"      graph: {G.number_of_nodes():,} nodes, {G.number_of_edges():,} edges "
        f"(hierarchy + Staples cross-listings)")

    results, same_as_all = [], []
    for rc in RETAILERS:
        comp = rc["name"]
        if comp == FOCAL:
            continue
        log(f"\n[3] Structure-aware matching: {comp} -> {FOCAL}")
        cn = nodes[(nodes.retailer == comp) & nodes.in_scope].reset_index(drop=False)
        par_a, ch_a = build_hierarchy(cn, xedges.get(comp))
        CL = lex.transform(cn["name_syn"])
        S0 = (CL @ FL.T).toarray().astype(np.float32)
        S = similarity_flooding(S0, par_a, ch_a, par_b, ch_b)
        order = np.argsort(-S, axis=1)[:, :50]
        Ssorted = np.take_along_axis(S, order, axis=1)
        cn["best_score"] = Ssorted[:, 0]
        cn["best_lex"] = S0[np.arange(len(cn)), order[:, 0]]
        cn["best_staples"] = focal["path_str"].values[order[:, 0]]
        cn["best_staples_uid"] = focal["uid"].values[order[:, 0]]
        cn["alt_staples"] = [" | ".join(focal["path_str"].values[order[i, 1:3]]) for i in range(len(cn))]
        cn["best_sem"] = np.nan

        # ---- calibration (same gold set + ablation negatives as Method V)
        gold_f = find_file(["gold_matches.csv"], CONFIG["data_dir"])
        gold = pd.read_csv(gold_f) if gold_f else pd.DataFrame(
            columns=["competitor", "comp_path", "staples_carries", "gold_staples_path", "source"])
        gold = gold[gold["competitor"] == comp]
        if len(gold) >= MIN_GOLD_ROWS:
            pos = gold[(gold.staples_carries == 1) & gold.gold_staples_path.notna()]
            row_of = dict(zip(cn["path_str"], range(len(cn))))
            abl = {}
            for _, r in pos.iterrows():
                i = row_of.get(r["comp_path"])
                if i is not None:
                    abl[r["comp_path"]] = float(S[i][ablation_mask(focal, r["gold_staples_path"], r["comp_path"])].max())
            cal_df = build_calibration_set(gold, dict(zip(cn["path_str"], cn["best_score"])), abl)
            tau, cal = calibrate_threshold(cal_df)
            s_enter, sinfo = gap_screen_threshold(cal_df, gold)
            cal.update(sinfo)
            cal["neg_median_score"] = float(cal_df.loc[cal_df.y == 0, "score"].median())
            cal["top1_accuracy"], cal["n_top1"] = top1_accuracy(gold, dict(zip(cn["path_str"], cn["best_staples"])))
            # wording-only baseline, to show what structure adds
            base_best = S0.max(axis=1)
            cal0 = build_calibration_set(gold, dict(zip(cn["path_str"], base_best)),
                                         {k: float(S0[row_of[k]][ablation_mask(focal, g, k)].max())
                                          for k, g in zip(pos["comp_path"], pos["gold_staples_path"]) if k in row_of})
            _, m0 = calibrate_threshold(cal0)
            cal["auc_wording_only"] = m0["auc_augmented"]
            cal["top1_wording_only"], _ = top1_accuracy(gold, dict(zip(cn["path_str"],
                                                                       focal["path_str"].values[S0.argmax(axis=1)])))
            gscored = cal_df
            log(f"      confident-pairing tau = {tau:.3f} | AUC {cal['auc_augmented']:.3f} (wording only "
                f"{cal['auc_wording_only']:.3f}) | top-1 shelf accuracy {cal['top1_accuracy']:.1%} (wording only "
                f"{cal['top1_wording_only']:.1%})")
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
        cn["gap_flag"] = cn["best_score"] < s_enter
        cn["margin_to_screen"] = cn["best_score"] - s_enter
        m = CONFIG["multi_match_margin"]
        cn["match_set"] = [[focal["uid"].values[order[i, j]] for j in range(order.shape[1])
                            if Ssorted[i, j] >= tau and Ssorted[i, j] >= Ssorted[i, 0] - m] for i in range(len(cn))]
        cn["n_shelves_matched"] = cn["match_set"].map(len)
        if "gap_screen_recall_on_real_gaps" in cal:
            log(f"      gap screen: score < {s_enter:.3f} | catches {cal['gap_screen_recall_on_real_gaps']:.0%} of real "
                f"gold gaps at a {CONFIG['gap_screen_false_alarm']:.0%} false-alarm budget | "
                f"{cal['base_rate_carried_random_sample']:.1%} of randomly drawn {comp} shelves exist at Staples")
        # SAME_AS edges into the graph
        for u, ms, sc in zip(cn["uid"], cn["match_set"], cn["best_score"]):
            for v in ms:
                G.add_edge(u, v, rel="SAME_AS", weight=float(sc))
                same_as_all.append((u, v, float(sc)))
        # CRS: direct SAME_AS-strength link to a Staples CORE shelf (0 = typical non-match)
        lo = cal["neg_median_score"]
        core_mat = focal["coreness"].values[order]
        f = np.clip((Ssorted - lo) / max(tau - lo, 1e-6), 0, 1)
        cn["crs_raw"] = (f * core_mat).max(axis=1)
        cn["crs_driver"] = [focal["path_str"].values[order[i, np.argmax(f[i] * core_mat[i])]] for i in range(len(cn))]
        cn["competitor"] = comp
        results.append([comp, cn, cal, gscored, tau])

    log("\n[4] Graph analytics: Personalised PageRank from Staples' shelves")
    st_leaves = focal[focal["count_raw"].notna()]
    seeds = {u: float(np.log1p(c)) for u, c in zip(st_leaves["uid"], st_leaves["count_raw"]) if u in G and c > 0}
    pr_seed, pr_glob = personalised_pagerank(G, seeds)
    concept = concepts_from_same_as(same_as_all, list(G.nodes))
    log(f"      {len(set(concept.values())):,} canonical concepts (SAME_AS components)")

    node_scores, unit_sets = [], []
    for k_res, (comp, cn, cal, gscored, tau) in enumerate(results):
        eps = 1e-12
        cn["A_raw"] = [np.log((pr_seed.get(u, 0) + eps) / (pr_glob.get(u, 0) + eps)) for u in cn["uid"]]
        cn["concept_id"] = cn["uid"].map(concept)
        cn = rollup_and_score(cn, focal, totals[comp], totals[FOCAL], tau)
        # sibling coverage: share of the parent aisle's items that Staples carries (excl. self)
        cov_items = dict(zip(cn["uid"], cn["covered_items"]))
        tot_items = dict(zip(cn["uid"], cn["leaf_items"]))
        sib = []
        for u, p in zip(cn["uid"], cn["parent_uid"]):
            if p in tot_items and tot_items[p] - tot_items[u] > 0:
                sib.append((cov_items[p] - cov_items[u]) / (tot_items[p] - tot_items[u]))
            else:
                sib.append(np.nan)
        cn["sibling_coverage"] = np.clip(sib, 0, 1)
        w = CONFIG["aas_ppr_weight"]
        cn["AAS_ppr"] = cn["AAS"]
        cn["AAS"] = np.where(cn["sibling_coverage"].notna(),
                             w * cn["AAS_ppr"] + (1 - w) * 100 * cn["sibling_coverage"].fillna(0), cn["AAS_ppr"])
        results[k_res][1] = cn
        gaps = cn["gap_type"] != "PARITY"
        cn.loc[gaps, "O_node"] = opportunity(cn[gaps], CONFIG["rank_weights"], CONFIG["gamma"], False)
        units = select_units(cn)
        units = decide_and_rank(units, comp)
        sens = sensitivity(units)
        units = units.merge(sens, on="uid", how="left")
        unit_sets.append((comp, units, cal, gscored, tau))
        staples_view = staples_taxonomy_view(nodes, focal, xedges.get(FOCAL, []), cn)
        write_outputs(comp, cn, units, cal, gscored, staples_view, "lexical + thesaurus + structure", "networkx", totals)
        charts(units, comp)
        node_scores.append(cn[["competitor", "uid", "path_str", "level", "is_leaf", "comp_items", "best_staples",
                               "best_score", "matched", "coverage", "DG", "WM", "gap_type", "AAS", "CRS",
                               "sibling_coverage", "concept_id", "O_node"]]
                           .merge(units[["uid", "label", "O"]], on="uid", how="left"))
        log("\n[7] Top recommendations (approved labels only)")
        show = units[units.approved].sort_values("O", ascending=False).head(CONFIG["top_n"])
        log(show[["path_str", "label", "comp_items", "staples_items", "coverage", "DG", "AAS", "CRS", "O",
                  "p_top_n"]].round(2).to_string(index=False))
    pd.concat(node_scores).to_csv(os.path.join(CONFIG["out_dir"], "node_scores_method_g.csv"), index=False)

    # cross-competitor gap concepts (needs >= 2 competitors): wording + thesaurus
    def sim_g(df):
        L = lex.transform([r["name_syn"] for r in df["row"]])
        return (L @ L.T).toarray()

    xw = cross_competitor_whitespace(unit_sets, sim_g, {r[0]: r[4] for r in unit_sets})
    if len(xw):
        xw.to_excel(os.path.join(CONFIG["out_dir"], "method_g_cross_competitor_gaps.xlsx"), index=False)
        log(f"\n[6b] Cross-competitor gap concepts: {len(xw)} concepts, "
            f"{(xw.n_competitors >= 2).sum()} carried by 2+ competitors")
    if CONFIG["export_neo4j"]:
        log("\n[8] Exporting the graph")
        export_graph(G, nodes[nodes.in_scope], CONFIG["out_dir"])
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


# %% [11] STAPLES TAXONOMY VIEW (graph-only insight) -----------------------------------
def staples_taxonomy_view(nodes, focal, xedges, cn):
    """Where Staples' own navigation graph helps or hides assortment: shelves whose items
    arrive only through cross-listing, hub shelves with many placements, and Staples L2s
    with no Office Depot counterpart (where Staples is ahead)."""
    placements = defaultdict(int)
    xparents = defaultdict(int)
    for c, p, _ in xedges:
        placements[c] += 1
        xparents[p] += 1
    f = focal[["uid", "path_str", "level", "is_leaf", "count_raw", "subtree_items", "coreness"]].copy()
    f["extra_placements"] = f["uid"].map(placements).fillna(0).astype(int)
    f["cross_listed_children"] = f["uid"].map(xparents).fillna(0).astype(int)
    f["items_only_via_crosslisting"] = f["is_leaf"] & f["count_raw"].isna() & (f["subtree_items"] > 0)
    matched_st = set(u for ms in cn["match_set"] for u in ms)
    f["has_competitor_counterpart"] = f["uid"].isin(matched_st)
    return f.drop(columns=["uid"]).sort_values(["items_only_via_crosslisting", "subtree_items"], ascending=False)


# %% [12] OUTPUTS ------------------------------------------------------------------------
def write_outputs(comp, cn, units, cal, gscored, staples_ahead, enc_kind, backend, totals):
    fp = os.path.join(CONFIG["out_dir"], f"method_g_results_{comp}.xlsx")
    readme = pd.DataFrame({
        "item": ["Method", "Competitor", "Matcher", "Graph engine", "Match threshold tau",
                 "Staples in-scope items", f"{comp} in-scope items", "Units scored", "Approved (CURATE)",
                 "How to read"],
        "value": ["G - graph-first (structure-aware matching, PageRank adjacency)", comp, enc_kind, backend,
                  round(cal.get("tau", float("nan")), 4), f"{totals[FOCAL]:,.0f}", f"{totals[comp]:,.0f}",
                  len(units), int(units["approved"].sum()),
                  "Recommendations = approved units ranked by O. AAS/CRS are 0-100. DG = log share gap "
                  "(>0: competitor devotes a bigger share of its catalogue). p_top_n = share of 500 "
                  "re-weighted runs in which the unit stays in the top-N."]})
    cols = ["rank_approved", "path_str", "level", "gap_type", "label", "action", "comp_items", "staples_items",
            "comp_per10k", "staples_per10k", "coverage", "DG", "WM", "AAS", "CRS", "O", "p_approved",
            "p_top_n", "top_uncovered", "matched_to", "crs_driver", "sibling_coverage", "concept_id"]
    rec = units[units["approved"]].sort_values("O", ascending=False)[cols]
    allu = units.sort_values(["label", "O"], ascending=[True, False])[cols]
    nm = cn[["path_str", "level", "is_leaf", "comp_items", "best_staples", "best_score", "best_lex",
             "alt_staples", "concept_id", "sibling_coverage", "matched", "gap_flag", "margin_to_screen", "n_shelves_matched", "coverage", "gap_type",
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
            staples_ahead.round(3).to_excel(xw, sheet_name="Staples_Taxonomy_Graph", index=False)
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
    ax.set_title(f"Method G decision surface - gaps vs {comp} (bubble size = competitor items)", loc="left",
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
    fig.savefig(os.path.join(CONFIG["out_dir"], f"method_g_decision_surface_{comp}.png"), dpi=160)
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
        fig.savefig(os.path.join(CONFIG["out_dir"], f"method_g_top_opportunities_{comp}.png"), dpi=160)
        plt.close(fig)


# %% [13] RUN ------------------------------------------------------------------------------
if __name__ == "__main__":
    main()
