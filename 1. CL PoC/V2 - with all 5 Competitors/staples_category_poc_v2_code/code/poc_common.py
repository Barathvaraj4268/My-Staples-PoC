#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
=====================================================================================
 STAPLES MARKETPLACE PoC  -  CATEGORY-LEVEL RECOMMENDATION  (v2, six retailers)
 poc_common.py : SHARED LIBRARY used by all three scripts
=====================================================================================

   method_v_vector.py   Method V - embeddings + Qdrant vector DB  -> node_scores_V.csv
   method_g_graph.py    Method G - category knowledge graph        -> node_scores_G.csv
   run_framework.py     Ensemble + Track C framework + every chart and table used in the
                        documents and the deck                     -> outputs/final/...

Everything a merchant might want to change lives at the top of this file:
   CONFIG            thresholds, weights, encoder, paths
   RETAILERS         one adapter per navigation tree (file, columns, counts, scope rules)
   CORE_*            1P "coreness" assumptions (replace with Staples sales bands later)

What changed from v1 (Office Depot only) - see the Methodology document, section 3:
   * counts are OPTIONAL per retailer (full / partial / none) - breadth is measured for all
   * embeddings use whatever path a node has (self + decaying ancestors), not fixed L1/L2/L3
   * navigation headings ("Shop By Category", "Featured") are skipped as context
   * department scope is one explicit, auditable universe applied to every retailer
   * calibration is per competitor (Office Depot gold + labelled samples for the others)

Install (Colab or local, Python 3.9+):
   pip install pandas numpy scipy scikit-learn openpyxl matplotlib networkx qdrant-client spacy
   python -m spacy download en_core_web_lg          # default encoder (offline once downloaded)
   pip install sentence-transformers                # optional: encoder="st" (needs HuggingFace)
   pip install wordllama                            # optional: encoder="wordllama" (bundled weights)
"""
import glob
import json
import math
import os
import re
import sys
import warnings
from collections import defaultdict

import numpy as np
import pandas as pd

warnings.filterwarnings("ignore")
pd.set_option("display.width", 220)
pd.set_option("display.max_columns", 40)

HERE = os.path.dirname(os.path.abspath(__file__))

# %% ------------------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------------------
CONFIG = {
    # --- paths (Colab: upload the six .xlsx files + gold_matches_v2.csv to /content) ------
    "data_dir": os.environ.get("POC_DATA_DIR", os.path.join(HERE, "..", "data")),
    "out_dir": os.environ.get("POC_OUT_DIR", os.path.join(HERE, "..", "outputs")),
    "gold_file": "gold_matches_v2.csv",
    "seed": 42,
    # --- tree handling ------------------------------------------------------------------
    "max_depth": 4,                 # Staples is 4 deep; deeper competitor levels fold into L4
    "max_unit_depth": 3,            # recommendations at L1-L3 (category level)
    # --- encoder (Method V) -------------------------------------------------------------
    # "spacy" reproduces the documents. "auto" runs a bake-off of every encoder that loads and
    # keeps the one with the best top-1 shelf accuracy on the labelled set.
    "encoder": "spacy",
    "encoder_candidates": ["st", "spacy", "wordllama", "tfidf"],
    "spacy_model": "en_core_web_lg",
    "st_model": "BAAI/bge-small-en-v1.5",
    # node vector = own label + ancestors (weight decays geometrically going up the path).
    # Asymmetric on purpose (context-weight grid, context_weight_grid.csv): competitor labels
    # are often generic ("Covers", "Bedroom") and need their context; Staples' upper levels
    # are noisy (patio furniture sits under Gift Shop > Professional Gifts > Compasses), so
    # Staples shelves lean on their own label.
    "w_self": 0.40, "ctx_decay": 0.35,              # competitor shelves
    "w_self_focal": 0.70, "ctx_decay_focal": 0.50,  # Staples shelves
    # --- retrieval + rerank (Method V) ----------------------------------------------------
    "k_retrieve": 50, "k_adjacency": 10,
    "w_sem": 0.65, "w_lex": 0.35,
    "multi_match_margin": 0.08,
    # --- graph (Method G) -------------------------------------------------------------------
    "lambda_parent": 0.40, "lambda_child": 0.35, "n_iter": 3,
    "ppr_alpha": 0.85, "aas_ppr_weight": 0.7, "xlist_weight_cap": 15,
    # --- calibration ------------------------------------------------------------------------
    "min_gold_rows": 40,
    "false_gap_cost": 1.0,
    "gap_screen_false_alarm": 0.10,
    # --- gap logic ----------------------------------------------------------------------------
    "enter_coverage_max": 0.20,     # < 20% of a node's leaves (items if counted) carried -> ENTER
    "deepen_min": math.log(2),      # competitor >= 2x Staples' relative depth/breadth -> DEEPEN
    "deepen_min_carriers": 2,       # ... shown by at least 2 competitors
    # --- Track C decision surface (0-100) --------------------------------------------------------
    "aas_hi": 60, "aas_lo": 40, "crs_hi": 60, "crs_lo": 40,
    # --- opportunity score -------------------------------------------------------------------------
    "rank_weights": {"pc": 0.35, "gs": 0.35, "aas": 0.30},   # peer consensus, gap size, adjacency
    "demand_weight": 0.25,          # used only if external_signals.csv is supplied
    "gamma": 1.0,                   # cannibalisation penalty exponent: O x (1 - CRS/100)^gamma
    "top_n": 20,
    "n_sensitivity": 500, "dirichlet_conc": 20, "threshold_jitter": 5.0,
}

FOCAL = "Staples"

# ---------------------------------------------------------------------------------------
# DEPARTMENT UNIVERSE - one rule set for every retailer (auditable in 00_data_audit.xlsx).
# In: anything a workplace / home-office / breakroom / facilities / lifestyle-workspace
#     customer could plausibly buy from Staples.  Out: media, apparel, beauty, auto, baby
#     consumables, toys, grocery, pets, instruments, collectibles, gift cards, services,
#     brand / character / merchandising hubs.  (Pet was a HOLD in v1; it is out of the
#     universe for consistency with the team's Amazon scoping.)
# ---------------------------------------------------------------------------------------
UNIVERSE_EXCLUDE_L1 = (
    r"^(Books|Books & Magazines|Kindle|Audible|Magazine|Music|Digital Music|CDs|Movies|Amazon Instant Video|"
    r"Video Games|Apps & Games|Software$|Clothing|Jewelry|Beauty|Premium Beauty|Auto|Automotive|Vehicles|"
    r"Baby$|Toys|Food$|Grocery|Amazon Fresh|Pets?$|Pet Supplies|Musical Instruments|Collectibles|"
    r"Gift Cards|Services|Photo Center|Shop by (Brand|Movie|TV Show|Video Game)|Character Shop|Feature$|"
    r"Gifts & Registry|Amazon Devices|Handmade)"
)
# merchandising / non-category pages anywhere in a path (drop the node and its subtree)
MERCH_NODE = (
    r"^(New|New Arrivals|Sale|Clearance|Deals|Best Sellers|Trending|Top Brands|Pierce & Ward.*|.*Inspiration|"
    r"Wood Swatches|Gift Guides?|Shop by Brand|Brands)$"
    r"|^New (?!Year)|^Sale |^Explore |Swatches|^In Stock|Build Your Own|in America$|by the Yard$"
    r"|Touch Up|Team Shop$|Fan Shop$|[Bb]y Brand"
)
# facets (colour, size, material, style, room) are filters, not categories
FACET_NODE = (
    r"(By|by) (Color|Colour|Fabric|Material|Style|Opacity|Wood Finish|Pattern|Room|Size|Shape|Price)"
    r"|Popular .*Sizes|Print & Pattern Shop|Shop By Color|^\d+'? ?x ?\d+'? "
)
# navigation headings: kept in the tree, skipped as semantic context, never a unit
HEADING_NODE = r"^(Shop By Category|Shop by Category|Shop By Type|Rugs By Type|More Rooms|Featured|Office Collections|" \
               r"Outdoor Collections|Shop All|Home Accessories|Categories|Departments|Shop by Department)$|^\(unnamed ID"
AGGREGATE_NODE = r"^(All|Shop All) "   # "All Office", "All Rugs" duplicate their parent

# ---------------------------------------------------------------------------------------
# RETAILER ADAPTERS - add a competitor by adding one entry
#   count_mode : "terminal_only" leaf pages carry counts, parents do not (Staples)
#                "cumulative"    parent pages carry totals incl. children (OD, West Elm)
#                "partial"       cumulative, but only on some levels (Walmart L1-L2)
#                "none"          no usable counts (Wayfair, Amazon) - breadth only
#   role       : mirror (B2B peer) | lifestyle (design-led) | scale (everything store)
# ---------------------------------------------------------------------------------------
RETAILERS = [
    {"name": "Staples", "role": "focal",
     "file_glob": ["*Staples_Navigation_Tree*.xlsx"], "sheet": "Navigation Tree",
     "level_cols": ["L1", "L2", "L3", "L4", "L5"], "count_col": "Count", "count_mode": "terminal_only",
     "id_col": "Category ID", "url_col": "URL",
     "crosslist": {"sheet": "Cross-Listings", "header_row": 3, "name_col": "Category",
                   "parents_col": "All parents"},
     "scope_exclude": [r"^Gift Cards", r"^Tech Services", r"^Warranties", r"Furniture Assembly$",
                       r"^Expanded Assortment > (Books|Clothing)"]},
    {"name": "OfficeDepot", "role": "mirror",
     "file_glob": ["*OfficeDepot_Navigation_Tree*.xlsx", "*officedepot_tree*.csv"], "sheet": "Category Tree",
     "level_cols": ["L1", "L2", "L3", "L4", "L5"], "count_col": "Count", "count_mode": "cumulative",
     "id_col": "Node ID", "url_col": "URL",
     # services / programme pages / custom print: Staples runs these outside its tree
     "scope_exclude": [r"^Ink & Toner", r"^GreenerOffice", r"^Services", r"^Security Solutions",
                       r"^Print & Copy", r"Warranties & Services", r"Protection Plans$", r"^Pet Supplies"]},
    {"name": "WestElm", "role": "lifestyle",
     "file_glob": ["*WestElm_Navigation_Tree*.xlsx"], "sheet": "Category Tree",
     "level_cols": ["L1", "L2", "L3", "L4", "L5"], "count_col": "Count", "count_mode": "cumulative",
     "id_col": "Node ID", "url_col": "URL",
     # named product collections (Hughes, Marlowe, "Sofa Collections") are ranges, not categories
     "scope_exclude": [r"West Elm Business to Business", r"West Elm Office at Work", r"Explore West Elm",
                       r"Collections?( >|$)"]},
    {"name": "Wayfair", "role": "lifestyle",
     "file_glob": ["*Wayfair_Navigation_Tree*.xlsx"], "sheet": "Navigation Tree",
     "level_cols": ["L1", "L2", "L3", "L4"], "count_col": None, "count_mode": "none",
     "id_col": "Category ID", "url_col": "URL",
     "row_filter": lambda d: d["Primary placement"].astype(str).str.upper().eq("Y"),
     "scope_exclude": [r"^Pet"]},
    {"name": "Amazon", "role": "scale",
     "file_glob": ["*Amazon_Navigation_Tree*.xlsx"], "sheet": "Navigation Tree",
     "level_cols": ["L1", "L2", "L3"], "count_col": None, "count_mode": "none",
     "id_col": "Category ID", "url_col": None,
     # team's PoC-relevant flag, plus phone accessories (Staples sells them)
     "row_filter": lambda d: d["PoC-relevant L1"].astype(str).eq("Yes") | d["L1"].eq("Cell Phones & Accessories"),
     "scope_exclude": []},
    {"name": "Walmart", "role": "scale",
     "file_glob": ["*Walmart_Navigation_Tree*.xlsx"], "sheet": "Navigation Tree",
     "level_cols": ["L1", "L2", "L3", "L4", "L5"], "count_col": "Count", "count_mode": "partial",
     "count_cap": 900000,          # Walmart shows "900,000+" - treated as censored (unknown)
     "id_col": "Category ID", "url_col": "URL",
     "row_filter": lambda d: d["Node Type"].eq("Category") & ~d["Name Source"].astype(str).str.startswith("No name"),
     "scope_exclude": []},
]
ROLE_OF = {r["name"]: r["role"] for r in RETAILERS}
COMPETITORS = [r["name"] for r in RETAILERS if r["name"] != FOCAL]

# ---------------------------------------------------------------------------------------
# 1P CORENESS - proxy for "how much Staples 1P revenue sits on this shelf" (0..1).
# Editable assumptions; replace with Staples category sales / margin bands when available.
# ---------------------------------------------------------------------------------------
CORE_L1_WEIGHT = {
    "Paper": 1.0, "Office Supplies": 1.0, "Shipping, Packing & Mailing Supplies": 1.0,
    "Cleaning Supplies": 0.9, "Coffee, Water & Snacks": 0.9, "Printers & Scanners": 0.9,
    "Shredders, Projectors & Office Machines": 0.8, "Computers & Accessories": 0.8,
    "Facilities": 0.7, "Safety Supplies": 0.7, "School Supplies": 0.7, "Batteries & Power": 0.7,
    "Furniture": 0.6, "Hard Drives & Data Storage": 0.6, "Security, Banking & Cash": 0.6,
    "Networking & WiFi": 0.5, "Healthcare Supplies": 0.5, "Retail Store Supplies": 0.5,
    "Party Supplies": 0.4, "Phones, Cameras & Electronics": 0.4, "Tablets & iPads": 0.4,
    "Bags, Backpacks & Luggage": 0.6, "Computer Software": 0.5, "Audio": 0.4,
    "Restaurant & Foodservice Supplies": 0.5, "Laboratory & Scientific Supplies": 0.5,
    "Tools, Parts & Supplies": 0.5, "Appliances & Kitchenware": 0.5, "Workwear": 0.4,
    "Gift Shop": 0.3, "Arts & Crafts": 0.3, "TV & Streaming Media": 0.3, "Smart Home & Security": 0.3,
    "Gaming": 0.3, "Wearable Technology": 0.3, "Education": 0.5,
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

GENERIC_NAMES = {"accessories", "other", "parts", "supplies", "storage", "more", "misc", "equipment",
                 "systems", "solutions", "kits", "refills", "tools", "products", "components",
                 "essentials", "basics", "appliances", "sets", "furniture", "decor", "hardware"}
STOP = {"and", "the", "of", "for", "with", "a", "an", "in", "to", "by", "all", "more", "other", "s", "shop"}
NOISE_WORDS = r"\b(shop( all| by)?|explore|west elm|pro\)|\(pro|professional)\b"

# Thesaurus (Method G): a small SKOS-style altLabel layer
THESAURUS = {
    "couch": "sofa", "sofas": "sofa", "loveseat": "sofa", "lamps": "lighting", "lamp": "lighting",
    "janitorial": "cleaning", "breakroom": "kitchen", "pantry": "grocery", "groceries": "grocery",
    "flatware": "cutlery", "dinnerware": "tableware", "serveware": "tableware", "carpet": "rug",
    "rugs": "rug", "bins": "container", "bin": "container", "containers": "container",
    "notepads": "pad", "notepad": "pad", "pads": "pad", "tissues": "tissue", "tv": "television",
    "pc": "computer", "laptop": "computer", "chromebooks": "laptop computer", "voip": "phone",
    "teleconferencing": "conference phone", "conferencing": "conference", "wellness": "health",
    "apparel": "clothing", "workwear": "clothing", "amenities": "personal care", "stationery": "paper",
    "drapes": "curtain", "curtains": "curtain", "sconces": "wall lighting", "chandeliers": "ceiling lighting",
    "pendants": "ceiling lighting", "barware": "drinkware", "throws": "blanket", "duvet": "bedding",
    "planters": "planter", "pots": "planter", "wearables": "wearable", "trackers": "tracker",
}

LABELS = ["CURATE", "VERTICAL EXTENSION", "REVIEW", "1P-CORE GAP", "OFF-BRAND", "VERIFY"]
# Zone colours: validated categorical steps (CVD-checked; OFF-BRAND is a deliberate neutral).
# Charts always add a second cue (shaded zones + direct labels, hollow VERIFY markers).
LABEL_COLOR = {"CURATE": "#008300", "VERTICAL EXTENSION": "#2a78d6", "REVIEW": "#eda100",
               "1P-CORE GAP": "#e34948", "OFF-BRAND": "#8c8c8c", "VERIFY": "#4a3aa7"}
ACTION = {
    "CURATE": "Open to marketplace sellers now (curated, design-led assortment)",
    "VERTICAL EXTENSION": "Phase 2 - enter through a vertical Staples already serves",
    "REVIEW": "Merchant decision - sits next to a 1P line",
    "1P-CORE GAP": "Fix in 1P merchandising - a marketplace here would cannibalise core",
    "OFF-BRAND": "Pass - real gap, wrong store for Staples",
    "VERIFY": "Check first - Staples may already carry it under another name (findability)",
}
RETAILER_COLOR = {"Staples": "#52514e", "OfficeDepot": "#2a78d6", "WestElm": "#eb6834", "Wayfair": "#1baf7a",
                  "Amazon": "#eda100", "Walmart": "#e87ba4"}
INK = {"primary": "#0b0b0b", "secondary": "#52514e", "muted": "#8a8984", "grid": "#e6e5e0", "surface": "#fcfcfb"}
STATUS_COLOR = {"carried": "#2a78d6", "likely": "#9dbfe9", "disputed": "#c9c8c3", "gap_soft": "#f0a3a2",
                "gap_hard": "#e34948"}

rng = np.random.default_rng(CONFIG["seed"])
RAW_ROWS = {}        # rows in each source file (filled by load_retailer; reported in the data audit)


def log(msg=""):
    print(msg, flush=True)


def out_path(*parts):
    p = os.path.join(CONFIG["out_dir"], *parts)
    os.makedirs(os.path.dirname(p), exist_ok=True)
    return p


# %% ------------------------------------------------------------------------------------
# FILES
# ---------------------------------------------------------------------------------------
def find_file(patterns, data_dir=None):
    """data_dir, one folder below it, the script folder and /content (Colab). Not recursive."""
    data_dir = data_dir or CONFIG["data_dir"]
    for root in [data_dir, os.path.join(data_dir, "*"), HERE, "/content"]:
        for pat in patterns:
            hits = sorted(glob.glob(os.path.join(root, pat)))
            if hits:
                return hits[0]
    return None


def read_table(path, sheet=None, header=0):
    if path.lower().endswith(".csv"):
        return pd.read_csv(path)
    return pd.read_excel(path, sheet_name=sheet, header=header)


# %% ------------------------------------------------------------------------------------
# TEXT
# ---------------------------------------------------------------------------------------
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


def clean_name(name):
    """Strip navigation noise ('Shop', 'Explore', '(Pro)') from a label."""
    s = re.sub(NOISE_WORDS, " ", str(name), flags=re.I)
    s = re.sub(r"\s+", " ", s).strip(" -&,")
    return s or str(name)


def is_generic(name):
    toks = [stem(t) for t in tokens(name)]
    gen = {stem(g) for g in GENERIC_NAMES}
    return (not toks) or all(t in gen for t in toks)


def synonymise(text):
    return " ".join(THESAURUS.get(t, t) for t in tokens(text))


def path_str(parts):
    return " > ".join(parts)


# %% ------------------------------------------------------------------------------------
# LOADING: one standard node table for every retailer
# ---------------------------------------------------------------------------------------
def load_retailer(rc):
    f = find_file(rc["file_glob"])
    if f is None:
        raise FileNotFoundError(f"No file for {rc['name']} matching {rc['file_glob']} in {CONFIG['data_dir']}")
    raw = read_table(f, rc.get("sheet"))
    n_raw = len(raw)
    if rc.get("row_filter") is not None:
        raw = raw[rc["row_filter"](raw)]
    lv = [c for c in rc["level_cols"] if c in raw.columns]
    raw = raw.dropna(subset=[lv[0]])
    cnt = rc.get("count_col")
    rows = []
    for rd in raw.to_dict("records"):
        parts = []
        for c in lv:
            v = rd.get(c)
            if v is not None and not (isinstance(v, float) and np.isnan(v)) and str(v).strip():
                parts.append(re.sub(r"\s+", " ", str(v)).strip())
        if not parts:
            continue
        cval = pd.to_numeric(rd.get(cnt), errors="coerce") if cnt and cnt in rd else np.nan
        rows.append({"path": tuple(parts), "count_raw": cval,
                     "native_id": str(rd.get(rc["id_col"], "")) if rc.get("id_col") in rd else "",
                     "url": rd.get(rc["url_col"], "") if rc.get("url_col") in rd else ""})
    if not rows:
        raise RuntimeError(f"{rc['name']}: no rows parsed")
    df = pd.DataFrame(rows).drop_duplicates("path").reset_index(drop=True)

    # implied parents (a breadcrumb with no row of its own)
    have = set(df["path"])
    extra = []
    for p in list(have):
        for i in range(1, len(p)):
            if p[:i] not in have:
                have.add(p[:i])
                extra.append({"path": p[:i], "count_raw": np.nan, "native_id": "", "url": ""})
    if extra:
        df = pd.concat([df, pd.DataFrame(extra)], ignore_index=True)

    df["retailer"] = rc["name"]
    df["role"] = rc["role"]
    df["depth"] = df["path"].map(len)
    df["name"] = df["path"].map(lambda p: p[-1])
    df["path_str"] = df["path"].map(path_str)

    # ---- node kinds and scope ------------------------------------------------------------
    re_merch, re_facet = re.compile(MERCH_NODE), re.compile(FACET_NODE)
    re_head, re_aggr = re.compile(HEADING_NODE), re.compile(AGGREGATE_NODE)
    re_univ = re.compile(UNIVERSE_EXCLUDE_L1)
    re_scope = [re.compile(x) for x in rc.get("scope_exclude", [])]

    parents_with_kids = {p[:-1] for p in df["path"] if len(p) > 1}

    def drop_reason(p):
        if rc["name"] != FOCAL and re_univ.search(p[0]):
            return "universe: department out of scope"
        ps = path_str(p)
        if any(x.search(ps) for x in re_scope):
            return "retailer scope rule"
        for q in p:
            if re_merch.search(q):
                return "merchandising page (new / sale / collection / sub-brand)"
            if re_facet.search(q):
                return "facet (colour / size / style / room filter)"
        if re_aggr.search(p[-1]) and p not in parents_with_kids:
            return "aggregate page duplicating its parent"
        if len(p) > CONFIG["max_depth"]:
            return f"deeper than L{CONFIG['max_depth']} (folded into L{CONFIG['max_depth']})"
        return ""

    df["drop_reason"] = df["path"].map(drop_reason)
    df["in_scope"] = df["drop_reason"].eq("")
    df["is_heading"] = df["name"].map(lambda n: bool(re_head.search(n) or re_aggr.search(n)))
    df["name_clean"] = df["name"].map(clean_name)

    df["uid"] = rc["name"] + "::" + df["path_str"]
    df["parent_uid"] = df["path"].map(lambda p: rc["name"] + "::" + path_str(p[:-1]) if len(p) > 1 else None)
    ins = df[df["in_scope"]]
    kids = ins.groupby("parent_uid")["uid"].apply(list).to_dict()
    df["children"] = df["uid"].map(lambda u: kids.get(u, []))
    df["is_leaf"] = df["children"].map(len).eq(0)
    df["source_file"] = os.path.basename(f)
    df.attrs["n_raw_rows"] = n_raw
    RAW_ROWS[rc["name"]] = int(n_raw)
    return df.reset_index(drop=True), f


def load_crosslist_edges(rc, file, nodes):
    """Staples cross-listings (child -> extra parent). Parents are given by NAME; ~50 names
    are shared by two shelves - each name resolves to ONE shelf (same-L1 non-leaf, else the
    shallowest non-leaf, else the shallowest)."""
    cl = rc.get("crosslist")
    if not cl:
        return []
    x = read_table(file, cl["sheet"], header=cl["header_row"])
    ins = nodes[nodes["in_scope"]]
    by_name = defaultdict(list)
    for u, n, leaf, lvl, pth in zip(ins["uid"], ins["name"], ins["is_leaf"], ins["depth"], ins["path"]):
        by_name[n].append((u, leaf, lvl, pth[0]))
    canon_parent = dict(zip(ins["uid"], ins["parent_uid"]))
    l1_of = dict(zip(ins["uid"], ins["path"].map(lambda p: p[0])))

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
        cands = by_name.get(child_name, [])
        if not cands:
            continue
        cu = sorted(cands, key=lambda c: (-int(c[1]), -c[2]))[0][0]
        for p in plist:
            pu = resolve(p, cu)
            if pu and pu != cu and pu != canon_parent.get(cu) and not pu.startswith(cu + " > ") \
                    and not cu.startswith(pu + " > "):
                edges.append((cu, pu, 1.0 / len(plist)))
    return list({(c, p): (c, p, w) for c, p, w in edges}.values())


def harmonise_counts(nodes, rc, xedges=None):
    """subtree_items: comparable item counts where the retailer publishes them (NaN = unknown).
    n_leaves: number of in-scope category leaves under the node (breadth, every retailer)."""
    mode = rc["count_mode"]
    ins = nodes["in_scope"].values
    uid_idx = {u: i for i, u in enumerate(nodes["uid"])}
    kids = defaultdict(set)
    for u, p, s in zip(nodes["uid"], nodes["parent_uid"], ins):
        if p and s:
            kids[p].add(u)
    for c, p, _ in (xedges or []):
        kids[p].add(c)
    leaf_flag = (nodes["is_leaf"] & nodes["in_scope"] & ~nodes["is_heading"]).values
    memo = {}
    sys.setrecursionlimit(20000)

    def leaves_under(u, stack=()):
        if u in memo:
            return memo[u]
        s = {u} if leaf_flag[uid_idx[u]] else set()
        for c in kids.get(u, ()):
            if c not in stack:
                s |= leaves_under(c, stack + (u,))
        memo[u] = s
        return s

    leafsets = [leaves_under(u) if s else set() for u, s in zip(nodes["uid"], ins)]
    nodes["leaf_set"] = leafsets
    nodes["n_leaves"] = [len(s) for s in leafsets]
    counts = nodes["count_raw"].values.astype(float)
    if mode == "terminal_only":
        nodes["subtree_items"] = [float(np.nansum([counts[uid_idx[v]] for v in s])) if s else 0.0
                                  for s in leafsets]
        nodes["count_known"] = ins
    elif mode in ("cumulative", "partial"):
        cap = rc.get("count_cap")
        stated = nodes["count_raw"].copy()
        if cap:
            stated[stated >= cap] = np.nan        # censored
        sub, known = {}, {}
        order = sorted(range(len(nodes)), key=lambda i: -nodes.at[i, "depth"])
        for i in order:
            u = nodes.at[i, "uid"]
            if pd.notna(stated.iat[i]):
                sub[u], known[u] = float(stated.iat[i]), True
            else:
                ch = [c for c in nodes.at[i, "children"]]
                if mode == "cumulative" and ch and all(known.get(c, False) for c in ch):
                    sub[u], known[u] = float(sum(sub[c] for c in ch)), True
                elif mode == "cumulative" and not ch:
                    sub[u], known[u] = np.nan, False
                else:
                    sub[u], known[u] = np.nan, False
        nodes["subtree_items"] = nodes["uid"].map(sub)
        nodes["count_known"] = nodes["uid"].map(known).fillna(False).astype(bool)
    else:
        nodes["subtree_items"] = np.nan
        nodes["count_known"] = False
    nodes["count_mode"] = mode
    return nodes


def retailer_totals(nodes, rc):
    """Catalogue size for normalising: items (if counted) and category leaves (always)."""
    ins = nodes[nodes["in_scope"]]
    leaves = int((ins["is_leaf"] & ~ins["is_heading"]).sum())
    if rc["count_mode"] == "terminal_only":
        items = float(ins.loc[ins["count_raw"].notna() & ins["is_leaf"], "count_raw"].sum())
    elif rc["count_mode"] == "cumulative":
        l1 = ins[ins["depth"] == 1]
        items = float(l1["subtree_items"].sum()) if l1["count_known"].all() else float(
            ins.loc[ins["is_leaf"] & ins["count_known"], "subtree_items"].sum())
    elif rc["count_mode"] == "partial":
        l2 = ins[(ins["depth"] == 2) & ins["count_known"]]
        items = float(l2["subtree_items"].sum())
    else:
        items = np.nan
    return {"items": items, "leaves": leaves}


def semantic_chain(path, heading_set):
    """Own label first, then ancestors bottom-up, skipping navigation headings and repeats."""
    chain, seen = [], set()
    for i in range(len(path), 0, -1):
        nm = path[i - 1]
        if i < len(path) and nm in heading_set:
            continue
        c = clean_name(nm)
        key = c.lower()
        if key in seen:
            continue
        seen.add(key)
        chain.append(c)
    return chain


def load_all(verbose=True):
    """Load, scope and harmonise all six trees. Returns (nodes, xedges, totals, files)."""
    frames, xedges, totals, files = [], {}, {}, {}
    for rc in RETAILERS:
        df, f = load_retailer(rc)
        files[rc["name"]] = f
        xe = load_crosslist_edges(rc, f, df) if rc.get("crosslist") else []
        xedges[rc["name"]] = xe
        df = harmonise_counts(df, rc, xe)
        totals[rc["name"]] = retailer_totals(df, rc)
        frames.append(df)
        if verbose:
            ins = df[df.in_scope]
            log(f"  {rc['name']:<12} rows {df.attrs.get('n_raw_rows', len(df)):>6} -> in scope {len(ins):>6} "
                f"(leaves {totals[rc['name']]['leaves']:>5}, max depth {ins['depth'].max()}, counts: {rc['count_mode']}"
                f"{', items %s' % format(int(totals[rc['name']]['items']), ',') if pd.notna(totals[rc['name']]['items']) else ''})"
                + (f"  + {len(xe)} cross-listing edges" if xe else ""))
    nodes = pd.concat(frames, ignore_index=True)
    heads = set(nodes.loc[nodes["is_heading"], "name"])
    nodes["sem_chain"] = nodes["path"].map(lambda p: semantic_chain(p, heads))
    nodes["name_expanded"] = [
        (f"{ch[1]} {ch[0]}" if len(ch) > 1 and is_generic(ch[0]) else ch[0]) for ch in nodes["sem_chain"]]
    nodes["sem_path"] = nodes["sem_chain"].map(lambda c: " > ".join(reversed(c)))
    nodes["unit_ok"] = nodes["in_scope"] & ~nodes["is_heading"]
    return nodes, xedges, totals, files


def staples_coreness(focal):
    depth_pct = focal.groupby("depth")["subtree_items"].rank(pct=True).fillna(0).values
    base = focal["path"].map(lambda p: CORE_L1_WEIGHT.get(p[0], CORE_DEFAULT)).values.astype(float)
    for pat, w in CORE_OVERRIDES:
        base[focal["path_str"].str.contains(pat, regex=True).values] = w
    return base * (0.6 + 0.4 * depth_pct)


def staples_descendants(focal, xedges):
    """All descendants of every Staples node, following canonical AND cross-listed placements."""
    kids = defaultdict(set)
    for u, p in zip(focal["uid"], focal["parent_uid"]):
        if p:
            kids[p].add(u)
    for c, p, _ in xedges:
        kids[p].add(c)
    memo = {}

    def desc(u, stack=()):
        if u in memo:
            return memo[u]
        s = {u}
        for c in kids.get(u, ()):
            if c not in stack:
                s |= desc(c, stack + (u,))
        memo[u] = s
        return s

    return {u: desc(u) for u in focal["uid"]}


# %% ------------------------------------------------------------------------------------
# LEXICAL SIMILARITY (both methods)
# ---------------------------------------------------------------------------------------
class Lexical:
    """Wording similarity: stemmed word uni/bi-grams + character 3-5 grams, averaged."""

    def __init__(self, texts):
        from sklearn.feature_extraction.text import TfidfVectorizer

        def words(s):
            t = [stem(x) for x in tokens(s)]
            return t + [a + "_" + b for a, b in zip(t, t[1:])]

        self.w = TfidfVectorizer(analyzer=words, sublinear_tf=True).fit(texts)
        self.c = TfidfVectorizer(analyzer="char_wb", ngram_range=(3, 5), sublinear_tf=True).fit(
            [t.lower() for t in texts])

    def transform(self, xs):
        from scipy.sparse import hstack
        xs = list(xs)
        return hstack([self.w.transform(xs) * np.sqrt(0.5),
                       self.c.transform([x.lower() for x in xs]) * np.sqrt(0.5)]).tocsr()


# %% ------------------------------------------------------------------------------------
# CALIBRATION (per competitor; both methods)
# ---------------------------------------------------------------------------------------
def load_gold(for_calibration=False):
    """Labelled rows. for_calibration=True keeps only the random / gap-enriched samples: rows from
    the verification loop were chosen BECAUSE the model flagged them, so they would bias tau; they
    are still used as overrides (truth beats the model) in run_framework.py."""
    f = find_file([CONFIG["gold_file"], "gold_matches*.csv"])
    if not f:
        log("  !! no gold file found - calibration falls back to score quantiles (indicative only)")
        return pd.DataFrame(columns=["competitor", "comp_path", "staples_carries", "gold_staples_path", "source"])
    g = pd.read_csv(f)
    g["staples_carries"] = g["staples_carries"].astype(int)
    if for_calibration:
        g = g[~g["source"].astype(str).str.contains("verification")]
    return g


def ablation_mask(focal, gold_path, comp_path):
    """True = keep. Hide the Staples department holding every gold alternative, plus any
    shelf with the same name as the answer or the competitor shelf."""
    hide = np.zeros(len(focal), dtype=bool)
    names = {comp_path.split(" > ")[-1].lower()}
    for alt in str(gold_path).split("|"):
        pre = " > ".join(alt.split(" > ")[:2])
        hide |= (focal["path_str"].eq(pre) | focal["path_str"].str.startswith(pre + " > ")).values
        names.add(alt.split(" > ")[-1].lower())
    hide |= focal["name"].str.lower().isin(names).values
    return ~hide


def calibrate(gold_c, best_score, ablation_scores, best_path):
    """tau  = 'same shelf' threshold (cost-weighted Youden's J on gold + ablation negatives)
       s_enter = gap screen: q-th percentile of scores of shelves Staples DOES carry
       neg  = median score of a non-match (anchors CRS at 0)."""
    rows = []
    for _, r in gold_c.iterrows():
        sc = best_score.get(r["comp_path"])
        if sc is None:
            continue
        rows.append({"comp_path": r["comp_path"], "y": int(r["staples_carries"]), "score": sc, "kind": "gold"})
        if int(r["staples_carries"]) == 1 and r["comp_path"] in ablation_scores:
            rows.append({"comp_path": r["comp_path"], "y": 0, "score": ablation_scores[r["comp_path"]],
                         "kind": "ablation"})
    cal_df = pd.DataFrame(rows)
    y, s = cal_df["y"].values, cal_df["score"].values
    c = CONFIG["false_gap_cost"]
    best = (float(np.median(s)), -1e9)
    for t in np.unique(np.round(s, 4)):
        pred = s >= t
        j = c * pred[y == 1].mean() - pred[y == 0].mean()
        if j > best[1]:
            best = (float(t), j)
    tau = best[0]
    from sklearn.metrics import roc_auc_score
    try:
        auc = float(roc_auc_score(y, s))
    except ValueError:
        auc = float("nan")
    g = cal_df[cal_df["kind"] == "gold"]
    pos = g[g.y == 1]["score"]
    s_enter = float(pos.quantile(CONFIG["gap_screen_false_alarm"])) if len(pos) else tau
    s_enter = min(s_enter, tau)
    negs = g[g.y == 0]["score"]
    # top-1: predicted Staples shelf equals a gold alternative, its parent or its child
    ok = n = 0
    for _, r in gold_c[(gold_c.staples_carries == 1) & gold_c.gold_staples_path.notna()].iterrows():
        pred = best_path.get(r["comp_path"])
        if pred is None:
            continue
        n += 1
        alts = str(r["gold_staples_path"]).split("|")
        ok += any(pred == a or pred.startswith(a + " > ") or a.startswith(pred + " > ") for a in alts)
    rnd = gold_c[gold_c.get("source", pd.Series("", index=gold_c.index)).astype(str).str.contains("stratified|random")]
    m = {"tau": tau, "s_enter": s_enter, "neg_median": float(cal_df.loc[cal_df.y == 0, "score"].median()),
         "auc": auc, "youden_j": best[1], "n_gold": int(len(g)), "n_pos": int((g.y == 1).sum()),
         "n_neg_real": int((g.y == 0).sum()), "n_neg_ablation": int((cal_df.kind == "ablation").sum()),
         "top1": ok / max(n, 1), "n_top1": n,
         "gold_accuracy": float(((g["score"] >= tau) == (g["y"] == 1)).mean()) if len(g) else np.nan,
         "false_gap_rate": float((pos < s_enter).mean()) if len(pos) else np.nan,
         "gap_recall_real": float((negs < s_enter).mean()) if len(negs) else np.nan,
         "base_rate_carried": float(rnd["staples_carries"].mean()) if len(rnd) else np.nan}
    return m, cal_df


def fallback_calibration(scores, pooled=None):
    """No labels for this competitor: borrow the pooled calibration (same scorer, same spine)."""
    if pooled:
        m = dict(pooled)
        m["note"] = "borrowed (pooled) calibration - label this competitor to confirm"
        return m
    return {"tau": float(np.quantile(scores, 0.10)), "s_enter": float(np.quantile(scores, 0.05)),
            "neg_median": float(np.quantile(scores, 0.02)), "note": "quantile fallback - indicative only"}


def status_from_scores(best, tau, s_enter):
    """matched (>= tau) | likely (carried, exact shelf not pinned) | gap (below the screen)"""
    return np.where(best >= tau, "matched", np.where(best < s_enter, "gap", "likely"))


# %% ------------------------------------------------------------------------------------
# NODE-LEVEL ROLL-UPS SHARED BY BOTH METHODS
# ---------------------------------------------------------------------------------------
def aas_scaler(A_leaves, matched_leaves):
    """AAS on ONE scale for every competitor and for both methods (quantile anchoring).
    F = empirical CDF of the method's raw adjacency over ALL competitor leaves, pooled;
        AAS = 100 x min(1, F(A) / F(median A of the leaves Staples carries))
    100 = as embedded in Staples' neighbourhood as a typical shelf Staples already carries;
    0 = the least embedded shelf in the panel. Pooling keeps one ruler across retailers;
    the quantile transform makes the vector and graph scales comparable before averaging."""
    ref = np.sort(np.asarray(A_leaves, dtype=float))
    f100 = np.searchsorted(ref, np.median(np.asarray(A_leaves)[np.asarray(matched_leaves)]), side="right") / len(ref)
    return lambda A: 100 * np.clip((np.searchsorted(ref, np.asarray(A, dtype=float), side="right") / len(ref))
                                   / max(f100, 1e-9), 0, 1)


def rollup_competitor(cn):
    """For each competitor node: breadth/items coverage and exposure-weighted AAS/CRS over leaves.
    cn needs: uid, is_leaf, unit_ok, status, subtree_items, count_known, AAS_leaf, CRS_leaf."""
    uid2row = {u: i for i, u in enumerate(cn["uid"])}
    leaf_rows = [i for i, (lf, ok) in enumerate(zip(cn["is_leaf"], cn["unit_ok"])) if lf and ok]
    desc = defaultdict(list)
    for i in leaf_rows:
        parts = cn.at[i, "path_str"].split(" > ")
        pre = cn.at[i, "uid"].split("::")[0] + "::"
        for k in range(1, len(parts) + 1):
            anc = pre + " > ".join(parts[:k])
            if anc in uid2row:
                desc[anc].append(i)
    status = cn["status"].values
    items = cn["subtree_items"].values.astype(float)
    known = cn["count_known"].values.astype(bool)
    aas, crs = cn["AAS_leaf"].values, cn["CRS_leaf"].values
    cov_leaf, cov_item, AAS, CRS, top_unc, n_gap_leaves = [], [], [], [], [], []
    for u in cn["uid"]:
        L = desc.get(u, [])
        own = uid2row[u]
        if not L:
            L = [own]
        gap = np.array([status[i] == "gap" for i in L])
        cov_leaf.append(1 - gap.mean())
        it = items[L]
        kn = known[L]
        if kn.all() and np.nansum(it) > 0:
            cov_item.append(float(np.nansum(it[~gap]) / np.nansum(it)))
            w = np.maximum(np.nan_to_num(it), 1.0)
        else:
            cov_item.append(np.nan)
            w = np.ones(len(L))
        AAS.append(float(np.average(aas[L], weights=w)))
        CRS.append(max(float(crs[own]), float(np.average(crs[L], weights=w))))
        n_gap_leaves.append(int(gap.sum()))
        unc = [cn.at[i, "name"] for i in L if status[i] == "gap"][:4]
        top_unc.append(", ".join(unc))
    cn["coverage_leaf"] = cov_leaf
    cn["coverage_items"] = cov_item
    cn["coverage"] = np.where(pd.notna(cn["coverage_items"]), cn["coverage_items"], cn["coverage_leaf"])
    cn["AAS"] = AAS
    cn["CRS"] = CRS
    cn["n_gap_leaves"] = n_gap_leaves
    cn["gap_leaves_example"] = top_unc
    return cn


# %% ------------------------------------------------------------------------------------
# TRACK C FRAMEWORK: labels, opportunity, sensitivity
# ---------------------------------------------------------------------------------------
def label_row(aas, crs, aas_hi=None, aas_lo=None, crs_hi=None, crs_lo=None):
    aas_hi = CONFIG["aas_hi"] if aas_hi is None else aas_hi
    aas_lo = CONFIG["aas_lo"] if aas_lo is None else aas_lo
    crs_hi = CONFIG["crs_hi"] if crs_hi is None else crs_hi
    crs_lo = CONFIG["crs_lo"] if crs_lo is None else crs_lo
    if crs >= crs_hi:
        return "1P-CORE GAP"
    if crs >= crs_lo:
        return "REVIEW"
    if aas >= aas_hi:
        return "CURATE"
    if aas >= aas_lo:
        return "VERTICAL EXTENSION"
    return "OFF-BRAND"


def opportunity(df, w=None, gamma=None, demand=None):
    """O = [w_pc*PC + w_gs*GS + w_aas*AAS/100] x (1 - CRS/100)^gamma   (all terms 0-1)
    PC  peer consensus  - share of the five competitors that show the gap
    GS  gap size        - how much of the carriers' catalogues it takes (percentile, 0-1)
    AAS adjacency, CRS cannibalisation risk (0-100)."""
    w = w or CONFIG["rank_weights"]
    gamma = CONFIG["gamma"] if gamma is None else gamma
    base = w["pc"] * df["PC"] + w["gs"] * df["GS"] + w["aas"] * df["AAS"] / 100
    if demand is not None and demand.notna().any():
        dw = CONFIG["demand_weight"]
        base = (1 - dw) * base + dw * demand.fillna(demand.median()) / 100
    return base * (1 - df["CRS"].clip(0, 100) / 100) ** gamma


def sensitivity(units, approved_label="CURATE"):
    """500 re-weightings (Dirichlet around the base weights), label thresholds jittered +-5,
    gamma in [0.5, 2]. Returns how often each unit is approved and lands in the top-N."""
    n, conc, jit = CONFIG["n_sensitivity"], CONFIG["dirichlet_conc"], CONFIG["threshold_jitter"]
    keys = list(CONFIG["rank_weights"])
    base = np.array([CONFIG["rank_weights"][k] for k in keys])
    r = np.random.default_rng(CONFIG["seed"])
    appr = np.zeros(len(units))
    topn = np.zeros(len(units))
    lockv = units["label"].eq("VERIFY").values
    for _ in range(n):
        w = dict(zip(keys, r.dirichlet(conc * base)))
        th = {k: CONFIG[k] + r.normal(0, jit) for k in ("aas_hi", "aas_lo", "crs_hi", "crs_lo")}
        lab = np.array([label_row(a, c, th["aas_hi"], th["aas_lo"], th["crs_hi"], th["crs_lo"])
                        for a, c in zip(units["AAS"], units["CRS"])])
        ok = (lab == approved_label) & ~lockv
        O = opportunity(units, w, CONFIG["gamma"] * r.uniform(0.5, 2.0)).values
        O = np.where(ok, O, -1)
        top = np.argsort(-O)[:CONFIG["top_n"]]
        top = top[O[top] > -1]
        appr += ok
        topn[top] += 1
    return pd.DataFrame({"unit_id": units["unit_id"].values, "p_approved": appr / n, "p_top_n": topn / n})


# %% ------------------------------------------------------------------------------------
# CHART STYLE
# ---------------------------------------------------------------------------------------
def set_style():
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    plt.rcParams.update({
        "font.family": "DejaVu Sans", "font.size": 10, "axes.titlesize": 12.5, "axes.titleweight": "bold",
        "axes.titlelocation": "left", "axes.labelsize": 10, "axes.spines.top": False, "axes.spines.right": False,
        "axes.edgecolor": INK["muted"], "axes.labelcolor": INK["secondary"], "xtick.color": INK["secondary"],
        "ytick.color": INK["secondary"], "text.color": INK["primary"], "axes.grid": False,
        "figure.facecolor": "white", "axes.facecolor": "white", "figure.dpi": 110, "savefig.dpi": 200,
        "savefig.bbox": "tight", "legend.frameon": False,
    })
    return plt


def _clean_json(o):
    """NaN/inf -> None and numpy scalars -> Python, so summary.json is strict JSON."""
    if isinstance(o, dict):
        return {str(k): _clean_json(v) for k, v in o.items()}
    if isinstance(o, (list, tuple)):
        return [_clean_json(v) for v in o]
    if hasattr(o, "item") and not isinstance(o, (str, bytes)):
        try:
            o = o.item()
        except Exception:  # noqa
            return str(o)
    if isinstance(o, float) and (math.isnan(o) or math.isinf(o)):
        return None
    return o


def save_json(obj, *parts):
    p = out_path(*parts)
    with open(p, "w") as fh:
        json.dump(_clean_json(obj), fh, indent=2, default=str, allow_nan=False)
    return p
