#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
=====================================================================================
 STAPLES MARKETPLACE PoC  -  CATEGORY-LEVEL RECOMMENDATION  (v2, six retailers)
 run_framework.py : ENSEMBLE + TRACK C FRAMEWORK + every chart and table in the documents
=====================================================================================

Run AFTER method_v_vector.py and method_g_graph.py. It reads their node scores and:

 A. ENSEMBLE    per competitor shelf, combines the two methods' verdicts
                (carried / likely / gap-hard / gap-soft / disputed) and applies labelled rows
 B. ENTER       finds competitor sub-trees Staples does not carry (>= 80% of leaves missing),
                at L1-L3 of the competitor tree
 C. CONCEPTS    merges ENTER sub-trees across the five competitors into category concepts
                (vector similarity, average linkage) and checks every competitor's Qdrant
                collection for the concept -> peer coverage "k of 5" (a count-free signal)
 D. DEEPEN      for every Staples L2/L3 aisle, compares each competitor's depth there with
                Staples' (items where both publish counts, else number of shelves), after
                size-factor normalisation -> "competitor is >= 2x deeper" votes
 E. TRACK C     AAS (adjacency) and CRS (cannibalisation) per unit = mean of Methods V and G;
                six labels: CURATE, VERTICAL EXTENSION, REVIEW, 1P-CORE GAP, OFF-BRAND, VERIFY
 F. RANK        O = [0.35 PC + 0.35 GS + 0.30 AAS/100] x (1 - CRS/100)^gamma; 500-run stress test
 G. OUTPUT      outputs/final/Category_Recommendations_v2.xlsx, figures fig01..fig14 (PNG),
                summary.json (every number quoted in the documents), graph/peer_edges.csv

Run:  python run_framework.py        (Colab: !python run_framework.py)
"""
# %% [0] IMPORTS ------------------------------------------------------------------------
import re
import time
from collections import Counter, defaultdict

import numpy as np
import pandas as pd

import poc_common as pc
from poc_common import CONFIG, FOCAL, LABELS, log

OUT = "final"

# Strategic themes (macro level): Staples department -> theme. Regex overrides on the Staples
# path come first, because Staples files some lifestyle shelves in odd places (patio furniture
# sits under Gift Shop > Professional Gifts > Compasses).
THEME_OVERRIDES = [
    (r"Sewing|Craft|Quilt|Party|Wedding|Christmas|Seasonal", "Celebrations, gifting & creativity"),
    (r"Patio|Outdoor (Decor|Furniture|Heaters|Lighting)|Compasses|Grounds|Garden|Planter|Gazebo|Weathervane|Cupola|"
     r"Railing|Shutter|Bird House|Birding|Chiminea|Fire Pit|Gazebo Canop|Pergola|Trellis|Awning|Greenhouse|Rain Barrel",
     "Outdoor living & grounds"),
    (r"Lamps & Lighting|Lighting|Decor|Rugs|Curtain|Accent|Sofa|Bedding|Mirror|Wall Art", "Home décor & furnishings"),
    (r"Fitness|Wearable|Massage|Sleep|Wellness|Health Monitors", "Wellness & ergonomics"),
    (r"Kitchen|Breakroom|Tableware|Drinkware|Cookware|Bakeware|Food Storage|Serveware|Appliances", "Breakroom, kitchen & hospitality"),
]
THEME_BY_L1 = {
    "Furniture": "Home décor & furnishings", "Decor": "Home décor & furnishings",
    "Fitness": "Wellness & ergonomics", "Wearable Technology": "Wellness & ergonomics",
    "Healthcare Supplies": "Health, safety & facilities", "Safety Supplies": "Health, safety & facilities",
    "Facilities": "Health, safety & facilities", "Cleaning Supplies": "Health, safety & facilities",
    "Tools, Parts & Supplies": "Health, safety & facilities", "Laboratory & Scientific Supplies": "Health, safety & facilities",
    "Workwear": "Health, safety & facilities",
    "Appliances & Kitchenware": "Breakroom, kitchen & hospitality", "Coffee, Water & Snacks": "Breakroom, kitchen & hospitality",
    "Restaurant & Foodservice Supplies": "Breakroom, kitchen & hospitality",
    "Party Supplies": "Celebrations, gifting & creativity", "Gift Shop": "Celebrations, gifting & creativity",
    "Arts & Crafts": "Celebrations, gifting & creativity", "School Supplies": "Celebrations, gifting & creativity",
    "Education": "Celebrations, gifting & creativity",
    "Office Supplies": "Core office & business supplies", "Paper": "Core office & business supplies",
    "Shipping, Packing & Mailing Supplies": "Core office & business supplies",
    "Retail Store Supplies": "Core office & business supplies", "Bags, Backpacks & Luggage": "Core office & business supplies",
    "Security, Banking & Cash": "Core office & business supplies", "Expanded Assortment": "Expanded assortment",
}
THEME_DEFAULT = "Technology & connected workspace"
THEME_ORDER = ["Home décor & furnishings", "Outdoor living & grounds", "Wellness & ergonomics",
               "Breakroom, kitchen & hospitality", "Celebrations, gifting & creativity",
               "Technology & connected workspace", "Health, safety & facilities", "Core office & business supplies",
               "Expanded assortment"]


def theme_of(staples_path):
    for pat, th in THEME_OVERRIDES:
        if re.search(pat, staples_path):
            return th
    return THEME_BY_L1.get(staples_path.split(" > ")[0], THEME_DEFAULT)


# %% [A] ENSEMBLE ------------------------------------------------------------------------
def ensemble_status(a, b):
    s = {a, b}
    if s <= {"matched", "likely"} and "matched" in s:
        return "carried"
    if s == {"likely"}:
        return "likely"
    if s == {"gap"}:
        return "gap_hard"
    if s == {"gap", "likely"}:
        return "gap_soft"
    return "disputed"                    # one method matched, the other found nothing close


def build_node_table(nodes, gold, focal_uid_of_path):
    V = pd.read_csv(pc.out_path("V", "node_scores_V.csv"))
    G = pd.read_csv(pc.out_path("G", "node_scores_G.csv"))
    cols = ["uid", "status", "AAS", "CRS", "best_score", "best_staples", "best_staples_uid", "crs_driver"]
    m = V[["competitor", "path_str", "sem_path", "depth", "is_leaf", "n_leaves", "subtree_items", "count_known"] + cols] \
        .merge(G[cols + ["sibling_coverage"]], on="uid", suffixes=("_V", "_G"))
    m["ens"] = [ensemble_status(a, b) for a, b in zip(m["status_V"], m["status_G"])]
    m["label_source"] = "model"
    # labelled rows override the models (human-in-the-loop)
    gl = gold.set_index(["competitor", "comp_path"])
    for i, (c, p) in enumerate(zip(m["competitor"], m["path_str"])):
        if (c, p) in gl.index:
            r = gl.loc[(c, p)]
            r = r.iloc[0] if isinstance(r, pd.DataFrame) else r
            if int(r["staples_carries"]) == 1:
                m.at[i, "ens"] = "carried"
                first = str(r["gold_staples_path"]).split("|")[0]
                if first in focal_uid_of_path:
                    m.at[i, "best_staples_uid_V"] = focal_uid_of_path[first]
                    m.at[i, "best_staples_V"] = first
            else:
                m.at[i, "ens"] = "gap_hard"
            m.at[i, "label_source"] = "labelled"
    m["AAS"] = (m["AAS_V"] + m["AAS_G"]) / 2
    m["CRS"] = (m["CRS_V"] + m["CRS_G"]) / 2
    m["is_gap"] = m["ens"].isin(["gap_hard", "gap_soft"])
    nm = nodes.set_index("uid")
    m["name"] = m["uid"].map(nm["name"])
    m["name_clean"] = m["uid"].map(nm["name_clean"])
    m["name_expanded"] = m["uid"].map(nm["name_expanded"])
    m["parent_uid"] = m["uid"].map(nm["parent_uid"])
    m["is_heading"] = m["uid"].map(nm["is_heading"]).astype(bool)
    return m


def gap_rollup(m):
    """For every competitor node: share of its leaves that are gaps (items-weighted when all
    leaves carry counts), share of those gaps that are only 'soft'."""
    out = []
    for comp, cn in m.groupby("competitor"):
        cn = cn.copy()
        uid2i = {u: i for i, u in enumerate(cn["uid"])}
        leaves = [i for i, lf in enumerate(cn["is_leaf"].values) if lf]
        desc = defaultdict(list)
        pre = comp + "::"
        paths = cn["path_str"].values
        for i in leaves:
            parts = paths[i].split(" > ")
            for k in range(1, len(parts) + 1):
                a = pre + " > ".join(parts[:k])
                if a in uid2i:
                    desc[a].append(i)
        gap = cn["is_gap"].values
        soft = cn["ens"].eq("gap_soft").values
        items = cn["subtree_items"].values.astype(float)
        known = cn["count_known"].values.astype(bool)
        gs, ss, ng = [], [], []
        for u in cn["uid"]:
            L = desc.get(u) or [uid2i[u]]
            L = np.array(L)
            if known[L].all() and np.nansum(items[L]) > 0:
                w = np.nan_to_num(items[L])
            else:
                w = np.ones(len(L))
            gs.append(float((w * gap[L]).sum() / max(w.sum(), 1e-9)))
            ss.append(float(soft[L][gap[L]].mean()) if gap[L].any() else 0.0)
            ng.append(int(gap[L].sum()))
        cn["gap_share"], cn["soft_share"], cn["n_gap_leaves"] = gs, ss, ng
        out.append(cn)
    return pd.concat(out, ignore_index=True)


# %% [B] ENTER UNITS ---------------------------------------------------------------------
def enter_units(m, totals):
    cand = m["is_gap"] & (m["gap_share"] >= 1 - CONFIG["enter_coverage_max"]) & ~m["is_heading"]
    cand &= m["depth"] <= CONFIG["max_unit_depth"]
    cset = set(m.loc[cand, "uid"])
    keep = cand & ~m["parent_uid"].isin(cset)
    u = m[keep].copy()
    # size filter: >= 10 items where counted, else >= 1 shelf
    small = u["count_known"].astype(bool) & (u["subtree_items"].fillna(0) < 10)
    u = u[~small]
    log(f"      {len(u):,} ENTER sub-trees (L1-L3, >= 80% of leaves missing at Staples)")
    return u.reset_index(drop=True)


# %% [C] CONCEPTS + PEER COVERAGE ---------------------------------------------------------
class PeerIndex:
    """Queries the Method V Qdrant DB (one collection per retailer, uid in the payload);
    falls back to NumPy on the saved vectors if the DB is not available."""

    def __init__(self, work_uids, V):
        self.pos = {u: i for i, u in enumerate(work_uids)}
        self.V = V
        self.client = None
        self.backend = "numpy"
        try:
            from qdrant_client import QdrantClient
            self.client = QdrantClient(path=pc.out_path("V", "qdrant_db"))
            self.backend = "qdrant"
        except Exception as e:  # noqa
            log(f"      Qdrant DB not available ({e}); NumPy fallback")
        self.by_ret = {}

    def register(self, retailer, uids):
        self.by_ret[retailer] = np.array(uids)

    def search(self, retailer, q, k=10):
        """-> (uids, cosine scores) of the k nearest shelves in that retailer's collection"""
        if self.client is not None:
            res = self.client.query_points(retailer.lower(), query=q.tolist(), limit=k, with_payload=True).points
            return [p.payload["uid"] for p in res], [p.score for p in res]
        uids = self.by_ret[retailer]
        M = self.V[[self.pos[u] for u in uids]] @ q
        I = np.argsort(-M)[:k]
        return list(uids[I]), list(M[I])


class SizeRef:
    """Gap size on one comparable scale: the percentile of a missing chunk of assortment among
    ALL of that competitor's L1-L3 categories - by number of shelves (every retailer) and, where
    the retailer publishes counts, by items too (the two percentiles are averaged)."""

    def __init__(self, m, totals):
        self.ref = {}
        for comp, cn in m[m["depth"] <= CONFIG["max_unit_depth"]].groupby("competitor"):
            sh = np.sort((cn["n_leaves"] / max(totals[comp]["leaves"], 1)).values)
            ki = cn["count_known"].astype(bool) & cn["subtree_items"].notna()
            it = np.sort((cn.loc[ki, "subtree_items"] / totals[comp]["items"]).values) \
                if pd.notna(totals[comp]["items"]) and ki.sum() >= 20 else None
            self.ref[comp] = (sh, it)

    def pct(self, comp, shelves_share, items_share=None):
        sh, it = self.ref[comp]
        p = [np.searchsorted(sh, shelves_share, side="right") / len(sh)]
        if it is not None and items_share is not None and pd.notna(items_share):
            p.append(np.searchsorted(it, items_share, side="right") / len(it))
        return float(np.mean(p))


def cluster_concepts(units, Vmap, lex, tau_peer):
    from sklearn.cluster import AgglomerativeClustering
    E = np.stack([Vmap[u] for u in units["uid"]])
    L = lex.transform(units["name_expanded"])
    S = CONFIG["w_sem"] * (E @ E.T) + CONFIG["w_lex"] * (L @ L.T).toarray()
    np.fill_diagonal(S, 1.0)
    if len(units) == 1:
        return np.array([0]), S
    cl = AgglomerativeClustering(n_clusters=None, metric="precomputed", linkage="average",
                                 distance_threshold=1 - tau_peer).fit(1 - np.clip(S, 0, 1))
    return cl.labels_, S


def concept_table(units, m, nodes, peer, Vmap, lex, cal_V, totals, focal):
    """One row per ENTER concept with peer coverage from every competitor's collection."""
    taus = cal_V.set_index("competitor")["tau"].to_dict()
    tau_peer = float(np.median(list(taus.values())))
    labels, S = cluster_concepts(units, Vmap, lex, tau_peer)
    units["concept"] = labels
    node_by_uid = m.set_index("uid")
    sref = SizeRef(m, totals)

    def size_of(u):
        if u not in node_by_uid.index:
            return 0.5
        r = node_by_uid.loc[u]
        c = r["competitor"]
        ish = r["subtree_items"] / totals[c]["items"] if (bool(r["count_known"]) and pd.notna(totals[c]["items"])) else None
        return sref.pct(c, r["n_leaves"] / max(totals[c]["leaves"], 1), ish)

    rows, members, peer_edges = [], [], []
    for cid, g in units.groupby("concept"):
        idx = g.index.values
        sub = S[np.ix_(idx, idx)]
        medoid = g.index[np.argmax(sub.mean(axis=1))] if len(g) > 1 else g.index[0]
        centroid = np.mean([Vmap[u] for u in g["uid"]], axis=0)
        centroid /= np.linalg.norm(centroid) + 1e-9
        cname = str(g.at[medoid, "name_clean"])
        carriers, contra, sizes = {}, 0, []
        for comp in pc.COMPETITORS:
            mine = g[g["competitor"] == comp]
            if len(mine):
                best = mine.sort_values("n_leaves", ascending=False).iloc[0]
                carriers[comp] = ("member", best["uid"], 1.0)
                sizes.append(size_of(best["uid"]))
                continue
            cand_uids, sc = peer.search(comp, centroid, k=10)
            Lc = lex.transform([node_by_uid.at[u, "name_expanded"] if u in node_by_uid.index else "" for u in cand_uids])
            lx = (Lc @ lex.transform([g.at[medoid, "name_expanded"]]).T).toarray().ravel()
            hyb = CONFIG["w_sem"] * np.array(sc) + CONFIG["w_lex"] * lx
            j = int(np.argmax(hyb))
            if hyb[j] >= taus.get(comp, tau_peer):
                u = cand_uids[j]
                carriers[comp] = ("peer", u, float(hyb[j]))
                sizes.append(size_of(u))
                if u in node_by_uid.index and node_by_uid.at[u, "ens"] == "carried":
                    contra += 1
                peer_edges.append((g.at[medoid, "uid"], u, float(hyb[j])))
        # home department at Staples: similarity-weighted vote of the 10 nearest Staples shelves
        st_uids, sc = peer.search(FOCAL, centroid, k=10)
        st_paths = [u.split("::", 1)[1] for u in st_uids]
        votes = Counter()
        for p_, s_ in zip(st_paths, sc):
            votes[p_.split(" > ")[0]] += max(s_, 0)
        home_l1 = votes.most_common(1)[0][0]
        best_in_l1 = [p_ for p_ in st_paths if p_.split(" > ")[0] == home_l1][0]
        home_path = " > ".join(best_in_l1.split(" > ")[:2])
        w = g["n_leaves"].clip(lower=1).values
        k = len(carriers)
        soft = float(np.average(g["soft_share"], weights=w))
        n_items = g.loc[g["count_known"].astype(bool), "subtree_items"].sum()
        rows.append({
            "unit_id": f"E{cid:04d}", "gap_type": "ENTER", "name": cname,
            "example_paths": " | ".join(f"{c}: {p}" for c, p in zip(g["competitor"], g["path_str"]))[:600],
            "home_L1": home_l1, "home_path": home_path, "nearest_staples_shelf": st_paths[0],
            "nearest_staples_score": float(sc[0]),
            "k": k, "carriers": ", ".join(sorted(carriers)), "members_n": len(g),
            "member_competitors": int(g["competitor"].nunique()),
            "contradicting_peers": contra, "soft_gap_share": soft,
            "PC": k / len(pc.COMPETITORS), "GS": float(np.mean(sizes)) if sizes else 0.0,
            "AAS": float(np.average(g["AAS"], weights=w)), "CRS": float(np.average(g["CRS"], weights=w)),
            "AAS_V": float(np.average(g["AAS_V"], weights=w)), "AAS_G": float(np.average(g["AAS_G"], weights=w)),
            "CRS_V": float(np.average(g["CRS_V"], weights=w)), "CRS_G": float(np.average(g["CRS_G"], weights=w)),
            "competitor_shelves": int(g["n_leaves"].sum()), "competitor_items_counted": float(n_items),
            "staples_shelves": 0, "crs_driver": g.sort_values("CRS", ascending=False)["crs_driver_V"].iloc[0],
            "centroid_uid": g.at[medoid, "uid"], "labelled_members": int(g["label_source"].eq("labelled").sum()),
        })
        for _, r in g.iterrows():
            members.append({"unit_id": f"E{cid:04d}", "competitor": r["competitor"], "path": r["path_str"],
                            "depth": r["depth"], "shelves": r["n_leaves"], "items": r["subtree_items"],
                            "ensemble": r["ens"], "V": r["status_V"], "G": r["status_G"],
                            "AAS": r["AAS"], "CRS": r["CRS"], "nearest_staples": r["best_staples_V"],
                            "score_V": r["best_score_V"], "label_source": r["label_source"]})
        for comp, (kind, u, s_) in carriers.items():
            if kind == "peer":
                members.append({"unit_id": f"E{cid:04d}", "competitor": comp, "path": u.split("::", 1)[1],
                                "depth": np.nan, "shelves": node_by_uid.at[u, "n_leaves"] if u in node_by_uid.index else np.nan,
                                "items": node_by_uid.at[u, "subtree_items"] if u in node_by_uid.index else np.nan,
                                "ensemble": "peer carrier" + (" (matches a Staples shelf)" if u in node_by_uid.index and
                                                              node_by_uid.at[u, "ens"] == "carried" else ""),
                                "V": "", "G": "", "AAS": np.nan, "CRS": np.nan, "nearest_staples": "",
                                "score_V": s_, "label_source": "peer search"})
    log(f"      {len(rows):,} ENTER concepts (peer threshold {tau_peer:.3f}); "
        f"{sum(r['k'] >= 2 for r in rows)} carried by 2+ competitors, {sum(r['k'] >= 3 for r in rows)} by 3+")
    return pd.DataFrame(rows), pd.DataFrame(members), peer_edges


# %% [D] DEEPEN UNITS --------------------------------------------------------------------
def deepen_units(m, nodes, focal, xedges, totals):
    desc = pc.staples_descendants(focal, xedges[FOCAL])
    anc = defaultdict(set)
    for X, ds in desc.items():
        for s in ds:
            anc[s].add(X)
    f = focal.set_index("uid")
    st_leaf = f["is_leaf"] & ~f["is_heading"]
    st_leaves = {X: [s for s in ds if st_leaf.get(s, False)] for X, ds in desc.items()}
    st_count = f["count_raw"]
    st_items = {X: float(np.nansum([st_count.get(s, np.nan) for s in L])) for X, L in st_leaves.items()}
    lv = m[m["is_leaf"] & m["ens"].isin(["carried", "likely", "disputed"])]
    per = defaultdict(lambda: defaultdict(lambda: {"leaves": 0, "items": 0.0, "known": 0, "aas": [], "crs": [],
                                                   "names": []}))
    for r in lv.itertuples(index=False):
        s = r.best_staples_uid_V
        for X in anc.get(s, ()):
            d = per[X][r.competitor]
            d["leaves"] += 1
            if r.count_known and pd.notna(r.subtree_items):
                d["items"] += float(r.subtree_items)
                d["known"] += 1
            d["aas"].append(r.AAS)
            d["crs"].append(r.CRS)
            d["names"].append(r.name_clean)
    # size factors per competitor
    sf_items, sf_leaves = {}, {}
    for comp in pc.COMPETITORS:
        c = lv[(lv.competitor == comp) & lv["ens"].eq("carried") & lv["count_known"].astype(bool)]
        ratios = [np.log(a) - np.log(st_count.get(s)) for a, s in zip(c["subtree_items"], c["best_staples_uid_V"])
                  if a > 0 and pd.notna(st_count.get(s, np.nan)) and st_count.get(s) > 0]
        sf_items[comp] = float(np.exp(np.median(ratios))) if len(ratios) >= 20 else np.nan
        rl = [np.log(per[X][comp]["leaves"]) - np.log(len(st_leaves[X])) for X in focal.loc[focal.depth == 2, "uid"]
              if per[X][comp]["leaves"] >= 3 and len(st_leaves[X]) >= 1]
        sf_leaves[comp] = float(np.exp(np.median(rl))) if len(rl) >= 10 else 1.0
    rows, evid = [], []
    # candidate aisles: real parents, or terminal shelves with a published count. Count-less
    # canonical leaves that only hold cross-listed siblings ("Plumb Bobs" lists 7 sibling shelves)
    # are navigation hubs, not aisles.
    cand = focal[(focal.depth.isin([2, 3])) & ~focal.is_heading & (~focal.is_leaf | focal.count_raw.notna())]
    for X, path, depth in zip(cand["uid"], cand["path_str"], cand["depth"]):
        comps = per.get(X, {})
        nL, nI = len(st_leaves[X]), st_items[X]
        dgs, votes, aas, crs, ex, names = {}, [], [], [], {}, Counter()
        for comp, d in comps.items():
            if d["leaves"] == 0:
                continue
            use_items = (d["known"] >= 0.8 * d["leaves"]) and pd.notna(sf_items[comp]) and nI > 0
            dg_sh = np.log1p(d["leaves"]) - np.log1p(sf_leaves[comp] * nL)
            ex_sh = max(d["leaves"] - sf_leaves[comp] * nL, 0) / max(totals[comp]["leaves"], 1)
            ex_it = None
            if use_items:
                dg = np.log1p(d["items"]) - np.log1p(sf_items[comp] * nI)
                comp_size, base = d["items"], sf_items[comp] * nI
                ex_it = max(d["items"] - sf_items[comp] * nI, 0) / totals[comp]["items"]
            else:
                dg, comp_size, base = dg_sh, d["leaves"], sf_leaves[comp] * nL
            dgs[comp] = (float(dg), "items" if use_items else "shelves", comp_size, base)
            ex[comp] = (ex_sh, ex_it)
            if dg >= CONFIG["deepen_min"]:
                votes.append(comp)
                names.update(d["names"])
            aas += d["aas"]
            crs += d["crs"]
        if not dgs:
            continue
        for comp, (dg, basis, cs, b) in dgs.items():
            evid.append({"staples_uid": X, "staples_path": path, "competitor": comp, "depth_gap": dg, "basis": basis,
                         "competitor_size": cs, "staples_size_scaled": b, "excess_shelf_share": ex[comp][0],
                         "excess_item_share": ex[comp][1], "deeper_2x": dg >= CONFIG["deepen_min"]})
        is_deepen = len(votes) >= CONFIG["deepen_min_carriers"] and len(votes) >= 0.5 * len(dgs)
        rows.append({"staples_uid": X, "path": path, "depth": depth, "deepen": is_deepen, "votes": len(votes),
                     "voters": ", ".join(sorted(votes)), "carriers_n": len(dgs),
                     "median_depth_gap": float(np.median([v[0] for v in dgs.values()])),
                     "AAS": float(np.mean(aas)), "CRS": float(np.mean(crs)),
                     "staples_shelves": nL, "staples_items": nI,
                     "excess": {c: ex[c] for c in votes},
                     "competitor_view": ", ".join(n for n, _ in names.most_common(3))})
    D = pd.DataFrame(rows)
    # keep L3 aisles; keep an L2 only when none of its L3 children is itself a DEEPEN aisle
    dd = D[D.deepen]
    l3_parents = {p.rsplit(" > ", 1)[0] for p in dd.loc[dd.depth == 3, "path"]}
    dd = dd[(dd.depth == 3) | ~dd["path"].isin(l3_parents)].copy()
    log("      size factors (items): " + ", ".join(f"{c} {v:.2f}" for c, v in sf_items.items() if pd.notna(v))
        + " | (shelves): " + ", ".join(f"{c} {v:.2f}" for c, v in sf_leaves.items()))
    log(f"      {len(dd):,} DEEPEN aisles (>= 2 competitors at least 2x Staples' relative depth)")
    return dd, pd.DataFrame(evid), sf_items, sf_leaves, SizeRef(m, totals)


def deepen_table(dd, sref):
    rows = []
    for i, r in enumerate(dd.itertuples(index=False)):
        voters = [v for v in r.voters.split(", ") if v]
        gs = float(np.mean([sref.pct(c, *r.excess[c]) for c in voters])) if voters else 0.0
        parts = r.path.split(" > ")
        # Staples' own label can be misleading (patio furniture sits under "Compasses"): show what the
        # competitors actually stock there when the two share no words
        own = {pc.stem(t) for t in pc.tokens(parts[-1])}
        seen = {pc.stem(t) for t in pc.tokens(r.competitor_view)}
        disp = parts[-1] if own & seen or not r.competitor_view else f"{parts[-1]} (competitors: {r.competitor_view})"
        rows.append({"unit_id": f"D{i:04d}", "gap_type": "DEEPEN", "name": disp, "competitor_view": r.competitor_view,
                     "example_paths": r.path, "home_L1": parts[0], "home_path": " > ".join(parts[:2]),
                     "nearest_staples_shelf": r.path, "nearest_staples_score": 1.0,
                     "k": r.votes, "carriers": r.voters, "members_n": r.carriers_n, "contradicting_peers": 0,
                     "soft_gap_share": 0.0, "PC": r.votes / len(pc.COMPETITORS), "GS": gs,
                     "AAS": r.AAS, "CRS": r.CRS, "AAS_V": np.nan, "AAS_G": np.nan, "CRS_V": np.nan, "CRS_G": np.nan,
                     "competitor_shelves": np.nan, "competitor_items_counted": np.nan,
                     "staples_shelves": r.staples_shelves, "staples_items": r.staples_items,
                     "crs_driver": r.path, "centroid_uid": r.staples_uid,
                     "labelled_members": 0, "member_competitors": r.votes,
                     "median_depth_gap": r.median_depth_gap, "staples_depth": r.depth})
    T = pd.DataFrame(rows)
    # method-specific AAS / CRS for DEEPEN aisles: mean over the competitor leaves assigned to them
    return T


# %% [E/F] LABEL, RANK, STRESS-TEST --------------------------------------------------------
def label_and_rank(U, stress=True):
    U["label"] = [pc.label_row(a, c) for a, c in zip(U["AAS"], U["CRS"])]
    U["label_V"] = [pc.label_row(a, c) if pd.notna(a) else "" for a, c in zip(U["AAS_V"], U["CRS_V"])]
    U["label_G"] = [pc.label_row(a, c) if pd.notna(a) else "" for a, c in zip(U["AAS_G"], U["CRS_G"])]
    U["zone_before_verify"] = U["label"]
    # VERIFY: an ENTER concept is probably carried under another name when most of its gap is
    # 'soft' (one method finds a plausible Staples shelf) or when most peers' equivalent shelves
    # DO match a Staples shelf
    # VERIFY: an actionable ENTER concept whose absence at Staples is not yet confirmed - most of
    # its gap is 'soft' (one method finds a plausible Staples shelf) or most peers' equivalent
    # shelves DO match a Staples shelf. A verified (labelled) member settles it; an OFF-BRAND
    # concept stays OFF-BRAND (the action - pass - is the same either way).
    peers = (U["k"] - U["member_competitors"].fillna(U["k"])).clip(lower=0)
    unsure = (U.soft_gap_share >= 0.5) | ((peers >= 1) & (U.contradicting_peers >= 0.5 * peers))
    ver = (U.gap_type == "ENTER") & (U.labelled_members == 0) & unsure & (U.label != "OFF-BRAND")
    U.loc[ver, "label"] = "VERIFY"
    U["O"] = pc.opportunity(U)
    agree = (U["label_V"] == U["label_G"]) | (U.gap_type == "DEEPEN")
    U["evidence"] = np.select([(U.k >= 3) & agree, (U.k >= 2) | agree], ["A", "B"], "C")
    U["action"] = U["label"].map(pc.ACTION)
    U["theme"] = [theme_of(p if gt == "DEEPEN" else f"{h} > {n} > {s}")
                  for p, gt, h, n, s in zip(U["example_paths"], U["gap_type"], U["home_path"], U["name"],
                                            U["nearest_staples_shelf"])]
    if stress:
        sens = pc.sensitivity(U)
        U = U.merge(sens, on="unit_id", how="left")
    U["verify_outcome"] = np.where(U["label"] == "VERIFY", "to check", "")
    U["rank_in_zone"] = U.groupby("label")["O"].rank(ascending=False, method="first").astype(int)
    return U.sort_values(["label", "O"], ascending=[True, False]).reset_index(drop=True)


# %% [G] MAIN ---------------------------------------------------------------------------
DEPT_CANON = {}   # filled in assign_families: punctuation-insensitive department names ("Arts Crafts & Sewing")


def assign_families(U, members, m):
    """Roll units up to the level Staples would act on: a Staples L2 aisle.
    DEEPEN   -> the unit's own Staples L2 aisle
    ENTER    -> the aisle it docks into in the knowledge graph: the competitor shelf's parent aisle
                (which Staples DOES carry) followed across its SAME_AS link to Staples (Method G;
                Method V's link, then the neighbour vote, as fallbacks)
    VERIFY   -> the Staples aisle where the 'missing' category was found
    OFF-BRAND-> the competitor department it comes from (there is no Staples aisle to dock into)"""
    nd = m.set_index("uid")
    l2 = lambda p: " > ".join(str(p).split(" > ")[:2])
    for p in members["path"].astype(str):
        k = p.split(" > ")[0]
        DEPT_CANON.setdefault(re.sub(r"[^a-z]", "", k.lower().replace(" and ", "")), k)
    dock, dept = [], []
    for _, r in U.iterrows():
        mm = members[(members.unit_id == r["unit_id"]) & (members.label_source != "peer search")]
        depts = Counter(p.split(" > ")[0] for p in mm["path"]) if len(mm) else Counter()
        depts = Counter({DEPT_CANON.get(re.sub(r"[^a-z]", "", k.lower().replace(" and ", "")), k): v for k, v in depts.items()})
        dept.append(depts.most_common(1)[0][0] if depts else "")
        if r["gap_type"] == "DEEPEN":
            dock.append(l2(r["example_paths"]))
            continue
        if r["label"] == "VERIFY" and str(r["verify_outcome"]).startswith("already carried: "):
            dock.append(l2(r["verify_outcome"].replace("already carried: ", "")))
            continue
        # vote: each member's carried parent aisle, followed across Method G's and Method V's SAME_AS link,
        # plus the Staples-neighbour vote (1.5 - breaks ties)
        votes = Counter({r["home_path"]: 1.5})
        for c, p in zip(mm["competitor"], mm["path"]):
            u = f"{c}::{p}"
            for _ in range(3):
                u = nd.at[u, "parent_uid"] if u in nd.index else None
                if u is None or pd.isna(u) or u not in nd.index:
                    break
                if nd.at[u, "ens"] in ("carried", "likely"):
                    for meth in ("G", "V"):
                        if nd.at[u, f"status_{meth}"] != "gap":
                            votes[l2(nd.at[u, f"best_staples_{meth}"])] += 1
                    break
        dock.append(votes.most_common(1)[0][0])
    U["aisle"] = dock
    U["competitor_dept"] = dept
    U["family"] = np.where((U["label"] == "OFF-BRAND") & (U["gap_type"] == "ENTER") & (U["competitor_dept"] != ""),
                           "Competitor dept: " + U["competitor_dept"], U["aisle"])
    return U


def cleared_by_verification(E1, members1, gold):
    """Concepts the models flagged (model-only pass) whose every model-flagged member was found,
    in the verification loop, to be carried by Staples under another name: findability issues."""
    ver = gold[gold["source"].astype(str).str.contains("verification")]
    vmap = {(c, p): (int(k), str(g)) for c, p, k, g in
            zip(ver["competitor"], ver["comp_path"], ver["staples_carries"], ver["gold_staples_path"])}
    rows = []
    for _, r in E1[(E1["gap_type"] == "ENTER") & (E1["zone_before_verify"] != "OFF-BRAND")].iterrows():
        allm = members1[(members1.unit_id == r["unit_id"]) & (members1.label_source != "peer search")]
        if (allm.label_source == "labelled").any():      # a labelled member says "not carried": keep it a gap
            continue
        mm = allm[allm.label_source == "model"]
        if mm.empty:
            continue
        res = [vmap.get((c, p)) for c, p in zip(mm["competitor"], mm["path"])]
        if any(x is None for x in res):
            continue
        if all(x[0] == 1 for x in res):
            rr = r.copy()
            where = Counter(x[1].split("|")[0] for x in res).most_common(1)[0][0]
            rr["label"] = "VERIFY"
            rr["verify_outcome"] = f"already carried: {where}"
            rr["nearest_staples_shelf"] = where
            rr["unit_id"] = "V" + r["unit_id"][1:]
            rr["action"] = "Already sold - fix naming / navigation so customers find it (findability, not assortment)"
            rows.append(rr)
    return pd.DataFrame(rows)


def main():
    t0 = time.time()
    log("\n[1] Loading trees, labels and both methods' node scores")
    nodes, xedges, totals, files = pc.load_all(verbose=False)
    gold = pc.load_gold()
    focal = nodes[(nodes.retailer == FOCAL) & nodes.unit_ok].reset_index(drop=True)
    focal["coreness"] = pc.staples_coreness(focal)
    uid_of_path = dict(zip(focal["path_str"], focal["uid"]))
    cal_V = pd.read_csv(pc.out_path("V", "calibration_V.csv"))
    cal_G = pd.read_csv(pc.out_path("G", "calibration_G.csv"))
    vec = np.load(pc.out_path("V", "vectors_V.npz"), allow_pickle=True)
    Vmap = {u: v for u, v in zip(vec["uid"], vec["V"])}
    lex = pc.Lexical(sorted(set(nodes.loc[nodes.in_scope, "name_expanded"])))
    peer = PeerIndex(vec["uid"], vec["V"])
    work = nodes[nodes.unit_ok]
    for r in [FOCAL] + pc.COMPETITORS:
        peer.register(r, work.loc[work.retailer == r, "uid"].values)

    log("\n[2] Model-only pass (calibration labels only) - what the models alone would flag")
    gold_model = gold[~gold["source"].astype(str).str.contains("verification")]
    m1 = gap_rollup(build_node_table(nodes, gold_model, uid_of_path))
    E1, members1, _ = concept_table(enter_units(m1, totals), m1, nodes, peer, Vmap, lex, cal_V, totals, focal)
    E1 = label_and_rank(E1, stress=False)
    log(f"      model-only: {int((E1.label == 'VERIFY').sum())} ENTER concepts would need verification")

    log("\n[3] Final pass with every labelled row (verification loop applied)")
    m = gap_rollup(build_node_table(nodes, gold, uid_of_path))
    log(f"      {len(m):,} competitor shelves; ensemble status: "
        + ", ".join(f"{k} {v:.0%}" for k, v in m["ens"].value_counts(normalize=True).items()))
    log("      ENTER: competitor sub-trees Staples does not carry")
    eu = enter_units(m, totals)
    log("      concepts across competitors + peer coverage from the Qdrant collections")
    E, members, peer_edges = concept_table(eu, m, nodes, peer, Vmap, lex, cal_V, totals, focal)

    log("\n[4] DEEPEN: Staples aisles where competitors are at least 2x deeper")
    dd, evid, sf_i, sf_l, sref = deepen_units(m, nodes, focal, xedges, totals)
    D = deepen_table(dd, sref)

    log("\n[5] Track C labels, opportunity score and 500-run stress test")
    U = label_and_rank(pd.concat([E, D], ignore_index=True))
    cleared = cleared_by_verification(E1, members1, gold)
    if len(cleared):
        U = pd.concat([U, cleared], ignore_index=True)
    U["rank_in_zone"] = U.groupby("label")["O"].rank(ascending=False, method="first").astype(int)
    U = U.sort_values(["label", "O"], ascending=[True, False]).reset_index(drop=True)
    U = assign_families(U, pd.concat([members, members1[members1.unit_id.isin(E1.unit_id)].assign(
        unit_id=lambda d: "V" + d.unit_id.str[1:])]), m)
    U["theme"] = [theme_of(f"{a} > {n}") for a, n in zip(U["aisle"], U["name"])]
    for lab in LABELS:
        log(f"      {lab:<20} {int((U.label == lab).sum()):>4} units")
    vstats = {"model_only_verify_concepts": int((E1.label == "VERIFY").sum()),
              "cleared_concepts": int(len(cleared)), "staples_crosslist_edges": int(len(xedges[FOCAL]))}

    log("\n[6] Writing tables, figures and summary")
    import framework_outputs as fo
    fo.write_all(U=U, members=members, evid=evid, m=m, nodes=nodes, focal=focal, totals=totals, cal_V=cal_V,
                 cal_G=cal_G, gold=gold, sf_items=sf_i, sf_leaves=sf_l, peer_edges=peer_edges, Vmap=Vmap,
                 files=files, theme_order=THEME_ORDER, vstats=vstats, E1=E1)
    log(f"\nDone in {time.time() - t0:.0f}s. Outputs in {pc.out_path(OUT, '')}")


if __name__ == "__main__":
    main()
