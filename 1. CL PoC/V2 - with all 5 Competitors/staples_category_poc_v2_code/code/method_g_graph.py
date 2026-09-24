#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
=====================================================================================
 STAPLES MARKETPLACE PoC  -  CATEGORY-LEVEL RECOMMENDATION  (v2, six retailers)
 METHOD G : GRAPH-FIRST  (category knowledge graph, structure-aware matching, PageRank)
=====================================================================================

Plain-language summary
----------------------
Method G treats the six navigation trees as ONE network: every category is a node, every
"sits under" link is an edge (Staples' cross-listings too - one shelf can sit under 41
parents), and every confident competitor-to-Staples match becomes a SAME_AS edge.
Two ideas that a vector search alone cannot use:
  * Structure-aware matching (similarity flooding): two shelves are more likely the same
    when their PARENTS and CHILDREN also match - "Standing Desks" under "Desks" beats a
    "Standing Desks" under "Mats".
  * Adjacency by network proximity (Personalised PageRank): start random walks on Staples'
    own shelves (weighted by assortment) and see how much of that "Staples gravity" reaches a
    competitor shelf through the category network. Whitespace that sits in an aisle Staples
    already occupies scores high; an island far from Staples scores low.

 1. LOAD        six trees (poc_common.load_all)
 2. GRAPH       CHILD_OF + CROSS_LISTED_UNDER edges (NetworkX; Neo4j export at the end)
 3. MATCH       wording + thesaurus similarity, refined by similarity flooding (3 rounds)
 4. CALIBRATE   per competitor on labelled rows (+ ablation negatives)
 5. SAME_AS     confident matches become edges; concepts = connected components
 6. ADJACENCY   Personalised PageRank from Staples shelves (+ sibling coverage)
 7. CANNIBAL.   SAME_AS-strength link to a Staples CORE shelf
 8. WRITE       outputs/G/node_scores_G.csv, calibration, Neo4j CSV + Cypher, GraphML

Run
---
 Colab : upload poc_common.py, this file, the six .xlsx trees and gold_matches_v2.csv, then
            !pip install -q networkx openpyxl
            !python method_g_graph.py
 Local : python method_g_graph.py
 Neo4j (optional): set NEO4J_URI / NEO4J_USER / NEO4J_PASSWORD to push the graph live, or
         load outputs/G/neo4j/*.csv with outputs/G/neo4j/load_graph.cypher.
"""
# %% [0] INSTALL (uncomment in Colab) ---------------------------------------------------
# !pip install -q pandas numpy scipy scikit-learn openpyxl networkx
# !pip install -q neo4j        # only to push the graph to a live Neo4j

# %% [1] IMPORTS ------------------------------------------------------------------------
import os
import time
from collections import defaultdict

import numpy as np
import pandas as pd

import poc_common as pc
from poc_common import CONFIG, FOCAL, log

OUT = "G"


# %% [2] GRAPH HELPERS ------------------------------------------------------------------
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


def similarity_flooding(S0, par_a, ch_a, par_b, ch_b):
    """sigma(a,b) starts as wording similarity S0 and absorbs structural evidence:
         parent term : how well a's parent matches (any of) b's parents
         child term  : for each child of a, its best match among b's children, averaged
    Terms are used only where both sides have them, so roots and leaves are not penalised.
    (Melnik, Garcia-Molina & Rahm 2002, adapted to taxonomies.)"""
    from scipy.sparse import csr_matrix
    na, nb = S0.shape
    lp, lc = CONFIG["lambda_parent"], CONFIG["lambda_child"]
    has_pa = np.array([len(p) > 0 for p in par_a])
    has_pb = np.array([len(p) > 0 for p in par_b])
    has_ca = np.array([len(c) > 0 for c in ch_a])
    has_cb = np.array([len(c) > 0 for c in ch_b])
    rows, cols, vals = [], [], []
    for a, kids in enumerate(ch_a):
        for k in kids:
            rows.append(a)
            cols.append(k)
            vals.append(1.0 / len(kids))
    Aavg = csr_matrix((vals, (rows, cols)), shape=(na, na))
    pa_first = np.array([p[0] if p else 0 for p in par_a])
    S = S0.copy()
    wp = (lp * np.outer(has_pa, has_pb)).astype(np.float32)
    wc = (lc * np.outer(has_ca, has_cb)).astype(np.float32)
    for _ in range(CONFIG["n_iter"]):
        Pa = S[pa_first, :]
        P = np.zeros_like(S)
        for b in np.where(has_pb)[0]:
            P[:, b] = Pa[:, par_b[b]].max(axis=1)
        R = np.zeros_like(S)
        for b in np.where(has_cb)[0]:
            R[:, b] = S[:, ch_b[b]].max(axis=1)
        C = Aavg @ R
        S = ((S0 + wp * P + wc * C) / (1 + wp + wc)).astype(np.float32)
    return S


def export_graph(G, nodes, out_dir):
    """Neo4j bulk-load CSVs + Cypher (load + example queries), and GraphML."""
    import networkx as nx
    nd = os.path.join(out_dir, "neo4j")
    os.makedirs(nd, exist_ok=True)
    n = nodes[["uid", "retailer", "name", "path_str", "depth", "subtree_items", "n_leaves", "is_leaf"]].copy()
    n["coreness"] = nodes.get("coreness", pd.Series(np.nan, index=nodes.index))
    n.columns = ["uid", "retailer", "name", "path", "level", "items", "n_leaves", "is_leaf", "coreness"]
    n.to_csv(os.path.join(nd, "nodes.csv"), index=False)
    e = pd.DataFrame([(u, v, d.get("rel", "LINK"), round(float(d.get("weight", 1.0)), 4))
                      for u, v, d in G.edges(data=True)], columns=["source", "target", "rel", "weight"])
    e.to_csv(os.path.join(nd, "edges.csv"), index=False)
    open(os.path.join(nd, "load_graph.cypher"), "w").write(
        """// ---- Staples marketplace PoC: six-retailer category knowledge graph ----
// Copy nodes.csv + edges.csv into Neo4j's import folder, then run:
CREATE CONSTRAINT cat_uid IF NOT EXISTS FOR (c:Category) REQUIRE c.uid IS UNIQUE;
LOAD CSV WITH HEADERS FROM 'file:///nodes.csv' AS r
MERGE (c:Category {uid: r.uid})
SET c.retailer = r.retailer, c.name = r.name, c.path = r.path, c.level = toInteger(r.level),
    c.items = toFloat(r.items), c.n_leaves = toInteger(r.n_leaves), c.is_leaf = (r.is_leaf = 'True'),
    c.coreness = toFloat(r.coreness)
MERGE (rt:Retailer {name: r.retailer})
MERGE (c)-[:SOLD_BY]->(rt);
// one LOAD per relationship type (no APOC needed)
LOAD CSV WITH HEADERS FROM 'file:///edges.csv' AS r WITH r WHERE r.rel = 'CHILD_OF'
MATCH (a:Category {uid: r.source}), (b:Category {uid: r.target}) MERGE (a)-[x:CHILD_OF]->(b) SET x.weight = toFloat(r.weight);
LOAD CSV WITH HEADERS FROM 'file:///edges.csv' AS r WITH r WHERE r.rel = 'CROSS_LISTED_UNDER'
MATCH (a:Category {uid: r.source}), (b:Category {uid: r.target}) MERGE (a)-[x:CROSS_LISTED_UNDER]->(b) SET x.weight = toFloat(r.weight);
LOAD CSV WITH HEADERS FROM 'file:///edges.csv' AS r WITH r WHERE r.rel = 'SAME_AS'
MATCH (a:Category {uid: r.source}), (b:Category {uid: r.target}) MERGE (a)-[x:SAME_AS]->(b) SET x.weight = toFloat(r.weight);
""")
    open(os.path.join(nd, "example_queries.cypher"), "w").write(
        """// 1. Consensus whitespace: competitor shelves with no Staples equivalent, carried by 3+ retailers
MATCH (c:Category) WHERE c.retailer <> 'Staples' AND NOT (c)-[:SAME_AS]->(:Category {retailer:'Staples'})
MATCH (c)-[:SIMILAR_TO]-(d:Category) WHERE d.retailer <> c.retailer AND d.retailer <> 'Staples'
WITH c, collect(DISTINCT d.retailer) + c.retailer AS carriers
WHERE size(carriers) >= 3 RETURN c.path, carriers ORDER BY size(carriers) DESC LIMIT 25;
// (SIMILAR_TO peer edges are written by run_framework.py into outputs/final/graph/peer_edges.csv)

// 2. Aisle whitespace: a missing shelf whose siblings Staples mostly carries
MATCH (c:Category)-[:CHILD_OF]->(p)<-[:CHILD_OF]-(sib)
WHERE c.retailer <> 'Staples' AND NOT (c)-[:SAME_AS]->(:Category {retailer:'Staples'})
WITH c, p, count(sib) AS n_sib,
     sum(CASE WHEN (sib)-[:SAME_AS]->(:Category {retailer:'Staples'}) THEN 1 ELSE 0 END) AS carried
WHERE n_sib >= 3 AND toFloat(carried)/n_sib >= 0.7
RETURN c.retailer, p.path AS aisle, c.name AS missing_shelf, carried, n_sib ORDER BY n_sib DESC LIMIT 25;

// 3. Cannibalisation check: competitor shelves pointing at a Staples CORE shelf
MATCH (c:Category)-[s:SAME_AS]->(t:Category {retailer:'Staples'}) WHERE t.coreness >= 0.7
RETURN c.retailer, c.path, t.path, s.weight, t.coreness ORDER BY t.coreness * s.weight DESC LIMIT 25;

// 4. Adjacency with Graph Data Science: PageRank seeded on Staples leaves
CALL gds.graph.project('cat', 'Category', {CHILD_OF:{orientation:'UNDIRECTED', properties:'weight'},
     CROSS_LISTED_UNDER:{orientation:'UNDIRECTED', properties:'weight'}, SAME_AS:{orientation:'UNDIRECTED', properties:'weight'}});
MATCH (s:Category {retailer:'Staples', is_leaf:true}) WITH collect(s) AS seeds
CALL gds.pageRank.stream('cat', {sourceNodes: seeds, relationshipWeightProperty:'weight'}) YIELD nodeId, score
WITH gds.util.asNode(nodeId) AS n, score WHERE n.retailer <> 'Staples'
RETURN n.retailer, n.path, score ORDER BY score DESC LIMIT 25;
""")
    H = nx.Graph()
    for u, d in G.nodes(data=True):
        H.add_node(u, **{k: (v if isinstance(v, (int, float, str)) else str(v)) for k, v in d.items()})
    for u, v, d in G.edges(data=True):
        H.add_edge(u, v, rel=d.get("rel", ""), weight=float(d.get("weight", 1.0)))
    nx.write_graphml(H, os.path.join(out_dir, "category_graph.graphml"))
    uri = os.environ.get("NEO4J_URI")
    if uri:
        try:
            from neo4j import GraphDatabase
            drv = GraphDatabase.driver(uri, auth=(os.environ.get("NEO4J_USER", "neo4j"),
                                                  os.environ.get("NEO4J_PASSWORD", "")))
            with drv.session() as ses:
                ses.run("UNWIND $rows AS r MERGE (c:Category {uid:r.uid}) SET c += r",
                        rows=n.where(pd.notna(n), None).to_dict("records"))
                for rel, g in e.groupby("rel"):
                    ses.run(f"UNWIND $rows AS r MATCH (a:Category {{uid:r.source}}),(b:Category {{uid:r.target}}) "
                            f"MERGE (a)-[x:{rel}]->(b) SET x.weight = r.weight", rows=g.to_dict("records"))
            log(f"      pushed {len(n)} nodes / {len(e)} edges to Neo4j at {uri}")
        except Exception as ex:  # noqa
            log(f"      Neo4j push skipped ({ex})")
    return e["rel"].value_counts().to_dict()


# %% [3] MAIN ---------------------------------------------------------------------------
def main():
    import networkx as nx
    t_start = time.time()
    log("\n[1] Loading the six navigation trees")
    nodes, xedges, totals, files = pc.load_all()
    gold = pc.load_gold(for_calibration=True)
    work = nodes[nodes.unit_ok].reset_index(drop=True)
    work["name_syn"] = work["name_expanded"].map(pc.synonymise)
    lex = pc.Lexical(sorted(set(work["name_syn"])))
    focal = work[work.retailer == FOCAL].reset_index(drop=True)
    focal["coreness"] = pc.staples_coreness(focal)
    work.loc[work.retailer == FOCAL, "coreness"] = focal["coreness"].values
    par_b, ch_b = build_hierarchy(focal, xedges.get(FOCAL))
    FL = lex.transform(focal["name_syn"])

    log("\n[2] Building the six-retailer category graph")
    G = nx.Graph()
    for r in work.itertuples(index=False):
        G.add_node(r.uid, retailer=r.retailer, name=r.name, level=int(r.depth), n_leaves=int(r.n_leaves))
    for u, p in zip(work["uid"], work["parent_uid"]):
        if p in G:
            G.add_edge(u, p, rel="CHILD_OF", weight=1.0)
    npl = defaultdict(int)
    for c, p, w in xedges.get(FOCAL, []):
        npl[c] += 1
    for c, p, w in xedges.get(FOCAL, []):
        if c in G and p in G and not G.has_edge(c, p):
            G.add_edge(c, p, rel="CROSS_LISTED_UNDER",
                       weight=(1.0 / npl[c]) if npl[c] <= CONFIG["xlist_weight_cap"] else 0.001)
    log(f"      {G.number_of_nodes():,} category nodes, {G.number_of_edges():,} edges (hierarchy + Staples cross-listings)")

    log("\n[3] Structure-aware matching (wording + thesaurus, similarity flooding) + calibration")
    results, pooled, calibs = [], [], {}
    for comp in pc.COMPETITORS:
        t0 = time.time()
        cn = work[work.retailer == comp].reset_index(drop=True)
        par_a, ch_a = build_hierarchy(cn)
        CL = lex.transform(cn["name_syn"])
        S0 = (CL @ FL.T).toarray().astype(np.float32)
        S = similarity_flooding(S0, par_a, ch_a, par_b, ch_b)
        order = np.argsort(-S, axis=1)[:, :CONFIG["k_retrieve"]]
        Ss = np.take_along_axis(S, order, axis=1)
        cn["best_score"] = Ss[:, 0]
        cn["best_lex"] = S0[np.arange(len(cn)), order[:, 0]]
        cn["best_staples"] = focal["path_str"].values[order[:, 0]]
        cn["best_staples_uid"] = focal["uid"].values[order[:, 0]]
        cn["alt_staples"] = [" | ".join(focal["path_str"].values[order[i, 1:4]]) for i in range(len(cn))]
        gc = gold[gold.competitor == comp]
        cal = None
        if len(gc) >= CONFIG["min_gold_rows"]:
            row_of = dict(zip(cn["path_str"], range(len(cn))))
            abl = {r["comp_path"]: float(S[row_of[r["comp_path"]]][pc.ablation_mask(
                focal, r["gold_staples_path"], r["comp_path"])].max())
                for _, r in gc[(gc.staples_carries == 1) & gc.gold_staples_path.notna()].iterrows()
                if r["comp_path"] in row_of}
            cal, cal_df = pc.calibrate(gc, dict(zip(cn["path_str"], cn["best_score"])), abl,
                                       dict(zip(cn["path_str"], cn["best_staples"])))
            # wording-only baseline: what the structure adds
            abl0 = {k: float(S0[row_of[k]][pc.ablation_mask(focal, g, k)].max())
                    for k, g in zip(gc["comp_path"], gc["gold_staples_path"]) if k in row_of and pd.notna(g)}
            m0, _ = pc.calibrate(gc, dict(zip(cn["path_str"], S0.max(axis=1))), abl0,
                                 dict(zip(cn["path_str"], focal["path_str"].values[S0.argmax(axis=1)])))
            cal["top1_wording_only"], cal["auc_wording_only"] = m0["top1"], m0["auc"]
            cal["source"] = "own labels"
            cal_df["competitor"] = comp
            pooled.append(cal_df)
        calibs[comp] = cal
        results.append([comp, cn, order, Ss])
        log(f"      {comp:<12} {len(cn):>5} shelves in {time.time() - t0:5.1f}s"
            + (f" | tau {cal['tau']:.3f}, AUC {cal['auc']:.3f} (wording only {cal['auc_wording_only']:.3f}), "
               f"top-1 {cal['top1']:.1%} (wording only {cal['top1_wording_only']:.1%})" if cal else " | no labels"))
        del S, S0

    pool = None
    if pooled:
        pdf = pd.concat(pooled)
        y, s = pdf["y"].values, pdf["score"].values
        cands = [(float(t), CONFIG["false_gap_cost"] * (s[y == 1] >= t).mean() - (s[y == 0] >= t).mean())
                 for t in np.unique(np.round(s, 4))]
        tau_p = max(cands, key=lambda z: z[1])[0]
        pos = pdf[(pdf.kind == "gold") & (pdf.y == 1)]["score"]
        pool = {"tau": tau_p, "s_enter": min(float(pos.quantile(CONFIG["gap_screen_false_alarm"])), tau_p),
                "neg_median": float(pdf.loc[pdf.y == 0, "score"].median()), "source": "pooled labels"}

    log("\n[4] SAME_AS edges, cannibalisation, Personalised PageRank adjacency")
    same_as = []
    for k, (comp, cn, order, Ss) in enumerate(results):
        cal = calibs[comp] or pc.fallback_calibration(cn["best_score"].values, pool)
        calibs[comp] = cal
        tau, s_enter, lo = cal["tau"], cal["s_enter"], cal["neg_median"]
        cn["status"] = pc.status_from_scores(cn["best_score"].values, tau, s_enter)
        m = CONFIG["multi_match_margin"]
        cn["match_set"] = [" | ".join(focal["uid"].values[order[i, j]] for j in range(order.shape[1])
                                      if Ss[i, j] >= tau and Ss[i, j] >= Ss[i, 0] - m) for i in range(len(cn))]
        for u, ms, sc in zip(cn["uid"], cn["match_set"], cn["best_score"]):
            for v in [x for x in ms.split(" | ") if x]:
                G.add_edge(u, v, rel="SAME_AS", weight=float(sc))
                same_as.append((u, v))
        f = np.clip((Ss - lo) / max(tau - lo, 1e-6), 0, 1)
        core = focal["coreness"].values[order]
        cn["crs_raw"] = (f * core).max(axis=1)
        cn["crs_driver"] = [focal["path_str"].values[order[i, np.argmax(f[i] * core[i])]] for i in range(len(cn))]
        cn["CRS_leaf"] = 100 * cn["crs_raw"]
        results[k][1] = cn
    log(f"      {len(same_as):,} SAME_AS edges")

    seeds = {u: float(np.log1p(c)) for u, c, lf in zip(focal["uid"], focal["subtree_items"], focal["is_leaf"])
             if lf and c > 0 and u in G}
    pr_seed = nx.pagerank(G, alpha=CONFIG["ppr_alpha"], personalization=seeds, weight="weight", max_iter=200)
    pr_glob = nx.pagerank(G, alpha=CONFIG["ppr_alpha"], weight="weight", max_iter=200)
    H = nx.Graph()
    H.add_nodes_from(G.nodes)
    H.add_edges_from(same_as)
    concept = {u: i for i, comp_ in enumerate(nx.connected_components(H)) for u in comp_}

    all_rows, cal_rows = [], []
    eps = 1e-12
    for res in results:
        res[1]["A_raw"] = [np.log((pr_seed.get(u, 0) + eps) / (pr_glob.get(u, 0) + eps)) for u in res[1]["uid"]]
    allc = pd.concat([r[1] for r in results])
    lf = allc["is_leaf"].values.astype(bool)
    scale = pc.aas_scaler(allc["A_raw"].values[lf], allc["status"].values[lf] == "matched")
    for comp, cn, order, Ss in results:
        cn["AAS_leaf"] = scale(cn["A_raw"].values)
        cn = pc.rollup_competitor(cn)
        # sibling coverage: share of the parent aisle's leaves Staples carries (excluding self)
        tot = dict(zip(cn["uid"], cn["n_leaves"].clip(lower=1)))
        carried = dict(zip(cn["uid"], cn["coverage_leaf"] * cn["n_leaves"].clip(lower=1)))
        sib = []
        for u, p in zip(cn["uid"], cn["parent_uid"]):
            if p in tot and tot[p] - tot[u] > 0:
                sib.append((carried[p] - carried[u]) / (tot[p] - tot[u]))
            else:
                sib.append(np.nan)
        cn["sibling_coverage"] = np.clip(sib, 0, 1)
        w = CONFIG["aas_ppr_weight"]
        cn["AAS_ppr"] = cn["AAS"]
        cn["AAS"] = np.where(cn["sibling_coverage"].notna(),
                             w * cn["AAS_ppr"] + (1 - w) * 100 * cn["sibling_coverage"].fillna(0), cn["AAS_ppr"])
        cn["concept_id"] = cn["uid"].map(concept)
        cn["competitor"], cn["method"] = comp, "G"
        cn["best_sem"] = np.nan
        keep = ["competitor", "method", "uid", "path_str", "sem_path", "depth", "is_leaf", "n_leaves",
                "subtree_items", "count_known", "best_staples", "best_staples_uid", "best_score", "best_sem",
                "best_lex", "alt_staples", "status", "match_set", "A_raw", "AAS_leaf", "CRS_leaf", "coverage",
                "coverage_leaf", "coverage_items", "AAS", "AAS_ppr", "sibling_coverage", "CRS", "crs_driver",
                "n_gap_leaves", "gap_leaves_example", "concept_id"]
        all_rows.append(cn[keep])
        cal = calibs[comp]
        cal_rows.append({"competitor": comp, "method": "G", "encoder": "wording + thesaurus + structure", **cal})
        st = cn["status"].value_counts(normalize=True)
        log(f"      {comp:<12} matched {st.get('matched', 0):.0%} | likely {st.get('likely', 0):.0%} | "
            f"gap {st.get('gap', 0):.0%}   ({cal.get('source', cal.get('note', ''))})")
    pd.concat(all_rows).to_csv(pc.out_path(OUT, "node_scores_G.csv"), index=False)
    pd.DataFrame(cal_rows).to_csv(pc.out_path(OUT, "calibration_G.csv"), index=False)
    if pooled:
        pd.concat(pooled).to_csv(pc.out_path(OUT, "calibration_rows_G.csv"), index=False)

    log("\n[5] Exporting the graph (Neo4j CSV + Cypher, GraphML)")
    wexp = work.merge(nodes[["uid"]], on="uid")
    rel_counts = export_graph(G, wexp, pc.out_path(OUT, ""))
    pd.DataFrame({"uid": list(pr_seed), "ppr_seed": list(pr_seed.values()),
                  "ppr_global": [pr_glob.get(u, 0) for u in pr_seed]}).to_csv(pc.out_path(OUT, "pagerank.csv"), index=False)
    pc.save_json({"n_nodes": G.number_of_nodes(), "n_edges": G.number_of_edges(), "edges_by_type": rel_counts,
                  "n_concepts": len(set(concept.values())), "seconds": round(time.time() - t_start, 1)},
                 OUT, "run_info_G.json")
    log(f"      edges by type: {rel_counts}")
    log(f"\nDone in {time.time() - t_start:.0f}s. Outputs in {pc.out_path(OUT, '')}")


if __name__ == "__main__":
    main()
