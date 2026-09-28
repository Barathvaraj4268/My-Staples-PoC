"""S2 Analysis-node definition: Staples leaf, split into pseudo-L5 by Staples' own type facet (§5.3).

define()   runs before mapping (Staples side only).
finalize() runs after mapping: folds thin children into 'Other', collapses weak splits, sets node status.
"""
from __future__ import annotations

import json
import re

import numpy as np
import pandas as pd

from .common import cfg, l2_config, load, log, save


def _entropy(counts: np.ndarray) -> float:
    p = counts / counts.sum()
    return float(-(p * np.log(p)).sum() / np.log(len(p))) if len(p) > 1 else 0.0


def choose_split(g: pd.DataFrame, c: dict, override=None):
    specs = g["specs"].map(json.loads)
    keys = sorted({k for d in specs for k in d if re.search(c["split_key_regex"], k)
                   and not re.search(c.get("split_key_exclude", "^$"), k)})
    if override is not None:
        keys = [override] if override else []
    best = None
    for k in keys:
        vals = specs.map(lambda d: d.get(k))
        cov = vals.notna().mean()
        vc = vals.value_counts()
        viable = vc[vc >= c["child_min_families"]]
        if cov < c["split_min_coverage"] or len(viable) < 2:
            continue
        score = cov * _entropy(vc.values.astype(float))
        if best is None or score > best[1]:
            best = (k, score, vals)
    return best


def define() -> pd.DataFrame:
    c = cfg()["nodes"]
    fam = load("families.parquet")
    st = fam[fam["retailer"] == "staples"]
    rows, assign = [], []
    for leaf, g in st.groupby("leaf_path"):
        l2 = g["l2_key"].iloc[0]
        override = (l2_config(l2).get("split_facet_overrides") or {}).get(g["leaf"].iloc[0])
        split = choose_split(g, c, override) if len(g) >= c["split_min_families"] else None
        if split is None:
            rows.append({"node_id": leaf, "leaf_path": leaf, "l2_key": l2, "split_key": "", "split_value": ""})
            assign += [(f, leaf) for f in g["family_id"]]
            continue
        key, _, vals = split
        vc = vals.value_counts()
        keep = set(vc[vc >= c["child_min_families"]].index)
        for v in sorted(keep) + ["Other"]:
            rows.append({"node_id": f"{leaf} > {v}", "leaf_path": leaf, "l2_key": l2, "split_key": key, "split_value": v})
        assign += [(f, f"{leaf} > {v if v in keep else 'Other'}") for f, v in zip(g["family_id"], vals)]
    nodes = pd.DataFrame(rows)
    save(nodes, "nodes_initial.parquet")
    save(pd.DataFrame(assign, columns=["family_id", "node_id"]), "staples_node_initial.parquet")
    log(f"S2 define: {nodes['leaf_path'].nunique()} leaves, {int((nodes['split_key'] != '').sum())} pseudo-L5 children")
    return nodes


def finalize() -> pd.DataFrame:
    """Uses the competitor node assignment written by S3 (mapping.parquet)."""
    c = cfg()["nodes"]
    nodes = load("nodes_initial.parquet")
    sa = load("staples_node_initial.parquet")
    mp = load("mapping.parquet")
    ca = mp.loc[mp["status"] == "mapped", ["family_id", "node_id"]]
    cnt_s = sa["node_id"].value_counts()
    cnt_c = ca["node_id"].value_counts()
    remap = {}
    for leaf, g in nodes[nodes["split_key"] != ""].groupby("leaf_path"):
        viable = [n for n in g["node_id"] if not n.endswith(" > Other")
                  and cnt_s.get(n, 0) >= c["child_min_families"] and cnt_c.get(n, 0) >= c["child_min_families"]]
        if len(viable) < 2:                              # split not supported by both sides -> collapse to leaf
            for n in g["node_id"]:
                remap[n] = leaf
        else:
            for n in g["node_id"]:
                if n not in viable:
                    remap[n] = f"{leaf} > Other"
    sa["node_id"] = sa["node_id"].replace(remap)
    ca = ca.assign(node_id=ca["node_id"].replace(remap))
    mp["node_id"] = mp["node_id"].replace(remap)
    save(sa, "staples_node.parquet")
    save(mp, "mapping.parquet")

    fam = load("families.parquet").set_index("family_id")
    ns = sa["node_id"].value_counts()
    nc = ca["node_id"].value_counts()
    out = []
    for nid in sorted(set(sa["node_id"]) | set(ca["node_id"])):
        leaf = nid if nid in set(nodes["leaf_path"]) else nid.rsplit(" > ", 1)[0]
        sub = nodes[nodes["leaf_path"] == leaf].iloc[0]
        n_s, n_c = int(ns.get(nid, 0)), int(nc.get(nid, 0))
        if n_c == 0:
            status = "staples_only"
        elif n_s >= c["min_families_scored"] and n_c >= c["min_families_scored"]:
            status = "scored"
        else:
            status = "thin"
        out.append({"node_id": nid, "leaf_path": leaf, "l2_key": sub["l2_key"],
                    "split_key": sub["split_key"] if nid != leaf else "",
                    "split_value": nid.rsplit(" > ", 1)[1] if nid != leaf else "",
                    "n_staples": n_s, "n_competitor": n_c, "status": status,
                    "l2_config": bool(l2_config(sub["l2_key"]))})
    nodes_final = pd.DataFrame(out)
    save(nodes_final, "nodes.parquet")
    log("S2 finalize: " + str(nodes_final["status"].value_counts().to_dict()))
    return nodes_final
