#!/usr/bin/env python3
"""
Helper for ADDING A COMPETITOR: draw a labelling sample so its matches can be calibrated.

After method_v_vector.py has run once, this writes outputs/labelling/labelling_sample_<Comp>.csv:
  * 10 shelves from each decile of Method V's best score (a "stratified sample" - its share of
    carried rows estimates how much of the competitor Staples already covers), plus
  * 20 extra shelves from the bottom two deciles ("gap-enriched" - where real gaps live).
Each row lists 10 candidate Staples shelves (vector top-10 + wording top-5, de-duplicated).
A reviewer fills in:
  staples_carries   1 = Staples has a category page for these products (same shelf, or a
                        broader / narrower shelf whose main assortment is these products); 0 = not
  gold_staples_path the best Staples path(s), several separated by "|" (parent/child also accepted)
then appends the rows to gold_matches_v2.csv and re-runs the three scripts.

Usage:  python label_sample.py WestElm Wayfair Amazon Walmart
"""
import sys

import numpy as np
import pandas as pd

import poc_common as pc
from poc_common import FOCAL


def main(comps):
    nodes, *_ = pc.load_all(verbose=False)
    vec = np.load(pc.out_path("V", "vectors_V.npz"), allow_pickle=True)
    pos = {u: i for i, u in enumerate(vec["uid"])}
    V = vec["V"]
    scores = pd.read_csv(pc.out_path("V", "node_scores_V.csv"))
    focal = nodes[(nodes.retailer == FOCAL) & nodes.unit_ok].reset_index(drop=True)
    FV = V[[pos[u] for u in focal["uid"]]]
    lex = pc.Lexical(sorted(set(nodes.loc[nodes.in_scope, "name_expanded"])))
    FL = lex.transform(focal["name_expanded"])
    rng = np.random.default_rng(pc.CONFIG["seed"])
    for comp in comps:
        s = scores[scores.competitor == comp].copy()
        s["decile"] = pd.qcut(s["best_score"].rank(method="first"), 10, labels=False)
        picks = []
        for d, g in s.groupby("decile"):
            picks.append(g.sample(min(10, len(g)), random_state=int(rng.integers(1e6))).assign(source="stratified sample"))
        rest = s[~s.uid.isin(pd.concat(picks).uid) & (s.decile <= 1)]
        picks.append(rest.sample(min(20, len(rest)), random_state=7).assign(source="gap-enriched sample"))
        smp = pd.concat(picks).reset_index(drop=True)
        cn = nodes.set_index("uid").loc[smp["uid"]].reset_index()
        Q = V[[pos[u] for u in cn["uid"]]]
        top_v = np.argsort(-(Q @ FV.T), axis=1)[:, :10]
        top_l = np.argsort(-(lex.transform(cn["name_expanded"]) @ FL.T).toarray(), axis=1)[:, :5]
        cands = []
        for i in range(len(cn)):
            seen = []
            for j in list(top_v[i]) + list(top_l[i]):
                p = focal["path_str"].values[j]
                if p not in seen:
                    seen.append(p)
            cands.append(seen[:12])
        out = pd.DataFrame({"competitor": comp, "comp_path": smp["path_str"], "context_path": cn["sem_path"],
                            "is_leaf": smp["is_leaf"], "source": smp["source"],
                            "candidates": [" || ".join(c) for c in cands],
                            "staples_carries": "", "gold_staples_path": "", "note": ""})
        f = pc.out_path("labelling", f"labelling_sample_{comp}.csv")
        out.to_csv(f, index=False)
        print(f"{comp}: {len(out)} rows -> {f}")


if __name__ == "__main__":
    main(sys.argv[1:] or [c for c in pc.COMPETITORS if c != "OfficeDepot"])
