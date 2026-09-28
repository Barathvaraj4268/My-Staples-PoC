"""S11 Outputs: PNG figures, CSV tables, and one self-contained static HTML report (§9)."""
from __future__ import annotations

import base64
import datetime as dt
import json

import numpy as np
import pandas as pd
from jinja2 import Environment, FileSystemLoader

from . import figures as F
from .common import ROOT, cfg, load, load_json, log, money, out, slug
from .s7_vector import APPROVE

COMP_LABEL = {"wayfair": "Wayfair"}


def _b64(p) -> str | None:
    if p is None:
        return None
    return "data:image/png;base64," + base64.b64encode(open(p, "rb").read()).decode()


def _pct(x, d=0):
    return "–" if x is None or (isinstance(x, float) and np.isnan(x)) else f"{x:.{d}%}"


def _num(x, d=0):
    return "–" if x is None or (isinstance(x, float) and np.isnan(x)) else f"{x:,.{d}f}"


def _money(x):
    return "–" if x is None or (isinstance(x, float) and np.isnan(x)) else money(x)


def node_display(node_id: str) -> str:
    return node_id.replace(" > ", " → ")


def gate(status: str, value: str, target: str, name: str, metric: str, note: str = "") -> dict:
    return {"name": name, "metric": metric, "value": value, "target": target, "status": status, "note": note}


def build_gates(qm, qx, qa6, qv, qi) -> list[dict]:
    q = cfg()["qa"]
    gs = []
    la, nb = qm["staples_loo_knn"], qm["none_balanced_acc"]
    ok1 = la >= q["g1_leaf_acc"] and nb >= q["g1_none_bal"]
    gs.append(gate("PASS" if ok1 else "FAIL",
                   f"leaf k-NN {la:.1%}; NONE bal. acc {nb:.1%} (in-scope kept {qm['in_scope_recall']:.0%}, out-of-scope rejected {qm['out_scope_rejection']:.0%})",
                   f"≥ {q['g1_leaf_acc']:.0%}; ≥ {q['g1_none_bal']:.0%}", "G1 Mapping",
                   "Leaf classifier (Staples leave-one-out) and the NONE decision on known in/out-of-scope products",
                   "Human-labelled gold set (200 families) still pending: this is a proxy."))
    fails = {l2: v["g2_fail"] for l2, v in qx["l2"].items()}
    passes = {l2: v["g2_pass"] for l2, v in qx["l2"].items()}
    n_f = sum(len(v) for v in fails.values())
    gs.append(gate("PASS" if n_f == 0 else "PARTIAL", f"{sum(len(v) for v in passes.values())} fields pass, {n_f} fail",
                   f"≥ {q['g2_field_acc']:.0%} per field", "G2 Extraction (physical facts)",
                   "Text instrument vs Staples specs, accuracy when found",
                   "Failing fields: " + "; ".join(f"{l2.split(' > ')[-1]}: {', '.join(v) or 'none'}" for l2, v in fails.items())
                   + ". Failing fields are kept out of gap scores and archetype grids."))
    gs.append(gate("PENDING", "–", "κ ≥ 0.6", "G3 Lifestyle tags & DFI", "Two-rater human labels",
                   "Needs 150 human-labelled families. Style, vibe tags and DFI are provisional until then."))
    gs.append(gate("PENDING", "audit files written", "precision ≥ 95%", "G4 Families & dedup",
                   "Manual audit of 100 family merges and identical-product pairs",
                   "Audit samples: outputs/tables/audit_*.csv"))
    g5 = [v["g5"] for k, v in qv.items() if isinstance(v, dict) and "g5" in v]
    nn_ = [x["neutral"] for x in g5 if x.get("neutral") is not None]
    tt_ = [x["title_card"] for x in g5 if x.get("title_card") is not None]
    a5, b5 = (float(np.mean(nn_)) if nn_ else None), (float(np.mean(tt_)) if tt_ else None)
    gs.append(gate("PASS" if a5 is not None and a5 <= q["g5_max_source_auc"] else ("FAIL" if a5 is not None else "PENDING"),
                   f"source AUC {_num(a5, 2)} (title-bearing card: {_num(b5, 2)})", f"≤ {q['g5_max_source_auc']}", "G5 Embedding",
                   "How well a classifier can tell the retailer from the card (0.5 = indistinguishable)",
                   "Some separability is genuine assortment difference; the title-bearing card shows how much copy style it removed."))
    aris = [v["ari"] for v in qa6.values() if v.get("ari") is not None]
    stabs = [v["stability"] for v in qa6.values() if v.get("stability") is not None]
    ma, ms = (float(np.median(aris)) if aris else None), (float(np.median(stabs)) if stabs else None)
    st6 = "PASS" if (ma is not None and ma >= q["g6_ari"] and ms is not None and ms >= q["g6_stability"]) else "FAIL"
    gs.append(gate(st6, f"median ARI {_num(ma, 2)}, stability {_num(ms, 2)}", f"ARI ≥ {q['g6_ari']}, stab ≥ {q['g6_stability']}",
                   "G6 Archetypes", "Facet grid vs HDBSCAN clusters (UMAP of card embeddings)",
                   "Low ARI means clusters split on something the grid does not use (often title wording); nameability test pending."))
    st = qi["sensitivity"]
    tgs, vos = st["tg"]["mean_topn_retention"], st["vos"]["mean_topn_retention"]
    aucs = [v["calibration"].get("auc") for k, v in qv.items() if isinstance(v, dict) and "calibration" in v]
    aucs = [a for a in aucs if a is not None]
    ok7 = all(x is not None and x >= q["g7_topn_stability"] for x in (tgs, vos)) and aucs and min(aucs) >= q["g7_auc"]
    gs.append(gate("PASS" if ok7 else "FAIL", f"top-5 kept: TG {_pct(tgs)}, VOS {_pct(vos)}; CRS AUC {', '.join(f'{a:.2f}' for a in aucs)}",
                   f"≥ {q['g7_topn_stability']:.0%}; AUC ≥ {q['g7_auc']}", "G7 Scores",
                   "Weight sensitivity (Dirichlet) and weak-supervision calibration AUC",
                   "AUC is on Staples-only weak pairs; Pat's calibration session replaces it."))
    de = qv["data_error_share"]
    gs.append(gate("PASS" if de <= q["g8_data_error"] else "FAIL", f"DATA-ERROR share {de:.1%}",
                   f"≤ {q['g8_data_error']:.0%}", "G8 Sanity", "Low fit + high cannibalisation (should be rare)",
                   "Merchant spot-check of 20 recommendations pending."))
    return gs


def node_insights(nd, ns_row, bands, gaps, fa, cand, se, vend, comp) -> list[str]:
    out_ = []
    out_.append(f"{comp} sample: {nd['n_competitor']:,} product families here vs {nd['n_staples']:,} at Staples. "
                f"Median price {_money(ns_row['price_median_competitor'])} vs {_money(ns_row['price_median_staples'])}.")
    b = bands[(bands["credibility"] >= 0.9)].assign(d=lambda x: x["share_competitor"] - x["share_staples"])
    b = b[b["d"] > 0.03]
    if len(b):
        r = b.nlargest(1, "d").iloc[0]
        out_.append(f"Price gap: {r['share_competitor']:.0%} of {comp}'s range sits in {r['band']} vs "
                    f"{r['share_staples']:.0%} of Staples' (credible).")
    g = gaps[(gaps["credibility"] >= 0.9) & (~gaps["descriptive_only"]) & (gaps["delta"] > 0.04)]
    for r in g.nlargest(3, "delta").itertuples():
        out_.append(f"{r.attribute.replace('_', ' ').capitalize()} “{r.value}”: {r.share_competitor:.0%} of {comp} "
                    f"vs {r.share_staples:.0%} of Staples (credible gap).")
    s = gaps[(gaps["credibility"] <= 0.1) & (~gaps["descriptive_only"]) & (gaps["delta"] < -0.08)]
    if len(s):
        r = s.nsmallest(1, "delta").iloc[0]
        out_.append(f"Staples is deeper in {r['attribute'].replace('_', ' ')} “{r['value']}” "
                    f"({r['share_staples']:.0%} vs {r['share_competitor']:.0%}): not a gap.")
    if not np.isnan(ns_row["design_forward_share_competitor"]):
        out_.append(f"Design-forward share (DFI ≥ 0.6): {comp} {ns_row['design_forward_share_competitor']:.0%} vs "
                    f"Staples {ns_row['design_forward_share_staples']:.0%}.")
    if fa is not None and len(fa):
        sl = fa[fa["shortlisted"]].sort_values("final_rank")
        if len(sl):
            r = sl.iloc[0]
            out_.append(f"Top recommendation: {r['name']}. {r['tier']} agreement between the two methods; "
                        f"{r['safe_share']:.0%} of its {comp} products are safe to add.")
        else:
            out_.append("No archetype passed the safety gate and the agreement bar here; see the excluded list below.")
    if cand is not None and len(cand):
        lc = cand["label"].value_counts()
        out_.append(f"Across {len(cand):,} {comp} products: {lc.reindex(APPROVE).fillna(0).sum():.0f} approvable "
                    f"({lc.get('CURATE', 0)} CURATE, {lc.get('STYLE-EXTENSION', 0)} STYLE-EXTENSION, {lc.get('TRADE-UP', 0)} TRADE-UP); "
                    f"{lc.get('SUBSTITUTE', 0) + lc.get('UNDERCUT', 0)} would substitute or undercut a Staples item.")
    if vend is not None and len(vend):
        k = int(vend["brand_on_staples"].sum())
        if k:
            out_.append(f"{k} brand(s) behind recommended products already sell on Staples: quickest recruits.")
    return out_


def run() -> dict:
    v = cfg()["vector"]
    nodes = load("nodes.parquet")
    fa = load("final_archetypes.parquet")
    cand = load("candidates.parquet")
    gaps = load("attr_gaps.parquet")
    bands = load("price_bands.parquet")
    nsum = load("node_summary.parquet").set_index("node_id")
    recs = load("sku_recs.parquet")
    se = load("style_extensions.parquet")
    vend = load("vendor_view.parquet")
    mp = load("mapping.parquet")
    ft = load("family_table.parquet")
    fam = load("families.parquet")
    qm, qx, qa6 = load_json("qa_mapping.json"), load_json("qa_extraction.json"), load_json("qa_archetypes.json")
    qv, qi = load_json("qa_vector.json"), load_json("qa_integration.json")
    comp_key = [r for r in ft["retailer"].unique() if r != "staples"][0]
    comp = COMP_LABEL.get(comp_key, comp_key.title())

    # ---- tables (CSV) + audit samples
    for name, df in {"final_archetypes": fa, "candidates": cand, "attribute_gaps": gaps, "price_bands": bands,
                     "sku_recommendations": recs, "style_extensions": se, "vendor_view": vend, "nodes": nodes}.items():
        df.to_csv(out("tables", f"{name}.csv"), index=False)
    backlog = mp[mp["status"] != "mapped"].merge(fam[["family_id", "title", "price", "url", "brand"]], on="family_id")
    backlog.to_csv(out("tables", "new_node_backlog.csv"), index=False)
    st = fam[(fam["retailer"] == "staples") & (fam["n_skus"] > 1)].sample(100, random_state=1)
    st[["family_id", "title", "n_skus", "sku_ids", "colours_observed", "leaf_path"]].to_csv(out("tables", "audit_family_merges.csv"), index=False)
    ident = cand[cand["identical"]].merge(ft[["family_id", "title", "price"]], on="family_id") \
        .merge(ft[["family_id", "title", "price"]].add_prefix("st_"), left_on="identical_staples", right_on="st_family_id")
    ident.head(100).to_csv(out("tables", "audit_identical_pairs.csv"), index=False)

    # ---- overview figures
    ov = {}
    ov["coverage"] = _b64(F.node_coverage(nodes, comp, out("figures", "overview", "node_coverage.png")))
    for l2, q in qx["l2"].items():
        ov[f"acc_{slug(l2)}"] = _b64(F.extractor_accuracy(q["validation"], cfg()["qa"]["g2_field_acc"],
                                                          out("figures", "overview", f"extractor_accuracy_{slug(l2)}.png")))

    # ---- per-node payloads
    node_list, sections = [], []
    order = nodes[nodes["status"].isin(["scored", "thin"])].sort_values("node_id")
    for _, nd in order.iterrows():
        nid = nd["node_id"]
        sid = slug(nid)
        if len(sid) > 70:                                   # keep it readable, unique via a short hash
            from .common import hash_text
            sid = sid[:60] + "_" + hash_text(nid)[:6]
        d = out("figures", sid, "x").parent
        nb = bands[bands["node_id"] == nid].sort_values("price_mid")
        ng = gaps[gaps["node_id"] == nid]
        nft = ft[ft["node_id"] == nid]
        figs = {
            "price": _b64(F.price_bands(nb, comp, d / "price_bands.png")) if len(nb) else None,
            "dfi": _b64(F.dfi_density(nft.loc[nft["retailer"] == "staples", "dfi"], nft.loc[nft["retailer"] != "staples", "dfi"],
                                      comp, d / "dfi_density.png")),
            "jsd": _b64(F.attribute_jsd(ng, d / "attribute_jsd.png")) if len(ng) else None,
            "gaps": _b64(F.value_gaps(ng, comp, d / "value_gaps.png")) if len(ng) else None,
        }
        nfa = fa[fa["node_id"] == nid].copy()
        ncand = cand[cand["node_id"] == nid]
        nvend = vend[vend["node_id"] == nid]
        scored = nd["status"] == "scored" and len(nfa) > 0
        if scored:
            T = qv[nd["l2_key"]]["thresholds"]
            figs["decision"] = _b64(F.decision_scatter(ncand, T, v, d / "decision_scatter.png"))
            figs["agreement"] = _b64(F.agreement(nfa, d / "agreement.png"))
            figs["labelmix"] = _b64(F.label_mix(nfa, d / "label_mix.png"))
        ns_row = nsum.loc[nid]
        ins = node_insights(nd, ns_row, nb, ng, nfa if scored else None, ncand if scored else None, se, nvend, comp)
        # archetype tables
        nfa = nfa.sort_values(["final", "tg"], ascending=False)
        def facets_str(s):
            return ", ".join(f"{k.replace('_', ' ')}: {v}" for k, v in json.loads(s).items()) or "–"
        m1 = [{"name": r.name, "def": facets_str(r.facets), "ns": r.n_staples, "nc": r.n_competitor, "vw": _num(r.vw, 2),
               "aas": _num(r.aas), "ad": _num(r.ad, 2), "crs": _num(r.crs), "ppr": _num(r.ppr, 2),
               "safe": _pct(r.safe_share), "vos": _num(r.vos, 1), "elig": "yes" if r.eligible else "no"} for r in nfa.itertuples()]
        m2 = [{"name": r.name, "ps": _pct(r.p_staples), "pc": _pct(r.p_competitor), "lsr": _num(r.lsr, 2),
               "cred": _pct(r.lsr_credibility), "flag": "absent at Staples" if r.absent else ("thin at Staples" if r.thin else ""),
               "ppg": (_num(r.ppg, 2) + (" ↑" if r.ppg_dir > 0 else " ↓" if r.ppg_dir < 0 else "")) if not pd.isna(r.ppg) else "–",
               "cg": _num(r.cg, 2), "msg": _num(r.msg, 2), "dfg": _num(r.dfg, 2), "tg": _num(r.tg, 1),
               "tgo": _num(r.tg_original, 1), "miss": (r.missing_colours or "").replace("|", ", ")} for r in nfa.itertuples()]
        fin = [{"rank": int(r.final_rank) if not pd.isna(r.final_rank) else "", "name": r.name, "tier": r.tier,
                "vos": _num(r.vos, 1), "tg": _num(r.tg, 1), "final": _num(r.final, 1),
                "stab": f"{_pct(r.stability_tg)} / {_pct(r.stability_vos)}", "short": bool(r.shortlisted),
                "def": facets_str(r.facets)} for r in nfa[nfa["eligible"]].sort_values("final_rank").itertuples()]
        excl = [{"name": r.name, "reason": r.exclusion_reason, "tier": r.tier, "vos": _num(r.vos, 1), "tg": _num(r.tg, 1)}
                for r in nfa[~nfa["eligible"]].itertuples()]
        # SKU cards grouped by shortlisted archetype
        nrec = recs[recs["node_id"] == nid] if len(recs) else recs
        sku_groups = []
        for r in nfa[nfa["shortlisted"]].sort_values("final_rank").itertuples():
            rows = nrec[nrec["archetype_id"] == r.archetype_id].sort_values("exemplar_rank")
            sku_groups.append({"rank": int(r.final_rank), "name": r.name, "tier": r.tier, "cards": [_card(x, comp) for x in rows.itertuples()]})
        nse = se[se["node_id"] == nid] if len(se) else se
        se_cards = [_card(x, comp) for x in nse.itertuples()]
        vend_rows = [{"brand": r.brand, "products": r.products, "arch": r.archetypes, "path": r.recruit_path,
                      "on": r.brand_on_staples, "house": r.house_brand} for r in nvend.head(25).itertuples()]
        gap_rows = [{"attr": r.attribute.replace("_", " "), "val": r.value, "sc": _pct(r.share_competitor), "ss": _pct(r.share_staples),
                     "delta": f"{r.delta:+.0%}", "cred": _pct(r.credibility), "desc": r.descriptive_only}
                    for r in ng[(ng["credibility"] >= 0.9) | (ng["credibility"] <= 0.1)].assign(a=lambda x: x["delta"].abs())
                    .nlargest(30, "a").itertuples()]
        lc = ncand["label"].value_counts().to_dict() if scored else {}
        node_list.append({"id": sid, "label": node_display(nid), "status": nd["status"], "l2": nd["l2_key"]})
        sections.append({
            "id": sid, "label": node_display(nid), "status": nd["status"], "l2": nd["l2_key"],
            "provisional": not nd["l2_config"], "ns": nd["n_staples"], "nc": nd["n_competitor"],
            "rho": qi["spearman_vos_tg"].get(nid), "insights": ins, "figs": figs,
            "facets": ", ".join(qa6.get(nid, {}).get("facets", [])).replace("_", " "),
            "m1": m1, "m2": m2, "final": fin, "excluded": excl, "skus": sku_groups, "style_ext": se_cards,
            "vendors": vend_rows, "gaps": gap_rows, "labels": lc,
            "lowconf": float(np.nan_to_num(mp.loc[(mp["node_id"] == nid) & (mp["status"] == "mapped"), "low_conf"].mean())),
        })

    # ---- overview numbers
    mapped = mp[mp["status"] == "mapped"]
    tiles = {
        "staples_fam": int(nodes.loc[nodes["status"].isin(["scored", "thin"]), "n_staples"].sum()),
        "comp_fam": int(len(fam[fam["retailer"] == comp_key])),
        "comp_mapped": int(len(mapped)),
        "scored": int((nodes["status"] == "scored").sum()),
        "thin": int((nodes["status"] == "thin").sum()),
        "backlog": int((mp["status"] != "mapped").sum()),
        "shortlisted": int(fa["shortlisted"].sum()),
        "exemplars": int(len(recs)),
    }
    bl = backlog.assign(page=backlog["pages"].str.split("|").str[0]).groupby(["page", "status"]).size().unstack(fill_value=0)
    backlog_rows = [{"page": p, "none_page": int(r.get("none_page", 0)), "none_far": int(r.get("none_far", 0))}
                    for p, r in bl.iterrows()]
    staples_only = nodes[nodes["status"] == "staples_only"]
    ctx = {
        "title": "Staples Assortment Gap PoC", "comp": comp, "date": dt.date.today().isoformat(),
        "encoder": qm["encoder"], "bakeoff": qm["bakeoff"], "backend": qx["backend"],
        "tiles": tiles, "gates": build_gates(qm, qx, qa6, qv, qi), "ov": ov, "l2s": list(qx["l2"].keys()),
        "nodes": node_list, "sections": sections, "backlog": backlog_rows,
        "staples_only": [{"node": node_display(r.node_id), "n": r.n_staples} for r in staples_only.sort_values("n_staples", ascending=False).head(25).itertuples()],
        "n_staples_only": len(staples_only),
        "cfg": cfg(), "label_counts": qv["label_counts"], "tier_counts": qi["tier_counts"],
        "calib": {l2: q["calibration"] for l2, q in qv.items() if isinstance(q, dict) and "calibration" in q},
        "thresholds": {l2: q["thresholds"] for l2, q in qv.items() if isinstance(q, dict) and "thresholds" in q},
    }
    env = Environment(loader=FileSystemLoader(str(ROOT / "src" / "alpoc" / "templates")), autoescape=True)
    html = env.get_template("report.html.j2").render(**ctx)
    p = out("report", "Staples_Assortment_Report.html")
    p.write_text(html, encoding="utf-8")
    log(f"S11 report: {p} ({p.stat().st_size / 1e6:.1f} MB), {len(sections)} node sections")
    return {"report": str(p), "size_mb": p.stat().st_size / 1e6}


def _s(v) -> str:
    return "–" if v is None or (isinstance(v, float) and np.isnan(v)) or v == "" else str(v)


def _card(x, comp) -> dict:
    return {"title": x.title, "url": x.url, "price": _money(x.price), "brand": x.brand, "house": bool(x.house_brand),
            "on": bool(x.brand_on_staples), "label": x.label, "colour": _s(x.colour_family),
            "material": _s(x.material_class), "style": _s(x.style_family), "dfi": _num(x.dfi, 2),
            "crs": _num(x.crs), "ppr": _num(x.ppr, 2), "ad": _num(x.ad, 2),
            "st_title": x.st_title, "st_url": x.st_url, "st_price": _money(x.st_price),
            "st_colour": _s(x.st_colour_family), "st_material": _s(x.st_material_class), "comp": comp}
