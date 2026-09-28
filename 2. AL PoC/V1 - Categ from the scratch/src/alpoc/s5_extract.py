"""S4/S5 Attribute schema + extraction + validation (methodology §6.1-6.2).

Instrument: backend 'rules' = controlled-vocabulary matchers + embedding zero-shot (local, no API).
Rules applied (§6.2.2):
  * physical facts (colour, material, Tier-2 with a spec key, dimensions): Staples = spec if present,
    else text; competitor = text.
  * style, Tier-3 lifestyle tags and DFI: the SAME text instrument on both sides.
  * G2: the text instrument is run on Staples too and scored against Staples specs; fields below the
    accuracy bar are marked 'descriptive only' and kept out of gap metrics and archetype grids.
"""
from __future__ import annotations

import json
import re

import numpy as np
import pandas as pd

from . import embed
from .common import cfg, l2_config, load, log, pct_of, save, save_json
from .s3_mapping import mapping_card
from .vocab import Matcher, colour_tone, matcher, universal

MULTI_ATTRS = ["aesthetic_tags", "use_context", "end_user_segment", "key_benefits"]
COLOUR_SPEC_KEYS = ["Color Family", "Furnishing Color", "True Color"]
DIM_SPEC = {"width": ["Width in Inches"], "depth": ["Depth in Inches"], "height": ["Height in Inches"],
            "capacity": ["Maximum Weight Capacity (Lbs)", "Capacity (lbs.)"]}
_U = r'(?:"|”|″|\'\'|in\.?|inches|-inch)?'
RX_TRIPLE = re.compile(rf"(\d+(?:\.\d+)?)\s*{_U}\s*W\s*x\s*(\d+(?:\.\d+)?)\s*{_U}\s*(?:D|L)\s*x?\s*(\d+(?:\.\d+)?)\s*{_U}\s*H\b", re.I)
RX_DIM = {
    "width": re.compile(rf"(\d+(?:\.\d+)?)\s*{_U}\s*(?:W\b|wide\b|width\b)|\bwidth\s*[:\-]?\s*(\d+(?:\.\d+)?)", re.I),
    "depth": re.compile(rf"(\d+(?:\.\d+)?)\s*{_U}\s*(?:D\b|deep\b|depth\b)|\bdepth\s*[:\-]?\s*(\d+(?:\.\d+)?)", re.I),
    "height": re.compile(rf"(\d+(?:\.\d+)?)\s*{_U}\s*(?:H\b|high\b|tall\b|height\b)|\bheight\s*[:\-]?\s*(\d+(?:\.\d+)?)", re.I),
}
RX_CAP = re.compile(r"(?:capacity|supports?|holds?|up to)[^.\d]{0,30}(\d{2,4})\s*(?:lbs?|pounds)\b", re.I)


def text_sources(r) -> list[tuple[str, str, float]]:
    if r["retailer"] == "staples":
        return [("title", r["title"], 0.9), ("bullets", r["bullets"], 0.75), ("desc", r["desc"], 0.6)]
    return [("title", r["title"], 0.9), ("choice", r["choice"], 0.85), ("desc", r["desc"], 0.6)]


def text_dims(text: str) -> dict:
    out = {}
    m = RX_TRIPLE.search(text)
    if m:
        out = {"width": float(m.group(1)), "depth": float(m.group(2)), "height": float(m.group(3))}
    for k, rx in RX_DIM.items():
        if k not in out:
            m = rx.search(text)
            if m:
                v = float(m.group(1) or m.group(2))
                if 3 <= v <= 150:
                    out[k] = v
    m = RX_CAP.search(text)
    if m:
        out["capacity"] = float(m.group(1))
    return out


def spec_number(specs: dict, keys: list[str]):
    for k in keys:
        if k in specs:
            m = re.search(r"\d+(?:\.\d+)?", specs[k].replace(",", ""))
            if m:
                return float(m.group(0))
    return np.nan


def spec_value(specs: dict, keys: list[str], m: Matcher, spec_map: dict | None = None):
    for k in keys or []:
        if k in specs:
            raw = specs[k]
            if spec_map and raw in spec_map:
                return spec_map[raw]
            v = m.first(raw)
            if v:
                return v
    return None


def _band(x, bands: dict):
    if x is None or np.isnan(x):
        return None
    for label, (lo, hi) in bands.items():
        if float(lo) <= x < float(hi):
            return label
    return None


def extract_l2(df: pd.DataFrame, l2: str) -> tuple[pd.DataFrame, dict]:
    c2 = l2_config(l2)
    ex = cfg()["extraction"]
    tier2 = {k: (Matcher(v["values"], v.get("mode", "priority")), v) for k, v in (c2.get("tier2") or {}).items()}
    mat_keys = c2.get("material_spec_keys") or ["Furnishing Material"]
    rows, val_rows = [], []
    for _, r in df.iterrows():
        specs = json.loads(r["specs"] or "{}")
        src = text_sources(r)
        is_st = r["retailer"] == "staples"
        o = {"family_id": r["family_id"]}

        # colour (physical): titles put the colour variant last, so read title segments from the end
        t_col, t_col_src = None, None
        for seg in reversed(re.sub(r"\([^()]*\)\s*$", "", r["title"]).split(",")):
            t_col = matcher("colour_family").first(seg)
            if t_col:
                t_col_src = "title"
                break
        if not t_col:
            t_col, t_col_src, _ = matcher("colour_family").from_sources(src[1:])
        s_col = spec_value(specs, COLOUR_SPEC_KEYS, matcher("colour_family")) if is_st else None
        o["colour_family"] = s_col or t_col
        o["colour_src"] = "spec" if s_col else t_col_src
        o["txt_colour_family"] = t_col
        # material (physical)
        t_mat, t_mat_src, _ = matcher("material_class").from_sources(src)
        s_mat = spec_value(specs, mat_keys, matcher("material_class")) if is_st else None
        o["material_class"] = s_mat or t_mat
        o["material_src"] = "spec" if s_mat else t_mat_src
        o["txt_material_class"] = t_mat
        # style (same instrument both sides; spec only as a validation indicator)
        t_sty, t_sty_src, _ = matcher("style_family").from_sources(src)
        o["style_family"], o["style_src"] = t_sty, t_sty_src
        # Tier 3 multi-label (same instrument)
        for a in MULTI_ATTRS:
            o[a] = matcher(a).from_sources(src)[0]
        # Tier 2 functional
        for a, (m, spec) in tier2.items():
            tv, tsrc, _ = m.from_sources(src)
            sv = spec_value(specs, spec.get("spec_keys"), m, spec.get("spec_map")) if is_st else None
            if set(m.values) == {"yes"}:
                # text can only ever find "yes": measure "feature mentioned" with the same instrument on
                # both sides (yes / no); the Staples spec is used for validation only
                o[a] = tv or "no"
                o["txt_" + a] = o[a]
                if is_st and sv:
                    val_rows.append((a, sv, o[a]))
                continue
            o[a] = sv or tv
            o["txt_" + a] = tv
            if is_st and sv and spec.get("spec_keys"):
                val_rows.append((a, sv, tv))
        # dimensions (physical)
        td = text_dims(" ".join(t for _, t, _ in src))
        for d, keys in DIM_SPEC.items():
            sv = spec_number(specs, keys) if is_st else np.nan
            o[d] = sv if not np.isnan(sv) else td.get(d, np.nan)
            o["txt_" + d] = td.get(d, np.nan)
            if is_st and not np.isnan(sv) and d != "capacity":
                val_rows.append((d, sv, td.get(d)))
        if is_st:
            val_rows.append(("colour_family", s_col, t_col))
            val_rows.append(("material_class", s_mat, t_mat))
            sty_spec = universal()["style_family"]["spec_style_map"].get(specs.get("Furnishing Style"))
            val_rows.append(("style_family (indicator)", sty_spec, t_sty))
        rows.append(o)
    att = pd.DataFrame(rows)
    for pre in ("", "txt_"):
        att[pre + "colour_tone"] = att[pre + "colour_family"].map(colour_tone)
        for band_name, spec in (c2.get("numeric_bands") or {}).items():
            att[pre + band_name] = att[pre + spec["source"]].map(lambda x: _band(x, spec["bands"]))
        att[pre + "size_class"] = att[pre + "width"].map(lambda x: _band(x, universal()["size_class_by_width"]))

    # style zero-shot for families without a style keyword (same instrument both sides)
    cards = [mapping_card(r) for _, r in df.iterrows()]
    E = embed.encode(cards)
    if ex["style_zero_shot"]:
        protos = universal()["style_family"]["zero_shot_prototypes"]
        P = embed.encode(list(protos.values()))
        sims = E @ P.T
        order = np.argsort(-sims, axis=1)
        names = list(protos.keys())
        for i in np.where(att["style_family"].isna())[0]:
            a, b = order[i, 0], order[i, 1]
            if sims[i, a] - sims[i, b] >= ex["style_zero_shot_margin"]:
                att.at[i, "style_family"] = names[a]
                att.at[i, "style_src"] = "zero-shot"
    # Design-Forward Index: anchor contrast, then pooled percentile within the L2 (both retailers)
    anchors = c2.get("dfi_anchors") or universal()["generic_dfi_anchors"]
    D = embed.encode(anchors["design"]).mean(axis=0)
    U = embed.encode(anchors["utility"]).mean(axis=0)
    # scored on the source-neutral card (extracted fields only), so retailer copywriting style cannot move it
    from .cards import neutral_card, tier2_attrs
    tmp = att.assign(node_id=df["node_id"].values)
    En = embed.encode([neutral_card(r, tier2_attrs(l2)) for _, r in tmp.iterrows()])
    att["dfi_raw"] = En @ D - En @ U
    att["dfi"] = pct_of(att["dfi_raw"].values, att["dfi_raw"].values)

    # validation table (Staples text instrument vs Staples spec)
    val = {}
    for field in sorted({v[0] for v in val_rows}):
        pairs = [(s, t) for f, s, t in val_rows if f == field and s is not None and not (isinstance(s, float) and np.isnan(s))]
        if not pairs:
            continue
        known = [(s, t) for s, t in pairs if t is not None and not (isinstance(t, float) and np.isnan(t))]
        if field in ("width", "depth", "height"):
            acc = np.mean([abs(t - s) <= 0.05 * max(s, 1) + 0.5 for s, t in known]) if known else np.nan
        else:
            acc = np.mean([s == t for s, t in known]) if known else np.nan
        val[field] = {"n_spec": len(pairs), "text_coverage": len(known) / len(pairs),
                      "accuracy_when_found": float(acc) if known else None}
    return att, val


def run() -> dict:
    nodes = load("nodes.parquet")
    fam = load("families.parquet")
    sn = load("staples_node.parquet")
    mp = load("mapping.parquet")
    active_l2 = sorted(nodes.loc[nodes["status"].isin(["scored", "thin"]), "l2_key"].unique())
    node_l2 = nodes.set_index("node_id")["l2_key"]
    assign = pd.concat([sn, mp.loc[mp["status"] == "mapped", ["family_id", "node_id"]]])
    assign["l2_key"] = assign["node_id"].map(node_l2)
    df = fam.merge(assign, on="family_id", how="inner", suffixes=("_fam", ""))
    df = df[df["l2_key"].isin(active_l2)].reset_index(drop=True)
    g2_min = cfg()["qa"]["g2_field_acc"]
    frames, qa = [], {}
    for l2, g in df.groupby("l2_key"):
        g = g.reset_index(drop=True)
        att, val = extract_l2(g, l2)
        att = g[["family_id", "retailer", "node_id", "l2_key", "price"]].merge(att, on="family_id")
        passing = [f for f, v in val.items() if v["accuracy_when_found"] is not None
                   and v["accuracy_when_found"] >= g2_min and "indicator" not in f]
        failing = [f for f, v in val.items() if v["accuracy_when_found"] is not None
                   and v["accuracy_when_found"] < g2_min and "indicator" not in f]
        qa[l2] = {"validation": val, "g2_pass": passing, "g2_fail": failing,
                  "provisional_config": not bool(l2_config(l2)),
                  "unknown_share": {a: {r: float(att.loc[att["retailer"] == r, a].isna().mean())
                                         for r in att["retailer"].unique()}
                                    for a in ["colour_family", "material_class", "style_family"]
                                    + list((l2_config(l2).get("tier2") or {}).keys())}}
        frames.append(att)
        log(f"S5 {l2}: {len(att)} families; G2 pass {passing}; fail {failing}")
    out = pd.concat(frames, ignore_index=True)
    for a in MULTI_ATTRS:
        out[a] = out[a].fillna("")
    save(out, "attributes.parquet")
    save_json({"backend": cfg()["extraction"]["backend"], "l2": qa}, "qa_extraction.json")
    return qa
