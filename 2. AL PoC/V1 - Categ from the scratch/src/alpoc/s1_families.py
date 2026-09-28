"""S1 Family grouping & dedup (methodology §5.2). The family is the counting unit everywhere."""
from __future__ import annotations

import json
import re
from collections import Counter

import numpy as np
import pandas as pd

from .common import hash_text, load, load_yaml, log, save
from .vocab import colour_family

COLOUR_SPEC_KEYS = ["Color Family", "Furnishing Color", "True Color"]


def _strip_model(title: str) -> str:
    return re.sub(r"\s*\([^()]*\)\s*$", "", title).strip()


def name_stem(title: str) -> str:
    """Title without trailing (MODEL) and without comma segments that name a colour/finish."""
    t = _strip_model(title)
    parts = [p.strip() for p in t.split(",")]
    keep = [parts[0]] + [p for p in parts[1:] if not colour_family(p)]
    return re.sub(r"\s+", " ", ", ".join(keep)).lower()


def staples_brands(titles: pd.Series) -> dict:
    """Brand = longest 1-3 word prefix still shared by >= 60% of the one-word prefix group."""
    words = titles.str.split()
    pre = {k: Counter(" ".join(w[:k]) for w in words if len(w) >= k) for k in (1, 2, 3)}
    out = {}
    for t, w in zip(titles, words):
        b = w[0] if w else ""
        base = pre[1].get(b, 0)
        for k in (2, 3):
            if len(w) >= k and pre[k][" ".join(w[:k])] >= max(3, 0.6 * base):
                b = " ".join(w[:k])
        out[t] = b
    return out


def clean_title(title: str, brand: str) -> str:
    t = _strip_model(title)
    if brand and t.lower().startswith(brand.lower()):
        t = t[len(brand):]
    parts = [p.strip() for p in t.split(",")]
    keep = [parts[0]] + [p for p in parts[1:] if not colour_family(p)]
    return re.sub(r"\s+", " ", ", ".join(keep)).strip(" ,-")


def sku_colour(specs: dict, title: str) -> str | None:
    for k in COLOUR_SPEC_KEYS:
        if specs.get(k):
            c = colour_family(specs[k])
            if c:
                return c
    segs = _strip_model(title).split(",")[1:]
    for s in segs:
        c = colour_family(s)
        if c:
            return c
    return None


def families_staples() -> pd.DataFrame:
    sku = load("sku_staples.parquet")
    sku["specs_d"] = sku["specs"].map(json.loads)
    brands = staples_brands(sku["title"])
    sku["brand"] = sku["title"].map(brands)
    sku["stem"] = sku["title"].map(name_stem)
    sku["family_id"] = [("S_" + hash_text(lp + "|" + st)[:12]) for lp, st in zip(sku["leaf_path"], sku["stem"])]
    sku["colour"] = [sku_colour(s, t) for s, t in zip(sku["specs_d"], sku["title"])]
    rows = []
    for fid, g in sku.groupby("family_id", sort=False):
        rep = g.iloc[0]
        merged = {}
        for d in g["specs_d"]:                       # representative first, then fill gaps
            for k, v in d.items():
                merged.setdefault(k, v)
        colours = sorted({c for c in g["colour"] if isinstance(c, str)})
        rows.append({
            "retailer": "staples", "family_id": fid,
            "title": rep["title"], "title_clean": clean_title(rep["title"], rep["brand"]),
            "brand": rep["brand"], "price": float(np.nanmedian(g["price"])) if g["price"].notna().any() else np.nan,
            "url": rep["url"], "desc": rep["desc"], "bullets": rep["bullets"],
            "specs": json.dumps(merged, ensure_ascii=False),
            "colours_observed": "|".join(colours), "n_colourways": max(1, len(colours)),
            "n_skus": len(g), "sku_ids": "|".join(g["sku_id"]),
            "leaf_path": rep["leaf_path"], "leaf": rep["leaf"], "l2_key": rep["l2_key"],
            "rating": float(g["rating"].mean()) if g["rating"].notna().any() else np.nan,
            "choice": "", "pages": "",
        })
    return pd.DataFrame(rows)


def families_competitor(name: str) -> pd.DataFrame:
    sku = load(f"sku_{name}.parquet")
    house = {b.lower() for b in (load_yaml("house_brands.yaml").get(name) or [])}
    rows = []
    for fid, g in sku.groupby("sku_id", sort=False):
        rep = g.iloc[0]
        choices = sorted({c for c in g["choice"] if c})
        colours = sorted({c for c in (colour_family(x) for x in choices) if c})
        desc = max(g["desc"], key=len)
        rows.append({
            "retailer": name, "family_id": f"{name[:1].upper()}_{fid}",
            "title": rep["title"], "title_clean": clean_title(rep["title"], rep["vendor"]),
            "brand": rep["vendor"], "house_brand": rep["vendor"].lower() in house,
            "price": float(np.nanmedian(g["price"])) if g["price"].notna().any() else np.nan,
            "url": rep["url"], "desc": desc, "bullets": "", "specs": "{}",
            "choice": " / ".join(choices),
            "colours_observed": "|".join(colours), "n_colourways": max(1, len(colours)),
            "n_skus": len(g), "sku_ids": fid,
            "pages": "|".join(sorted(set(g["page"]))),
            "leaf_path": "", "leaf": "", "l2_key": "", "rating": np.nan,
        })
    return pd.DataFrame(rows)


def run() -> dict:
    fs = families_staples()
    stats = {"staples": {"skus": int(fs["n_skus"].sum()), "families": len(fs)}}
    frames = [fs]
    for name, ad in load_yaml("retailers.yaml").items():
        if ad["role"] == "competitor":
            fc = families_competitor(name)
            stats[name] = {"rows": int(fc["n_skus"].sum()), "families": len(fc),
                           "multi_page": int((fc["pages"].str.count(r"\|") > 0).sum())}
            frames.append(fc)
    fam = pd.concat(frames, ignore_index=True)
    fam["house_brand"] = fam["house_brand"].fillna(False).astype(bool)
    # brand_on_staples: the competitor display brand also prefixes Staples titles (existing supplier)
    st_titles = fs["title"].str.lower()
    brand_hits = {}
    for b in fam.loc[fam["retailer"] != "staples", "brand"].dropna().unique():
        bl = b.lower().strip()
        brand_hits[b] = bool(bl) and bool(st_titles.str.startswith(bl + " ").any())
    fam["brand_on_staples"] = fam["brand"].map(brand_hits).fillna(False).astype(bool)
    save(fam, "families.parquet")
    log(f"S1 families: {stats}")
    return stats
