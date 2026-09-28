"""S0 Ingest & clean: retailer adapters -> one SKU table per retailer (methodology §5.1)."""
from __future__ import annotations

import ast
import json
import re

import numpy as np
import pandas as pd
from ftfy import fix_text

from .common import interim, load_yaml, log, path, save


def _read_excel_cached(fname: str) -> pd.DataFrame:
    src = path("raw_dir") / fname
    cache = interim("raw_" + re.sub(r"\W+", "_", fname) + ".pkl")
    if cache.exists() and cache.stat().st_mtime > src.stat().st_mtime:
        return pd.read_pickle(cache)
    log(f"reading {fname}")
    df = pd.read_excel(src)
    df.to_pickle(cache)
    return df


def parse_price(v) -> tuple[float, bool]:
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return np.nan, False
    s = str(v)
    per_item = "per item" in s.lower()
    m = re.search(r"[\d,]+(\.\d+)?", s.replace("$", ""))
    return (float(m.group(0).replace(",", "")) if m else np.nan), per_item


_MOJIBAKE = {"â„¢": "™", "â€™": "'", "â€œ": '"', "â€\x9d": '"', "â€“": "–", "â€”": "—", "Â®": "®", "Â°": "°"}


def _clean(s) -> str:
    if s is None or (isinstance(s, float) and np.isnan(s)):
        return ""
    s = str(s)
    for bad, good in _MOJIBAKE.items():
        s = s.replace(bad, good)
    return re.sub(r"\s+", " ", fix_text(s)).strip()


def parse_staples_description(raw) -> tuple[str, str, dict]:
    """-> (paragraph text, bullets joined, {spec name: value})"""
    if not isinstance(raw, str) or not raw.startswith("{"):
        return _clean(raw), "", {}
    try:
        d = ast.literal_eval(raw)
    except (ValueError, SyntaxError):
        return _clean(raw), "", {}
    para = " ".join(_clean(p) for p in d.get("paragraph", []) or [])
    bullets = " • ".join(_clean(b) for b in d.get("bullets", []) or [])
    specs = {}
    for sp in d.get("specification", []) or []:
        name, val = _clean(sp.get("name")), _clean(sp.get("value"))
        if name and val and name not in specs:
            specs[name] = val
    return para, bullets, specs


def ingest_staples(ad: dict) -> pd.DataFrame:
    raw = _read_excel_cached(ad["file"])
    c = ad["columns"]
    raw = raw.drop_duplicates()
    rows = []
    for r in raw.to_dict("records"):
        levels = [str(r[k]) for k in ad["path_columns"] if isinstance(r.get(k), str) and r[k].strip()]
        para, bullets, specs = parse_staples_description(r.get(c["description"]))
        price, per_item = parse_price(r.get(c["price"]))
        rating = re.search(r"([\d.]+)\s*stars", str(r.get(c["rating"]) or ""))
        rows.append({
            "retailer": "staples",
            "sku_id": str(r[c["id"]]),
            "title": _clean(r[c["title"]]),
            "price": price, "per_item_price": per_item,
            "url": str(r.get(c["url"]) or ""),
            "model": _clean(r.get(c["model"])),
            "rating": float(rating.group(1)) if rating else np.nan,
            "desc": para, "bullets": bullets, "specs": json.dumps(specs, ensure_ascii=False),
            "leaf_path": " > ".join(levels),
            "l2_key": " > ".join(levels[:2]),
            "leaf": str(r[ad["leaf_column"]]),
        })
    df = pd.DataFrame(rows)
    # a SKU listed under several leaves keeps its first leaf as primary (secondary kept for reference)
    sec = df.groupby("sku_id")["leaf_path"].agg(lambda s: "|".join(sorted(set(s))[1:]) if s.nunique() > 1 else "")
    df = df.drop_duplicates("sku_id", keep="first").copy()
    df["secondary_leaves"] = df["sku_id"].map(sec).fillna("")
    return df


def ingest_competitor(name: str, ad: dict) -> pd.DataFrame:
    raw = _read_excel_cached(ad["file"])
    c = ad["columns"]
    raw = raw.drop_duplicates()
    rows = []
    for r in raw.to_dict("records"):
        desc_full = _clean(r.get(c["description"]))
        desc = desc_full.split(" | ")[0]            # right of " | " = customer review (PII): dropped
        page = str(r[c["page"]])
        for k, v in (ad.get("page_name_fixes") or {}).items():   # before whitespace is collapsed
            page = page.replace(k, v)
        page = _clean(page)
        price, per_item = parse_price(r.get(c["price"]))
        rows.append({
            "retailer": name,
            "sku_id": str(r[c["id"]]),
            "title": _clean(r[c["title"]]),
            "price": price, "per_item_price": per_item,
            "url": str(r.get(c["url"]) or ""),
            "vendor": re.sub(r"[®™]", "", _clean(r.get(c["vendor"]))).strip(),
            "choice": _clean(r.get(c["choice"])),
            "desc": desc, "bullets": "", "specs": "{}",
            "page": page,
            "had_review_text": " | " in desc_full,
        })
    return pd.DataFrame(rows)


def run() -> dict:
    retailers = load_yaml("retailers.yaml")
    stats = {}
    for name, ad in retailers.items():
        df = ingest_staples(ad) if ad["role"] == "base" else ingest_competitor(name, ad)
        save(df, f"sku_{name}.parquet")
        stats[name] = {"rows": int(len(df)), "skus": int(df["sku_id"].nunique()),
                       "price_missing": int(df["price"].isna().sum())}
        log(f"S0 {name}: {stats[name]}")
    return stats
