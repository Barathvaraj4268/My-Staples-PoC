"""Canonical product cards (§6.4): the same template on both sides, no price, no raw spec dump."""
from __future__ import annotations

import re

import pandas as pd

from .common import l2_config, split_multi
from .vocab import matcher

_AESTHETIC = ["colour_family", "material_class", "style_family"]


def node_label(node_id: str) -> str:
    parts = node_id.split(" > ")
    return " > ".join(parts[-2:]) if len(parts) > 2 else node_id


def _val(v) -> str:
    return "" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v)


def full_card(r: pd.Series, tier2: list[str]) -> str:
    feats = ", ".join(f"{a.replace('_', ' ')} {_val(r.get(a))}" for a in tier2 if _val(r.get(a)))
    return (f"{node_label(r['node_id'])} | {r['title_clean']} | colour: {_val(r['colour_family'])} "
            f"{_val(r['colour_tone'])} | material: {_val(r['material_class'])} | style: {_val(r['style_family'])} | "
            f"features: {feats} | vibe: {', '.join(split_multi(r['aesthetic_tags']))} | "
            f"use: {', '.join(split_multi(r['use_context']))}")


def strip_aesthetic_words(title: str) -> str:
    """Remove colour / material / style words so the functional view ignores the look."""
    t = title
    for a in _AESTHETIC:
        for plist in matcher(a).values.values():
            for p in plist:
                t = re.sub(matcher(a)._wrap(p), " ", t, flags=re.I)
    return re.sub(r"\s+", " ", t).strip(" ,-|")


def functional_card(r: pd.Series, core: list[str]) -> str:
    feats = ", ".join(f"{a.replace('_', ' ')} {_val(r.get(a))}" for a in core if _val(r.get(a)))
    return f"{node_label(r['node_id'])} | {strip_aesthetic_words(r['title_clean'])} | {feats}"


TYPE_ATTRS = ["form_factor", "desk_type"]


def _tx(r: pd.Series, a: str) -> str:
    """Text-instrument value (same instrument on both sides, including its misses); falls back to the field."""
    return _val(r.get("txt_" + a)) if ("txt_" + a) in r.index else _val(r.get(a))


def neutral_card(r: pd.Series, tier2: list[str]) -> str:
    """Source-neutral card: only text-instrument extracted fields, no free text, so neither retailer's copy style
    nor its spec-sheet completeness can separate the two sides."""
    typ = " ".join(_tx(r, a) for a in TYPE_ATTRS if _tx(r, a))
    feats = ", ".join(f"{a.replace('_', ' ')} {_tx(r, a)}" for a in tier2 if a not in TYPE_ATTRS and _tx(r, a))
    return (f"{node_label(r['node_id'])} | type: {typ} | colour: {_tx(r, 'colour_family')} {_tx(r, 'colour_tone')} | "
            f"material: {_tx(r, 'material_class')} | style: {_val(r['style_family'])} | features: {feats} | "
            f"vibe: {', '.join(split_multi(r['aesthetic_tags']))} | use: {', '.join(split_multi(r['use_context']))} | "
            f"size: {_tx(r, 'size_class')}")


def neutral_functional_card(r: pd.Series, core: list[str]) -> str:
    typ = " ".join(_tx(r, a) for a in TYPE_ATTRS if _tx(r, a))
    feats = ", ".join(f"{a.replace('_', ' ')} {_tx(r, a)}" for a in core if a not in TYPE_ATTRS and _tx(r, a))
    return f"{node_label(r['node_id'])} | type: {typ} | features: {feats} | size: {_tx(r, 'size_class')}"


def tier2_attrs(l2: str) -> list[str]:
    c2 = l2_config(l2)
    return list((c2.get("tier2") or {}).keys()) + list((c2.get("numeric_bands") or {}).keys())


def functional_core(l2: str) -> list[str]:
    return l2_config(l2).get("functional_core") or tier2_attrs(l2)
