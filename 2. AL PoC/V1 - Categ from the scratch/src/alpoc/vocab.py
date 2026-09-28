"""Controlled-vocabulary matchers (the local 'rules' extraction instrument)."""
from __future__ import annotations

import re
from functools import lru_cache

from .common import load_yaml


class Matcher:
    """Maps free text to a canonical value from a {value: [regex, ...]} vocabulary.

    mode 'position': the earliest match wins (longest pattern first at a tie).
    mode 'priority': the first value, in listed order, with any match wins.
    mode 'multi'   : every value with a match.
    """

    def __init__(self, values: dict, mode: str = "position"):
        self.mode = mode
        self.values = {str(k): [str(p) for p in v] for k, v in values.items()}
        pats = []
        for val, plist in self.values.items():
            for p in plist:
                pats.append((val, p))
        pats.sort(key=lambda t: -len(t[1]))
        self._alts = pats
        self._one = re.compile("|".join(f"(?P<g{i}>{self._wrap(p)})" for i, (_, p) in enumerate(pats)), re.I)
        self._per_value = {val: re.compile("|".join(self._wrap(p) for p in plist), re.I)
                           for val, plist in self.values.items()}

    @staticmethod
    def _wrap(p: str) -> str:
        return p if ("\\b" in p or p.startswith("(?")) else rf"\b(?:{p})\b"

    def first(self, text: str):
        if not text:
            return None
        if self.mode == "priority":
            for val, rx in self._per_value.items():
                if rx.search(text):
                    return val
            return None
        m = self._one.search(text)
        if not m:
            return None
        i = int(m.lastgroup[1:])
        return self._alts[i][0]

    def all(self, text: str) -> list[str]:
        if not text:
            return []
        return [val for val, rx in self._per_value.items() if rx.search(text)]

    def from_sources(self, sources: list[tuple[str, str, float]]):
        """sources = [(name, text, confidence)] in precedence order -> (value, source, confidence)."""
        if self.mode == "multi":
            found, srcs = [], []
            for name, text, _ in sources:
                for v in self.all(text):
                    if v not in found:
                        found.append(v)
                        srcs.append(name)
            return ("|".join(found), "+".join(sorted(set(srcs))), 0.7 if found else 0.0)
        for name, text, conf in sources:
            v = self.first(text)
            if v is not None:
                return v, name, conf
        return None, None, 0.0


@lru_cache(maxsize=None)
def universal() -> dict:
    return load_yaml("schema_universal.yaml")


@lru_cache(maxsize=None)
def matcher(attr: str) -> Matcher:
    spec = universal()[attr]
    return Matcher(spec["values"], spec.get("mode", "position"))


def colour_family(text: str):
    return matcher("colour_family").first(text)


def colour_tone(family):
    return universal()["colour_tone_map"].get(family) if family else None
