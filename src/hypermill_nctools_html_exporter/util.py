from __future__ import annotations

import re


_WS = re.compile(r"\s+")


def clean_text(s: str) -> str:
    return _WS.sub(" ", (s or "").strip())


def sanitize_filename(name: str) -> str:
    invalid = '<>:"/\\|?*'
    out = name
    for ch in invalid:
        out = out.replace(ch, "_")
    out = out.rstrip(". ").strip()
    return out or "output"
