
from __future__ import annotations
from typing import Dict, Any, List


def build_appendix_rows(doc: Dict[str, Any]) -> List[Dict[str, str]]:
    """
    Select 'non-primary' fields to list at the end.
    Rule:
      - Not in layout.primary_fields_allowlist
      - Not mentioned in layout.used_in_body (if provided)
    Output rows with: name, type, preview (stringified, possibly truncated)
    """
    fields = doc.get("fields", {}) or {}
    layout = doc.get("layout", {}) or {}
    allow = set(layout.get("primary_fields_allowlist", []) or [])
    used = set(layout.get("used_in_body", []) or [])
    rows: List[Dict[str, str]] = []
    for name, meta in fields.items():
        if name in allow or name in used:
            continue
        ftype = meta.get("type", "unknown")
        val = meta.get("value")
        preview = ""
        if isinstance(val, (list, tuple)):
            preview = ", ".join(str(x) for x in val[:5])
            if len(val) > 5:
                preview += " …"
        elif isinstance(val, dict):
            preview = "{…}"
        elif val is None:
            preview = ""
        else:
            preview = str(val)
        if len(preview) > 200:
            preview = preview[:200] + " …"
        rows.append({"name": name, "type": ftype, "preview": preview})
    rows.sort(key=lambda r: r["name"].lower())
    return rows
