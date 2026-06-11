"""Read ModelChoice's per-workbook tree storage.

ModelChoice persists every decision tree as JSON in a *very hidden*
worksheet named ``_MC_Store`` (see ``TreeJsonStore.cs`` in the add-in).
The payload lives in row 1, chunked across columns A1, B1, C1, … when it
exceeds Excel's ~32k-character cell limit. The reassembled string is a
v2 envelope::

    {"_v": 2, "trees": {"<sheetName>": "<model-json-string>", ...}}

where each *value* is itself a JSON string (the serialized
``DecisionTreeModel``). A v1 legacy form stored the raw model JSON
directly (detected by a top-level ``RootId``), keyed as ``MC_TreeView``.

This module is pure: it turns the raw cell string(s) into a mapping of
sheet-name → model-JSON-string. The Excel/COM reading lives in the
bridge so this stays unit-testable without Excel.
"""

from __future__ import annotations

import json

STORE_SHEET_NAME = "_MC_Store"
_LEGACY_TREE_KEY = "MC_TreeView"


def parse_store(raw: str) -> dict[str, str]:
    """Parse the reassembled ``_MC_Store`` A1 payload into a mapping of
    sheet name → model JSON string. Mirrors ``TreeJsonStore.ParseStorageFormat``.

    Returns an empty dict for empty input.
    """
    if not raw or not raw.strip():
        return {}

    try:
        root = json.loads(raw)
    except json.JSONDecodeError:
        # Not valid JSON at all — treat the whole thing as a single v1 model.
        return {_LEGACY_TREE_KEY: raw}

    if isinstance(root, dict):
        # v2 envelope: {"_v": >=2, "trees": {...}}
        v = root.get("_v")
        trees = root.get("trees")
        if isinstance(v, int) and v >= 2 and isinstance(trees, dict):
            return {
                k: (val if isinstance(val, str) else json.dumps(val))
                for k, val in trees.items()
            }

        # v1: raw model JSON (has a RootId / rootId but no envelope).
        if "RootId" in root or "rootId" in root:
            return {_LEGACY_TREE_KEY: raw}

    # Fallback: opaque content, treat as a single model.
    return {_LEGACY_TREE_KEY: raw}


def reassemble_chunks(cells: list[str | None]) -> str:
    """Reassemble the chunked A1, B1, C1, … payload. Stops at the first
    empty cell. Mirrors ``TreeJsonStore.ReadChunkedContent``."""
    parts: list[str] = []
    for c in cells:
        if c is None or c == "":
            break
        parts.append(c)
    return "".join(parts)
