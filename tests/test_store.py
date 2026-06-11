"""Tests for parsing the `_MC_Store` payload formats."""

from __future__ import annotations

import json

from modelchoice_mcp.store import parse_store, reassemble_chunks

_MODEL = json.dumps(
    {"RootId": "T", "Nodes": {"T": {"type": "terminal", "id": "T", "name": "t", "value": 1}}}
)


def test_v2_envelope_with_nested_json_string() -> None:
    raw = json.dumps({"_v": 2, "trees": {"MC_Tree_1": _MODEL, "MC_Tree_2": _MODEL}})
    trees = parse_store(raw)
    assert set(trees) == {"MC_Tree_1", "MC_Tree_2"}
    # values are the (still-encoded) model JSON strings
    assert json.loads(trees["MC_Tree_1"])["RootId"] == "T"


def test_v1_legacy_raw_model_keyed_to_treeview() -> None:
    trees = parse_store(_MODEL)  # raw model JSON, has RootId, no envelope
    assert list(trees) == ["MC_TreeView"]


def test_empty_returns_empty() -> None:
    assert parse_store("") == {}
    assert parse_store("   ") == {}


def test_non_json_treated_as_single_model() -> None:
    assert parse_store("garbage") == {"MC_TreeView": "garbage"}


def test_reassemble_chunks_stops_at_first_empty() -> None:
    assert reassemble_chunks(["aaa", "bbb", None, "ccc"]) == "aaabbb"
    assert reassemble_chunks([None]) == ""
    assert reassemble_chunks(["solo"]) == "solo"
