"""Tree parser + roller tests, validated against ModelChoice's own
authoritative semantics (the C# `Model_Validate_Compile_Rollback_Works`
test in ModelChoice.Core.Tests) so the Python rollback matches the
add-in byte-for-byte."""

from __future__ import annotations

import json

import pytest

from modelchoice_mcp.tree import TreeParseError, parse_model, rollup, to_model_json

# Ground truth from ModelChoice.Core.Tests/UnitTest1.cs:
# Decision D -> {Option1->chance C, Option2->T2 value 50};
# C -> {Bad 0.5 value -100 -> T1, Good 0.5 value +50 -> T2}. Maximize.
# Asserted: rollback EV = 50 (Option2 dominates chance EV -25).
_GROUND_TRUTH = {
    "RootId": "D",
    "Settings": {"Maximize": True, "ChanceProbabilities": "AutoNormalize"},
    "Nodes": {
        "D": {"type": "decision", "id": "D", "name": "Drill?", "options": [
            {"name": "Option1", "childId": "C"},
            {"name": "Option2", "childId": "T2", "value": 50}]},
        "C": {"type": "chance", "id": "C", "name": "Geology", "branches": [
            {"name": "Bad", "probability": 0.5, "childId": "T1", "value": -100},
            {"name": "Good", "probability": 0.5, "childId": "T2", "value": 50}]},
        "T1": {"type": "terminal", "id": "T1", "name": "Loss", "value": 0},
        "T2": {"type": "terminal", "id": "T2", "name": "Win", "value": 0},
    },
}


def test_rollup_matches_modelchoice_ground_truth() -> None:
    r = rollup(parse_model(json.dumps(_GROUND_TRUTH)))
    assert r.expected_value == 50.0
    assert r.optimal_path == ["Option2"]
    assert r.node_results["C"].expected_value == -25.0  # chance branch
    assert r.node_results["D"].optimal_child_id == "T2"


def test_minimize_flips_decision() -> None:
    model = dict(_GROUND_TRUTH)
    model["Settings"] = {"Maximize": False, "ChanceProbabilities": "AutoNormalize"}
    r = rollup(parse_model(json.dumps(model)))
    # Minimizing prefers the chance node (EV -25) over Option2 (50).
    assert r.expected_value == -25.0
    assert r.optimal_path == ["Option1"]


def test_chance_auto_normalizes_probabilities() -> None:
    model = {
        "RootId": "C", "Settings": {"ChanceProbabilities": "AutoNormalize"},
        "Nodes": {
            "C": {"type": "chance", "id": "C", "name": "x", "branches": [
                {"name": "a", "probability": 1, "childId": "T1"},
                {"name": "b", "probability": 3, "childId": "T2"}]},  # 1:3, not summing to 1
            "T1": {"type": "terminal", "id": "T1", "name": "A", "value": 0},
            "T2": {"type": "terminal", "id": "T2", "name": "B", "value": 100},
        },
    }
    r = rollup(parse_model(json.dumps(model)))
    assert r.expected_value == pytest.approx(75.0)  # 0.25*0 + 0.75*100


def test_accumulated_branch_values_reach_terminal() -> None:
    model = {
        "RootId": "D", "Settings": {"Maximize": True},
        "Nodes": {
            "D": {"type": "decision", "id": "D", "name": "d", "options": [
                {"name": "go", "childId": "T", "value": 123.0}]},
            "T": {"type": "terminal", "id": "T", "name": "t", "value": 0},
        },
    }
    assert rollup(parse_model(json.dumps(model))).expected_value == 123.0


def test_serialize_round_trips_through_parser() -> None:
    # parse -> serialize -> parse must preserve structure and rollback.
    tree = parse_model(json.dumps(_GROUND_TRUTH))
    again = parse_model(to_model_json(tree))
    assert again.root_id == tree.root_id
    assert set(again.nodes) == set(tree.nodes)
    assert rollup(again).expected_value == rollup(tree).expected_value == 50.0


def test_serialize_emits_modelchoice_shape() -> None:
    tree = parse_model(json.dumps(_GROUND_TRUTH))
    raw = json.loads(to_model_json(tree))
    assert raw["RootId"] == "D"
    assert raw["Nodes"]["D"]["type"] == "decision"
    assert raw["Nodes"]["C"]["type"] == "chance"
    # chance branch carries probability; the converter shape uses childId
    branch = raw["Nodes"]["C"]["branches"][0]
    assert "probability" in branch and "childId" in branch
    # terminal carries its own value
    assert raw["Nodes"]["T1"]["value"] == 0
    assert raw["Settings"]["Maximize"] is True


def test_bad_json_raises() -> None:
    with pytest.raises(TreeParseError):
        parse_model("not json")


def test_missing_root_raises() -> None:
    with pytest.raises(TreeParseError):
        parse_model(json.dumps({"RootId": "X", "Nodes": {}}))
