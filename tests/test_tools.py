"""MCP-tool tests with a fake bridge (no Excel). The bridge only needs
to return tree-name → model-JSON, so we stub `list_trees`."""

from __future__ import annotations

import json
from collections.abc import Iterator

import pytest

from modelchoice_mcp import tools
from modelchoice_mcp.schemas import RollupResponse, TreeList, TreeStructure

# Ground-truth tree (matches ModelChoice's C# rollback test): EV 50, Option2.
_MODEL = json.dumps({
    "RootId": "D",
    "Settings": {"Maximize": True, "ModelName": "Oil"},
    "Nodes": {
        "D": {"type": "decision", "id": "D", "name": "Drill?", "options": [
            {"name": "Drill", "childId": "C"},
            {"name": "Sell", "childId": "T2", "value": 50}]},
        "C": {"type": "chance", "id": "C", "name": "Geology", "branches": [
            {"name": "Dry", "probability": 0.5, "childId": "T1", "value": -100},
            {"name": "Wet", "probability": 0.5, "childId": "T2", "value": 50}]},
        "T1": {"type": "terminal", "id": "T1", "name": "Loss", "value": 0},
        "T2": {"type": "terminal", "id": "T2", "name": "Win", "value": 0},
    },
})


class _FakeBridge:
    def __init__(self, trees: dict[str, str]) -> None:
        self._trees = trees

    def list_trees(self, workbook: str | None = None) -> dict[str, str]:
        return dict(self._trees)


@pytest.fixture
def bridge() -> Iterator[_FakeBridge]:
    b = _FakeBridge({"MC_Tree_1": _MODEL})
    tools.set_bridge_for_testing(b)  # type: ignore[arg-type]
    yield b
    tools.set_bridge_for_testing(None)


def test_list_trees_summarizes(bridge: _FakeBridge) -> None:
    out = tools.list_trees()
    assert isinstance(out, TreeList)
    assert out.count == 1
    s = out.trees[0]
    assert s.name == "MC_Tree_1" and s.model_name == "Oil"
    assert s.decision_count == 1 and s.chance_count == 1 and s.terminal_count == 2
    assert s.root_name == "Drill?"


def test_get_tree_structure(bridge: _FakeBridge) -> None:
    out = tools.get_tree()
    assert isinstance(out, TreeStructure)
    assert out.root_id == "D" and out.maximize is True
    decision = next(n for n in out.nodes if n.kind == "decision")
    assert {b.name for b in decision.branches} == {"Drill", "Sell"}
    chance = next(n for n in out.nodes if n.kind == "chance")
    assert all(b.probability is not None for b in chance.branches)


def test_roll_up_recommends_optimal(bridge: _FakeBridge) -> None:
    out = tools.roll_up()
    assert isinstance(out, RollupResponse)
    assert out.expected_value == 50.0
    assert out.optimal_path == ["Sell"]
    assert "Sell" in out.recommendation and "50" in out.recommendation
    d = next(n for n in out.nodes if n.id == "D")
    assert d.optimal_branch_name == "Sell"


def test_named_tree_selection(bridge: _FakeBridge) -> None:
    out = tools.roll_up(tree_name="MC_Tree_1")
    assert out.name == "MC_Tree_1"
