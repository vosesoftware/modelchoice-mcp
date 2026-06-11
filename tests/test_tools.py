"""MCP-tool tests with a fake bridge (no Excel). The bridge only needs
to return tree-name → model-JSON, so we stub `list_trees`."""

from __future__ import annotations

import json
from collections.abc import Iterator

import pytest

from modelchoice_mcp import tools
from modelchoice_mcp.schemas import (
    BranchSpec,
    BuildTreeResult,
    EditOp,
    EvpiResult,
    NodeSpec,
    RollbackVerification,
    RollupResponse,
    TreeList,
    TreeSpec,
    TreeStructure,
)

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
    def __init__(
        self, trees: dict[str, str], node_values: dict[str, float] | None = None
    ) -> None:
        self._trees = trees
        self._node_values = node_values or {}

    def list_trees(self, workbook: str | None = None) -> dict[str, str]:
        return dict(self._trees)

    def read_node_values(self, workbook: str | None = None) -> dict[str, float]:
        return dict(self._node_values)

    def run_evpi(self, workbook: str | None = None) -> dict[str, object]:
        return {
            "model_name": "Oil",
            "objective": "Maximize",
            "optimal_ev": 50.0,
            "evpi": 25.0,
            "value_with_perfect_info": 75.0,
        }

    def run_analysis(self, command_name: str, workbook: str | None = None) -> dict[str, object]:
        self.last_command = command_name
        return {
            "command": command_name,
            "new_sheets": ["MC_RB_Verdict"],
            "sheets": [
                "MC_Tree_1", "MC_RB_Verdict", "MC_SensReport", "MC_Tornado",
                "MC_StrategyTable", "MC_StratRegions",
            ],
        }

    def write_tree(self, model_json: str, sheet_name: str | None = None,
                   workbook: str | None = None) -> str:
        self.written_json = model_json
        return sheet_name or "MC_Tree_1"

    def render_tree(self, sheet_name: str, workbook: str | None = None) -> None:
        self.rendered = sheet_name

    def read_sheet(self, sheet_name: str, workbook: str | None = None,
                   max_rows: int = 200, max_cols: int = 20) -> list[list[object]]:
        return [
            ["Robustness verdict", "Robust"],
            ["Robustness Score", "82 / 100"],
            ["Min distance to flip", 0.42],
            [None, None],
        ]


@pytest.fixture
def bridge() -> Iterator[_FakeBridge]:
    b = _FakeBridge({"MC_Tree_1": _MODEL})
    tools.set_bridge_for_testing(b)  # type: ignore[arg-type]
    yield b
    tools.set_bridge_for_testing(None)


def _set(b: _FakeBridge) -> None:
    tools.set_bridge_for_testing(b)  # type: ignore[arg-type]


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


# Computed EVs for the ground-truth tree: D=50, C=-25, T1=-100, T2=50.
_TRUE_CELLS = {"D": 50.0, "C": -25.0, "T1": -100.0, "T2": 50.0}


def test_verify_rollback_matches() -> None:
    _set(_FakeBridge({"MC_Tree_1": _MODEL}, _TRUE_CELLS))
    try:
        out = tools.verify_rollback()
        assert isinstance(out, RollbackVerification)
        assert out.rendered and out.compared_count == 4
        assert out.matches == 4 and not out.mismatches
        assert out.max_abs_diff < 0.01
        assert "verified" in out.verdict
    finally:
        tools.set_bridge_for_testing(None)


def test_verify_rollback_flags_mismatch() -> None:
    bad = dict(_TRUE_CELLS, C=-20.0)  # chance EV off by 5
    _set(_FakeBridge({"MC_Tree_1": _MODEL}, bad))
    try:
        out = tools.verify_rollback()
        assert [m.node_id for m in out.mismatches] == ["C"]
        assert out.mismatches[0].diff == pytest.approx(-5.0)
        assert out.max_abs_diff == pytest.approx(5.0)
    finally:
        tools.set_bridge_for_testing(None)


def test_verify_rollback_not_rendered() -> None:
    _set(_FakeBridge({"MC_Tree_1": _MODEL}, {}))  # no MC_V_ cells
    try:
        out = tools.verify_rollback()
        assert out.rendered is False and out.compared_count == 0
        assert "not rendered" in out.verdict.lower()
    finally:
        tools.set_bridge_for_testing(None)


def test_run_evpi(bridge: _FakeBridge) -> None:
    out = tools.run_evpi()
    assert isinstance(out, EvpiResult)
    assert out.evpi == 25.0 and out.optimal_ev == 50.0
    assert out.value_with_perfect_info == 75.0
    assert "25" in out.interpretation


def test_run_analysis_maps_friendly_name(bridge: _FakeBridge) -> None:
    out = tools.run_analysis("robustness")
    assert out.command == "MC_Robustness_Auto"
    assert out.new_sheets == ["MC_RB_Verdict"]
    assert "read_sheet" in out.note
    assert bridge.last_command == "MC_Robustness_Auto"


def test_run_analysis_unknown_raises(bridge: _FakeBridge) -> None:
    with pytest.raises(ValueError):
        tools.run_analysis("teleport")


def test_run_analysis_covers_all_known(bridge: _FakeBridge) -> None:
    for name in ("risk_profile", "sensitivity", "strategy_table", "evpi"):
        assert tools.run_analysis(name).command.startswith("MC_")


def test_read_sheet(bridge: _FakeBridge) -> None:
    out = tools.read_sheet("MC_RB_Verdict")
    assert out.sheet == "MC_RB_Verdict" and out.row_count == 4
    assert out.rows[2] == ["Min distance to flip", 0.42]
    assert out.truncated is False


def test_run_robustness_extracts_headline(bridge: _FakeBridge) -> None:
    out = tools.run_robustness()
    assert out.verdict == "Robust"
    assert out.robustness_score == "82 / 100"
    assert out.min_distance == 0.42
    # All non-empty label rows are captured as details.
    assert {kv.label for kv in out.details} == {
        "Robustness verdict", "Robustness Score", "Min distance to flip"
    }
    assert out.sheets == ["MC_RB_Verdict"]
    assert bridge.last_command == "MC_Robustness_Auto"


def _oil_spec() -> TreeSpec:
    return TreeSpec(
        root_id="D",
        model_name="Oil",
        maximize=True,
        nodes=[
            NodeSpec(id="D", type="decision", name="Drill?", branches=[
                BranchSpec(name="Drill", child_id="C"),
                BranchSpec(name="Sell", child_id="T2", value=50),
            ]),
            NodeSpec(id="C", type="chance", name="Geology", branches=[
                BranchSpec(name="Dry", child_id="T1", value=-100, probability=0.5),
                BranchSpec(name="Wet", child_id="T2", value=50, probability=0.5),
            ]),
            NodeSpec(id="T1", type="terminal", name="Loss", value=0),
            NodeSpec(id="T2", type="terminal", name="Win", value=0),
        ],
    )


def test_build_tree_dry_run_validates_and_rolls_up() -> None:
    tools.set_bridge_for_testing(_FakeBridge({}))  # type: ignore[arg-type]
    try:
        out = tools.build_tree(_oil_spec())
        assert isinstance(out, BuildTreeResult)
        assert out.written is False and out.sheet is None
        assert out.node_count == 4
        assert out.expected_value == 50.0
        assert out.optimal_path == ["Sell"]
        # The emitted JSON is real ModelChoice model JSON.
        assert '"RootId": "D"' in out.model_json and '"type": "chance"' in out.model_json
    finally:
        tools.set_bridge_for_testing(None)


def test_build_tree_commit_writes_and_renders() -> None:
    fake = _FakeBridge({})
    tools.set_bridge_for_testing(fake)  # type: ignore[arg-type]
    try:
        out = tools.build_tree(_oil_spec(), dry_run=False)
        assert out.written is True and out.sheet == "MC_Tree_1"
        assert fake.rendered == "MC_Tree_1"
    finally:
        tools.set_bridge_for_testing(None)


def test_edit_tree_changes_payoff_and_flips_decision() -> None:
    fake = _FakeBridge({"MC_Tree_1": _MODEL})
    tools.set_bridge_for_testing(fake)  # type: ignore[arg-type]
    try:
        # Drop the 'Sell' payoff below the chance EV (-25) so Drill wins.
        out = tools.edit_tree(
            [EditOp(op="set_branch_value", node_id="D", branch_name="Sell", value=-30)]
        )
        assert isinstance(out, BuildTreeResult)
        assert out.written is False
        assert out.optimal_path == ["Drill"]
        assert out.expected_value == -25.0
    finally:
        tools.set_bridge_for_testing(None)


def test_edit_tree_set_objective_minimize() -> None:
    fake = _FakeBridge({"MC_Tree_1": _MODEL})
    tools.set_bridge_for_testing(fake)  # type: ignore[arg-type]
    try:
        out = tools.edit_tree([EditOp(op="set_objective", maximize=False)])
        # Minimizing prefers the chance node (EV -25) over Sell (50).
        assert out.expected_value == -25.0 and out.optimal_path == ["Drill"]
    finally:
        tools.set_bridge_for_testing(None)


def test_edit_tree_commit_writes_same_sheet() -> None:
    fake = _FakeBridge({"MC_Tree_1": _MODEL})
    tools.set_bridge_for_testing(fake)  # type: ignore[arg-type]
    try:
        out = tools.edit_tree(
            [EditOp(op="set_probability", node_id="C", branch_name="Wet", value=0.8)],
            dry_run=False,
        )
        assert out.written is True and out.sheet == "MC_Tree_1"
        assert fake.rendered == "MC_Tree_1"
        # The edited probability is reflected in the written JSON.
        assert '"probability": 0.8' in fake.written_json
    finally:
        tools.set_bridge_for_testing(None)


def test_edit_tree_bad_target_raises() -> None:
    from modelchoice_mcp.tree import TreeParseError

    tools.set_bridge_for_testing(_FakeBridge({"MC_Tree_1": _MODEL}))  # type: ignore[arg-type]
    try:
        with pytest.raises(TreeParseError):
            tools.edit_tree([EditOp(op="set_terminal_value", node_id="NOPE", value=1)])
    finally:
        tools.set_bridge_for_testing(None)


def test_run_decision_report_strategy(bridge: _FakeBridge) -> None:
    out = tools.run_decision_report("strategy_table")
    assert bridge.last_command == "MC_ExportStrategyTable_Auto"
    assert out.primary_sheet == "MC_StrategyTable"
    assert out.sheets == ["MC_StrategyTable", "MC_StratRegions"]
    assert out.rows  # primary sheet content came through


def test_run_decision_report_unknown_raises(bridge: _FakeBridge) -> None:
    with pytest.raises(ValueError):
        tools.run_decision_report("astrology")


def test_run_decision_report_maps_all(bridge: _FakeBridge) -> None:
    cmds = {
        "policy_suggestion": "MC_PolicySuggestion_Auto",
        "decision_brief": "MC_DecisionBrief_Auto",
        "mcda_report": "MC_McdaReport_Auto",
    }
    for name, cmd in cmds.items():
        out = tools.run_decision_report(name)
        assert out.command == cmd


def test_run_sensitivity(bridge: _FakeBridge) -> None:
    out = tools.run_sensitivity()
    assert bridge.last_command == "MC_SensitivityAnalysis_Auto"
    assert out.sheets == ["MC_SensReport", "MC_Tornado"]
    assert out.report_rows  # MC_SensReport rows came through
    assert "tornado" in out.note.lower()
