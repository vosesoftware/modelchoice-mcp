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
    ControlPanelResult,
    EditOp,
    EviiResult,
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

    def run_utility(self, function: str = "exponential", risk_tolerance: float = 1.0,
                    workbook: str | None = None) -> dict[str, object]:
        self.utility_args = (function, risk_tolerance)
        return {
            "function": "Exponential",
            "risk_tolerance": 50.0,
            "expected_value": 50.0,
            "certainty_equivalent": 30.0,
            "risk_premium": 20.0,
            "optimal_decision_ev": "Sell",
            "optimal_decision_utility": "Sell",
        }

    def run_analysis(self, command_name: str, workbook: str | None = None) -> dict[str, object]:
        self.last_command = command_name
        return {
            "command": command_name,
            "new_sheets": ["MC_RB_Verdict"],
            "sheets": [
                "MC_Tree_1", "MC_RB_Verdict", "MC_SensReport", "MC_Tornado",
                "MC_StrategyTable", "MC_StratRegions", "MC_RiskProfile",
                "MC_ForceOutcome", "MC_TwoWaySens",
            ],
        }

    def import_precisiontree(self, file_path: str) -> dict[str, object]:
        self.imported_path = file_path
        return {"workbook": "oil_ModelChoice.xlsx", "trees": ["MC_Tree_1"]}

    def apply_mcda(self, spec_json: str, workbook: str | None = None) -> None:
        self.mcda_spec = spec_json

    def run_evii(self, chance_node: str, likelihoods: list[list[float]],
                 test_cost: float = 0.0, signals: list[str] | None = None,
                 test_name: str = "Test", workbook: str | None = None) -> dict[str, object]:
        self.evii_args = (chance_node, likelihoods, test_cost, signals, test_name)
        # Mimic the MC_EVII sheet layout: page header (1-5), Key Values
        # section (6), header (7), data rows 8-12, blank, recommendation.
        rows: list[list[object]] = [
            [None, "ModelChoice — EVII", None],
            [None, "Performed by …", None],
            [None, "Date …", None],
            [None, "Test …", None],
            [None, None, None],
            [None, "Key Values", None],
            [None, "Metric", "Value"],
            [None, "Prior EV", "100.00"],
            [None, "EVPI (upper bound)", "50.00"],
            [None, "EVII (test value)", "30.00"],
            [None, "Test cost", "5.00"],
            [None, "Net value", "25.00"],
            [None, None, None],
            [None, "The test is worthwhile.", None],
        ]
        return {"rows": rows, "sheets": ["MC_Tree_1", "MC_EVII"]}

    def build_control_panel(self, sheet_name: str | None = None,
                            workbook: str | None = None) -> dict[str, object]:
        self.panel_sheet = sheet_name
        # Mimic the panel block read back from the top of the tree sheet.
        rows: list[list[object]] = [
            [None, "Control Panel — Inputs", None],
            [None, "Input", "Value"],
            [None, "Geology: P(Dry)", 0.5],
            [None, "Geology: Dry — value", -100.0],
            [None, "Geology: P(Wet)", 0.5],
            [None, "Geology: Wet — value", 50.0],
            [None, "Drill?: Sell — value", 50.0],
            [None, None, None],
        ]
        return {"sheet": sheet_name or "MC_Tree_1", "rows": rows}

    def set_input_formula(self, named_range: str, formula: str,
                          sheet_name: str | None = None,
                          workbook: str | None = None) -> str:
        self.input_formula = (named_range, formula)
        return sheet_name or "MC_Tree_1"

    def open_workbook(self, path: str) -> dict[str, object]:
        self.opened_path = path
        return {
            "workbook": "oil.xlsx",
            "sheets": ["MC_Tree_1", "_MC_Store"],
            "trees": ["MC_Tree_1"],
        }

    def close_workbook(self, workbook: str, save: bool = False) -> dict[str, object]:
        self.closed = (workbook, save)
        return {"closed": workbook, "saved": save}

    def write_tree(self, model_json: str, sheet_name: str | None = None,
                   workbook: str | None = None) -> str:
        self.written_json = model_json
        return sheet_name or "MC_Tree_1"

    def render_tree(self, sheet_name: str, workbook: str | None = None) -> None:
        self.rendered = sheet_name

    def read_sheet(self, sheet_name: str, workbook: str | None = None,
                   max_rows: int = 200, max_cols: int = 20) -> list[list[object]]:
        if sheet_name == "MC_RiskProfile":
            return [
                [None, "Risk Profile", None, None],
                [None, "Performed by ...", None, None],
                [None, "Date ...", None, None],
                [None, "Start node ...", None, None],
                [None, None, None, None],
                [None, "Statistical Summary", None, None],
                [None, "Statistic", "Drill", "Sell"],
                [None, "Expected Value", -25.0, 50.0],
                [None, "Minimum", -100.0, 50.0],
                [None, "Maximum", 50.0, 50.0],
                [None, "Std Deviation", 75.0, 0.0],
                [None, None, None, None],
            ]
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


def test_build_control_panel_parses_inputs(bridge: _FakeBridge) -> None:
    out = tools.build_control_panel("MC_Tree_1")
    assert isinstance(out, ControlPanelResult)
    assert bridge.panel_sheet == "MC_Tree_1"
    assert out.sheet == "MC_Tree_1"
    # Title and header rows are skipped; only labelled numeric rows count.
    assert out.linked_count == 5
    labels = {kv.label for kv in out.inputs}
    assert "Geology: P(Dry)" in labels
    assert "Drill?: Sell — value" in labels
    assert "Input" not in labels
    assert "persist across re-renders" in out.note


def test_open_workbook(bridge: _FakeBridge) -> None:
    from modelchoice_mcp.schemas import OpenWorkbookResult

    out = tools.open_workbook(r"C:\models\oil.xlsx")
    assert isinstance(out, OpenWorkbookResult)
    assert bridge.opened_path == r"C:\models\oil.xlsx"
    assert out.workbook == "oil.xlsx"
    assert out.trees == ["MC_Tree_1"]
    assert "MC_Tree_1" in out.note


def test_close_workbook(bridge: _FakeBridge) -> None:
    from modelchoice_mcp.schemas import CloseWorkbookResult

    out = tools.close_workbook("oil.xlsx", save=True)
    assert isinstance(out, CloseWorkbookResult)
    assert bridge.closed == ("oil.xlsx", True)
    assert out.workbook == "oil.xlsx" and out.saved is True
    assert "saved" in out.note


def test_close_workbook_discards_by_default(bridge: _FakeBridge) -> None:
    out = tools.close_workbook("oil.xlsx")
    assert bridge.closed == ("oil.xlsx", False)
    assert out.saved is False
    assert "discarded" in out.note


def test_set_input_distribution_value(bridge: _FakeBridge) -> None:
    from modelchoice_mcp.schemas import InputDistributionResult

    out = tools.set_input_distribution("C", "Wet", "VoseNormal(50,10)", kind="value")
    assert isinstance(out, InputDistributionResult)
    # Wet -> child T2, value input => MC_BV_C_T2; leading '=' added.
    assert out.named_range == "MC_BV_C_T2"
    assert out.formula == "=VoseNormal(50,10)"
    assert bridge.input_formula == ("MC_BV_C_T2", "=VoseNormal(50,10)")
    assert out.tree == "MC_Tree_1"
    assert "modelrisk-mcp" in out.note


def test_set_input_distribution_probability(bridge: _FakeBridge) -> None:
    out = tools.set_input_distribution(
        "C", "Dry", "=VosePERT(0.3,0.5,0.7)", kind="probability"
    )
    # Dry -> child T1, probability input => MC_BP_C_T1; existing '=' kept.
    assert out.named_range == "MC_BP_C_T1"
    assert out.formula == "=VosePERT(0.3,0.5,0.7)"


def test_set_input_distribution_rejects_probability_on_decision(bridge: _FakeBridge) -> None:
    with pytest.raises(ValueError, match="chance node"):
        tools.set_input_distribution("D", "Drill", "VoseUniform(0,1)", kind="probability")


def test_set_input_distribution_unknown_branch(bridge: _FakeBridge) -> None:
    with pytest.raises(ValueError, match="no branch"):
        tools.set_input_distribution("C", "Nope", "VoseNormal(0,1)")


def test_set_input_distribution_bad_kind(bridge: _FakeBridge) -> None:
    with pytest.raises(ValueError, match=r"value.*probability"):
        tools.set_input_distribution("C", "Wet", "VoseNormal(0,1)", kind="bogus")


def test_run_utility(bridge: _FakeBridge) -> None:
    from modelchoice_mcp.schemas import UtilityResult

    out = tools.run_utility(risk_tolerance=50.0, function="exponential")
    assert isinstance(out, UtilityResult)
    assert bridge.utility_args == ("exponential", 50.0)
    assert out.expected_value == 50.0
    assert out.certainty_equivalent == 30.0
    assert out.risk_premium == 20.0
    assert out.decision_changed is False  # Sell == Sell
    assert "risk premium" in out.interpretation


def test_run_evii_parses_key_values(bridge: _FakeBridge) -> None:
    out = tools.run_evii("Geology", [[0.8, 0.2], [0.3, 0.7]], test_cost=5.0)
    assert isinstance(out, EviiResult)
    assert out.prior_ev == 100.0
    assert out.evpi_upper_bound == 50.0
    assert out.evii == 30.0
    assert out.test_cost == 5.0
    assert out.net_value == 25.0
    assert out.worthwhile is True
    assert out.recommendation == "The test is worthwhile."
    assert out.sheets == ["MC_EVII"]
    assert "worth running" in out.interpretation


def test_run_evii_passes_spec_through(bridge: _FakeBridge) -> None:
    tools.run_evii(
        "Geology", [[0.9, 0.1], [0.2, 0.8]],
        test_cost=12.0, signals=["Pos", "Neg"], test_name="Seismic",
    )
    chance_node, likelihoods, cost, signals, name = bridge.evii_args
    assert chance_node == "Geology"
    assert likelihoods == [[0.9, 0.1], [0.2, 0.8]]
    assert cost == 12.0
    assert signals == ["Pos", "Neg"]
    assert name == "Seismic"


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


def test_export_tree_json(bridge: _FakeBridge) -> None:
    from modelchoice_mcp.schemas import TreeExport

    out = tools.export_tree_json()
    assert isinstance(out, TreeExport)
    assert out.tree == "MC_Tree_1"
    assert out.model_name == "Oil"
    assert out.node_count == 4
    assert out.model_json == _MODEL


def test_import_tree_json_roundtrip() -> None:
    fake = _FakeBridge({"MC_Tree_1": _MODEL})
    tools.set_bridge_for_testing(fake)  # type: ignore[arg-type]
    try:
        # dry-run validates + rolls back, no write
        prev = tools.import_tree_json(_MODEL)
        assert prev.written is False
        assert prev.expected_value == 50.0 and prev.optimal_path == ["Sell"]
        assert prev.node_count == 4
        # commit writes the raw JSON as-is + renders
        out = tools.import_tree_json(_MODEL, dry_run=False)
        assert out.written is True and out.sheet == "MC_Tree_1"
        assert fake.written_json == _MODEL
        assert fake.rendered == "MC_Tree_1"
    finally:
        tools.set_bridge_for_testing(None)


def test_import_tree_json_bad_json_raises() -> None:
    from modelchoice_mcp.tree import TreeParseError

    tools.set_bridge_for_testing(_FakeBridge({}))  # type: ignore[arg-type]
    try:
        with pytest.raises((TreeParseError, ValueError)):
            tools.import_tree_json("{not valid model json}")
    finally:
        tools.set_bridge_for_testing(None)


def _mcda_specs():
    from modelchoice_mcp.schemas import BranchSpec, CriterionSpec, McdaSpec, NodeSpec, TreeSpec

    tree = TreeSpec(
        root_id="D", model_name="Vendor choice", maximize=True,
        nodes=[
            NodeSpec(id="D", type="decision", name="Vendor", branches=[
                BranchSpec(name="Acme", child_id="T1"),
                BranchSpec(name="Globex", child_id="T2")]),
            NodeSpec(id="T1", type="terminal", name="Acme outcome", value=0),
            NodeSpec(id="T2", type="terminal", name="Globex outcome", value=0),
        ],
    )
    mcda = McdaSpec(
        criteria=[CriterionSpec(id="safety", name="Safety", weight=0.5, maximize=True,
                                options=["Low", "Medium", "High"])],
        financial_weight=0.5,
        aggregation="weighted_sum",
        terminal_scores={"T1": {"safety": "Low"}, "T2": {"safety": "High"}},
    )
    return tree, mcda


def test_build_mcda_dry_run(bridge: _FakeBridge) -> None:
    from modelchoice_mcp.schemas import McdaBuildResult

    tree, mcda = _mcda_specs()
    out = tools.build_mcda(tree, mcda)
    assert isinstance(out, McdaBuildResult)
    assert out.written is False
    assert out.criteria == ["Safety"]
    assert out.terminals_scored == 2
    assert out.node_count == 3


def test_build_mcda_commit() -> None:
    fake = _FakeBridge({})
    tools.set_bridge_for_testing(fake)  # type: ignore[arg-type]
    try:
        tree, mcda = _mcda_specs()
        out = tools.build_mcda(tree, mcda, dry_run=False)
        assert out.written is True and out.sheet == "MC_Tree_1"
        assert fake.rendered == "MC_Tree_1"
        import json as _json
        spec = _json.loads(fake.mcda_spec)
        assert spec["financialWeight"] == 0.5
        assert spec["criteria"][0]["id"] == "safety"
        assert spec["terminalScores"] == {"T1": {"safety": "Low"}, "T2": {"safety": "High"}}
    finally:
        tools.set_bridge_for_testing(None)


def test_build_mcda_bad_option_raises(bridge: _FakeBridge) -> None:
    tree, mcda = _mcda_specs()
    mcda.terminal_scores = {"T1": {"safety": "Stellar"}}  # not a defined option
    with pytest.raises(ValueError):
        tools.build_mcda(tree, mcda)


def _mcda_ahp_specs():
    """Two non-financial criteria + a consistent 3x3 AHP matrix
    [financial, cost, safety] proportional to weights [4, 2, 1]."""
    from modelchoice_mcp.schemas import BranchSpec, CriterionSpec, McdaSpec, NodeSpec, TreeSpec

    tree = TreeSpec(
        root_id="D", model_name="Site", maximize=True,
        nodes=[
            NodeSpec(id="D", type="decision", name="Site", branches=[
                BranchSpec(name="A", child_id="T1"),
                BranchSpec(name="B", child_id="T2")]),
            NodeSpec(id="T1", type="terminal", name="A", value=0),
            NodeSpec(id="T2", type="terminal", name="B", value=0),
        ],
    )
    mcda = McdaSpec(
        criteria=[
            CriterionSpec(id="cost", name="Cost", weight=0.0, maximize=False,
                          options=["High", "Low"]),
            CriterionSpec(id="safety", name="Safety", weight=0.0, maximize=True,
                          options=["Low", "High"]),
        ],
        weight_source="ahp",
        ahp_matrix=[[1.0, 2.0, 4.0], [0.5, 1.0, 2.0], [0.25, 0.5, 1.0]],
        terminal_scores={"T1": {"cost": "Low", "safety": "Low"},
                         "T2": {"cost": "High", "safety": "High"}},
    )
    return tree, mcda


def test_build_mcda_ahp_dry_run(bridge: _FakeBridge) -> None:
    tree, mcda = _mcda_ahp_specs()
    out = tools.build_mcda(tree, mcda)
    assert out.weight_source == "ahp"
    # Perfectly consistent matrix -> CR ~ 0.
    assert out.consistency_ratio is not None and out.consistency_ratio < 0.01
    by_label = {kv.label: kv.value for kv in out.weights}
    assert by_label["Financial"] == pytest.approx(4 / 7, abs=1e-3)
    assert by_label["Cost"] == pytest.approx(2 / 7, abs=1e-3)
    assert by_label["Safety"] == pytest.approx(1 / 7, abs=1e-3)


def test_build_mcda_ahp_commit_sends_matrix() -> None:
    fake = _FakeBridge({})
    tools.set_bridge_for_testing(fake)  # type: ignore[arg-type]
    try:
        tree, mcda = _mcda_ahp_specs()
        out = tools.build_mcda(tree, mcda, dry_run=False)
        assert out.written is True and out.weight_source == "ahp"
        import json as _json
        spec = _json.loads(fake.mcda_spec)
        assert spec["weightSource"] == "ahp"
        assert spec["ahpMatrix"] == [[1.0, 2.0, 4.0], [0.5, 1.0, 2.0], [0.25, 0.5, 1.0]]
    finally:
        tools.set_bridge_for_testing(None)


def test_build_mcda_ahp_missing_matrix_raises(bridge: _FakeBridge) -> None:
    tree, mcda = _mcda_ahp_specs()
    mcda.ahp_matrix = None
    with pytest.raises(ValueError, match="requires ahp_matrix"):
        tools.build_mcda(tree, mcda)


def test_build_mcda_ahp_bad_size_raises(bridge: _FakeBridge) -> None:
    tree, mcda = _mcda_ahp_specs()
    mcda.ahp_matrix = [[1.0, 2.0], [0.5, 1.0]]  # 2x2, need 3x3 (financial + 2 criteria)
    with pytest.raises(ValueError, match="3x3"):
        tools.build_mcda(tree, mcda)


def test_run_scenarios_compares(bridge: _FakeBridge) -> None:
    from modelchoice_mcp.schemas import ScenarioComparison, ScenarioSpec

    out = tools.run_scenarios([
        ScenarioSpec(name="Wet likely", edits=[
            EditOp(op="set_probability", node_id="C", branch_name="Wet", value=0.9),
            EditOp(op="set_probability", node_id="C", branch_name="Dry", value=0.1)]),
        ScenarioSpec(name="Sell cheaper", edits=[
            EditOp(op="set_branch_value", node_id="D", branch_name="Sell", value=-30)]),
    ])
    assert isinstance(out, ScenarioComparison)
    assert out.objective == "maximize"
    assert out.baseline_ev == 50.0 and out.baseline_optimal_path == ["Sell"]
    by_name = {s.name: s for s in out.scenarios}
    # Wet 0.9: Drill EV = .9*50 + .1*-100 = 35 -> still < Sell 50, decision holds.
    assert by_name["Wet likely"].expected_value == 50.0
    assert by_name["Wet likely"].decision_changed is False
    # Sell -30: chance EV -25 > -30, so Drill becomes optimal -> flip.
    assert by_name["Sell cheaper"].expected_value == -25.0
    assert by_name["Sell cheaper"].decision_changed is True
    assert by_name["Sell cheaper"].delta_vs_baseline == -75.0
    assert out.best_scenario == "Wet likely"


def test_edit_tree_add_option_flips_decision() -> None:
    fake = _FakeBridge({"MC_Tree_1": _MODEL})
    tools.set_bridge_for_testing(fake)  # type: ignore[arg-type]
    try:
        # Baseline optimum is 'Sell' at EV 50. Add a 'License' strategy worth 60.
        out = tools.edit_tree(
            [EditOp(op="add_option", node_id="D", name="License", value=60)]
        )
        assert out.optimal_path == ["License"]
        assert out.expected_value == 60.0
        # An outcome terminal was auto-created and linked.
        assert '"name": "License"' in out.model_json
        assert '"T_License"' in out.model_json
    finally:
        tools.set_bridge_for_testing(None)


def test_edit_tree_add_branch_to_chance() -> None:
    fake = _FakeBridge({"MC_Tree_1": _MODEL})
    tools.set_bridge_for_testing(fake)  # type: ignore[arg-type]
    try:
        out = tools.edit_tree(
            [EditOp(op="add_branch", node_id="C", name="Marginal",
                    value=0, probability=0.5)]
        )
        # C now has three outcomes (Dry, Wet, Marginal).
        import json as _json
        model = _json.loads(out.model_json)
        assert len(model["Nodes"]["C"]["branches"]) == 3
    finally:
        tools.set_bridge_for_testing(None)


def test_edit_tree_add_branch_needs_probability() -> None:
    from modelchoice_mcp.tree import TreeParseError

    tools.set_bridge_for_testing(_FakeBridge({"MC_Tree_1": _MODEL}))  # type: ignore[arg-type]
    try:
        with pytest.raises(TreeParseError):
            tools.edit_tree([EditOp(op="add_branch", node_id="C", name="Marginal", value=0)])
    finally:
        tools.set_bridge_for_testing(None)


def test_edit_tree_add_option_wrong_node_kind_raises() -> None:
    from modelchoice_mcp.tree import TreeParseError

    tools.set_bridge_for_testing(_FakeBridge({"MC_Tree_1": _MODEL}))  # type: ignore[arg-type]
    try:
        # 'add_option' onto a chance node should fail.
        with pytest.raises(TreeParseError):
            tools.edit_tree([EditOp(op="add_option", node_id="C", name="X", value=1)])
    finally:
        tools.set_bridge_for_testing(None)


def test_edit_tree_remove_branch() -> None:
    fake = _FakeBridge({"MC_Tree_1": _MODEL})
    tools.set_bridge_for_testing(fake)  # type: ignore[arg-type]
    try:
        # Remove 'Sell'; only 'Drill' (-> chance EV -25) remains.
        out = tools.edit_tree(
            [EditOp(op="remove_branch", node_id="D", branch_name="Sell")]
        )
        assert out.optimal_path == ["Drill"]
        assert out.expected_value == -25.0
        assert '"name": "Sell"' not in out.model_json
    finally:
        tools.set_bridge_for_testing(None)


def test_edit_tree_remove_last_branch_raises() -> None:
    from modelchoice_mcp.tree import TreeParseError

    tools.set_bridge_for_testing(_FakeBridge({"MC_Tree_1": _MODEL}))  # type: ignore[arg-type]
    try:
        with pytest.raises(TreeParseError):
            tools.edit_tree([
                EditOp(op="remove_branch", node_id="D", branch_name="Sell"),
                EditOp(op="remove_branch", node_id="D", branch_name="Drill"),
            ])
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
        "force_to_outcome": "MC_ForceToOutcome_Auto",
    }
    for name, cmd in cmds.items():
        out = tools.run_decision_report(name)
        assert out.command == cmd
    # force_to_outcome finds its primary sheet
    fto = tools.run_decision_report("force_to_outcome")
    assert fto.primary_sheet == "MC_ForceOutcome"
    # two_way_sensitivity maps + finds its grid sheet
    tw = tools.run_decision_report("two_way_sensitivity")
    assert tw.command == "MC_TwoWaySensitivity_Auto"
    assert tw.primary_sheet == "MC_TwoWaySens"


def test_import_precisiontree(bridge: _FakeBridge) -> None:
    from modelchoice_mcp.schemas import ImportResult

    out = tools.import_precisiontree(r"C:\models\oil.xlsx")
    assert isinstance(out, ImportResult)
    assert bridge.imported_path == r"C:\models\oil.xlsx"
    assert out.workbook == "oil_ModelChoice.xlsx"
    assert out.trees == ["MC_Tree_1"]


def test_run_risk_profile(bridge: _FakeBridge) -> None:
    out = tools.run_risk_profile()
    assert bridge.last_command == "MC_RiskProfile_Auto"
    assert out.sheets == ["MC_RiskProfile"]
    by_name = {s.name: s for s in out.series}
    assert set(by_name) == {"Drill", "Sell"}
    assert by_name["Drill"].expected_value == -25.0
    assert by_name["Drill"].minimum == -100.0
    assert by_name["Drill"].maximum == 50.0
    assert by_name["Drill"].std_dev == 75.0
    assert by_name["Sell"].expected_value == 50.0
    assert by_name["Sell"].std_dev == 0.0


def test_run_sensitivity(bridge: _FakeBridge) -> None:
    out = tools.run_sensitivity()
    assert bridge.last_command == "MC_SensitivityAnalysis_Auto"
    assert out.sheets == ["MC_SensReport", "MC_Tornado"]
    assert out.report_rows  # MC_SensReport rows came through
    assert "tornado" in out.note.lower()
