"""MCP tools (Phase 1, read-only) over the ModelChoice bridge.

Each tool obtains a process-global :class:`ModelChoiceBridge` via
``get_bridge()`` (lazy — the first call attaches to Excel). Tests inject
a fake with ``set_bridge_for_testing`` to avoid touching COM.
"""

from __future__ import annotations

from typing import Annotated

from pydantic import Field

from modelchoice_mcp.bridge import ModelChoiceBridge
from modelchoice_mcp.schemas import (
    AnalysisRun,
    BranchView,
    EvpiResult,
    NodeDiff,
    NodeResultView,
    NodeView,
    RollbackVerification,
    RollupResponse,
    SheetData,
    TreeList,
    TreeStructure,
    TreeSummary,
)
from modelchoice_mcp.server import mcp
from modelchoice_mcp.tree import DecisionTree, parse_model, rollup

# Friendly analysis name -> ModelChoice headless ExcelCommand.
_ANALYSES: dict[str, str] = {
    "risk_profile": "MC_RiskProfile_Auto",
    "robustness": "MC_Robustness_Auto",
    "sensitivity": "MC_SensitivityAnalysis_Auto",
    "strategy_table": "MC_ExportStrategyTable_Auto",
    "policy_suggestion": "MC_PolicySuggestion_Auto",
    "decision_brief": "MC_DecisionBrief_Auto",
    "mcda_report": "MC_McdaReport_Auto",
    "evpi": "MC_EVPI_Auto",
}

_bridge: ModelChoiceBridge | None = None


def get_bridge() -> ModelChoiceBridge:
    global _bridge
    if _bridge is None:
        _bridge = ModelChoiceBridge()
    return _bridge


def set_bridge_for_testing(bridge: ModelChoiceBridge | None) -> None:
    global _bridge
    _bridge = bridge


def _counts(tree: DecisionTree) -> tuple[int, int, int]:
    d = sum(1 for n in tree.nodes.values() if n.kind == "decision")
    c = sum(1 for n in tree.nodes.values() if n.kind == "chance")
    t = sum(1 for n in tree.nodes.values() if n.kind == "terminal")
    return d, c, t


@mcp.tool(
    description=(
        "ModelChoice: List the decision trees stored in a workbook, with a "
        "summary of each (model name, root node, and node-type counts). Trees "
        "live in the workbook's hidden ModelChoice store; the add-in does not "
        "need to be loaded. Omit workbook_name for the active workbook."
    )
)
def list_trees(workbook_name: str | None = None) -> TreeList:
    trees = get_bridge().list_trees(workbook_name)
    summaries: list[TreeSummary] = []
    for name, model_json in trees.items():
        t = parse_model(model_json)
        d, c, term = _counts(t)
        root = t.nodes[t.root_id]
        summaries.append(
            TreeSummary(
                name=name,
                model_name=t.model_name,
                root_id=t.root_id,
                root_name=root.name,
                node_count=len(t.nodes),
                decision_count=d,
                chance_count=c,
                terminal_count=term,
            )
        )
    return TreeList(workbook=workbook_name, count=len(summaries), trees=summaries)


@mcp.tool(
    description=(
        "ModelChoice: Read the full structure of one decision tree — every "
        "node (decision / chance / terminal) with its branches, probabilities, "
        "and branch values. Pass tree_name to pick a specific tree, else the "
        "first/active one. Read-only."
    )
)
def get_tree(
    tree_name: Annotated[
        str | None, Field(description="Tree sheet name, e.g. 'MC_Tree_1'. Omit for the first tree.")
    ] = None,
    workbook_name: str | None = None,
) -> TreeStructure:
    bridge = get_bridge()
    trees = bridge.list_trees(workbook_name)
    if tree_name is None:
        tree_name = next(iter(trees))
    t = parse_model(trees[tree_name])
    nodes = [
        NodeView(
            id=n.id,
            name=n.name,
            kind=n.kind,
            value=n.value,
            branches=[
                BranchView(
                    name=b.name, child_id=b.child_id, value=b.value, probability=b.probability
                )
                for b in n.branches
            ],
        )
        for n in t.nodes.values()
    ]
    return TreeStructure(
        name=tree_name,
        model_name=t.model_name,
        root_id=t.root_id,
        maximize=t.maximize,
        node_count=len(t.nodes),
        nodes=nodes,
    )


@mcp.tool(
    description=(
        "ModelChoice: Roll a decision tree back to its expected values and "
        "OPTIMAL POLICY — the decision recommendation. Returns the root "
        "expected value, the optimal sequence of decisions, a plain-English "
        "recommendation, and the per-node expected values. Computed directly "
        "from the stored model (terminal payoff = accumulated branch values; "
        "chance = probability-weighted EV; decision = best EV by maximize/"
        "minimize) — reproduces ModelChoice's rollback without the add-in."
    )
)
def roll_up(
    tree_name: Annotated[
        str | None, Field(description="Tree sheet name. Omit for the first tree.")
    ] = None,
    workbook_name: str | None = None,
) -> RollupResponse:
    bridge = get_bridge()
    trees = bridge.list_trees(workbook_name)
    if tree_name is None:
        tree_name = next(iter(trees))
    t = parse_model(trees[tree_name])
    r = rollup(t)

    direction = "maximize" if t.maximize else "minimize"
    if r.optimal_path:
        choices = " → ".join(r.optimal_path)
        rec = (
            f"Optimal decision ({direction} EV): take {choices}. "
            f"Expected value {r.expected_value:,.2f}."
        )
    else:
        rec = (
            f"No decision nodes to optimize; expected value {r.expected_value:,.2f} "
            f"({direction})."
        )

    nodes = [
        NodeResultView(
            id=nr.id,
            name=nr.name,
            kind=nr.kind,
            expected_value=nr.expected_value,
            optimal_branch_name=nr.optimal_branch_name,
        )
        for nr in r.node_results.values()
    ]
    return RollupResponse(
        name=tree_name,
        model_name=t.model_name,
        maximize=t.maximize,
        expected_value=r.expected_value,
        optimal_path=r.optimal_path,
        recommendation=rec,
        nodes=nodes,
    )


@mcp.tool(
    description=(
        "ModelChoice: Cross-check our rollback against ModelChoice's own. The "
        "add-in writes each node's rolled-back EV into an MC_V_<nodeId> named "
        "range when it renders a tree; this compares those cells to our Python "
        "rollback and reports any node that differs beyond `tolerance`. A clean "
        "result is strong evidence the read+rollback is faithful. If the tree "
        "hasn't been rendered (no MC_V_ cells), `rendered` is false — open the "
        "tree in ModelChoice so it renders, then re-run."
    )
)
def verify_rollback(
    tree_name: Annotated[
        str | None, Field(description="Tree sheet name. Omit for the first tree.")
    ] = None,
    workbook_name: str | None = None,
    tolerance: Annotated[
        float, Field(gt=0, description="Max allowed |computed - cell| difference. Default 0.01.")
    ] = 0.01,
) -> RollbackVerification:
    bridge = get_bridge()
    trees = bridge.list_trees(workbook_name)
    if tree_name is None:
        tree_name = next(iter(trees))
    r = rollup(parse_model(trees[tree_name]))
    cells = bridge.read_node_values(workbook_name)

    common = [nid for nid in r.node_results if nid in cells]
    mismatches: list[NodeDiff] = []
    max_abs = 0.0
    for nid in common:
        computed = r.node_results[nid].expected_value
        cell = cells[nid]
        diff = computed - cell
        max_abs = max(max_abs, abs(diff))
        if abs(diff) > tolerance:
            mismatches.append(
                NodeDiff(
                    node_id=nid, name=r.node_results[nid].name,
                    computed=computed, cell=cell, diff=diff,
                )
            )

    rendered = len(cells) > 0
    matches = len(common) - len(mismatches)
    if not rendered:
        verdict = "Tree not rendered — no MC_V_ cells found. Open it in ModelChoice first."
    elif not common:
        verdict = "No overlapping node IDs between the model and the rendered cells."
    elif not mismatches:
        verdict = f"Rollback verified: all {matches} compared nodes match (max diff {max_abs:.2g})."
    else:
        verdict = f"{len(mismatches)} of {len(common)} nodes differ beyond {tolerance}."

    return RollbackVerification(
        name=tree_name,
        rendered=rendered,
        compared_count=len(common),
        matches=matches,
        max_abs_diff=max_abs,
        mismatches=mismatches,
        verdict=verdict,
    )


@mcp.tool(
    description=(
        "ModelChoice: Compute the Expected Value of Perfect Information (EVPI) "
        "for the active decision tree — the most a decision-maker should pay "
        "for perfect information about all uncertainties before deciding. "
        "Drives ModelChoice's headless analysis (MC_EVPI_Auto) and reads the "
        "result. Requires the ModelChoice add-in loaded in Excel with a tree "
        "open. EVPI is not defined for MCDA models."
    )
)
def run_evpi(workbook_name: str | None = None) -> EvpiResult:
    r = get_bridge().run_evpi(workbook_name)
    evpi = r.get("evpi")
    if evpi is None:
        interp = "EVPI could not be read."
    elif evpi <= 1e-9:
        interp = (
            "EVPI is essentially zero — perfect information would not change the "
            "optimal decision, so don't pay for more information."
        )
    else:
        interp = (
            f"Perfect information about all uncertainties is worth up to "
            f"{evpi:,.2f}; that's the ceiling on what to spend gathering it."
        )
    return EvpiResult(
        model_name=r.get("model_name"),
        objective=r.get("objective"),
        optimal_ev=r.get("optimal_ev"),
        evpi=evpi,
        value_with_perfect_info=r.get("value_with_perfect_info"),
        interpretation=interp,
    )


@mcp.tool(
    description=(
        "ModelChoice: Run a decision-analysis on the active tree and report the "
        "result sheets it produced. `analysis` is one of: 'risk_profile', "
        "'robustness' (how much inputs must change to flip the decision), "
        "'sensitivity', 'strategy_table', 'policy_suggestion', 'decision_brief', "
        "'mcda_report', 'evpi'. Drives ModelChoice's headless command and lists "
        "the new sheets — read them with read_sheet. Requires the ModelChoice "
        "add-in loaded in Excel with a tree open."
    )
)
def run_analysis(
    analysis: Annotated[
        str, Field(description="One of: " + ", ".join(sorted(_ANALYSES)))
    ],
    workbook_name: str | None = None,
) -> AnalysisRun:
    key = analysis.lower()
    if key not in _ANALYSES:
        raise ValueError(
            f"Unknown analysis {analysis!r}. Choose from: {', '.join(sorted(_ANALYSES))}."
        )
    command = _ANALYSES[key]
    r = get_bridge().run_analysis(command, workbook_name)
    new = r.get("new_sheets", [])
    note = (
        f"{analysis} wrote: {', '.join(new)} — read with read_sheet."
        if new
        else f"{analysis} ran but added no new sheet; it may have updated an existing one."
    )
    return AnalysisRun(
        analysis=key,
        command=command,
        new_sheets=new,
        sheets=r.get("sheets", []),
        note=note,
    )


@mcp.tool(
    description=(
        "ModelChoice: Read a worksheet's used range (capped) as rows of cell "
        "values — for pulling back the result sheet an analysis produced (e.g. "
        "'MC_EVPI', 'MC_RB_Verdict', a sensitivity report). Returns numbers, "
        "text, or null per cell."
    )
)
def read_sheet(
    sheet_name: Annotated[str, Field(description="Worksheet name to read.")],
    workbook_name: str | None = None,
    max_rows: Annotated[
        int, Field(ge=1, le=2000, description="Max rows to return (default 200).")
    ] = 200,
    max_cols: Annotated[
        int, Field(ge=1, le=100, description="Max columns to return (default 20).")
    ] = 20,
) -> SheetData:
    rows = get_bridge().read_sheet(sheet_name, workbook_name, max_rows=max_rows, max_cols=max_cols)
    return SheetData(
        sheet=sheet_name,
        row_count=len(rows),
        rows=rows,
        truncated=len(rows) >= max_rows,
    )


__all__ = [
    "get_bridge",
    "get_tree",
    "list_trees",
    "read_sheet",
    "roll_up",
    "run_analysis",
    "run_evpi",
    "set_bridge_for_testing",
    "verify_rollback",
]
