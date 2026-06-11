"""MCP tools (Phase 1, read-only) over the ModelChoice bridge.

Each tool obtains a process-global :class:`ModelChoiceBridge` via
``get_bridge()`` (lazy — the first call attaches to Excel). Tests inject
a fake with ``set_bridge_for_testing`` to avoid touching COM.
"""

from __future__ import annotations

from typing import Annotated, Any

from pydantic import Field

from modelchoice_mcp.bridge import ModelChoiceBridge
from modelchoice_mcp.schemas import (
    AnalysisRun,
    BranchView,
    BuildTreeResult,
    EvpiResult,
    KeyValue,
    NodeDiff,
    NodeResultView,
    NodeView,
    RobustnessSummary,
    RollbackVerification,
    RollupResponse,
    SensitivityReport,
    SheetData,
    TreeList,
    TreeSpec,
    TreeStructure,
    TreeSummary,
)
from modelchoice_mcp.server import mcp
from modelchoice_mcp.tree import (
    Branch,
    DecisionTree,
    Node,
    TreeParseError,
    parse_model,
    rollup,
    to_model_json,
)

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


def _spec_to_tree(spec: TreeSpec) -> DecisionTree:
    nodes: dict[str, Node] = {}
    for ns in spec.nodes:
        kind = ns.type.lower()
        if kind == "terminal":
            nodes[ns.id] = Node(id=ns.id, name=ns.name, kind="terminal", value=ns.value)
        elif kind in ("chance", "decision"):
            branches = [
                Branch(
                    name=b.name,
                    child_id=b.child_id,
                    value=b.value,
                    probability=(b.probability if kind == "chance" else None),
                )
                for b in ns.branches
            ]
            nodes[ns.id] = Node(id=ns.id, name=ns.name, kind=kind, branches=branches)
        else:
            raise TreeParseError(f"node {ns.id!r}: unknown type {ns.type!r}")
    return DecisionTree(
        root_id=spec.root_id,
        nodes=nodes,
        maximize=spec.maximize,
        model_name=spec.model_name,
    )


@mcp.tool(
    description=(
        "ModelChoice: Build a decision tree from a structured description and "
        "(optionally) write it into Excel. Assemble `spec` from the decision: "
        "decision / chance / terminal nodes, with branch probabilities and "
        "cash flows (branch `value`) and terminal payoffs. The tool serializes "
        "it to ModelChoice's model format, validates it (rolls it back — "
        "catching cycles, missing children, bad structure), and returns the "
        "rolled-back EV + optimal policy so you can confirm before writing. "
        "dry_run=True (default) previews without touching Excel; dry_run=False "
        "writes the tree into the workbook and renders it (needs the "
        "ModelChoice add-in loaded). This is the build-from-a-prompt path."
    )
)
def build_tree(
    spec: TreeSpec,
    dry_run: bool = True,
    workbook_name: str | None = None,
) -> BuildTreeResult:
    tree = _spec_to_tree(spec)
    model_json = to_model_json(tree)
    # Validate by round-tripping through the parser + roller.
    parsed = parse_model(model_json)
    r = rollup(parsed)

    direction = "maximize" if tree.maximize else "minimize"
    if r.optimal_path:
        rec = (
            f"Optimal decision ({direction} EV): take {' → '.join(r.optimal_path)}. "
            f"Expected value {r.expected_value:,.2f}."
        )
    else:
        rec = f"Expected value {r.expected_value:,.2f} ({direction})."

    written = False
    sheet: str | None = None
    if not dry_run:
        bridge = get_bridge()
        sheet = bridge.write_tree(model_json, workbook=workbook_name)
        bridge.render_tree(sheet, workbook_name)
        written = True

    return BuildTreeResult(
        written=written,
        sheet=sheet,
        node_count=len(parsed.nodes),
        expected_value=r.expected_value,
        optimal_path=r.optimal_path,
        recommendation=rec,
        model_json=model_json,
    )


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


def _key_values(rows: list[list[Any]]) -> list[KeyValue]:
    """Extract label→value pairs from a two-column report sheet: a row
    contributes a pair when its first cell is a non-empty string label and
    a later cell holds the value."""
    pairs: list[KeyValue] = []
    for row in rows:
        if not row:
            continue
        label = row[0]
        if not isinstance(label, str) or not label.strip():
            continue
        value = next((c for c in row[1:] if c is not None and c != ""), None)
        if value is not None:
            pairs.append(KeyValue(label=label.strip(), value=value))
    return pairs


@mcp.tool(
    description=(
        "ModelChoice: Run the robustness ('break the decision') analysis and "
        "return a structured read — the verdict, robustness score, and the "
        "minimum change to an input that flips the optimal decision, plus all "
        "label→value pairs from the verdict sheet. Higher distance / score = "
        "more robust. Requires the ModelChoice add-in loaded with a tree open."
    )
)
def run_robustness(workbook_name: str | None = None) -> RobustnessSummary:
    bridge = get_bridge()
    run = bridge.run_analysis("MC_Robustness_Auto", workbook_name)
    all_sheets = run.get("sheets", [])
    rb_sheets = [s for s in all_sheets if s.startswith("MC_RB_")]
    rows: list[list[Any]] = []
    if "MC_RB_Verdict" in all_sheets:
        rows = bridge.read_sheet("MC_RB_Verdict", workbook_name)
    pairs = _key_values(rows)

    def _find(*needles: str) -> KeyValue | None:
        for p in pairs:
            low = p.label.lower()
            if any(n in low for n in needles):
                return p
        return None

    verdict = _find("verdict")
    score = _find("score")
    dist = _find("distance", "flip")
    dist_val = dist.value if dist and isinstance(dist.value, (int, float)) else None

    return RobustnessSummary(
        verdict=str(verdict.value) if verdict else None,
        robustness_score=str(score.value) if score else None,
        min_distance=float(dist_val) if dist_val is not None else None,
        details=pairs,
        sheets=rb_sheets,
        note=(
            "Read from the robustness report. Other sheets (vulnerability, "
            "boundary, threats) are available via read_sheet."
        ),
    )


@mcp.tool(
    description=(
        "ModelChoice: Run one-way sensitivity analysis and return the report. "
        "Auto-selects the model's variable inputs and ranks them by how much "
        "they swing the expected value (tornado order, largest first) — showing "
        "which assumptions the decision is most sensitive to. Returns the "
        "MC_SensReport rows, the baseline EV, and the report sheets "
        "(MC_SensReport / MC_Tornado / …). Requires the ModelChoice add-in "
        "loaded with a tree open."
    )
)
def run_sensitivity(workbook_name: str | None = None) -> SensitivityReport:
    bridge = get_bridge()
    run = bridge.run_analysis("MC_SensitivityAnalysis_Auto", workbook_name)
    all_sheets = run.get("sheets", [])
    sens_sheets = [s for s in all_sheets if s.startswith(("MC_Sens", "MC_Tornado", "MC_Spider"))]
    rows: list[list[Any]] = []
    if "MC_SensReport" in all_sheets:
        rows = bridge.read_sheet("MC_SensReport", workbook_name)

    baseline: float | None = None
    for row in rows:
        if row and isinstance(row[0], str) and "baseline" in row[0].lower():
            val = next(
                (c for c in row[1:] if isinstance(c, (int, float)) and not isinstance(c, bool)),
                None,
            )
            if val is not None:
                baseline = float(val)
                break

    return SensitivityReport(
        baseline_ev=baseline,
        report_rows=rows,
        sheets=sens_sheets,
        note=(
            "Variables are tornado-ordered (largest EV swing first). "
            "The MC_Tornado sheet has the chart; read other sheets with read_sheet."
        ),
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
    "build_tree",
    "get_bridge",
    "get_tree",
    "list_trees",
    "read_sheet",
    "roll_up",
    "run_analysis",
    "run_evpi",
    "run_robustness",
    "run_sensitivity",
    "set_bridge_for_testing",
    "verify_rollback",
]
