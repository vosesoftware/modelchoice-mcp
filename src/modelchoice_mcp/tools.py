"""MCP tools (Phase 1, read-only) over the ModelChoice bridge.

Each tool obtains a process-global :class:`ModelChoiceBridge` via
``get_bridge()`` (lazy — the first call attaches to Excel). Tests inject
a fake with ``set_bridge_for_testing`` to avoid touching COM.
"""

from __future__ import annotations

import dataclasses
import json
from typing import Annotated, Any

from pydantic import Field

from modelchoice_mcp.bridge import ModelChoiceBridge
from modelchoice_mcp.schemas import (
    AnalysisRun,
    BranchView,
    BuildTreeResult,
    CloseWorkbookResult,
    ControlPanelResult,
    DecisionReport,
    EditOp,
    EviiResult,
    EvpiResult,
    ImportResult,
    InputDistributionResult,
    KeyValue,
    LicenseStatus,
    McdaBuildResult,
    McdaSpec,
    NodeDiff,
    NodeResultView,
    NodeView,
    OpenWorkbookResult,
    RiskProfileReport,
    RiskProfileSeries,
    RobustnessSummary,
    RollbackVerification,
    RollupResponse,
    ScenarioComparison,
    ScenarioOutcome,
    ScenarioSpec,
    SensitivityReport,
    SheetData,
    TreeExport,
    TreeList,
    TreeSpec,
    TreeStructure,
    TreeSummary,
    UtilityResult,
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


_AHP_RANDOM_INDEX = [
    0.0, 0.0, 0.58, 0.90, 1.12, 1.24, 1.32, 1.41, 1.45, 1.49,
    1.51, 1.48, 1.56, 1.57, 1.59,
]


def _ahp_weights(matrix: list[list[float]]) -> tuple[list[float], float]:
    """AHP principal-eigenvector weights + consistency ratio, mirroring the
    add-in's AhpCalculator (power iteration; Saaty's RI table). `matrix` is an
    n-by-n positive reciprocal pairwise-comparison matrix; returns (weights, CR)."""
    n = len(matrix)
    w = [1.0 / n] * n
    for _ in range(100):
        t = [sum(matrix[i][j] * w[j] for j in range(n)) for i in range(n)]
        s = sum(t)
        if s < 1e-15:
            raise ValueError("ahp_matrix is degenerate (zero column sum).")
        t = [x / s for x in t]
        converged = max(abs(t[i] - w[i]) for i in range(n)) < 1e-10
        w = t
        if converged:
            break
    if n <= 2:
        return w, 0.0
    lam = 0.0
    for i in range(n):
        aw = sum(matrix[i][j] * w[j] for j in range(n))
        if w[i] > 1e-15:
            lam += aw / w[i]
    lam /= n
    ci = (lam - n) / (n - 1)
    ri = _AHP_RANDOM_INDEX[min(n, len(_AHP_RANDOM_INDEX)) - 1]
    return w, (ci / ri if ri > 1e-15 else 0.0)


def _validate_ahp_matrix(matrix: list[list[float]] | None, n: int) -> list[list[float]]:
    """Validate an AHP matrix is present, n-by-n, and strictly positive."""
    if not matrix:
        raise ValueError("weight_source='ahp' requires ahp_matrix.")
    if len(matrix) != n or any(len(row) != n for row in matrix):
        raise ValueError(
            f"ahp_matrix must be {n}x{n} (financial + {n - 1} criteria), "
            "ordered [financial, then criteria in spec order]."
        )
    for i, row in enumerate(matrix):
        for j, v in enumerate(row):
            if isinstance(v, bool) or not isinstance(v, (int, float)) or v <= 0:
                raise ValueError(f"ahp_matrix[{i}][{j}] must be a positive number.")
    return matrix


@mcp.tool(
    description=(
        "ModelChoice: Build a MULTI-CRITERIA (MCDA) decision model — when the "
        "choice isn't pure money. Give the tree structure (`tree`, same as "
        "build_tree) plus `mcda`: the criteria (each with discrete ordinal "
        "options worst→best, a weight, and a direction), the financial weight, "
        "an aggregation method (weighted_sum / weighted_geometric / waspas), and "
        "each terminal's option per criterion (`terminal_scores`). The tool "
        "validates structure + weights + that every score references a real "
        "criterion option. Weights are entered directly (weight_source='direct') "
        "OR derived from an AHP pairwise-comparison matrix (weight_source='ahp' "
        "with `ahp_matrix` — the tool computes the eigenvector weights + "
        "consistency ratio and the add-in recomputes them authoritatively). "
        "dry_run=True (default) previews; dry_run=False builds the tree, sets MCDA "
        "mode, and renders — the add-in computes the composite scores. Needs the "
        "add-in loaded (a build with the MCDA command)."
    )
)
def build_mcda(
    tree: TreeSpec,
    mcda: McdaSpec,
    dry_run: bool = True,
    workbook_name: str | None = None,
) -> McdaBuildResult:
    t = _spec_to_tree(tree)
    model_json = to_model_json(t)
    parse_model(model_json)  # structural validation (cycles, missing children)

    if not mcda.criteria:
        raise ValueError("MCDA needs at least one criterion.")
    crit_by_id: dict[str, Any] = {}
    for c in mcda.criteria:
        if len(c.options) < 2:
            raise ValueError(f"criterion {c.id!r} needs >=2 options.")
        crit_by_id[c.id] = c
    term_ids = {n.id for n in tree.nodes if n.type.lower() == "terminal"}
    for tid, scores in mcda.terminal_scores.items():
        if tid not in term_ids:
            raise ValueError(f"terminal_scores references unknown terminal {tid!r}.")
        for cid, opt in scores.items():
            if cid not in crit_by_id:
                raise ValueError(f"terminal {tid!r} scores unknown criterion {cid!r}.")
            if opt not in crit_by_id[cid].options:
                raise ValueError(
                    f"terminal {tid!r} criterion {cid!r}: {opt!r} is not one of "
                    f"{crit_by_id[cid].options}."
                )

    crit_names = [c.name for c in mcda.criteria]
    scored = len(mcda.terminal_scores)
    ws = mcda.weight_source.strip().lower()

    # Resolve effective weights for preview/reporting. For AHP the matrix is
    # authoritative (the add-in recomputes it on commit); we mirror the math
    # here so the preview shows the derived weights + consistency ratio.
    consistency_ratio: float | None = None
    ahp_matrix: list[list[float]] | None = None
    if ws == "ahp":
        ahp_matrix = _validate_ahp_matrix(mcda.ahp_matrix, len(mcda.criteria) + 1)
        derived, consistency_ratio = _ahp_weights(ahp_matrix)
        financial_w, crit_w = derived[0], derived[1:]
        weight_note = f" AHP weights (CR={consistency_ratio:.3f}" + (
            "; INCONSISTENT, CR>0.10 — consider revising judgments)"
            if consistency_ratio > 0.10
            else ")"
        )
    else:
        financial_w = mcda.financial_weight
        crit_w = [c.weight for c in mcda.criteria]
        total = financial_w + sum(crit_w)
        weight_note = "" if abs(total - 1.0) < 1e-3 else f" (weights sum {total:.3f}, not 1.0)"

    weights_kv = [KeyValue(label="Financial", value=round(financial_w, 4))] + [
        KeyValue(label=c.name, value=round(crit_w[i], 4)) for i, c in enumerate(mcda.criteria)
    ]

    if dry_run:
        return McdaBuildResult(
            written=False,
            sheet=None,
            node_count=len(t.nodes),
            criteria=crit_names,
            terminals_scored=scored,
            weight_source=ws,
            weights=weights_kv,
            consistency_ratio=consistency_ratio,
            note=(
                f"Validated MCDA preview: {len(crit_names)} criteria, {scored} terminals "
                f"scored;{weight_note}. Composite scores are computed by the add-in on commit."
            ),
        )

    bridge = get_bridge()
    sheet = bridge.write_tree(model_json, workbook=workbook_name)
    bridge.render_tree(sheet, workbook_name)
    spec: dict[str, Any] = {
        "financialWeight": mcda.financial_weight,
        "aggregation": mcda.aggregation,
        "waspasLambda": mcda.waspas_lambda,
        "weightSource": ws,
        "criteria": [
            {
                "id": c.id,
                "name": c.name,
                "weight": c.weight,
                "maximize": c.maximize,
                "options": c.options,
            }
            for c in mcda.criteria
        ],
        "terminalScores": mcda.terminal_scores,
    }
    if ahp_matrix is not None:
        spec["ahpMatrix"] = ahp_matrix
    bridge.apply_mcda(json.dumps(spec), workbook_name)
    return McdaBuildResult(
        written=True,
        sheet=sheet,
        node_count=len(t.nodes),
        criteria=crit_names,
        terminals_scored=scored,
        weight_source=ws,
        weights=weights_kv,
        consistency_ratio=consistency_ratio,
        note=(
            f"Built MCDA model on {sheet!r}: {len(crit_names)} criteria, {scored} "
            f"terminals scored, aggregation {mcda.aggregation}{weight_note}."
        ),
    )


def _apply_edits(tree: DecisionTree, edits: list[EditOp]) -> DecisionTree:
    """Return a new tree with the edits applied. Frozen dataclasses are
    rebuilt via dataclasses.replace; unknown ops or missing targets raise
    TreeParseError."""
    nodes = dict(tree.nodes)
    maximize = tree.maximize
    model_name = tree.model_name

    for e in edits:
        op = e.op.lower()
        if op == "set_objective":
            if e.maximize is None:
                raise TreeParseError("set_objective needs `maximize`.")
            maximize = e.maximize
            continue

        if not e.node_id or e.node_id not in nodes:
            raise TreeParseError(f"edit {op!r}: no node {e.node_id!r}.")
        node = nodes[e.node_id]

        if op == "rename_node":
            if not e.name:
                raise TreeParseError("rename_node needs `name`.")
            nodes[e.node_id] = dataclasses.replace(node, name=e.name)
        elif op == "set_terminal_value":
            if node.kind != "terminal":
                raise TreeParseError(f"{e.node_id!r} is not a terminal node.")
            if e.value is None:
                raise TreeParseError("set_terminal_value needs `value`.")
            nodes[e.node_id] = dataclasses.replace(node, value=e.value)
        elif op in ("set_probability", "set_branch_value", "rename_branch"):
            found = False
            new_branches = []
            for b in node.branches:
                if b.name == e.branch_name:
                    found = True
                    if op == "set_probability":
                        b = dataclasses.replace(b, probability=e.value)
                    elif op == "set_branch_value":
                        b = dataclasses.replace(b, value=e.value or 0.0)
                    else:  # rename_branch
                        if not e.name:
                            raise TreeParseError("rename_branch needs `name`.")
                        b = dataclasses.replace(b, name=e.name)
                new_branches.append(b)
            if not found:
                raise TreeParseError(
                    f"node {e.node_id!r} has no branch named {e.branch_name!r}."
                )
            nodes[e.node_id] = dataclasses.replace(node, branches=new_branches)
        elif op in ("add_option", "add_branch"):
            want_kind = "decision" if op == "add_option" else "chance"
            label = "option" if op == "add_option" else "outcome"
            if node.kind != want_kind:
                raise TreeParseError(
                    f"{op!r} needs a {want_kind} node, but {e.node_id!r} is {node.kind}."
                )
            if not e.name:
                raise TreeParseError(f"{op!r} needs `name` for the new {label}.")
            if any(b.name == e.name for b in node.branches):
                raise TreeParseError(
                    f"node {e.node_id!r} already has a branch named {e.name!r}."
                )
            if op == "add_branch" and e.probability is None:
                raise TreeParseError("add_branch needs `probability` for a chance outcome.")

            # Resolve the child: link an existing node, or auto-create a terminal.
            if e.child_id is not None:
                if e.child_id not in nodes:
                    raise TreeParseError(f"{op!r}: child node {e.child_id!r} does not exist.")
                child_id = e.child_id
            else:
                child_id = _fresh_terminal_id(nodes, e.name)
                nodes[child_id] = Node(
                    id=child_id, name=e.name, kind="terminal", value=0.0
                )

            new_branch = Branch(
                name=e.name,
                child_id=child_id,
                value=e.value or 0.0,
                probability=(e.probability if op == "add_branch" else None),
            )
            nodes[e.node_id] = dataclasses.replace(
                node, branches=[*node.branches, new_branch]
            )
        elif op in ("remove_branch", "remove_option"):
            if not any(b.name == e.branch_name for b in node.branches):
                raise TreeParseError(
                    f"node {e.node_id!r} has no branch named {e.branch_name!r}."
                )
            if len(node.branches) <= 1:
                raise TreeParseError(
                    f"can't remove the last branch of {e.node_id!r} — a decision/chance "
                    "node needs at least one."
                )
            kept = [b for b in node.branches if b.name != e.branch_name]
            nodes[e.node_id] = dataclasses.replace(node, branches=kept)
        else:
            raise TreeParseError(f"unknown edit op {e.op!r}.")

    return dataclasses.replace(
        tree, nodes=nodes, maximize=maximize, model_name=model_name
    )


def _fresh_terminal_id(nodes: dict[str, Node], label: str) -> str:
    """A node id not already in `nodes`, derived from a label."""
    slug = "".join(c if c.isalnum() else "_" for c in label).strip("_") or "node"
    base = f"T_{slug}"
    if base not in nodes:
        return base
    i = 1
    while f"{base}_{i}" in nodes:
        i += 1
    return f"{base}_{i}"


@mcp.tool(
    description=(
        "ModelChoice: Edit an existing decision tree in place — change "
        "probabilities, branch cash flows, terminal payoffs, node/branch "
        "labels, the objective, OR add/remove whole options and outcomes — then "
        "re-roll it. Pass `edits` as a list of operations. Value/label ops: "
        "'set_probability', 'set_branch_value', 'set_terminal_value', "
        "'rename_node', 'rename_branch', 'set_objective'. Structural ops: "
        "'add_option' (new strategy on a decision node — give node_id, name, "
        "value; omit child_id to auto-create its outcome), 'add_branch' (new "
        "outcome on a chance node — also needs probability), 'remove_branch' "
        "(drop an option/outcome by branch_name). "
        "Reads the named tree, applies the edits, validates by rolling back, and "
        "returns the new EV + optimal policy. dry_run=True (default) previews; "
        "dry_run=False writes and re-renders. The 'tweak it by talking' path."
    )
)
def edit_tree(
    edits: list[EditOp],
    tree_name: Annotated[
        str | None, Field(description="Tree sheet name. Omit for the first tree.")
    ] = None,
    dry_run: bool = True,
    workbook_name: str | None = None,
) -> BuildTreeResult:
    bridge = get_bridge()
    trees = bridge.list_trees(workbook_name)
    if tree_name is None:
        tree_name = next(iter(trees))
    if tree_name not in trees:
        raise TreeParseError(f"no tree {tree_name!r}; available: {', '.join(trees)}.")

    edited = _apply_edits(parse_model(trees[tree_name]), edits)
    model_json = to_model_json(edited)
    parsed = parse_model(model_json)
    r = rollup(parsed)

    direction = "maximize" if edited.maximize else "minimize"
    if r.optimal_path:
        rec = (
            f"After edits, optimal decision ({direction} EV): take "
            f"{' → '.join(r.optimal_path)}. Expected value {r.expected_value:,.2f}."
        )
    else:
        rec = f"After edits, expected value {r.expected_value:,.2f} ({direction})."

    written = False
    if not dry_run:
        bridge.write_tree(model_json, sheet_name=tree_name, workbook=workbook_name)
        bridge.render_tree(tree_name, workbook_name)
        written = True

    return BuildTreeResult(
        written=written,
        sheet=tree_name if written else None,
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
        "ModelChoice: Open a workbook (.xlsx) from disk in the running Excel so "
        "the other tools can act on it. Pass an absolute file path. Reports the "
        "workbook's sheets and any ModelChoice tree sheets it contains (so you "
        "can tell at a glance whether it's a decision-tree workbook). If a "
        "workbook with the same file name is already open, that one is used. "
        "Requires Excel running (with the ModelChoice add-in for rendering)."
    )
)
def open_workbook(
    path: Annotated[
        str,
        Field(description="Absolute path to the workbook, e.g. r'C:\\models\\oil.xlsx'."),
    ],
) -> OpenWorkbookResult:
    r = get_bridge().open_workbook(path)
    trees = r.get("trees", [])
    wb = r.get("workbook", "")
    note = (
        f"Opened {wb!r} with {len(trees)} ModelChoice tree(s): {', '.join(trees)}."
        if trees
        else f"Opened {wb!r} — no ModelChoice trees found (it's not a ModelChoice workbook, "
        "or the tree store is absent)."
    )
    return OpenWorkbookResult(
        workbook=wb, sheets=r.get("sheets", []), trees=trees, note=note
    )


@mcp.tool(
    description=(
        "ModelChoice: Report the ModelChoice add-in's licence state (fully "
        "licensed / trial / expired / not activated). Building and analysis "
        "ACTIONS require a FULL licence; reading trees works regardless. Use "
        "this to explain why an action was refused. Read-only."
    )
)
def license_status(workbook_name: str | None = None) -> LicenseStatus:
    s = get_bridge().license_status(workbook_name)
    if not s:
        return LicenseStatus(
            available=False,
            is_complete=False,
            actions_allowed=False,
            note=(
                "Could not read the licence status — the ModelChoice add-in isn't "
                "loaded, or is older than the MC_LicenseStatus_Auto command. "
                "Building/analysis actions are blocked until a licensed add-in reports in."
            ),
        )
    is_complete = bool(s.get("isComplete"))
    days = s.get("daysLeft")
    days_left = int(days) if isinstance(days, (int, float)) and not isinstance(days, bool) else None
    note = (
        "Fully licensed — all actions allowed."
        if is_complete
        else (
            f"Not fully licensed ({s.get('statusText') or 'see flags'}). Reading trees "
            "works; building and analysis actions are blocked until activation."
        )
    )
    return LicenseStatus(
        available=True,
        is_complete=is_complete,
        is_trial=bool(s.get("isTrial")),
        is_expired=bool(s.get("isExpired")),
        is_not_activated=bool(s.get("isNotActivated")),
        days_left=days_left,
        status_text=s.get("statusText"),
        actions_allowed=is_complete,
        note=note,
    )


@mcp.tool(
    description=(
        "ModelChoice: Close an open workbook by file name. By DEFAULT unsaved "
        "changes are DISCARDED (save=False) — pass save=True to write them "
        "first. The counterpart to open_workbook. Raises if Excel isn't running "
        "or the named workbook isn't open."
    )
)
def close_workbook(
    workbook_name: Annotated[
        str, Field(description="File name of an open workbook, e.g. 'oil.xlsx'.")
    ],
    save: Annotated[
        bool,
        Field(description="Save before closing. False (default) discards unsaved changes."),
    ] = False,
) -> CloseWorkbookResult:
    r = get_bridge().close_workbook(workbook_name, save)
    closed = r.get("closed", workbook_name)
    saved = bool(r.get("saved", save))
    note = (
        f"Closed {closed!r}" + (" (saved)." if saved else " — unsaved changes discarded.")
    )
    return CloseWorkbookResult(workbook=closed, saved=saved, note=note)


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
        "ModelChoice: Export a tree's raw ModelChoice model JSON — for saving "
        "to a file, sharing, version control, or re-importing elsewhere with "
        "import_tree_json. Read-only. Pass tree_name to pick a tree, else the "
        "first/active one."
    )
)
def export_tree_json(
    tree_name: Annotated[
        str | None, Field(description="Tree sheet name. Omit for the first tree.")
    ] = None,
    workbook_name: str | None = None,
) -> TreeExport:
    bridge = get_bridge()
    trees = bridge.list_trees(workbook_name)
    if tree_name is None:
        tree_name = next(iter(trees))
    if tree_name not in trees:
        raise ValueError(f"Tree {tree_name!r} not found. Available: {', '.join(trees)}.")
    raw = trees[tree_name]
    t = parse_model(raw)
    return TreeExport(
        tree=tree_name,
        model_name=t.model_name,
        node_count=len(t.nodes),
        model_json=raw,
    )


@mcp.tool(
    description=(
        "ModelChoice: Import a PrecisionTree workbook (.xls/.xlsx) into "
        "ModelChoice. Converts a copy (the original is never modified), opens "
        "it as the active workbook, and returns the converted workbook name and "
        "its tree sheets. Drives ModelChoice's headless MC_ImportPrecisionTree_Auto; "
        "requires the add-in loaded (a build that includes the import command)."
    )
)
def import_precisiontree(
    file_path: Annotated[str, Field(description="Path to the PrecisionTree .xls/.xlsx file.")],
) -> ImportResult:
    r = get_bridge().import_precisiontree(file_path)
    trees = r.get("trees", [])
    wb = r.get("workbook")
    return ImportResult(
        workbook=wb,
        trees=trees,
        note=(
            f"Imported to {wb!r}; trees: {', '.join(trees)}."
            if trees
            else "Import ran; no ModelChoice trees were found in the result — "
            "check the source is a PrecisionTree model."
        ),
    )


@mcp.tool(
    description=(
        "ModelChoice: Import a tree from raw ModelChoice model JSON (e.g. one "
        "produced by export_tree_json). Validates it by parsing + rolling it "
        "back (catching bad JSON, cycles, missing children), then returns the "
        "EV + optimal policy. dry_run=True (default) previews without touching "
        "Excel; dry_run=False writes the tree into the workbook (preserving the "
        "JSON as-is) and renders it (needs the add-in loaded)."
    )
)
def import_tree_json(
    model_json: Annotated[str, Field(description="Raw ModelChoice model JSON to import.")],
    sheet_name: Annotated[
        str | None, Field(description="Tree sheet name to write to. Omit to auto-name.")
    ] = None,
    dry_run: bool = True,
    workbook_name: str | None = None,
) -> BuildTreeResult:
    parsed = parse_model(model_json)  # raises TreeParseError on bad JSON/structure
    r = rollup(parsed)
    direction = "maximize" if parsed.maximize else "minimize"
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
        sheet = bridge.write_tree(model_json, sheet_name=sheet_name, workbook=workbook_name)
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
        "ModelChoice: What-if scenario comparison. Give a list of named "
        "`scenarios`, each a set of input `edits` (value changes — "
        "set_probability / set_branch_value / set_terminal_value, the same ops "
        "edit_tree uses). Each scenario is applied to a fresh copy of the tree "
        "and rolled back; the tool returns each scenario's expected value, its "
        "optimal decisions, the change vs the baseline, whether the optimal "
        "decision flipped, and which scenario wins by the tree's objective. "
        "Pure Python — reads the stored tree, no Excel writes, no add-in needed. "
        "Pairs with build_control_panel: scenarios are sets of panel-input values."
    )
)
def run_scenarios(
    scenarios: Annotated[
        list[ScenarioSpec],
        Field(description="Named scenarios, each a bundle of input overrides."),
    ],
    tree_name: Annotated[
        str | None, Field(description="Tree sheet name. Omit for the first tree.")
    ] = None,
    workbook_name: str | None = None,
) -> ScenarioComparison:
    bridge = get_bridge()
    trees = bridge.list_trees(workbook_name)
    if tree_name is None:
        tree_name = next(iter(trees))
    base = parse_model(trees[tree_name])
    baseline = rollup(base)

    outcomes: list[ScenarioOutcome] = []
    for sc in scenarios:
        t = _apply_edits(base, sc.edits)
        r = rollup(t)
        outcomes.append(
            ScenarioOutcome(
                name=sc.name,
                expected_value=r.expected_value,
                optimal_path=r.optimal_path,
                delta_vs_baseline=r.expected_value - baseline.expected_value,
                decision_changed=(r.optimal_path != baseline.optimal_path),
            )
        )

    best: str | None = None
    if outcomes:
        chooser = max if base.maximize else min
        best = chooser(outcomes, key=lambda o: o.expected_value).name

    flips = [o.name for o in outcomes if o.decision_changed]
    flip_note = (
        f"Decision flips under: {', '.join(flips)}. "
        if flips
        else "Decision holds across all scenarios. "
    )
    best_note = f"Best by {'max' if base.maximize else 'min'} EV: {best}." if best else ""
    note = f"Baseline EV {baseline.expected_value:,.2f}. " + flip_note + best_note

    return ScenarioComparison(
        tree=tree_name,
        objective="maximize" if base.maximize else "minimize",
        baseline_ev=baseline.expected_value,
        baseline_optimal_path=baseline.optimal_path,
        scenarios=outcomes,
        best_scenario=best,
        note=note,
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


def _parse_money(value: Any) -> float | None:
    """Parse a number from a report cell. The MC_EVII renderer writes values
    as culture-formatted text (e.g. '1,234.56'), so accept strings with
    thousands separators as well as native numbers."""
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        text = value.strip().replace(",", "").replace(" ", "")
        if not text:
            return None
        try:
            return float(text)
        except ValueError:
            return None
    return None


@mcp.tool(
    description=(
        "ModelChoice: Compute the Expected Value of Imperfect Information "
        "(EVII / EVSI) for a specific test on the active tree — how much a "
        "real, imperfect information source is worth before you decide. Unlike "
        "EVPI, EVII needs a test spec: `chance_node` (the uncertainty the test "
        "informs), `likelihood_matrix` = P(signal|state) with one ROW per state "
        "of that node (in branch order) and one COLUMN per test signal (each row "
        "sums to 1), an optional `test_cost`, and optional `signals` names. "
        "Returns EVII, its net value after cost (positive ⇒ worth running), the "
        "EVPI ceiling, and the recommendation. Drives ModelChoice's headless "
        "MC_EVII_Auto; requires the add-in loaded with a tree open. Not defined "
        "for MCDA models."
    )
)
def run_evii(
    chance_node: Annotated[
        str, Field(description="Name of the chance node whose states the test informs.")
    ],
    likelihood_matrix: Annotated[
        list[list[float]],
        Field(
            description=(
                "P(signal|state): one row per state (in the node's branch order), "
                "one column per signal; each row sums to 1."
            )
        ),
    ],
    test_cost: Annotated[
        float, Field(description="Cost of running the test. Default 0.")
    ] = 0.0,
    signals: Annotated[
        list[str] | None,
        Field(description="Optional signal names; defaults to Signal1..N."),
    ] = None,
    test_name: Annotated[
        str, Field(description="Optional test label for the report. Default 'Test'.")
    ] = "Test",
    workbook_name: str | None = None,
) -> EviiResult:
    bridge = get_bridge()
    r = bridge.run_evii(
        chance_node, likelihood_matrix, test_cost, signals, test_name, workbook_name
    )
    rows: list[list[Any]] = r.get("rows", [])
    sheets = [s for s in r.get("sheets", []) if s.startswith("MC_EVII")]

    def _cell(row_idx: int, col_idx: int) -> Any:
        if 0 <= row_idx < len(rows):
            row = rows[row_idx]
            if 0 <= col_idx < len(row):
                return row[col_idx]
        return None

    # Key Values block (renderer-fixed): labels in column B, values in C.
    # Rows 8-12 (1-based) = data; column C = index 2.
    prior_ev = _parse_money(_cell(7, 2))
    evpi = _parse_money(_cell(8, 2))
    evii = _parse_money(_cell(9, 2))
    cost = _parse_money(_cell(10, 2))
    net = _parse_money(_cell(11, 2))

    # Recommendation: first non-empty column-B cell below the key-values table.
    recommendation: str | None = None
    for row in rows[12:]:
        if len(row) > 1 and isinstance(row[1], str) and row[1].strip():
            recommendation = row[1].strip()
            break

    worthwhile = (net > 0) if net is not None else None
    if net is None or evii is None:
        interp = "EVII could not be read from the MC_EVII report."
    elif net > 0:
        interp = (
            f"The test is worth running: its information is worth {evii:,.2f}, "
            f"exceeding its {cost or 0:,.2f} cost by {net:,.2f} (net). EVII can "
            f"never beat the EVPI ceiling of {evpi:,.2f}."
            if evpi is not None
            else f"The test is worth running: net value {net:,.2f} after cost."
        )
    else:
        interp = (
            f"The test is not worth running: its information ({evii:,.2f}) does "
            f"not justify its {cost or 0:,.2f} cost (net {net:,.2f})."
        )

    return EviiResult(
        chance_node=chance_node,
        test_name=test_name,
        prior_ev=prior_ev,
        evpi_upper_bound=evpi,
        evii=evii,
        test_cost=cost if cost is not None else test_cost,
        net_value=net,
        worthwhile=worthwhile,
        recommendation=recommendation,
        report_rows=rows,
        sheets=sheets,
        interpretation=interp,
    )


@mcp.tool(
    description=(
        "ModelChoice: Lift the active tree's inputs into a labelled CONTROL "
        "PANEL at the top of its sheet, with each tree cell linked back to its "
        "panel cell — so an analyst can drive the whole model from one tidy "
        "input block and every probability/value flows through automatically. "
        "Covers chance-branch probabilities and branch/option cash-flow values "
        "(terminal payoffs are a follow-up). The links survive re-renders. "
        "Drives ModelChoice's headless MC_BuildControlPanel_Auto; requires the "
        "add-in loaded with a RENDERED tree (run build_tree/render first)."
    )
)
def build_control_panel(
    tree_name: Annotated[
        str | None,
        Field(description="Tree sheet name to build the panel on. Omit for the active sheet."),
    ] = None,
    workbook_name: str | None = None,
) -> ControlPanelResult:
    bridge = get_bridge()
    r = bridge.build_control_panel(tree_name, workbook_name)
    rows: list[list[Any]] = r.get("rows", [])

    # Panel layout: title in B1, "Input"/"Value" header in B2/C2, then one
    # labelled row per input (label in column B = index 1, value in column C
    # = index 2). Collect rows whose label is text and value is numeric.
    inputs: list[KeyValue] = []
    for row in rows:
        if len(row) < 3:
            continue
        label, value = row[1], row[2]
        if not isinstance(label, str) or not label.strip():
            continue
        if label.strip().lower() in ("input", "control panel — inputs"):
            continue
        if isinstance(value, (int, float)) and not isinstance(value, bool):
            inputs.append(KeyValue(label=label.strip(), value=value))

    return ControlPanelResult(
        sheet=r.get("sheet"),
        linked_count=len(inputs),
        inputs=inputs,
        note=(
            f"Built a control panel of {len(inputs)} linked inputs at the top of "
            f"{r.get('sheet')!r}. Edit a panel value and the tree re-rolls; the "
            "links persist across re-renders."
            if inputs
            else "Control-panel command ran but no linked inputs were read back — "
            "check that the tree is rendered."
        ),
    )


@mcp.tool(
    description=(
        "ModelChoice: Assign an UNCERTAINTY to a tree input — put a ModelRisk "
        "`Vose*` distribution on a branch's cash flow or probability, the way the "
        "ModelChoice UI lets you type a distribution into an input cell. Pass the "
        "node id, the branch/option label, the distribution formula (e.g. "
        "'VoseNormal(100,20)', 'VosePERT(80,100,150)'), and kind='value' (cash "
        "flow) or 'probability'. The distribution is stored as the cell's "
        "user-formula and re-rendered, so it survives edits. This is the "
        "decision-tree half of a Monte Carlo: once inputs are distributions, run "
        "the simulation with modelrisk-mcp (it samples these cells and collects "
        "the tree's output distribution). Requires the add-in loaded with a tree "
        "rendered. Branch values / probabilities only (not terminal payoffs)."
    )
)
def set_input_distribution(
    node_id: Annotated[str, Field(description="Node whose branch/option carries the input.")],
    outcome: Annotated[
        str, Field(description="The branch/option label to put the distribution on.")
    ],
    distribution: Annotated[
        str,
        Field(
            description=(
                "The ModelRisk distribution formula, e.g. 'VoseNormal(100,20)' or "
                "'=VosePERT(80,100,150)'. A leading '=' is optional."
            )
        ),
    ],
    kind: Annotated[
        str, Field(description="'value' (branch cash flow) or 'probability'.")
    ] = "value",
    tree_name: Annotated[
        str | None, Field(description="Tree sheet name. Omit for the first tree.")
    ] = None,
    workbook_name: str | None = None,
) -> InputDistributionResult:
    k = kind.lower()
    if k not in ("value", "probability"):
        raise ValueError(f"kind must be 'value' or 'probability', got {kind!r}.")

    bridge = get_bridge()
    trees = bridge.list_trees(workbook_name)
    if tree_name is None:
        tree_name = next(iter(trees))
    if tree_name not in trees:
        raise TreeParseError(f"no tree {tree_name!r}; available: {', '.join(trees)}.")

    tree = parse_model(trees[tree_name])
    node = tree.nodes.get(node_id)
    if node is None:
        raise ValueError(f"node {node_id!r} not found in {tree_name!r}.")
    branch = next((b for b in node.branches if b.name == outcome), None)
    if branch is None:
        names = ", ".join(b.name for b in node.branches) or "(none)"
        raise ValueError(
            f"node {node_id!r} ({node.name!r}) has no branch {outcome!r}. "
            f"Branches: {names}."
        )
    if k == "probability" and node.kind != "chance":
        raise ValueError(
            f"kind='probability' is only valid on a chance node; {node_id!r} is a "
            f"{node.kind} node. Use kind='value' for a decision option's cash flow."
        )

    prefix = "MC_BV_" if k == "value" else "MC_BP_"
    named_range = f"{prefix}{node_id}_{branch.child_id}"
    formula = distribution.strip()
    if not formula.startswith("="):
        formula = "=" + formula

    sheet = bridge.set_input_formula(named_range, formula, tree_name, workbook_name)

    return InputDistributionResult(
        tree=sheet,
        node_id=node_id,
        outcome=outcome,
        kind=k,
        named_range=named_range,
        formula=formula,
        note=(
            f"Put {formula} on {node.name!r} → {outcome!r} ({named_range}). It's now "
            "an uncertain input. Run the Monte Carlo with modelrisk-mcp's "
            "run_simulation (track the root EV cell as the output) to get the "
            "tree's outcome distribution."
        ),
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
        "ModelChoice: Apply a decision-maker's RISK ATTITUDE (utility function) "
        "to the active tree and return the certainty equivalent, the risk "
        "premium (EV minus CE), and the optimal decision under risk aversion — "
        "which can differ from the EV-maximising choice. `function` is "
        "'exponential' (needs `risk_tolerance` R — roughly the 50/50 "
        "double-or-nothing amount you're indifferent to), 'logarithmic', 'sqrt', "
        "or 'linear'. Drives ModelChoice's headless MC_Utility_Auto; requires the "
        "add-in loaded with a tree open. Not defined for MCDA models."
    )
)
def run_utility(
    risk_tolerance: Annotated[
        float,
        Field(description="Risk tolerance R for exponential utility; ignored by others."),
    ] = 1.0,
    function: Annotated[
        str, Field(description="exponential | logarithmic | sqrt | linear")
    ] = "exponential",
    workbook_name: str | None = None,
) -> UtilityResult:
    r = get_bridge().run_utility(function, risk_tolerance, workbook_name)
    ev = r.get("expected_value")
    ce = r.get("certainty_equivalent")
    premium = r.get("risk_premium")
    opt_ev = r.get("optimal_decision_ev")
    opt_u = r.get("optimal_decision_utility")
    changed = (opt_ev != opt_u) if (opt_ev is not None and opt_u is not None) else None

    if ce is None:
        interp = "Utility result could not be read from the MC_Utility sheet."
    elif changed:
        interp = (
            f"Risk attitude CHANGES the decision: by EV you'd pick {opt_ev!r}, but under "
            f"this risk attitude the optimal choice is {opt_u!r} (certainty equivalent "
            f"{ce:,.2f})."
        )
    else:
        prem_txt = (
            f"; risk premium {premium:,.2f} (positive means risk-averse)."
            if premium is not None
            else "."
        )
        interp = (
            f"Certainty equivalent {ce:,.2f}{prem_txt} The optimal decision "
            f"({opt_u!r}) is unchanged from the EV choice."
        )

    return UtilityResult(
        function=r.get("function"),
        risk_tolerance=r.get("risk_tolerance"),
        expected_value=ev,
        certainty_equivalent=ce,
        risk_premium=premium,
        optimal_decision_ev=opt_ev,
        optimal_decision_utility=opt_u,
        decision_changed=changed,
        interpretation=interp,
    )


def _rp_num(v: Any) -> float | None:
    return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else None


def _parse_risk_profile_series(rows: list[list[Any]]) -> list[RiskProfileSeries]:
    """Parse the per-series stats block (Statistic | series… header, then EV /
    Min / Max / StdDev rows). Series names come from the header; stat fields are
    matched by keyword (English) and degrade to None otherwise."""
    header_idx = None
    for i, row in enumerate(rows):
        if (
            len(row) > 2
            and isinstance(row[1], str)
            and any(isinstance(c, str) and c.strip() for c in row[2:])
        ):
            nxt = rows[i + 1] if i + 1 < len(rows) else []
            if len(nxt) > 2 and _rp_num(nxt[2]) is not None:
                header_idx = i
                break
    if header_idx is None:
        return []

    names = [str(c).strip() for c in rows[header_idx][2:] if isinstance(c, str) and c.strip()]
    stats: dict[str, dict[str, float]] = {n: {} for n in names}
    for row in rows[header_idx + 1:]:
        if len(row) < 2 or not isinstance(row[1], str) or not row[1].strip():
            break
        label = row[1].strip().lower()
        if "expected" in label or "mean" in label:
            field = "expected_value"
        elif "min" in label:
            field = "minimum"
        elif "max" in label:
            field = "maximum"
        elif "std" in label or "deviation" in label:
            field = "std_dev"
        else:
            continue
        for j, n in enumerate(names):
            v = _rp_num(row[2 + j]) if 2 + j < len(row) else None
            if v is not None:
                stats[n][field] = v
    return [RiskProfileSeries(name=n, **stats[n]) for n in names]


@mcp.tool(
    description=(
        "ModelChoice: Run the risk profile and return the outcome distribution "
        "for each decision option — not just the expected value, but the spread "
        "(expected value, minimum, maximum, std dev per option) plus the full "
        "cumulative-probability table rows. Shows downside/upside, not just the "
        "average. Drives ModelChoice's headless MC_RiskProfile_Auto; requires "
        "the add-in loaded with a tree open."
    )
)
def run_risk_profile(workbook_name: str | None = None) -> RiskProfileReport:
    bridge = get_bridge()
    run = bridge.run_analysis("MC_RiskProfile_Auto", workbook_name)
    all_sheets = run.get("sheets", [])
    rp_sheets = [s for s in all_sheets if s.startswith("MC_RiskProfile")]
    rows: list[list[Any]] = []
    if "MC_RiskProfile" in all_sheets:
        rows = bridge.read_sheet("MC_RiskProfile", workbook_name)
    series = _parse_risk_profile_series(rows)
    return RiskProfileReport(
        series=series,
        report_rows=rows,
        sheets=rp_sheets,
        note=(
            f"Outcome stats for {len(series)} option(s). The report rows include "
            "the cumulative-probability table (read other sheets via read_sheet)."
            if series
            else "Risk profile ran; parse the report rows for the distribution."
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


# Friendly report name -> (command, primary result sheet, sheet-name prefixes).
_REPORTS: dict[str, tuple[str, str, tuple[str, ...]]] = {
    "strategy_table": ("MC_ExportStrategyTable_Auto", "MC_StrategyTable", ("MC_Strat",)),
    "policy_suggestion": ("MC_PolicySuggestion_Auto", "MC_PolicyTable", ("MC_Policy",)),
    "decision_brief": ("MC_DecisionBrief_Auto", "MC_DecisionBrief", ("MC_Brief", "MC_Decision")),
    "mcda_report": ("MC_McdaReport_Auto", "MC_MCDA_Summary", ("MC_MCDA", "MC_Mcda")),
    "force_to_outcome": ("MC_ForceToOutcome_Auto", "MC_ForceOutcome", ("MC_Force",)),
    "two_way_sensitivity": ("MC_TwoWaySensitivity_Auto", "MC_TwoWaySens", ("MC_TwoWay",)),
}


@mcp.tool(
    description=(
        "ModelChoice: Run a decision report and read it back in one call — "
        "'strategy_table' (the optimal action for every scenario), "
        "'policy_suggestion' (recommended policy + rationale), "
        "'decision_brief' (an executive summary of the decision), "
        "'mcda_report' (multi-criteria scores), 'force_to_outcome' (what "
        "inputs would have to change to force a chosen outcome), or "
        "'two_way_sensitivity' (the EV grid over the two most-sensitive inputs). "
        "Returns the primary sheet's "
        "rows, any label→value pairs found, and the related report sheets. "
        "Requires the ModelChoice add-in loaded with a tree open."
    )
)
def run_decision_report(
    report: Annotated[
        str, Field(description="One of: " + ", ".join(sorted(_REPORTS)))
    ],
    workbook_name: str | None = None,
) -> DecisionReport:
    key = report.lower()
    if key not in _REPORTS:
        raise ValueError(
            f"Unknown report {report!r}. Choose from: {', '.join(sorted(_REPORTS))}."
        )
    command, primary, prefixes = _REPORTS[key]
    bridge = get_bridge()
    run = bridge.run_analysis(command, workbook_name)
    all_sheets = run.get("sheets", [])
    related = [s for s in all_sheets if s.startswith(prefixes)]

    rows: list[Any] = []
    primary_sheet: str | None = None
    if primary in all_sheets:
        primary_sheet = primary
    elif related:
        primary_sheet = related[0]
    if primary_sheet is not None:
        rows = bridge.read_sheet(primary_sheet, workbook_name)

    return DecisionReport(
        report=key,
        command=command,
        primary_sheet=primary_sheet,
        rows=rows,
        details=_key_values(rows),
        sheets=related,
        note=(
            f"{report} report. Other sheets ({', '.join(related)}) are "
            "available via read_sheet."
            if related
            else f"{report} ran; no report sheet was found."
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
    "build_control_panel",
    "build_mcda",
    "build_tree",
    "close_workbook",
    "edit_tree",
    "export_tree_json",
    "get_bridge",
    "get_tree",
    "import_precisiontree",
    "import_tree_json",
    "license_status",
    "list_trees",
    "open_workbook",
    "read_sheet",
    "roll_up",
    "run_analysis",
    "run_decision_report",
    "run_evii",
    "run_evpi",
    "run_risk_profile",
    "run_robustness",
    "run_scenarios",
    "run_sensitivity",
    "run_utility",
    "set_bridge_for_testing",
    "set_input_distribution",
    "verify_rollback",
]
