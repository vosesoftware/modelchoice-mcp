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
    BranchView,
    NodeResultView,
    NodeView,
    RollupResponse,
    TreeList,
    TreeStructure,
    TreeSummary,
)
from modelchoice_mcp.server import mcp
from modelchoice_mcp.tree import DecisionTree, parse_model, rollup

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


__all__ = ["get_bridge", "get_tree", "list_trees", "roll_up", "set_bridge_for_testing"]
