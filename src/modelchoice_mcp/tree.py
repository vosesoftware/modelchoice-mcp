"""Parse and roll back a ModelChoice decision tree from its model JSON.

The JSON schema is defined by ``DecisionTreeModel`` / the polymorphic
``NodeDefinitionJsonConverter`` in ModelChoice.Core:

- Top level: ``{"RootId": str, "Nodes": {id: nodeObj}, "Settings": {...}}``
- Each node has a ``"type"`` discriminator: ``"terminal" | "chance" | "decision"``.
  - terminal: ``{"value": float}``
  - chance:   ``{"branches": [{"name", "probability", "childId", "value"?}]}``
  - decision: ``{"options":  [{"name", "childId", "value"?}]}``
- ``Settings``: ``Maximize`` (true→max EV), ``ChanceProbabilities`` (AutoNormalize
  by default), ``CalculationMethod`` (Cumulative by default).

Rollback mirrors ``DecisionTreeModel.Compile`` + ``TreeNode.Rollback``:
terminal payoff = accumulated branch values along the path (+ the
terminal's own value in classic/Cumulative mode); chance EV = the
(optionally normalized) probability-weighted child EVs; decision EV =
the max (or min) child EV, and the winning option is the optimal policy.
This reproduces ModelChoice's `MC_V_<id>` rollback values in pure Python
— no add-in required.
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from typing import Any


class TreeParseError(ValueError):
    """The model JSON could not be parsed into a valid tree."""


@dataclass(frozen=True)
class Branch:
    name: str
    child_id: str
    value: float = 0.0
    probability: float | None = None  # None for decision options


@dataclass(frozen=True)
class Node:
    id: str
    name: str
    kind: str  # "terminal" | "chance" | "decision"
    value: float = 0.0  # terminal's own value
    branches: list[Branch] = field(default_factory=list)


@dataclass(frozen=True)
class DecisionTree:
    root_id: str
    nodes: dict[str, Node]
    maximize: bool = True
    auto_normalize: bool = True
    cumulative: bool = True
    model_name: str = "Untitled"


@dataclass(frozen=True)
class NodeResult:
    id: str
    name: str
    kind: str
    expected_value: float
    optimal_child_id: str | None = None  # for decision nodes: the chosen option's child
    optimal_branch_name: str | None = None


@dataclass(frozen=True)
class RollupResult:
    root_id: str
    expected_value: float
    optimal_path: list[str]  # branch/option names from root along the optimal policy
    node_results: dict[str, NodeResult]


def _as_float(v: Any, default: float = 0.0) -> float:
    try:
        return float(v)
    except (TypeError, ValueError):
        return default


def parse_model(model_json: str) -> DecisionTree:
    """Parse a serialized ``DecisionTreeModel`` JSON string into a
    :class:`DecisionTree`. Property lookups are case-insensitive on the
    top-level keys (C# serializes PascalCase; we tolerate either)."""
    try:
        raw = json.loads(model_json)
    except json.JSONDecodeError as exc:
        raise TreeParseError(f"model JSON is not valid JSON: {exc}") from exc
    if not isinstance(raw, dict):
        raise TreeParseError("model JSON root must be an object")

    def get(d: dict[str, Any], *names: str, default: Any = None) -> Any:
        lower = {k.lower(): v for k, v in d.items()}
        for n in names:
            if n.lower() in lower:
                return lower[n.lower()]
        return default

    root_id = get(raw, "RootId")
    nodes_raw = get(raw, "Nodes")
    if not isinstance(root_id, str) or not isinstance(nodes_raw, dict):
        raise TreeParseError("model JSON must have a string RootId and a Nodes object")

    settings = get(raw, "Settings", default={}) or {}
    maximize = bool(get(settings, "Maximize", default=True))
    chance_mode = str(get(settings, "ChanceProbabilities", default="AutoNormalize"))
    # enum may serialize as the name or its int (0=MustTotal100, 1=AutoNormalize)
    auto_normalize = chance_mode.lower() in ("autonormalize", "1")
    calc = str(get(settings, "CalculationMethod", default="Cumulative"))
    cumulative = calc.lower() != "mcda"
    model_name = str(get(settings, "ModelName", default="Untitled"))

    nodes: dict[str, Node] = {}
    for nid, n in nodes_raw.items():
        if not isinstance(n, dict):
            raise TreeParseError(f"node {nid!r} is not an object")
        kind = str(n.get("type", "")).lower()
        name = str(n.get("name", nid))
        if kind == "terminal":
            nodes[nid] = Node(id=nid, name=name, kind="terminal", value=_as_float(n.get("value")))
        elif kind == "chance":
            branches = [
                Branch(
                    name=str(b.get("name", "")),
                    child_id=str(b.get("childId", "")),
                    value=_as_float(b.get("value")),
                    probability=_as_float(b.get("probability")),
                )
                for b in n.get("branches", [])
            ]
            nodes[nid] = Node(id=nid, name=name, kind="chance", branches=branches)
        elif kind == "decision":
            options = [
                Branch(
                    name=str(o.get("name", "")),
                    child_id=str(o.get("childId", "")),
                    value=_as_float(o.get("value")),
                    probability=None,
                )
                for o in n.get("options", [])
            ]
            nodes[nid] = Node(id=nid, name=name, kind="decision", branches=options)
        else:
            raise TreeParseError(f"node {nid!r} has unsupported type {kind!r}")

    if root_id not in nodes:
        raise TreeParseError(f"RootId {root_id!r} is not present in Nodes")
    return DecisionTree(
        root_id=root_id,
        nodes=nodes,
        maximize=maximize,
        auto_normalize=auto_normalize,
        cumulative=cumulative,
        model_name=model_name,
    )


def _branch_json(b: Branch, *, include_probability: bool) -> dict[str, Any]:
    out: dict[str, Any] = {"name": b.name, "childId": b.child_id}
    if include_probability:
        out["probability"] = b.probability if b.probability is not None else 0.0
    # Match ModelChoice's writer: omit branch/option value when it is 0.
    if b.value != 0:
        out["value"] = b.value
    return out


def to_model_json(tree: DecisionTree) -> str:
    """Serialize a :class:`DecisionTree` to the exact ModelChoice model
    JSON (the inverse of :func:`parse_model`). Produces the
    ``{"RootId", "Nodes": {id: {"type", ...}}, "Settings"}`` shape that
    ``DecisionTreeModel.FromJson`` consumes — terminal nodes carry
    ``value``; chance nodes carry ``branches`` with ``probability``;
    decision nodes carry ``options``. A partial ``Settings`` is fine —
    ModelChoice fills the rest with defaults on load."""
    nodes: dict[str, Any] = {}
    for nid, n in tree.nodes.items():
        if n.kind == "terminal":
            nodes[nid] = {"type": "terminal", "id": n.id, "name": n.name, "value": n.value}
        elif n.kind == "chance":
            nodes[nid] = {
                "type": "chance", "id": n.id, "name": n.name,
                "branches": [_branch_json(b, include_probability=True) for b in n.branches],
            }
        elif n.kind == "decision":
            nodes[nid] = {
                "type": "decision", "id": n.id, "name": n.name,
                "options": [_branch_json(b, include_probability=False) for b in n.branches],
            }
        else:  # pragma: no cover - guarded by construction
            raise TreeParseError(f"node {nid!r} has unsupported kind {n.kind!r}")

    return json.dumps(
        {
            "RootId": tree.root_id,
            "Nodes": nodes,
            "Settings": {"ModelName": tree.model_name, "Maximize": tree.maximize},
        }
    )


def rollup(tree: DecisionTree) -> RollupResult:
    """Roll the tree back to expected values and the optimal policy.

    Returns the root EV, the optimal path (the sequence of branch/option
    names taken under the optimal policy, following the chosen option at
    each decision and *all* branches at chance nodes is not a single
    path — so the path lists decisions and, at the first chance node,
    stops descending decisions only), and a per-node EV map.
    """
    node_results: dict[str, NodeResult] = {}

    def ev(node_id: str, accumulated: float) -> float:
        node = tree.nodes.get(node_id)
        if node is None:
            raise TreeParseError(f"missing child node {node_id!r}")

        if node.kind == "terminal":
            payoff = accumulated + (node.value if tree.cumulative else 0.0)
            node_results[node_id] = NodeResult(node_id, node.name, "terminal", payoff)
            return payoff

        if node.kind == "chance":
            total_p = sum((b.probability or 0.0) for b in node.branches)
            normalize = tree.auto_normalize or abs(total_p - 1.0) > 1e-12
            value = 0.0
            for b in node.branches:
                p = (b.probability or 0.0)
                p = (p / total_p) if (normalize and total_p > 0) else p
                value += p * ev(b.child_id, accumulated + b.value)
            node_results[node_id] = NodeResult(node_id, node.name, "chance", value)
            return value

        # decision: pick the optimal option
        best_ev: float | None = None
        best: Branch | None = None
        for o in node.branches:
            child_ev = ev(o.child_id, accumulated + o.value)
            if (
                best_ev is None
                or (tree.maximize and child_ev > best_ev)
                or (not tree.maximize and child_ev < best_ev)
            ):
                best_ev, best = child_ev, o
        if best is None or best_ev is None:
            raise TreeParseError(f"decision node {node_id!r} has no options")
        node_results[node_id] = NodeResult(
            node_id, node.name, "decision", best_ev,
            optimal_child_id=best.child_id, optimal_branch_name=best.name,
        )
        return best_ev

    root_ev = ev(tree.root_id, 0.0)

    # Build the optimal path: follow chosen options at decisions until a
    # chance/terminal node ends the deterministic prefix.
    path: list[str] = []
    cur = tree.root_id
    while True:
        res = node_results.get(cur)
        if res is None or res.kind != "decision" or res.optimal_child_id is None:
            break
        path.append(res.optimal_branch_name or "")
        cur = res.optimal_child_id

    return RollupResult(tree.root_id, root_ev, path, node_results)
