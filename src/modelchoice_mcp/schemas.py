"""Pydantic response schemas for the MCP tools."""

from __future__ import annotations

from pydantic import BaseModel, Field


class BranchView(BaseModel):
    name: str
    child_id: str
    value: float = Field(description="Branch/option value added along the path.")
    probability: float | None = Field(
        default=None, description="Probability (chance branches only; null for decision options)."
    )


class NodeView(BaseModel):
    id: str
    name: str
    kind: str = Field(description="'decision', 'chance', or 'terminal'.")
    value: float = Field(default=0.0, description="Terminal node's own value (0 otherwise).")
    branches: list[BranchView] = Field(default_factory=list)


class TreeSummary(BaseModel):
    name: str = Field(description="Tree sheet name (e.g. 'MC_Tree_1').")
    model_name: str
    root_id: str
    root_name: str
    node_count: int
    decision_count: int
    chance_count: int
    terminal_count: int


class TreeList(BaseModel):
    workbook: str | None = None
    count: int
    trees: list[TreeSummary]


class TreeStructure(BaseModel):
    name: str
    model_name: str
    root_id: str
    maximize: bool = Field(description="True = the tree maximizes EV; False = minimizes.")
    node_count: int
    nodes: list[NodeView]


class NodeResultView(BaseModel):
    id: str
    name: str
    kind: str
    expected_value: float
    optimal_branch_name: str | None = Field(
        default=None, description="At a decision node, the option chosen by rollback."
    )


class RollupResponse(BaseModel):
    name: str
    model_name: str
    maximize: bool
    expected_value: float = Field(description="Rolled-back expected value at the root.")
    optimal_path: list[str] = Field(
        description="The decisions taken under the optimal policy, from the root."
    )
    recommendation: str = Field(description="Plain-English decision recommendation.")
    nodes: list[NodeResultView]
