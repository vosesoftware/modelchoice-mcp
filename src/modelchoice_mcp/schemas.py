"""Pydantic response schemas for the MCP tools."""

from __future__ import annotations

from typing import Any

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


class NodeDiff(BaseModel):
    node_id: str
    name: str
    computed: float = Field(description="Our Python rollback EV for this node.")
    cell: float = Field(description="The MC_V_<nodeId> value ModelChoice wrote.")
    diff: float = Field(description="computed - cell.")


class AnalysisRun(BaseModel):
    """Result of driving a ModelChoice headless analysis command."""

    analysis: str = Field(description="Friendly analysis name requested.")
    command: str = Field(description="The MC_*_Auto ExcelCommand that was run.")
    new_sheets: list[str] = Field(
        description="Sheets the analysis created (read them with read_sheet)."
    )
    sheets: list[str] = Field(description="All sheets in the workbook afterwards.")
    note: str


class SheetData(BaseModel):
    """A capped block of a worksheet's used range."""

    sheet: str
    row_count: int
    rows: list[list[Any]] = Field(description="Cell values, row-major (numbers, text, or null).")
    truncated: bool = Field(description="True if the sheet had more rows/cols than returned.")


class KeyValue(BaseModel):
    label: str
    value: Any


class RobustnessSummary(BaseModel):
    """Structured read of ModelChoice's robustness ('break the decision')
    analysis — how much inputs must change to flip the optimal choice."""

    verdict: str | None = Field(
        default=None, description="Overall robustness verdict (e.g. 'Robust', 'Fragile')."
    )
    robustness_score: str | None = Field(
        default=None, description="Score out of 100, if reported."
    )
    min_distance: float | None = Field(
        default=None,
        description="Smallest normalized change to an input that flips the decision.",
    )
    details: list[KeyValue] = Field(
        description="All label→value pairs read from the verdict sheet."
    )
    sheets: list[str] = Field(description="Robustness report sheets produced (MC_RB_*).")
    note: str


class EvpiResult(BaseModel):
    """Expected Value of Perfect Information for the active tree, produced
    by ModelChoice's headless MC_EVPI_Auto command."""

    model_name: str | None = None
    objective: str | None = Field(default=None, description="'Maximize' or 'Minimize'.")
    optimal_ev: float | None = Field(default=None, description="Rolled-back expected value.")
    evpi: float | None = Field(
        default=None,
        description="Full-tree EVPI — the most you'd pay for perfect information.",
    )
    value_with_perfect_info: float | None = Field(
        default=None, description="Optimal EV plus (or minus) the EVPI."
    )
    interpretation: str = Field(
        description="Plain-English reading of the EVPI."
    )


class RollbackVerification(BaseModel):
    """Cross-check of our Python rollback against the EVs ModelChoice
    itself wrote into the MC_V_<nodeId> named ranges."""

    name: str
    rendered: bool = Field(
        description="True if the tree has MC_V_ cells (i.e. it has been rendered by the add-in)."
    )
    compared_count: int = Field(description="Nodes present in both our rollback and the cells.")
    matches: int
    max_abs_diff: float = Field(description="Largest |computed - cell| across compared nodes.")
    mismatches: list[NodeDiff] = Field(
        default_factory=list, description="Nodes whose values differ beyond the tolerance."
    )
    verdict: str
