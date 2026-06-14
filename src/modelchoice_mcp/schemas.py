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


class BranchSpec(BaseModel):
    """A branch (chance) or option (decision) in a tree being built."""

    name: str = Field(description="Branch/option label.")
    child_id: str = Field(description="Id of the node this branch leads to.")
    value: float = Field(
        default=0.0, description="Cash flow added along this branch (cost/payoff)."
    )
    probability: float | None = Field(
        default=None,
        description="Probability (chance branches only; auto-normalized if they don't sum to 1).",
    )


class NodeSpec(BaseModel):
    """One node of a tree being built."""

    id: str = Field(description="Unique node id (e.g. 'D', 'C', 'T1').")
    type: str = Field(description="'decision', 'chance', or 'terminal'.")
    name: str = Field(description="Node label.")
    value: float = Field(default=0.0, description="Terminal node's own payoff (terminal only).")
    branches: list[BranchSpec] = Field(
        default_factory=list,
        description="Children — chance branches (with probability) or decision options.",
    )


class TreeSpec(BaseModel):
    """A decision tree to build, assembled from a description."""

    root_id: str = Field(description="Id of the root node.")
    model_name: str = Field(default="Untitled", description="Tree name.")
    maximize: bool = Field(default=True, description="True = maximize EV; False = minimize.")
    nodes: list[NodeSpec] = Field(description="All nodes in the tree.")


class EditOp(BaseModel):
    """A single targeted edit to an existing tree. `op` selects the
    change; the other fields are the target/new value for that op."""

    op: str = Field(
        description=(
            "Value/label edits: 'set_probability', 'set_branch_value', "
            "'set_terminal_value', 'rename_node', 'rename_branch', 'set_objective'. "
            "Structural edits: 'add_option' (new option on a decision node), "
            "'add_branch' (new outcome on a chance node), 'remove_branch' "
            "(drop an option/outcome by name)."
        )
    )
    node_id: str | None = Field(default=None, description="Target node id (most ops).")
    branch_name: str | None = Field(
        default=None,
        description="Target branch/option name (branch-level ops, incl. remove_branch).",
    )
    value: float | None = Field(
        default=None,
        description=(
            "New number — probability, branch cash flow, or terminal payoff; "
            "for add_option/add_branch it is the new branch's cash flow (default 0)."
        ),
    )
    name: str | None = Field(
        default=None,
        description="New label — for a rename op, or the new branch/option name on add_*.",
    )
    maximize: bool | None = Field(
        default=None, description="New objective for 'set_objective' (true=maximize)."
    )
    child_id: str | None = Field(
        default=None,
        description=(
            "For add_option/add_branch: id of the node the new branch leads to. "
            "Omit to auto-create a new terminal outcome."
        ),
    )
    probability: float | None = Field(
        default=None,
        description="For add_branch on a chance node: the new outcome's probability.",
    )


class BuildTreeResult(BaseModel):
    """Outcome of building a tree from a spec."""

    written: bool = Field(description="True if written into the workbook (else a dry-run preview).")
    sheet: str | None = Field(
        default=None, description="Tree sheet it was written to, if committed."
    )
    node_count: int
    expected_value: float = Field(description="Rolled-back EV of the built tree.")
    optimal_path: list[str]
    recommendation: str
    model_json: str = Field(
        description="The ModelChoice model JSON that was (or would be) written."
    )


class ScenarioSpec(BaseModel):
    """A named what-if scenario: a set of input edits applied together. Edits
    are typically value changes (set_probability / set_branch_value /
    set_terminal_value), the same ops edit_tree uses."""

    name: str = Field(description="Scenario label, e.g. 'Pessimistic' or 'Strong market'.")
    edits: list[EditOp] = Field(description="Input overrides applied for this scenario.")


class ScenarioOutcome(BaseModel):
    name: str
    expected_value: float = Field(description="Rolled-back EV under this scenario.")
    optimal_path: list[str] = Field(description="Optimal decisions under this scenario.")
    delta_vs_baseline: float = Field(description="Scenario EV minus the baseline EV.")
    decision_changed: bool = Field(
        description="True if the optimal decision differs from the baseline's."
    )


class ScenarioComparison(BaseModel):
    """Side-by-side what-if comparison of the tree under several scenarios,
    all rolled back in pure Python (no Excel needed)."""

    tree: str
    objective: str = Field(description="'maximize' or 'minimize'.")
    baseline_ev: float
    baseline_optimal_path: list[str]
    scenarios: list[ScenarioOutcome]
    best_scenario: str | None = Field(
        default=None, description="Scenario with the best EV by the tree's objective."
    )
    note: str


class KeyValue(BaseModel):
    label: str
    value: Any


class CriterionSpec(BaseModel):
    """One non-financial MCDA criterion with discrete ordinal options."""

    id: str = Field(description="Stable criterion id (e.g. 'safety').")
    name: str = Field(description="Display name.")
    weight: float = Field(description="Relative weight 0..1 (financial + all criteria sum to 1).")
    maximize: bool = Field(
        default=True, description="True = last option is best; False = first option is best."
    )
    options: list[str] = Field(description="Ordered discrete options (>=2), worst to best.")


class McdaSpec(BaseModel):
    """The multi-criteria layer for build_mcda: criteria, weights, aggregation,
    and each terminal's per-criterion option selections."""

    criteria: list[CriterionSpec] = Field(description="Non-financial criteria.")
    financial_weight: float = Field(
        default=0.5, description="Weight of the auto-included financial (payoff) criterion."
    )
    aggregation: str = Field(
        default="weighted_sum",
        description="weighted_sum | weighted_geometric | waspas.",
    )
    waspas_lambda: float = Field(default=0.5, description="WASPAS blend (only for waspas).")
    terminal_scores: dict[str, dict[str, str]] = Field(
        default_factory=dict,
        description="Per terminal: {terminal_id: {criterion_id: option_label}}.",
    )
    weight_source: str = Field(
        default="direct",
        description=(
            "'direct' uses the per-criterion `weight` + `financial_weight` fields; "
            "'ahp' derives all weights from `ahp_matrix` (those fields are ignored)."
        ),
    )
    ahp_matrix: list[list[float]] | None = Field(
        default=None,
        description=(
            "AHP pairwise comparison matrix on Saaty's 1-9 scale — required when "
            "weight_source='ahp'. Must be (1 + number_of_criteria) square, ordered "
            "[financial, criterion0, criterion1, ...], positive, and reciprocal "
            "across the diagonal (a[j][i] = 1/a[i][j], diagonal = 1)."
        ),
    )


class McdaBuildResult(BaseModel):
    """Outcome of building an MCDA model."""

    written: bool = Field(description="True if written + rendered (else a validated preview).")
    sheet: str | None = None
    node_count: int
    criteria: list[str] = Field(description="Criterion names defined.")
    terminals_scored: int
    weight_source: str = Field(
        default="direct", description="'direct' or 'ahp' — how the weights were set."
    )
    weights: list[KeyValue] = Field(
        default_factory=list,
        description="Effective weights (label -> weight): 'Financial' + each criterion.",
    )
    consistency_ratio: float | None = Field(
        default=None,
        description="AHP consistency ratio (CR < 0.10 acceptable); None for direct weights.",
    )
    note: str


class InputDistributionResult(BaseModel):
    """Outcome of assigning an uncertainty (a Vose* distribution) to a tree
    input — a branch value or probability — the way the ModelChoice UI lets
    you type a distribution into a cell. The distribution is stored as the
    cell's user-formula and re-rendered, ready for ModelRisk (via
    modelrisk-mcp) to sample during a Monte Carlo simulation."""

    tree: str = Field(description="Tree sheet the input belongs to.")
    node_id: str = Field(description="Node whose branch/option the distribution was put on.")
    outcome: str = Field(description="Branch/option label the distribution was assigned to.")
    kind: str = Field(description="'value' (branch cash flow) or 'probability'.")
    named_range: str = Field(
        description="The MC_BV_/MC_BP_ named range now holding the distribution formula."
    )
    formula: str = Field(description="The cell formula assigned (e.g. '=VoseNormal(100,20)').")
    note: str


class OpenWorkbookResult(BaseModel):
    """Outcome of opening a workbook from disk in the running Excel."""

    workbook: str = Field(description="The opened workbook's file name (now active).")
    sheets: list[str] = Field(description="Worksheet names in the workbook.")
    trees: list[str] = Field(
        description="ModelChoice tree sheets found in the workbook (empty if none)."
    )
    note: str


class TreeExport(BaseModel):
    """A tree's stored ModelChoice model JSON, for saving or sharing."""

    tree: str = Field(description="Tree sheet name.")
    model_name: str
    node_count: int
    model_json: str = Field(
        description="The raw ModelChoice model JSON (round-trips via import_tree_json)."
    )


class ImportResult(BaseModel):
    """Outcome of importing a PrecisionTree workbook into ModelChoice."""

    workbook: str | None = Field(
        default=None, description="The converted ModelChoice workbook, now active."
    )
    trees: list[str] = Field(description="Tree sheet names found in the converted workbook.")
    note: str


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


class DecisionReport(BaseModel):
    """Structured read of one of ModelChoice's report analyses
    (strategy table, policy suggestion, decision brief, MCDA report)."""

    report: str = Field(description="Report requested.")
    command: str = Field(description="The MC_*_Auto command run.")
    primary_sheet: str | None = Field(
        default=None, description="Main result sheet (its rows are returned)."
    )
    rows: list[Any] = Field(description="Rows of the primary sheet (cell values).")
    details: list[KeyValue] = Field(
        default_factory=list, description="Label→value pairs found on the primary sheet."
    )
    sheets: list[str] = Field(description="All related report sheets produced.")
    note: str


class ControlPanelResult(BaseModel):
    """Outcome of lifting a tree's inputs into a linked control panel at the
    top of its sheet (ModelChoice's headless MC_BuildControlPanel_Auto)."""

    sheet: str | None = Field(default=None, description="Tree sheet the panel was built on.")
    linked_count: int = Field(description="Number of inputs lifted into the panel and linked.")
    inputs: list[KeyValue] = Field(
        description="The control-panel rows: input label -> current value."
    )
    note: str


class UtilityResult(BaseModel):
    """Risk-attitude (utility) rollback — the certainty equivalent and the
    optimal decision under a decision-maker's risk attitude, vs plain EV.
    Produced by ModelChoice's headless MC_Utility_Auto."""

    function: str | None = Field(default=None, description="Utility function applied.")
    risk_tolerance: float | None = Field(
        default=None, description="Risk tolerance R (exponential utility)."
    )
    expected_value: float | None = Field(default=None, description="Plain rolled-back EV.")
    certainty_equivalent: float | None = Field(
        default=None,
        description="Guaranteed amount equivalent to the gamble under this risk attitude.",
    )
    risk_premium: float | None = Field(
        default=None, description="EV minus certainty equivalent; positive => risk-averse."
    )
    optimal_decision_ev: str | None = Field(
        default=None, description="Optimal decision by expected value."
    )
    optimal_decision_utility: str | None = Field(
        default=None, description="Optimal decision under the utility function."
    )
    decision_changed: bool | None = Field(
        default=None, description="True if risk attitude changes the optimal decision."
    )
    interpretation: str = Field(description="Plain-English reading of the result.")


class EviiResult(BaseModel):
    """Expected Value of Imperfect Information (EVII / EVSI) for a specific
    test, produced by ModelChoice's headless MC_EVII_Auto command. Unlike
    EVPI, EVII is always relative to a test — a target chance node and a
    likelihood matrix P(signal|state)."""

    chance_node: str = Field(description="The chance node the test informs.")
    test_name: str
    prior_ev: float | None = Field(
        default=None, description="EV under the original (prior) probabilities."
    )
    evpi_upper_bound: float | None = Field(
        default=None, description="EVPI — the perfect-information ceiling EVII cannot exceed."
    )
    evii: float | None = Field(
        default=None, description="Value of the imperfect test's information (before its cost)."
    )
    test_cost: float | None = None
    net_value: float | None = Field(
        default=None, description="EVII minus test cost. Positive => the test is worth running."
    )
    worthwhile: bool | None = Field(
        default=None, description="True when net_value > 0."
    )
    recommendation: str | None = Field(
        default=None, description="ModelChoice's plain-English recommendation."
    )
    report_rows: list[list[Any]] = Field(
        description="Rows of the MC_EVII sheet (key values, per-signal detail, likelihood matrix)."
    )
    sheets: list[str] = Field(description="Report sheets produced (MC_EVII).")
    interpretation: str = Field(description="Plain-English reading of the EVII result.")


class RiskProfileSeries(BaseModel):
    """Outcome statistics for one decision series (option) in a risk profile."""

    name: str = Field(description="Decision option / series name.")
    expected_value: float | None = None
    minimum: float | None = None
    maximum: float | None = None
    std_dev: float | None = None


class RiskProfileReport(BaseModel):
    """ModelChoice's risk profile — the distribution of possible outcomes for
    each decision option (not just the EV). Drives MC_RiskProfile_Auto."""

    series: list[RiskProfileSeries] = Field(
        description="Per-option outcome stats (expected value, min, max, std dev)."
    )
    report_rows: list[list[Any]] = Field(
        description="Rows of the MC_RiskProfile sheet (stats + the cumulative-probability table)."
    )
    sheets: list[str] = Field(description="Risk-profile report sheets produced.")
    note: str


class SensitivityReport(BaseModel):
    """Read of ModelChoice's one-way sensitivity analysis. Variables are
    reported tornado-ordered (largest EV swing first)."""

    baseline_ev: float | None = Field(
        default=None, description="The model's baseline expected value, if found."
    )
    report_rows: list[list[Any]] = Field(
        description="Rows of the MC_SensReport sheet (variables ranked by swing)."
    )
    sheets: list[str] = Field(
        description="Sensitivity report sheets produced (MC_Sens*/MC_Tornado)."
    )
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
