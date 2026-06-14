# Changelog

All notable changes to `modelchoice-mcp`. Versions are tag-driven; pushing a
`vX.Y.Z` tag publishes to PyPI via the release workflow.

## 0.0.26
- **`open_workbook`** — open a decision-tree workbook (.xlsx) from disk in the
  running Excel so the other tools can act on it. Reports the workbook's sheets
  and any ModelChoice tree sheets it contains; reuses an already-open workbook
  of the same name. (Mirrors the new `open_workbook` in modelrisk-mcp.)

## 0.0.25
- **`build_mcda` gains AHP weight elicitation (AB#2646)** — set
  `weight_source="ahp"` and pass an `ahp_matrix` (Saaty 1-9 pairwise comparisons,
  `(1 + criteria)` square, ordered `[financial, …criteria]`) instead of direct
  weights. The tool computes the eigenvector weights + consistency ratio for the
  preview and forwards `weightSource`/`ahpMatrix` to `MC_ApplyMcda_Auto`, which
  recomputes them authoritatively (CR surfaced; CR > 0.10 flagged inconsistent).
  `McdaBuildResult` now reports `weight_source`, `weights`, `consistency_ratio`.

## 0.0.24
- **`/decision-tree-monte-carlo` prompt** — guides the cross-server Monte Carlo
  hand-off: assign `Vose*` distributions to a tree's uncertain inputs
  (`set_input_distribution`), wrap the root EV (`MC_V_<rootId>`) as a ModelRisk
  output, run the simulation in **modelrisk-mcp**, and read the EV's output
  distribution back. The `/design-decision-tree` prompt's step 5 now points to it.
- **README roadmap** refreshed — Phases 0–3 marked delivered (22 tools, 2
  prompts), with simulation-orchestration and AHP-for-MCDA as candidate next
  steps.

## 0.0.23
- **`set_input_distribution`** — assign an uncertainty (a ModelRisk `Vose*`
  distribution) to a tree input (branch cash flow or probability), the way the
  UI lets you type a distribution into a cell. Stored as the input's
  user-formula and re-rendered, so it persists. This is the decision-tree half
  of a Monte Carlo: once inputs are distributions, run the simulation with
  **modelrisk-mcp** (it samples these cells and collects the tree's output
  distribution). Pure `_MC_Store` edit + re-render — no new add-in command.

## 0.0.22
- **`build_mcda`** — build a multi-criteria (MCDA) model: tree + criteria
  (ordinal options, weights, direction) + aggregation + per-terminal scores.
  Validates structure/weights/scores in Python; drives the new
  `MC_ApplyMcda_Auto` (AB#2637) to set MCDA mode + render. v1 = direct weights
  (AHP follow-up). No-op until that add-in build ships.

## 0.0.21
- **`run_utility`** — risk-attitude (utility) rollback: certainty equivalent,
  risk premium, and the optimal decision under risk aversion vs EV. Drives the
  new `MC_Utility_Auto` (AB#2635); no-op until that add-in build ships.

## 0.0.20
- **Drivers for the C# add-in batch** — `import_precisiontree` (drives
  `MC_ImportPrecisionTree_Auto`) and `two_way_sensitivity` via
  `run_decision_report` (drives `MC_TwoWaySensitivity_Auto`). Both no-op until
  the matching add-in build (PR !2032) is installed. (Terminal-payoff linking
  in the control panel is add-in-side; no new MCP tool needed.)

## 0.0.19
- **Housekeeping** — CHANGELOG; gated live integration tests
  (`MODELCHOICE_LIVE=1`) exercising the real COM attach/read/rollback path.

## 0.0.18
- **Knowledge resources** — `modelchoice://guide/*` curated decision-analysis
  guidance (building trees, EVPI vs EVII, which analysis to use, pitfalls).

## 0.0.17
- **Import / export** — `export_tree_json` (a tree's raw ModelChoice JSON, to
  save/share/version) and `import_tree_json` (validate raw JSON → write/render).

## 0.0.16
- `run_decision_report` gains **`force_to_outcome`** (what inputs would have to
  change to force a chosen outcome).

## 0.0.15
- **`run_risk_profile`** — per-option outcome distribution (expected value, min,
  max, std dev + the cumulative-probability table), not just the EV.

## 0.0.14
- **`run_scenarios`** — what-if comparison: named bundles of input edits, each
  rolled back and compared to the baseline (EV, optimal decision, Δ, flips).

## 0.0.13
- **Excel attach fix** — bind to the running Excel via `GetActiveObject` (ROT)
  instead of xlwings' window-handle walk, which failed (`0x800A01A8`) once the
  ModelChoice add-in was loaded. Fixes timeouts/hangs from the MCP server.

## 0.0.12
- **`build_control_panel`** — lift a tree's inputs into a labelled control panel
  at the top of its sheet, with the tree linked back to it.

## 0.0.11
- `edit_tree` gains structural ops: `add_option`, `add_branch`, `remove_branch`.

## 0.0.10
- **`run_evii`** — Expected Value of Imperfect Information for a specific test.

## 0.0.9
- **`run_decision_report`** — strategy table / policy / decision brief / MCDA.

## 0.0.1 – 0.0.8
- Phases 0–2: read + rollback engine (`list_trees`, `get_tree`, `roll_up`,
  `verify_rollback`), build/edit (`build_tree`, `edit_tree`), and analysis
  drivers (`run_evpi`, `run_robustness`, `run_sensitivity`, `run_analysis`,
  `read_sheet`), plus the `/design-decision-tree` prompt. CI + tag-driven PyPI
  release. First PyPI publish at 0.0.1.
