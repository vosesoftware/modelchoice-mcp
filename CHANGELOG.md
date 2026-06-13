# Changelog

All notable changes to `modelchoice-mcp`. Versions are tag-driven; pushing a
`vX.Y.Z` tag publishes to PyPI via the release workflow.

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
