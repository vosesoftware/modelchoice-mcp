# ModelChoice MCP

**An open Model Context Protocol server for [Vose Software's ModelChoice](https://www.vosesoftware.com) — the decision-tree add-in for Excel.**

A sibling to [`modelrisk-mcp`](https://github.com/vosesoftware/modelrisk-mcp): where that server brings Monte Carlo risk modelling into a conversation, this one brings **decision analysis** — building, reading, rolling back, and analysing decision trees in Excel.

> **Status: `0.0.27` — Phase 3 (build + drive).** 24 tools: build_tree / build_mcda / edit_tree (incl. add/remove options & outcomes) / set_input_distribution (put a Vose distribution on an input) / build_control_panel / export_tree_json / import_tree_json / import_precisiontree, read + roll + verify, run_scenarios (what-if comparison), plus run_evpi / run_evii / run_risk_profile / run_robustness / run_sensitivity / run_decision_report (strategy/policy/brief/mcda/force-to-outcome/two-way) / run_analysis / read_sheet over ModelChoice's headless commands.

## Tools

| Tool | What it does |
|---|---|
| `build_tree` | **Build a tree from a structured description** and write it into Excel (dry-run by default). Validates + rolls it back so you confirm the recommendation before writing. The build-from-a-prompt path. |
| `build_mcda` | **Build a multi-criteria (MCDA) model** — for choices that aren't pure money. Give the tree + criteria (ordinal options, weights, direction), aggregation, and each terminal's option per criterion; the add-in computes the composite scores. Weights can be entered directly or derived from an **AHP pairwise-comparison matrix** (`weight_source="ahp"` → eigenvector weights + consistency ratio). Drives `MC_ApplyMcda_Auto`. |
| `edit_tree` | **Tweak an existing tree** — change probabilities, payoffs, labels, or the objective, **add or remove whole options/outcomes** (`add_option` / `add_branch` / `remove_branch`), then re-roll it (dry-run by default). The "tweak it by talking" path. |
| `set_input_distribution` | **Make a tree input uncertain** — put a ModelRisk `Vose*` distribution on a branch's cash flow or probability (like typing a distribution into the cell in the UI). The decision-tree half of a Monte Carlo: once inputs are distributions, run the simulation with **modelrisk-mcp** to get the tree's outcome distribution. Pure store-edit + re-render — no add-in command. |
| `build_control_panel` | **Lift the tree's inputs into a control panel** at the top of the sheet — every probability and cash flow as a labelled cell, with the tree linked back to it, so you drive the whole model from one input block. Drives `MC_BuildControlPanel_Auto`. |
| `export_tree_json` / `import_tree_json` | Round-trip a tree's raw ModelChoice JSON — export to save/share/version, import (validated) to write it back into a workbook. |
| `import_precisiontree` | Import a PrecisionTree workbook (.xls/.xlsx) into ModelChoice — converts a copy (original untouched). Drives `MC_ImportPrecisionTree_Auto`. |
| `open_workbook` | **Open a workbook (.xlsx) from disk** in the running Excel so the other tools can act on it. Reports its sheets + any ModelChoice tree sheets; reuses an already-open workbook of the same name. |
| `close_workbook` | **Close an open workbook** by file name (counterpart to `open_workbook`). Discards unsaved changes by default; pass `save=True` to save first. |
| `list_trees` | List the decision trees in a workbook with node-type counts. |
| `get_tree` | Full structure of one tree — decision / chance / terminal nodes, branches, probabilities, values. |
| `roll_up` | Roll the tree back to its expected values and **optimal policy** — the decision recommendation, in plain English. |
| `verify_rollback` | Cross-check our rollback against the `MC_V_<id>` cells ModelChoice itself wrote — a correctness guarantee when the tree has been rendered. |
| `run_scenarios` | **What-if comparison** — give named scenarios (bundles of input changes); each is rolled back and compared to the baseline (EV, optimal decision, Δ, whether the decision flips, which scenario wins). Pure Python, no Excel needed. |
| `run_evpi` | Expected Value of Perfect Information for the active tree — the most you'd pay for perfect information before deciding. Drives ModelChoice's headless `MC_EVPI_Auto`. |
| `run_risk_profile` | The **outcome distribution** for each decision option — expected value, min, max, std dev, plus the cumulative-probability table. Shows downside/upside, not just the average. Drives `MC_RiskProfile_Auto`. |
| `run_utility` | Apply a **risk attitude** (utility function) — returns the certainty equivalent, risk premium, and the optimal decision under risk aversion (can differ from the EV choice). Drives `MC_Utility_Auto`. |
| `run_evii` | Expected Value of Imperfect Information for a specific **test** — pass the target chance node and a likelihood matrix P(signal\|state); returns EVII, its net value after cost (worth running?), and the EVPI ceiling. Drives `MC_EVII_Auto`. |
| `run_decision_report` | Run a report — `strategy_table`, `policy_suggestion`, `decision_brief`, `mcda_report`, or `force_to_outcome` — and read it back (rows + label→value pairs + sheets) in one call. |
| `run_analysis` | Run any decision-analysis (`robustness`, `sensitivity`, `strategy_table`, `policy_suggestion`, `decision_brief`, `mcda_report`, `risk_profile`, `evpi`) and report the result sheets it produced. |
| `run_robustness` | Run the robustness ("break the decision") analysis and return a structured read — verdict, score, and the minimum input change that flips the decision. |
| `run_sensitivity` | Run one-way sensitivity and return the tornado-ordered report (which assumptions the decision is most sensitive to) + baseline EV + sheets. |
| `run_decision_report` | Run a decision report — `strategy_table`, `policy_suggestion`, `decision_brief`, or `mcda_report` — and read the primary result sheet back in one call (rows + label→value pairs + related sheets). |
| `read_sheet` | Read a result sheet's cells back (numbers / text), e.g. the robustness verdict or a sensitivity report. |

**Prompts:** `/design-decision-tree` walks an analyst through building a tree from a description (or pasted data) and analysing it. `/decision-tree-monte-carlo` then turns it into a Monte Carlo — put `Vose*` distributions on the uncertain inputs and run the simulation in `modelrisk-mcp` to get the distribution of the tree's expected value.

**Resources:** `modelchoice://guide/*` — curated decision-analysis guidance (building trees, EVPI vs EVII, which analysis answers which question, common pitfalls) the model can read for grounding.

Run it: `uv run python -m modelchoice_mcp` (stdio), or wire `modelchoice-mcp` into Claude Desktop like any MCP server.

## How it works

ModelChoice is a C# .NET (Excel-DNA) add-in. Unlike ModelRisk it exposes **no COM object model and no compute worksheet functions** — but it persists the *entire* decision tree as JSON in a very-hidden worksheet (`_MC_Store`), and the rollback logic is fully captured in that model. So this server can:

- **Read** every tree in a workbook by reassembling the `_MC_Store` payload (chunked across columns, v2-envelope or v1-legacy format) and parsing the model JSON.
- **Roll back** each tree to its expected values and **optimal policy** in pure Python — reproducing ModelChoice's own rollback (terminal payoff = accumulated branch values; chance = auto-normalized probability-weighted EV; decision = max/min by `Maximize`). **No add-in needs to be loaded** — it reads the saved model directly.

The Python roller is validated against ModelChoice's own authoritative C# test (`Model_Validate_Compile_Rollback_Works`): a decision/chance tree that rolls back to EV 50 with the optimal option selected — and matches exactly.

## Roadmap

- **Phase 0 — read + rollback engine (done):** parse `_MC_Store`, reconstruct each tree, roll back to EV + optimal policy. Validated against ModelChoice's own C# rollback test.
- **Phase 1 — the MCP server (done):** `list_trees`, `get_tree`, `roll_up`, proven live against a real workbook. `verify_rollback` cross-checks our EVs against the `MC_V_<id>` named ranges (live-verified: 4/4 nodes, 0 diff). CI (ruff + mypy + pytest) + a tag-driven PyPI release pipeline.
- **Phase 2 — drive analyses (done):** `run_analysis` + `read_sheet` drive any headless `MC_*_Auto` command; structured readers for `run_robustness`, `run_sensitivity`, `run_risk_profile`, and `run_decision_report` (strategy / policy / brief / mcda / force-to-outcome / two-way), plus standalone `run_evpi` (AB#2558) and `run_evii` (AB#2568).
- **Phase 3 — build / edit / extend trees (done):** `build_tree`, `edit_tree` (incl. add/remove options & outcomes), `build_control_panel`, `export_tree_json` / `import_tree_json`, `import_precisiontree` (AB#2632), `run_utility` (risk attitude, AB#2635), `build_mcda` (multi-criteria, AB#2637), and `set_input_distribution` (put a `Vose*` distribution on an input — the decision-tree half of a Monte Carlo). All live-verified against the shipped add-in.

The original phased roadmap is delivered. Candidate next steps (not yet scheduled):

- **Simulation hand-off (cross-server):** the `decision-tree-monte-carlo` prompt guides assigning distributions → wrapping the root EV as a ModelRisk output → running the simulation in `modelrisk-mcp` → reading the EV's output distribution back. A thin orchestration tool could automate the hand-off end-to-end.
- ~~**AHP weight elicitation for MCDA**~~ — *done (0.0.25, AB#2646):* `build_mcda` accepts a pairwise-comparison matrix (`weight_source="ahp"`) and derives the weights + consistency ratio via the add-in's `AhpCalculator`.

## Architecture

A separate server that reuses the patterns proven in `modelrisk-mcp` (xlwings bridge, dry-run-by-default safety on any future writes, packaging, MCP-registry flow). Decision analysis and Monte Carlo are distinct domains, so they ship as distinct servers.

## Development

```powershell
uv sync --extra dev
uv run pytest        # parser/roller + store-format tests (no Excel needed)
```

MIT licensed.
