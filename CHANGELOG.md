# Changelog

All notable changes to `modelchoice-mcp`. Versions are tag-driven; pushing a
`vX.Y.Z` tag publishes to PyPI via the release workflow.

## Unreleased
- **The wheel is now installed and booted before it is published (AB#3143).** The wheel is
  what PyPI serves and what `pip install` gives you, but nothing ever installed it — the exe
  check gates the PyInstaller binary built from source, a different artifact with different
  dependency resolution, so a wheel-only fault (a missing `packages` entry, a bad hatch
  build glob, an undeclared dependency) would have shipped. The `build` job now installs the
  freshly-built wheel into a clean virtualenv, makes it answer a real MCP `initialize`
  handshake, and asserts the version it reports matches the tag being released — so a
  forgotten version bump can no longer burn a PyPI version number.
- **The PyPI publish now waits for the Windows exe verification (AB#3142).** `publish-pypi`
  depended on `build` alone, so `build-windows-exe` — which packs the single-file exe and
  asserts it answers an MCP `initialize` handshake — raced the publish instead of gating
  it, and a packaging break would still have reached PyPI. Publishing is the one
  irreversible step in the workflow (a version number can never be reused), so it now runs
  after every verification. Costs a few minutes per release; buys the ability to fix a bad
  build by deleting a tag rather than burning a version. Release-workflow only — no change
  to the package.

## 0.0.31
- **Migrated to the mcp 2.0 SDK (AB#3134).** mcp 2.0 removed `mcp.server.fastmcp` and
  renamed the high-level server class `FastMCP` → **`MCPServer`** (`from mcp.server import
  MCPServer`). The `@mcp.tool` / `@mcp.resource` / `@mcp.prompt` decorators are unchanged
  and still return the plain function, so all 25 tools, 4 resources and 2 prompts register
  exactly as before — verified name-for-name against the 1.x baseline. The dependency is
  now `mcp>=2,<3`; the major-version cap stays, deliberately.
  - **HTTP transports:** host and port are now `run()` keyword arguments
    (`mcp.run(transport=..., host=..., port=...)`) instead of mutating `mcp.settings`
    beforehand. `--transport`, `--host` and `--port` behave as before.
  - The server now reports its own version in `serverInfo` — `MCPServer` takes a `version`
    argument, which `FastMCP` had no equivalent for.

## 0.0.30
- **CI stops drifting with upstream releases.** `uv.lock` is now committed (it was in
  `.gitignore`), and CI syncs with `--locked`. Previously CI had no pins at all: it
  re-resolved the open floors in `pyproject.toml` on every run, so a new upstream major
  landed the day it shipped — `mcp 2.0.0` moved `mcp.server.fastmcp` and broke `mypy`
  across every `@mcp.tool` on an **unchanged `main`**. `mcp` is also capped to `<2` for
  people installing from PyPI, who resolve against the constraints rather than the lock.
  Editing `pyproject.toml` now requires re-running `uv lock` and committing the result;
  CI fails with a clear message if you forget, instead of silently upgrading you.
- **`export_tree_json` now carries a `generator` field (AB#3123)** — product
  (`ModelChoice by Vose Software`), server version, UTC export timestamp and the
  product URL, so an exported tree stays attributable once it lands in version
  control, a ticket, or someone else's repository. It sits *alongside*
  `model_json`, never inside it: `model_json` still round-trips byte-identically
  through `import_tree_json`. Part of the ModelChoice output-branding work
  (Feature AB#3118).

## 0.0.29
- **Licence gate (AB#2659)** — building and analysis ACTIONS now require a fully
  licensed ModelChoice. The bridge reads the add-in's licence state via the new
  headless `MC_LicenseStatus_Auto` and refuses actions (build/edit commit,
  build_mcda, control panel, set_input_distribution, run_utility / evii / evpi /
  risk_profile / decision_report / robustness / sensitivity / analysis, import)
  unless `isComplete` (full licence). **Reading is unaffected** — list/get/roll_up/
  verify/scenarios/export and open/close workbook work regardless. New read-only
  **`license_status`** tool reports the state. Fail-closed: if the status can't be
  read (add-in missing/old), actions are blocked with an actionable message.
  (Trial/expired users can read but not drive actions.) Needs the add-in build
  with `MC_LicenseStatus_Auto`.
- **`.mcpb` now built with the official `mcpb` CLI** (`@anthropic-ai/mcpb`,
  validates during pack) instead of a hand-rolled zip; plain-zip fallback when
  node isn't present. README install section reworked: PyPI + config is the
  recommended path; the `.mcpb` one-click carries a note that the **Claude
  Desktop Extensions installer silently no-ops on the latest Windows MSIX builds
  (a client bug, not the bundle)** — use `pip install` until Anthropic patches it.

## 0.0.28
- **One-click install: Claude Desktop Extension (`.mcpb`)** — the release now
  also builds a standalone Windows `.exe` (new PyInstaller spec) and wraps it in
  a `.mcpb` Desktop Extension attached to the GitHub release. Installing into
  Claude Desktop becomes **1. open `modelchoice-mcp.mcpb`  2. restart Claude** —
  no Python, no `claude_desktop_config.json` editing. Built CLI-free in CI
  (`scripts/build_mcpb.py`); the bundle version is injected from the release tag.
  Excel + the ModelChoice add-in are still required for rendering.

## 0.0.27
- **`close_workbook`** — close an open workbook by file name (counterpart to
  `open_workbook`). By default unsaved changes are discarded (`save=False`);
  pass `save=True` to write them first. (Mirrors `close_workbook` in
  modelrisk-mcp.)
- **`open_workbook` suppresses Excel's open prompts** — `Workbooks.Open` is now
  called with `update_links=False`, `ignore_read_only_recommended=True`,
  `notify=False`, `add_to_mru=False` and with `DisplayAlerts`/`AskToUpdateLinks`
  off, so a headless open can't hang on an Update-Links / read-only dialog.
  (External links aren't refreshed on open; values stay as last saved.)

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
