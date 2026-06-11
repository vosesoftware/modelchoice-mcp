# ModelChoice MCP

**An open Model Context Protocol server for [Vose Software's ModelChoice](https://www.vosesoftware.com) — the decision-tree add-in for Excel.**

A sibling to [`modelrisk-mcp`](https://github.com/vosesoftware/modelrisk-mcp): where that server brings Monte Carlo risk modelling into a conversation, this one brings **decision analysis** — building, reading, rolling back, and analysing decision trees in Excel.

> **Status: `0.0.1` — Phase 1 MCP server (read-only).** Three tools (`list_trees`, `get_tree`, `roll_up`) over a proven read path. Building/analysis tools are on the roadmap below.

## Tools

| Tool | What it does |
|---|---|
| `list_trees` | List the decision trees in a workbook with node-type counts. |
| `get_tree` | Full structure of one tree — decision / chance / terminal nodes, branches, probabilities, values. |
| `roll_up` | Roll the tree back to its expected values and **optimal policy** — the decision recommendation, in plain English. |
| `verify_rollback` | Cross-check our rollback against the `MC_V_<id>` cells ModelChoice itself wrote — a correctness guarantee when the tree has been rendered. |
| `run_evpi` | Expected Value of Perfect Information for the active tree — the most you'd pay for perfect information before deciding. Drives ModelChoice's headless `MC_EVPI_Auto`. *(Needs a ModelChoice build that includes that command — Phase 2.)* |

Run it: `uv run python -m modelchoice_mcp` (stdio), or wire `modelchoice-mcp` into Claude Desktop like any MCP server.

## How it works

ModelChoice is a C# .NET (Excel-DNA) add-in. Unlike ModelRisk it exposes **no COM object model and no compute worksheet functions** — but it persists the *entire* decision tree as JSON in a very-hidden worksheet (`_MC_Store`), and the rollback logic is fully captured in that model. So this server can:

- **Read** every tree in a workbook by reassembling the `_MC_Store` payload (chunked across columns, v2-envelope or v1-legacy format) and parsing the model JSON.
- **Roll back** each tree to its expected values and **optimal policy** in pure Python — reproducing ModelChoice's own rollback (terminal payoff = accumulated branch values; chance = auto-normalized probability-weighted EV; decision = max/min by `Maximize`). **No add-in needs to be loaded** — it reads the saved model directly.

The Python roller is validated against ModelChoice's own authoritative C# test (`Model_Validate_Compile_Rollback_Works`): a decision/chance tree that rolls back to EV 50 with the optimal option selected — and matches exactly.

## Roadmap

- **Phase 0 (done):** the read + rollback engine — parse `_MC_Store`, reconstruct each tree, roll back to EV + optimal policy. Validated against ModelChoice's own C# rollback test.
- **Phase 1 (done):** the MCP server — `list_trees`, `get_tree`, `roll_up`, proven live against a real workbook.
- **Phase 1b (done):** `verify_rollback` cross-checks our EVs against the `MC_V_<id>` named ranges; CI (ruff + mypy + pytest) and a tag-driven PyPI release pipeline; wheel + sdist build. *(First publish awaits a PyPI trusted-publisher; the live `MC_V_` cross-check awaits a real rendered tree.)*
- **Phase 2 (in progress):** drive analyses. ModelChoice's analyses are UI/modal-dialog only; the clean path — since we own the source — is to add headless `[ExcelCommand]` entry points mirroring `MC_RiskProfile_Auto`, then invoke + read them via `Application.Run`. **EVPI** is the first vertical slice: `MC_EVPI_Auto` (ModelChoice, AB#2558) + the `run_evpi` tool here. Sensitivity, robustness, strategy table, and EVII follow.
- **Phase 3:** build/edit trees — write the model JSON into `_MC_Store` and trigger a re-render.

## Architecture

A separate server that reuses the patterns proven in `modelrisk-mcp` (xlwings bridge, dry-run-by-default safety on any future writes, packaging, MCP-registry flow). Decision analysis and Monte Carlo are distinct domains, so they ship as distinct servers.

## Development

```powershell
uv sync --extra dev
uv run pytest        # parser/roller + store-format tests (no Excel needed)
```

MIT licensed.
