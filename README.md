# ModelChoice MCP

**An open Model Context Protocol server for [Vose Software's ModelChoice](https://www.vosesoftware.com) — the decision-tree add-in for Excel.**

A sibling to [`modelrisk-mcp`](https://github.com/vosesoftware/modelrisk-mcp): where that server brings Monte Carlo risk modelling into a conversation, this one brings **decision analysis** — building, reading, rolling back, and analysing decision trees in Excel.

> **Status: `0.0.1` — Phase 0 (read-only) spike.** The core read path is proven end to end. Building/analysis tools are on the roadmap below.

## How it works

ModelChoice is a C# .NET (Excel-DNA) add-in. Unlike ModelRisk it exposes **no COM object model and no compute worksheet functions** — but it persists the *entire* decision tree as JSON in a very-hidden worksheet (`_MC_Store`), and the rollback logic is fully captured in that model. So this server can:

- **Read** every tree in a workbook by reassembling the `_MC_Store` payload (chunked across columns, v2-envelope or v1-legacy format) and parsing the model JSON.
- **Roll back** each tree to its expected values and **optimal policy** in pure Python — reproducing ModelChoice's own rollback (terminal payoff = accumulated branch values; chance = auto-normalized probability-weighted EV; decision = max/min by `Maximize`). **No add-in needs to be loaded** — it reads the saved model directly.

The Python roller is validated against ModelChoice's own authoritative C# test (`Model_Validate_Compile_Rollback_Works`): a decision/chance tree that rolls back to EV 50 with the optimal option selected — and matches exactly.

## Roadmap

- **Phase 0 (done):** read-only — list trees, read structure, roll back to EV + optimal policy, cross-check against the `MC_V_<id>` named ranges.
- **Phase 1:** drive analyses. ModelChoice's analyses (sensitivity, EVPI, EVII, strategy table, robustness) are currently UI/modal-dialog only; the clean path — since we own the source — is to add headless `[ExcelCommand]` entry points mirroring the existing `MC_RiskProfile_Auto`, then invoke + read them via `Application.Run`.
- **Phase 2:** build/edit trees — write the model JSON into `_MC_Store` and trigger a re-render.

## Architecture

A separate server that reuses the patterns proven in `modelrisk-mcp` (xlwings bridge, dry-run-by-default safety on any future writes, packaging, MCP-registry flow). Decision analysis and Monte Carlo are distinct domains, so they ship as distinct servers.

## Development

```powershell
uv sync --extra dev
uv run pytest        # parser/roller + store-format tests (no Excel needed)
```

MIT licensed.
