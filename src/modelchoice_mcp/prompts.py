"""Guided prompts for ModelChoice MCP.

Importing this module side-effects the ``@mcp.prompt`` registrations
into the shared FastMCP instance, like the tools module.
"""

from __future__ import annotations

from modelchoice_mcp.server import mcp

_DESIGN_DESCRIPTION = (
    "Walk an analyst through building a decision tree by describing the "
    "decision (or pasting data), then build it with build_tree and analyse it."
)

_DESIGN_TEMPLATE = """\
You are helping an analyst turn a decision into a ModelChoice decision
tree, using the modelchoice-mcp tools. The goal: the analyst describes
the decision in plain language (or pastes data), and you build the tree
for them — asking only what you genuinely need.

Work in this order:

1. **Elicit the structure.** Identify, in order:
   - The **decisions** (choices the analyst controls — "launch or not").
   - The **chance** events (uncertain outcomes — "market is good/bad"),
     each with its branches and **probabilities** (they need not sum to
     1; ModelChoice auto-normalises, but check the intent).
   - The **terminal payoffs** and the **cash flow on each branch**
     (cost/benefit incurred down that path).
   - Whether the objective is to **maximize** (profit/value) or
     **minimize** (cost/loss).
   Draw the structure back to the analyst in words before building.

2. **If the analyst pasted data**, first lay the **input parameters out
   in a table** (parameter, value, where it's used), and ask:
   - Which of these are **uncertain** vs fixed?
   - For the uncertain ones, is there a point estimate plus a range, or
     historical data to fit? (These become Monte Carlo distributions —
     see step 5.)
   Confirm the mapping from data → tree before building.

3. **Preview the tree.** Call `build_tree` with `dry_run=True`. It
   validates the structure (catching cycles, missing children) and
   returns the **rolled-back expected value and the optimal policy**.
   Show the analyst the recommendation and the structure. Iterate the
   spec until they're happy. *Never commit on the first pass.*

4. **Commit.** When confirmed, call `build_tree` with `dry_run=False`
   to write and render it in Excel. Then `roll_up` to restate the
   recommendation, and offer `run_robustness` (how much would have to
   change to flip the decision?) and `run_sensitivity` (which
   assumptions matter most) and `run_evpi` (is more information worth
   paying for?).

5. **Uncertain inputs → Monte Carlo (optional, powerful).** A branch
   value or probability that the analyst flagged as uncertain can be a
   Vose distribution rather than a point number. Use
   `set_input_distribution` to put a `Vose*` formula on that input, then
   (if the `modelrisk` MCP server is available) run the simulation to get
   the *distribution* of the tree's expected value, not just its point
   estimate. The `/decision-tree-monte-carlo` prompt walks this through.

Pace it like a conversation. Build the analyst's understanding, confirm
each commitment, and let the tree fill in around the discussion.
"""


_SIMULATE_DESCRIPTION = (
    "Turn a built decision tree into a Monte Carlo: put Vose distributions on "
    "its uncertain inputs, then run the simulation in modelrisk-mcp to get the "
    "distribution of the tree's expected value (not just the point estimate)."
)

_SIMULATE_TEMPLATE = """\
You are taking an existing ModelChoice decision tree and evaluating it
under input uncertainty — a Monte Carlo over the decision. This spans
**two MCP servers**: `modelchoice` (the tree) and `modelrisk` (the
simulation engine). Confirm both are connected before you start; if
`modelrisk` is not available, stop and say so.

The idea: a normal rollback gives ONE expected value from point inputs.
If the inputs are uncertain, the EV is itself a distribution. ModelRisk
samples the `Vose*` formulas you place on the tree's input cells, the
add-in re-rolls the tree each iteration, and you read back the
distribution of the root EV.

Work in this order:

1. **Pick the tree and its inputs.** `list_trees` / `get_tree` to find
   the tree and its branch values and probabilities. With the analyst,
   decide which inputs are genuinely uncertain (cash flows, market
   probabilities) versus fixed. Note the **root node id** — the root EV
   lives in the `MC_V_<rootId>` named range, which will be your
   simulation output.

2. **Assign a distribution to each uncertain input.** For each, choose a
   ModelRisk distribution that matches what's known — a point estimate
   plus a range → `VosePERT(min, most_likely, max)`; a mean and spread →
   `VoseNormal(mean, sd)`; a bounded fraction → `VoseBeta`/`VosePERT`; raw
   data → ask `modelrisk` to fit one. Then call
   `set_input_distribution(node_id, outcome, distribution, kind)` with
   `kind='value'` for a cash flow or `kind='probability'`. (If `modelrisk`
   exposes `propose_distributions_for_inputs`, use it to suggest the
   forms first.) Probabilities across a chance node should still be
   coherent — if you randomise one, consider how the complement behaves.

3. **Mark the output.** Using the `modelrisk` server, wrap the
   `MC_V_<rootId>` cell as a simulation output (e.g. `wrap_with_output`)
   so ModelRisk collects the rolled-back EV each iteration. Name it
   clearly (e.g. "Decision EV").

4. **Run the simulation.** Use `modelrisk`'s `run_simulation`. ModelRisk
   resamples the `Vose*` input cells, Excel recalculates, the ModelChoice
   UDFs re-roll the tree, and the output cell captures the EV per
   iteration.

5. **Read the result back and interpret.** Pull the output distribution
   (`get_simulation_results` / `get_samples` / `get_tail_risk`): report
   the **mean** (compare it to the deterministic rollback EV — they should
   be close if the model is roughly linear), the **spread** (P10 to P90), the
   **downside** (probability the EV is below zero or below a threshold),
   and which decision stays optimal across the runs. The headline: not
   "the EV is X" but "the EV is centred near X with this much downside
   risk, and the recommended decision is/ isn't robust to it."

6. **Tidy up (optional).** If the analyst wants the deterministic model
   back, the distributions live as user-formulas on the input cells;
   note that re-running `set_input_distribution` (or editing the cell)
   changes them, and ModelRisk's restore tools revert any output wrapper.

Keep the analyst oriented: every distribution you add is an explicit
statement about what's uncertain and how much. Surface those choices,
don't bury them.
"""


@mcp.prompt(name="design-decision-tree", description=_DESIGN_DESCRIPTION)
def design_decision_tree_prompt() -> str:
    return _DESIGN_TEMPLATE


@mcp.prompt(name="decision-tree-monte-carlo", description=_SIMULATE_DESCRIPTION)
def decision_tree_monte_carlo_prompt() -> str:
    return _SIMULATE_TEMPLATE
