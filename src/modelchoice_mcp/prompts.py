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
   Vose distribution rather than a point number. If the `modelrisk`
   MCP server is available, propose the right distribution for each
   uncertain input and place it in the tree's cell, so the decision is
   evaluated under uncertainty — decision analysis and Monte Carlo
   together.

Pace it like a conversation. Build the analyst's understanding, confirm
each commitment, and let the tree fill in around the discussion.
"""


@mcp.prompt(name="design-decision-tree", description=_DESIGN_DESCRIPTION)
def design_decision_tree_prompt() -> str:
    return _DESIGN_TEMPLATE
