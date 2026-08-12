"""Curated decision-analysis knowledge, exposed as MCP resources.

Importing this module side-effects the ``@mcp.resource`` registrations
into the shared MCPServer instance, like the tools and prompts modules.
These give the model grounded guidance on *how* to use the ModelChoice
tools well — what each analysis answers, how rollback works, and the
common modelling pitfalls — without it having to guess.
"""

from __future__ import annotations

from modelchoice_mcp.server import mcp

_DECISION_TREES = """\
# Building a decision tree

A ModelChoice tree has three node kinds:

- **Decision (□)** — a choice you control (launch / wait / license). Its
  options each carry a cash flow (`value`, e.g. a cost is negative).
- **Chance (○)** — an uncertainty you don't control (market is strong/weak).
  Its branches carry a `probability` (auto-normalised if they don't sum to 1)
  and a cash flow.
- **Terminal (◁)** — an end point with a final payoff (often 0 when all the
  cash sits on the branches leading to it).

**Rollback** computes the value bottom-up:
- terminal = the accumulated branch cash flows down its path,
- chance = the probability-weighted average of its branches (an *expected* value),
- decision = the **best** option by the objective (maximise EV by default; set
  `Maximize=false` to minimise, e.g. cost/loss).

The root's rolled-back value is the model's EV; the sequence of best options is
the **optimal policy** — the recommendation. Build with `build_tree`, tweak with
`edit_tree`, read the recommendation with `roll_up`.

**Conventions that avoid double-counting:** put a cost/revenue on the branch
where it occurs, not also on the terminal. A decision option that "does nothing"
(e.g. *Don't launch*) usually leads straight to a terminal of 0.
"""

_VALUE_OF_INFORMATION = """\
# Value of information: EVPI vs EVII

Both answer "how much should I pay to learn more before deciding?"

- **EVPI** (Expected Value of *Perfect* Information) — the ceiling. The most you
  should ever pay for a perfect crystal ball on the uncertainties. If EVPI is
  ~0, no information can change your decision — don't pay for any. Tool: `run_evpi`.
- **EVII** (Expected Value of *Imperfect* Information, a.k.a. EVSI) — the value of
  a *real, fallible* test (a survey, a pilot, a diagnostic). Always **≤ EVPI**.
  You give it the test's reliability as a likelihood matrix P(signal | state).
  Tool: `run_evii` — it returns EVII, the test cost, and the **net** value
  (EVII minus cost); run the test only if net > 0.

Rule of thumb: compute EVPI first. If it's below the test's cost, skip EVII — even
perfect information isn't worth that much here.
"""

_CHOOSING_ANALYSIS = """\
# Which analysis answers which question

| Question | Tool |
|---|---|
| What should I decide, and what's the EV? | `roll_up` |
| How good/bad could it get (downside, upside, spread)? | `run_risk_profile` |
| Which assumptions matter most? | `run_sensitivity` (one-way, tornado) |
| How two inputs interact / where the decision flips | two-way sensitivity |
| How wrong can my inputs be before the decision flips? | `run_robustness` |
| What would have to be true to justify a chosen option? | `run_decision_report force_to_outcome` |
| Should I gather more information first? | `run_evpi` (ceiling), `run_evii` (a specific test) |
| The optimal action in *every* scenario | `run_decision_report strategy_table` |
| A recommended policy + plain-English rationale | `run_decision_report policy_suggestion` |
| An executive summary | `run_decision_report decision_brief` |
| Compare several named what-if cases | `run_scenarios` |
| Multi-criteria choice (not pure money) | MCDA (`run_decision_report mcda_report`) |

EV is the headline; a risk profile is what stops you over-trusting it.
"""

_PITFALLS = """\
# Common decision-tree pitfalls

1. **Double-counting cash** — putting the same cost on a branch *and* its
   terminal. Put each cash flow in one place.
2. **EV tunnel vision** — two options with equal EV can have very different
   risk. Always look at `run_risk_profile` before committing.
3. **Forgetting the "do nothing" option** — the status quo is a real
   alternative; model it (usually a 0 terminal) so the tree can reject action.
4. **Probabilities that don't reflect evidence** — sensitivity + robustness tell
   you which guesses actually matter, so you can refine only those.
5. **Paying for information reflexively** — check EVPI first; if it's near zero,
   no study is worth running.
6. **Minimise vs maximise** — for costs/losses set `Maximize=false`, or the tree
   will pick the worst option.
7. **Sunk costs in the tree** — only model cash flows that change with the
   decision from here forward.
"""

_RESOURCES = {
    "decision-trees": ("Building a decision tree", _DECISION_TREES),
    "value-of-information": ("EVPI vs EVII", _VALUE_OF_INFORMATION),
    "choosing-analysis": ("Which analysis answers which question", _CHOOSING_ANALYSIS),
    "pitfalls": ("Common decision-tree pitfalls", _PITFALLS),
}


@mcp.resource("modelchoice://guide/decision-trees", name="Building a decision tree")
def guide_decision_trees() -> str:
    return _DECISION_TREES


@mcp.resource("modelchoice://guide/value-of-information", name="EVPI vs EVII")
def guide_value_of_information() -> str:
    return _VALUE_OF_INFORMATION


@mcp.resource("modelchoice://guide/choosing-analysis", name="Which analysis to use")
def guide_choosing_analysis() -> str:
    return _CHOOSING_ANALYSIS


@mcp.resource("modelchoice://guide/pitfalls", name="Common decision-tree pitfalls")
def guide_pitfalls() -> str:
    return _PITFALLS


__all__ = [
    "guide_choosing_analysis",
    "guide_decision_trees",
    "guide_pitfalls",
    "guide_value_of_information",
]
