## Big picture

The optimiser is trying to find the best mix of the seven composites, subject to rules: minimum and maximum weights per asset class, a volatility target and, for the Income strategies, a minimum yield. Penalisation is the way those rules are enforced. The model does not forbid a portfolio that breaks a rule. Instead, it reduces that portfolio's score by an amount that grows with the size of the breach. A portfolio that breaks the rules therefore looks unattractive to the optimiser rather than impossible.

## Conceptual approach

In constrained optimisation, constraints can be handled in two ways.

**Hard constraints** reject any candidate that violates them outright.

**Soft constraints (penalty method)** turn the constrained problem into an unconstrained one by changing the objective:

> Penalised score = Objective − Σ Penalty(deviation)

Deviation is how far a candidate is outside a limit, and it is zero when the constraint is met. Three design choices then decide how the penalty behaves:

1. **Shape of the penalty function.** A quadratic penalty, (k × deviation)², is smooth and forgiving of tiny breaches. Because it squares the deviation, it gets steep quickly for larger ones.
2. **Scaling factor (k).** This sets the trade-off between return and compliance. If k is too low, the optimiser happily "buys" return by breaching limits. If k is too high, the search landscape becomes cliff-like, and a heuristic search can get stuck.
3. **Relative scaling across constraints.** A larger k makes a constraint behave more like a hard limit.

Penalties are especially natural for RISKOptimizer. It uses a genetic or evolutionary search over trial portfolios rather than an analytical solver, so it needs a single number to rank each candidate.

## How it works in this model

- **Objective:** maximise the *mean* of W97, which is simulated portfolio value at year 20, starting from a notional value of 2, across 500 (then 1,000) Monte Carlo iterations.
- **Decision variables:** C48:C54, integers from 0 to 1,000 (weights × 1,000). This gives a granularity of 0.1%.
- **Constraints:** all are soft. They are min/max limits per composite (D48:D54 compared against the bounds in C7:C21), plus annualised volatility and min yield.
- **Penalty functions:**
  - Volatility and yield: (1000 × deviation)²
  - All other constraints: (100 × deviation)²

Deviations are expressed as decimals (1% = 0.01), so the penalty for each 1% breach is:

| Breach | Weight constraint | Volatility / yield |
|---|---|---|
| 0.1% | 0.01 | 1 |
| 0.5% | 0.25 | 25 |
| 1.0% | 1 | 100 |

For scale, a portfolio compounding at about 4% for 20 years ends at roughly 4.4. A 1% volatility breach (penalty 100) therefore wipes out any conceivable return advantage, so volatility and yield behave almost as hard constraints. Weight limits are firm for breaches of 1% or more. However, a breach of 0.1–0.2% is nearly free.

## Simple example

Take a hypothetical Balanced strategy with max equity of 60% and a volatility target of 9%:

| Candidate | Equity | Vol | Mean W97 | Penalty | Score |
|---|---|---|---|---|---|
| A (compliant) | 60.0% | 9.0% | 4.30 | 0 | **4.30** |
| B | 60.5% | 9.0% | 4.36 | (100×0.005)² = 0.25 | 4.11 |
| C | 60.0% | 9.2% | 4.40 | (1000×0.002)² = 4.00 | 0.40 |
| D | 60.1% | 9.0% | 4.32 | (100×0.001)² = 0.01 | **4.31** |

B and C lose, as intended. **D wins despite breaching the equity limit**, because its extra return (0.02) outweighs a tiny quadratic penalty (0.01).

## Points worth noting for the validation

- **Final weights can breach limits.** Soft constraints mean the final weights may sit marginally outside policy limits. Nothing in the documented process checks the optimised output against the limits after the run.
- **The scaling factors are expert judgement.** The 100 and 1,000 values are undocumented, and their effect depends on the units of the constraint cells. I've assumed decimals here, and that is worth confirming with the model owner.
- **Penalties are compared against a noisy objective.** The objective is a simulated mean, so Monte Carlo noise in mean W97 from 500 iterations may be of the same order as small penalties. The ranking between near-compliant candidates (A and D above) could partly reflect simulation noise rather than genuine differences.

Since UKRM couldn't run @RISK, these points come from the documented settings rather than tested behaviour.

### When both constraints are breached
When both are breached, the penalties are calculated separately and then added together:

> Score = Mean W97 − (100 × weight deviation)² − (1000 × volatility deviation)²

Each constraint is squared on its own and the results are summed. There is no joint or interaction term.

**Example** (max equity 60%, vol target 9%):

| Candidate | Equity | Vol | Mean W97 | Weight penalty | Vol penalty | Score |
|---|---|---|---|---|---|---|
| E | 60.5% | 9.1% | 4.38 | (100×0.005)² = 0.25 | (1000×0.001)² = 1.00 | **3.13** |

A few points follow from this:

- **Volatility dominates.** A 0.1% volatility breach costs four times as much as a 0.5% equity breach. The optimiser will fix the volatility breach first.
- **The two breaches tend to come together.** Adding equity usually raises portfolio volatility, so an equity overweight often triggers both penalties at once. In practice that makes equity breaches costlier than the weight penalty alone suggests.
- **Total penalty is lower than for one large breach.** Because each breach is squared separately, two small breaches cost less than a single breach of their combined size, since 0.5² + 0.5² < 1². Spreading small violations across several constraints is therefore relatively cheap. This adds to the earlier point about marginal breaches surviving in the final output.
