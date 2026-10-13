# Model Optimisation and Outputs

Oct 13, 2026 · @Ariel Gimenez

This section describes the methodology behind the optimisation and simulation steps of the SAA model, and the outputs they produce. It complements the Model Implementation section: where that section follows the process step by step through the Excel workbooks, this one looks at what each step does conceptually and what it assumes. The composite-level inputs (expected returns, volatilities, correlations and yields) are taken as given, as they are covered in earlier sections.

## Optimisation

### Problem formulation

For each strategy, the model chooses the weights of seven asset class composites: equity, developed market government bonds, other bonds, property, gold, hedge funds and cash. The aim is to maximise the expected value of the portfolio at the end of a 20-year horizon, subject to the strategy's constraints:

- a target annualised volatility, which defines the strategy's risk level;
- minimum and maximum weights for each composite;
- for the Income strategies, a minimum portfolio yield;
- a long-only, fully invested portfolio (weights between 0% and 100%, summing to 100%).

The volatility targets are set at equally spaced points between 30% and 90% of expected equity volatility, so the Core strategies form a ladder of increasing risk. The Income and Responsible strategies take their yield and hedge fund constraints from the optimised Core results.

### What the objective implies

If composite returns are independent over time and the portfolio is rebalanced to fixed weights, the expected terminal value grows with the portfolio's expected return. Maximising it is therefore equivalent to maximising the portfolio's expected return. In mean-variance terms, the problem is:

> maximise w′μ subject to √(w′Σw) ≤ σ\*, w′y ≥ y\_min (Income only), lower ≤ w ≤ upper, Σw = 1

where w is the vector of composite weights, μ the expected returns, Σ the covariance matrix implied by the volatilities and correlations, y the yields and σ\* the volatility target.

The objective is linear, and every constraint is either linear or convex. The problem therefore has a well-defined global optimum, which a standard deterministic solver can find exactly. The result is a single point on the constrained efficient frontier at each strategy's volatility target. When the volatility constraint binds, the optimum takes all the risk allowed and allocates it to the composites with the highest expected return per unit of risk, after allowing for correlation.

### Solution method

The model does not solve this problem analytically. It uses a stochastic, population-based search (Palisade RISKOptimizer) in which each candidate portfolio is scored by a Monte Carlo simulation of its terminal value. Constraints are applied as soft constraints: a breach is not forbidden but reduces the score by a quadratic penalty, (1,000 × deviation)^2 for the volatility and yield constraints and (100 × deviation)^2 for the composite bounds. The search starts from the prior-year weights and runs in two passes of increasing simulation size.

This method has three practical consequences:

- **Approximate solution.** The result lies close to, but not exactly at, the true optimum, and may slightly breach constraints where the extra return outweighs the penalty.
- **Simulation noise.** Because the score of each candidate is itself an estimate, two runs with the same inputs can return different weights.
- **Anchoring to the prior year.** Starting from last year's allocation favours stability between years, but the final weights can depend on where the search began.

### Stability of the solution

Return-maximising optimisers are highly sensitive to differences in expected returns and tend to concentrate in a few assets. The model limits this in two ways: optimising across a small number of broad composites rather than individual indices, and imposing minimum and maximum weights on each composite. It does not use statistical techniques such as resampling or shrinkage of the inputs, so the stability of the result relies on these structural choices and on the bounds.

## Simulation

Once the weights are set, each strategy is assessed through two complementary simulations: a forward-looking Monte Carlo projection and a backward-looking historic back-test.

### Monte Carlo projection

The projection generates 10,000 scenarios of portfolio value over the 20-year horizon. In each scenario, returns for the seven composites and for inflation are drawn jointly from their expected returns and volatilities. Dependence between them is imposed with @RISK's rank-order method, which applies the input correlation matrix as Spearman rank correlations and leaves the shape of each marginal distribution unchanged. The marginal distributions assumed for each composite are not set out in the model documentation.

The same simulation engine scores candidate portfolios during the optimisation, with fewer iterations. The projected characteristics are therefore consistent with the objective that produced the weights.

The projection inherits the assumptions of its inputs: expected returns, volatilities and correlations are constant over the 20 years, returns are independent from one period to the next, and there is no valuation reversion, which the Bank handles through its tactical process. Under these assumptions, the projection describes a stable "steady-state" world. It is likely to understate the clustering of losses seen in real crises, so projected drawdowns should be read as indicative.

### Historic back-test

The back-test applies the strategy's weights to the actual monthly returns of the composites over the last 20 years. The portfolio is held without trading within each calendar year and rebalanced to its target weights at each year-end. The portfolio yield is the value-weighted average of the composite yields. Where an index within a composite has a shorter history, the composite is rebuilt from its remaining indices for the missing period.

### How the two views fit together

The two simulations answer different questions. The projection shows the range of outcomes the model expects under its own assumptions. The back-test shows how the allocation would have behaved through real events, including the 2008 financial crisis, the 2020 pandemic sell-off and the 2022 rise in interest rates, with their fat tails, changing correlations and persistent losses.

The back-test is not an independent test of the allocation. The volatilities and correlations used to set the weights are estimated over a history that overlaps with the back-test window, so the allocation is, to some extent, fitted to the period over which it is evaluated. The back-test also follows a single historical path, and historic returns over the window differ from the forward-looking expected returns. Its results illustrate behaviour under stress but should not be read as a forecast.

## Model outputs

For each of the ten strategies, the model produces three sets of outputs.

### Optimal weights

The primary output is the optimal weight of each of the seven composites. These are broken down to individual indices using the fixed within-composite weights, which gives the index-level allocation used to build portfolios. The expected return, volatility, yield and correlation inputs behind each strategy are stored alongside its weights, so each set of results can be traced to the assumptions that produced it.

### Historic and projected characteristics

The historic and projected characteristics describe each strategy along the same dimensions, so the two views can be read side by side.

| Dimension | Historic back-test (last 20 years) | Monte Carlo projection (next 20 years) |
| --- | --- | --- |
| Return | Annualised return; excess over cash and over inflation | Mean annualised return; returns at the 90% and 10% probability levels; mean excess over cash and over inflation |
| Volatility | Annualised volatility; downside volatility | Mean annualised volatility |
| Short-term losses | Maximum and minimum 12-month and monthly returns | Mean minimum 12-month return |
| Drawdown | Maximum drawdown, its peak and trough dates and months to recover | Maximum drawdown, as a mean and at the 10% probability level |
| Income | Average gross yield | Current gross yield |
| Risk-adjusted measures | Sharpe, Sortino, Calmar and adjusted Sharpe ratios; skewness and kurtosis of monthly returns | None |

The historic characteristics are a single realised outcome, while the projected characteristics are averages and percentiles across 10,000 scenarios. Differences between them, for example a deeper historic drawdown than the projected mean, reflect both the assumptions of the projection and the particular events in the back-test window, as discussed above.

Together, these outputs show whether each strategy's risk and return profile is consistent with its intended position on the risk ladder, and they support the communication of the SAA to the wider investment team.

### Validation Assessment

The optimisation and simulation steps of the SAA model are run within Palisade @RISK and its RISKOptimizer module, a third-party Excel add-in. UKRM does not have access to a licensed installation of the add-in. As a result, UKRM could not independently run the optimisation and simulation steps in the Model Owner's workbooks, or reproduce the optimal weights and the projected characteristics reported by the model.

UKRM's assessment of these steps is therefore based on a desk-based review of three sources:

- the MO's instructions document, including the optimiser set-up, constraint definitions, penalty functions and runtime settings shown in it;
- the structure of the workbooks;
- the VBA macros in the SAA Model workbook.

The review of the macros is a static code review. The logic was read and assessed line by line, but the code was not executed. Any observations on the macros, including those raised in the Model Implementation section, have therefore not been confirmed by running the code and should be validated with the MO.

The lack of access also limited the scope of the independent testing. In particular, UKRM was not able to:

- reproduce the MO's results for any strategy, or confirm that the reported weights are the output of the documented set-up;
- assess how sensitive the results are to changes in the volatility targets, the composite bounds, the minimum yield or the penalty functions;
- test alternative strategies or constraint sets beyond those run by the MO;
- assess the stability and convergence of the optimiser across repeated runs, random seeds, starting points or numbers of iterations;
- verify how @RISK applies the correlation matrix and what marginal distributions it uses, beyond what is stated in the vendor documentation.

UKRM's independent challenger implementation partly mitigates this limitation. It replicates the optimisation and simulation approach outside @RISK and compares it with an exact solution of the same problem, as described in the challenger section of this report. However, the challenger is an approximation of the vendor tool, not a re-run of it. It cannot confirm the behaviour of RISKOptimizer itself, and differences between the challenger and the model's outputs cannot be fully attributed without access to the original environment.

UKRM considers this a limitation on the scope of the validation, which should be recorded as such. Consistent with the expectation that third-party models and tools are subject to the same standards as internally developed ones, the MO should address it through one of the following routes:

- give UKRM access to a licensed installation, or arrange a supervised run in which UKRM specifies the scenarios to be executed;
- provide execution evidence for each annual run, including the optimiser settings, random seed, run logs and full outputs for each strategy;
- obtain and retain vendor documentation and evidence of testing for the @RISK functionality the model relies on.

Until one of these is in place, the conclusions of this report on the optimisation and simulation steps rest on review of the documentation and code, and on the challenger, rather than on independent execution.
