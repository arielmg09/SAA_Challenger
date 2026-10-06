# X. Independent Challenger Model

## X.1 Purpose

MRM developed an independent challenger model in Python to support the review of the Strategic Asset Allocation (SAA) model. It has two aims. The first is to make the model's calculations transparent by rebuilding its main steps in code that can be inspected line by line. The second is to provide an independent reference point for sense-checking the developer's outputs.

The challenger is a validation tool only. It does not replace the production model and is not intended for setting the SAA.

## X.2 What the SAA model does

### X.2.1 How the model is built

The SAA model is a set of linked Excel workbooks. Market and economic data are sourced from FactSet, so running the model depends on access to FactSet data services. The workbooks are as follows:

- **Composite Weights.xlsx** refreshes, once a year, the weight of each index within its asset class composite. For Equities, Developed Market Government Bonds and Other Bonds, it updates the market capitalisation of each index and calculates the weights from these values.
- **Forward Looking Return Assumptions.xlsx** applies the Building Block Approach (BBA) to estimate an expected return for each index.
- **SAA Input Data.xlsm** is the main store of inputs. It extends the index data to the latest period and combines the BBA and Historic Volatility Approach (HVA) into a single expected return. It also holds:
  - the volatilities, calculated from monthly historical returns as `STDEV.S()×√12`;
  - the correlations, calculated with `CORREL()`;
  - the yields.
- **SAA Model.xlsm** runs the optimisation and the simulations. These inputs are pasted into its MonteCarloModel tab as values.

The SAA Model is run separately for each of ten strategies: Defensive, Cautious, Balanced, Growth, Income, Income Plus, and the Responsible versions of Defensive, Cautious, Balanced and Growth. Each strategy has its own volatility target. The targets are evenly spaced between 30% and 90% of the volatility of equities. Each portfolio starts with a value of £2. The Core strategies are optimised first, because some constraints for the Income and Responsible strategies are based on the Core results.

The model produces three blocks of output for each strategy.

### X.2.2 Block A – Optimisation

**Main output: the weight of each asset class composite for each strategy, and the same weights broken down to the 19 underlying indices.**

The optimisation uses Palisade @RISK 7.6 and its RISKOptimizer add-in, and works at the level of seven asset class composites: Equity, Developed Market Government Bonds, Other Bonds, Property, Gold, Hedge Funds and Cash. Inflation is also simulated but cannot be invested in.

The weight of each composite is set as a whole number between 0 and 1,000, then rescaled so that the weights add up to 100%. For each candidate portfolio, @RISK simulates 20 years of correlated annual returns. The goal is to maximise the average portfolio value in year 20 (cell W97).

Constraints are "soft": a portfolio may break a constraint, but its score is reduced by a penalty. The penalty is (1,000 × size of breach)² for the volatility target and the minimum yield, and (100 × size of breach)² for the minimum and maximum weight of each asset class. The optimisation runs twice per strategy, first with 500 simulation iterations and then with 1,000.

Once the composite weights are fixed, each one is multiplied by the weights of the indices within that composite. This gives the final allocation to the 19 indices.

### X.2.3 Block B – Historic Characteristics

**Main output: a set of statistics showing how each strategy would have performed over the last 20 years.**

Each optimised portfolio is tested on the last 20 years of monthly historical data. The statistics reported are:

- the annualised return, and the annualised return in excess of cash and of inflation;
- the highest and lowest 12-month returns;
- the annualised volatility;
- the largest fall from a peak (maximum drawdown), with the dates of the peak and the low point, and the number of months taken to recover;
- the average gross portfolio yield.

These figures are calculated in a hidden HistoricSimulation tab and shown on the Summary tab.

### X.2.4 Block C – Projected Characteristics

**Main output: a set of statistics showing the range of outcomes each strategy might produce over the next 20 years.**

Each optimised portfolio is put through a Monte Carlo simulation of 10,000 possible 20-year futures. The statistics reported are:

- the average annualised return;
- the annualised return achieved with 90% and with 10% probability;
- the average annualised return in excess of cash and of inflation;
- the average volatility;
- the average of the lowest 12-month return;
- the maximum drawdown, as an average and at a 10% probability;
- the current gross portfolio yield.

## X.3 The challenger: Replica and Analytical

The challenger reproduces all three blocks. For Block A, the optimisation, it uses two independent methods.

**The Replica** rebuilds the developer's method as closely as the available information allows. It uses:

- the same goal (the highest average portfolio value in year 20);
- the same starting value of £2;
- the same whole-number weights from 0 to 1,000;
- the same penalty formulas;
- the same two runs of 500 and 1,000 iterations;
- the same rank-based correlation method.

The main difference is the search method. Instead of RISKOptimizer's genetic algorithm, the Replica uses differential evolution, a similar trial-and-error method from the open-source SciPy library.

**The Analytical** method is an addition by UKRM with no counterpart in the developer's model. It skips the simulation and calculates directly the portfolio with the highest expected return that meets every constraint exactly.

This is a fair benchmark for a simple reason. The portfolio is rebalanced every year and each year's returns are drawn independently. Under these conditions, the average year-20 value equals the starting value multiplied by (1 + expected portfolio return) to the power of 20. Maximising the average year-20 value is therefore the same as maximising expected return. The Analytical result is the answer the Replica is aiming for, without the effect of simulation noise, penalties or rounding.

The Replica, Analytical and developer weights are all run through the same historic back-test (Block B) and the same set of 10,000 simulated futures (Block C). Any differences in the results therefore come from the weights alone.

## X.4 Differences between the challenger and the developer's model

| Aspect | Developer | Challenger |
|---|---|---|
| Software | Linked Excel workbooks with @RISK 7.6 and RISKOptimizer | Python (NumPy, SciPy, pandas) in a Jupyter notebook |
| Inputs | Calculated in the upstream workbooks from FactSet data | Taken as given from the developer's workbooks and entered into an input template |
| Optimisation search | Genetic algorithm | Replica: differential evolution. Analytical: direct calculation (SLSQP) from 20 starting points |
| Constraints | Soft, with penalties | Replica: same penalties. Analytical: must be met exactly |
| Stopping rule | 1,000 trials, 60 minutes, or improvement below 0.001% over 100 trials | Fixed number of search rounds (about 10,500 trials per run) |
| Distribution of returns | Not visible in the values-only file | Normal |
| Historic back-test | Hidden HistoricSimulation tab; method not visible | Monthly data; weights reset to target every month (yearly is an option) |
| Monte Carlo | 10,000 iterations, annual steps (inferred from the model layout) | 10,000 iterations, annual steps |

Because the search methods differ, the number of trials in the challenger is not directly comparable with RISKOptimizer's.

## X.5 Key assumptions

The challenger relies on the following assumptions. Where an assumption could not be confirmed from the information available, this is stated.

1. **Optimisation goal.** Cell W97 is understood to be the average portfolio value in year 20, including capital and income. This is based on the model's layout and has not been confirmed from the formulas.
2. **Returns.** Expected returns are treated as average annual returns. Each year's returns are independent of the previous year's. Portfolios are rebalanced to their target weights every year, with income reinvested.
3. **Distribution of returns.** Returns are assumed to be normally distributed. The vendor describes the method as "distribution-free", but this refers to how correlations are applied. @RISK still assigns a distribution to each input, and this cannot be seen in a values-only file. The choice has little effect on the optimal weights, but it does affect the spread of outcomes in Block C.
4. **Volatility targets.** Targets are set at evenly spaced points between 30% and 90% of equity volatility, using four risk levels. The level assigned to the Income, Income Plus and Responsible strategies is a placeholder and needs to be confirmed with the model owner.
5. **Portfolio volatility in the optimisation.** The Replica measures portfolio volatility from the simulated annual returns, and the Analytical method calculates it from the asset class volatilities and correlations. The developer's exact measure has not been confirmed.
6. **Penalty scale.** The size of a breach is measured in decimal terms (a 1% breach counts as 0.01). With a starting value of £2, the penalties are large compared with the year-20 portfolio value, so in practice the constraints behave almost as fixed limits.
7. **Correlations.** The correlation matrix is treated as rank correlations, as @RISK does, even though it is estimated with `CORREL()`. The challenger keeps this approach so that its results remain comparable with the developer's. For normally distributed returns, the difference between the two types of correlation is no more than about 0.02.
8. **Constraints for the Income and Responsible strategies.** The developer bases some of these constraints on the Core results. In the challenger they are entered directly into the input template.
9. **Historic Characteristics.**
   - The back-test uses monthly composite returns, with the weights reset to target each month.
   - Excess returns are the annualised portfolio return minus the annualised return on cash or inflation.
   - Months to recovery are counted from the low point of the largest fall.
   - The average yield is the weighted average of the monthly composite yields.
10. **Projected Characteristics.**
    - Returns are simulated in annual steps, so the "lowest 12-month return" is the lowest single-year return in each simulated future.
    - Drawdowns are measured on year-end values, so they will be smaller than drawdowns measured on monthly data.
    - "Return with 90% probability" is read as the return beaten in 90% of the simulated futures (the 10th percentile).

## X.6 How the challenger works: step by step

The challenger consists of a Jupyter notebook, which runs the analysis in six sections, and a supporting file of helper functions (`saa_helpers.py`). While it runs, the notebook prints a line for every step, such as "BLOCK A | REPLICA | Balanced | Pass 1", so it is always clear what is being calculated.

**Section 1: Setup.** The user sets:

- the input and output file names;
- the strategies to run, in order (Core first);
- whether volatility targets are calculated automatically;
- how long the optimiser searches;
- how often the back-test rebalances.

The developer's composite weights can also be entered here, so that they appear alongside the challenger's results.

**Section 2: Input data reading.** The notebook reads the input template and displays its contents. It then carries out three checks:

- that the correlation matrix is mathematically valid (a Cholesky decomposition, the same check @RISK uses);
- that the index weights within each composite add up to 100%;
- whether historic data is available.

**Section 3: Block A, optimisation.** This section has five steps.

1. *Volatility targets.* The target for each strategy is calculated from equity volatility, or taken from the Setup tab.
2. *Replica.* For each strategy:
   - the model simulates 20 years of annual returns for each asset class, based on its expected return and volatility;
   - it reorders these values so that the asset classes move together according to the target correlations (the Iman–Conover method, as used by @RISK);
   - it scores each candidate portfolio as its average year-20 value minus any penalties;
   - it searches for the highest-scoring portfolio, first using 500 simulations and then using 1,000 new simulations, starting from the first answer.
3. *Analytical.* The model calculates directly the highest-return portfolio that meets all constraints exactly.
4. *Composite output.* The weights are shown in one table per method, with strategies in rows and composites in columns, together with the differences between methods and a chart.
5. *Unbundling.* Each composite weight is multiplied by the weights of the indices within it, giving the allocation to the 19 indices.

**Section 4: Block B, Historic Characteristics.** Each set of weights is tested on the last 20 years of monthly composite returns. The notebook then calculates the historic statistics listed in X.2.3 for each strategy.

**Section 5: Block C, Projected Characteristics.** The model simulates 10,000 possible 20-year futures, applies every set of weights to the same futures, and calculates the projected statistics listed in X.2.4. A chart shows each composite and each optimised portfolio by risk and return.

**Section 6: Output creation.** All results are saved to one Excel file, with separate sheets for composite weights, index weights, Historic Characteristics and Projected Characteristics for each method.

## X.7 Limitations

The challenger rebuilds the developer's method in concept, not line by line. Because the search methods differ, it cannot be expected to reproduce RISKOptimizer's weights exactly. Several features of the developer's model could not be confirmed from the values-only file:

- the distribution of returns;
- the exact volatility measure used in the optimisation;
- the method used in the hidden HistoricSimulation tab;
- how the Income and Responsible constraints are derived from the Core results.

These are covered by the assumptions in X.5. The challenger also relies on the template containing an accurate copy of the developer's inputs and historic data.

## X.8 Interpreting the comparison

The results are read in three steps.

**Replica against Analytical.** With a starting value of £2, the two methods are expected to agree closely. A noticeable gap in weights alongside an almost identical expected return means that quite different portfolios score almost equally well. In that situation the optimal weights are sensitive to small changes in inputs or to simulation noise.

**Developer against challenger, Block A.** If the developer's weights are close to the challenger's, this indicates that the optimisation reaches a sensible answer for the given inputs. A developer portfolio with a materially lower expected return at a similar volatility may indicate one of two things. Either the RISKOptimizer search stopped early, or the inputs, constraints or assumptions in the developer's model differ from those in the challenger.

**Developer against challenger, Blocks B and C.** When the same weights are entered, the challenger's statistics should be close to the developer's Summary tab. Differences then point to the calculation method, not the weights. The relevant assumptions in X.5 are the starting point for investigating them.

[Results and conclusions to be added once the challenger has been run on the current-year inputs and historic data.]
