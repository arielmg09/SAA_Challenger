## Model Implementation

The SAA model is implemented as a chain of linked Excel workbooks, run once a year by the model owner following the documented SAA Instructions. Market and economic data are taken from FactSet, so the model can only be run where FactSet access is available. The workbooks are run in a fixed order. Data passes between them mostly as manual copy-and-paste of values, not through live links. Before each annual run, the prior year's materials are copied to an archive folder, which keeps them available for the year-on-year comparison and the materiality step described below.

The first stage prepares the inputs in two workbooks, both refreshed from FactSet. The `Composite Weights` workbook sets the weight of each index within its composite. For the Equity, Developed Market Government Bonds and Other Bonds composites, it updates the market capitalisation or market value of the underlying indices and derives the weights from these, following the home-bias and averaging rules in the developer documentation. The resulting weights (sheet 'Summary', C2:C22) are pasted as values into the hidden 'AssetClassMap' sheet of both the `SAA Input Data` and `SAA Output` workbooks. The `Forward Looking Return Assumptions` workbook implements the Building Block Approach (BBA) and produces an expected return for each index. Some of its inputs, on the 'Defaults' and 'RealGBPGro' tabs, are updated by hand rather than through the FactSet refresh.

The `SAA Input Data` workbook is the central store of model inputs. Running the 'Update Data' macro on the 'RiskReturnStats' tab refreshes the index histories. The workbook then calculates expected volatilities, correlations and yields at composite level, together with the blended BBA/Historic Volatility Approach (HVA) expected returns. Volatilities are calculated as the annualised standard deviation of historical monthly returns, and correlations between composites as the Pearson correlation of the same series. The macro settings specified in the instructions are "all available data" for the volatility analysis, and data "limited by the shortest series" for both the correlation matrix and the return assumptions. The outputs are held on four tabs: 'ReturnAssumptions', 'VolatilityMatrix', 'YieldData' and 'CorrelationMatrix'. In the version of the model currently in use, 19 indices are grouped into seven composites. The developer documentation, which describes 21 indices, has not been updated to reflect this.

The four input blocks are pasted as values into the 'MonteCarloModel' sheet of the `SAA Model` workbook. Optimisation and simulation are run in this workbook with Palisade @RISK and RISKOptimizer. For each strategy, the optimiser adjusts the integer composite allocations (C48:C54, expressed per 1,000) to maximise the mean of the simulated portfolio value at cell W97. This maximisation is subject to soft constraints on minimum and maximum composite weights, annualised volatility and, where relevant, minimum yield, each enforced through a quadratic penalty function. Each strategy is optimised twice, first at 500 and then at 1,000 iterations, with the aim of reaching a global rather than a local maximum. The Core strategies are optimised first, because the yield and hedge fund constraints of the Income and Responsible strategies are derived from the optimised Core results. Once all strategies have been optimised, the workbook is switched to simulation mode. A 10,000-iteration Monte Carlo simulation is then run for each strategy, the 'Update Historic Simulation' and 'Export Output' routines are executed on the 'Summary' sheet, and the file is saved as the "SAA Unconstrained Model".

The exported results are reviewed in the `SAA Output` workbook. The prior year's Summary is pasted from the archive into the 'Prior Year' tab, and the 'Change' tab compares current and prior allocations. The 'Materiality' tab applies two tests to the proposed changes: the absolute sum of weight changes, and a tracking-error measure based on the matrix taken from `SAA Model` (B57:J66). Any strategy whose proposed change does not meet the materiality threshold is reverted to its prior-year allocation in the unconstrained model, and the simulation stage is re-run. This produces the final `SAA Model` and `SAA Output` files.

## Model Outputs

The model produces results for ten strategies: Defensive, Cautious, Balanced, Growth, Income, Income Plus, and the Responsible variants of Defensive, Cautious, Balanced and Growth. The volatility targets of the strategies are spaced evenly between 30% and 90% of the volatility of the Equity composite. For each strategy, the model produces four sets of outputs.

The first output is the strategic asset allocation itself: the optimised weight assigned to each of the seven composites (Equity, Developed Market Government Bonds, Other Bonds, Property, Gold, Hedge Funds and Cash). Combining these weights with the composite weights in the 'AssetClassMap' sheet gives the implied allocation to each underlying index.

The second output is a set of historic characteristics, calculated by applying each strategy's allocation to 20 years of historical index returns (the 'HistoricSimulation' sheet). These are:

| Metric | Description |
|---|---|
| Annualised return | Annualised total return over the 20-year history |
| Excess return over cash / inflation | Annualised return in excess of the cash and inflation series |
| Maximum / minimum 12-month return | Best and worst rolling 12-month returns |
| Annualised volatility | Annualised standard deviation of monthly returns |
| Maximum drawdown | Largest peak-to-trough decline, with the peak and trough dates and the months taken to recover |
| Average gross yield | Average historical portfolio yield |

The third output is a set of projected characteristics, derived from the 10,000-iteration Monte Carlo simulation in the 'MonteCarloModel' sheet. These are:

| Metric | Description |
|---|---|
| Annualised return | Mean outcome, and the outcomes at the 90% and 10% probability levels |
| Excess return over cash / inflation | Mean projected excess return |
| Volatility | Mean projected annualised volatility |
| Minimum 12-month return | Mean of the worst 12-month return across simulations |
| Maximum drawdown | Mean outcome and the outcome at the 10% probability level |
| Current gross yield | Portfolio yield based on the current yield inputs |

The fourth output is the year-on-year change and materiality assessment in the `SAA Output` workbook. This shows the change in each strategy's allocation from the prior year, together with the absolute-sum-of-changes and tracking-error results that determine whether the proposed allocation is adopted or the prior-year allocation is kept.

---


**X.X Historic Simulation and Forward Projections**

**X.X.1 Overview**

After optimisation, each strategy's asset allocation is assessed in two complementary ways. A historic simulation back-tests the optimised weights against realised market data. A forward projection, a Monte Carlo simulation, characterises the distribution of future outcomes implied by the forward-looking capital market assumptions. Both are executed within the 'SAA Model.xlsm' workbook for one strategy at a time and repeated for each of the ten strategies.

The optimisation step (RISKOptimizer) is treated as a black box for the purposes of this section. Its outputs, the composite weights for each strategy, are the inputs to the processes described below.

**X.X.2 Execution**

Following optimisation of all strategies, the model owner sets the run mode in cell C3 of the 'MonteCarloModel' sheet to 'Simulation', selects the strategy in cell C5 and runs @Risk with 10,000 iterations and a single simulation. The historic simulation and export are then triggered from the 'Summary' sheet using the 'Update Historic Simulation' and 'Export Output' buttons. The sequence is repeated for each strategy, and the workbook is saved as 'SAA Unconstrained Model'.

**X.X.3 Forward projections (Monte Carlo simulation)**

The forward projection is performed in @Risk using worksheet formulas on the 'MonteCarloModel' sheet. The simulation uses the following inputs, transferred manually from 'SAA Input Data.xlsm':

- expected returns (E48:E55)
- expected volatilities (F48:F55)
- expected yields (G48:G55)
- the composite correlation matrix (C70:J77)

These cover the seven composites plus inflation. Each iteration projects portfolio values over a 20-year horizon from a notional starting value of 2. This value is a convention adopted by the developers and does not represent a realistic portfolio size. Because all reported characteristics are expressed in return terms, they are not affected by the starting value. The terminal portfolio value at year 20 (cell W97) appears to be the value the optimiser maximises.

According to the vendor documentation, @Risk applies Spearman rank correlations through a distribution-free rank-order pairing method. The marginal distributions assumed for each composite were not documented by the model owner and have not been independently confirmed. The simulation engine is therefore treated as a vendor component, and its outputs are reviewed rather than its internal mechanics.

**X.X.4 Historic simulation (back-test)**

The historic simulation is implemented in VBA (module 'ModHistoricSimulation'). It applies the strategy's optimised composite weights to realised monthly composite returns, sourced from 'SAA Input Data.xlsm', over a fixed 20-year window. The procedure runs as follows:

1. **Data preparation.** The macro loads composite returns, sub-index returns, composite yields and composite weights. It truncates the series to 240 monthly observations and checks that the dates span exactly 20 years.
2. **Proxying of short histories.** Where a composite lacks returns at the start of the window, the macro rebuilds it from the constituent sub-indices that have sufficient history. The available weights are re-scaled pro rata, allowed to drift monthly, and reset to target each December.
3. **Portfolio simulation.** The notional starting value from the 'MonteCarloModel' sheet is allocated to composites according to the strategy weights. Each holding compounds at its realised monthly return, with weights drifting intra-year and rebalanced to target at each December year-end. Portfolio yield is calculated monthly as the value-weighted average of composite yields.
4. **Statistics.** Return, risk, drawdown and distributional statistics are calculated from the monthly portfolio value series and written to the 'HistoricSimulation' sheet.

The historic simulation uses realised returns only. It is independent of the forward-looking expected returns and therefore provides an out-of-model view of how the allocation would have performed historically.

**X.X.5 Export and final outputs**

The 'Export Output' macro (module 'ModExport') writes three blocks to the strategy-specific sheet of 'SAA Output.xlsm', overwriting the previous contents:

- the set-up parameters (composite weights, expected returns, volatilities, yields and correlation matrix)
- the Monte Carlo results
- the historic simulation results

These populate the 'Summary' tab of 'SAA Output.xlsm', which forms the final model output. For each strategy, the 'Summary' tab reports the allocation at both composite and individual index level, together with the following characteristics.

| Output | Source | Characteristics reported |
|---|---|---|
| Strategic asset allocation | RISKOptimizer, with within-composite weights from 'Composite Weights.xlsx' | Weight per composite and weight per individual index |
| Historic Characteristics | Historic simulation (20 years, realised data) | Annualised return; annualised excess over cash; annualised excess over inflation; maximum and minimum 12-month return; annualised volatility; maximum drawdown with peak and trough dates; months to recover; average gross yield |
| Projected Characteristics | Monte Carlo simulation (10,000 iterations, 20-year horizon) | Mean annualised return, with the 90% and 10% probability outcomes; mean excess over cash and over inflation; mean volatility; mean minimum 12-month return; maximum drawdown (mean and 10% probability); current gross yield |

The 'HistoricSimulation' sheet also calculates downside volatility, skewness, kurtosis and the Sharpe, Sortino, Calmar and adjusted Sharpe ratios. These are not reported on the 'Summary' tab.

The 'Change' tab compares the current-year allocations with the prior year. The 'Materiality' tab tests whether proposed changes exceed the materiality threshold, using an absolute sum of changes and a tracking-error approach. For strategies that do not breach the threshold, the prior-year allocation is reinstated and the simulation stage is re-run. That re-run produces the final 'SAA Model' and 'SAA Output' files, which constitute the approved strategic asset allocation and its reported characteristics for the year.
