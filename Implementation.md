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

There are three points to check before this goes in the report. First, I've described the B57:J66 block as "the matrix taken from `SAA Model`" rather than as a covariance matrix. You've confirmed it holds correlations, while the instructions call it a covariance matrix, so whether the tracking-error test scales it correctly belongs in your findings rather than in this descriptive section. Second, the step from composite weights to index-level allocation through 'AssetClassMap' is my reading of why the weights are pasted into `SAA Output`, so please confirm it with the model owner. Third, the macro settings and file references come from the instructions document, which you've noted differs from the model version you hold, so it's worth checking them against the current build.
