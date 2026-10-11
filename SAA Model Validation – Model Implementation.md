# Model Implementation

Oct 9, 2026 · @Ariel Gimenez

## Implementation overview

The SAA model is implemented as a chain of linked Excel workbooks, supported by VBA macros and the Palisade @RISK add-in (including its RISKOptimizer module). It is run once a year by the Model Owner (MO), ahead of the annual re-optimisation of the strategies. Market and economic data are sourced through the FactSet Excel add-in, so a full run depends on FactSet access.

The workbooks sit on a shared network drive. Before each annual refresh, the prior year's materials are archived by copying them into a new folder within the archive area. Each workbook then covers one stage of the process and hands its results to the next, as summarised below.

| Workbook | Role in the process | Main output | How the output is passed on |
| --- | --- | --- | --- |
| Composite Weights | Refreshes index market capitalisations and market values, and derives the weight of each index within its composite | Index-to-composite weights | Values copied manually into a hidden mapping sheet in both SAA Input Data and SAA Output |
| Forward Looking Return Assumptions | Implements the Building Block Approach (BBA) for each index; FactSet refresh plus manual updates to the default parameters and real GDP growth tabs | BBA expected returns by index | Used by SAA Input Data in its data update; no separate transfer step is described |
| SAA Input Data | Main input repository: refreshes index histories, builds the composite histories and calculates composite volatility, correlation, yield and blended BBA/HVA returns, through a macro-driven data update | Composite-level return, volatility, yield and correlation inputs, including inflation | Values copied manually into SAA Model; composite histories also read directly by the historic simulation macro |
| SAA Model | Hosts the optimisation and the Monte Carlo simulation in @RISK, and the historic back-test of each strategy | Optimal composite weights and historic and projected characteristics by strategy | Written by an export macro to a sheet named after the strategy in SAA Output |
| SAA Output | Collates composite- and index-level weights and characteristics, the prior-year comparison, the change analysis and the materiality check | Final SAA and supporting governance material | End of the chain |

&#91;embedded content: SAA implementation workflow · workbooks and hand-offs\]

Five of the twelve hand-offs shown, marked with dashed lines, rely on manual copy-paste or re-entry; the prior-year starting weights entered before each optimisation are a further manual step not drawn.

One run produces ten strategies: four Core strategies (Defensive, Cautious, Balanced and Growth), two Income strategies (Income and Income Plus) and four Responsible variants of the Core strategies, referred to as Sustainable in the instructions document. The Core strategies are optimised first, because the yield and hedge fund constraints of the Income and Responsible strategies are set by reference to the optimised Core output.

The calculation of composite weights, BBA and HVA expected returns, volatilities and correlations is covered in earlier sections. The remainder of this section takes those inputs as given and focuses on how they are carried through the optimisation and simulation steps.

## Annual refresh and preparation of inputs

The annual run follows a step-by-step instructions document maintained by the MO. Up to the start of the optimisation, the process runs as follows.

1. **Archive.** The prior year's workbooks are copied into a new archive folder, so that the previous SAA can be retrieved for the prior-year comparison and the materiality check.
2. **Composite weights.** The Composite Weights workbook is refreshed from FactSet. The resulting within-composite index weights are then copied as values into a hidden mapping sheet in SAA Input Data and in SAA Output. These weights drive both the construction of the composite histories and the later breakdown of composite weights to index level.
3. **Forward-looking returns.** The Forward Looking Return Assumptions workbook is refreshed from FactSet. Two tabs, holding default parameters and real GDP growth, are updated manually following instructions embedded in the workbook.
4. **Input data.** In SAA Input Data, a macro refreshes the index data and recalculates the risk and return statistics. The macro asks the user to choose the data window through a dialog box. The instructions prescribe all available data for the general analysis, data limited by the shortest index series for the correlation matrix and for the return assumptions, and no adjustment to the asset class weighting used to combine the return assumptions. The workbook is saved and kept open for the next step.
5. **Transfer to the optimisation workbook.** The composite-level inputs are copied from SAA Input Data and pasted as values into the Monte Carlo sheet of SAA Model, as set out below.

| Input (composite level) | Source in SAA Input Data | Destination in SAA Model |
| --- | --- | --- |
| Expected returns (8 series, incl. inflation) | Return assumptions sheet | Expected return column of the Monte Carlo input table |
| Expected volatilities (8 series) | Volatility sheet | Volatility column of the same table |
| Expected yields (7 composites) | Yield sheet | Yield column of the same table |
| Correlation matrix (8 x 8) | Correlation sheet | Correlation block of the Monte Carlo sheet |

At this point, all inputs needed for the optimisation sit as static values in SAA Model. No live link remains to the upstream workbooks, so any later change upstream must be re-transferred by hand.

## Optimisation

### Set-up of the optimisation workbook

SAA Model is a macro-enabled workbook with two user-facing sheets: a Summary sheet and a Monte Carlo sheet, plus a sheet holding the Monte Carlo summary statistics and a hidden sheet holding the historic simulation. Two control cells on the Monte Carlo sheet drive each run. The first sets the mode (Optimisation or Simulation). The second selects the strategy, which loads that strategy's constraint set: minimum and maximum weights for each composite, the volatility target and, where relevant, the minimum yield.

The Monte Carlo sheet projects the value of a portfolio over a 20-year horizon from a notional starting value of 2. Composite returns are drawn by @RISK using the expected returns and volatilities transferred from SAA Input Data. The correlation matrix is imposed through @RISK's rank-order correlation method, which, according to the vendor documentation, works on Spearman rank correlations and is distribution-free. A designated output cell holds the projected portfolio value at the end of the horizon.

### Decision variables, starting point and objective

The optimiser adjusts seven cells, one per composite (inflation is a simulated variable, not an investable composite). Each cell takes an integer value between 0 and 1,000. Portfolio weights are derived from these integers so that they always sum to 100%; UKRM's replication indicates that each integer is divided by the sum of the seven.

Before each run, the user enters the strategy's prior-year SAA weights, multiplied by 1,000, as the starting values. The optimisation therefore starts from last year's allocation.

The objective is to maximise the mean, across simulation iterations, of the projected portfolio value at the end of the 20-year horizon. RISKOptimizer runs in standard mode (a single optimal portfolio per strategy), not in efficient-frontier mode.

### Constraints and penalties

All constraints are set as soft constraints. A candidate portfolio that breaches a constraint is not rejected; instead, a penalty is deducted from its objective value. The penalty grows with the square of the breach:

- (1,000 x deviation)^2 for the annualised volatility and minimum yield constraints;
- (100 x deviation)^2 for the minimum and maximum composite weights.

The heavier penalty means the optimiser tolerates breaches of the volatility and yield targets less than breaches of individual weight bounds. Volatility targets for the strategies are set at equally spaced points between 30% and 90% of expected equity volatility.

### Optimiser settings and run sequence

RISKOptimizer performs a stochastic, population-based search: it generates candidate portfolios (trials), runs a full @RISK simulation on each one and keeps the best penalised results to breed the next candidates. The instructions do not specify the search engine selected. Each optimisation stops at the first of three limits:

- 1,000 trials;
- 60 minutes of run time;
- an improvement of less than 0.001% in the objective over the last 100 trials.

Each strategy is optimised twice. The first pass uses 500 simulation iterations per trial, and the second pass uses 1,000. The instructions state that the second pass ensures a global maximum is found.

After the second pass, the resulting composite weights are copied and transposed by hand into the strategy results table, one row per strategy. The process is repeated for all ten strategies, the Core strategies first.

## Simulation, outputs and materiality

### Simulation of the optimised strategies

Once all ten strategies have been optimised, the mode cell is switched to Simulation. For each strategy in turn, the user selects the strategy, sets @RISK to 10,000 iterations and runs the Monte Carlo simulation on the optimised allocation; its summary statistics are collected on a dedicated results sheet. On the Summary sheet, the user then runs two macros: the historic simulation and the export, described below.

### Historic simulation

The historic simulation back-tests the strategy's current weights over the last 20 years of monthly data. The window is hard-coded in the macro, and a comment in the code notes it cannot exceed 22 years because of data availability. The macro runs in five steps.

1. **Data load.** It opens SAA Input Data, which must sit in the same folder as SAA Model, and reads the composite monthly returns, composite yields, index returns and composite mapping directly from that workbook. Unlike the Monte Carlo inputs, these data are not copied by hand. The macro checks that the first and last dates of the window are exactly 20 years apart.
2. **Missing data.** If a composite has no return at the start of the window, because one of its indices has a shorter history, the macro rebuilds that composite from its remaining indices. Their weights are rescaled to sum to 100%, the rebuilt series is rebalanced annually, and it fills the gap.
3. **Weights.** The strategy weights are read from the Return & Risk Assumptions table on the Monte Carlo sheet, and each composite is matched to its return series by name.
4. **Portfolio path.** The portfolio is held buy-and-hold within each calendar year and rebalanced to the target weights at each December month-end. The monthly portfolio yield is the value-weighted average of the composite yields.
5. **Statistics.** The macro calculates annualised return and volatility, excess return over cash and inflation, maximum and minimum 12-month and monthly returns, maximum drawdown with its peak, trough and recovery time, downside volatility, skewness, kurtosis, and the Sharpe, Sortino, Calmar and adjusted Sharpe ratios. These, together with the monthly cumulative return index, returns, drawdown profile and yield, are written to the hidden historic simulation sheet.

### Export to SAA Output

The export macro collects, for the selected strategy, the strategy name, the Return & Risk Assumptions table (weights, expected return, volatility and latest yield for the seven composites and inflation), the correlation matrix, the Monte Carlo summary statistics and the historic simulation results. It writes them to the sheet of SAA Output named after the strategy, overwriting its previous contents, then saves and closes the file. SAA Output must sit in the same folder as SAA Model, and a sheet for each strategy must already exist. Each strategy sheet therefore holds both the results and the inputs used to produce them.

Both macros check the layout of the sheets they read, such as the position of key labels, and stop if it is not as expected. Neither checks that the Monte Carlo and historic results on the sheets belong to the strategy being exported.

When all strategies have been simulated and exported, the workbook is saved as the "SAA Unconstrained Model".

For each strategy, the Summary sheet reports two sets of characteristics.

| Historic characteristics (20-year back-test) | Projected characteristics (Monte Carlo, 20-year horizon) |
| --- | --- |
| Annualised return | Mean annualised return, and the return at 90% and 10% probability |
| Excess return over cash and over inflation | Mean excess return over cash and over inflation |
| Annualised volatility | Mean volatility |
| Maximum and minimum 12-month return | Mean minimum 12-month return |
| Maximum drawdown, peak and trough dates, months to recovery | Maximum drawdown (mean and at 10% probability) |
| Average gross yield | Current gross yield |

### Output workbook and prior-year comparison

SAA Output collates the exported results. Its Summary tab shows weights at composite level and, using the within-composite weights in its hidden mapping sheet, at individual index level, together with the historic and projected characteristics. The MO reviews and sense-checks the output; the instructions do not define what this review covers.

The prior-year Summary tab is then copied as values from the archived output file into a Prior Year tab. Because some headings use merged cells, this may need to be done one table at a time. A Change tab presents tables and charts of the year-on-year changes for the wider investment team.

### Materiality check and final run

The final step decides, strategy by strategy, whether the new allocation is different enough from last year's to be adopted. The Materiality tab offers two approaches:

- **Absolute sum of changes.** This updates automatically and compares the total absolute change in weights against a threshold. The threshold is currently the ex-ante standard deviation discussed in the 2022 SAA review document, and the instructions state that it should be re-evaluated each year.
- **Tracking error.** This measures the ex-ante tracking error of the new allocation against the prior-year allocation. It needs a matrix copied by hand from SAA Model into the Materiality tab, which the instructions describe as the covariance matrix.

The output workbook is saved as the "SAA Unconstrained Output". For strategies whose change does not meet the materiality threshold, the user reopens the SAA Unconstrained Model and overwrites the optimised allocation with the prior-year allocation, taken from the archived files. The simulation step is then re-run for all strategies, and the workbook is saved as the "SAA Model". The resulting SAA Output is the final output used by the business.

## UKRM Independent Assessment

UKRM reviewed the implementation against the MO's instructions document, the structure of the workbooks and macros, and the results of UKRM's independent challenger implementation. The review identified seven findings for the MO to address, set out below. Further points that do not require remediation are recorded under Observations.

### Finding 1: Reliance on manual copy-and-paste between workbooks

**Rating:** \[to be assigned\]

The implementation depends on repeated manual transfers of values between workbooks. A single annual run involves at least the following manual interventions:

- within-composite weights pasted into hidden sheets in two separate workbooks;
- manual updates to the default parameters and real GDP growth tabs;
- four blocks of composite inputs (returns, volatilities, yields and correlations) pasted into SAA Model;
- prior-year weights, scaled by 1,000, typed in as starting values for each of the ten strategies;
- optimised weights copied and transposed into the results table for each of the ten strategies;
- the prior-year Summary pasted table by table into the Prior Year tab;
- the matrix for the tracking error approach pasted into the Materiality tab;
- prior-year allocations re-entered for strategies that fail the materiality threshold.

No reconciliation control or maker-checker review of these transfers is described or evidenced. Several destinations are hidden sheets, which makes a misplaced or partial paste harder to spot. Because values replace live links, the data lineage from FactSet to the final allocation is broken at each step. The historic simulation, by contrast, reads its data directly from SAA Input Data and matches series by name, which shows that automation is feasible within the current set-up. It also means the back-test and the Monte Carlo projection reach the same workbook by different routes: if SAA Input Data changes after the manual paste, the two sets of results rest on different inputs.

An error at any of these points, such as a range shifted by one row or a stale value left from the prior year, would pass silently into customer-facing strategies. The number and repetition of the steps make such an error more likely over time.

**Recommendation.** The MO should automate the transfer of data between workbooks, for example through direct links or a macro-driven import that matches series by name and checks dimensions and totals before loading. Until that is in place, each manual transfer should be subject to a documented maker-checker reconciliation, retained as evidence for each annual run.

### Finding 2: Reproducibility, convergence and constraint compliance of the optimisation

**Rating:** \[to be assigned\]

RISKOptimizer is a stochastic search, and each candidate portfolio is itself evaluated on a finite Monte Carlo sample. Two runs with the same inputs can therefore return different allocations. The instructions do not set a random seed, do not record the optimiser settings used for each strategy, and do not require evidence of convergence. The claim that the two-pass approach ensures a global maximum is not supported by any testing.

The soft-constraint design adds a second source of uncertainty. A portfolio that slightly breaches a constraint can score higher than a fully compliant one if the extra return outweighs the penalty, and no post-optimisation check confirms that the final allocation meets every constraint.

UKRM's challenger results are consistent with these concerns. When the RISKOptimizer set-up is replicated (penalised stochastic search), the resulting allocations breach the volatility target by up to 0.16 percentage points and some maximum-weight bounds by small amounts. Compared with an exact solution of the same problem, composite weights differ by up to around 3 percentage points. The allocations found in the first and second passes also differ materially for some strategies. Further detail is given in the challenger section of this report.

Under the model's assumptions of independent returns and a fixed-weight portfolio, the objective (the mean terminal portfolio value) depends only on the portfolio's expected return. The optimisation problem can therefore be solved exactly with a deterministic solver, without Monte Carlo noise.

**Recommendation.** The MO should fix and record the random seed and all optimiser settings for each run, and evidence convergence, for example by repeating the optimisation from several starting points and reporting the spread of results. A hard check that every constraint is met should be added after optimisation. The MO should also benchmark the RISKOptimizer output against a deterministic solution of the same problem and explain any differences.

### Finding 3: Implementation of the materiality check

**Rating:** \[to be assigned\]

The materiality check determines whether each strategy adopts the new allocation or keeps the prior-year one, so it directly drives the final SAA. Three weaknesses were identified.

- The instructions describe the matrix pasted into the Materiality tab as a covariance matrix, but the range referenced in SAA Model holds the correlation matrix. Unless the Materiality tab rescales it by the volatilities, which UKRM has not been able to confirm, the tracking error is misstated.
- The instructions present two approaches but do not state which one decides the outcome, or what happens if they disagree.
- The threshold for the absolute sum of changes is described as subject to annual re-evaluation, but no evidence of that re-evaluation or of its approval was provided.

**Recommendation.** The MO should confirm and, if needed, correct the matrix used in the tracking error calculation. The decision rule, the threshold and its annual review should be documented and approved, and the outcome for each strategy recorded.

### Finding 4: Instructions document not aligned with the model in use

**Rating:** \[to be assigned\]

The instructions document is the only description of the implementation, but it does not match the model currently in use:

- the cell ranges for the composite weights are sized for 21 indices, while the current model uses 19;
- the strategy results table is referenced with two different ranges in different steps;
- the worked example shows starting values that do not match the weights displayed next to them;
- the matrix used for the tracking error is labelled as a covariance matrix (see Finding 3);
- the optimiser search engine and random seed are not specified, and the review and sense-check steps carry no criteria;
- the SAA Model version made available to UKRM differs from the version described.

Out-of-date instructions increase the risk of errors during the annual run, particularly if it is performed by someone other than the usual operator.

**Recommendation.** The MO should update the instructions to reflect the current model, place them under version control alongside the workbooks, and review them as part of each annual run.

### Finding 5: Process controls, audit trail and key-person dependency

**Rating:** \[to be assigned\]

The annual run is a long, sequential, manual process with no run log. Mode and strategy selection, iteration counts and file naming are all set by hand, and the version history of the model relies on folder copies in the archive area. No record is kept of who ran each step, which settings were used, or what checks were performed.

The process also appears to depend on the knowledge of a small number of individuals, and on access to FactSet and to a licensed installation of @RISK.

**Recommendation.** The MO should introduce a run checklist with sign-off at each stage, retaining the settings and outputs of each run. The workbooks should be brought under the Bank's end-user computing controls, including access restrictions, change control and protection of formulas and hidden sheets. At least one additional person should be trained to perform the annual run.

### Finding 6: Error handling and run-state checks in the macros

**Rating:** \[to be assigned\]

The macros in SAA Model are structured as a chain of steps that each report success or failure, but a failure is not handled safely:

- when a step fails, the macro stops without any error message, leaving Excel with calculation set to manual and screen updating switched off; the only sign of failure is the absence of the completion message;
- after a failed historic simulation, the hidden sheet still holds the previous strategy's results, and the export macro would write them to SAA Output under the current strategy's name, since neither macro checks that the Monte Carlo and historic results match the strategy selected;
- file names, the same-folder location of SAA Input Data and SAA Output, the number of composites (seven) and the 20-year window are hard-coded;
- the missing-data routine records only the last index found with missing data, so a composite with more than one short index would not be rebuilt correctly.

These weaknesses make it possible for incorrect or mismatched results to reach SAA Output without the user being aware.

**Recommendation.** The MO should add explicit error messages and restore Excel's settings when a macro fails. Before export, the macros should confirm that the Monte Carlo and historic results were produced for the selected strategy in the current run, for example by stamping each results sheet with the strategy name and run time. Hard-coded parameters should be moved to a visible set-up area, and the missing-data routine should either handle several short indices or stop with a clear message.

### Finding 7: Errors and non-standard definitions in the historic statistics

**Rating:** \[to be assigned\]

Review of the historic simulation code identified the following issues in the statistics it calculates:

- **Adjusted Sharpe ratio.** The standard adjustment (Pézier and White) is SR x \[1 + (S/6) x SR − (K/24) x SR^2\], where S is skewness and K is excess kurtosis. The code uses skewness in the kurtosis term, so the reported ratio ignores kurtosis.
- **Maximum and minimum returns.** The running maximum and minimum of the 12-month and monthly returns start at zero. The reported minimum can therefore never be positive and the reported maximum never negative. Over a 20-year window that includes market falls the practical impact is likely to be small, but the logic is incorrect.
- **Downside volatility.** This is calculated as the standard deviation of the negative monthly returns around their own mean, annualised. The usual definition measures deviations below zero or a target return; the code's measure understates downside risk and therefore overstates the Sortino ratio.

The export macro copies the whole historic simulation sheet into SAA Output, so these statistics are carried into the output workbook. UKRM has not confirmed which of them are presented to stakeholders.

**Recommendation.** The MO should correct the adjusted Sharpe ratio and the maximum and minimum calculations, document the definition of downside volatility and align it with standard practice, and test the historic statistics against an independent calculation.

## Observations

The following points do not require remediation by the MO but are recorded for completeness and for consideration at the next model change.

- **Correlation measure.** The input correlations are Pearson correlations of historic composite returns, while @RISK imposes correlations as Spearman rank correlations. For approximately normal returns the difference is small, but the two measures are not the same, and the instructions do not document the marginal distributions assumed for each composite. A short note in the model documentation would close this gap.
- **Encoding of the weights.** Because weights are derived by dividing each integer by the total, many different integer combinations give the same portfolio, and the effective step size varies from run to run. This makes the search less efficient but is not expected to affect the result materially.
- **Notional starting value.** The projection starts from a portfolio value of 2, which the historic simulation scales to 2 million. This has no effect on the reported returns, volatilities or drawdowns, but it is not intuitive for a reader of the intermediate outputs.
- **Order of optimisation.** The Income and Responsible strategies depend on the Core results. If a Core strategy is re-run, its dependent strategies must be re-run too; the instructions do not say so.
- **Run time.** At the stated limits (two passes of up to 60 minutes each for ten strategies), the optimisation could take up to 20 hours in the worst case. A long run time creates an incentive to stop early or skip the second pass.
- **Software version.** The instructions refer to @RISK 7.6. The MO should confirm that the installed release remains supported by the vendor.
- **Data window settings.** The data window chosen in the SAA Input Data dialog (all available data for the general analysis) should be checked for consistency with the methodology documentation, which describes volatilities as limited by the shortest index series. This point is covered further in the data and inputs sections.

* **Rebalancing assumption.** The historic simulation rebalances to target weights at each calendar year-end. The MO should confirm that the Monte Carlo projection uses the same rebalancing assumption, so that historic and projected characteristics are comparable.
* **Regional settings.** The macros convert numbers and dates to text and back, and format dates as dd/mm/yyyy. Running them on a machine with different regional settings could cause dates to be misread.
* **Shared code library.** The shared utility module is a general-purpose library that includes routines not used by the SAA process, such as SEDOL formatting, sorting and frequency conversion. The SEDOL routine runs on every sheet the macros write, with no effect on the SAA sheets. Removing unused code would make the macros easier to review and maintain.
