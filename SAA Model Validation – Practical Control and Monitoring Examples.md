# Practical Control and Monitoring Examples

Oct 16, 2026 · @Ariel Gimenez

This note gives a practical example for each validation finding: a simple control or monitoring step that the Model Owner could add to the existing process with little effort. The examples are deliberately concrete, mostly a new tab in an existing Excel workbook or a short one-page document, so that the team can see exactly what is being suggested. They are starting points to adapt, not fixed requirements, and MRM is happy to help set them up.

## Finding 1: Monitoring – a written monitoring statement

The model is already reviewed each year when the optimal weights are refreshed. The suggestion is to write that review down as a short monitoring statement, so that everyone knows what is checked, how, by whom, and what happens if something looks wrong.

### Example monitoring statement

| Element | What to write down (example) |
| --- | --- |
| Frequency | Once a year, as part of the annual refresh; plus an off-cycle review if a trigger below is hit |
| Outputs reviewed | Optimal weights for all ten strategies, projected and historic characteristics, and the change from last year |
| Type of review | Quantitative: the metrics below, each with a threshold. Qualitative: a short written comment on whether the results make sense given market conditions |
| Who prepares | The analyst who runs the model |
| Who reviews | A second member of the team who did not run the model |
| Who decides | The investment committee that approves the SAA |
| Record kept | A one-page monitoring summary saved with the year's files and presented to the committee |

### Example metrics and triggers

| Metric | How it is measured | Check if (amber) | Escalate if (red) | Action |
| --- | --- | --- | --- | --- |
| Realised vs expected return | Each strategy's actual return over the last 3 years, compared with the range the model projected | Outside the 25%–75% range | Outside the 10%–90% range | Explain in the summary; red triggers a review of the return assumptions |
| Realised vs target volatility | Actual volatility of each strategy over the last 3 years vs its volatility target | More than 1 percentage point from target | More than 2 percentage points from target | Check whether the strategy still fits its client risk profile |
| Risk ladder order | Strategies ranked by realised volatility | Two neighbouring strategies very close | Order reversed (e.g. Cautious riskier than Balanced) | Escalate to the committee |
| Change in weights | Sum of absolute changes in weights vs last year | Above half of the materiality threshold | Above the materiality threshold | Explain what drove the change (returns, risk or constraints) |
| Binding constraints | Which minimum/maximum limits the optimal weights sit on | Same limit binding 2 years running | Same limit binding 3 years running | Review whether the limit is still appropriate |

The thresholds above are illustrative. The team should set them based on experience, and review them once a year.

### Example triggers for an off-cycle review

- A major market event that changes the outlook for an asset class, such as the 2022 rise in interest rates.
- An index used in the model being discontinued or replaced.
- A red result on any metric above.
- A significant change to the strategies or how they are used with clients.

## Finding 2: Data quality – a "Data Checks" tab

Add a "Data Checks" tab to the SAA Input Data workbook, with one row per index series. It updates automatically after each FactSet refresh and shows at a glance whether any series looks wrong before the data are used.

### Example layout

| Column | What it shows | Simple Excel formula (example) | Flag if |
| --- | --- | --- | --- |
| Series name | Index name as used in the model | – | – |
| Source / ticker | FactSet identifier | – | – |
| First date | Earliest date with data | =MINIFS(dates, values, "<>") | Later than last year's first date (history has shortened) |
| Last date | Most recent date with data | =MAXIFS(dates, values, "<>") | Not equal to the expected refresh month-end |
| Number of observations | Count of monthly values | =COUNT(values) | Different from the expected number |
| Missing values | Blank cells within the history | =COUNTBLANK(values) | Greater than 0 |
| Repeated values | Months where the value is identical to the month before | =SUMPRODUCT(--(values=previous values)) | Greater than 0 (may signal a series that stopped updating) |
| Largest monthly move | Biggest monthly return, up or down | =MAX(ABS(returns)) | Above an agreed limit (for example 25%) |
| Change vs last year | Largest difference against last year's archived data over the same dates | =MAX(ABS(this year − last year)) | Greater than 0 (history has been restated) |
| Status | Overall result for the row | =IF(any flag, "CHECK", "OK") | "CHECK" |

A summary cell at the top of the tab can count the number of "CHECK" rows. The run should not continue until it shows zero, or until each flag has been explained in a short note next to the row.

### Two further simple checks

- **Manual inputs.** For values typed in by hand, such as real GDP growth, show the value used last year next to this year's and flag any change above a set amount, so that typing errors stand out.
- **Archived raw data.** Save a copy of the refreshed raw data, as values, in the year's archive folder. This is what makes the "Change vs last year" column possible the following year.

## Finding 3: Implementation – a "Run Log" tab and a "Control Checks" tab

### Run Log tab

Add a "Run Log" tab to the SAA Model workbook. Each time the model is run, one new row is added. Over time, this becomes the record of who did what, when and with which settings.

| Column | Example entry |
| --- | --- |
| Run date | 12/01/2027 |
| Run by | Name of analyst |
| Reviewed by | Name of reviewer (a different person) |
| Purpose of run | Annual refresh / re-run after materiality check / test |
| Data refresh date | Date FactSet data were refreshed in SAA Input Data |
| Data Checks result | "0 CHECK rows" or reference to explanations |
| Strategies run | All ten / list of strategies |
| Optimiser settings | Iterations (500 / 1,000), maximum trials, maximum time, random seed used |
| Simulation settings | Iterations (10,000), random seed used |
| Control Checks result | All passed / list of exceptions and how they were resolved |
| Files saved as | File names and archive folder |
| Comments | Anything unusual, e.g. a step re-run or a manual adjustment |

Recording the random seed is important: with the same inputs, settings and seed, the run can be repeated and should give the same result.

### Control Checks tab

Add a "Control Checks" tab that shows a simple PASS or FAIL for each check, using conditional formatting (green or red). The run is complete only when every check shows PASS, or each FAIL has been explained in the Run Log.

| Check | How it works | PASS if |
| --- | --- | --- |
| Inputs pasted correctly | Compare the sum of each pasted block (returns, volatilities, yields, correlations) in SAA Model with the same sum in SAA Input Data | Sums match |
| Composite weights pasted correctly | Compare the sum of the within-composite weights in each workbook that uses them | Sums match, and each composite totals 100% |
| Correlation matrix is valid | Diagonal values all equal 1, and the matrix is symmetrical | Both true |
| Weights add up | Sum of the seven composite weights for each strategy | Equals 100% |
| Weights within limits | Each weight compared with its minimum and maximum | All within limits |
| Volatility on target | Strategy volatility compared with its target | Within an agreed tolerance (e.g. ±0.1%) |
| Yield met | Portfolio yield compared with the minimum (Income strategies) | At or above the minimum |
| Right strategy exported | Strategy name on the simulation results compared with the strategy selected | Names match |
| Results are current | Date and time stamp on the results compared with the Run Log entry | Same run |

Most of these checks need only simple formulas such as SUM, ABS and IF. The last two may need a small change to the macros, so that each results sheet is stamped with the strategy name and run time.

## Finding 4: Roles and decision rights – a one-page responsibilities table and a judgement register

### Who does what

A single table makes clear who does each task and who signs it off. In the example below, **R** is the person who does the task, **A** the person who signs it off and is accountable for it, **C** someone consulted beforehand, and **I** someone kept informed.

| Activity | Model Owner | Analyst (runs the model) | Reviewer | Investment Committee | MRM |
| --- | --- | --- | --- | --- | --- |
| Refresh data and run Data Checks | A | R | I | – | – |
| Run optimisation and simulation | A | R | I | – | – |
| Complete Control Checks and Run Log | A | R | R | – | – |
| Review results and monitoring summary | A | C | R | I | I |
| Approve the new SAA | R | – | C | A | I |
| Change judgements (targets, limits, thresholds) | R | C | C | A | C |
| Change indices, data sources or the model itself | A | R | C | I | C |
| Manage @RISK licence and vendor support | A | R | – | – | – |

Names rather than job titles can be added, together with a named back-up for the analyst, so that the annual run never depends on one person.

### Expert judgement register

A short table, kept with the model files, listing the main judgements that shape the allocation. It is reviewed once a year as part of the annual refresh.

| Judgement | Current value | Why this value | Owner | Approved by / date | Next review |
| --- | --- | --- | --- | --- | --- |
| Volatility targets | Equally spaced between 30% and 90% of equity volatility | e.g. to match client risk profiles | Model Owner | Investment Committee, \[date\] | Next annual refresh |
| Composite minimum and maximum weights | Per strategy (see model) | e.g. diversification and liquidity | Model Owner | Investment Committee, \[date\] | Next annual refresh |
| Home bias | UK weight four times its market weight | e.g. client demand and currency risk | Model Owner | \[to be confirmed\] | Next annual refresh |
| Blend of expected return approaches | 50% BBA / 50% HVA | e.g. balance between the two approaches | Model Owner | \[to be confirmed\] | Next annual refresh |
| Materiality threshold | Ex-ante standard deviation | e.g. avoid unnecessary portfolio changes | Model Owner | \[to be confirmed\] | Next annual refresh |

The "Why this value" column should be completed by the team; the entries shown are placeholders. Where a judgement has no recorded approval, the register makes that visible so it can be confirmed at the next committee.

## Finding 5: Documentation – a "Read Me" tab and an annual change log

### Read Me tab

Add a "Read Me" tab as the first tab of each workbook (Composite Weights, Forward Looking Return Assumptions, SAA Input Data, SAA Model and SAA Output). Anyone opening the file can then understand it without asking the person who built it.

| Section | What to include |
| --- | --- |
| Purpose | One or two sentences on what the workbook does and where it sits in the process |
| Inputs | Where the data come from (FactSet, other workbooks, manual entries) and which tabs hold them |
| Outputs | What the workbook produces and where it goes next |
| Tabs | A one-line description of each tab, including hidden ones |
| Macros | What each button does, in plain words, and what to do if it fails |
| Linked files | Other workbooks this file reads from or writes to, and the folder they must sit in |
| Owner and version | Model Owner, current version number and date of last change |

### Annual change log

Add a "Change Log" tab to SAA Model, or keep it as a short table in the model documentation. One row is added for each change made at an annual refresh or in between.

| Date | What changed | Why | Effect on the allocation | Approved by |
| --- | --- | --- | --- | --- |
| \[date\] | e.g. LIBOR series replaced by SONIA | LIBOR discontinued | e.g. small change to cash and bond weights | \[name / committee\] |
| \[date\] | e.g. two indices removed from the composites | e.g. no longer investable | \[description\] | \[name / committee\] |
| \[date\] | e.g. volatility target for Cautious changed | e.g. alignment with client risk profiles | \[description\] | \[name / committee\] |

The entries above are illustrative. Over time, the change log answers the question that is otherwise hardest to answer: why does the model look the way it does today?

### Model document

The Read Me tabs and the change log can then feed a short model document, which brings together the purpose and use of the model, its data, method, key judgements, limitations, controls and monitoring. MRM can provide a simple template and help the team reuse the existing methodology papers rather than rewriting them.
