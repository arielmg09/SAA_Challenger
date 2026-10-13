# Validation Findings

Oct 13, 2026 · @Ariel Gimenez

## Introduction

This is the first independent validation of the SAA model, which has supported the Bank's strategic asset allocation for several years. The model's core engine is a vendor tool, Palisade @RISK, used within a set of Excel workbooks. In practice, the model works as a largely mechanical process: data are refreshed, values are transferred between workbooks, and the vendor tool runs the optimisation and simulation. UKRM has therefore focused this review less on the mathematics of the vendor engine and more on the controls, monitoring, governance and documentation that surround it, which is where the main model risk lies.

UKRM recognises the effort involved in running this process each year, and that the gaps identified are typical of a long-standing model that pre-dates the Bank's current model risk framework under PRA SS1/23. UKRM has not identified evidence that the current SAA is unreliable. Together, however, they point to areas where an undetected issue could affect the strategies offered to clients, and where the Bank could struggle to evidence how a given allocation was produced.

The findings are intended as the starting point for a joint, medium-term programme of work between Wealth Management and MRM, not as a list of corrections. For each finding, the recommended actions focus on where UKRM believes the greatest risk sits, and on how MRM can support their design and implementation.

### Overarching recommendation: a joint control enhancement roadmap

UKRM recommends that the Model Owner (MO) and MRM agree a single control enhancement roadmap covering the five findings below, delivered over 12 to 18 months and aligned with the annual refresh cycle. The roadmap should prioritise the actions that most reduce the risk of an undetected error reaching client portfolios (Findings 2 and 3), followed by monitoring and governance (Findings 1 and 4), with documentation (Finding 5) built up progressively alongside them. MRM will support the design of each element and review progress at agreed checkpoints, rather than only at the next validation.

### Summary of findings

| # | Finding | Suggested rating | Core risk |
| --- | --- | --- | --- |
| 1 | No defined monitoring plan | Medium | Strategies may drift away from the risk levels promised to clients without anyone noticing |
| 2 | Absence of data quality controls | Medium | Changes in source data could alter the allocation without being visible as an error |
| 3 | Implementation controls | Medium | A past SAA cannot be reproduced, and manual or silent errors may reach the outputs |
| 4 | Roles, responsibilities and decision rights | Low | The judgements that most drive the allocation have no clear owner or approver |
| 5 | Documentation below minimum standards | Low | The model cannot be properly understood, tiered or reviewed from its documentation |

The suggested ratings reflect the collaborative approach agreed for this first validation and are subject to confirmation by UKRM.

## Finding 1: No defined monitoring plan

**Suggested rating:** Medium

### Observation

The model is re-run once a year, but there is no defined plan to monitor its performance or the behaviour of its outputs between runs. No metrics, tolerances, escalation routes or triggers for an off-cycle review have been set. Review of the outputs relies on an undefined "sense check" during the annual run.

### Why it matters

For a model of this kind, the main risk is not a calculation error but a gradual drift between what the strategies are designed to deliver and what they actually deliver. Each strategy is defined by its volatility target, and that target is what distinguishes Defensive from Growth. If realised risk moves outside the intended band, the mapping between strategies and client risk profiles may no longer hold. This can have implications for suitability and for the outcomes the Bank must evidence under the FCA's Consumer Duty, a link that is not currently made in the model's governance.

Without monitoring, the Bank also cannot tell whether year-on-year changes in the allocation reflect genuine changes in market evidence or noise in the inputs and the optimiser. Nor can it tell whether the constraints, rather than the model, are effectively setting the allocation.

### Recommended actions

1. **Track realised against projected outcomes.** Each year, compare each strategy's realised return and volatility with the range projected by the model, for example the 10th to 90th percentile, over rolling 1-, 3- and 5-year periods. Persistent outcomes outside the range are the clearest sign that the model's assumptions need revisiting.
2. **Monitor risk-ladder integrity.** Check that the strategies remain correctly ordered by risk and return, and that each strategy's realised volatility stays within an agreed tolerance of its target.
3. **Attribute changes in the allocation.** Break down each year's change in weights into the part driven by changes in expected returns, volatilities, correlations and constraints. This turns the annual refresh from "the optimiser said so" into an explainable decision.
4. **Monitor binding constraints.** Record which constraints bind for each strategy each year. Constraints that bind year after year are, in effect, setting the allocation and should be reviewed as expert judgements (see Finding 4).
5. **Define off-cycle review triggers.** Examples include a significant market regime change (such as the 2022 rise in interest rates), the discontinuation or replacement of an index, or a breach of the realised-risk tolerance.
6. **Formalise reporting and escalation.** Summarise the results in a short annual monitoring report to the relevant committee, with clear tolerances and escalation routes.

### How MRM can support

MRM can work with the MO to select a small set of proportionate metrics and tolerances, share monitoring templates used for other models, and provide the change attribution analysis for the first year using UKRM's challenger implementation.

## Finding 2: Absence of data quality controls

**Suggested rating:** Medium

### Observation

Each year, the model rebuilds more than 20 years of index history from FactSet, supplemented by manually updated inputs such as real GDP growth. No data quality checks are defined or evidenced at any stage. In addition:

- the set of indices has changed since the methodology was documented (19 indices in use against 21 described, and LIBOR-based series replaced), with no record of when or why;
- where an index has a shorter history, the composite is rebuilt from the remaining indices, but these proxies are not recorded or approved;
- raw data are not archived, so the history used in one year cannot be compared with that used in the next.

### Why it matters

The most significant data risks for this model would not appear as errors. They would simply produce a different allocation. Examples include:

- **restated history:** data vendors can revise or backfill past index values, which changes volatilities and correlations without any change in the model;
- **index events:** index methodology changes, rebasing or discontinuation, as happened with LIBOR, can break the continuity of a series;
- **silent failures:** a series that fails to refresh may return stale or repeated values that look plausible;
- **index biases:** hedge fund indices in particular can be affected by survivorship and reporting biases, which flatter historic risk and return;
- **unrecorded proxies:** a proxied composite has a different risk profile from the composite the methodology describes.

Because the optimiser responds to small differences in inputs, any of these could shift the allocation of every strategy.

### Recommended actions

1. **Build a data inventory.** For each series, record the source, identifier, frequency, history start date, owner and the fallback if it becomes unavailable.
2. **Archive a snapshot of the raw data at each refresh.** Compare it with the prior year's snapshot over the common period, so that any restated history is identified and explained before it reaches the optimiser.
3. **Introduce simple automated checks.** These should cover completeness, stale or repeated values, outliers against an agreed tolerance, and alignment of dates across series. Results should be retained as evidence of the run.
4. **Govern index changes and proxies.** Keep a register of every index substitution and proxy, with its rationale and approval, and assess the impact of each change on the allocation before adoption.
5. **Review index suitability periodically.** Confirm that each index remains representative of the investable opportunity set, with particular attention to hedge fund and alternative asset indices.

### How MRM can support

MRM can help define a proportionate set of data checks and tolerances, review the data inventory, and share approaches used for data quality controls on other models within the Bank.

## Finding 3: Implementation controls

**Suggested rating:** Medium

### Observation

The annual run relies on a long sequence of manual steps, with limited controls around them:

- inputs are copied as values between workbooks at least eight times, several into hidden sheets, with no reconciliation or second review evidenced;
- the optimiser is stochastic, no random seed is set, and the settings used for each strategy are not recorded;
- constraints are applied as penalties, and no check confirms that the final allocation meets every constraint;
- when a macro step fails, it stops without an error message, and results from a previous strategy can then be exported under the current strategy's name;
- the process depends on a licensed installation of a vendor tool (@RISK), whose version support is unconfirmed and which UKRM could not access to run the model independently.

The detail is set out in the Model Implementation section.

### Why it matters

Beyond the risk of an individual error, the central issue is reproducibility. Because the optimiser is stochastic, inputs are pasted as static values and no run evidence is kept, the Bank would probably be unable to reproduce a past SAA exactly, or to show how a given allocation was produced, if asked by a client, internal audit or a regulator. The same gaps mean that an error introduced during the run, such as a misplaced paste or a stale result after a failed macro, could pass unnoticed into the strategies offered to clients.

### Recommended actions

1. **Create a reproducibility pack for each annual run.** Archive the input data, the values transferred between workbooks, the optimiser settings and random seed, the outputs for each strategy, and the sign-offs. The test is simple: could someone else re-create this year's SAA from the pack?
2. **Introduce key control points with reconciliations.** At each transfer between workbooks, check control totals or a small set of reference values, and record the check on a run checklist signed by the preparer and a reviewer.
3. **Add an independent "dual run".** Each year, solve the same optimisation with a deterministic solver alongside @RISK, and explain any material difference. This gives assurance on the optimiser's result without needing to inspect the vendor engine.
4. **Confirm constraint compliance after optimisation.** Add an explicit check that every constraint is met before the weights are adopted.
5. **Make the macros fail safely.** Display clear error messages, restore Excel's settings on failure, and stamp each results sheet with the strategy and run time so that mismatched results cannot be exported.
6. **Bring the workbooks under end-user computing controls.** Apply version control, access restrictions, protection of formulas and hidden sheets, and a change log.
7. **Plan for continuity of the vendor tool.** Confirm licence ownership and vendor support for the version in use, and agree a fallback if the tool becomes unavailable. Granting UKRM access, or arranging a supervised run, would also allow future validations to test the model directly.

In the longer term, the MO may wish to consider whether linking the workbooks directly, or consolidating the process, would remove most of the manual steps altogether.

### How MRM can support

UKRM's challenger implementation, which already solves the optimisation deterministically, could serve as the starting point for the dual run. MRM can also help design the run checklist and key control points, and share end-user computing control templates used elsewhere in the Bank.

## Finding 4: Roles, responsibilities and decision rights

**Suggested rating:** Low

### Observation

The roles of model owner, developer, user, reviewer and approver are not formally defined for the SAA model. The same individuals appear to refresh the data, run the model, review the outputs and decide on the allocation, and the process appears to depend on very few people. No one is named as responsible for the relationship with the @RISK vendor.

The allocation also depends heavily on a set of expert judgements. Their ownership, rationale and approval are not documented. They include:

- the volatility targets for each strategy, set between 30% and 90% of equity volatility;
- the minimum and maximum weights for each composite;
- the four-times home bias in equities and government bonds;
- the equal blend of the two expected return approaches (BBA and HVA);
- the materiality threshold and the decision to retain the prior-year allocation where a change is not material.

### Why it matters

In a model where the optimiser is a vendor tool, these judgements, rather than the optimiser itself, are often what shape the final allocation. A volatility target or a composite bound can matter more than any change in expected returns. If no one is clearly accountable for setting and challenging them, they can remain unchanged for years after the reasons behind them have moved on.

The lack of segregation between running and reviewing the model removes a natural line of defence against errors, and dependence on a few individuals creates a continuity risk for an annual process that is rarely repeated.

### Recommended actions

1. **Agree a RACI for the model lifecycle.** Cover data refresh, model run, output review, approval of the SAA, model changes, monitoring and vendor management.
2. **Create an expert-judgement register.** For each judgement listed above, record its rationale, the evidence supporting it, its owner, the approving committee and how often it is reviewed. Revisit the register as part of each annual refresh.
3. **Separate running from reviewing.** Ensure that the person who runs the model is not the only person who reviews its outputs before approval.
4. **Plan for continuity.** Train at least one additional person to perform the full annual run, using the updated instructions from Finding 5.
5. **Assign vendor ownership.** Name an owner for the @RISK licence, version upgrades and vendor support.

### How MRM can support

MRM can share RACI and expert-judgement templates used for other models, and help the MO identify which judgements are most material by testing their effect on the allocation with UKRM's challenger.

## Finding 5: Documentation below minimum standards

**Suggested rating:** Low

### Observation

The model is documented through two methodology papers dated 2020 to 2022, covering inputs and expected returns, and a step-by-step instructions document. Measured against the Bank's minimum standards under PRA SS1/23:

- there is no document describing the optimisation and simulation methodology, the model's assumptions and limitations, or the testing performed;
- the instructions no longer match the model in use, for example in the number of indices, LIBOR references, cell ranges and the description of the matrix used for the materiality check;
- there is no statement of how the model's outputs are used downstream, which portfolios and client processes rely on them, or what the impact of an error would be;
- changes made at each annual refresh are not recorded.

### Why it matters

Documentation is what allows the model to be understood, run and challenged by someone other than its current operators. Without a clear statement of the model's purpose and use, the Bank cannot reliably assess its materiality or tier, set proportionate controls, or judge the impact of a limitation. Out-of-date instructions also increase the chance of an error during the annual run, particularly if it is performed by someone new (see Finding 4).

### Recommended actions

1. **Produce a proportionate model document.** Cover the purpose and use of the model, data, methodology (including optimisation and simulation), key assumptions and limitations, expert judgements, testing and monitoring. Existing methodology papers can be referenced rather than rewritten.
2. **Write a model-use statement.** Describe which portfolios, products and client processes rely on the SAA, so that the model's materiality and tier can be confirmed.
3. **Maintain a limitations register.** Record the model's known limitations, such as the steady-state assumptions, the reliance on historic volatilities and correlations, and the absence of independent execution of the vendor tool, with any compensating controls.
4. **Bring the instructions up to date and keep them current.** Align them with the model in use, place them under version control with the workbooks, and review them at each annual refresh.
5. **Keep an annual change log.** Record what changed at each refresh (data, indices, judgements, settings) and why.

### How MRM can support

MRM can provide a model documentation template aligned with SS1/23, review drafts as they are developed, and help identify which existing material can be reused, so that the documentation effort stays proportionate to the model.
