"""
saa_helpers.py - helper functions for the SAA challenger notebook (UKRM).

Written to work with older library versions (numpy 1.16+, scipy 1.2+, pandas 0.24+).

The notebook uses these functions in four blocks:
  A. Optimisation        -> composite weights per strategy, then index-level weights
  B. Historic simulation -> "Historic Characteristics" (20-year back-test, monthly data)
  C. Monte Carlo         -> "Projected Characteristics" (20-year forward simulation)
  D. Output              -> Excel export
"""

import numpy as np
import pandas as pd
from scipy import stats
from scipy.optimize import differential_evolution, minimize

# The 7 investable composites, plus Inflation (simulated, but not investable).
ASSETS = ["Equity", "DMGovt", "BondOther", "Property", "Gold", "HFs", "Cash"]
ALL_VARS = ASSETS + ["Inflation"]


# =====================================================================================
# INPUT DATA READING
# =====================================================================================
def read_inputs(file_path):
    """Read the input template. The two historic tabs are optional."""
    xl = pd.ExcelFile(file_path)

    returns = xl.parse("ReturnAssumptions", index_col=0)["ExpectedReturn"]
    vols = xl.parse("Volatility", index_col=0)["Volatility"]
    yields = xl.parse("Yield", index_col=0)["Yield"]
    corr = xl.parse("CorrelationMatrix", index_col=0).loc[ALL_VARS, ALL_VARS]

    asset_map = xl.parse("AssetClassMap", usecols="A:C")
    asset_map.columns = ["Composite", "Index", "Weight"]
    asset_map = asset_map.dropna(subset=["Index"])
    asset_map = asset_map[asset_map["Composite"].isin(ASSETS)].reset_index(drop=True)

    setup = xl.parse("Setup", index_col=0).dropna(how="all")
    setup = setup.loc[[r for r in setup.index if not str(r).startswith(("Check", "Strategy columns"))]]
    setup = setup[[c for c in setup.columns if not str(c).startswith("Unnamed")]]

    settings = xl.parse("Settings", index_col=0)["Value"].to_dict()

    inputs = {
        "returns": returns.loc[ALL_VARS].astype(float),
        "vols": vols.loc[ALL_VARS].astype(float),
        "yields": yields.loc[ASSETS].astype(float),
        "corr": corr.astype(float),
        "asset_map": asset_map,
        "setup": setup.astype(float),
        "settings": settings,
        "hist_returns": None,
        "hist_yields": None,
    }

    # Optional: monthly historic returns of the composites (+ Cash and Inflation)
    if "HistoricReturns" in xl.sheet_names:
        h = xl.parse("HistoricReturns", index_col=0)
        h.index = pd.to_datetime(h.index)
        inputs["hist_returns"] = h[ALL_VARS].astype(float).sort_index()
    # Optional: monthly historic yields of the composites
    if "HistoricYields" in xl.sheet_names:
        y = xl.parse("HistoricYields", index_col=0)
        y.index = pd.to_datetime(y.index)
        inputs["hist_yields"] = y[ASSETS].astype(float).sort_index()

    return inputs


def check_inputs(inputs):
    """Simple checks on the inputs; prints the results."""
    try:
        np.linalg.cholesky(inputs["corr"].values)
        print("OK  Correlation matrix passes the Cholesky check.")
    except np.linalg.LinAlgError:
        print("!!  Correlation matrix FAILS the Cholesky check - results unreliable.")

    totals = inputs["asset_map"].groupby("Composite")["Weight"].sum()
    bad = totals[(totals - 1).abs() > 1e-4]
    if len(bad) == 0:
        print("OK  Index weights add up to 100% in every composite.")
    else:
        print("!!  Index weights do not add up to 100% for:", dict(bad.round(4)))

    for name in ["hist_returns", "hist_yields"]:
        df = inputs[name]
        if df is None:
            print(f"--  No '{name}' tab found - this part of the analysis will be skipped.")
        else:
            print(f"OK  {name}: {len(df)} months, {df.index[0]:%b %Y} to {df.index[-1]:%b %Y}")


def covariance_matrix(inputs):
    """Covariance = vol_i x vol_j x corr_ij (investable assets only)."""
    v = inputs["vols"][ASSETS].values
    return np.outer(v, v) * inputs["corr"].loc[ASSETS, ASSETS].values


def volatility_targets(inputs, risk_levels, n_levels=4, low=0.30, high=0.90):
    """
    Volatility targets evenly spaced between 30% and 90% of equity volatility.
    risk_levels: {strategy name: level 1..n_levels}. Strategies not in the
    dictionary keep the target typed in the Setup tab.
    """
    eq_vol = inputs["vols"]["Equity"]
    steps = np.linspace(low, high, n_levels)
    targets = {}
    for s in inputs["setup"].columns:
        if s in risk_levels:
            targets[s] = eq_vol * steps[risk_levels[s] - 1]
        else:
            targets[s] = inputs["setup"].loc["Annualised Volatility", s]
    return pd.Series(targets)


# =====================================================================================
# MONTE CARLO SIMULATION (used in Block A and Block C)
# =====================================================================================
def simulate_returns(inputs, n_iterations, n_years, seed=1, distribution="Normal"):
    """
    Simulate correlated ANNUAL returns. Result shape: (n_iterations, n_years, 8).

    Step 1: draw independent returns for each asset (Normal or Lognormal).
    Step 2: re-order each column so the RANK correlation matches the target
            matrix (Iman-Conover, as used by @RISK). Values are not changed,
            only how they are paired across assets.
    """
    rng = np.random.RandomState(seed)
    n = n_iterations * n_years
    mu = inputs["returns"].values
    sd = inputs["vols"].values

    z = rng.standard_normal((n, len(ALL_VARS)))
    if distribution == "Lognormal":
        s2 = np.log(1 + sd**2 / (1 + mu) ** 2)
        draws = np.exp(np.log(1 + mu) - s2 / 2 + np.sqrt(s2) * z) - 1
    else:
        draws = mu + sd * z

    scores = stats.norm.ppf(np.arange(1, n + 1) / (n + 1))
    s = np.column_stack([rng.permutation(scores) for _ in ALL_VARS])
    f = np.linalg.cholesky(np.corrcoef(s, rowvar=False))
    p = np.linalg.cholesky(inputs["corr"].values)
    t = s @ np.linalg.inv(f).T @ p.T
    correlated = np.empty_like(draws)
    for j in range(len(ALL_VARS)):
        order = stats.rankdata(t[:, j], method="ordinal").astype(int) - 1
        correlated[:, j] = np.sort(draws[:, j])[order]

    return correlated.reshape(n_iterations, n_years, len(ALL_VARS))


def portfolio_volatility(port_returns, measure):
    """
    Volatility of simulated annual portfolio returns (iterations x years).
      'pooled'    : one standard deviation across all simulated years
      'mean_path' : standard deviation within each 20-year path, then averaged
    """
    if measure == "mean_path":
        return port_returns.std(axis=1, ddof=1).mean()
    return port_returns.std(ddof=1)


# =====================================================================================
# BLOCK A - OPTIMISATION
# =====================================================================================
def penalised_objective(weights, sims, inputs, constraints, vol_target, vol_measure):
    """
    Value to MAXIMISE = mean portfolio value in the final year (cell W97)
                        minus penalties for constraint breaches:
        (1000 x breach)^2 for volatility and minimum yield
        (100  x breach)^2 for each asset-class min / max
    """
    st = inputs["settings"]
    k_big, k_small = st["PenaltyFactorVolYield"], st["PenaltyFactorOther"]

    port_returns = sims[:, :, :7] @ weights
    final_value = st["InitialValue"] * np.prod(1 + port_returns, axis=1)

    vol = portfolio_volatility(port_returns, vol_measure)
    if st["VolConstraintType"] == "target":
        vol_breach = abs(vol - vol_target)
    else:
        vol_breach = max(0, vol - vol_target)
    penalty = (k_big * vol_breach) ** 2
    penalty += (k_big * max(0, constraints["Min Yield"] - weights @ inputs["yields"].values)) ** 2
    for i, a in enumerate(ASSETS):
        penalty += (k_small * max(0, constraints["Min " + a] - weights[i])) ** 2
        penalty += (k_small * max(0, weights[i] - constraints["Max " + a])) ** 2

    return final_value.mean() - penalty


def optimise_replica(inputs, constraints, vol_target, n_iterations, seed=1,
                     start=None, generations=300, vol_measure="pooled"):
    """
    Search over whole-number weights 0..1000 (rescaled to 100%), as in the
    RISKOptimizer adjustable cells. Differential evolution = genetic-style search.
    Returns (raw whole numbers, weights).
    """
    st = inputs["settings"]
    sims = simulate_returns(inputs, n_iterations, int(st["HorizonYears"]), seed, st["Distribution"])

    def to_minimise(x):
        x = np.round(x)
        if x.sum() == 0:
            return 1e9
        return -penalised_objective(x / x.sum(), sims, inputs, constraints, vol_target, vol_measure)

    pop_size = int(st["PopulationPerAsset"]) * 7
    init = np.random.RandomState(seed).uniform(0, 1000, size=(pop_size, 7))
    if start is not None:
        init[0] = start

    result = differential_evolution(to_minimise, bounds=[(0, 1000)] * 7, init=init,
                                    maxiter=generations, tol=1e-10, seed=seed, polish=False)
    raw = np.round(result.x)
    return raw, raw / raw.sum()


def optimise_analytical(inputs, constraints, vol_target):
    """
    Benchmark with no simulation and no penalties:
        maximise w'mu   subject to   sum(w)=1, min<=w<=max,
                                     sqrt(w'Cov w) <= vol target, w'yield >= min yield
    """
    mu = inputs["returns"][ASSETS].values
    cov = covariance_matrix(inputs)
    y = inputs["yields"].values
    bounds = [(constraints["Min " + a], constraints["Max " + a]) for a in ASSETS]
    cons = [
        {"type": "eq", "fun": lambda w: w.sum() - 1},
        {"type": "ineq", "fun": lambda w: vol_target**2 - w @ cov @ w},
        {"type": "ineq", "fun": lambda w: w @ y - constraints["Min Yield"]},
    ]
    if inputs["settings"]["VolConstraintType"] == "target":
        cons[1]["type"] = "eq"

    best = None
    rng = np.random.RandomState(0)
    for _ in range(20):
        res = minimize(lambda w: -(w @ mu), rng.dirichlet(np.ones(7)), method="SLSQP",
                       bounds=bounds, constraints=cons)
        if res.success and (best is None or res.fun < best.fun):
            best = res
    if best is None:
        print("     Analytical: no portfolio meets all constraints.")
        return None
    return best.x / best.x.sum()


def unbundle(composite_weights, asset_map):
    """Index weight = composite weight x index weight within the composite.
    composite_weights: DataFrame, rows = portfolios, columns = ASSETS."""
    out = pd.DataFrame(index=composite_weights.index)
    for _, row in asset_map.iterrows():
        out[row["Index"]] = composite_weights[row["Composite"]] * row["Weight"]
    return out


# =====================================================================================
# BLOCK B - HISTORIC CHARACTERISTICS (back-test on monthly data)
# =====================================================================================
def backtest_returns(weights, hist_returns, rebalance="monthly"):
    """Monthly portfolio returns for fixed target weights.
    'monthly' = back to target every month; 'annual' = weights drift for 12 months."""
    r = hist_returns[ASSETS].values
    if rebalance == "monthly":
        return pd.Series(r @ weights, index=hist_returns.index)

    out, w = [], np.array(weights, dtype=float)
    for t in range(len(r)):
        if t % 12 == 0:
            w = np.array(weights, dtype=float)       # rebalance
        port = w @ r[t]
        out.append(port)
        w = w * (1 + r[t]) / (1 + port)            # weights drift with returns
    return pd.Series(out, index=hist_returns.index)


def annualised(monthly):
    """Compound annual growth rate of a monthly return series."""
    return (1 + monthly).prod() ** (12.0 / len(monthly)) - 1


def historic_characteristics(weights, inputs, years=20, rebalance="monthly"):
    """Back-test the portfolio over the last `years` years of monthly data."""
    h = inputs["hist_returns"].iloc[-12 * years:]
    port = backtest_returns(weights, h, rebalance)

    value = (1 + port).cumprod()
    rolling_12m = (1 + port).rolling(12).apply(np.prod, raw=True) - 1
    peak_so_far = value.cummax()
    drawdown = value / peak_so_far - 1

    valley = drawdown.idxmin()
    peak = value.loc[:valley].idxmax()
    after = value.loc[valley:]
    recovered = after[after >= value.loc[peak]]
    if len(recovered) > 0:
        months_to_recovery = len(value.loc[valley:recovered.index[0]]) - 1
    else:
        months_to_recovery = "Not recovered"

    if inputs["hist_yields"] is not None:
        hy = inputs["hist_yields"].reindex(h.index).dropna()
        avg_yield = (hy.values @ weights).mean()
    else:
        avg_yield = np.nan

    return {
        "Annualised return": annualised(port),
        "Annualised excess over cash": annualised(port) - annualised(h["Cash"]),
        "Annualised excess over inflation": annualised(port) - annualised(h["Inflation"]),
        "Maximum 12-month return": rolling_12m.max(),
        "Annualised portfolio volatility": port.std() * np.sqrt(12),
        "Minimum 12-month return": rolling_12m.min(),
        "Maximum drawdown": drawdown.min(),
        "Drawdown peak (period)": peak.strftime("%b %Y"),
        "Drawdown valley (period)": valley.strftime("%b %Y"),
        "Months to recovery": months_to_recovery,
        "Average gross portfolio yield": avg_yield,
    }


# =====================================================================================
# BLOCK C - PROJECTED CHARACTERISTICS (Monte Carlo, annual steps)
# =====================================================================================
def projected_characteristics(weights, sims, inputs):
    """Statistics of the 20-year forward simulation for one portfolio."""
    years = sims.shape[1]
    port = sims[:, :, :7] @ weights                       # (paths, years)
    growth = np.prod(1 + port, axis=1)
    ann_ret = growth ** (1.0 / years) - 1                  # annualised return per path
    ann_cash = np.prod(1 + sims[:, :, 6], axis=1) ** (1.0 / years) - 1
    ann_infl = np.prod(1 + sims[:, :, 7], axis=1) ** (1.0 / years) - 1

    value = np.cumprod(1 + port, axis=1)
    value = np.hstack([np.ones((value.shape[0], 1)), value])   # include the start value
    drawdown = (value / np.maximum.accumulate(value, axis=1) - 1).min(axis=1)

    return {
        "Annualised return (mean)": ann_ret.mean(),
        "Annualised return (90% probability)": np.percentile(ann_ret, 10),
        "Annualised return (10% probability)": np.percentile(ann_ret, 90),
        "Mean annualised excess over cash": (ann_ret - ann_cash).mean(),
        "Mean annualised excess over inflation": (ann_ret - ann_infl).mean(),
        "Mean annualised portfolio volatility": port.std(axis=1, ddof=1).mean(),
        "Mean minimum 12-month return": port.min(axis=1).mean(),
        "Maximum drawdown (mean)": drawdown.mean(),
        "Maximum drawdown (10% probability)": np.percentile(drawdown, 10),
        "Current gross portfolio yield": weights @ inputs["yields"].values,
    }


# =====================================================================================
# OUTPUT
# =====================================================================================
def to_table(results):
    """{strategy: {statistic: value}} -> table with statistics in their original order
    (older pandas versions would otherwise sort the rows alphabetically)."""
    first = next(iter(results.values()))
    return pd.DataFrame(results).loc[list(first.keys()), list(results.keys())]


def save_results(file_path, sheets):
    """Write {sheet name: DataFrame} to one Excel file."""
    writer = pd.ExcelWriter(file_path)
    for name, df in sheets.items():
        df.to_excel(writer, sheet_name=name)
    writer.close()
    print("Results saved to", file_path)
