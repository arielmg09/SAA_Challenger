"""
saa_helpers.py - simple helper functions for the SAA challenger notebook.

Each function does ONE job. The notebook calls them in order:
    read_inputs  ->  simulate_returns  ->  optimise_replica / optimise_analytical
    ->  portfolio_stats  ->  look_through  ->  save_results
"""

import numpy as np
import pandas as pd
from scipy import stats
from scipy.optimize import differential_evolution, minimize

# The 7 investable composites, and Inflation (simulated, but not investable).
ASSETS = ["Equity", "DMGovt", "BondOther", "Property", "Gold", "HFs", "Cash"]
ALL_VARS = ASSETS + ["Inflation"]


# -------------------------------------------------------------------------------------
# 1. READ INPUTS
# -------------------------------------------------------------------------------------
def read_inputs(file_path):
    """Read all tabs of the input template and return them in a dictionary."""
    returns = pd.read_excel(file_path, sheet_name="ReturnAssumptions", index_col=0)["ExpectedReturn"]
    vols = pd.read_excel(file_path, sheet_name="Volatility", index_col=0)["Volatility"]
    yields = pd.read_excel(file_path, sheet_name="Yield", index_col=0)["Yield"]

    corr = pd.read_excel(file_path, sheet_name="CorrelationMatrix", index_col=0)
    corr = corr.loc[ALL_VARS, ALL_VARS]        # keep only the 8x8 block (ignores check cells)

    asset_map = pd.read_excel(file_path, sheet_name="AssetClassMap", usecols="A:C")
    asset_map.columns = ["Composite", "Index", "Weight"]
    asset_map = asset_map.dropna(subset=["Index"])
    asset_map = asset_map[asset_map["Composite"].isin(ASSETS)]   # drops the source note

    setup = pd.read_excel(file_path, sheet_name="Setup", index_col=0)
    setup = setup.dropna(how="all")
    setup = setup.loc[[r for r in setup.index if not str(r).startswith(("Check", "Strategy columns"))]]

    settings = pd.read_excel(file_path, sheet_name="Settings", index_col=0)["Value"].to_dict()

    return {
        "returns": returns.loc[ALL_VARS].astype(float),
        "vols": vols.loc[ALL_VARS].astype(float),
        "yields": yields.loc[ASSETS].astype(float),
        "corr": corr.astype(float),
        "asset_map": asset_map,
        "setup": setup.astype(float),
        "settings": settings,
    }


def check_correlation(corr):
    """Cholesky check: a valid correlation matrix must be decomposable."""
    try:
        np.linalg.cholesky(corr.values)
        print("Correlation matrix OK - Cholesky decomposition succeeded.")
    except np.linalg.LinAlgError:
        print("WARNING: correlation matrix is NOT positive definite - results unreliable.")
    print("Smallest eigenvalue:", round(np.linalg.eigvalsh(corr.values).min(), 4))


def covariance_matrix(inputs):
    """Covariance = vol_i * vol_j * corr_ij (investable assets only)."""
    v = inputs["vols"][ASSETS].values
    c = inputs["corr"].loc[ASSETS, ASSETS].values
    return np.outer(v, v) * c


# -------------------------------------------------------------------------------------
# 2. SIMULATION
# -------------------------------------------------------------------------------------
def simulate_returns(inputs, n_iterations, n_years, seed=1, distribution="Normal"):
    """
    Simulate correlated annual returns. Result has shape (n_iterations, n_years, 8).

    Step 1: draw independent returns for each asset (Normal or Lognormal).
    Step 2: re-order (shuffle) each column so the RANK correlation matches the
            target matrix (Iman-Conover method - what @RISK does). The values
            themselves are not changed, only their pairing across assets.
    """
    rng = np.random.default_rng(seed)
    n = n_iterations * n_years
    mu = inputs["returns"].values
    sd = inputs["vols"].values

    # Step 1 - independent draws
    z = rng.standard_normal((n, len(ALL_VARS)))
    if distribution == "Lognormal":
        s2 = np.log(1 + sd**2 / (1 + mu) ** 2)
        m = np.log(1 + mu) - s2 / 2
        draws = np.exp(m + np.sqrt(s2) * z) - 1
    else:
        draws = mu + sd * z

    # Step 2 - Iman-Conover re-ordering
    target = inputs["corr"].values
    scores = stats.norm.ppf(np.arange(1, n + 1) / (n + 1))       # a ranked "template" column
    s = np.column_stack([rng.permutation(scores) for _ in ALL_VARS])
    f = np.linalg.cholesky(np.corrcoef(s, rowvar=False))         # correlation it happens to have
    p = np.linalg.cholesky(target)                               # correlation we want
    t = s @ np.linalg.inv(f).T @ p.T                             # template with target correlation
    correlated = np.empty_like(draws)
    for j in range(len(ALL_VARS)):
        order = stats.rankdata(t[:, j], method="ordinal").astype(int) - 1
        correlated[:, j] = np.sort(draws[:, j])[order]

    return correlated.reshape(n_iterations, n_years, len(ALL_VARS))


# -------------------------------------------------------------------------------------
# 3. OBJECTIVE AND PENALTIES (mirrors the RISKOptimizer set-up)
# -------------------------------------------------------------------------------------
def penalised_objective(weights, sims, inputs, constraints):
    """
    Value to MAXIMISE = mean portfolio value at the end of the horizon (W97)
                        minus penalties for any constraint breach.
    Penalty = (1000 x breach)^2 for volatility and yield, (100 x breach)^2 otherwise.
    """
    st = inputs["settings"]
    k_big, k_small = st["PenaltyFactorVolYield"], st["PenaltyFactorOther"]

    port_returns = sims[:, :, :7] @ weights                          # (iterations, years)
    final_value = st["InitialValue"] * np.prod(1 + port_returns, axis=1)
    mean_value = final_value.mean()

    vol = port_returns.std()
    port_yield = weights @ inputs["yields"].values

    if st["VolConstraintType"] == "target":
        vol_breach = abs(vol - constraints["Annualised Volatility"])
    else:
        vol_breach = max(0, vol - constraints["Annualised Volatility"])
    penalty = (k_big * vol_breach) ** 2
    penalty += (k_big * max(0, constraints["Min Yield"] - port_yield)) ** 2
    for i, a in enumerate(ASSETS):
        penalty += (k_small * max(0, constraints[f"Min {a}"] - weights[i])) ** 2
        penalty += (k_small * max(0, weights[i] - constraints[f"Max {a}"])) ** 2

    return mean_value - penalty


# -------------------------------------------------------------------------------------
# 4. OPTIMISERS
# -------------------------------------------------------------------------------------
def optimise_replica(inputs, constraints, n_iterations, seed=1, start=None, generations=300):
    """
    Search over integer weights 0..1000 (normalised to sum to 100%), like the
    RISKOptimizer adjustable cells. Uses differential evolution, a genetic-style
    algorithm. Returns (raw integers, weights).
    """
    st = inputs["settings"]
    sims = simulate_returns(inputs, n_iterations, int(st["HorizonYears"]), seed, st["Distribution"])

    def to_minimise(x):
        if x.sum() == 0:
            return 1e9
        w = x / x.sum()
        return -penalised_objective(w, sims, inputs, constraints)

    result = differential_evolution(
        to_minimise,
        bounds=[(0, 1000)] * 7,
        integrality=[True] * 7,
        popsize=int(st["PopulationPerAsset"]),
        maxiter=generations,
        tol=1e-10,
        seed=seed,
        x0=start,
        polish=False,
    )
    raw = np.round(result.x)
    return raw, raw / raw.sum()


def optimise_analytical(inputs, constraints):
    """
    Deterministic benchmark (no simulation, no penalties):
        maximise   w' mu
        subject to sum(w) = 1, min <= w <= max,
                   sqrt(w' Cov w) <= target vol,  w' yield >= min yield
    """
    mu = inputs["returns"][ASSETS].values
    cov = covariance_matrix(inputs)
    y = inputs["yields"].values
    target = constraints["Annualised Volatility"]
    bounds = [(constraints[f"Min {a}"], constraints[f"Max {a}"]) for a in ASSETS]

    cons = [
        {"type": "eq", "fun": lambda w: w.sum() - 1},
        {"type": "ineq", "fun": lambda w: target**2 - w @ cov @ w},
        {"type": "ineq", "fun": lambda w: w @ y - constraints["Min Yield"]},
    ]
    if inputs["settings"]["VolConstraintType"] == "target":
        cons[1]["type"] = "eq"

    best = None
    rng = np.random.default_rng(0)
    for _ in range(20):                                  # several starting points
        start = rng.dirichlet(np.ones(7))
        res = minimize(lambda w: -(w @ mu), start, method="SLSQP", bounds=bounds, constraints=cons)
        if res.success and (best is None or res.fun < best.fun):
            best = res
    if best is None:
        print("   Analytical: no feasible portfolio for these constraints.")
        return None
    return best.x / best.x.sum()


# -------------------------------------------------------------------------------------
# 5. OUTPUT HELPERS
# -------------------------------------------------------------------------------------
def portfolio_stats(weights, inputs, constraints, sims):
    """Key statistics for one set of weights, using the final (large) simulation."""
    st = inputs["settings"]
    port_returns = sims[:, :, :7] @ weights
    final_value = st["InitialValue"] * np.prod(1 + port_returns, axis=1)
    inflation = np.prod(1 + sims[:, :, 7], axis=1)

    breaches = []
    vol_sim = port_returns.std()
    if vol_sim > constraints["Annualised Volatility"] + 1e-4:
        breaches.append("Volatility")
    if weights @ inputs["yields"].values < constraints["Min Yield"] - 1e-4:
        breaches.append("Min Yield")
    for i, a in enumerate(ASSETS):
        if weights[i] < constraints[f"Min {a}"] - 1e-4 or weights[i] > constraints[f"Max {a}"] + 1e-4:
            breaches.append(a)

    return {
        "Expected return": weights @ inputs["returns"][ASSETS].values,
        "Volatility (analytic)": np.sqrt(weights @ covariance_matrix(inputs) @ weights),
        "Volatility (simulated)": vol_sim,
        "Yield": weights @ inputs["yields"].values,
        "Mean final value (W97)": final_value.mean(),
        "Median final value": np.median(final_value),
        "5th percentile final value": np.percentile(final_value, 5),
        "Mean real final value": (final_value / inflation).mean(),
        "Breaches (>0.01%)": ", ".join(breaches) if breaches else "None",
    }


def look_through(weights_table, asset_map):
    """Composite weights x index weights within each composite = index-level weights."""
    out = asset_map[["Composite", "Index"]].copy()
    for col in weights_table.columns:
        out[col] = [weights_table.loc[c, col] * w for c, w in zip(asset_map["Composite"], asset_map["Weight"])]
    return out


def save_results(file_path, sheets):
    """Write a dictionary {sheet name: DataFrame} to Excel."""
    with pd.ExcelWriter(file_path) as writer:
        for name, df in sheets.items():
            df.to_excel(writer, sheet_name=name)
    print("Results saved to", file_path)
