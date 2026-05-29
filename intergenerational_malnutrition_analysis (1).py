"""
Intergenerational Malnutrition Analysis — NFHS-5 (India, 29 States)
=====================================================================
Objectives
----------
1. Describe the maternal-child nutritional landscape across Indian states.
2. Quantify the intergenerational gradient: child stunting by maternal BMI category.
3. Identify ecological predictors of the Thin-Fat Index (OLS regression).
4. Test the DOHaD mediation hypothesis: Maternal Anaemia → Child Stunting → TFI.

Usage
-----
Place the six NFHS-5 Excel files in the same folder as this script (or set
DATA_DIR below), then run:

    python intergenerational_malnutrition_analysis.py

Required packages
-----------------
    pip install pandas numpy matplotlib seaborn scipy statsmodels pingouin openpyxl

Output
------
All figures (.png) and a summary Excel workbook are written to OUT_DIR.
"""

# ── Standard library ──────────────────────────────────────────────────────────
import os
import warnings

# ── Third-party ───────────────────────────────────────────────────────────────
import numpy as np
import pandas as pd

import matplotlib
matplotlib.use("Agg")                           # headless rendering
import matplotlib.patches as mpatches
import matplotlib.pyplot as plt

import seaborn as sns
from scipy import stats
from scipy.stats import f_oneway, kruskal, shapiro

import statsmodels.api as sm
import statsmodels.formula.api as smf
from statsmodels.stats.outliers_influence import variance_inflation_factor

import pingouin as pg
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

warnings.filterwarnings("ignore")
pd.set_option("display.float_format", "{:.2f}".format)

# ── Paths ──────────────────────────────────────────────────────────────────────
# DATA_DIR  : folder containing the six NFHS-5 Excel files
# OUT_DIR   : folder where figures and the Excel report are saved
DATA_DIR = os.path.dirname(os.path.abspath(__file__))   # same folder as script
OUT_DIR  = os.path.join(DATA_DIR, "outputs")
os.makedirs(OUT_DIR, exist_ok=True)

# ── Visual style ───────────────────────────────────────────────────────────────
PAL  = ["#2E4057", "#048A81", "#54C6EB", "#EF8C5B", "#E03C31", "#6A0572"]
CMAP = "YlOrRd"
sns.set_theme(style="whitegrid", font_scale=1.05)

# ══════════════════════════════════════════════════════════════════════════════
# 1. DATA LOADING
# ══════════════════════════════════════════════════════════════════════════════

def load_excel(filename: str) -> pd.DataFrame:
    """Load an Excel file from DATA_DIR; raise a clear error if not found."""
    path = os.path.join(DATA_DIR, filename)
    if not os.path.exists(path):
        raise FileNotFoundError(
            f"\n[ERROR] Expected file not found:\n  {path}\n"
            "Please place all six NFHS-5 Excel files in the same directory as "
            "this script and re-run."
        )
    return pd.read_excel(path)


bmi_r   = load_excel("BMI - Copy.xlsx")
whr_r   = load_excel("WHR.xlsx")
ana_a_r = load_excel("Prevalence of anaemia in adults.xlsx")
ana_c_r = load_excel("Prevalence of anaemia in children.xlsx")
ch_r    = load_excel("Nutritional status of children.xlsx")
anc_r   = load_excel("Antenatal care indicators.xlsx")

# ── Column renaming ────────────────────────────────────────────────────────────
bmi_r.columns = [
    "State", "Var", "SubVar",
    "pct_thin", "pct_mild_thin", "pct_mod_sev_thin",
    "pct_ow_obese", "pct_ow", "pct_obese", "N",
]

whr_r.columns = [
    "State", "Var", "SubVar",
    "pct_whr_normal", "pct_whr_high", "N",
]

ana_a_r.columns = [
    "State", "Var", "SubVar",
    "pct_anaemia_mild", "pct_anaemia_mod", "pct_anaemia_sev",
    "pct_anaemia_any", "N",
]

ana_c_r.columns = [
    "State", "Var", "SubVar",
    "pct_child_anaemia_mild", "pct_child_anaemia_mod",
    "pct_child_anaemia_sev", "N",
]

ch_r.columns = [
    "State", "Var", "SubVar",
    "pct_stunting", "N_stunt",
    "pct_wasting", "pct_child_ow", "N",
]

# ── State key helper ───────────────────────────────────────────────────────────
def norm_state(s: str) -> str:
    return str(s).strip().lower().replace(" ", "_")


for df in [bmi_r, whr_r, ana_a_r, ana_c_r, ch_r]:
    df["state_key"] = df["State"].apply(norm_state)

# ══════════════════════════════════════════════════════════════════════════════
# 2. AGGREGATION HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def pct_to_count_agg(
    df: pd.DataFrame,
    var_filter: str,
    pct_cols: list,
    n_col: str = "N",
) -> pd.DataFrame:
    """
    Filter rows where Var == var_filter, convert % columns to absolute counts,
    then return one weighted-mean % row per state.
    """
    sub = df[df["Var"] == var_filter].copy()
    sub[n_col] = pd.to_numeric(sub[n_col], errors="coerce").fillna(0)

    for col in pct_cols:
        sub[col] = pd.to_numeric(sub[col], errors="coerce").fillna(0)
        sub[f"cnt_{col}"] = sub[col] * sub[n_col] / 100.0

    cnt_cols = [f"cnt_{c}" for c in pct_cols]
    agg = sub.groupby("state_key").agg(
        {n_col: "sum", **{c: "sum" for c in cnt_cols}}
    ).reset_index()

    for col, cnt in zip(pct_cols, cnt_cols):
        agg[col] = np.where(
            agg[n_col] > 0,
            agg[cnt] / agg[n_col] * 100,
            np.nan,
        )
        agg.drop(columns=cnt, inplace=True)

    return agg


# ── State-level summaries ──────────────────────────────────────────────────────
bmi_state = pct_to_count_agg(
    bmi_r, "Residence",
    ["pct_thin", "pct_mild_thin", "pct_mod_sev_thin",
     "pct_ow_obese", "pct_ow", "pct_obese"],
).rename(columns={"N": "N_bmi"})

whr_state = pct_to_count_agg(
    whr_r, "Residence",
    ["pct_whr_normal", "pct_whr_high"],
).rename(columns={"N": "N_whr"})

ana_state = pct_to_count_agg(
    ana_a_r, "Residence",
    ["pct_anaemia_mild", "pct_anaemia_mod",
     "pct_anaemia_sev", "pct_anaemia_any"],
).rename(columns={"N": "N_ana"})

stunt_state = pct_to_count_agg(
    ch_r, "Residence",
    ["pct_stunting", "pct_wasting", "pct_child_ow"],
    n_col="N",
).rename(columns={"N": "N_ch"})

# ── Rurality ───────────────────────────────────────────────────────────────────
res_bmi = bmi_r[bmi_r["Var"] == "Residence"].copy()
res_bmi["N"] = pd.to_numeric(res_bmi["N"], errors="coerce")

state_N = res_bmi.groupby("state_key")["N"].sum().rename("N_total")
rural_N = (
    res_bmi[res_bmi["SubVar"].str.strip() == "Rural"]
    .groupby("state_key")["N"].sum()
    .rename("N_rural")
)
rurality = pd.concat([state_N, rural_N], axis=1).reset_index()
rurality["pct_rural"] = rurality["N_rural"] / rurality["N_total"] * 100

# ── Education ──────────────────────────────────────────────────────────────────
edu_bmi = bmi_r[bmi_r["Var"] == "Schooling"].copy()
edu_bmi["N"] = pd.to_numeric(edu_bmi["N"], errors="coerce").fillna(0)
edu_bmi["pct_thin"] = pd.to_numeric(edu_bmi["pct_thin"], errors="coerce").fillna(0)

edu_total = edu_bmi.groupby("state_key")["N"].sum()
edu_high = (
    edu_bmi[edu_bmi["SubVar"].str.strip() == "12 or more years complete"]
    .groupby("state_key")["N"].sum()
)
education = pd.DataFrame(
    {"N_edu_total": edu_total, "N_edu_12plus": edu_high}
).reset_index()
education["pct_edu_12plus"] = (
    education["N_edu_12plus"] / education["N_edu_total"] * 100
)

# ── Thin-Fat Index ─────────────────────────────────────────────────────────────
tfi = bmi_state[["state_key", "pct_thin", "N_bmi"]].merge(
    whr_state[["state_key", "pct_whr_high", "N_whr"]], on="state_key"
)
tfi["TFI"] = tfi["pct_thin"] * tfi["pct_whr_high"] / 100.0

# ── Master dataset ─────────────────────────────────────────────────────────────
master = (
    tfi
    .merge(
        ana_state[["state_key", "pct_anaemia_any", "pct_anaemia_sev", "N_ana"]],
        on="state_key", how="left",
    )
    .merge(
        stunt_state[["state_key", "pct_stunting", "pct_wasting", "N_ch"]],
        on="state_key", how="left",
    )
    .merge(rurality[["state_key", "pct_rural"]], on="state_key", how="left")
    .merge(education[["state_key", "pct_edu_12plus"]], on="state_key", how="left")
)

state_label = bmi_r[["state_key", "State"]].drop_duplicates()
master = master.merge(state_label, on="state_key", how="left")
master["State"] = master["State"].apply(
    lambda s: str(s).replace("_", " ").title()
)
master.set_index("State", inplace=True)
master.dropna(subset=["TFI", "pct_anaemia_any", "pct_stunting"], inplace=True)

print(f"\n{'='*70}")
print(f" MASTER DATASET — {len(master)} Indian States (NFHS-5)")
print(f"{'='*70}")
print(
    master[[
        "pct_thin", "pct_whr_high", "TFI",
        "pct_anaemia_any", "pct_stunting",
        "pct_rural", "pct_edu_12plus",
    ]].round(2).to_string()
)

# ══════════════════════════════════════════════════════════════════════════════
# 3. OBJECTIVE 2 — INTERGENERATIONAL GRADIENT
# ══════════════════════════════════════════════════════════════════════════════

mns_mask = ch_r["Var"].str.contains("nutritional", case=False, na=False)
mns      = ch_r[mns_mask].copy()


def clean_bmi_cat(s: str) -> str:
    s = str(s).strip()
    if "Underweight" in s or "18.5)" in s:
        return "Underweight\n(BMI <18.5)"
    elif "Normal" in s or "18.5-24.9" in s:
        return "Normal\n(BMI 18.5–24.9)"
    elif "Overweight" in s or "25.0" in s:
        return "Overweight/Obese\n(BMI ≥25.0)"
    return "Other"


mns["bmi_cat"]   = mns["SubVar"].apply(clean_bmi_cat)
mns["N_stunt"]   = pd.to_numeric(mns["N_stunt"],      errors="coerce").fillna(0)
mns["pct_stunt"] = pd.to_numeric(mns["pct_stunting"], errors="coerce").fillna(0)
mns["cnt_stunt"] = mns["pct_stunt"] * mns["N_stunt"] / 100.0

cat_agg = (
    mns.groupby("bmi_cat")
    .agg(N_total=("N_stunt", "sum"), cnt_stunt=("cnt_stunt", "sum"))
    .reset_index()
)
cat_agg["weighted_pct_stunting"] = (
    cat_agg["cnt_stunt"] / cat_agg["N_total"] * 100
)

cat_order = [
    "Underweight\n(BMI <18.5)",
    "Normal\n(BMI 18.5–24.9)",
    "Overweight/Obese\n(BMI ≥25.0)",
]
cat_agg["bmi_cat"] = pd.Categorical(
    cat_agg["bmi_cat"], categories=cat_order, ordered=True
)
cat_agg.sort_values("bmi_cat", inplace=True)

# ── State-level stunting per BMI category ─────────────────────────────────────
stunt_by_cat = (
    mns.groupby(["state_key", "bmi_cat"])
    .apply(
        lambda g: (
            g["cnt_stunt"].sum() / g["N_stunt"].sum() * 100
            if g["N_stunt"].sum() > 0
            else np.nan
        )
    )
    .reset_index(name="pct_stunt_state")
)
stunt_by_cat["bmi_cat"] = pd.Categorical(
    stunt_by_cat["bmi_cat"], categories=cat_order, ordered=True
)

uw = stunt_by_cat[stunt_by_cat["bmi_cat"] == cat_order[0]]["pct_stunt_state"].dropna()
nm = stunt_by_cat[stunt_by_cat["bmi_cat"] == cat_order[1]]["pct_stunt_state"].dropna()
ow = stunt_by_cat[stunt_by_cat["bmi_cat"] == cat_order[2]]["pct_stunt_state"].dropna()

kw_stat, kw_p = kruskal(uw, nm, ow)

# ── Pairwise Mann-Whitney U with Bonferroni correction ────────────────────────
pairs = [
    ("Underweight vs Normal",    uw, nm),
    ("Underweight vs Overweight", uw, ow),
    ("Normal vs Overweight",      nm, ow),
]
mw_results = []
for label, g1, g2 in pairs:
    u, p   = stats.mannwhitneyu(g1, g2, alternative="two-sided")
    p_bonf = min(p * 3, 1.0)
    r      = 1 - 2 * u / (len(g1) * len(g2))
    mw_results.append(
        {"Comparison": label, "U": u, "p_raw": p, "p_Bonf": p_bonf, "r": r}
    )
mw_df = pd.DataFrame(mw_results)

# ── Relative risk ──────────────────────────────────────────────────────────────
_cat_idx    = cat_agg.set_index("bmi_cat")["weighted_pct_stunting"]
pct_uw      = _cat_idx[cat_order[0]]
pct_nm      = _cat_idx[cat_order[1]]
pct_ow      = _cat_idx[cat_order[2]]
RR_nm_vs_uw = pct_nm / pct_uw
RR_ow_vs_uw = pct_ow / pct_uw

# ══════════════════════════════════════════════════════════════════════════════
# 4. OBJECTIVE 1 — DESCRIPTIVE STATISTICS
# ══════════════════════════════════════════════════════════════════════════════

indicators = {
    "Maternal Underweight\n(BMI <18.5) %":      "pct_thin",
    "Maternal Central Obesity\n(WHR ≥0.85) %":  "pct_whr_high",
    "Thin-Fat Index\n(ecological proxy)":        "TFI",
    "Maternal Anaemia\n(Any, <12 g/dL) %":       "pct_anaemia_any",
    "Child Stunting\n(HAZ <-2 SD) %":            "pct_stunting",
}

desc_stats = []
for label, col in indicators.items():
    d = master[col].dropna()
    _, sw_p = shapiro(d)
    desc_stats.append({
        "Indicator":       label.replace("\n", " "),
        "n":               int(d.count()),
        "Mean ± SD":       f"{d.mean():.1f} ± {d.std():.1f}",
        "Median":          f"{d.median():.1f}",
        "IQR":             f"{d.quantile(0.25):.1f}–{d.quantile(0.75):.1f}",
        "Min":             f"{d.min():.1f}",
        "Max":             f"{d.max():.1f}",
        "Shapiro-Wilk p":  f"{sw_p:.3f}",
        "Distribution":    "Normal" if sw_p > 0.05 else "Non-normal",
    })

desc_df = pd.DataFrame(desc_stats)

print(f"\n{'='*70}")
print(" OBJECTIVE 1 — DESCRIPTIVE STATISTICS")
print(f"{'='*70}")
print(desc_df.to_string(index=False))

# ── Pearson / Spearman correlation matrices ────────────────────────────────────
corr_cols   = list(indicators.values())
corr_labels = [l.replace("\n", " ") for l in indicators.keys()]
corr_data   = master[corr_cols].dropna()

pearson_r  = corr_data.corr(method="pearson")
spearman_r = corr_data.corr(method="spearman")

n_corr   = len(corr_data)
pearson_p = pd.DataFrame(
    np.ones_like(pearson_r.values),
    index=pearson_r.index, columns=pearson_r.columns,
)
for i, c1 in enumerate(corr_cols):
    for j, c2 in enumerate(corr_cols):
        if i != j:
            r, p = stats.pearsonr(corr_data[c1], corr_data[c2])
            pearson_p.iloc[i, j] = p

pearson_r.columns  = corr_labels
pearson_r.index    = corr_labels
spearman_r.columns = corr_labels
spearman_r.index   = corr_labels

print("\nPearson Correlation Matrix:")
print(pearson_r.round(3).to_string())

# ══════════════════════════════════════════════════════════════════════════════
# 5. OBJECTIVE 3 — OLS REGRESSION: PREDICTORS OF TFI
# ══════════════════════════════════════════════════════════════════════════════

reg_data = master[
    ["TFI", "pct_stunting", "pct_edu_12plus", "pct_rural", "pct_anaemia_any"]
].dropna().copy()
reg_data.columns = ["TFI", "stunting", "edu", "rural", "anaemia"]

X_vif  = sm.add_constant(reg_data[["stunting", "edu", "rural", "anaemia"]])
vif_df = pd.DataFrame({
    "Variable": X_vif.columns,
    "VIF":      [
        variance_inflation_factor(X_vif.values, i)
        for i in range(X_vif.shape[1])
    ],
})

ols_result = smf.ols(
    "TFI ~ stunting + edu + rural + anaemia", data=reg_data
).fit(cov_type="HC3")

print(f"\n{'='*70}")
print(" OBJECTIVE 3 — MULTIPLE REGRESSION: PREDICTORS OF TFI")
print(f"{'='*70}")
print(ols_result.summary())
print("\nVariance Inflation Factors:")
print(vif_df.round(3).to_string(index=False))

# ── Standardised betas ─────────────────────────────────────────────────────────
std_data = reg_data.copy()
for c in std_data.columns:
    std_data[c] = (std_data[c] - std_data[c].mean()) / std_data[c].std()

std_result = smf.ols(
    "TFI ~ stunting + edu + rural + anaemia", data=std_data
).fit(cov_type="HC3")
std_coefs = std_result.params.drop("Intercept").rename("Std. Beta")

reg_table = pd.DataFrame({
    "β (unstd.)":  ols_result.params.drop("Intercept").round(3),
    "95% CI Low":  ols_result.conf_int().drop("Intercept")[0].round(3),
    "95% CI High": ols_result.conf_int().drop("Intercept")[1].round(3),
    "Std. Beta":   std_coefs.round(3),
    "p-value":     ols_result.pvalues.drop("Intercept").round(4),
    "Sig.": ols_result.pvalues.drop("Intercept").apply(
        lambda p: "***" if p < 0.001 else
                  "**"  if p < 0.01  else
                  "*"   if p < 0.05  else
                  "†"   if p < 0.10  else "ns"
    ),
})

print("\nRegression Coefficients:")
print(reg_table.to_string())
print(
    f"\nR² = {ols_result.rsquared:.3f} | Adj. R² = {ols_result.rsquared_adj:.3f}"
    f" | F({int(ols_result.df_model)},{int(ols_result.df_resid)}) = "
    f"{ols_result.fvalue:.2f}, p = {ols_result.f_pvalue:.4f}"
)

# ══════════════════════════════════════════════════════════════════════════════
# 6. OBJECTIVE 4 — MEDIATION ANALYSIS (DOHaD HYPOTHESIS)
# ══════════════════════════════════════════════════════════════════════════════

med_data = master[["pct_anaemia_any", "pct_stunting", "TFI"]].dropna().copy()
med_data.columns = ["anaemia", "stunting", "TFI"]

# Standardise for mediation
med_std = med_data.copy()
for c in med_std.columns:
    med_std[c] = (med_std[c] - med_std[c].mean()) / med_std[c].std()

# Step 1 — Total effect c  (anaemia → TFI)
m1      = smf.ols("TFI ~ anaemia",          data=med_std).fit()
c_total = m1.params["anaemia"]
c_p     = m1.pvalues["anaemia"]

# Step 2 — Path a  (anaemia → stunting)
m2  = smf.ols("stunting ~ anaemia",         data=med_std).fit()
a   = m2.params["anaemia"]
a_p = m2.pvalues["anaemia"]

# Step 3 — Path b + direct effect c'  (stunting + anaemia → TFI)
m3        = smf.ols("TFI ~ anaemia + stunting", data=med_std).fit()
b         = m3.params["stunting"]
b_p       = m3.pvalues["stunting"]
c_prime   = m3.params["anaemia"]
c_prime_p = m3.pvalues["anaemia"]

# Indirect effect (ACME) and Sobel test
indirect    = a * b
se_indirect = np.sqrt(
    b**2 * m2.bse["anaemia"]**2 + a**2 * m3.bse["stunting"]**2
)
sobel_z       = indirect / se_indirect
sobel_p       = 2 * (1 - stats.norm.cdf(abs(sobel_z)))
prop_mediated = indirect / c_total if c_total != 0 else np.nan

pg_med = pg.mediation_analysis(
    data=med_std, x="anaemia", m="stunting", y="TFI",
    n_boot=5000, seed=42, alpha=0.05,
)

print(f"\n{'='*70}")
print(" OBJECTIVE 4 — MEDIATION ANALYSIS (DOHaD Hypothesis)")
print(f"{'='*70}")
print("\nBaron-Kenny Pathway Summary:")
print(f"  Total effect (c):    β = {c_total:.3f}, p = {c_p:.4f}")
print(f"  Path a (X→M):        β = {a:.3f},  p = {a_p:.4f}")
print(f"  Path b (M→Y|X):      β = {b:.3f},  p = {b_p:.4f}")
print(f"  Direct effect (c'):  β = {c_prime:.3f}, p = {c_prime_p:.4f}")
print(f"  Indirect (ACME):     β = {indirect:.3f}")
print(f"  Sobel Z = {sobel_z:.3f}, p = {sobel_p:.4f}")
print(f"  Proportion mediated: {prop_mediated*100:.1f}%")
print("\nPingouin Bootstrap Mediation (5000 iterations):")
print(pg_med.to_string(index=True))

bk_summary = pd.DataFrame({
    "Path": [
        "Total Effect (c)",
        "Path a: Anaemia→Stunting",
        "Path b: Stunting→TFI (adj.)",
        "Direct Effect (c')",
        "Indirect Effect (ACME)",
    ],
    "β (std.)": [
        round(c_total,  3), round(a,        3), round(b,       3),
        round(c_prime,  3), round(indirect, 3),
    ],
    "p-value": [
        round(c_p,      4), round(a_p,      4), round(b_p,     4),
        round(c_prime_p,4), round(sobel_p,  4),
    ],
    "Interpretation": [
        "Sig." if c_p      < 0.05 else "ns",
        "Sig." if a_p      < 0.05 else "ns",
        "Sig." if b_p      < 0.05 else "ns",
        "Sig." if c_prime_p< 0.05 else "ns",
        "Sig." if sobel_p  < 0.05 else "ns",
    ],
})

# ══════════════════════════════════════════════════════════════════════════════
# 7. FIGURES
# ══════════════════════════════════════════════════════════════════════════════

# ── Figure 1: Descriptive landscape ───────────────────────────────────────────
fig1, axes = plt.subplots(2, 3, figsize=(18, 11))
fig1.suptitle(
    "Figure 1 — Objective 1: Maternal-Child Nutritional Landscape\n"
    "across 29 Indian States (NFHS-5)",
    fontsize=14, fontweight="bold", y=1.01,
)
axes = axes.flatten()

panel_vars = [
    ("pct_thin",        "Maternal Underweight\n(BMI <18.5) %",     PAL[0]),
    ("pct_whr_high",    "Central Obesity\n(WHR ≥0.85) %",          PAL[1]),
    ("TFI",             "Thin-Fat Index\n(ecological proxy)",        PAL[2]),
    ("pct_anaemia_any", "Maternal Anaemia\n(Any, <12 g/dL) %",     PAL[3]),
    ("pct_stunting",    "Child Stunting\n(HAZ <-2 SD) %",           PAL[4]),
]

for ax, (col, title, color) in zip(axes[:5], panel_vars):
    d = master[col].sort_values()
    ax.barh(d.index, d.values, color=color, edgecolor="white", linewidth=0.4)
    ax.set_title(title, fontsize=11, fontweight="bold", pad=6)
    ax.set_xlabel(
        "Prevalence (%)" if "TFI" not in col else "TFI Score", fontsize=9
    )
    ax.tick_params(axis="y", labelsize=7.5)
    ax.axvline(
        d.median(), color="black", ls="--", lw=1.2, alpha=0.7,
        label=f"Median: {d.median():.1f}",
    )
    ax.legend(fontsize=8)
    for sp in ["top", "right"]:
        ax.spines[sp].set_visible(False)

# Panel 6 — Correlation heatmap
ax6         = axes[5]
short_labels = ["UW %", "CO %", "TFI", "Anaemia %", "Stunting %"]
hm_data      = pearson_r.copy()
hm_data.index   = short_labels
hm_data.columns = short_labels
sns.heatmap(
    hm_data, ax=ax6, annot=True, fmt=".2f",
    cmap="RdYlGn", center=0, vmin=-1, vmax=1,
    linewidths=0.5, annot_kws={"size": 9},
    cbar_kws={"shrink": 0.7, "label": "Pearson r"},
)
ax6.set_title(
    "Pearson Correlation Matrix\n(State-Level Indicators)",
    fontsize=11, fontweight="bold",
)
ax6.tick_params(axis="x", rotation=30, labelsize=9)
ax6.tick_params(axis="y", rotation=0,  labelsize=9)

plt.tight_layout()
fig1.savefig(
    os.path.join(OUT_DIR, "Figure1_Objective1_Descriptive.png"),
    dpi=180, bbox_inches="tight",
)
plt.close(fig1)
print("\n[Saved] Figure 1")

# ── Figure 2: Stunting by maternal BMI ────────────────────────────────────────
fig2, (ax_left, ax_right) = plt.subplots(1, 2, figsize=(14, 7))
fig2.suptitle(
    "Figure 2 — Objective 2: Intergenerational Biological Gradient\n"
    "Child Stunting Across Maternal BMI Categories (NFHS-5, 29 States)",
    fontsize=13, fontweight="bold",
)

bars = ax_left.bar(
    cat_agg["bmi_cat"].astype(str),
    cat_agg["weighted_pct_stunting"],
    color=[PAL[0], PAL[1], PAL[3]],
    edgecolor="white", width=0.55, zorder=2,
)
ax_left.set_ylabel("Weighted Prevalence of Child Stunting (%)", fontsize=11)
ax_left.set_xlabel("Maternal BMI Category", fontsize=11)
ax_left.set_title(
    "Country-Level Weighted Stunting Prevalence\nby Maternal BMI Category",
    fontsize=11,
)
for bar, val in zip(bars, cat_agg["weighted_pct_stunting"]):
    ax_left.text(
        bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.3,
        f"{val:.1f}%", ha="center", va="bottom",
        fontsize=10, fontweight="bold",
    )

ax_left.text(
    0.05, 0.92,
    f"RR (Normal vs Underweight) = {RR_nm_vs_uw:.2f}\n"
    f"RR (Overweight vs Underweight) = {RR_ow_vs_uw:.2f}",
    transform=ax_left.transAxes, fontsize=9.5,
    bbox=dict(boxstyle="round", facecolor="lightyellow", alpha=0.8),
)
ax_left.set_ylim(0, cat_agg["weighted_pct_stunting"].max() * 1.22)
ax_left.grid(axis="y", alpha=0.4)
for sp in ["top", "right"]:
    ax_left.spines[sp].set_visible(False)

stunt_plot = stunt_by_cat.copy()
stunt_plot["bmi_cat_short"] = stunt_plot["bmi_cat"].map({
    cat_order[0]: "Underweight\n<18.5",
    cat_order[1]: "Normal\n18.5–24.9",
    cat_order[2]: "Overweight\n≥25.0",
})
box_order = ["Underweight\n<18.5", "Normal\n18.5–24.9", "Overweight\n≥25.0"]
sns.boxplot(
    data=stunt_plot, x="bmi_cat_short", y="pct_stunt_state",
    palette=[PAL[0], PAL[1], PAL[3]], ax=ax_right,
    order=box_order, width=0.45, linewidth=1.2, fliersize=4,
)
sns.stripplot(
    data=stunt_plot, x="bmi_cat_short", y="pct_stunt_state",
    palette=[PAL[0], PAL[1], PAL[3]], ax=ax_right,
    order=box_order, size=5, jitter=True, alpha=0.6,
)
ax_right.set_ylabel("State-Level Stunting Prevalence (%)", fontsize=11)
ax_right.set_xlabel("Maternal BMI Category", fontsize=11)
ax_right.set_title(
    f"State-Level Distribution (n=29 States)\n"
    f"Kruskal-Wallis H = {kw_stat:.2f}, p = {kw_p:.4f}",
    fontsize=11,
)
ax_right.grid(axis="y", alpha=0.4)
for sp in ["top", "right"]:
    ax_right.spines[sp].set_visible(False)

plt.tight_layout()
fig2.savefig(
    os.path.join(OUT_DIR, "Figure2_Objective2_StuntingByBMI.png"),
    dpi=180, bbox_inches="tight",
)
plt.close(fig2)
print("[Saved] Figure 2")

# ── Figure 3: Scatter plots — predictors of TFI ───────────────────────────────
fig3, axes3 = plt.subplots(2, 2, figsize=(14, 11))
fig3.suptitle(
    "Figure 3 — Objective 3: Ecological Predictors of the Thin-Fat Index\n"
    "Scatter Plots with OLS Regression Lines (n = 29 States, NFHS-5)",
    fontsize=13, fontweight="bold",
)

scatter_vars = [
    ("pct_stunting",    "Child Stunting (%)",                  PAL[0]),
    ("pct_edu_12plus",  "Women with ≥12 Yrs\nSchooling (%)",   PAL[1]),
    ("pct_rural",       "Rural Residence (%)",                  PAL[2]),
    ("pct_anaemia_any", "Maternal Anaemia (%)",                 PAL[3]),
]

for ax, (col, xlabel, color) in zip(axes3.flatten(), scatter_vars):
    x = master[col].dropna()
    y = master.loc[x.index, "TFI"]
    mask = y.notna()
    x, y = x[mask], y[mask]
    r, p = stats.pearsonr(x, y)

    ax.scatter(x, y, color=color, edgecolors="white", s=60, alpha=0.85, zorder=3)
    m_fit = np.polyfit(x, y, 1)
    xline = np.linspace(x.min(), x.max(), 100)
    ax.plot(xline, np.polyval(m_fit, xline), color="black", lw=1.6, ls="--")

    combined = pd.DataFrame({"x": x, "y": y})
    extremes = (
        combined.nlargest(3, "y").index.tolist()
        + combined.nsmallest(2, "y").index.tolist()
    )
    for st in extremes:
        ax.annotate(
            str(st)[:12],
            (combined.loc[st, "x"], combined.loc[st, "y"]),
            textcoords="offset points", xytext=(5, 3),
            fontsize=7, color="dimgrey",
        )

    sig = "***" if p < 0.001 else "**" if p < 0.01 else "*" if p < 0.05 else "ns"
    ax.set_xlabel(xlabel, fontsize=10)
    ax.set_ylabel("Thin-Fat Index", fontsize=10)
    ax.set_title(f"r = {r:.3f} ({sig})", fontsize=10)
    for sp in ["top", "right"]:
        ax.spines[sp].set_visible(False)
    ax.grid(alpha=0.3)

plt.tight_layout()
fig3.savefig(
    os.path.join(OUT_DIR, "Figure3_Objective3_Regression_Scatter.png"),
    dpi=180, bbox_inches="tight",
)
plt.close(fig3)
print("[Saved] Figure 3")

# ── Figure 3b: Coefficient plot ────────────────────────────────────────────────
fig3b, ax_coef = plt.subplots(figsize=(9, 5))
coef_plot_data = reg_table.copy()
coef_plot_data.index = [
    "Child Stunting", "Education\n(≥12 yrs)", "Rurality", "Maternal\nAnaemia"
]
errors    = (coef_plot_data["95% CI High"] - coef_plot_data["95% CI Low"]) / 2
colors_cp = [
    PAL[0] if p < 0.05 else "grey"
    for p in ols_result.pvalues.drop("Intercept")
]

ax_coef.barh(
    coef_plot_data.index, coef_plot_data["Std. Beta"],
    xerr=errors.values, color=colors_cp,
    edgecolor="white", height=0.45,
    error_kw={"elinewidth": 1.6, "capsize": 4, "ecolor": "black"},
)
ax_coef.axvline(0, color="black", lw=1.2)
ax_coef.set_xlabel("Standardised Regression Coefficient (β)", fontsize=11)
ax_coef.set_title(
    f"OLS Regression: Predictors of Thin-Fat Index\n"
    f"R² = {ols_result.rsquared:.3f} (Adj. R² = {ols_result.rsquared_adj:.3f})"
    f" | HC3 robust SE",
    fontsize=11, fontweight="bold",
)
sig_patch   = mpatches.Patch(color=PAL[0], label="p < 0.05 (significant)")
insig_patch = mpatches.Patch(color="grey",  label="p ≥ 0.05 (not significant)")
ax_coef.legend(handles=[sig_patch, insig_patch], fontsize=9)
for sp in ["top", "right"]:
    ax_coef.spines[sp].set_visible(False)
ax_coef.grid(axis="x", alpha=0.4)

plt.tight_layout()
fig3b.savefig(
    os.path.join(OUT_DIR, "Figure3b_Objective3_CoefficientPlot.png"),
    dpi=180, bbox_inches="tight",
)
plt.close(fig3b)
print("[Saved] Figure 3b")

# ── Figure 4: Mediation pathway diagram ───────────────────────────────────────
fig4, (ax_path, ax_boot) = plt.subplots(1, 2, figsize=(16, 7))
fig4.suptitle(
    "Figure 4 — Objective 4: DOHaD Mediation Analysis\n"
    "Maternal Anaemia → Child Stunting (Mediator) → Thin-Fat Index",
    fontsize=13, fontweight="bold",
)

ax_path.set_xlim(0, 10)
ax_path.set_ylim(0, 6)
ax_path.axis("off")


def draw_box(ax, x, y, text, w=2.6, h=0.9, color="#2E4057"):
    rect = mpatches.FancyBboxPatch(
        (x - w / 2, y - h / 2), w, h,
        boxstyle="round,pad=0.12",
        facecolor=color, edgecolor="white", linewidth=1.5, alpha=0.88,
    )
    ax.add_patch(rect)
    ax.text(
        x, y, text, ha="center", va="center",
        fontsize=10, fontweight="bold", color="white", wrap=True,
    )


def draw_arrow(ax, x1, y1, x2, y2, label, color="black"):
    ax.annotate(
        "", xy=(x2, y2), xytext=(x1, y1),
        arrowprops=dict(
            arrowstyle="-|>", color=color, lw=1.8, mutation_scale=15
        ),
    )
    mx, my = (x1 + x2) / 2, (y1 + y2) / 2 + 0.22
    ax.text(
        mx, my, label, ha="center", va="bottom", fontsize=8.5,
        color=color, fontweight="bold",
        bbox=dict(boxstyle="round,pad=0.2", fc="white", ec=color, alpha=0.7),
    )


draw_box(ax_path, 1.5, 3, "Maternal\nAnaemia\n(X)",  color=PAL[3])
draw_box(ax_path, 5.0, 5, "Child\nStunting\n(M)",    color=PAL[0])
draw_box(ax_path, 8.5, 3, "Thin-Fat\nIndex\n(Y)",    color=PAL[2])

draw_arrow(ax_path, 2.4,  3.4,  4.0,  4.65,
           f"a = {a:.3f}{'*' if a_p < 0.05 else ''}",            color=PAL[3])
draw_arrow(ax_path, 5.95, 4.65, 7.65, 3.4,
           f"b = {b:.3f}{'*' if b_p < 0.05 else ''}",            color=PAL[0])
draw_arrow(ax_path, 2.85, 2.85, 7.15, 2.85,
           f"c' = {c_prime:.3f}{'*' if c_prime_p < 0.05 else ''} (direct)",
           color=PAL[2])

ax_path.text(
    5.0, 1.7,
    f"Indirect Effect (ACME) = {indirect:.3f}  |  "
    f"Sobel Z = {sobel_z:.2f}, p = {sobel_p:.4f}\n"
    f"Total Effect (c) = {c_total:.3f}, p = {c_p:.4f}  |  "
    f"Proportion Mediated = {prop_mediated*100:.1f}%",
    ha="center", va="center", fontsize=9.5,
    bbox=dict(
        boxstyle="round", facecolor="lightyellow",
        edgecolor="goldenrod", alpha=0.9,
    ),
)
ax_path.set_title(
    "Baron-Kenny Pathway Diagram\n(Standardised Coefficients)",
    fontsize=11, fontweight="bold",
)

ax_boot.set_title(
    "Pingouin Bootstrap Mediation Results\n(5000 iterations, 95% CI)",
    fontsize=11, fontweight="bold",
)

med_labels = pg_med.index.tolist()
coefs      = pg_med["coef"].values
ci_low     = pg_med["CI2.5"].values
ci_high    = pg_med["CI97.5"].values
p_vals = (
    pg_med["pval"].values
    if "pval" in pg_med.columns
    else np.ones(len(med_labels))
)

y_pos       = np.arange(len(med_labels))
bar_colors  = [PAL[0] if p < 0.05 else "lightgrey" for p in p_vals]
ax_boot.barh(y_pos, coefs, color=bar_colors, height=0.5, edgecolor="white", zorder=2)
ax_boot.errorbar(
    coefs, y_pos,
    xerr=[coefs - ci_low, ci_high - coefs],
    fmt="none", color="black", capsize=4, lw=1.5, zorder=3,
)
ax_boot.axvline(0, color="black", lw=1.2, ls="--")
ax_boot.set_yticks(y_pos)
ax_boot.set_yticklabels(med_labels, fontsize=9)
ax_boot.set_xlabel("Standardised Coefficient", fontsize=10)

sig_p  = mpatches.Patch(color=PAL[0],      label="p < 0.05")
nsig_p = mpatches.Patch(color="lightgrey", label="p ≥ 0.05")
ax_boot.legend(handles=[sig_p, nsig_p], fontsize=9, loc="lower right")
ax_boot.grid(axis="x", alpha=0.3)
for sp in ["top", "right"]:
    ax_boot.spines[sp].set_visible(False)

plt.tight_layout()
fig4.savefig(
    os.path.join(OUT_DIR, "Figure4_Objective4_Mediation.png"),
    dpi=180, bbox_inches="tight",
)
plt.close(fig4)
print("[Saved] Figure 4")

# ── Figure 5: Supplementary bubble chart ──────────────────────────────────────
fig5, ax5 = plt.subplots(figsize=(14, 9))
plot_d = master[["TFI", "pct_stunting", "pct_anaemia_any", "pct_rural", "N_bmi"]].dropna()

sc = ax5.scatter(
    plot_d["pct_stunting"], plot_d["TFI"],
    s=plot_d["pct_rural"] * 8,
    c=plot_d["pct_anaemia_any"],
    cmap=CMAP, edgecolors="dimgrey", linewidths=0.5,
    alpha=0.87, zorder=3,
)
cbar = plt.colorbar(sc, ax=ax5, shrink=0.6)
cbar.set_label("Maternal Anaemia Prevalence (%)", fontsize=10)

for idx, row in plot_d.iterrows():
    ax5.annotate(
        str(idx)[:12],
        (row["pct_stunting"], row["TFI"]),
        textcoords="offset points", xytext=(4, 3),
        fontsize=7.5, color="dimgrey",
    )

x_s, y_s = plot_d["pct_stunting"], plot_d["TFI"]
r5, p5   = stats.pearsonr(x_s, y_s)
mfit     = np.polyfit(x_s, y_s, 1)
xfit     = np.linspace(x_s.min(), x_s.max(), 100)
ax5.plot(xfit, np.polyval(mfit, xfit), "k--", lw=1.8, label=f"r = {r5:.3f}")

ax5.set_xlabel("Child Stunting Prevalence (%)", fontsize=12)
ax5.set_ylabel("Thin-Fat Index (Ecological Proxy)", fontsize=12)
ax5.set_title(
    "Supplementary Figure — Ecological Association: Child Stunting vs Thin-Fat Index\n"
    "Bubble size ∝ % rural population | Colour = Maternal Anaemia %",
    fontsize=12, fontweight="bold",
)
ax5.legend(fontsize=10)
ax5.grid(alpha=0.3)
for sp in ["top", "right"]:
    ax5.spines[sp].set_visible(False)

plt.tight_layout()
fig5.savefig(
    os.path.join(OUT_DIR, "Figure5_Supp_BubbleChart.png"),
    dpi=180, bbox_inches="tight",
)
plt.close(fig5)
print("[Saved] Figure 5")

# ══════════════════════════════════════════════════════════════════════════════
# 8. EXCEL REPORT
# ══════════════════════════════════════════════════════════════════════════════

def style_header_row(ws, row_idx: int, fill_hex: str = "2E4057") -> None:
    fill = PatternFill("solid", fgColor=fill_hex)
    font = Font(bold=True, color="FFFFFF", name="Calibri", size=10)
    for cell in ws[row_idx]:
        cell.fill      = fill
        cell.font      = font
        cell.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )


def auto_col_width(ws) -> None:
    for col in ws.columns:
        max_len = max(
            (len(str(c.value)) if c.value else 0) for c in col
        )
        ws.column_dimensions[
            get_column_letter(col[0].column)
        ].width = min(max_len + 4, 40)


def write_df_to_sheet(
    wb: Workbook, sheet_name: str, df: pd.DataFrame, fill_hex: str = "2E4057"
):
    ws = wb.create_sheet(title=sheet_name)
    ws.append(list(df.columns))
    for row in df.itertuples(index=False):
        ws.append(list(row))
    style_header_row(ws, 1, fill_hex)
    auto_col_width(ws)
    return ws


wb = Workbook()
wb.remove(wb.active)

# Sheet 1: Master dataset
master_export = master.reset_index()[[
    "State", "pct_thin", "pct_whr_high", "TFI",
    "pct_anaemia_any", "pct_anaemia_sev",
    "pct_stunting", "pct_wasting",
    "pct_rural", "pct_edu_12plus",
]].round(2)
master_export.columns = [
    "State", "Underweight %", "Central Obesity (WHR≥0.85) %", "Thin-Fat Index",
    "Any Anaemia %", "Severe Anaemia %",
    "Child Stunting %", "Child Wasting %",
    "Rural %", "Education (≥12 yrs) %",
]
write_df_to_sheet(wb, "Master_State_Dataset", master_export, "2E4057")

# Sheet 2: Descriptive statistics (Obj 1)
write_df_to_sheet(wb, "Obj1_Descriptive_Stats", desc_df, "048A81")

# Sheet 3: Stunting by maternal BMI (Obj 2)
cat_export = cat_agg.copy()
cat_export["bmi_cat"] = cat_export["bmi_cat"].astype(str)
cat_export.columns = [
    "Maternal BMI Category", "N Children",
    "Stunted Children", "Weighted Stunting %",
]
write_df_to_sheet(wb, "Obj2_Stunting_by_BMI", cat_export.round(2), "54C6EB")
write_df_to_sheet(wb, "Obj2_Mann_Whitney_Tests", mw_df.round(4), "54C6EB")

rr_ws = wb.create_sheet("Obj2_Relative_Risk")
rr_rows = [
    ["Comparison", "Reference", "Comparator", "RR", "Interpretation"],
    ["Normal vs Underweight", "Underweight", "Normal", round(RR_nm_vs_uw, 3),
     "Protective (<1)" if RR_nm_vs_uw < 1 else "Higher risk (>1)"],
    ["Overweight vs Underweight", "Underweight", "Overweight/Obese",
     round(RR_ow_vs_uw, 3),
     "Protective (<1)" if RR_ow_vs_uw < 1 else "Higher risk (>1)"],
]
for r in rr_rows:
    rr_ws.append(r)
style_header_row(rr_ws, 1, "54C6EB")
auto_col_width(rr_ws)

# Sheet 4: Regression (Obj 3)
reg_export = reg_table.reset_index()
reg_export.columns = ["Predictor", "β", "CI_Low", "CI_High", "Std_Beta", "p", "Sig"]
write_df_to_sheet(wb, "Obj3_Regression_Results", reg_export.round(4), "EF8C5B")
write_df_to_sheet(wb, "Obj3_VIF", vif_df.round(3), "EF8C5B")

# Sheet 5: Mediation (Obj 4)
write_df_to_sheet(wb, "Obj4_BK_Mediation", bk_summary, "E03C31")
pg_export = pg_med.reset_index().round(4)
write_df_to_sheet(wb, "Obj4_Bootstrap_Mediation", pg_export, "E03C31")

out_xlsx = os.path.join(OUT_DIR, "NFHS5_Malnutrition_Analysis_Results.xlsx")
wb.save(out_xlsx)
print(f"\n[Saved] Excel report → {out_xlsx}")

# ══════════════════════════════════════════════════════════════════════════════
# 9. SUMMARY
# ══════════════════════════════════════════════════════════════════════════════
print(f"""
{'='*70}
 ANALYSIS COMPLETE — SUMMARY OF KEY FINDINGS
{'='*70}

OBJECTIVE 1 (Descriptive Landscape):
  • Maternal underweight:       {master['pct_thin'].mean():.1f}% (±{master['pct_thin'].std():.1f}) — range {master['pct_thin'].min():.1f}–{master['pct_thin'].max():.1f}%
  • Central obesity (WHR≥0.85): {master['pct_whr_high'].mean():.1f}% (±{master['pct_whr_high'].std():.1f})
  • Thin-Fat Index (mean):       {master['TFI'].mean():.1f} (±{master['TFI'].std():.1f})
  • Maternal anaemia (any):     {master['pct_anaemia_any'].mean():.1f}% (±{master['pct_anaemia_any'].std():.1f})
  • Child stunting:             {master['pct_stunting'].mean():.1f}% (±{master['pct_stunting'].std():.1f})

OBJECTIVE 2 (Intergenerational Gradient):
  • Stunting: Underweight mothers = {pct_uw:.1f}%,
              Normal mothers      = {pct_nm:.1f}% (RR = {RR_nm_vs_uw:.2f}),
              Overweight mothers  = {pct_ow:.1f}% (RR = {RR_ow_vs_uw:.2f})
  • Kruskal-Wallis H = {kw_stat:.2f}, p = {kw_p:.4f}
  • {'Significant' if kw_p < 0.05 else 'No significant'} difference across maternal BMI categories

OBJECTIVE 3 (Predictors of TFI — OLS Regression):
  • R² = {ols_result.rsquared:.3f}, Adj. R² = {ols_result.rsquared_adj:.3f}
  • Child stunting:   β = {ols_result.params['stunting']:.3f}, p = {ols_result.pvalues['stunting']:.4f}
  • Education:        β = {ols_result.params['edu']:.3f}, p = {ols_result.pvalues['edu']:.4f}
  • Rurality:         β = {ols_result.params['rural']:.3f}, p = {ols_result.pvalues['rural']:.4f}
  • Maternal anaemia: β = {ols_result.params['anaemia']:.3f}, p = {ols_result.pvalues['anaemia']:.4f}

OBJECTIVE 4 (DOHaD Mediation):
  • Total effect (c):    β = {c_total:.3f}, p = {c_p:.4f}
  • Indirect effect:     β = {indirect:.3f}  (Sobel Z = {sobel_z:.2f}, p = {sobel_p:.4f})
  • Direct effect (c'):  β = {c_prime:.3f}, p = {c_prime_p:.4f}
  • Proportion mediated: {prop_mediated*100:.1f}%
  • {'✓ Significant' if sobel_p < 0.05 else '✗ Non-significant'} mediation (child stunting as mediator)

OUTPUT FILES:
  {out_xlsx}
  {os.path.join(OUT_DIR, 'Figure1_Objective1_Descriptive.png')}
  {os.path.join(OUT_DIR, 'Figure2_Objective2_StuntingByBMI.png')}
  {os.path.join(OUT_DIR, 'Figure3_Objective3_Regression_Scatter.png')}
  {os.path.join(OUT_DIR, 'Figure3b_Objective3_CoefficientPlot.png')}
  {os.path.join(OUT_DIR, 'Figure4_Objective4_Mediation.png')}
  {os.path.join(OUT_DIR, 'Figure5_Supp_BubbleChart.png')}
{'='*70}
""")
