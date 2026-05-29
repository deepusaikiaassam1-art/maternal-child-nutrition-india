# maternal-child-nutrition-india
State-level analysis of intergenerational malnutrition across 29 Indian states using NFHS-5 data. Examines maternal BMI, anaemia, child stunting, and the Thin-Fat Index via descriptive statistics, OLS regression, and DOHaD mediation analysis.

# nfhs5-intergenerational-malnutrition

> State-level analysis of intergenerational malnutrition across 29 Indian states using NFHS-5 data. Examines maternal BMI, anaemia, child stunting, and the Thin-Fat Index via descriptive statistics, OLS regression, and DOHaD mediation analysis.

---

## Overview

This repository contains a fully reproducible Python analysis of maternal and child nutritional indicators across **29 Indian states**, using data from the **National Family Health Survey — Round 5 (NFHS-5, 2019–21)**.

The analysis is structured around four research objectives:

| # | Objective | Method |
|---|-----------|--------|
| 1 | Describe the maternal-child nutritional landscape | Descriptive statistics, Shapiro-Wilk normality, Pearson correlation |
| 2 | Quantify the intergenerational gradient in child stunting by maternal BMI category | Kruskal-Wallis, Mann-Whitney U (Bonferroni), Relative Risk |
| 3 | Identify ecological predictors of the Thin-Fat Index (TFI) | OLS regression with HC3 robust SE, VIF, standardised β |
| 4 | Test the DOHaD mediation hypothesis | Baron-Kenny steps, Sobel test, Pingouin bootstrap mediation (5000 iter.) |

---

## Background

The **Developmental Origins of Health and Disease (DOHaD)** hypothesis proposes that nutritional insults during critical developmental windows have lasting intergenerational consequences. India presents a paradox — simultaneous high prevalence of maternal underweight and central obesity (the **Thin-Fat phenotype**) — which this study operationalises as the **Thin-Fat Index (TFI)**:

```
TFI = (% maternal underweight) × (% central obesity, WHR ≥ 0.85) / 100
```

---

## Repository Structure

```
nfhs5-intergenerational-malnutrition/
│
├── intergenerational_malnutrition_analysis.py   # Main analysis script
├── README.md
│
├── data/                                        # Place NFHS-5 Excel files here
│   ├── BMI - Copy.xlsx
│   ├── WHR.xlsx
│   ├── Prevalence of anaemia in adults.xlsx
│   ├── Prevalence of anaemia in children.xlsx
│   ├── Nutritional status of children.xlsx
│   └── Antenatal care indicators.xlsx
│
└── outputs/                                     # Auto-generated on first run
    ├── NFHS5_Malnutrition_Analysis_Results.xlsx
    ├── Figure1_Objective1_Descriptive.png
    ├── Figure2_Objective2_StuntingByBMI.png
    ├── Figure3_Objective3_Regression_Scatter.png
    ├── Figure3b_Objective3_CoefficientPlot.png
    ├── Figure4_Objective4_Mediation.png
    └── Figure5_Supp_BubbleChart.png
```

---

## Requirements

Python 3.8 or higher is required. Install all dependencies with:

```bash
pip install pandas numpy matplotlib seaborn scipy statsmodels pingouin openpyxl
```

| Package | Purpose |
|---------|---------|
| `pandas`, `numpy` | Data wrangling and aggregation |
| `matplotlib`, `seaborn` | Figure generation |
| `scipy` | Pearson/Spearman correlation, Mann-Whitney U, Kruskal-Wallis, Shapiro-Wilk |
| `statsmodels` | OLS regression, VIF |
| `pingouin` | Bootstrap mediation analysis |
| `openpyxl` | Formatted Excel report output |

---

## Data

Data are sourced from **NFHS-5 (2019–21)**, publicly available from the International Institute for Population Sciences (IIPS):

> International Institute for Population Sciences (IIPS) and ICF. (2021). *National Family Health Survey (NFHS-5), India, 2019–21*. Mumbai: IIPS.

Download the state-level factsheets from: https://rchiips.org/nfhs/nfhs-5Reports/NFHS-5_INDIA_REPORT.pdf

Place the six required Excel files in the `data/` folder (or in the same directory as the script) before running.

---

## Usage

```bash
# Clone the repository
git clone https://github.com/<your-username>/nfhs5-intergenerational-malnutrition.git
cd nfhs5-intergenerational-malnutrition

# Install dependencies
pip install pandas numpy matplotlib seaborn scipy statsmodels pingouin openpyxl

# Add NFHS-5 Excel files to the data/ folder, then run
python intergenerational_malnutrition_analysis.py
```

All outputs are saved automatically to the `outputs/` folder.

---

## Outputs

### Figures

| File | Description |
|------|-------------|
| `Figure1_Objective1_Descriptive.png` | Horizontal bar charts for 5 nutritional indicators + Pearson correlation heatmap |
| `Figure2_Objective2_StuntingByBMI.png` | Child stunting by maternal BMI category — bar chart (national) + box-strip plot (state-level) |
| `Figure3_Objective3_Regression_Scatter.png` | Scatter plots of TFI vs each predictor with OLS regression lines |
| `Figure3b_Objective3_CoefficientPlot.png` | Standardised coefficient plot with 95% CI |
| `Figure4_Objective4_Mediation.png` | Baron-Kenny pathway diagram + Pingouin bootstrap forest plot |
| `Figure5_Supp_BubbleChart.png` | Supplementary bubble chart: child stunting vs TFI, coloured by anaemia, sized by rurality |

### Excel Report

`NFHS5_Malnutrition_Analysis_Results.xlsx` contains seven formatted sheets:

- `Master_State_Dataset` — state-level dataset for all indicators
- `Obj1_Descriptive_Stats` — summary statistics with normality tests
- `Obj2_Stunting_by_BMI` — weighted stunting prevalence by maternal BMI category
- `Obj2_Mann_Whitney_Tests` — pairwise tests with Bonferroni correction
- `Obj2_Relative_Risk` — relative risk estimates (underweight as reference)
- `Obj3_Regression_Results` — OLS coefficients, CI, standardised β, p-values
- `Obj3_VIF` — variance inflation factors
- `Obj4_BK_Mediation` — Baron-Kenny pathway summary
- `Obj4_Bootstrap_Mediation` — Pingouin bootstrap results (5000 iterations)

---

## Key Variables

| Variable | Description | Source Sheet |
|----------|-------------|-------------|
| `pct_thin` | % women with BMI < 18.5 | BMI |
| `pct_whr_high` | % women with WHR ≥ 0.85 | WHR |
| `TFI` | Thin-Fat Index (ecological proxy) | Computed |
| `pct_anaemia_any` | % women with any anaemia (Hb < 12 g/dL) | Anaemia (adults) |
| `pct_stunting` | % children with HAZ < −2 SD | Child nutrition |
| `pct_wasting` | % children with WHZ < −2 SD | Child nutrition |
| `pct_rural` | % rural population (residence-weighted) | BMI |
| `pct_edu_12plus` | % women with ≥12 years of schooling | BMI |

---

## Citation

If you use this code or analysis in your research, please cite:

```
Saikia, D. (2024). Intergenerational malnutrition analysis across Indian states
using NFHS-5 data. GitHub: https://github.com/<your-username>/nfhs5-intergenerational-malnutrition
```

---

## License

This project is licensed under the MIT License. See `LICENSE` for details.

---

## Contact

**Dr. Deepjyoti Saikia**
For queries regarding the analysis or data, please open a GitHub Issue.
