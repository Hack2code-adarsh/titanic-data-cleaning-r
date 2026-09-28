# Data Cleaning & Preliminary Analysis with R — Titanic Dataset

> A complete, reproducible data-cleaning and exploratory analysis pipeline built entirely with **base R** — no external packages required.

![R](https://img.shields.io/badge/R-4.6.1-276DC3?logo=r&logoColor=white)
![Status](https://img.shields.io/badge/status-complete-success)
![Dataset](https://img.shields.io/badge/dataset-Titanic-891%20rows-blue)
![License](https://img.shields.io/badge/license-MIT-lightgrey)
![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20R-blueviolet)

---

## Overview

This project delivers a **structured data-cleaning, preprocessing, and preliminary analysis** workflow on the classic **Titanic passenger survival dataset** using **R**. The pipeline walks through every essential stage of a real-world data-science task — from raw data exploration to missing-value imputation, outlier handling, normalization, categorical encoding, and exploratory data analysis (EDA) — and packages the entire process into a **professional Word report** with embedded code, console outputs, and visualizations.

The project was designed to run on a **base-R-only** environment, meaning it executes cleanly in any standard R installation without needing `install.packages()`. Every figure and statistic in the final report is generated directly by running the script, making the analysis **fully reproducible**.

---

## Dataset

| Property | Value |
|---|---|
| **Source** | [Kaggle Titanic Dataset](https://www.kaggle.com/c/titanic/data) (publicly mirrored) |
| **Records** | 891 passengers |
| **Variables** | 12 (mixed categorical + numerical) |
| **Target** | `Survived` (0 = Died, 1 = Survived) |
| **Missingness** | Age (~20%), Cabin (~77%), Embarked (~0.2%) |

### Variable Dictionary

| Variable | Type | Role | Description |
|---|---|---|---|
| `PassengerId` | integer | identifier | Unique passenger ID (dropped) |
| `Survived` | integer → factor | **target** | 0 = Died, 1 = Survived |
| `Pclass` | integer → factor | feature | Ticket class: 1st, 2nd, 3rd |
| `Name` | character | identifier | Passenger name (dropped) |
| `Sex` | character | feature | male / female |
| `Age` | numeric | feature | Age in years (177 missing) |
| `SibSp` | integer | feature | # siblings / spouses aboard |
| `Parch` | integer | feature | # parents / children aboard |
| `Ticket` | character | identifier | Ticket number (dropped) |
| `Fare` | numeric | feature | Passenger fare (right-skewed) |
| `Cabin` | character | feature | Cabin number (687 missing → flag) |
| `Embarked` | character | feature | Port: C, Q, S (2 missing) |

---

## Methodology

The analysis follows a **seven-stage pipeline**, each stage producing saved artifacts (text outputs + PNG plots) for the final report.

```
┌─────────────────────┐    ┌─────────────────────┐    ┌─────────────────────┐
│ 1. Initial          │ → │ 2. Missing-Value    │ → │ 3. Data Cleaning    │
│    Exploration      │    │    Analysis         │    │    & Imputation     │
│ dim/str/summary     │    │ colSums(is.na)      │    │ group-wise median   │
└─────────────────────┘    └─────────────────────┘    └─────────────────────┘
                                                                      ↓
┌─────────────────────┐    ┌─────────────────────┐    ┌─────────────────────┐
│ 6. Categorical      │ ← │ 5. Normalization    │ ← │ 4. Outlier          │
│    Encoding         │    │    & Scaling        │    │    Detection (IQR)  │
│ label + one-hot     │    │ Min-Max & Z-score   │    │ winsorization       │
└─────────────────────┘    └─────────────────────┘    └─────────────────────┘
        ↓
┌─────────────────────┐
│ 7. Exploratory      │
│    Data Analysis    │  →  Final Word Report (.docx)
│ tables, cor, plots  │
└─────────────────────┘
```

### Stage 1 — Initial Exploration
Loaded raw CSV with `na.strings` to surface genuine missingness. Used `dim()`, `names()`, `str()`, `head()`, `tail()`, and `summary()` to profile structure, types, and distributions.

### Stage 2 — Missing-Value Analysis
Quantified missingness per variable and overall (8.1% of all cells). Strategy set by variable:
- **Age** (~20%) → group-wise median imputation by `Pclass × Sex`
- **Cabin** (~77%) → convert to binary `HasCabin` flag, drop raw column
- **Embarked** (~0.2%) → mode imputation

### Stage 3 — Data Cleaning
Trimmed whitespace, dropped identifier columns (`PassengerId`, `Name`, `Ticket`), converted coded integers to labelled factors, imputed remaining NAs, engineered `HasCabin`, and verified true duplicates on the raw identifier-inclusive data. **Result: 891 rows, 0 missing values.**

### Stage 4 — Outlier Detection (IQR)
Applied the 1.5×IQR rule to all numeric variables. Age/SibSp/Parch outliers retained as plausible real values; **Fare winsorized at the 99th percentile** to curb extreme ticket prices without losing valid records.

### Stage 5 — Normalization
Generated both **Min-Max** (range [0,1]) and **Z-score** (mean 0, sd 1) scaled versions of `Age` and `Fare`, giving downstream modelling flexibility.

### Stage 6 — Categorical Encoding
- **Label encoding** for binary categoricals: `Sex` → `Sex_code`, `HasCabin` → `HasCabin_code`
- **One-hot encoding** for nominal categoricals with >2 levels: `Embarked` → 3 dummy columns, `Pclass` → 3 dummy columns

### Stage 7 — Exploratory Data Analysis
Computed grouped survival rates, a full correlation matrix, and per-feature correlations with `Survived`. Generated seven publication-ready plots (barplots, histograms, boxplots, scatter, correlation heatmap).

---

## Key Findings

| Insight | Detail |
|---|---|
| **Overall survival** | 38.4% (342 of 891 passengers) |
| **Sex is the strongest predictor** | 74.2% of females survived vs 18.9% of males (corr = -0.54) |
| **Class gradient** | Class 1: 63.0% · Class 2: 47.3% · Class 3: 24.2% |
| **Cherbourg advantage** | Passengers embarking at Cherbourg (C) had the highest survival (55.4%), largely due to Class 1 proportion |
| **Fare ↔ Survival** | Positive correlation (0.27); higher fare → higher survival odds |
| **Age ↔ Survival** | Weakly correlated (-0.06); age alone is not a strong predictor |
| **HasCabin ↔ Survival** | Positive correlation (0.32); cabin presence is a proxy for upper classes |

---

## Project Structure

```
titanic-data-cleaning-r/
├── titanic_analysis.R                    # Reproducible R analysis script (base R)
├── generate_docx.ps1                      # PowerShell script to build the .docx report
├── Week1_Titanic_R_DataCleaning_Analysis.docx  # Final professional Word report
├── .gitignore
│
├── data/
│   ├── titanic.csv                        # Raw dataset (891 × 12)
│   └── titanic_cleaned.csv                # Cleaned, scaled & encoded dataset
│
├── outputs/                               # Captured R console output (per stage)
│   ├── 01_exploration.txt
│   ├── 02_missing.txt
│   ├── 03_cleaning.txt
│   ├── 04_outliers.txt
│   ├── 05_normalization.txt
│   ├── 06_encoding.txt
│   └── 07_eda.txt
│
└── plots/                                 # Generated visualizations (PNG, 110 DPI)
    ├── 02_missing.png                     # Missing-value barplot
    ├── 04_outliers.png                    # Age & Fare boxplots
    ├── 05_normalization.png               # Post-normalization histograms
    ├── 07_correlation.png                 # Correlation matrix heatmap
    ├── 08_histograms.png                  # Age & Fare by survival
    ├── 09_barplots.png                    # Survival by Sex & Pclass
    └── 10_age_fare_scatter.png            # Age vs Fare scatter
```

---

## Tech Stack

| Tool | Purpose |
|---|---|
| **R 4.6.1** | Data analysis, cleaning, EDA, visualization (base R only) |
| **PowerShell 5.1** | Word COM automation to generate the `.docx` report |
| **Git + GitHub CLI** | Version control and repository management |
| **Microsoft Word** | Final report rendering |

---

## How to Reproduce

### Prerequisites
- **R** (≥ 4.0) installed and `Rscript` available on PATH
- The raw dataset (`data/titanic.csv`) is already included in the repo

### Run the Analysis

```bash
# 1. Clone the repository
git clone https://github.com/Hack2code-adarsh/titanic-data-cleaning-r.git
cd titanic-data-cleaning-r

# 2. Run the R analysis script
Rscript titanic_analysis.R
```

This will:
- Regenerate all files in `outputs/` (console text)
- Regenerate all plots in `plots/` (PNG)
- Produce the cleaned dataset at `data/titanic_cleaned.csv`

### (Optional) Regenerate the Word Report

```powershell
# Requires Microsoft Word installed (uses Word COM automation)
powershell -ExecutionPolicy Bypass -File generate_docx.ps1
```

---

## Report

The complete analysis is documented in **`Week1_Titanic_R_DataCleaning_Analysis.docx`**, a professionally formatted Word report containing:

- Title page & auto-generated table of contents
- Dataset overview with a variable dictionary table
- R code snippets with syntax-styled formatting
- Real captured R console output for every stage
- Seven embedded figures with captions
- Conclusions and recommended next steps
- Full R script in the appendix

---

## What I Learned

- Designing a **reproducible, package-free R pipeline** that runs anywhere
- Choosing the right imputation strategy per variable (group-wise median vs mode vs flag)
- Avoiding a subtle data-leakage pitfall: checking duplicates on identifier-inclusive raw data, not on the cleaned subset
- Balancing outlier handling — **winsorizing** rather than deleting to preserve valid minority-class records
- Automating professional report generation via **Word COM** (PowerShell), embedding live code outputs and figures
- End-to-end Git/GitHub workflow including repo creation and remote management via the `gh` CLI

---

## Future Work

- Engineer a `FamilySize` feature (`SibSp + Parch`) and an `IsAlone` flag
- Extract a `Title` feature from passenger names (`Mr`, `Mrs`, `Miss`, `Master`)
- Perform stratified train/test splitting
- Build baseline predictive models (logistic regression, decision trees)
- Evaluate with cross-validation and ROC curves

---

## Author

**Adarsh Upadhyay** — [Hack2code-adarsh](https://github.com/Hack2code-adarsh)

---

## License

This project is licensed under the MIT License — see the dataset's original terms for data-usage restrictions.

---

<p align="center">
  <i>Built with R · Powered by reproducible data science</i>
</p>
