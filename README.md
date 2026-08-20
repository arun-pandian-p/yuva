# Week 2 Task: Exploratory Data Analysis (EDA) and Visualization Framework Design

This repository contains the comprehensive framework, automated Python analysis pipeline, publication-grade visualization charts, and formal documentation for **Week 2: Exploratory Data Analysis (EDA) and Visualization Framework Design**.

---

## 📁 Repository Contents

- 📄 **`EDA_and_Visualization_Framework.docx`** — Full technical framework document in Microsoft Word format, featuring custom typography, callouts, data tables, code snippets, and embedded figures.
- 📄 **`EDA_and_Visualization_Framework.doc`** — Formatted HTML-DOC deliverable with embedded base64 publication charts and CSS styling for 1-click opening in MS Word, LibreOffice, or web browsers.
- 🐍 **`eda_framework.py`** — Production-grade automated Python pipeline executing univariate, bivariate, multivariate, missing data diagnostics, and outlier detection.
- 📥 **`download_data.py`** — Automated terminal data ingestion script.
- 📊 **`telco_customer_churn.csv`** — Dataset used for practical demonstration (IBM Telco Customer Churn, 7,043 records).
- 🖼️ **`output_visuals/`** — High-resolution figures generated during EDA:
  - `fig1_missing_and_target.png`: Missing data diagnostics & target balance (26.5% churn).
  - `fig2_univariate_numeric.png`: Histograms + KDE + Box plots for tenure, monthly charges, and total charges.
  - `fig3_univariate_categorical.png`: Proportions of contracts, internet services, and payment methods.
  - `fig4_bivariate_continuous_target.png`: Violin plots of continuous metrics vs. churn status.
  - `fig5_correlation_heatmap.png`: Pearson correlation matrix across numeric and encoded binary features.
  - `fig6_categorical_churn_crosstabs.png`: Cross-tabulated churn proportions.
  - `fig7_multivariate_interaction.png`: Multi-attribute interaction scatter plot (tenure vs. monthly charges by contract and churn).

---

## 🚀 Quickstart & Reproduction

### 1. Install Dependencies
```bash
pip install pandas numpy matplotlib seaborn scipy scikit-learn python-docx requests
```

### 2. Download Dataset via Terminal
```bash
python download_data.py
```

### 3. Run Automated EDA Pipeline & Generate Visuals
```bash
python eda_framework.py
```

### 4. Rebuild DOC / DOCX Deliverables
```bash
python generate_doc.py
```

---

## 📊 Key Analytical Findings (Case Study)
1. **Missing Data Diagnostics**: 11 records in `TotalCharges` contained whitespace strings for new customers (`tenure = 0`). Imputed with median ($1,397.48).
2. **Target Distribution**: Churn rate is 26.5% (1,869 churners vs. 5,174 retained).
3. **Statistical Significance**:
   - **Welch's Independent $t$-Test** (Tenure vs. Churn): $t = -34.82, p = 1.20 \times 10^{-232}$ (Churners have median tenure of 10 months vs. 38 months for retained).
   - **Chi-Square Independence Test** (Contract vs. Churn): $\chi^2 = 1184.60, p = 5.86 \times 10^{-258}$ (Month-to-month contracts have 42.7% churn rate vs. 2.8% for 2-year contracts).
   - **Fiber Optic Service**: Highest revenue tier but 41.9% churn rate.
