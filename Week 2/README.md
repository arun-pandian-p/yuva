# Week 2 Task: Exploratory Data Analysis (EDA) and Visualization Framework Design

This folder contains the comprehensive framework, automated Python analysis pipeline, publication-grade visualization charts, and formal documentation for **Week 2: Exploratory Data Analysis (EDA) and Visualization Framework Design**.

---

## 📁 Folder Contents

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

## 🚀 Quickstart & Execution

```bash
# 1. Install Dependencies
pip install pandas numpy matplotlib seaborn scipy scikit-learn python-docx requests

# 2. Download Dataset via Terminal
python download_data.py

# 3. Run Automated EDA Pipeline & Generate Visuals
python eda_framework.py

# 4. Rebuild DOC / DOCX Deliverables
python generate_doc.py
```
