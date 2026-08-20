"""
Comprehensive Exploratory Data Analysis (EDA) Pipeline
Demonstration on the Telco Customer Churn Dataset.
Generates publication-quality charts and comprehensive statistical summaries.
Optimized for high visual fidelity with compressed file footprints.
"""

import os
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from scipy import stats
from PIL import Image

# Configure styling
plt.style.use('seaborn-v0_8-whitegrid' if 'seaborn-v0_8-whitegrid' in plt.style.available else 'default')
plt.rcParams['font.family'] = 'sans-serif'
plt.rcParams['font.size'] = 10
plt.rcParams['axes.titlesize'] = 11
plt.rcParams['axes.labelsize'] = 10
plt.rcParams['figure.titlesize'] = 13

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(SCRIPT_DIR, "output_visuals")
os.makedirs(OUTPUT_DIR, exist_ok=True)

def save_optimized_fig(fig, filepath, dpi=130):
    """Saves figure with PNG compression to ensure small file size without visual loss."""
    fig.savefig(filepath, bbox_inches='tight', dpi=dpi)
    # Further optimize with PIL
    try:
        img = Image.open(filepath)
        img = img.convert('RGB')
        img.save(filepath, format='PNG', optimize=True)
    except Exception as e:
        pass

def run_eda_pipeline():
    print("=" * 70)
    print("STARTING END-TO-END EXPLORATORY DATA ANALYSIS (EDA) FRAMEWORK")
    print("=" * 70)

    # -------------------------------------------------------------
    # PHASE 1: DATA INGESTION & STRUCTURAL INSPECTION
    # -------------------------------------------------------------
    csv_path = os.path.join(SCRIPT_DIR, "telco_customer_churn.csv")
    df = pd.read_csv(csv_path)
    print(f"\n[1] Dataset Loaded Successfully.")
    print(f"    - Rows: {df.shape[0]:,}")
    print(f"    - Columns: {df.shape[1]}")
    print(f"    - Memory Usage: {df.memory_usage().sum() / 1024:.2f} KB")

    # Data Hygiene: Clean TotalCharges (contains whitespace strings)
    df['TotalCharges'] = pd.to_numeric(df['TotalCharges'].str.strip(), errors='coerce')
    
    # -------------------------------------------------------------
    # PHASE 2: MISSING VALUE DIAGNOSTICS
    # -------------------------------------------------------------
    missing_counts = df.isnull().sum()
    missing_pct = (missing_counts / len(df)) * 100
    missing_df = pd.DataFrame({
        'Feature': df.columns,
        'Missing_Count': missing_counts.values,
        'Missing_Pct': missing_pct.values,
        'Dtype': df.dtypes.astype(str).values
    }).sort_values(by='Missing_Pct', ascending=False)
    
    print("\n[2] Missing Data Diagnostics:")
    print(missing_df[missing_df['Missing_Count'] > 0])

    # Impute TotalCharges with median for subsequent modeling readiness
    median_total = df['TotalCharges'].median()
    df_clean = df.copy()
    df_clean['TotalCharges'] = df_clean['TotalCharges'].fillna(median_total)

    # -------------------------------------------------------------
    # PHASE 3: VARIABLE TAXONOMY CLASSIFICATION
    # -------------------------------------------------------------
    numeric_cols = ['tenure', 'MonthlyCharges', 'TotalCharges']
    categorical_cols = [c for c in df.columns if c not in numeric_cols and c != 'customerID']
    print(f"\n[3] Variable Classification:")
    print(f"    - Numeric Features ({len(numeric_cols)}): {numeric_cols}")
    print(f"    - Categorical Features ({len(categorical_cols)}): {categorical_cols}")

    # Summary Statistics for Numeric
    num_summary = df_clean[numeric_cols].describe().T
    num_summary['skewness'] = df_clean[numeric_cols].skew()
    num_summary['kurtosis'] = df_clean[numeric_cols].kurtosis()
    print("\n[4] Numerical Summary Statistics:")
    print(num_summary[['mean', 'std', 'min', '50%', 'max', 'skewness', 'kurtosis']])

    # -------------------------------------------------------------
    # PHASE 4: GENERATING PUBLICATION-QUALITY FIGURES (OPTIMIZED)
    # -------------------------------------------------------------

    # Figure 1: Missing Data & Target Class Balance
    fig, axes = plt.subplots(1, 2, figsize=(11, 4.2))
    
    # Subplot 1A: Missing Value Proportions
    non_zero_missing = missing_df[missing_df['Missing_Count'] > 0]
    if len(non_zero_missing) == 0:
        axes[0].text(0.5, 0.5, "No Missing Values Detected", ha='center', va='center', fontsize=12)
    else:
        sns.barplot(x='Feature', y='Missing_Count', data=non_zero_missing, ax=axes[0], color='#1B365D')
        axes[0].set_title("Missing Value Counts by Feature", fontweight='bold', color='#1B365D')
        axes[0].set_ylabel("Count of Nulls")
        for p in axes[0].patches:
            axes[0].annotate(f"{int(p.get_height())} ({p.get_height()/len(df)*100:.2f}%)",
                             (p.get_x() + p.get_width() / 2., p.get_height()),
                             ha='center', va='bottom', fontsize=9, color='#1B365D')
    
    # Subplot 1B: Target Distribution (Churn Rate)
    churn_counts = df['Churn'].value_counts()
    colors = ['#4B6B94', '#E05A47']
    axes[1].pie(churn_counts, labels=churn_counts.index, autopct='%1.1f%%',
                startangle=140, colors=colors, explode=(0, 0.08),
                wedgeprops={'edgecolor': 'white', 'linewidth': 2})
    axes[1].set_title("Target Distribution: Churn vs Retained", fontweight='bold', color='#1B365D')
    plt.tight_layout()
    fig1_path = os.path.join(OUTPUT_DIR, "fig1_missing_and_target.png")
    save_optimized_fig(fig, fig1_path)
    plt.close(fig)
    print(f"[+] Saved {fig1_path} ({os.path.getsize(fig1_path)/1024:.1f} KB)")

    # Figure 2: Univariate Numerical Distributions
    fig, axes = plt.subplots(len(numeric_cols), 2, figsize=(11, 8.5))
    for i, col in enumerate(numeric_cols):
        sns.histplot(df_clean[col], kde=True, ax=axes[i, 0], color='#1B365D', bins=25, alpha=0.6)
        mean_val = df_clean[col].mean()
        median_val = df_clean[col].median()
        axes[i, 0].axvline(mean_val, color='#E05A47', linestyle='--', label=f'Mean ({mean_val:.1f})')
        axes[i, 0].axvline(median_val, color='#2ECC71', linestyle='-', label=f'Median ({median_val:.1f})')
        axes[i, 0].set_title(f"Distribution of {col} (Skew: {df_clean[col].skew():.2f})", fontweight='bold')
        axes[i, 0].legend(loc='upper right', fontsize=8)
        
        sns.boxplot(x=df_clean[col], ax=axes[i, 1], color='#85C1E9', flierprops={'markerfacecolor':'#E05A47', 'markersize':3})
        axes[i, 1].set_title(f"Boxplot & Spread of {col}", fontweight='bold')
    
    plt.tight_layout()
    fig2_path = os.path.join(OUTPUT_DIR, "fig2_univariate_numeric.png")
    save_optimized_fig(fig, fig2_path)
    plt.close(fig)
    print(f"[+] Saved {fig2_path} ({os.path.getsize(fig2_path)/1024:.1f} KB)")

    # Figure 3: Categorical Feature Frequency Distributions
    key_cat_cols = ['Contract', 'InternetService', 'PaymentMethod', 'TechSupport']
    fig, axes = plt.subplots(2, 2, figsize=(12, 7.5))
    axes = axes.flatten()
    for i, col in enumerate(key_cat_cols):
        order = df[col].value_counts().index
        sns.countplot(x=col, hue=col, data=df, order=order, ax=axes[i], palette='crest', legend=False)
        axes[i].set_title(f"Univariate Breakdown: {col}", fontweight='bold', color='#1B365D')
        axes[i].tick_params(axis='x', rotation=15)
        total_len = float(len(df))
        for p in axes[i].patches:
            height = p.get_height()
            if height > 0:
                axes[i].annotate(f'{height/total_len:.1%}', (p.get_x() + p.get_width() / 2., height + 30),
                                 ha='center', va='bottom', fontsize=8)
    plt.tight_layout()
    fig3_path = os.path.join(OUTPUT_DIR, "fig3_univariate_categorical.png")
    save_optimized_fig(fig, fig3_path)
    plt.close(fig)
    print(f"[+] Saved {fig3_path} ({os.path.getsize(fig3_path)/1024:.1f} KB)")

    # Figure 4: Bivariate Relationships
    fig, axes = plt.subplots(1, 3, figsize=(13, 4.2))
    for i, col in enumerate(numeric_cols):
        sns.violinplot(x='Churn', y=col, hue='Churn', data=df_clean, ax=axes[i], palette=['#4B6B94', '#E05A47'], split=False, inner='quartile', legend=False)
        axes[i].set_title(f"{col} by Churn Status", fontweight='bold', color='#1B365D')
    plt.tight_layout()
    fig4_path = os.path.join(OUTPUT_DIR, "fig4_bivariate_continuous_target.png")
    save_optimized_fig(fig, fig4_path)
    plt.close(fig)
    print(f"[+] Saved {fig4_path} ({os.path.getsize(fig4_path)/1024:.1f} KB)")

    # Figure 5: Correlation Matrix Heatmap
    corr_df = df_clean[numeric_cols].copy()
    corr_df['SeniorCitizen'] = df_clean['SeniorCitizen']
    corr_df['Churn_Binary'] = (df_clean['Churn'] == 'Yes').astype(int)
    corr_df['Partner'] = (df_clean['Partner'] == 'Yes').astype(int)
    corr_df['Dependents'] = (df_clean['Dependents'] == 'Yes').astype(int)
    corr_df['PaperlessBilling'] = (df_clean['PaperlessBilling'] == 'Yes').astype(int)
    
    corr_matrix = corr_df.corr()
    fig, ax = plt.subplots(figsize=(8, 6))
    mask = np.triu(np.ones_like(corr_matrix, dtype=bool))
    sns.heatmap(corr_matrix, mask=mask, annot=True, fmt='.2f', cmap='coolwarm',
                vmin=-1, vmax=1, square=True, linewidths=0.8, cbar_kws={'shrink': 0.8}, ax=ax)
    ax.set_title("Pearson Correlation Heatmap", fontweight='bold', color='#1B365D')
    plt.tight_layout()
    fig5_path = os.path.join(OUTPUT_DIR, "fig5_correlation_heatmap.png")
    save_optimized_fig(fig, fig5_path)
    plt.close(fig)
    print(f"[+] Saved {fig5_path} ({os.path.getsize(fig5_path)/1024:.1f} KB)")

    # Figure 6: Categorical Bivariate Proportions
    fig, axes = plt.subplots(1, 2, figsize=(12, 4.2))
    
    contract_churn = pd.crosstab(df_clean['Contract'], df_clean['Churn'], normalize='index') * 100
    contract_churn.plot(kind='bar', stacked=True, color=['#4B6B94', '#E05A47'], ax=axes[0], alpha=0.85)
    axes[0].set_title("Churn Rate by Contract Type (%)", fontweight='bold', color='#1B365D')
    axes[0].set_ylabel("Percentage (%)")
    axes[0].tick_params(axis='x', rotation=0)
    for p in axes[0].containers:
        axes[0].bar_label(p, fmt='%.1f%%', label_type='center', color='white', fontweight='bold', fontsize=8.5)

    internet_churn = pd.crosstab(df_clean['InternetService'], df_clean['Churn'], normalize='index') * 100
    internet_churn.plot(kind='bar', stacked=True, color=['#4B6B94', '#E05A47'], ax=axes[1], alpha=0.85)
    axes[1].set_title("Churn Rate by Internet Service (%)", fontweight='bold', color='#1B365D')
    axes[1].set_ylabel("Percentage (%)")
    axes[1].tick_params(axis='x', rotation=0)
    for p in axes[1].containers:
        axes[1].bar_label(p, fmt='%.1f%%', label_type='center', color='white', fontweight='bold', fontsize=8.5)

    plt.tight_layout()
    fig6_path = os.path.join(OUTPUT_DIR, "fig6_categorical_churn_crosstabs.png")
    save_optimized_fig(fig, fig6_path)
    plt.close(fig)
    print(f"[+] Saved {fig6_path} ({os.path.getsize(fig6_path)/1024:.1f} KB)")

    # Figure 7: Multivariate Interaction (Sampled to avoid massive scatter plot overhead)
    # Downsample points slightly for clean render and tiny file footprint
    sample_df = df_clean.sample(n=min(2500, len(df_clean)), random_state=42)
    fig, ax = plt.subplots(figsize=(9, 5))
    sns.scatterplot(
        data=sample_df,
        x='tenure',
        y='MonthlyCharges',
        hue='Churn',
        style='Contract',
        palette={'No': '#4B6B94', 'Yes': '#E05A47'},
        alpha=0.65,
        s=30,
        ax=ax
    )
    ax.set_title("Multivariate Interaction: Tenure vs. Monthly Charges", fontweight='bold', color='#1B365D')
    ax.set_xlabel("Tenure (Months)")
    ax.set_ylabel("Monthly Charges ($)")
    plt.legend(bbox_to_anchor=(1.02, 1), loc='upper left', fontsize=8.5)
    plt.tight_layout()
    fig7_path = os.path.join(OUTPUT_DIR, "fig7_multivariate_interaction.png")
    save_optimized_fig(fig, fig7_path)
    plt.close(fig)
    print(f"[+] Saved {fig7_path} ({os.path.getsize(fig7_path)/1024:.1f} KB)")

    print("\n[+] Optimized EDA Pipeline execution completed!")

if __name__ == "__main__":
    run_eda_pipeline()
