"""
CSCI218 Group Project: Dry Bean Dataset Classification
=======================================================
This script performs classification of dry bean types using multiple
machine learning algorithms on the UCI Dry Bean Dataset.

Dataset: https://archive.ics.uci.edu/ml/datasets/Dry+Bean+Dataset
7 bean types: SEKER, BARBUNYA, BOMBAY, CALI, HOROZ, SIRA, DERMASON
16 features derived from grain images (shape, size, etc.)

Algorithms implemented:
1. K-Nearest Neighbours (KNN)
2. Decision Tree
3. Random Forest
4. Support Vector Machine (SVM)
5. Logistic Regression
6. Naive Bayes (Gaussian)
"""

import os
import warnings
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import seaborn as sns

from sklearn.model_selection import train_test_split, cross_val_score, StratifiedKFold
from sklearn.preprocessing import StandardScaler, LabelEncoder
from sklearn.metrics import (
    accuracy_score, precision_score, recall_score, f1_score,
    classification_report, confusion_matrix
)
from sklearn.neighbors import KNeighborsClassifier
from sklearn.tree import DecisionTreeClassifier
from sklearn.ensemble import RandomForestClassifier
from sklearn.svm import SVC
from sklearn.linear_model import LogisticRegression
from sklearn.naive_bayes import GaussianNB

warnings.filterwarnings('ignore')

# ============================================================
# Configuration
# ============================================================
OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")
os.makedirs(OUTPUT_DIR, exist_ok=True)

RANDOM_STATE = 42
TEST_SIZE = 0.2

# ============================================================
# 1. Load Dataset
# ============================================================
print("=" * 60)
print("CSCI218 Group Project: Dry Bean Classification")
print("=" * 60)

# Try loading from ucimlrepo first, fall back to local CSV
try:
    from ucimlrepo import fetch_ucirepo
    print("\n[1] Loading Dry Bean dataset from UCI ML Repository...")
    dataset = fetch_ucirepo(id=602)
    X = dataset.data.features
    y = dataset.data.targets.values.ravel()
    print(f"    Dataset loaded successfully via ucimlrepo.")
except Exception:
    # Fallback: try loading from a local CSV
    local_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Dry_Bean_Dataset.csv")
    if os.path.exists(local_path):
        print(f"\n[1] Loading Dry Bean dataset from local file: {local_path}")
        df = pd.read_csv(local_path)
        X = df.iloc[:, :-1]
        y = df.iloc[:, -1].values
    else:
        # Last resort: download directly
        print("\n[1] Downloading Dry Bean dataset from UCI archive...")
        url = "https://archive.ics.uci.edu/ml/machine-learning-databases/00602/DryBeanDataset.zip"
        import urllib.request
        import zipfile
        import io
        response = urllib.request.urlopen(url)
        z = zipfile.ZipFile(io.BytesIO(response.read()))
        for name in z.namelist():
            if name.endswith('.xlsx') or name.endswith('.csv'):
                z.extract(name, os.path.dirname(os.path.abspath(__file__)))
        # Try to find the extracted file
        for f in os.listdir(os.path.dirname(os.path.abspath(__file__))):
            if 'dry' in f.lower() and (f.endswith('.xlsx') or f.endswith('.csv')):
                if f.endswith('.xlsx'):
                    df = pd.read_excel(os.path.join(os.path.dirname(os.path.abspath(__file__)), f))
                else:
                    df = pd.read_csv(os.path.join(os.path.dirname(os.path.abspath(__file__)), f))
                X = df.iloc[:, :-1]
                y = df.iloc[:, -1].values
                break

print(f"    Samples: {X.shape[0]}, Features: {X.shape[1]}")
print(f"    Bean classes: {np.unique(y)}")
print(f"    Feature names: {list(X.columns)}")

# ============================================================
# 2. Exploratory Data Analysis & Visualizations
# ============================================================
print("\n[2] Exploratory Data Analysis...")

# 2a. Class distribution
class_counts = pd.Series(y).value_counts().sort_index()
print(f"    Class distribution:\n{class_counts.to_string()}")

fig, ax = plt.subplots(figsize=(10, 6))
colors = sns.color_palette("husl", len(class_counts))
bars = ax.bar(class_counts.index, class_counts.values, color=colors, edgecolor='black')
ax.set_xlabel("Bean Type", fontsize=13)
ax.set_ylabel("Number of Samples", fontsize=13)
ax.set_title("Class Distribution in Dry Bean Dataset", fontsize=15, fontweight='bold')
for bar, val in zip(bars, class_counts.values):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 50,
            str(val), ha='center', va='bottom', fontsize=10, fontweight='bold')
plt.xticks(rotation=30, ha='right')
plt.tight_layout()
plt.savefig(os.path.join(OUTPUT_DIR, "class_distribution.png"), dpi=150)
plt.close()
print("    Saved: class_distribution.png")

# 2b. Feature correlation heatmap
fig, ax = plt.subplots(figsize=(14, 11))
corr = X.corr()
mask = np.triu(np.ones_like(corr, dtype=bool))
sns.heatmap(corr, mask=mask, annot=True, fmt=".2f", cmap="RdBu_r",
            center=0, linewidths=0.5, ax=ax, annot_kws={"size": 7})
ax.set_title("Feature Correlation Heatmap", fontsize=15, fontweight='bold')
plt.tight_layout()
plt.savefig(os.path.join(OUTPUT_DIR, "correlation_heatmap.png"), dpi=150)
plt.close()
print("    Saved: correlation_heatmap.png")

# 2c. Feature distributions by class (selected features)
key_features = ['Area', 'Perimeter', 'roundness', 'Compactness']
available_features = [f for f in key_features if f in X.columns]
if len(available_features) < 4:
    available_features = list(X.columns[:4])

fig, axes = plt.subplots(2, 2, figsize=(14, 10))
for idx, feat in enumerate(available_features[:4]):
    ax = axes[idx // 2][idx % 2]
    for cls in np.unique(y):
        subset = X[pd.Series(y) == cls][feat]
        ax.hist(subset, bins=30, alpha=0.5, label=cls, density=True)
    ax.set_title(f"Distribution of {feat}", fontsize=12, fontweight='bold')
    ax.set_xlabel(feat)
    ax.set_ylabel("Density")
    ax.legend(fontsize=7, loc='upper right')
plt.suptitle("Feature Distributions by Bean Type", fontsize=15, fontweight='bold')
plt.tight_layout()
plt.savefig(os.path.join(OUTPUT_DIR, "feature_distributions.png"), dpi=150)
plt.close()
print("    Saved: feature_distributions.png")

# 2d. Boxplots for key features
fig, axes = plt.subplots(2, 2, figsize=(14, 10))
df_plot = X.copy()
df_plot['Class'] = y
for idx, feat in enumerate(available_features[:4]):
    ax = axes[idx // 2][idx % 2]
    sns.boxplot(data=df_plot, x='Class', y=feat, ax=ax, palette="husl")
    ax.set_title(f"Boxplot of {feat}", fontsize=12, fontweight='bold')
    ax.tick_params(axis='x', rotation=30)
plt.suptitle("Feature Boxplots by Bean Type", fontsize=15, fontweight='bold')
plt.tight_layout()
plt.savefig(os.path.join(OUTPUT_DIR, "feature_boxplots.png"), dpi=150)
plt.close()
print("    Saved: feature_boxplots.png")

# ============================================================
# 3. Data Preprocessing
# ============================================================
print("\n[3] Data Preprocessing...")

# Encode target labels
le = LabelEncoder()
y_encoded = le.fit_transform(y)
class_names = le.classes_
print(f"    Encoded classes: {dict(zip(class_names, range(len(class_names))))}")

# Check for missing values
missing = X.isnull().sum().sum()
print(f"    Missing values: {missing}")
if missing > 0:
    X = X.fillna(X.median())
    print("    Filled missing values with median.")

# Train/test split
X_train, X_test, y_train, y_test = train_test_split(
    X, y_encoded, test_size=TEST_SIZE, random_state=RANDOM_STATE, stratify=y_encoded
)
print(f"    Training set: {X_train.shape[0]} samples")
print(f"    Test set:     {X_test.shape[0]} samples")

# Feature scaling
scaler = StandardScaler()
X_train_scaled = scaler.fit_transform(X_train)
X_test_scaled = scaler.transform(X_test)
print("    Features standardized (zero mean, unit variance).")

# ============================================================
# 4. Model Training & Evaluation
# ============================================================
print("\n[4] Training and Evaluating Models...")
print("-" * 60)

models = {
    "K-Nearest Neighbours": KNeighborsClassifier(n_neighbors=5),
    "Random Forest": RandomForestClassifier(n_estimators=100, random_state=RANDOM_STATE, n_jobs=-1),
    "SVM (RBF)": SVC(kernel='rbf', C=10, gamma='scale', random_state=RANDOM_STATE),
}

results = {}
cv = StratifiedKFold(n_splits=5, shuffle=True, random_state=RANDOM_STATE)

for name, model in models.items():
    print(f"\n  Training: {name}...")

    # Cross-validation on training set
    cv_scores = cross_val_score(model, X_train_scaled, y_train, cv=cv, scoring='accuracy')

    # Fit on full training set, predict on test set
    model.fit(X_train_scaled, y_train)
    y_pred = model.predict(X_test_scaled)

    acc = accuracy_score(y_test, y_pred)
    prec = precision_score(y_test, y_pred, average='weighted')
    rec = recall_score(y_test, y_pred, average='weighted')
    f1 = f1_score(y_test, y_pred, average='weighted')

    results[name] = {
        'cv_mean': cv_scores.mean(),
        'cv_std': cv_scores.std(),
        'accuracy': acc,
        'precision': prec,
        'recall': rec,
        'f1': f1,
        'y_pred': y_pred,
        'model': model,
    }

    print(f"    5-Fold CV Accuracy:  {cv_scores.mean():.4f} (+/- {cv_scores.std():.4f})")
    print(f"    Test Accuracy:       {acc:.4f}")
    print(f"    Test Precision (W):  {prec:.4f}")
    print(f"    Test Recall (W):     {rec:.4f}")
    print(f"    Test F1-Score (W):   {f1:.4f}")

# ============================================================
# 5. Results Comparison & Visualization
# ============================================================
print("\n[5] Generating Result Visualizations...")

# 5a. Model comparison bar chart
model_names = list(results.keys())
metrics_data = {
    'Accuracy': [results[m]['accuracy'] for m in model_names],
    'Precision': [results[m]['precision'] for m in model_names],
    'Recall': [results[m]['recall'] for m in model_names],
    'F1-Score': [results[m]['f1'] for m in model_names],
}

fig, ax = plt.subplots(figsize=(14, 7))
x = np.arange(len(model_names))
width = 0.18
multiplier = 0
colors_metrics = ['#2196F3', '#4CAF50', '#FF9800', '#F44336']

for (metric, values), color in zip(metrics_data.items(), colors_metrics):
    offset = width * multiplier
    bars = ax.bar(x + offset, values, width, label=metric, color=color, edgecolor='black', linewidth=0.5)
    for bar, val in zip(bars, values):
        ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.003,
                f"{val:.3f}", ha='center', va='bottom', fontsize=7, fontweight='bold')
    multiplier += 1

ax.set_xlabel("Model", fontsize=13)
ax.set_ylabel("Score", fontsize=13)
ax.set_title("Model Performance Comparison on Dry Bean Dataset", fontsize=15, fontweight='bold')
ax.set_xticks(x + width * 1.5)
ax.set_xticklabels(model_names, rotation=20, ha='right', fontsize=10)
ax.legend(loc='lower right', fontsize=10)
ax.set_ylim(0.5, 1.05)
ax.grid(axis='y', alpha=0.3)
plt.tight_layout()
plt.savefig(os.path.join(OUTPUT_DIR, "model_comparison.png"), dpi=150)
plt.close()
print("    Saved: model_comparison.png")

# 5b. Cross-validation comparison
fig, ax = plt.subplots(figsize=(12, 6))
cv_means = [results[m]['cv_mean'] for m in model_names]
cv_stds = [results[m]['cv_std'] for m in model_names]
bars = ax.bar(model_names, cv_means, yerr=cv_stds, capsize=5,
              color=sns.color_palette("viridis", len(model_names)), edgecolor='black')
for bar, mean, std in zip(bars, cv_means, cv_stds):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + std + 0.005,
            f"{mean:.4f}", ha='center', va='bottom', fontsize=9, fontweight='bold')
ax.set_xlabel("Model", fontsize=13)
ax.set_ylabel("5-Fold CV Accuracy", fontsize=13)
ax.set_title("Cross-Validation Accuracy Comparison", fontsize=15, fontweight='bold')
ax.set_ylim(0.7, 1.0)
ax.grid(axis='y', alpha=0.3)
plt.xticks(rotation=20, ha='right')
plt.tight_layout()
plt.savefig(os.path.join(OUTPUT_DIR, "cv_comparison.png"), dpi=150)
plt.close()
print("    Saved: cv_comparison.png")

# 5c. Confusion matrices for all models (3 models: 1 row x 3 columns)
fig, axes = plt.subplots(1, 3, figsize=(18, 6))
for idx, (name, res) in enumerate(results.items()):
    ax = axes[idx]
    cm = confusion_matrix(y_test, res['y_pred'])
    sns.heatmap(cm, annot=True, fmt='d', cmap='Blues', ax=ax,
                xticklabels=class_names, yticklabels=class_names)
    ax.set_title(f"{name}\n(Acc: {res['accuracy']:.4f})", fontsize=11, fontweight='bold')
    ax.set_xlabel("Predicted")
    ax.set_ylabel("Actual")
    ax.tick_params(axis='x', rotation=30)
    ax.tick_params(axis='y', rotation=0)
plt.suptitle("Confusion Matrices for All Models", fontsize=16, fontweight='bold')
plt.tight_layout()
plt.savefig(os.path.join(OUTPUT_DIR, "confusion_matrices.png"), dpi=150)
plt.close()
print("    Saved: confusion_matrices.png")

# 5d. Best model detailed classification report
best_model_name = max(results, key=lambda m: results[m]['accuracy'])
best_res = results[best_model_name]
print(f"\n{'=' * 60}")
print(f"BEST MODEL: {best_model_name} (Accuracy: {best_res['accuracy']:.4f})")
print(f"{'=' * 60}")
report = classification_report(y_test, best_res['y_pred'], target_names=class_names)
print(report)

# Save classification report
with open(os.path.join(OUTPUT_DIR, "classification_report.txt"), 'w') as f:
    f.write(f"Best Model: {best_model_name}\n")
    f.write(f"Test Accuracy: {best_res['accuracy']:.4f}\n\n")
    f.write(report)
print("    Saved: classification_report.txt")

# 5e. Feature importance (from Random Forest)
rf_model = results['Random Forest']['model']
importances = rf_model.feature_importances_
feat_imp = pd.Series(importances, index=X.columns).sort_values(ascending=True)

fig, ax = plt.subplots(figsize=(10, 8))
feat_imp.plot(kind='barh', ax=ax, color=sns.color_palette("viridis", len(feat_imp)), edgecolor='black')
ax.set_xlabel("Feature Importance", fontsize=13)
ax.set_title("Random Forest Feature Importance", fontsize=15, fontweight='bold')
ax.grid(axis='x', alpha=0.3)
plt.tight_layout()
plt.savefig(os.path.join(OUTPUT_DIR, "feature_importance.png"), dpi=150)
plt.close()
print("    Saved: feature_importance.png")

# ============================================================
# 6. Summary Table
# ============================================================
print("\n[6] Results Summary Table")
print("-" * 80)
print(f"{'Model':<25} {'CV Acc':>10} {'Test Acc':>10} {'Precision':>10} {'Recall':>10} {'F1':>10}")
print("-" * 80)
for name in model_names:
    r = results[name]
    print(f"{name:<25} {r['cv_mean']:>10.4f} {r['accuracy']:>10.4f} {r['precision']:>10.4f} {r['recall']:>10.4f} {r['f1']:>10.4f}")
print("-" * 80)

# Save summary to CSV
summary_df = pd.DataFrame({
    'Model': model_names,
    'CV_Accuracy_Mean': [results[m]['cv_mean'] for m in model_names],
    'CV_Accuracy_Std': [results[m]['cv_std'] for m in model_names],
    'Test_Accuracy': [results[m]['accuracy'] for m in model_names],
    'Precision_Weighted': [results[m]['precision'] for m in model_names],
    'Recall_Weighted': [results[m]['recall'] for m in model_names],
    'F1_Weighted': [results[m]['f1'] for m in model_names],
})
summary_df.to_csv(os.path.join(OUTPUT_DIR, "results_summary.csv"), index=False)
print("\n    Saved: results_summary.csv")

print(f"\nAll outputs saved to: {OUTPUT_DIR}")
print("Done!")
