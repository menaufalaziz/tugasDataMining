
import numpy as np
import pandas as pd
from itertools import combinations
from scipy.stats import chi2_contingency

df = pd.read_excel("output/datsert_setelah_cleaning.xlsx")

nominal = ["Survived", "Pclass", "Sex", "Embarked"]
numerik = ["Age", "Fare"]


def cramers_v(a, b):
    tabel = pd.crosstab(df[a], df[b])
    chi2, p, dof, _ = chi2_contingency(tabel, correction=False)
    v = np.sqrt(chi2 / (tabel.values.sum() * (min(tabel.shape) - 1)))
    return chi2, p, v


# Korelasi nominal: Chi-Square dan Cramer's V
print("KORELASI NOMINAL")
for a, b in combinations(nominal, 2):
    chi2, p, v = cramers_v(a, b)
    print(f"{a} - {b}: chi2={chi2:.3f}, p={p:.4g}, Cramer's V={v:.3f}")

# Korelasi numerik: Pearson
print("\nKORELASI NUMERIK (Pearson)")
print(df[numerik].corr(method="pearson").round(4))
