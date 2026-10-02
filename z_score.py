import pandas as pd

df = pd.read_excel("output/datsert_setelah_cleaning.xlsx")
 
numerik = ["Age", "Fare"]
 
# Z-Score: (x - rata-rata) / simpangan baku
for kolom in numerik:
    df[kolom + "_zscore"] = ((df[kolom] - df[kolom].mean()) / df[kolom].std()).round(4)
 
print(df.head(10))
print("\nPemeriksaan (mean harus 0, std harus 1):")
print(df[["Age_zscore", "Fare_zscore"]].agg(["mean", "std"]).round(4))
 
df.to_excel("hasil_normalisasi.xlsx", index=False)
