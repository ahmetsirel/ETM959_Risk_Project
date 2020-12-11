import pandas as pd
import numpy as np
import os

path = os.path.join(os.getcwd(), "data")
files = [file for file in os.listdir(path) if file.endswith("csv")]

data = pd.DataFrame()
for file in files:
    file_data = pd.read_csv(os.path.join(path,file))
    file_data["Şirket"] = file[: file.find(" ")]
    data = data.append(file_data)


# Process Columns
data["Tarih"] = pd.to_datetime(data["Tarih"], dayfirst=True)
data[["Şimdi", "Açılış", "Yüksek", "Düşük"]] = data[
    ["Şimdi", "Açılış", "Yüksek", "Düşük"]
].applymap(lambda x: round(float(str(x).replace(",", ".")), 3))
data["Fark %"] = data["Fark %"].apply(
    lambda x: round(float(str(x).replace(",", ".").replace("%", "")) / 100, 3)
)

data = data[["Şirket", "Tarih", "Şimdi", "Açılış", "Yüksek", "Düşük", "Fark %"]]
data = data.sort_values(["Şirket", "Tarih"])

data_agg = data.groupby("Şirket").agg({"Fark %": ["mean", "var"]})


# Cov
cov_data = pd.DataFrame()
for company in data["Şirket"].unique():
    current = data[data["Şirket"] == company]["Fark %"].rename(company)
    cov_data = pd.concat([cov_data, current], axis=1)
cov = cov_data.cov()

writer = pd.ExcelWriter("risk.xlsx", engine="xlsxwriter")
data.to_excel(writer, sheet_name="Data")
data_agg.to_excel(writer, sheet_name="Data Agg")
cov_data.to_excel(writer, sheet_name="Cov Data")
cov.to_excel(writer, sheet_name="Covariance")
writer.save()

print("Success")
