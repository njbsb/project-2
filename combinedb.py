import pandas as pd
import os

directory = os.getcwd()
zhpla_path = os.path.join(directory, "database/input/", "dummyzhpla.xlsx")
zpdev_path = os.path.join(directory, "database/input/", "dummyzpdev.xlsx")

df_zh = pd.read_excel(zhpla_path, header=None)
df_zp = pd.read_excel(zpdev_path, header=None)
df
print(df_zh.head())
print("\n")
print(df_zp.head())
