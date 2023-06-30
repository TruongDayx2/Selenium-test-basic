import pandas as pd

df = pd.read_excel("input.xlsx")
for index, row in df.iterrows():
    print((row[6]))
# print(df)