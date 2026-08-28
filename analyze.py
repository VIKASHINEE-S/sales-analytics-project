import pandas as pd
df = pd.read_csv("sales_data.csv")
print(df)
revenue_by_region = df.groupby("Region")["Revenue"].sum()
print(revenue_by_region)
revenue_by_product = df.groupby("Product")["Revenue"].sum()
print(revenue_by_product)

import pandas as pd
import matplotlib.pyplot as plt

df = pd.read_csv("sales_data.csv")
print(df)

revenue_by_region = df.groupby("Region")["Revenue"].sum()
print(revenue_by_region)

revenue_by_product = df.groupby("Product")["Revenue"].sum()
print(revenue_by_product)

revenue_by_product.plot(kind="bar")
plt.title("Revenue by Product")
plt.ylabel("Revenue ($)")
plt.tight_layout()
plt.savefig("revenue_chart.png")
print("Chart saved as revenue_chart.png")