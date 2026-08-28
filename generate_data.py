import pandas as pd
import random
products = [
    ("Wireless Earbuds", 49.99),
    ("Bluetooth Speaker", 39.99),
    ("Coffee Maker", 59.99),
    ("Running Shoes", 79.99),
    ("Office Chair", 149.99),
]
regions = ["North", "South", "East", "West"]
orders = []

for i in range(20):
    product_name, price = random.choice(products)
    region = random.choice(regions)
    quantity = random.randint(1, 5)
    revenue = round(price * quantity, 2)

    order = {
        "OrderID": i + 1,
        "Product": product_name,
        "Region": region,
        "Quantity": quantity,
        "Price": price,
        "Revenue": revenue,
    }
    orders.append(order)














df = pd.DataFrame(orders)
df.to_csv("sales_data.csv", index=False)
print("Done! Created sales_data.csv with", len(df), "orders")
