import pandas as pd
from pathlib import Path

# Create data folder
data_folder = Path("data")
data_folder.mkdir(exist_ok=True)

# Sample sales dataset
sales_data = {
    "Order_ID": [
        1001, 1002, 1003, 1004, 1005,
        1006, 1007, 1008, 1009, 1010,
        1011, 1012, 1013, 1014, 1015
    ],

    "Date": [
        "2026-01-01", "2026-01-02", "2026-01-03",
        "2026-01-04", "2026-01-05", "2026-01-06",
        "2026-01-07", "2026-01-08", "2026-01-09",
        "2026-01-10", "2026-01-11", "2026-01-12",
        "2026-01-13", "2026-01-14", "2026-01-15"
    ],

    "Product": [
        "Laptop", "Mouse", "Keyboard", "Monitor", "Laptop",
        "Headphones", "Mouse", "Monitor", "Keyboard", "Laptop",
        "Headphones", "Mouse", "Monitor", "Keyboard", "Laptop"
    ],

    "Category": [
        "Electronics", "Accessories", "Accessories",
        "Electronics", "Electronics", "Accessories",
        "Accessories", "Electronics", "Accessories",
        "Electronics", "Accessories", "Accessories",
        "Electronics", "Accessories", "Electronics"
    ],

    "Quantity": [
        2, 5, 3, 2, 1, 4, 6, 1, 5, 2, 3, 4, 2, 3, 1
    ],

    "Unit_Price": [
        50000, 500, 1200, 15000, 50000,
        2500, 500, 15000, 1200, 50000,
        2500, 500, 15000, 1200, 50000
    ],

    "Region": [
        "South", "North", "East", "West", "South",
        "North", "East", "West", "South", "North",
        "East", "West", "South", "North", "East"
    ]
}

# Create DataFrame
df = pd.DataFrame(sales_data)

# Calculate total sales
df["Total_Sales"] = df["Quantity"] * df["Unit_Price"]

# Convert date column
df["Date"] = pd.to_datetime(df["Date"])

# Save Excel file
file_path = data_folder / "sales_data.xlsx"

df.to_excel(file_path, index=False)

print("=" * 45)
print("     SAMPLE SALES DATA GENERATED")
print("=" * 45)

print(f"Total Records: {len(df)}")
print(f"Total Sales: Rs. {df['Total_Sales'].sum():,.2f}")
print(f"Excel File: {file_path}")
print("Status: Successfully Created!")
print("=" * 45)