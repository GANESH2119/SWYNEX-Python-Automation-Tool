import pandas as pd
from pathlib import Path
import json


class SalesAnalyzer:

    def __init__(self, input_file):
        self.input_file = Path(input_file)
        self.output_folder = Path("output")
        self.output_folder.mkdir(exist_ok=True)

    def load_data(self):
        """Load cleaned sales data."""

        try:
            if not self.input_file.exists():
                raise FileNotFoundError(
                    "Cleaned Excel file not found!"
                )

            df = pd.read_excel(self.input_file)

            required = [
                "Order_ID", "Date", "Product",
                "Category", "Quantity",
                "Unit_Price", "Region", "Total_Sales"
            ]

            missing = [
                col for col in required
                if col not in df.columns
            ]

            if missing:
                raise ValueError(
                    f"Missing columns: {missing}"
                )

            print("Sales data loaded successfully!")

            return df

        except Exception as error:
            print(f"Loading Error: {error}")
            return None

    def analyze_sales(self, df):
        """Perform complete sales analysis."""

        total_revenue = float(df["Total_Sales"].sum())
        total_orders = int(len(df))
        total_quantity = int(df["Quantity"].sum())

        average_order_value = (
            total_revenue / total_orders
            if total_orders > 0 else 0
        )

        # Product-wise sales
        product_sales = (
            df.groupby("Product")["Total_Sales"]
            .sum()
            .sort_values(ascending=False)
        )

        # Category-wise sales
        category_sales = (
            df.groupby("Category")["Total_Sales"]
            .sum()
            .sort_values(ascending=False)
        )

        # Region-wise sales
        region_sales = (
            df.groupby("Region")["Total_Sales"]
            .sum()
            .sort_values(ascending=False)
        )

        # Monthly sales
        df["Date"] = pd.to_datetime(df["Date"])

        monthly_sales = (
            df.groupby(df["Date"].dt.strftime("%Y-%m"))
            ["Total_Sales"]
            .sum()
        )

        top_product = product_sales.index[0]

        results = {
            "total_revenue": total_revenue,
            "total_orders": total_orders,
            "total_quantity": total_quantity,
            "average_order_value": round(
                average_order_value, 2
            ),
            "top_product": top_product,
            "product_sales": product_sales.to_dict(),
            "category_sales": category_sales.to_dict(),
            "region_sales": region_sales.to_dict(),
            "monthly_sales": monthly_sales.to_dict()
        }

        return results

    def display_results(self, results):
        """Display analysis in terminal."""

        print("\n" + "=" * 55)
        print("           SALES ANALYSIS REPORT")
        print("=" * 55)

        print(f"Total Revenue       : Rs. {results['total_revenue']:,.2f}")
        print(f"Total Orders        : {results['total_orders']}")
        print(f"Total Quantity Sold : {results['total_quantity']}")
        print(f"Average Order Value : Rs. {results['average_order_value']:,.2f}")
        print(f"Top Product         : {results['top_product']}")

        print("\nPRODUCT-WISE SALES")
        print("-" * 40)

        for product, sales in results["product_sales"].items():
            print(f"{product:<20} Rs. {sales:,.2f}")

        print("\nCATEGORY-WISE SALES")
        print("-" * 40)

        for category, sales in results["category_sales"].items():
            print(f"{category:<20} Rs. {sales:,.2f}")

        print("\nREGION-WISE SALES")
        print("-" * 40)

        for region, sales in results["region_sales"].items():
            print(f"{region:<20} Rs. {sales:,.2f}")

        print("\nMONTHLY SALES")
        print("-" * 40)

        for month, sales in results["monthly_sales"].items():
            print(f"{month:<20} Rs. {sales:,.2f}")

        print("=" * 55)

    def save_analysis(self, results):
        """Save analysis results as JSON."""

        output_file = (
            self.output_folder / "sales_analysis.json"
        )

        with open(output_file, "w", encoding="utf-8") as file:
            json.dump(results, file, indent=4)

        print(f"\nAnalysis saved to: {output_file}")

    def run(self):
        """Run complete sales analysis."""

        print("\nStarting Sales Analysis...")

        df = self.load_data()

        if df is None:
            return

        try:
            if df.empty:
                print("No sales records available!")
                return

            results = self.analyze_sales(df)

            self.display_results(results)

            self.save_analysis(results)

            print("\nSALES ANALYSIS COMPLETED SUCCESSFULLY!")

        except Exception as error:
            print(f"Analysis Error: {error}")


if __name__ == "__main__":

    analyzer = SalesAnalyzer(
        "output/cleaned_sales_data.xlsx"
    )

    analyzer.run()