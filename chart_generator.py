import pandas as pd
import plotly.express as px
from pathlib import Path


class ChartGenerator:

    def __init__(self):
        self.input_file = Path(
            "output/cleaned_sales_data.xlsx"
        )

        self.chart_folder = Path("charts")
        self.chart_folder.mkdir(exist_ok=True)

    def load_data(self):
        """Load cleaned sales data."""

        if not self.input_file.exists():
            raise FileNotFoundError(
                "Cleaned sales Excel file not found!"
            )

        df = pd.read_excel(self.input_file)

        if df.empty:
            raise ValueError("No sales data available!")

        return df

    def generate_product_chart(self, df):
        """Generate product-wise sales bar chart."""

        data = (
            df.groupby("Product")["Total_Sales"]
            .sum()
            .reset_index()
            .sort_values("Total_Sales", ascending=False)
        )

        fig = px.bar(
            data,
            x="Product",
            y="Total_Sales",
            title="Product-wise Sales Analysis",
            text_auto=".2s",
            labels={
                "Product": "Product",
                "Total_Sales": "Revenue (Rs.)"
            },
            template="plotly_white"
        )

        fig.update_layout(
            title_x=0.5,
            xaxis_title="Products",
            yaxis_title="Total Revenue (Rs.)"
        )

        fig.write_html(
            self.chart_folder / "product_sales.html"
        )

    def generate_category_chart(self, df):
        """Generate category-wise pie chart."""

        data = (
            df.groupby("Category")["Total_Sales"]
            .sum()
            .reset_index()
        )

        fig = px.pie(
            data,
            names="Category",
            values="Total_Sales",
            title="Category-wise Sales Distribution",
            hole=0.4,
            template="plotly_white"
        )

        fig.update_layout(title_x=0.5)

        fig.write_html(
            self.chart_folder / "category_sales.html"
        )

    def generate_region_chart(self, df):
        """Generate region-wise sales chart."""

        data = (
            df.groupby("Region")["Total_Sales"]
            .sum()
            .reset_index()
            .sort_values("Total_Sales", ascending=False)
        )

        fig = px.bar(
            data,
            x="Region",
            y="Total_Sales",
            title="Region-wise Sales Performance",
            text_auto=".2s",
            template="plotly_white"
        )

        fig.update_layout(
            title_x=0.5,
            xaxis_title="Region",
            yaxis_title="Revenue (Rs.)"
        )

        fig.write_html(
            self.chart_folder / "region_sales.html"
        )

    def generate_monthly_chart(self, df):
        """Generate monthly revenue trend chart."""

        df["Date"] = pd.to_datetime(df["Date"])

        data = (
            df.groupby(df["Date"].dt.to_period("M"))
            ["Total_Sales"]
            .sum()
            .reset_index()
        )

        data["Month"] = (
            data["Date"].astype(str)
        )

        fig = px.line(
            data,
            x="Month",
            y="Total_Sales",
            markers=True,
            title="Monthly Revenue Trend",
            template="plotly_white"
        )

        fig.update_layout(
            title_x=0.5,
            xaxis_title="Month",
            yaxis_title="Revenue (Rs.)"
        )

        fig.write_html(
            self.chart_folder / "monthly_sales.html"
        )

    def run(self):
        """Execute all chart generation functions."""

        print("\n" + "=" * 50)
        print("    AUTOMATED CHART GENERATOR")
        print("=" * 50)

        try:
            df = self.load_data()

            print("Sales data loaded successfully!")

            self.generate_product_chart(df)
            print("1. Product Sales Chart Generated")

            self.generate_category_chart(df)
            print("2. Category Sales Chart Generated")

            self.generate_region_chart(df)
            print("3. Region Sales Chart Generated")

            self.generate_monthly_chart(df)
            print("4. Monthly Sales Chart Generated")

            print("\nAll charts generated successfully!")
            print(f"Chart Location: {self.chart_folder}")
            print("=" * 50)

        except Exception as error:
            print(f"\nChart Generation Error: {error}")


if __name__ == "__main__":

    generator = ChartGenerator()
    generator.run()