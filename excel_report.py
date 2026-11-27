import pandas as pd
from pathlib import Path
from datetime import datetime

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo


class ExcelReportGenerator:

    def __init__(self):
        self.input_file = Path(
            "output/cleaned_sales_data.xlsx"
        )

        self.output_folder = Path("output")
        self.output_folder.mkdir(exist_ok=True)

        self.output_file = (
            self.output_folder /
            "Automated_Sales_Report.xlsx"
        )

    def load_data(self):
        """Load cleaned Excel dataset."""

        if not self.input_file.exists():
            raise FileNotFoundError(
                "Cleaned sales data not found!"
            )

        df = pd.read_excel(self.input_file)

        if df.empty:
            raise ValueError("No sales records found!")

        required = [
            "Order_ID", "Date", "Product",
            "Category", "Quantity", "Unit_Price",
            "Region", "Total_Sales"
        ]

        missing = [
            col for col in required if col not in df.columns
        ]

        if missing:
            raise ValueError(
                f"Missing columns: {missing}"
            )

        return df

    def generate_sheets(self, df):
        """Generate multiple Excel report sheets."""

        total_revenue = df["Total_Sales"].sum()
        total_orders = len(df)
        total_quantity = df["Quantity"].sum()

        average_order_value = (
            total_revenue / total_orders
        )

        top_product = (
            df.groupby("Product")["Total_Sales"]
            .sum()
            .idxmax()
        )

        summary = pd.DataFrame({
            "Metric": [
                "Total Revenue",
                "Total Orders",
                "Total Quantity Sold",
                "Average Order Value",
                "Top Performing Product"
            ],
            "Value": [
                total_revenue,
                total_orders,
                total_quantity,
                average_order_value,
                top_product
            ]
        })

        product_report = (
            df.groupby("Product")
            .agg(
                Total_Quantity=("Quantity", "sum"),
                Total_Revenue=("Total_Sales", "sum"),
                Number_of_Orders=("Order_ID", "count")
            )
            .reset_index()
            .sort_values(
                "Total_Revenue", ascending=False
            )
        )

        category_report = (
            df.groupby("Category")
            .agg(
                Total_Quantity=("Quantity", "sum"),
                Total_Revenue=("Total_Sales", "sum"),
                Number_of_Orders=("Order_ID", "count")
            )
            .reset_index()
            .sort_values(
                "Total_Revenue", ascending=False
            )
        )

        region_report = (
            df.groupby("Region")
            .agg(
                Total_Quantity=("Quantity", "sum"),
                Total_Revenue=("Total_Sales", "sum"),
                Number_of_Orders=("Order_ID", "count")
            )
            .reset_index()
            .sort_values(
                "Total_Revenue", ascending=False
            )
        )

        df["Date"] = pd.to_datetime(df["Date"])

        monthly_report = (
            df.groupby(df["Date"].dt.strftime("%Y-%m"))
            .agg(
                Total_Revenue=("Total_Sales", "sum"),
                Number_of_Orders=("Order_ID", "count")
            )
            .reset_index()
            .rename(columns={"Date": "Month"})
        )

        with pd.ExcelWriter(
            self.output_file,
            engine="openpyxl"
        ) as writer:

            summary.to_excel(
                writer,
                sheet_name="Executive Summary",
                index=False,
                startrow=3
            )

            product_report.to_excel(
                writer,
                sheet_name="Product Analysis",
                index=False
            )

            category_report.to_excel(
                writer,
                sheet_name="Category Analysis",
                index=False
            )

            region_report.to_excel(
                writer,
                sheet_name="Region Analysis",
                index=False
            )

            monthly_report.to_excel(
                writer,
                sheet_name="Monthly Analysis",
                index=False
            )

            df.to_excel(
                writer,
                sheet_name="Cleaned Data",
                index=False
            )

        print("Excel sheets generated successfully!")

    def format_workbook(self):
        """Apply professional Excel formatting."""

        workbook = load_workbook(self.output_file)

        header_fill = PatternFill(
            fill_type="solid",
            fgColor="17365D"
        )

        title_fill = PatternFill(
            fill_type="solid",
            fgColor="1F4E78"
        )

        header_font = Font(
            bold=True,
            color="FFFFFF",
            size=11
        )

        title_font = Font(
            bold=True,
            color="FFFFFF",
            size=18
        )

        thin_border = Border(
            bottom=Side(
                style="thin",
                color="D9E2F3"
            )
        )

        for ws in workbook.worksheets:

            ws.sheet_view.showGridLines = False

            if ws.title == "Executive Summary":

                ws.merge_cells("A1:B1")

                ws["A1"] = "AUTOMATED SALES REPORT"
                ws["A1"].fill = title_fill
                ws["A1"].font = title_font
                ws["A1"].alignment = Alignment(
                    horizontal="center",
                    vertical="center"
                )

                ws.row_dimensions[1].height = 35

                ws["A2"] = "Generated On"
                ws["B2"] = datetime.now().strftime(
                    "%Y-%m-%d %H:%M:%S"
                )

                header_row = 4

            else:
                header_row = 1

            # Format headers
            for cell in ws[header_row]:

                cell.fill = header_fill
                cell.font = header_font

                cell.alignment = Alignment(
                    horizontal="center",
                    vertical="center"
                )

            ws.row_dimensions[header_row].height = 25

            # Format data cells
            for row in ws.iter_rows(
                min_row=header_row + 1
            ):

                for cell in row:

                    cell.border = thin_border

                    cell.alignment = Alignment(
                        vertical="center"
                    )

                    column_name = ws.cell(
                        header_row,
                        cell.column
                    ).value

                    if column_name in [
                        "Total_Sales",
                        "Total_Revenue",
                        "Average Order Value"
                    ]:
                        cell.number_format = (
                            '"Rs. "#,##0.00'
                        )

            # Automatically adjust column widths
            for column in ws.columns:

                max_length = 0
                column_letter = get_column_letter(
                    column[0].column
                )

                for cell in column:

                    if cell.value is not None:
                        max_length = max(
                            max_length,
                            len(str(cell.value))
                        )

                ws.column_dimensions[
                    column_letter
                ].width = min(max_length + 4, 35)

            # Freeze headers and enable filters
            ws.freeze_panes = (
                f"A{header_row + 1}"
            )

            if ws.title != "Executive Summary":

                last_column = get_column_letter(
                    ws.max_column
                )

                table_range = (
                    f"A1:{last_column}{ws.max_row}"
                )

                table_name = (
                    ws.title.replace(" ", "")
                    .replace("-", "")
                )

                table = Table(
                    displayName=table_name,
                    ref=table_range
                )

                style = TableStyleInfo(
                    name="TableStyleMedium2",
                    showFirstColumn=False,
                    showLastColumn=False,
                    showRowStripes=True,
                    showColumnStripes=False
                )

                table.tableStyleInfo = style

                ws.add_table(table)

        workbook.save(self.output_file)

        print("Professional formatting applied!")

    def run(self):
        """Execute report generation workflow."""

        print("\n" + "=" * 55)
        print("     AUTOMATED EXCEL REPORT GENERATOR")
        print("=" * 55)

        try:
            df = self.load_data()

            print(f"Records Loaded: {len(df)}")

            self.generate_sheets(df)

            self.format_workbook()

            print("\nREPORT GENERATION SUCCESSFUL!")
            print("-" * 55)
            print(f"Output File: {self.output_file}")
            print("Sheets Created: 6")
            print("Status: Successfully Generated!")
            print("=" * 55)

        except Exception as error:
            print(f"\nReport Generation Error: {error}")


if __name__ == "__main__":

    generator = ExcelReportGenerator()
    generator.run()