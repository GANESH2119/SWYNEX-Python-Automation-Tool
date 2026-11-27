import pandas as pd
from pathlib import Path


class DataProcessor:

    def __init__(self, input_file):
        self.input_file = Path(input_file)
        self.output_folder = Path("output")
        self.output_folder.mkdir(exist_ok=True)

        self.required_columns = [
            "Order_ID",
            "Date",
            "Product",
            "Category",
            "Quantity",
            "Unit_Price",
            "Region"
        ]

    def load_data(self):
        """Load Excel data safely."""
        try:
            if not self.input_file.exists():
                raise FileNotFoundError(
                    f"Excel file not found: {self.input_file}"
                )

            df = pd.read_excel(self.input_file)

            print("\nExcel file loaded successfully!")
            print(f"Original Records: {len(df)}")

            return df

        except Exception as error:
            print(f"Error loading Excel file: {error}")
            return None

    def validate_columns(self, df):
        """Check required columns."""
        missing = [
            col for col in self.required_columns
            if col not in df.columns
        ]

        if missing:
            raise ValueError(
                f"Missing required columns: {missing}"
            )

        print("Column validation successful!")

    def clean_data(self, df):
        """Clean and validate dataset."""

        original_count = len(df)

        # Remove completely empty rows
        df = df.dropna(how="all").copy()

        # Remove duplicate records
        df = df.drop_duplicates(
            subset=["Order_ID"]
        ).copy()

        # Clean text columns
        for col in ["Product", "Category", "Region"]:
            df[col] = df[col].astype("string").str.strip()

        # Convert numeric columns
        df["Quantity"] = pd.to_numeric(
            df["Quantity"], errors="coerce"
        )

        df["Unit_Price"] = pd.to_numeric(
            df["Unit_Price"], errors="coerce"
        )

        # Convert date column
        df["Date"] = pd.to_datetime(
            df["Date"], errors="coerce"
        )

        # Remove records with invalid required values
        df = df.dropna(
            subset=self.required_columns
        ).copy()

        # Remove invalid quantities and prices
        df = df[
            (df["Quantity"] > 0) &
            (df["Unit_Price"] >= 0)
        ].copy()

        # Calculate total sales
        df["Total_Sales"] = (
            df["Quantity"] * df["Unit_Price"]
        )

        removed = original_count - len(df)

        print("\nData Cleaning Completed!")
        print(f"Original Records : {original_count}")
        print(f"Valid Records    : {len(df)}")
        print(f"Removed Records  : {removed}")

        return df

    def save_data(self, df):
        """Save cleaned data to Excel."""

        output_file = (
            self.output_folder /
            "cleaned_sales_data.xlsx"
        )

        df.to_excel(output_file, index=False)

        print("\nCleaned Excel file saved!")
        print(f"Location: {output_file}")

        return output_file

    def process(self):
        """Execute complete processing workflow."""

        print("=" * 50)
        print("       AUTOMATED DATA PROCESSOR")
        print("=" * 50)

        df = self.load_data()

        if df is None:
            return None

        try:
            self.validate_columns(df)

            cleaned_df = self.clean_data(df)

            if cleaned_df.empty:
                print("No valid records available.")
                return None

            self.save_data(cleaned_df)

            print("\nPROCESSING SUMMARY")
            print("-" * 50)
            print(f"Total Orders : {len(cleaned_df)}")
            print(
                f"Total Sales  : Rs. "
                f"{cleaned_df['Total_Sales'].sum():,.2f}"
            )
            print(
                f"Average Sale : Rs. "
                f"{cleaned_df['Total_Sales'].mean():,.2f}"
            )

            print("=" * 50)
            print("DATA PROCESSING SUCCESSFUL!")
            print("=" * 50)

            return cleaned_df

        except Exception as error:
            print(f"\nProcessing Error: {error}")
            return None


if __name__ == "__main__":

    processor = DataProcessor(
        "data/sales_data.xlsx"
    )

    processor.process()