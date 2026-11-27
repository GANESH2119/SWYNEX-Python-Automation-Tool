import subprocess
import sys
from pathlib import Path


def run_module(script_name):

    print("\n" + "=" * 60)
    print(f"RUNNING: {script_name}")
    print("=" * 60)

    try:
        subprocess.run(
            [sys.executable, script_name],
            check=True
        )

        print(f"\n{script_name} completed successfully!")
        return True

    except subprocess.CalledProcessError:
        print(f"\nERROR: {script_name} failed!")
        return False

    except Exception as error:
        print(f"\nUnexpected Error: {error}")
        return False


def main():

    print("=" * 60)
    print(" AUTOMATED EXCEL REPORT GENERATOR")
    print(" Complete Business Reporting Automation System")
    print("=" * 60)

    modules = [
        "generate_data.py",
        "data_processor.py",
        "sales_analyzer.py",
        "excel_report.py",
        "chart_generator.py",
        "pdf_report.py"
    ]

    for module in modules:

        if not Path(module).exists():
            print(f"\nRequired file missing: {module}")
            return

        success = run_module(module)

        if not success:
            print("\nAPPLICATION STOPPED DUE TO ERROR.")
            return

    print("\n" + "=" * 60)
    print(" ALL OPERATIONS COMPLETED SUCCESSFULLY!")
    print("=" * 60)

    print("\nGENERATED PROJECT OUTPUTS:")

    print("1. Cleaned Excel Dataset")
    print("2. Sales Analysis JSON")
    print("3. Professional Excel Report")
    print("4. Product Sales Chart")
    print("5. Category Sales Chart")
    print("6. Region Sales Chart")
    print("7. Monthly Sales Chart")
    print("8. Professional PDF Report")

    print("\nThank you for using the application!")


if __name__ == "__main__":
    main()