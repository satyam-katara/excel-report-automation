
import argparse
from pathlib import Path

from data_loader import load_data
from data_cleaning import clean_data

from report_metrics import (
    calculate_kpis,
    category_analysis,
    subcategory_analysis,
    region_analysis,
    state_analysis,
    monthly_analysis,
    yearly_analysis
)

from excel_formatter import export_report


# Project paths
PROJECT_DIR = Path(__file__).resolve().parent.parent

DATA_PATH = (
    PROJECT_DIR / "Data" / "Sample - Superstore.csv"
)

REPORT_DIR = PROJECT_DIR / "reports"


def generate_mis_report(selected_month):

    print("\n" + "=" * 50)
    print("SUPERSTORE MIS AUTOMATION")
    print("=" * 50)

    # 1. Load data
    print("\nStep 1: Loading dataset...")
    df = load_data(DATA_PATH)

    # 2. Clean and transform
    print("\nStep 2: Validating and transforming data...")
    cleaned_df = clean_data(df)

    # 3. Calculate business metrics
    print("\nStep 3: Calculating business metrics...")

    kpis = calculate_kpis(cleaned_df)

    category_report = category_analysis(cleaned_df)
    subcategory_report = subcategory_analysis(cleaned_df)
    region_report = region_analysis(cleaned_df)
    state_report = state_analysis(cleaned_df)
    monthly_report = monthly_analysis(cleaned_df)
    yearly_report = yearly_analysis(cleaned_df)

    # 4. Generate Excel report
    print("\nStep 4: Generating Excel report...")

    output_path = (
        REPORT_DIR /
        f"Superstore_MIS_{selected_month}.xlsx"
    )

    export_report(
        df=cleaned_df,
        kpis=kpis,
        category_report=category_report,
        subcategory_report=subcategory_report,
        region_report=region_report,
        state_report=state_report,
        monthly_report=monthly_report,
        yearly_report=yearly_report,
        output_path=output_path,
        selected_month=selected_month
    )

    print("\n" + "=" * 50)
    print("AUTOMATION COMPLETED SUCCESSFULLY!")
    print("Report:", output_path)
    print("=" * 50)


if __name__ == "__main__":

    parser = argparse.ArgumentParser(
        description="Superstore Monthly MIS Automation"
    )

    parser.add_argument(
        "--month",
        required=True,
        help="Reporting month in YYYY-MM format"
    )

    args = parser.parse_args()

    generate_mis_report(args.month)
