
from pathlib import Path
import subprocess
import sys

import pandas as pd


PROJECT_DIR = Path(__file__).resolve().parent

DATA_PATH = PROJECT_DIR / "Data" / "Sample - Superstore.csv"
PYTHON_PATH = PROJECT_DIR / ".venv" / "Scripts" / "python.exe"
MAIN_SCRIPT = PROJECT_DIR / "src" / "main.py"


def generate_monthly_report():

    print("=" * 50)
    print("MONTHLY MIS AUTOMATION")
    print("=" * 50)

    if not DATA_PATH.exists():
        raise FileNotFoundError(f"Dataset not found: {DATA_PATH}")

    # Find the latest month available in the dataset
    df = pd.read_csv(
        DATA_PATH,
        encoding="cp1252",
        usecols=["Order Date"]
    )

    order_dates = pd.to_datetime(
        df["Order Date"],
        format="mixed",
        errors="coerce"
    )

    if order_dates.isna().all():
        raise ValueError("No valid Order Dates found.")

    selected_month = order_dates.max().strftime("%Y-%m")

    print(f"Latest available reporting month: {selected_month}")

    # Run the existing MIS automation
    result = subprocess.run(
        [
            str(PYTHON_PATH),
            str(MAIN_SCRIPT),
            "--month",
            selected_month
        ],
        cwd=PROJECT_DIR
    )

    if result.returncode != 0:
        raise RuntimeError("Monthly MIS report generation failed.")

    print("\nMonthly automation completed successfully!")


if __name__ == "__main__":
    try:
        generate_monthly_report()
    except Exception as error:
        print(f"\nERROR: {error}")
        sys.exit(1)
