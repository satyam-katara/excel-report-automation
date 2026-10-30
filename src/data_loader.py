
from pathlib import Path
import pandas as pd


# Required columns for our MIS report
REQUIRED_COLUMNS = [
    "Row ID",
    "Order ID",
    "Order Date",
    "Ship Date",
    "Ship Mode",
    "Customer ID",
    "Customer Name",
    "Segment",
    "Country",
    "City",
    "State",
    "Postal Code",
    "Region",
    "Product ID",
    "Category",
    "Sub-Category",
    "Product Name",
    "Sales",
    "Quantity",
    "Discount",
    "Profit"
]


def load_data(file_path):
    """
    Load the Superstore CSV and validate its structure.
    """

    file_path = Path(file_path)

    # Check file existence
    if not file_path.exists():
        raise FileNotFoundError(
            f"Dataset not found: {file_path}"
        )

    # Load CSV
    df = pd.read_csv(
        file_path,
        encoding="cp1252"
    )

    # Validate required columns
    missing_columns = [
        col for col in REQUIRED_COLUMNS
        if col not in df.columns
    ]

    if missing_columns:
        raise ValueError(
            f"Missing required columns: {missing_columns}"
        )

    # Convert dates
    for col in ["Order Date", "Ship Date"]:
        df[col] = pd.to_datetime(
            df[col],
            format="mixed",
            errors="coerce"
        )

    print("Dataset loaded successfully!")
    print(f"Records: {len(df)}")
    print(f"Columns: {len(df.columns)}")

    return df
