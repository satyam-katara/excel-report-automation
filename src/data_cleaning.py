
import pandas as pd


def clean_data(df):
    """
    Validate and transform Superstore data
    for monthly MIS reporting.
    """

    # Work on a copy to preserve original data
    df = df.copy()

    print("Starting data validation...")

    # 1. Check missing values
    missing_values = df.isnull().sum().sum()

    # 2. Check duplicate records
    duplicate_rows = df.duplicated().sum()

    # 3. Check date validity
    invalid_order_dates = df["Order Date"].isna().sum()
    invalid_ship_dates = df["Ship Date"].isna().sum()

    # 4. Check business rules
    invalid_sales = (df["Sales"] <= 0).sum()
    invalid_quantity = (df["Quantity"] <= 0).sum()

    # Display validation results
    print("\nDATA QUALITY SUMMARY")
    print("=" * 35)
    print("Missing Values:", missing_values)
    print("Duplicate Rows:", duplicate_rows)
    print("Invalid Order Dates:", invalid_order_dates)
    print("Invalid Ship Dates:", invalid_ship_dates)
    print("Invalid Sales:", invalid_sales)
    print("Invalid Quantity:", invalid_quantity)

    # Stop processing if critical issues exist
    if missing_values > 0:
        raise ValueError("Dataset contains missing values.")

    if duplicate_rows > 0:
        raise ValueError("Dataset contains duplicate rows.")

    if invalid_order_dates > 0 or invalid_ship_dates > 0:
        raise ValueError("Dataset contains invalid dates.")

    if invalid_sales > 0 or invalid_quantity > 0:
        raise ValueError("Invalid Sales or Quantity values found.")

    # 5. Create date-based columns
    df["Year"] = df["Order Date"].dt.year
    df["Month"] = df["Order Date"].dt.month
    df["Month Name"] = df["Order Date"].dt.month_name()
    df["Year-Month"] = df["Order Date"].dt.to_period("M")

    print("\nData validation completed successfully!")
    print("Data transformation completed successfully!")
    print("Final Shape:", df.shape)

    return df
