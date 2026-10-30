
from pathlib import Path
import sys

import pandas as pd
from openpyxl import load_workbook

# Add src folder to Python path
PROJECT_DIR = Path(__file__).resolve().parent.parent
SRC_DIR = PROJECT_DIR / "src"
sys.path.insert(0, str(SRC_DIR))

from data_loader import load_data
from data_cleaning import clean_data
from report_metrics import (
    calculate_kpis,
    category_analysis,
    subcategory_analysis,
    region_analysis,
    state_analysis,
    monthly_analysis,
    yearly_analysis,
)

DATA_PATH = PROJECT_DIR / "Data" / "Sample - Superstore.csv"
REPORT_PATH = PROJECT_DIR / "reports" / "Superstore_MIS_2017-12.xlsx"


def get_cleaned_data():
    df = load_data(DATA_PATH)
    return clean_data(df)


def test_dataset_loading():
    df = load_data(DATA_PATH)
    assert len(df) == 9994
    assert len(df.columns) == 21


def test_data_cleaning():
    df = get_cleaned_data()
    assert df.isnull().sum().sum() == 0
    assert df.duplicated().sum() == 0
    assert df.shape == (9994, 25)


def test_kpi_calculations():
    df = get_cleaned_data()
    kpis = calculate_kpis(df)

    assert round(kpis["Total Sales"], 2) == 2297200.86
    assert round(kpis["Total Profit"], 2) == 286397.02
    assert kpis["Total Orders"] == 5009
    assert kpis["Total Customers"] == 793


def test_category_analysis():
    df = get_cleaned_data()
    result = category_analysis(df)

    assert round(result["Total_Sales"].sum(), 2) == 2297200.86
    assert round(result["Total_Profit"].sum(), 2) == 286397.02


def test_other_analysis_totals():
    df = get_cleaned_data()

    subcategory = subcategory_analysis(df)
    region = region_analysis(df)
    state = state_analysis(df)
    monthly = monthly_analysis(df)
    yearly = yearly_analysis(df)

    assert round(subcategory["Total_Sales"].sum(), 2) == 2297200.86
    assert round(region["Total_Sales"].sum(), 2) == 2297200.86
    assert round(state["Total_Sales"].sum(), 2) == 2297200.86
    assert round(monthly["Total_Sales"].sum(), 2) == 2297200.86
    assert round(yearly["Total_Sales"].sum(), 2) == 2297200.86


def test_excel_report():
    assert REPORT_PATH.exists(), "Excel report does not exist!"

    workbook = load_workbook(REPORT_PATH, data_only=False)

    expected_sheets = [
        "Executive Dashboard",
        "Category Analysis",
        "Subcategory Analysis",
        "Regional Analysis",
        "State Analysis",
        "Monthly Summary",
        "Yearly Summary",
        "Monthly Data",
    ]

    assert workbook.sheetnames == expected_sheets

    dashboard = workbook["Executive Dashboard"]
    assert len(dashboard._charts) == 4

    monthly_data = workbook["Monthly Data"]
    assert monthly_data.max_row > 1

    workbook.close()
