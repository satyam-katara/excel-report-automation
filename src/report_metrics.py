
import pandas as pd


def calculate_kpis(df):
    """Calculate overall business KPIs."""

    total_sales = df["Sales"].sum()
    total_profit = df["Profit"].sum()
    total_orders = df["Order ID"].nunique()
    total_customers = df["Customer ID"].nunique()

    return {
        "Total Sales": total_sales,
        "Total Profit": total_profit,
        "Total Orders": total_orders,
        "Total Customers": total_customers,
        "Profit Margin (%)": (
            total_profit / total_sales * 100
        ),
        "Average Order Value": (
            total_sales / total_orders
        )
    }


def category_analysis(df):
    """Calculate category-wise performance."""

    result = (
        df.groupby("Category")
        .agg(
            Total_Sales=("Sales", "sum"),
            Total_Profit=("Profit", "sum"),
            Total_Quantity=("Quantity", "sum"),
            Total_Orders=("Order ID", "nunique")
        )
        .reset_index()
    )

    result["Profit_Margin"] = (
        result["Total_Profit"] /
        result["Total_Sales"] * 100
    )

    return result.sort_values(
        "Total_Sales", ascending=False
    )


def subcategory_analysis(df):
    """Calculate sub-category-wise performance."""

    result = (
        df.groupby(["Category", "Sub-Category"])
        .agg(
            Total_Sales=("Sales", "sum"),
            Total_Profit=("Profit", "sum"),
            Total_Quantity=("Quantity", "sum"),
            Total_Orders=("Order ID", "nunique")
        )
        .reset_index()
    )

    result["Profit_Margin"] = (
        result["Total_Profit"] /
        result["Total_Sales"] * 100
    )

    return result.sort_values(
        "Total_Sales", ascending=False
    )


def region_analysis(df):
    """Calculate region-wise performance."""

    result = (
        df.groupby("Region")
        .agg(
            Total_Sales=("Sales", "sum"),
            Total_Profit=("Profit", "sum"),
            Total_Orders=("Order ID", "nunique"),
            Total_Customers=("Customer ID", "nunique")
        )
        .reset_index()
    )

    result["Profit_Margin"] = (
        result["Total_Profit"] /
        result["Total_Sales"] * 100
    )

    return result.sort_values(
        "Total_Sales", ascending=False
    )


def state_analysis(df):
    """Calculate state-wise performance."""

    result = (
        df.groupby("State")
        .agg(
            Total_Sales=("Sales", "sum"),
            Total_Profit=("Profit", "sum"),
            Total_Orders=("Order ID", "nunique")
        )
        .reset_index()
    )

    result["Profit_Margin"] = (
        result["Total_Profit"] /
        result["Total_Sales"] * 100
    )

    return result.sort_values(
        "Total_Sales", ascending=False
    )


def monthly_analysis(df):
    """Calculate monthly business performance."""

    result = (
        df.groupby("Year-Month")
        .agg(
            Total_Sales=("Sales", "sum"),
            Total_Profit=("Profit", "sum"),
            Total_Orders=("Order ID", "nunique"),
            Total_Quantity=("Quantity", "sum")
        )
        .reset_index()
    )

    return result.sort_values("Year-Month")


def yearly_analysis(df):
    """Calculate annual business performance."""

    result = (
        df.groupby("Year")
        .agg(
            Total_Sales=("Sales", "sum"),
            Total_Profit=("Profit", "sum"),
            Total_Orders=("Order ID", "nunique"),
            Total_Quantity=("Quantity", "sum")
        )
        .reset_index()
    )

    result["Profit_Margin"] = (
        result["Total_Profit"] /
        result["Total_Sales"] * 100
    )

    result["Sales_Growth_%"] = (
        result["Total_Sales"].pct_change() * 100
    )

    return result
