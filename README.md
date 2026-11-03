<div align="center">

# 📊 Excel Report Automation
### Superstore MIS Dashboard

**One Python command turns raw sales data into a formatted 8-sheet Excel MIS workbook with an executive dashboard.**

![Python](https://img.shields.io/badge/Python-3.x-3776AB?logo=python&logoColor=white)
![Pandas](https://img.shields.io/badge/Pandas-data%20processing-150458?logo=pandas&logoColor=white)
![OpenPyXL](https://img.shields.io/badge/OpenPyXL-Excel%20output-217346?logo=microsoftexcel&logoColor=white)
![VBA](https://img.shields.io/badge/VBA-control%20panel-5C2D91)
![Pytest](https://img.shields.io/badge/Pytest-6%20tests-0A9EDC?logo=pytest&logoColor=white)
![Platform](https://img.shields.io/badge/Platform-Windows-0078D6?logo=windows&logoColor=white)

[Overview](#-overview) · [Impact](#-project-impact) · [How it works](#-how-it-works) · [Workbook](#-workbook-contents) · [Quickstart](#-quickstart) · [Scheduling](#-scheduling) · [VBA](#-vba-integration) · [Tests](#-tests)

</div>

---

## 📌 Overview

Monthly MIS reporting is repetitive. Someone loads the data, cleans it, builds summaries, makes charts, formats the file, and then does it all again next month. Manual work like this also leads to copy-paste errors, broken formulas, and reports that look different each cycle.

This project automates the whole process. The estimated manual effort is **3.5 analyst-hours per report cycle**.

![Executive Dashboard](screenshots/executive_dashboard.png)

### What it does

- Loads and validates **9,994 Superstore sales records** (2014–2017)
- Cleans the data: missing values, duplicates, invalid dates
- Calculates KPIs: sales, profit, profit margin, order volume, average order value (AOV)
- Builds category, subcategory, region, state, monthly, and yearly analysis
- Creates an Excel workbook with:
  - **8 worksheets**
  - **6 KPI cards**
  - **4 charts**
  - conditional formatting, AutoFilters, and print-ready layouts
- Detects the latest reporting month in the input data
- Includes a **VBA control panel** to pick a month and start the report
- Runs on a schedule with **Windows Task Scheduler**

---

## ⚡ Project Impact

| Metric | Manual process | Automated pipeline |
|---|---:|---:|
| Time per report | ~3.5 hours *(estimated)* | **2.71 seconds** *(measured)* |
| Workbook structure | Built by hand each time | 8 consistent worksheets |
| Repeated calculations | Manual / spreadsheet-based | Automated in Python |

**Estimated time saved:** about **210 minutes per run**, which is an estimated **99%+ reduction** in report-generation time.

<details>
<summary><b>📏 Measurement notes (click to expand)</b></summary>

<br>

- The **3.5 hours (210 minutes)** manual baseline is an **estimate**. It was not timed with a stopwatch on a real company process.
- The **2.71 seconds** runtime was **measured** on the development machine during a successful run of `python run_monthly_report.py`. Runtime will change with hardware, data size, and workload.
- This is a project-level comparison, not a promise of business-wide savings. Review, interpretation, and sending the report still take human time.

To measure the runtime on your own machine, run this in PowerShell from the project folder:

```powershell
Measure-Command { python run_monthly_report.py }
```

</details>

---

## 🔄 How It Works

```text
Superstore CSV
     │
     ▼
Data Loading ........... Pandas: schema checks and date parsing
     │
     ▼
Cleaning & Validation .. missing values, duplicates, invalid dates
     │
     ▼
Transformation ......... Year, Month, and Year-Month fields
     │
     ▼
Business Metrics ....... KPIs and category / region / state / time breakdowns
     │
     ▼
Excel Generation ....... OpenPyXL: 8 worksheets, KPI cards, charts, formatting
     │
     ▼
reports/Superstore_MIS_YYYY-MM.xlsx
```

### Three ways to run it

| # | Option | How it works |
|:-:|---|---|
| 1 | **Python entry point** | Run `run_monthly_report.py`. It finds the latest month in the dataset and builds the report. |
| 2 | **VBA control panel** | Choose a month in `MIS Control Panel!B8` inside the macro-enabled workbook. The macro starts the Python process. |
| 3 | **Windows Task Scheduler** | Runs the Python entry point automatically each month. |

> [!IMPORTANT]
> The scheduler only runs the report script. It **does not fetch new Superstore data**. You must update the input CSV yourself. If the CSV does not change, the script will keep detecting the same latest month.

---

## 📑 Workbook Contents

| # | Worksheet | What it shows |
|:-:|---|---|
| 1 | **Executive Dashboard** | 6 KPI cards, monthly sales and profit trends, category and regional charts |
| 2 | **Category Analysis** | Furniture, Technology, and Office Supplies performance |
| 3 | **Subcategory Analysis** | Sales and profit by subcategory |
| 4 | **Regional Analysis** | West, East, Central, and South compared |
| 5 | **State Analysis** | Sales, profit, and profitability by state |
| 6 | **Monthly Summary** | Month-by-month sales and profit trends |
| 7 | **Yearly Summary** | Annual performance, 2014–2017 |
| 8 | **Monthly Data** | Transaction-level rows for the selected reporting month |

> [!NOTE]
> **Reporting scope:** The dashboard KPIs and the category, subcategory, regional, state, monthly, and yearly sheets cover the **full dataset (2014–2017)**. Only the **Monthly Data** sheet is limited to the selected reporting month.

---

## 🚀 Quickstart

**1. Clone the repository**

```bash
git clone https://github.com/satyam-katara/excel-report-automation.git
cd excel-report-automation
```

**2. Create and activate a virtual environment (Windows PowerShell)**

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
```

<details>
<summary>PowerShell blocks the activation script?</summary>

<br>

Use Command Prompt instead:

```cmd
.venv\Scripts\activate.bat
```

</details>

**3. Install dependencies**

```bash
pip install -r requirements.txt
```

**4. Generate the report**

```bash
python run_monthly_report.py
```

The workbook is saved in the `reports/` folder as `Superstore_MIS_YYYY-MM.xlsx`.

---

## ⏰ Scheduling

The project is set up with these Windows Task Scheduler settings:

| Setting | Value |
|---|---|
| Task name | Superstore Monthly MIS Automation |
| Trigger | Monthly, on the 1st at 8:00 AM |
| Script | `run_monthly_report.py` |
| Reporting month | Detected automatically from the latest month in the input data |

**Before you rely on it:**

- The computer must be on, or set to wake for the task
- The task must be enabled
- The task must use the project's virtual environment and the correct working directory
- The input CSV must be refreshed separately when new data arrives

---

## 🧩 VBA Integration

The `vba/` folder holds the macro source and the control panel workbook:

- `GenerateReport.bas` — VBA macro source
- `Superstore_MIS_Automation.xlsm` — macro-enabled control panel workbook

The control panel reads the reporting month from `MIS Control Panel!B8` and starts the Python report process. Python and the project dependencies must be installed on the machine that runs the macro.

---

## ✅ Tests

Run the tests from the project root:

```bash
pytest -v
```

There are **6 tests** covering the data loader, the cleaning logic, and the report metrics. Run them locally to confirm they pass before you publish changes.

---

## 🗂️ Project Structure

```text
excel-report-automation/
├── Data/
│   └── Sample - Superstore.csv
├── logs/
├── reports/                      # Generated workbooks (git-ignored)
├── screenshots/
│   └── executive_dashboard.png
├── src/
│   ├── main.py                   # Pipeline orchestration
│   ├── data_loader.py            # CSV loading, schema and date checks
│   ├── data_cleaning.py          # Cleaning and validation
│   ├── report_metrics.py         # KPI and breakdown calculations
│   ├── excel_formatter.py        # Workbook, dashboard, charts, formatting
│   └── inspect_data.ipynb        # Data exploration notebook
├── tests/                        # Pytest suite
├── vba/
│   ├── GenerateReport.bas        # Control panel macro source
│   └── Superstore_MIS_Automation.xlsm
├── run_monthly_report.py         # Month-aware entry point
├── requirements.txt
├── .gitignore
└── README.md
```

---

## 🛠️ Technology Stack

| Area | Tools |
|---|---|
| Language | Python |
| Data processing | Pandas, NumPy |
| Excel generation | OpenPyXL |
| Visualization | Matplotlib, Seaborn |
| Excel automation | VBA |
| Scheduling | Windows Task Scheduler |
| Testing | Pytest |

---

## 🔮 Possible Future Enhancements

- [ ] Send generated reports by email automatically
- [ ] Budget-versus-actual analysis when budget data is available
- [ ] Native Excel slicers for interactive dashboard filtering
- [ ] Automatic ingestion of refreshed source data

---

## 👤 Author

**Satyam Katara**
B.Tech, Computer Science and Engineering (Data Science), 2027

[![GitHub](https://img.shields.io/badge/GitHub-satyam--katara-181717?logo=github&logoColor=white)](https://github.com/satyam-katara)
