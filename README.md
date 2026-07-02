# PlateMaster

**PlateMaster** is a Streamlit dashboard for analyzing multi-well plate and plate-reader data stored in Excel files. It helps you load one file, many files, or a ZIP archive, choose plate rows and columns, aggregate measurements, remove outliers, compare groups statistically, visualize time-course results, and export a complete Excel report.

The project is designed for researchers who need a fast, visual, and repeatable workflow for processing plate experiments without manually copying values between spreadsheets.

---

## Table of contents

- [Key features](#key-features)
- [Typical workflow](#typical-workflow)
- [Input data format](#input-data-format)
- [Installation](#installation)
- [Running the app](#running-the-app)
- [How to use PlateMaster](#how-to-use-platemaster)
- [Analysis modes](#analysis-modes)
- [Statistics and outlier handling](#statistics-and-outlier-handling)
- [Visualization options](#visualization-options)
- [Exported report](#exported-report)
- [Project structure](#project-structure)
- [Troubleshooting](#troubleshooting)
- [Notes and limitations](#notes-and-limitations)

---

## Key features

### Flexible data loading

- Load Excel files from a server/local folder path.
- Upload one or many Excel files directly from your computer.
- Upload a ZIP archive containing multiple Excel workbooks.
- Supports common Excel extensions: `.xlsx`, `.xls`, and `.xlsm`.
- Tries multiple Excel engines and includes fallbacks for difficult files, including delimited text exports and optional LibreOffice conversion.

### Plate-focused selection and aggregation

- Select any subset of rows and columns from the loaded tables.
- Aggregate values by original columns, by original rows, or by custom plate-template groups.
- Available aggregation methods:
  - None
  - Mean
  - Median
  - Min
  - Max
  - Standard deviation
  - Sum

### Outlier filtering

PlateMaster can remove outliers before aggregation using:

- IQR filtering
- Z-score filtering
- MAD filtering
- No filtering

The app also keeps a detailed log of removed values so results remain transparent and auditable.

### Statistical analysis

- Basic descriptive statistics for selected values:
  - n
  - mean
  - median
  - mode
  - standard deviation
  - SEM
  - MAD
  - min/max
  - 95% confidence interval approximation
  - IQR/2
  - approximate p-value for mean different from zero
- Welch two-sample t-test between selected groups.
- Pairwise p-value matrix across time/metric combinations.

### Interactive charts

PlateMaster includes several Altair-based chart styles:

- Line + points
- Line + points by columns
- Scatter plot
- Bar plot by time
- Bar plot by columns
- Box plot by time
- Box plot by columns
- Violin plot by time
- Violin plot by columns

Optional error bars can display standard deviation, SEM, 95% CI, IQR/2, or MAD.

### Export

Generate a downloadable `plate_master_report.xlsx` workbook with:

- Aggregated table
- Statistical analysis
- T-test analysis
- Pairwise t-test p-values
- Outlier log
- Outlier statistics
- Plot data

---

## Typical workflow

1. Start the Streamlit app.
2. Upload Excel files or load a folder containing Excel files.
3. Review the representative input table.
4. Select rows and columns to analyze.
5. Choose an aggregation layout and aggregation function.
6. Optionally enable outlier filtering.
7. Press **Run**.
8. Inspect aggregated tables, statistics, t-tests, outlier logs, and charts.
9. Export the final Excel report.

---

## Input data format

PlateMaster expects each spreadsheet to contain a table where:

- The **first column** contains row identifiers, such as plate row names, well labels, sample names, or group labels.
- The remaining columns contain numeric measurements.
- Column names may represent time points, wavelengths, treatments, concentrations, or any other measurement dimension.

Example:

| Well | 0 | 10 | 20 | 30 |
| --- | ---: | ---: | ---: | ---: |
| A1 | 0.10 | 0.15 | 0.22 | 0.31 |
| A2 | 0.11 | 0.16 | 0.24 | 0.33 |
| B1 | 0.08 | 0.12 | 0.19 | 0.27 |

During loading, the app renames the first column internally to `row_id` and converts numeric measurement columns automatically.

### Time extraction from filenames

For cross-file time-course analysis, PlateMaster extracts the last number from each filename and uses it as `Time`.

Examples:

| Filename | Extracted time |
| --- | ---: |
| `plate_0.xlsx` | `0` |
| `experiment_10min.xlsx` | `10` |
| `sample_time_2.5.xls` | `2.5` |

If a filename does not contain a number, the file can still be loaded, but time-based statistics and plotting may be limited.

---

## Installation

### 1. Clone the repository

```bash
git clone https://github.com/narek-abelyan/PlateMaster.git
cd PlateMaster
```

### 2. Create and activate a virtual environment

```bash
python -m venv .venv
source .venv/bin/activate
```

On Windows PowerShell:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

### 3. Install dependencies

If the repository contains a `requirements.txt`, install it with:

```bash
pip install -r requirements.txt
```

If there is no requirements file, install the main dependencies manually:

```bash
pip install streamlit pandas numpy altair matplotlib openpyxl xlrd python-calamine
```

Optional, but useful for old or problematic `.xls` files:

- Install **LibreOffice** so PlateMaster can attempt automatic conversion to `.xlsx`.

---

## Running the app

From the repository directory, run:

```bash
streamlit run PlateMaster.py
```

You can also run the script directly:

```bash
python PlateMaster.py
```

When executed directly outside Streamlit, the script tries to restart itself through `streamlit run`.

---

## How to use PlateMaster

### 1. Choose a data source

In the sidebar, select one of the available source modes:

#### Server folder path

Use this mode when the Excel files are already available on the same machine or server where Streamlit is running.

1. Enter the folder path.
2. Click **Load folder**.
3. PlateMaster scans the folder for `.xlsx`, `.xls`, and `.xlsm` files.

#### Local PC upload

Use this mode when you want to upload files through the browser.

1. Upload one or many Excel files.
2. Optionally upload a ZIP archive containing Excel files.
3. Click **Load uploaded files**.

### 2. Review the representative input table

After loading, PlateMaster displays one representative table. Use it to confirm that rows, columns, and numeric values were recognized correctly.

### 3. Select rows and columns

Use the sidebar multiselect controls to choose the rows and columns included in the analysis.

### 4. Choose aggregation settings

Select:

- Aggregation layout
- Aggregation function
- Outlier filtering method
- Color scale mode
- Color palette

### 5. Save or load presets

Presets store frequently used analysis settings, including selected rows, columns, aggregation mode, filtering settings, chart options, and color options.

### 6. Run analysis

Click **Run** to calculate results for the current settings.

---

## Analysis modes

### By original columns

This mode aggregates selected rows within each selected column. It is useful when columns represent time points, wavelengths, or measurement channels.

Example result columns:

```text
0, 10, 20, 30
```

### By original rows

This mode aggregates selected columns within each selected row. It is useful when each row represents a sample, well, or condition and columns represent repeated measurements.

Example result columns:

```text
A1, A2, B1, B2
```

### By plate groups (template)

This mode allows custom grouping directly in the app.

1. Select rows and columns.
2. Enter group labels into the template grid.
3. Cells with the same label are combined into one group.

This is useful for biological replicates, treatment groups, control wells, or custom plate layouts.

---

## Statistics and outlier handling

### Outlier methods

| Method | Parameter | Description |
| --- | --- | --- |
| None | Not applicable | Keeps all numeric values. |
| IQR | IQR multiplier | Removes values outside the interquartile range bounds. |
| Z-score | Z-score threshold | Removes values with high absolute z-scores. |
| MAD | MAD threshold | Removes values using median absolute deviation. |

### Error metrics

The app computes several metrics that can be used in tables or chart error bars:

- Standard deviation
- Standard error of the mean
- Approximate 95% confidence interval
- IQR/2
- MAD

### T-tests

PlateMaster uses Welch-style two-sample t-test calculations for comparing groups. This is useful when comparing conditions with unequal variance or unequal sample counts.

The p-values are approximate and should be interpreted as a quick exploratory analysis rather than a substitute for a full statistical workflow.

---

## Visualization options

PlateMaster builds interactive charts using Altair. Depending on your data shape, you can choose between time-oriented, column-oriented, and distribution-oriented plots.

Recommended chart choices:

| Goal | Suggested chart |
| --- | --- |
| Follow values over time | Line + points |
| Compare columns within each time point | Bar plot by time |
| Compare time points within each column | Bar plot by columns |
| Show distribution shape | Violin plot |
| Show spread and min/max | Box plot |
| Inspect individual aggregated points | Scatter plot |

---

## Exported report

The export button creates `plate_master_report.xlsx` with multiple sheets:

| Sheet | Contents |
| --- | --- |
| Aggregated table | Final combined table across files. |
| Statistical analysis | Descriptive statistics for the selected row/column combination. |
| T-test analysis | Welch t-test comparison between selected groups. |
| Pairwise T-test p-values | Matrix of approximate pairwise p-values. |
| Outlier log | Removed values with file, time, row, column, and value. |
| Outlier statistics | Removal counts and percentages. |
| Plot data | Data used to generate the selected chart. |

---

## Project structure

```text
PlateMaster.py   # Main Streamlit application and analysis logic
README.md        # Project documentation
```

The current implementation is contained in a single Python script. This makes the project easy to run, copy, and deploy, while still keeping data loading, cleaning, aggregation, statistics, plotting, and export logic together.

---

## Troubleshooting

### Streamlit is not installed

Install Streamlit:

```bash
pip install streamlit
```

Then run:

```bash
streamlit run PlateMaster.py
```

### Excel file cannot be parsed

Try the following:

1. Install optional Excel readers:

   ```bash
   pip install openpyxl xlrd python-calamine
   ```

2. If the file is an old `.xls`, install LibreOffice and try loading it again.
3. Open the file manually and save it as `.xlsx`.
4. Check that the table has a first identifier column and at least one numeric measurement column.

### No numeric columns detected

PlateMaster only analyzes columns that can be converted to numbers. Check that numeric cells do not contain incompatible text, symbols, or formatting artifacts.

### No time values available

Make sure your filenames contain a number. PlateMaster uses the last number in each filename as the time value.

Example:

```text
experiment_0.xlsx
experiment_10.xlsx
experiment_20.xlsx
```

### No chart appears

Check that:

- You clicked **Run**.
- At least one plot column is selected.
- The selected columns exist in the aggregated result.
- The filenames contain extractable time values if you are using time-based plots.

---

## Notes and limitations

- The app is optimized for exploratory plate analysis and fast reporting.
- Approximate p-values are intended for quick comparison and should be validated for publication-grade statistical analysis.
- Very large Excel or ZIP uploads may be limited by Streamlit server settings.
- Presets are stored in Streamlit session state, so they may reset when the session restarts.
- The first column of every input table is treated as the row identifier.

---

## License

No license file is included in this checkout. Add a license before distributing or reusing the project publicly.

---

## Citation / author

Repository: <https://github.com/narek-abelyan/PlateMaster>

If you use PlateMaster in your work, consider citing or linking to the repository so others can find the tool.
