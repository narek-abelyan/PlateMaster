# 🧫 PlateMaster

<p align="center">
  <a href="#-english-version">🇬🇧 English</a> •
  <a href="#-русская-версия">🇷🇺 Русский</a>
</p>

## 🇬🇧 English version

**PlateMaster** is an interactive **Streamlit** application for analyzing multi-well plate and plate-reader data stored in Excel files. It helps you load one workbook, multiple workbooks, or a ZIP archive, select plate rows and columns, aggregate measurements, remove outliers, compare groups statistically, visualize time-course results, and export a complete Excel report.

The project is particularly useful for researchers who need a fast, intuitive, and reproducible workflow without manually transferring data between spreadsheets.

> The current main application file is **`PlateMaster1.2.py`**.

---

## 📚 Table of contents

- [✨ Key features](#-key-features)
- [🧭 Typical workflow](#-typical-workflow)
- [📥 Input data format](#-input-data-format)
- [⚙️ Installation](#️-installation)
- [🚀 Running the app](#-running-the-app)
- [🧪 How to use PlateMaster](#-how-to-use-platemaster)
- [🧩 Analysis modes](#-analysis-modes)
- [📊 Statistics and outlier handling](#-statistics-and-outlier-handling)
- [🎨 Visualization options](#-visualization-options)
- [📤 Exported report](#-exported-report)
- [🗂 Project structure](#-project-structure)
- [🛠 Troubleshooting](#-troubleshooting)
- [⚠️ Notes and limitations](#️-notes-and-limitations)
---

## ✨ Key features

### 📁 Flexible data loading

PlateMaster supports several ways to bring experiment data into the app:

- load Excel files from a server or local folder path;
- upload one or many Excel files through the browser;
- upload a ZIP archive containing multiple Excel workbooks;
- read common Excel extensions: `.xlsx`, `.xls`, and `.xlsm`;
- try multiple Excel engines, including `openpyxl`, `xlrd`, and optional `python-calamine`;
- fall back to delimited text parsing for some instrument exports mislabeled as Excel files;
- optionally use **LibreOffice** conversion for old or problematic `.xls` files.

### 🧫 Plate-focused selection and aggregation

- select any subset of rows and columns from the loaded tables;
- aggregate values by original columns;
- aggregate values by original rows;
- aggregate values by custom plate-template groups;
- save and reapply analysis presets.

Available aggregation methods:

| Method | Purpose |
| --- | --- |
| Mean | average value |
| Median | median value |
| Min | minimum value |
| Max | maximum value |
| Standard deviation | standard deviation |
| Sum | total sum |

### 🧹 Outlier filtering

PlateMaster can remove outliers before aggregation using:

- **IQR** filtering;
- **Z-score** filtering;
- **MAD** filtering;
- **None** to keep all values.

The app also keeps a detailed log of removed values so results remain transparent and auditable.

### 📈 Statistical analysis

PlateMaster calculates:

- `n`;
- mean, median, mode;
- standard deviation;
- SEM;
- MAD;
- min/max;
- Student-t 95% confidence interval for the mean (bounds and half-width);
- IQR/2;
- two-sided one-sample Student t-test for a mean different from zero;
- Welch two-sample t-test between selected groups;
- pairwise p-value matrix across time/metric combinations.

### 🎨 Interactive charts

Charts are built with **Altair** and are designed for quick visual inspection:

- Line + points;
- Line + points by columns;
- Scatter plot;
- Bar plot by time;
- Bar plot by columns;
- Box plot by time;
- Box plot by columns;
- Violin plot by time;
- Violin plot by columns.

Optional error bars can display standard deviation, SEM, 95% CI, IQR/2, or MAD. SEM and 95% CI require Mean aggregation.

---

## 🧭 Typical workflow

1. Start the Streamlit app.
2. Upload Excel files or a ZIP archive.
3. Review the representative input table.
4. Select rows and columns to analyze.
5. Choose the aggregation layout and aggregation function.
6. Optionally enable outlier filtering.
7. Configure the color palette, charts, and error bars.
8. Press **Run**.
9. Inspect aggregated tables, statistics, t-tests, outlier logs, and charts.
10. Export the final Excel report.

---

## 📥 Input data format

PlateMaster expects each spreadsheet to contain a table where:

- the **first column** contains row identifiers such as well labels, sample names, or group labels;
- the remaining columns contain numeric measurements;
- column names may represent time points, wavelengths, treatments, concentrations, or other measurement dimensions.

Example:

| Well | 0 | 10 | 20 | 30 |
| --- | ---: | ---: | ---: | ---: |
| A1 | 0.10 | 0.15 | 0.22 | 0.31 |
| A2 | 0.11 | 0.16 | 0.24 | 0.33 |
| B1 | 0.08 | 0.12 | 0.19 | 0.27 |

During loading, the app renames the first column internally to `row_id` and converts numeric measurement columns automatically.

Open **Import options** before loading to select the worksheet by number (starting at 1) or exact name, and the header row (starting at 1). These choices apply to the whole batch. The default reads the first worksheet and first row. Reload files after changing these options.

By default, input must have unique plate-row letters (A–AF) and numbered measurement columns (1–48). **Auto / partial table** allows subsets; choose an **Expected plate** size to require the complete matrix and valid coordinates. For sample names, full well labels such as A1, or other measurement-column labels, explicitly enable **Allow custom table labels** and check the warning and preview. Duplicate/missing row identifiers and infinite measurements are rejected. Purely textual auxiliary columns, discarded empty measurements, and non-numeric cells are reported. Structural checks cannot determine the biological meaning of valid numeric cells.

Use **Preview file** to inspect each accepted table before analysis. Import errors identify skipped files. If a loaded file fails during aggregation, analysis stops with the filename and error, without exporting partial results. JSON settings and Excel reports record the import choices and warnings for each input; use those choices when reloading the original files.

### ⏱ Time extraction from filenames

For cross-file time-course analysis, PlateMaster extracts the **last number from each filename** and uses it as `Time`.

| Filename | Extracted time |
| --- | ---: |
| `plate_0.xlsx` | `0` |
| `experiment_10min.xlsx` | `10` |
| `sample_time_2.5.xls` | `2.5` |

If a filename does not contain a number, the file can still be loaded, but time-based statistics and plotting may be limited.

### Documented validation example

[`example_2`](example_2/README.md) is a CC0 synthetic validation set with five Excel files, a known plate map, reloadable analysis settings, exact instructions and fixed expected numerical results. Its regression test reads the distributed workbooks and compares PlateMaster output with the reference means, Student-t intervals and Welch test.

---

## ⚙️ Installation

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

```bash
pip install -r requirements.txt
```

If there is no `requirements.txt`, install the main dependencies manually:

```bash
pip install streamlit pandas numpy scipy altair openpyxl xlrd python-calamine
```

Optional, but useful for old or problematic `.xls` files: install **LibreOffice** so PlateMaster can attempt automatic conversion to `.xlsx`.

---

## 🚀 Running the app

From the repository directory, run:

```bash
streamlit run PlateMaster1.2.py
```

You can also run the script directly:

```bash
python PlateMaster1.2.py
```

When executed directly outside Streamlit, the script tries to restart itself through `streamlit run`.

---

## 🧪 How to use PlateMaster

### 1. Choose a data source

In the sidebar, select one of the available source modes.

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

### 4. Configure analysis settings

Select:

- aggregation layout;
- aggregation function;
- outlier filtering method;
- color scale mode;
- color palette;
- plot columns;
- chart type;
- error bar metric.

### 5. Save or load presets

Presets now include the plate-group map, selected rows/columns, aggregation and outlier settings, chart/color choices, and the selected descriptive-statistics and t-test groups.

1. Load your spreadsheets, choose the settings and (if needed) fill the plate-group grid.
2. Click **Run**, then choose the statistics, comparison groups and plot options to save.
3. Click **Download settings.json**. Keep this file **and the original spreadsheets**; the JSON does not embed raw data.
4. In a new session, load the spreadsheets first. Choose the JSON under **Import analysis settings (.json)** and click **Apply imported settings**.
5. Review the input/environment check, then click **Run**. Matching inputs and settings reproduce the calculation; changed files or software produce a warning. Missing selected rows/columns or invalid settings are rejected without partially applying them.

**Save preset now / Apply preset** remain shortcuts within the current session. Download JSON to preserve a preset across restarts. Import does not silently start an analysis. Settings using the removed `None` aggregation are rejected; choose a supported aggregation and save a new file.

Input records contain original-file SHA-256 checksums, parsed-table checksums and the currently inferred filename times. Time is still extracted from filenames; its unit is recorded as unspecified. JSON saving does not add a new time-entry feature.

### 6. Run analysis

Click **Run** to calculate results for the current settings.

---

## 🧩 Analysis modes

### By original columns

This mode aggregates selected rows within each selected column. It is useful when columns represent time points, wavelengths, or measurement channels.

```text
0, 10, 20, 30
```

### By original rows

This mode aggregates selected columns within each selected row. It is useful when each row represents a sample, well, or condition and columns represent repeated measurements.

```text
A1, A2, B1, B2
```

### By plate groups (template)

This mode allows custom grouping directly in the app or through a reusable CSV plate map.

1. Select rows and columns.
2. Enter group labels into the template grid, or upload a CSV with the columns `row_id,column,group` and click **Apply plate template**.
3. Cells with the same label are combined into one group.
4. Click **Download plate template.csv** to save the current map. The exported file includes every selected well; unassigned wells have a blank `group` value.

Uploaded templates are checked for duplicate wells, unknown row or column identifiers, and missing coordinates. This is useful for biological replicates, treatment groups, control wells, or custom plate layouts. The plate map can be reused independently of the complete analysis-settings JSON.

---

## 📊 Statistics and outlier handling

### Outlier methods

| Method | Parameter | Description |
| --- | --- | --- |
| None | Not applicable | Keeps all numeric values. |
| IQR | IQR multiplier | Removes values outside the interquartile range bounds. |
| Z-score | Z-score threshold | Removes values with high absolute z-scores. |
| MAD | MAD threshold | Removes values using median absolute deviation. |

### Error metrics

The app computes several metrics that can be used in tables or chart error bars:

- standard deviation;
- standard error of the mean;
- Student-t 95% confidence interval for the mean (bounds and half-width);
- IQR/2;
- MAD.

### T-tests

PlateMaster uses `scipy.stats.ttest_ind(..., equal_var=False)` for the two-sided Welch test. Both the selected-group result and the pairwise matrix use the same calculation and the t distribution with Welch degrees of freedom. The mean-versus-zero result uses `scipy.stats.ttest_1samp`.

**Optional pairwise comparisons:** the matrix is off by default. Enable **Enable pairwise comparisons**, then explicitly select 2–20 time/metric groups (at most 190 tests). Only the lower triangle is filled; self-comparisons and the mirrored upper triangle are omitted. Values below 0.005 are shown as `p < 0.005`; other values are rounded to three decimal places. P-values are unadjusted for multiple comparisons; a warning is shown in the interface and included in the Excel report. No Holm or other multiplicity correction is applied. The enable flag and selected groups are saved in presets and JSON. Older JSON files load with this option off. The pairwise sheets are exported only when comparisons have been calculated.


- Missing/non-numeric observations are omitted; `n` counts available values. Infinite values are rejected.
- Sample SD uses `ddof=1`; SEM is SD / sqrt(n). The 95% CI half-width is `t.ppf(0.975, n-1) * SEM`, and bounds are `mean ± half-width`. Every aggregation layout uses this definition.
- With fewer than two observations, SD, SEM, confidence intervals and inferential p-values are unavailable (`NaN` / blank in Excel), not falsely reported as zero.
- A constant one-sample group, or two Welch groups that both have zero sample variance, receives an undefined-test status and no p-value. A single constant group compared with a variable group is supported. For a constant sample with n ≥ 2 the calculated mean interval collapses to its observed mean; this is not evidence that the underlying population has zero uncertainty.
- SEM and 95% CI chart options are available only with **Mean** aggregation. They describe uncertainty of a mean, not a sum, median or extreme. Undefined error bars are omitted.
- Tests assume independent observations. Technical wells from one sample are not automatically independent biological replicates. Pairwise p-values are still **unadjusted** for multiple testing.

Numerical checks: `[1, 2]` versus `[4, 5]` gives Welch p = **0.0513167019494862**; `[1, 2]` gives a mean-CI half-width of **6.353102368087347**. These differ deliberately from the old normal approximation.

The `None` aggregation has been removed in all layouts. `None` remains available to disable outlier filtering or chart error bars. Other aggregation functions are unchanged; `Std` aggregation still uses `ddof=0`.

---

## 🎨 Visualization options

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

## 📤 Exported report

The export button creates **`plate_master_report.xlsx`** with multiple sheets:

| Sheet | Contents |
| --- | --- |
| Aggregated table | Final combined table across files. |
| Statistical analysis | Descriptive statistics for the selected row/column combination. |
| T-test analysis | Welch t-test comparison between selected groups. |
| Pairwise T-test p-values | Optional lower-triangular matrix of formatted, unadjusted Welch p-values, with row labels. |
| Pairwise notes | Interpretation and unadjusted-p-value warning; included with the matrix. |
| Outlier log | Removed values with file, time, row, column, and value. |
| Outlier statistics | Removal counts and percentages. |
| Plot data | Data used to generate the selected chart. |
| Analysis settings | All effective analysis, chart and statistical-selection settings. |
| Plate groups | Saved row/column/group map. |
| Input files | Original-file and parsed-table SHA-256 checksums, names and inferred times. |
| Software versions | Application version, Git revision, modified-tree flag, actual source checksums, Python and libraries. |
| Settings JSON | Exact reloadable JSON, split into numbered parts to respect Excel cell limits. |

To recover JSON from Excel, join the `JSON` cells in `Settings JSON` in ascending `Part` order, then save as UTF-8 `.json`. Source checksums identify local code changes even before a new Git commit is made.

---

## 🗂 Project structure

```text
PlateMaster1.2.py   # Main Streamlit application and analysis logic
analysis_settings.py # Settings validation, JSON and provenance
example_2/          # CC0 documented validation dataset and expected results
tests/              # Numerical, settings and Streamlit regression tests
requirements-dev.txt # Dependencies for running tests
requirements.txt    # Python dependencies
README.md           # English and Russian documentation
LICENSE             # Full MIT License text
```

The Streamlit entry point remains `PlateMaster1.2.py`. Settings validation and provenance are kept in `analysis_settings.py`; this revision does not perform a full architectural refactor. The stable release is identified as `1.3.0` inside the app and exported settings. The immutable release page is [v1.3.0](https://github.com/narek-abelyan/PlateMaster/releases/tag/v1.3.0), and citation metadata are provided in [`CITATION.cff`](CITATION.cff).

### Run the regression checks

The [Tests workflow](https://github.com/narek-abelyan/PlateMaster/actions/workflows/tests.yml) runs the test suite automatically on pushes and pull requests, and can also be started manually from the Actions tab. Its matrix includes Linux (Ubuntu) and Windows with Python 3.12 and 3.13. Each job installs the test dependencies, checks their consistency, records their resolved versions, and runs `python -m pytest -q`. Check the individual jobs for the result of a particular revision; a passing run confirms the covered scenarios, not every possible input or workflow.

Tested locally on Linux with Python 3.13.12, Streamlit 1.63.0, pandas 2.3.3, SciPy 1.17.1, openpyxl 3.1.5, xlrd 2.0.2 and python-calamine 0.8.2. These are recorded test-environment versions, not a claim that all platforms have been validated. SciPy ≥ 1.11 and Streamlit ≥ 1.49 are required for APIs used here.

```bash
python -m pip install -r requirements-dev.txt
python -m pytest -q
```

Tests cover known Welch/CI results, missing and degenerate samples, every aggregation layout, rejected settings, saved plate groups, changed-input detection, named presets, fresh-session JSON replay, and Excel provenance. Interface tests simulate the uploaded JSON bytes because Streamlit AppTest does not provide a browser file-picker API.

The distributed validation dataset can be checked separately with `python -m pytest -q tests/test_validation_example.py`.

---

## 🛠 Troubleshooting

### Streamlit is not installed

```bash
pip install streamlit
streamlit run PlateMaster1.2.py
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

```text
experiment_0.xlsx
experiment_10.xlsx
experiment_20.xlsx
```

### No chart appears

Check that:

- you clicked **Run**;
- at least one plot column is selected;
- the selected columns exist in the aggregated result;
- filenames contain extractable time values if you are using time-based plots.

---

## ⚠️ Notes and limitations

- The app is optimized for exploratory plate analysis and fast reporting.
- Correct t-test calculations do not replace checking independence, study design and the effect of multiple comparisons.
- Very large Excel or ZIP uploads may be limited by Streamlit server settings.
- Named presets are session-local; download settings JSON and keep the original files for use after restart.
- The first column of every input table is treated as the row identifier.

---

## 📝 License

This project is distributed under the **MIT License**. See the full license text in [`LICENSE`](LICENSE).

---

## 🙌 Citation / author

Repository: <https://github.com/narek-abelyan/PlateMaster>

If you use PlateMaster in your work, consider citing or linking to the repository so others can find the tool.

---

## 🇷🇺 Русская версия

**PlateMaster** — это интерактивное приложение на **Streamlit** для анализа данных многолуночных планшетов и Excel-файлов с результатами plate-reader экспериментов. Оно помогает быстро загружать один файл, набор файлов или ZIP-архив, выбирать строки и столбцы планшета, объединять измерения, удалять выбросы, сравнивать группы статистически, строить графики временных рядов и выгружать готовый Excel-отчёт.

Проект особенно полезен исследователям, которым нужен быстрый, наглядный и воспроизводимый workflow без ручного копирования данных между таблицами.

> Текущая основная версия приложения находится в файле **`PlateMaster1.2.py`**.

---

## 📚 Содержание

- [✨ Возможности](#-возможности)
- [🧭 Типичный сценарий работы](#-типичный-сценарий-работы)
- [📥 Формат входных данных](#-формат-входных-данных)
- [⚙️ Установка](#️-установка)
- [🚀 Запуск приложения](#-запуск-приложения)
- [🧪 Как пользоваться PlateMaster](#-как-пользоваться-platemaster)
- [🧩 Режимы анализа](#-режимы-анализа)
- [📊 Статистика и выбросы](#-статистика-и-выбросы)
- [🎨 Визуализация](#-визуализация)
- [📤 Excel-отчёт](#-excel-отчёт)
- [🗂 Структура проекта](#-структура-проекта)
- [🛠 Решение проблем](#-решение-проблем)
- [⚠️ Ограничения](#️-ограничения)
---

## ✨ Возможности

### 📁 Гибкая загрузка данных

PlateMaster умеет работать с разными источниками данных:

- загрузка Excel-файлов из папки на сервере или локальной машине;
- загрузка одного или нескольких Excel-файлов через браузер;
- загрузка ZIP-архива с несколькими Excel-книгами;
- поддержка расширений `.xlsx`, `.xls`, `.xlsm`;
- несколько стратегий чтения Excel-файлов: `openpyxl`, `xlrd`, опционально `python-calamine`;
- fallback для некоторых экспортов приборов, которые выглядят как `.xls`, но фактически являются текстовыми таблицами;
- опциональная конвертация через **LibreOffice** для старых или проблемных `.xls` файлов.

### 🧫 Работа с планшетной логикой

- выбор любых строк и столбцов из загруженной таблицы;
- агрегация по исходным столбцам;
- агрегация по исходным строкам;
- агрегация по пользовательским группам в шаблоне планшета;
- сохранение и применение пресетов анализа.

Доступные функции агрегации:

| Метод | Назначение |
| --- | --- |
| Mean | среднее значение |
| Median | медиана |
| Min | минимум |
| Max | максимум |
| Standard deviation | стандартное отклонение |
| Sum | сумма |

### 🧹 Фильтрация выбросов

Перед агрегацией можно удалить выбросы одним из методов:

- **IQR** — фильтрация по межквартильному размаху;
- **Z-score** — фильтрация по z-оценке;
- **MAD** — фильтрация по median absolute deviation;
- **None** — без удаления выбросов.

Приложение сохраняет подробный журнал удалённых значений, чтобы результаты оставались прозрачными и проверяемыми.

### 📈 Статистический анализ

PlateMaster рассчитывает:

- количество значений `n`;
- mean, median, mode;
- standard deviation;
- SEM;
- MAD;
- min/max;
- 95% доверительный интервал среднего по t-распределению (границы и полуширина);
- IQR/2;
- двусторонний одновыборочный t-тест для отличия среднего от нуля;
- Welch two-sample t-test между выбранными группами;
- матрицу попарных p-value для комбинаций времени и метрик.

### 🎨 Интерактивные графики

Графики строятся на **Altair** и помогают быстро увидеть динамику и различия между группами:

- Line + points;
- Line + points by columns;
- Scatter plot;
- Bar plot by time;
- Bar plot by columns;
- Box plot by time;
- Box plot by columns;
- Violin plot by time;
- Violin plot by columns.

Для error bars можно выбрать standard deviation, SEM, 95% CI, IQR/2 или MAD. SEM и 95% CI доступны только при агрегации Mean.

---

## 🧭 Типичный сценарий работы

1. Запустить Streamlit-приложение.
2. Загрузить Excel-файлы или ZIP-архив.
3. Проверить representative input table.
4. Выбрать нужные строки и столбцы.
5. Выбрать режим агрегации и функцию агрегации.
6. При необходимости включить фильтрацию выбросов.
7. Настроить цветовую палитру, графики и error bars.
8. Нажать **Run**.
9. Просмотреть агрегированную таблицу, статистику, t-test, журнал выбросов и графики.
10. Скачать итоговый Excel-отчёт.

---

## 📥 Формат входных данных

PlateMaster ожидает таблицу, где:

- **первый столбец** содержит идентификаторы строк: well labels, sample names, group names или другие названия;
- остальные столбцы содержат числовые измерения;
- названия столбцов могут обозначать время, длину волны, концентрацию, обработку, канал измерения и т.д.

Пример:

| Well | 0 | 10 | 20 | 30 |
| --- | ---: | ---: | ---: | ---: |
| A1 | 0.10 | 0.15 | 0.22 | 0.31 |
| A2 | 0.11 | 0.16 | 0.24 | 0.33 |
| B1 | 0.08 | 0.12 | 0.19 | 0.27 |

При загрузке первый столбец внутренне переименовывается в `row_id`, а числовые столбцы автоматически конвертируются в числа.

Перед загрузкой откройте **Import options**: укажите номер листа (с 1) или его точное имя и номер строки заголовка (с 1). Настройки общие для всей партии. По умолчанию читаются первый лист и первая строка; после изменения параметров загрузите файлы заново.

По умолчанию нужны уникальные буквы строк планшета (A–AF) и номера столбцов (1–48). **Auto / partial table** допускает часть планшета; конкретный размер в **Expected plate** требует полной матрицы с правильными координатами. Для названий образцов, обозначений A1 и других заголовков явно включите **Allow custom table labels** и проверьте предупреждение и просмотр данных. Повторяющиеся и пустые идентификаторы строк, а также бесконечные значения отклоняются. Программа сообщает об исключённых текстовых служебных столбцах, пустых измерениях и нечисловых ячейках. Проверка структуры не определяет биологический смысл чисел.

В **Preview file** можно просмотреть каждый принятый файл до анализа. Ошибки загрузки перечисляют пропущенные файлы. При ошибке агрегации программа показывает имя файла и причину и останавливается без экспорта частичных результатов. JSON и Excel сохраняют параметры импорта и предупреждения для каждого файла; используйте эти параметры при повторной загрузке исходных данных.

### ⏱ Извлечение времени из имени файла

Для анализа временных рядов PlateMaster берёт **последнее число из имени файла** и использует его как `Time`.

| Имя файла | Извлечённое время |
| --- | ---: |
| `plate_0.xlsx` | `0` |
| `experiment_10min.xlsx` | `10` |
| `sample_time_2.5.xls` | `2.5` |

Если в имени файла нет числа, файл всё равно может быть загружен, но time-based статистика и графики могут быть ограничены.

### Документированный проверочный пример

[`example_2`](example_2/README.md) — синтетический набор под лицензией CC0: пять Excel-файлов, известная карта планшета, готовые настройки, точная инструкция и зафиксированные ожидаемые численные результаты. Регрессионный тест читает приложенные файлы и сравнивает результаты PlateMaster с эталонными средними, интервалами Стьюдента и тестом Уэлча.

---

## ⚙️ Установка

### 1. Клонировать репозиторий

```bash
git clone https://github.com/narek-abelyan/PlateMaster.git
cd PlateMaster
```

### 2. Создать виртуальное окружение

```bash
python -m venv .venv
source .venv/bin/activate
```

Windows PowerShell:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

### 3. Установить зависимости

```bash
pip install -r requirements.txt
```

Если файла `requirements.txt` нет, установите основные зависимости вручную:

```bash
pip install streamlit pandas numpy scipy altair openpyxl xlrd python-calamine
```

Опционально установите **LibreOffice**, если нужно открывать старые или повреждённые `.xls` файлы.

---

## 🚀 Запуск приложения

Из папки репозитория выполните:

```bash
streamlit run PlateMaster1.2.py
```

Также можно запустить Python-файл напрямую:

```bash
python PlateMaster1.2.py
```

При прямом запуске скрипт пытается перезапустить себя через `streamlit run`.

---

## 🧪 Как пользоваться PlateMaster

### 1. Выберите источник данных

В sidebar доступны два режима:

#### Server folder path

Используйте этот режим, если Excel-файлы уже лежат на сервере или машине, где запущен Streamlit.

1. Укажите путь к папке.
2. Нажмите **Load folder**.
3. PlateMaster найдёт файлы `.xlsx`, `.xls`, `.xlsm`.

#### Local PC upload

Используйте этот режим для загрузки через браузер.

1. Загрузите один или несколько Excel-файлов.
2. При необходимости загрузите ZIP-архив.
3. Нажмите **Load uploaded files**.

### 2. Проверьте representative table

После загрузки приложение показывает пример входной таблицы. Проверьте, что строки, столбцы и числовые значения распознаны корректно.

### 3. Выберите строки и столбцы

Через multiselect в sidebar выберите rows и columns, которые должны попасть в анализ.

### 4. Настройте анализ

Выберите:

- aggregation layout;
- aggregation function;
- outlier filtering method;
- color scale mode;
- color palette;
- plot columns;
- chart type;
- error bar metric.

### 5. Сохраните пресет

Пресет сохраняет карту групп лунок, выбранные строки и столбцы, агрегацию, фильтры, оформление графиков, а также выбранные группы для описательной статистики и t-теста.

1. Загрузите таблицы, настройте анализ и при необходимости заполните карту групп.
2. Нажмите **Run**, затем выберите статистику, сравниваемые группы и параметры графика.
3. Нажмите **Download settings.json**. Сохраните JSON **вместе с исходными таблицами**: сами измерения в JSON не включаются.
4. В новой сессии сначала загрузите таблицы. Выберите JSON в **Import analysis settings (.json)** и нажмите **Apply imported settings**.
5. Проверьте сообщение о файлах и окружении, затем нажмите **Run**. При изменении файлов или версий программа предупреждает, что это не точное повторение. Отсутствие нужных строк/столбцов и некорректные настройки приводят к отказу от загрузки без частичного применения.

**Save preset now / Apply preset** работают внутри текущей сессии. Для хранения после перезапуска скачайте JSON. Импорт не запускает расчёт автоматически. Настройки с удалённой агрегацией `None` отклоняются: выберите поддерживаемую функцию и сохраните новый файл.

В сведения о входах входят контрольные суммы SHA-256 исходных файлов, контрольные суммы прочитанных таблиц и время, извлечённое из имени. Единицы времени пока не задаются и отмечаются как неизвестные. Сохранение JSON не меняет существующий способ определения времени.

### 6. Запустите расчёт

Нажмите **Run**, чтобы получить таблицы, статистику и графики.

---

## 🧩 Режимы анализа

### By original columns

Агрегирует выбранные строки внутри каждого выбранного столбца. Удобно, когда столбцы — это время, длины волн или каналы измерения.

```text
0, 10, 20, 30
```

### By original rows

Агрегирует выбранные столбцы внутри каждой выбранной строки. Удобно, когда строки — это лунки, образцы или условия, а столбцы — повторы или измерения.

```text
A1, A2, B1, B2
```

### By plate groups (template)

Позволяет создать группы прямо в интерфейсе или загрузить готовую карту планшета в CSV:

1. Выберите строки и столбцы.
2. Введите названия групп в template grid либо загрузите CSV с колонками `row_id,column,group` и нажмите **Apply plate template**.
3. Ячейки с одинаковым названием объединяются в одну группу.
4. Нажмите **Download plate template.csv**, чтобы сохранить текущую карту. В файл входят все выбранные лунки; у неназначенных лунок поле `group` остаётся пустым.

При загрузке проверяются повторяющиеся лунки, неизвестные строки и столбцы, а также пустые координаты. Карту можно использовать повторно независимо от полного JSON с настройками анализа.

---

## 📊 Статистика и выбросы

### Методы удаления выбросов

| Метод | Параметр | Описание |
| --- | --- | --- |
| None | — | оставить все числовые значения |
| IQR | IQR multiplier | удалить значения вне границ межквартильного размаха |
| Z-score | threshold | удалить значения с высокой абсолютной z-оценкой |
| MAD | threshold | удалить значения по median absolute deviation |

### Метрики ошибок

- Standard deviation;
- Standard error of the mean;
- 95% CI for the mean using Student’s t distribution;
- IQR/2;
- MAD.

### T-tests

Для двустороннего теста Уэлча используется `scipy.stats.ttest_ind(..., equal_var=False)`. Отдельное сравнение и попарная матрица используют один расчёт с t-распределением и степенями свободы Уэлча. Для сравнения среднего с нулём используется `scipy.stats.ttest_1samp`.

**Необязательные попарные сравнения:** матрица выключена по умолчанию. Включите **Enable pairwise comparisons** и выберите от 2 до 20 комбинаций времени и метрики (не более 190 тестов). Заполняется только нижняя половина матрицы; диагональ и зеркальная верхняя половина остаются пустыми. Значения меньше 0,005 показываются как `p < 0.005`, остальные округляются до трёх знаков после точки. P-value не скорректированы на множественные проверки: предупреждение есть в интерфейсе и Excel. Поправка Холма и другие поправки не применяются. Включение и выбор групп сохраняются в пресетах и JSON. Старые JSON загружаются с выключенной матрицей. Листы попарных сравнений экспортируются только после расчёта.

- Пропущенные и нечисловые наблюдения исключаются; `n` отражает число доступных значений. Бесконечные значения отклоняются.
- Выборочное SD рассчитывается с `ddof=1`; SEM = SD / sqrt(n). Полуширина 95% интервала среднего равна `t.ppf(0.975, n-1) * SEM`, границы — `mean ± полуширина`. Формула одинакова для всех способов группировки.
- При n < 2 SD, SEM, доверительный интервал и p-value недоступны (`NaN` / пустая ячейка в Excel), а не равны нулю.
- Для одной постоянной выборки и для двух групп Уэлча с нулевой дисперсией в обеих группах возвращается пояснение и неопределённый p-value. Сравнение постоянной группы с переменной поддерживается. При n ≥ 2 интервал постоянной выборки сводится к наблюдаемому среднему; это не доказывает отсутствие неопределённости в исходной популяции.
- SEM и 95% CI на графиках доступны только для **Mean**. Это погрешность среднего, а не суммы, медианы или экстремума. Неопределённые полосы погрешности не рисуются.
- Тесты предполагают независимые наблюдения. Технические лунки одной пробы не становятся автоматически независимыми биологическими повторами. Попарные p-value пока **не корректируются** на множество сравнений.

Контрольные числа: для `[1, 2]` и `[4, 5]` p-value Уэлча = **0,0513167019494862**; для `[1, 2]` полуширина доверительного интервала среднего = **6,353102368087347**. Отличие от прежнего нормального приближения ожидаемо.

Агрегация `None` удалена во всех режимах. Вариант `None` для отключения фильтра выбросов и полос погрешности сохранён. Остальные агрегации не менялись; `Std` по-прежнему использует `ddof=0`.

---

## 🎨 Визуализация

| Цель | Рекомендуемый график |
| --- | --- |
| Проследить динамику во времени | Line + points |
| Сравнить столбцы внутри time point | Bar plot by time |
| Сравнить time points внутри столбца | Bar plot by columns |
| Показать форму распределения | Violin plot |
| Показать разброс и min/max | Box plot |
| Посмотреть отдельные агрегированные точки | Scatter plot |

---

## 📤 Excel-отчёт

Кнопка экспорта создаёт файл **`plate_master_report.xlsx`** с несколькими листами:

| Лист | Содержимое |
| --- | --- |
| Aggregated table | итоговая агрегированная таблица по всем файлам |
| Statistical analysis | descriptive statistics |
| T-test analysis | Welch t-test между выбранными группами |
| Pairwise T-test p-values | нижняя половина матрицы нескорректированных p-value с подписями строк; значения меньше 0,005 показаны как `p < 0.005`, остальные округлены до трёх знаков |
| Pairwise notes | пояснение матрицы и предупреждение об отсутствии поправки |
| Outlier log | удалённые значения с файлом, временем, строкой, столбцом и значением |
| Outlier statistics | количество и процент удалённых значений |
| Plot data | данные, использованные для графика |
| Analysis settings | настройки анализа, графиков и выбранных статистических сравнений |
| Plate groups | карта строк, столбцов и групп лунок |
| Input files | имена, SHA-256 исходных файлов и прочитанных таблиц, извлечённое время |
| Software versions | версия программы, Git-ревизия, признак изменённого кода, его контрольные суммы, Python и библиотеки |
| Settings JSON | точный JSON настроек, разбитый на части из-за ограничения размера ячейки Excel |

Для восстановления JSON из Excel объедините ячейки `JSON` листа `Settings JSON` по возрастанию `Part` и сохраните результат в UTF-8 `.json`. Контрольные суммы исходников позволяют отличить локально изменённый код даже до нового Git-коммита.

---

## 🗂 Структура проекта

```text
PlateMaster1.2.py   # Main Streamlit application and analysis logic
analysis_settings.py # Settings validation, JSON and provenance
tests/              # Numerical, settings and Streamlit regression tests
requirements-dev.txt # Dependencies for running tests
requirements.txt    # Python dependencies
README.md           # English and Russian documentation
LICENSE             # Full MIT License text
```

Точка запуска остаётся `PlateMaster1.2.py`. Проверка настроек и сведения об анализе вынесены в `analysis_settings.py`; полного разделения архитектуры эта ревизия не выполняет. В приложении и экспортируемых настройках указана стабильная версия `1.3.0`. Неизменяемая страница релиза: [v1.3.0](https://github.com/narek-abelyan/PlateMaster/releases/tag/v1.3.0); данные для цитирования находятся в [`CITATION.cff`](CITATION.cff).

### Запуск проверок

[Проверка Tests на GitHub](https://github.com/narek-abelyan/PlateMaster/actions/workflows/tests.yml) автоматически запускает тесты при отправке изменений и создании или обновлении pull request. Её также можно запустить вручную на вкладке Actions. Проверки настроены для Linux (Ubuntu) и Windows с Python 3.12 и 3.13. Каждое задание устанавливает зависимости тестов, проверяет их совместимость, записывает установленные версии и выполняет `python -m pytest -q`. Результат конкретной версии кода виден в заданиях: успешный запуск подтверждает проверенные сценарии, а не все возможные данные и действия.

Локально проверено на Linux: Python 3.13.12, Streamlit 1.63.0, pandas 2.3.3, SciPy 1.17.1, openpyxl 3.1.5, xlrd 2.0.2 и python-calamine 0.8.2. Это сведения о проверенном окружении, а не заявление о проверке всех платформ. Используемые API требуют SciPy ≥ 1.11 и Streamlit ≥ 1.49.

```bash
python -m pip install -r requirements-dev.txt
python -m pytest -q
```

Тесты проверяют известные результаты Уэлча и доверительных интервалов, пропуски и вырожденные выборки, все способы группировки, отказ от неверных настроек, сохранение групп, изменение входных файлов, пресеты, повторение JSON в новой сессии и состав Excel-отчёта. Загрузка JSON в тестах интерфейса имитируется на уровне полученных байтов: Streamlit AppTest не управляет браузерным выбором файла.

Отдельный проверочный набор запускается командой `python -m pytest -q tests/test_validation_example.py`.

---

## 🛠 Решение проблем

### Streamlit не установлен

```bash
pip install streamlit
streamlit run PlateMaster1.2.py
```

### Excel-файл не читается

Попробуйте:

1. Установить дополнительные reader-пакеты:

   ```bash
   pip install openpyxl xlrd python-calamine
   ```

2. Для старых `.xls` установить LibreOffice.
3. Открыть файл вручную и сохранить как `.xlsx`.
4. Проверить, что первый столбец содержит идентификаторы, а остальные — числовые значения.

### No numeric columns detected

Проверьте, что числовые ячейки не содержат лишний текст, символы или нестандартное форматирование.

### Нет time values

Убедитесь, что имена файлов содержат число:

```text
experiment_0.xlsx
experiment_10.xlsx
experiment_20.xlsx
```

### График не появляется

Проверьте, что:

- нажата кнопка **Run**;
- выбран хотя бы один plot column;
- выбранные столбцы есть в aggregated result;
- для time-based plots имена файлов содержат извлекаемое время.

---

## ⚠️ Ограничения

- Приложение оптимизировано для exploratory plate analysis и быстрого отчёта.
- Корректный t-тест не заменяет проверку независимости наблюдений, дизайна опыта и множества сравнений.
- Очень большие Excel/ZIP файлы могут быть ограничены настройками Streamlit-сервера.
- Именованные пресеты остаются в текущей сессии; для работы после перезапуска скачайте JSON и сохраните исходные файлы.
- Первый столбец каждого входного файла всегда считается идентификатором строки.

---

## 📝 Лицензия

Этот проект распространяется под лицензией **MIT License**. Полный текст лицензии находится в файле [`LICENSE`](LICENSE).

---

## 🙌 Автор / цитирование

Repository: <https://github.com/narek-abelyan/PlateMaster>

Если PlateMaster помог в вашей работе, укажите ссылку на репозиторий, чтобы другие пользователи могли найти инструмент.

---
