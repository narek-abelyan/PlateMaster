# 🧫 PlateMaster

## 🇬🇧 English version
**PlateMaster** is an interactive **Streamlit** application for analyzing multi-well plate and plate-reader data stored in Excel files. It helps you load one workbook, multiple workbooks, or a ZIP archive, select plate rows and columns, aggregate measurements, remove outliers, compare groups statistically, visualize time-course results, and export a complete Excel report.

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
- [🇬🇧 English version](#-english-version)

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
| None | оставить значения без агрегации |
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
- приблизительный 95% доверительный интервал;
- IQR/2;
- приблизительный p-value для отличия среднего от нуля;
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

Для error bars можно выбрать standard deviation, SEM, 95% CI, IQR/2 или MAD.

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

### ⏱ Извлечение времени из имени файла

Для анализа временных рядов PlateMaster берёт **последнее число из имени файла** и использует его как `Time`.

| Имя файла | Извлечённое время |
| --- | ---: |
| `plate_0.xlsx` | `0` |
| `experiment_10min.xlsx` | `10` |
| `sample_time_2.5.xls` | `2.5` |

Если в имени файла нет числа, файл всё равно может быть загружен, но time-based статистика и графики могут быть ограничены.

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
pip install streamlit pandas numpy altair matplotlib openpyxl xlrd python-calamine
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

Presets помогают сохранить часто используемые настройки: строки, столбцы, режим агрегации, фильтрацию, тип графика и цветовую палитру.

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

Позволяет создать группы прямо в интерфейсе:

1. Выберите строки и столбцы.
2. Введите названия групп в template grid.
3. Ячейки с одинаковым названием объединяются в одну группу.

Это удобно для biological replicates, treatment groups, control wells и пользовательских plate layouts.

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
- Approximate 95% CI;
- IQR/2;
- MAD.

### T-tests

PlateMaster использует Welch-style two-sample t-test для сравнения групп. P-values являются приблизительными и подходят для exploratory analysis; для публикационных выводов рекомендуется отдельная статистическая проверка.

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
| Pairwise T-test p-values | матрица попарных p-values |
| Outlier log | удалённые значения с файлом, временем, строкой, столбцом и значением |
| Outlier statistics | количество и процент удалённых значений |
| Plot data | данные, использованные для графика |

---

## 🗂 Структура проекта

```text
PlateMaster1.2.py   # Main Streamlit application and analysis logic
requirements.txt    # Python dependencies
README.md           # Russian and English documentation
```

Приложение сейчас находится в одном Python-файле, что упрощает запуск, перенос и деплой.

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
- [🇷🇺 Русская версия](#-русская-версия)

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
| None | keep values without aggregation |
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
- approximate 95% confidence interval;
- IQR/2;
- approximate p-value for mean different from zero;
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

Optional error bars can display standard deviation, SEM, 95% CI, IQR/2, or MAD.

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

### ⏱ Time extraction from filenames

For cross-file time-course analysis, PlateMaster extracts the **last number from each filename** and uses it as `Time`.

| Filename | Extracted time |
| --- | ---: |
| `plate_0.xlsx` | `0` |
| `experiment_10min.xlsx` | `10` |
| `sample_time_2.5.xls` | `2.5` |

If a filename does not contain a number, the file can still be loaded, but time-based statistics and plotting may be limited.

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
pip install streamlit pandas numpy altair matplotlib openpyxl xlrd python-calamine
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

Presets store frequently used analysis settings, including selected rows, columns, aggregation mode, filtering settings, chart options, and color options.

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

This mode allows custom grouping directly in the app.

1. Select rows and columns.
2. Enter group labels into the template grid.
3. Cells with the same label are combined into one group.

This is useful for biological replicates, treatment groups, control wells, or custom plate layouts.

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
- approximate 95% confidence interval;
- IQR/2;
- MAD.

### T-tests

PlateMaster uses Welch-style two-sample t-test calculations for comparing groups. The p-values are approximate and should be interpreted as quick exploratory analysis rather than a substitute for a full statistical workflow.

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
| Pairwise T-test p-values | Matrix of approximate pairwise p-values. |
| Outlier log | Removed values with file, time, row, column, and value. |
| Outlier statistics | Removal counts and percentages. |
| Plot data | Data used to generate the selected chart. |

---

## 🗂 Project structure

```text
PlateMaster1.2.py   # Main Streamlit application and analysis logic
requirements.txt    # Python dependencies
README.md           # English and Russian documentation
LICENSE             # Full MIT License text
```

The current implementation is contained in a single Python script. This makes the project easy to run, copy, and deploy, while still keeping data loading, cleaning, aggregation, statistics, plotting, and export logic together.

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
- Approximate p-values are intended for quick comparison and should be validated for publication-grade statistical analysis.
- Very large Excel or ZIP uploads may be limited by Streamlit server settings.
- Presets are stored in Streamlit session state, so they may reset when the session restarts.
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
- [🇬🇧 English version](#-english-version)

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
| None | оставить значения без агрегации |
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
- приблизительный 95% доверительный интервал;
- IQR/2;
- приблизительный p-value для отличия среднего от нуля;
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

Для error bars можно выбрать standard deviation, SEM, 95% CI, IQR/2 или MAD.

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

### ⏱ Извлечение времени из имени файла

Для анализа временных рядов PlateMaster берёт **последнее число из имени файла** и использует его как `Time`.

| Имя файла | Извлечённое время |
| --- | ---: |
| `plate_0.xlsx` | `0` |
| `experiment_10min.xlsx` | `10` |
| `sample_time_2.5.xls` | `2.5` |

Если в имени файла нет числа, файл всё равно может быть загружен, но time-based статистика и графики могут быть ограничены.

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
pip install streamlit pandas numpy altair matplotlib openpyxl xlrd python-calamine
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

Presets помогают сохранить часто используемые настройки: строки, столбцы, режим агрегации, фильтрацию, тип графика и цветовую палитру.

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

Позволяет создать группы прямо в интерфейсе:

1. Выберите строки и столбцы.
2. Введите названия групп в template grid.
3. Ячейки с одинаковым названием объединяются в одну группу.

Это удобно для biological replicates, treatment groups, control wells и пользовательских plate layouts.

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
- Approximate 95% CI;
- IQR/2;
- MAD.

### T-tests

PlateMaster использует Welch-style two-sample t-test для сравнения групп. P-values являются приблизительными и подходят для exploratory analysis; для публикационных выводов рекомендуется отдельная статистическая проверка.

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
| Pairwise T-test p-values | матрица попарных p-values |
| Outlier log | удалённые значения с файлом, временем, строкой, столбцом и значением |
| Outlier statistics | количество и процент удалённых значений |
| Plot data | данные, использованные для графика |

---

## 🗂 Структура проекта

```text
PlateMaster1.2.py   # Main Streamlit application and analysis logic
requirements.txt    # Python dependencies
README.md           # English and Russian documentation
LICENSE             # Full MIT License text
```

Приложение сейчас находится в одном Python-файле, что упрощает запуск, перенос и деплой.

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
- Приблизительные p-values не заменяют полноценный статистический pipeline.
- Очень большие Excel/ZIP файлы могут быть ограничены настройками Streamlit-сервера.
- Presets хранятся в session state и могут сбрасываться после перезапуска сессии.
- Первый столбец каждого входного файла всегда считается идентификатором строки.

---

## 📝 License

Этот проект распространяется под лицензией **MIT License**. Полный текст лицензии находится в файле [`LICENSE`](LICENSE).

---

## 🙌 Автор / цитирование

Repository: <https://github.com/narek-abelyan/PlateMaster>

Если PlateMaster помог в вашей работе, укажите ссылку на репозиторий, чтобы другие пользователи могли найти инструмент.

---
