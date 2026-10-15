from __future__ import annotations

from copy import deepcopy
from io import BytesIO, StringIO
from pathlib import Path
from typing import Callable
import importlib.util
import os
import re
import shutil
import math
import subprocess
import tempfile
import zipfile

import pandas as pd
import numpy as np
import streamlit as st
import altair as alt
from scipy import stats as scipy_stats

import analysis_settings as _analysis_settings

# A running Cloud process can retain the pre-pairwise module after a Git update.
# Reload that older module before importing the newly added settings constant.
if not hasattr(_analysis_settings, "MAX_PAIRWISE_GROUPS"):
    importlib.reload(_analysis_settings)

from analysis_settings import (
    APP_VERSION, DEFAULT_SETTINGS, ENUMS, LAYOUTS, MAX_PAIRWISE_GROUPS, PLOT_TYPES, dumps_config,
    input_record, loads_config, make_config, metric_names, provenance_sheets,
    replay_warnings, software_record, validate_config,
)

try:
    from streamlit.runtime.scriptrunner import get_script_run_ctx
except Exception:  # pragma: no cover
    get_script_run_ctx = None

EXCEL_EXTENSIONS = (".xlsx", ".xls", ".xlsm")
OLE_SIGNATURE = bytes.fromhex("D0CF11E0A1B11AE1")





def _available_excel_engines_for_suffix(suffix: str) -> list[str]:
    """Return reader candidates ordered for a file suffix, preferring installed Calamine."""
    suffix = suffix.lower()
    engines = ["openpyxl", "xlrd"] if suffix in {".xlsx", ".xlsm", "xlsx", "xlsm"} else ["xlrd", "openpyxl"]
    if importlib.util.find_spec("python_calamine") is not None:
        # Calamine can read both .xls and .xlsx and is often more tolerant.
        engines.insert(0, "calamine")
    return engines


def _try_libreoffice_convert_bytes_to_xlsx(content: bytes, filename: str = "input.xls") -> bytes | None:
    """Convert spreadsheet bytes to XLSX in a temporary directory using LibreOffice.

    Use the filename suffix for the temporary input. Return converted bytes,
    or None when LibreOffice is unavailable or does not produce an XLSX file."""
    office_bin = shutil.which("libreoffice") or shutil.which("soffice")
    if office_bin is None:
        return None

    suffix = Path(filename).suffix or ".xls"
    with tempfile.TemporaryDirectory() as tmpdir:
        input_path = Path(tmpdir) / f"input{suffix}"
        input_path.write_bytes(content)

        cmd = [office_bin, "--headless", "--convert-to", "xlsx", "--outdir", tmpdir, str(input_path)]
        run = subprocess.run(cmd, capture_output=True, text=True)
        if run.returncode != 0:
            return None

        output_path = input_path.with_suffix(".xlsx")
        if not output_path.exists():
            matches = list(Path(tmpdir).glob("*.xlsx"))
            if not matches:
                return None
            output_path = matches[0]

        return output_path.read_bytes()

def discover_excel_files(folder: str) -> list[Path]:
    """Return sorted Excel paths directly inside folder, or [] for an invalid folder."""
    base = Path(folder).expanduser().resolve()
    if not base.exists() or not base.is_dir():
        return []
    patterns = ("*.xlsx", "*.xls", "*.xlsm")
    files: list[Path] = []
    for pattern in patterns:
        files.extend(base.glob(pattern))
    return sorted(files)




def _read_spreadsheet_flexible_from_bytes(content: bytes, filename: str = "", sheet_name: str | int = 0, header: int = 0) -> pd.DataFrame:
    """Read spreadsheet bytes using Excel engines, text fallbacks, then LibreOffice.

    The filename selects reader order; sheet_name and header select the worksheet
    and zero-based header row. Return the parsed DataFrame, or raise ValueError
    with reader diagnostics if all attempts fail."""
    errors: list[str] = []
    suffix = filename.lower().rsplit('.', 1)[-1] if '.' in filename else ""

    engine_attempts = _available_excel_engines_for_suffix(suffix)
    for engine in engine_attempts:
        try:
            return pd.read_excel(BytesIO(content), engine=engine, sheet_name=sheet_name, header=header)
        except Exception as exc:
            errors.append(f"{engine}: {exc}")

    # Fallback for instrument exports mislabeled as .xls but containing delimited text.
    if not content.startswith((OLE_SIGNATURE, b"PK")) and sheet_name == 0:
        text = content.decode("utf-8", errors="ignore")
        for sep in ["\t", "   ", ";", ","]:
            try:
                frame = pd.read_csv(StringIO(text), sep=sep, header=header)
                if frame.shape[1] >= 2:
                    return frame
            except Exception as exc:
                errors.append(f"csv sep={sep!r}: {exc}")

    converted_xlsx = _try_libreoffice_convert_bytes_to_xlsx(content, filename or "input.xls")
    if converted_xlsx is not None:
        try:
            return pd.read_excel(BytesIO(converted_xlsx), engine="openpyxl", sheet_name=sheet_name, header=header)
        except Exception as exc:
            errors.append(f"libreoffice-convert: {exc}")

    raise ValueError(
        "Unable to parse spreadsheet content. "
        "If this is old .xls, ensure python-calamine is installed and optionally LibreOffice for auto-convert. Attempts: "
        + " | ".join(errors)
    )


def _read_spreadsheet_flexible_from_path(path: Path, sheet_name: str | int = 0, header: int = 0) -> pd.DataFrame:
    """Read a path with the same explicit worksheet/header rules as uploaded bytes."""
    return _read_spreadsheet_flexible_from_bytes(path.read_bytes(), path.name, sheet_name, header)


def sanitize_table(frame: pd.DataFrame, allow_custom: bool = False,
                   plate_shape: tuple[int, int] | None = None) -> pd.DataFrame:
    """Validate a plate matrix before dropping missing measurements.

    Reject missing/duplicate row IDs, non-finite values and unexpected dimensions.
    Non-plate labels require explicit opt-in and produce a retained warning.
    Empty measurement rows/columns and non-numeric cells are reported in attrs.
    """
    if frame.empty or frame.shape[1] < 2:
        raise ValueError("No measurement table found; check the worksheet and header row")
    frame = frame.copy()
    frame.columns = ["row_id", *[str(c).strip() for c in frame.columns[1:]]]
    if frame.columns.duplicated().any():
        raise ValueError("Duplicate column identifiers; check the header row")
    # Completely empty rows are harmless; a numeric row without an ID is not.
    frame = frame.dropna(how="all")
    labels = frame["row_id"].astype("string").str.strip()
    if labels.isna().any() or labels.eq("").any():
        raise ValueError("Missing row identifiers; check the worksheet and header row")
    if labels.duplicated().any():
        raise ValueError("Duplicate row identifiers (possibly instrument metadata): " +
                         ", ".join(labels[labels.duplicated()].unique()[:5]))
    frame["row_id"] = labels.astype(str)
    warnings = []
    # Ignore purely textual auxiliary columns (e.g. an OD unit column), but
    # retain numbered empty wells when checking the expected plate dimensions.
    auxiliary = [c for c in frame.columns[1:] if not re.fullmatch(r"[1-9][0-9]*", c)
                 and not pd.to_numeric(frame[c], errors="coerce").notna().any()]
    if auxiliary:
        frame = frame.drop(columns=auxiliary)
        warnings.append("Ignored non-measurement columns without numbers: " + ", ".join(auxiliary))
    def row_number(label):
        if not re.fullmatch(r"[A-Z]{1,2}", label):
            return 0
        value = 0
        for char in label:
            value = value * 26 + ord(char) - ord("A") + 1
        return value
    row_numbers = [row_number(x) for x in frame["row_id"]]
    columns = list(frame.columns[1:])
    column_numbers = [int(c) if re.fullmatch(r"[1-9][0-9]*", c) else 0 for c in columns]
    max_rows, max_cols = plate_shape or (32, 48)
    standard = (all(0 < r <= max_rows for r in row_numbers) and
                all(0 < c <= max_cols for c in column_numbers))
    if plate_shape and (frame.shape != (max_rows, max_cols + 1) or not standard):
        raise ValueError(f"Expected a {max_rows} x {max_cols} plate with row letters and numbered columns; "
                         f"found {len(frame)} rows and {len(columns)} measurement columns")
    if not standard:
        message = "Non-standard row/column labels or plate dimensions; numeric instrument metadata may be mistaken for measurements."
        if not allow_custom:
            raise ValueError(message + " Select the correct worksheet/header, or explicitly enable custom table labels after checking the file.")
        warnings.append(message)
    numeric = frame.iloc[:, 1:].apply(pd.to_numeric, errors="coerce")
    if np.isinf(numeric.to_numpy(dtype=float)).any():
        raise ValueError("Infinite measurement values are not supported")
    invalid = int((frame.iloc[:, 1:].notna() & numeric.isna()).to_numpy().sum())
    if invalid:
        warnings.append(f"{invalid} non-numeric measurement cell(s) treated as missing.")
    kept_columns = numeric.columns[numeric.notna().any()].tolist()
    if not kept_columns:
        raise ValueError("No numeric measurements; check the worksheet and header row")
    keep_rows = numeric[kept_columns].notna().any(axis=1)
    result = pd.concat([frame[["row_id"]], numeric[kept_columns]], axis=1).loc[keep_rows].copy()
    dropped_rows, dropped_cols = len(frame) - len(result), len(columns) - len(kept_columns)
    if dropped_rows or dropped_cols:
        warnings.append(f"Removed {dropped_rows} row(s) and {dropped_cols} column(s) without numeric measurements.")
    result.attrs["import_warnings"] = warnings
    return result


def load_table_from_path(path: Path, **options) -> pd.DataFrame:
    """Read a path using the explicit import options supported by the byte reader."""
    return load_table_from_bytes(path.read_bytes(), path.name, **options)


def load_table_from_bytes(content: bytes, filename: str = "", sheet_name: str | int = 0,
                          header: int = 0, allow_custom: bool = False,
                          plate_shape: tuple[int, int] | None = None) -> pd.DataFrame:
    """Read and validate a spreadsheet; retain import choices for provenance."""
    raw = _read_spreadsheet_flexible_from_bytes(content, filename, sheet_name, header)
    result = sanitize_table(raw, allow_custom, plate_shape)
    result.attrs["import_options"] = {"sheet": sheet_name, "header_row": header + 1,
                                     "allow_custom": allow_custom, "plate_shape": plate_shape}
    return result



def _column_sort_key(label: str) -> tuple[int, float | str]:
    """Return a key that orders numeric labels by value before other text labels."""
    label_str = str(label).strip()
    try:
        return (0, float(label_str.replace(",", ".")))
    except ValueError:
        return (1, label_str)


def _sort_columns_natural(labels: list[str]) -> list[str]:
    """Return column labels sorted numerically where possible, then lexically."""
    return sorted(labels, key=_column_sort_key)


def _extract_time_from_filename(filename: str) -> float | None:
    """Return the last unsigned numeric token in the filename stem, or None.

    Accept a period or comma as decimal separator. No underscore is required;
    sample or replicate identifiers can consequently be interpreted as time."""
    stem = Path(filename).stem
    matches = re.findall(r"(\d+(?:[\.,]\d+)?)", stem)
    if not matches:
        return None
    try:
        return float(matches[-1].replace(",", "."))
    except ValueError:
        return None




def _format_time_label(value: float | int | str | None) -> str:
    """Format a time to at most six decimals; map None to empty text and retain other text."""
    if value is None:
        return ""
    try:
        num = float(value)
        text = f"{num:.6f}".rstrip("0").rstrip(".")
        return text
    except Exception:
        return str(value)

def _remove_outliers(series: pd.Series, method: str, strictness: float) -> pd.Series:
    """Return non-missing observations retained by the selected outlier rule.

    Preserve original indices for logging. Strictness sets the IQR multiplier,
    Z-score threshold, or modified MAD-score threshold. None disables filtering;
    zero dispersion or an unrecognized method leaves available values unchanged."""
    s = series.dropna().astype(float)
    if s.empty or method == "None":
        return s

    if method == "IQR":
        q1 = s.quantile(0.25)
        q3 = s.quantile(0.75)
        iqr = q3 - q1
        if iqr == 0:
            return s
        lower = q1 - strictness * iqr
        upper = q3 + strictness * iqr
        return s[(s >= lower) & (s <= upper)]

    if method == "Z-score":
        std = s.std(ddof=0)
        if std == 0:
            return s
        z = (s - s.mean()).abs() / std
        return s[z <= strictness]

    if method == "MAD":
        median = s.median()
        mad = (s - median).abs().median()
        if mad == 0:
            return s
        modified_z = 0.6745 * (s - median).abs() / mad
        return s[modified_z <= strictness]

    return s

def _finite_observations(values) -> pd.Series:
    """Use available numeric observations; reject infinities rather than reporting invalid tests."""
    series = pd.to_numeric(pd.Series(values, dtype="object"), errors="coerce").dropna().astype(float)
    if not np.isfinite(series).all():
        raise ValueError("Statistical inputs must not contain infinite values")
    return series


def _dispersion_metrics(values) -> dict[str, float]:
    """Sample SD, SEM and Student-t 95% CI half-width for the arithmetic mean."""
    series = _finite_observations(values)
    n = len(series)
    sd = float(series.std(ddof=1)) if n > 1 else float("nan")
    sem = sd / math.sqrt(n) if n > 1 else float("nan")
    ci95 = float(scipy_stats.t.ppf(0.975, n - 1) * sem) if n > 1 else float("nan")
    return {"std": sd, "sem": sem, "ci95": ci95,
            "iqr_half": float((series.quantile(0.75) - series.quantile(0.25)) / 2.0),
            "mad": float((series - series.median()).abs().median())}


def aggregate_rows(
    table: pd.DataFrame,
    selected_rows: list[str],
    selected_columns: list[str],
    func: Callable[[pd.Series], float],
    outlier_method: str = "None",
    outlier_strictness: float = 1.5,
) -> tuple[pd.DataFrame, list[dict[str, float | str]], dict[str, dict[str, float]], dict[str, list[float]]]:
    """Aggregate selected rows into one value per selected measurement column.

    Apply func to observations remaining after missing-value removal and
    optional outlier filtering. Return the one-row DataFrame, removed-value
    records, per-column dispersion metrics, and retained observations.
    Raise ValueError if no selected rows or usable values remain."""
    filtered = table[table["row_id"].isin(selected_rows)]
    if filtered.empty:
        raise ValueError("Selected rows not found in table")

    removed_rows: list[dict[str, float | str]] = []
    values: dict[str, float] = {}
    column_stats: dict[str, dict[str, float]] = {}
    source_values: dict[str, list[float]] = {}
    for column in selected_columns:
        if column not in filtered.columns:
            continue
        original_series = filtered[column].dropna()
        cleaned_series = _remove_outliers(original_series, outlier_method, outlier_strictness)

        removed_indices = original_series.index.difference(cleaned_series.index)
        for idx in removed_indices:
            removed_rows.append({
                "row_id": str(filtered.loc[idx, "row_id"]),
                "column": str(column),
                "value": float(original_series.loc[idx]),
            })

        if cleaned_series.empty:
            continue
        values[column] = func(cleaned_series)
        source_values[column] = [float(x) for x in cleaned_series.tolist()]

        column_stats[column] = _dispersion_metrics(cleaned_series)

    if not values:
        raise ValueError("No values for selected columns")

    return pd.DataFrame([values], index=["aggregated"]), removed_rows, column_stats, source_values






def aggregate_by_groups(
    table: pd.DataFrame,
    group_map: dict[str, list[tuple[str, str]]],
    func: Callable[[pd.Series], float],
    outlier_method: str = "None",
    outlier_strictness: float = 1.5,
) -> tuple[pd.DataFrame, list[dict[str, float | str]], dict[str, dict[str, float]], dict[str, list[float]]]:
    """Aggregate plate cells listed as (row_id, column) pairs in each named group.

    Apply func after removing missing values and optional outliers. Return
    the one-row group DataFrame, removed-cell records, dispersion metrics,
    and retained observations. Skip unavailable cells and empty groups;
    raise ValueError when no group has usable values."""
    removed_rows: list[dict[str, float | str]] = []
    values: dict[str, float] = {}
    group_stats: dict[str, dict[str, float]] = {}
    source_values: dict[str, list[float]] = {}

    table_indexed = table.set_index("row_id", drop=False)

    for group_name, members in group_map.items():
        collected: list[tuple[str, str, float]] = []
        for row_id, col in members:
            if row_id in table_indexed.index and col in table_indexed.columns:
                v = pd.to_numeric(table_indexed.at[row_id, col], errors="coerce")
                if pd.notna(v):
                    collected.append((row_id, col, float(v)))

        if not collected:
            continue

        original_series = pd.Series([x[2] for x in collected])
        cleaned_series = _remove_outliers(original_series, outlier_method, outlier_strictness)

        removed_indices = original_series.index.difference(cleaned_series.index)
        for idx in removed_indices:
            row_id, col, val = collected[int(idx)]
            removed_rows.append({"row_id": row_id, "column": str(col), "group": group_name, "value": float(val)})

        if cleaned_series.empty:
            continue

        values[group_name] = func(cleaned_series)
        source_values[group_name] = [float(x) for x in cleaned_series.tolist()]
        group_stats[group_name] = _dispersion_metrics(cleaned_series)

    if not values:
        raise ValueError("No values for selected groups")

    return pd.DataFrame([values], index=["aggregated"]), removed_rows, group_stats, source_values


def aggregate_columns_by_row(
    table: pd.DataFrame,
    selected_rows: list[str],
    selected_columns: list[str],
    func: Callable[[pd.Series], float],
    outlier_method: str = "None",
    outlier_strictness: float = 1.5,
) -> tuple[pd.DataFrame, list[dict[str, float | str]], dict[str, dict[str, float]], dict[str, list[float]]]:
    """Aggregate selected measurement columns into one value per selected row.

    Apply func after missing-value removal and optional outlier filtering.
    Return the one-row DataFrame, removed-cell records, per-row dispersion
    metrics, and retained observations. Raise ValueError if no usable data remain."""
    filtered = table[table["row_id"].isin(selected_rows)]
    if filtered.empty:
        raise ValueError("Selected rows not found in table")

    removed_rows: list[dict[str, float | str]] = []
    values: dict[str, float] = {}
    row_stats: dict[str, dict[str, float]] = {}
    source_values: dict[str, list[float]] = {}

    for _, row in filtered.iterrows():
        row_id = str(row["row_id"])
        row_series = pd.to_numeric(row[selected_columns], errors="coerce").dropna()
        if row_series.empty:
            continue

        cleaned_series = _remove_outliers(row_series, outlier_method, outlier_strictness)
        removed_indices = row_series.index.difference(cleaned_series.index)
        for col in removed_indices:
            removed_rows.append({"row_id": row_id, "column": str(col), "value": float(row_series.loc[col])})

        if cleaned_series.empty:
            continue

        values[row_id] = func(cleaned_series)
        source_values[row_id] = [float(x) for x in cleaned_series.tolist()]
        row_stats[row_id] = _dispersion_metrics(cleaned_series)

    if not values:
        raise ValueError("No values for selected rows")

    return pd.DataFrame([values], index=["aggregated"]), removed_rows, row_stats, source_values


def _basic_stat_metrics(values: list[float]) -> pd.DataFrame:
    """Describe available observations and test their mean against zero using Student's t."""
    series = _finite_observations(values)
    if series.empty:
        return pd.DataFrame([{"Metric": "No data", "Value": "N/A"}])
    n = len(series)
    mean = float(series.mean())
    spread = _dispersion_metrics(series)
    mode_vals = series.mode()
    p_value = float("nan")
    status = "Need at least 2 observations" if n < 2 else "Undefined (zero sample variance)"
    if n > 1 and spread["std"] > 0:
        p_value = float(scipy_stats.ttest_1samp(series.to_numpy(), popmean=0.0).pvalue)
        status = "Two-sided Student t-test; null mean = 0"
    rows = [
        {"Metric": "n", "Value": n},
        {"Metric": "mean", "Value": mean},
        {"Metric": "median", "Value": float(series.median())},
        {"Metric": "mode", "Value": float(mode_vals.iloc[0]) if not mode_vals.empty else float("nan")},
        {"Metric": "sd", "Value": spread["std"]},
        {"Metric": "sem", "Value": spread["sem"]},
        {"Metric": "mad", "Value": spread["mad"]},
        {"Metric": "min", "Value": float(series.min())},
        {"Metric": "max", "Value": float(series.max())},
        {"Metric": "95% CI half-width (mean, Student t)", "Value": spread["ci95"]},
        {"Metric": "95% CI lower (mean)", "Value": mean - spread["ci95"]},
        {"Metric": "95% CI upper (mean)", "Value": mean + spread["ci95"]},
        {"Metric": "IQR/2", "Value": spread["iqr_half"]},
        {"Metric": "p-value (mean != 0, Student t)", "Value": p_value},
        {"Metric": "One-sample test status", "Value": status},
    ]
    return pd.DataFrame(rows)


def _welch_result(values_a, values_b) -> dict:
    """One validated calculation shared by the two-group table and pairwise matrix."""
    a, b = _finite_observations(values_a), _finite_observations(values_b)
    result = {"n_a": len(a), "n_b": len(b), "mean_a": float(a.mean()), "mean_b": float(b.mean()),
              "t": float("nan"), "df": float("nan"), "p": float("nan"), "status": "OK"}
    if len(a) < 2 or len(b) < 2:
        result["status"] = "Need at least 2 observations in each group"
    elif a.var(ddof=1) == 0 and b.var(ddof=1) == 0:
        result["status"] = "Undefined (both groups have zero sample variance)"
    else:
        test = scipy_stats.ttest_ind(a.to_numpy(), b.to_numpy(), equal_var=False)
        result.update(t=float(test.statistic), df=float(test.df), p=float(test.pvalue))
    return result


def _welch_t_test(values_a: list[float], values_b: list[float]) -> pd.DataFrame:
    """Return Welch's unequal-variance t-test, including its degrees of freedom."""
    r = _welch_result(values_a, values_b)
    return pd.DataFrame([
        {"Metric": "Group A n", "Value": r["n_a"]}, {"Metric": "Group B n", "Value": r["n_b"]},
        {"Metric": "Group A mean", "Value": r["mean_a"]}, {"Metric": "Group B mean", "Value": r["mean_b"]},
        {"Metric": "Welch t", "Value": r["t"]}, {"Metric": "df (Welch)", "Value": r["df"]},
        {"Metric": "p-value (2-sided, Welch t)", "Value": r["p"]},
        {"Metric": "Test status", "Value": r["status"]},
    ])


def _welch_p_value(values_a: list[float], values_b: list[float]) -> float:
    """Return exactly the same Welch p-value as the two-group comparison table."""
    return _welch_result(values_a, values_b)["p"]


def _pairwise_p_values(group_values: dict[str, list[float]]) -> pd.DataFrame:
    """Return formatted unadjusted Welch p-values once per pair, in the lower triangle.

    Require 2–20 explicitly selected groups. Leave the diagonal and upper
    triangle blank. Show undefined comparisons as N/A, values below 0.005 as
    ``p < 0.005``, and all other p-values rounded to three decimal places.
    """
    labels = list(group_values)
    if not 2 <= len(labels) <= MAX_PAIRWISE_GROUPS:
        raise ValueError(f"Select between 2 and {MAX_PAIRWISE_GROUPS} groups.")
    matrix = pd.DataFrame("", index=labels, columns=labels, dtype=object)
    matrix.index.name = "Group (Time | metric)"
    for i, first in enumerate(labels):
        for second in labels[i + 1:]:
            p_value = _welch_p_value(group_values[first], group_values[second])
            if pd.isna(p_value):
                displayed = "N/A"
            elif p_value < 0.005:
                displayed = "p < 0.005"
            else:
                displayed = f"{p_value:.3f}"
            matrix.loc[second, first] = displayed
    return matrix


def _build_excel_report(sheets: dict[str, pd.DataFrame]) -> bytes:
    """Serialize named DataFrames as an XLSX workbook in memory.

    Replace invalid sheet-name characters, truncate names to 31 characters,
    and write each table without its DataFrame index. Return workbook bytes."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            safe_sheet = re.sub(r"[\[\]\*\:/\\\?]", "_", sheet_name)[:31]
            df.to_excel(writer, index=False, sheet_name=safe_sheet or "Sheet1")
    output.seek(0)
    return output.getvalue()


def _to_numeric_series(values: pd.Series | list[float] | np.ndarray) -> pd.Series:
    """Coerce a one-dimensional collection to numbers and omit missing or invalid entries."""
    return pd.to_numeric(pd.Series(values, dtype="object"), errors="coerce").dropna()


AGGREGATIONS: dict[str, Callable[[pd.Series | list[float] | np.ndarray], float]] = {
    "Mean": lambda s: float(_to_numeric_series(s).mean()),
    "Median": lambda s: float(_to_numeric_series(s).median()),
    "Min": lambda s: float(_to_numeric_series(s).min()),
    "Max": lambda s: float(_to_numeric_series(s).max()),
    "Std": lambda s: float(_to_numeric_series(s).std(ddof=0)),
    "Sum": lambda s: float(_to_numeric_series(s).sum()),
}




def _gradient_style(df: pd.DataFrame, mode: str = "Global scale", palette: str = "Rose"):
    """Return a pandas Styler with custom RGB cell backgrounds for numeric data.

    Exclude Time and normalize colors globally, by column, or by row using
    the requested palette. Return the original DataFrame if empty, without
    numeric columns, or if styling fails."""
    if df.empty:
        return df
    numeric_cols = [c for c in df.select_dtypes(include=["number"]).columns.tolist() if str(c).strip().lower() != "time"]
    if not numeric_cols:
        return df

    try:
        values = df[numeric_cols].apply(pd.to_numeric, errors="coerce")
        style_map = pd.DataFrame("", index=df.index, columns=df.columns)

        if mode == "Per-column scale":
            mins = values.min(axis=0)
            spans = (values.max(axis=0) - mins).replace(0, pd.NA)
            ratios = values.sub(mins, axis=1).div(spans, axis=1)
        elif mode == "Per row scale":
            mins = values.min(axis=1)
            spans = (values.max(axis=1) - mins).replace(0, pd.NA)
            ratios = values.sub(mins, axis=0).div(spans, axis=0)
        else:
            global_min = values.min().min()
            global_max = values.max().max()
            span = (global_max - global_min) if pd.notna(global_max) and pd.notna(global_min) else pd.NA
            if pd.isna(span) or span == 0:
                ratios = values * 0 + 0.5
            else:
                ratios = (values - global_min) / span

        palette_rgb: dict[str, tuple[tuple[float, float, float], tuple[float, float, float]]] = {
            "Rose": ((255.0, 246.0, 250.0), (255.0, 92.0, 175.0)),
            "Pink-Violet": ((255.0, 244.0, 252.0), (220.0, 120.0, 230.0)),
            "Purple": ((248.0, 244.0, 255.0), (171.0, 71.0, 188.0)),
            "Blue": ((240.0, 248.0, 255.0), (66.0, 135.0, 245.0)),
            "Green": ((243.0, 252.0, 245.0), (52.0, 168.0, 83.0)),
            "Orange": ((255.0, 248.0, 240.0), (245.0, 124.0, 0.0)),
            "Teal": ((240.0, 252.0, 252.0), (0.0, 137.0, 123.0)),
            "Gray": ((250.0, 250.0, 250.0), (97.0, 97.0, 97.0)),
        }
        start_rgb, end_rgb = palette_rgb.get(palette, palette_rgb["Rose"])

        for row_idx in ratios.index:
            for col in ratios.columns:
                ratio = ratios.loc[row_idx, col]
                if pd.isna(ratio):
                    continue
                ratio = max(0.0, min(1.0, float(ratio)))
                r = int(start_rgb[0] + (end_rgb[0] - start_rgb[0]) * ratio)
                g = int(start_rgb[1] + (end_rgb[1] - start_rgb[1]) * ratio)
                b = int(start_rgb[2] + (end_rgb[2] - start_rgb[2]) * ratio)
                style_map.loc[row_idx, col] = f"background-color: rgb({r}, {g}, {b}); color: #111;"

        return df.style.apply(lambda _: style_map, axis=None)
    except Exception:
        return df


def _input_manifest() -> list[dict]:
    """Return sorted loaded-file metadata, using parsed-table records when bytes are unavailable."""
    records = st.session_state.get("input_records", {})
    return [records.get(name) or input_record(name, None, table, _extract_time_from_filename(name))
            for name, table in sorted(st.session_state.loaded_tables.items())]


def _capture_analysis_config() -> dict:
    """Capture effective choices, including dynamically created statistical selectors."""
    settings = deepcopy(DEFAULT_SETTINGS)
    settings.update(st.session_state.get("_last_settings", {}))
    threshold_keys = {"IQR": "iqr_multiplier", "Z-score": "zscore_threshold", "MAD": "mad_threshold"}
    active_threshold = threshold_keys.get(st.session_state.get("outlier_method", settings["outlier_method"]))
    for key in settings:
        # Streamlit may retain a stale widget default after an inactive slider is removed.
        # Its permanent saved value, not hidden widget state, is authoritative.
        if key in threshold_keys.values() and key != active_threshold:
            continue
        if key in st.session_state:
            settings[key] = deepcopy(st.session_state[key])
    groups = [deepcopy(cell) for cell in st.session_state.get("_plate_groups", [])
              if cell["row_id"] in settings["selected_rows"] and cell["column"] in settings["selected_columns"]]
    metrics = metric_names(settings, groups)
    if "plot_columns" not in st.session_state and "plot_columns" not in st.session_state.get("_last_settings", {}):
        settings["plot_columns"] = list(metrics)
    settings["plot_columns"] = [c for c in settings["plot_columns"] if c in metrics]
    for key in ("stat_metric", "ttest_metric_a", "ttest_metric_b"):
        if settings[key] not in metrics:
            settings[key] = None
    return make_config(settings, groups, _input_manifest(), st.session_state["_software_record"])


def _save_named_preset() -> None:
    """Save current analysis settings under the session preset name, or show a missing-name error."""
    name = st.session_state.get("preset_name", "").strip()
    if not name:
        st.session_state["_settings_error"] = "Enter a preset name before saving."
        return
    st.session_state.analysis_presets[name] = _capture_analysis_config()
    st.session_state["_settings_message"] = f"Preset '{name}' saved. Download its settings to keep it after restart."


def _queue_named_preset() -> None:
    """Queue a copied session preset for application before widgets are recreated."""
    name = st.session_state.get("preset_to_load", "")
    if not name:
        st.session_state["_settings_error"] = "Choose a preset to load."
        return
    st.session_state["_pending_config"] = deepcopy(st.session_state.analysis_presets[name])


def _queue_settings_upload() -> None:
    """Validate uploaded JSON and queue settings, storing any validation error for display."""
    uploaded = st.session_state.get("settings_upload")
    if uploaded is None:
        st.session_state["_settings_error"] = "Choose a settings JSON file first."
        return
    try:
        st.session_state["_pending_config"] = loads_config(uploaded.getvalue())
    except (ValueError, TypeError) as exc:
        st.session_state["_settings_error"] = str(exc)


def _apply_pending_config(all_rows: list[str], all_columns: list[str]) -> None:
    """Apply a queued document atomically, before any settings widgets are instantiated."""
    pending = st.session_state.pop("_pending_config", None)
    if pending is None:
        return
    try:
        config = validate_config(pending)
        settings = config["settings"]
        missing_rows = set(settings["selected_rows"]) - set(all_rows)
        missing_columns = set(settings["selected_columns"]) - set(all_columns)
        if missing_rows or missing_columns:
            raise ValueError(f"Settings were not applied: input rows/columns are missing. Rows: {sorted(missing_rows)}; columns: {sorted(missing_columns)}")
    except (ValueError, TypeError) as exc:
        st.session_state["_settings_error"] = str(exc)
        return
    for key, value in settings.items():
        if value is None:
            st.session_state.pop(key, None)
        else:
            st.session_state[key] = deepcopy(value)
    st.session_state["_last_settings"] = deepcopy(settings)
    st.session_state["_plate_groups"] = deepcopy(config["plate_groups"])
    st.session_state.pop("_template_axes", None)
    st.session_state["_replay_reference"] = config
    st.session_state.analysis_ready = False
    st.session_state["_settings_message"] = "Settings applied. Review the inputs, then press Run."


def _prepare_selectbox(key: str, options: list) -> None:
    """Discard stale selections before a widget is created, never after it is instantiated."""
    if st.session_state.get(key) not in options:
        st.session_state.pop(key, None)


def _plate_template_csv(rows: list[str], columns: list[str], cells: list[dict]) -> bytes:
    """Serialize every selected well and its optional group as a reusable CSV template."""
    labels = {(str(cell["row_id"]), str(cell["column"])): str(cell["group"]) for cell in cells}
    records = [
        {"row_id": str(row), "column": str(column), "group": labels.get((str(row), str(column)), "")}
        for row in rows
        for column in columns
    ]
    return pd.DataFrame(records, columns=["row_id", "column", "group"]).to_csv(index=False).encode("utf-8")


def _load_plate_template_csv(content: bytes, rows: list[str], columns: list[str]) -> list[dict]:
    """Validate a plate-template CSV and return its non-empty group assignments."""
    try:
        frame = pd.read_csv(BytesIO(content), dtype=str, keep_default_na=False, encoding="utf-8-sig")
    except Exception as exc:
        raise ValueError(f"Unable to read plate template CSV: {exc}") from exc
    frame.columns = [str(column).strip() for column in frame.columns]
    required = ["row_id", "column", "group"]
    if list(frame.columns) != required:
        raise ValueError("Plate template must contain exactly these columns: row_id, column, group.")
    for column in required:
        frame[column] = frame[column].astype(str).str.strip()
    if frame[["row_id", "column"]].eq("").any(axis=None):
        raise ValueError("Plate template contains an empty row_id or column.")
    duplicates = frame.duplicated(["row_id", "column"], keep=False)
    if duplicates.any():
        wells = sorted({f"{row.row_id}{row.column}" for row in frame.loc[duplicates].itertuples()})
        raise ValueError("Plate template contains duplicate wells: " + ", ".join(wells))
    unknown_rows = sorted(set(frame["row_id"]) - set(map(str, rows)))
    unknown_columns = _sort_columns_natural(list(set(frame["column"]) - set(map(str, columns))))
    if unknown_rows or unknown_columns:
        raise ValueError(
            f"Plate template does not match the selected data. Unknown rows: {unknown_rows}; "
            f"unknown columns: {unknown_columns}."
        )
    assignments = frame.loc[frame["group"] != "", required].to_dict(orient="records")
    if not assignments:
        raise ValueError("Plate template does not contain any group assignments.")
    return assignments


def _apply_plate_template_upload(rows: list[str], columns: list[str]) -> None:
    """Apply an uploaded CSV plate map to the current template axes."""
    uploaded = st.session_state.get("plate_template_upload")
    if uploaded is None:
        st.session_state["_plate_template_error"] = "Choose a plate template CSV file first."
        return
    try:
        st.session_state["_plate_groups"] = _load_plate_template_csv(uploaded.getvalue(), rows, columns)
        st.session_state.pop("_template_axes", None)
        st.session_state.analysis_ready = False
        st.session_state["_plate_template_message"] = "Plate template applied. Review the groups, then press Run."
    except (ValueError, TypeError) as exc:
        st.session_state["_plate_template_error"] = str(exc)


def _edit_plate_groups(rows: list[str], columns: list[str]) -> dict[str, list[tuple[str, str]]]:
    """Keep the data editor's base stable across edits; replace it only on load/axis changes."""
    axes = (tuple(rows), tuple(columns))
    if st.session_state.get("_template_axes") != axes:
        template = pd.DataFrame("", index=rows, columns=columns)
        for cell in st.session_state.get("_plate_groups", []):
            if cell["row_id"] in rows and cell["column"] in columns:
                template.loc[cell["row_id"], cell["column"]] = cell["group"]
        st.session_state["_template_base"] = template
        st.session_state["_template_axes"] = axes
        st.session_state["_template_revision"] = st.session_state.get("_template_revision", 0) + 1
    edited = st.data_editor(st.session_state["_template_base"], width="stretch",
                            key=f"plate_group_template_{st.session_state['_template_revision']}")
    cells = []
    mapping = {}
    for row_id in edited.index:
        for column in edited.columns:
            value = edited.loc[row_id, column]
            label = "" if pd.isna(value) else str(value).strip()
            if label:
                cells.append({"row_id": str(row_id), "column": str(column), "group": label})
                mapping.setdefault(label, []).append((str(row_id), str(column)))
    st.session_state["_plate_groups"] = cells
    return mapping


def _render_settings_download(slot) -> None:
    """Capture current settings, retain them in session state, and render a JSON download."""
    config = _capture_analysis_config()
    st.session_state["_current_analysis_config"] = config
    st.session_state["_last_settings"] = deepcopy(config["settings"])
    with slot.container():
        st.download_button(
            "Download settings.json",
            data=dumps_config(config),
            file_name="plate_master_analysis.json",
            mime="application/json",
            width="stretch",
        )


def run_dashboard() -> None:
    """Render the Streamlit import, analysis, plotting, and report-download workflow."""
    st.set_page_config(page_title="PlateMaster - Dashboard for Plate Analysis", layout="wide")
    st.title("PlateMaster - Dashboard for Plate Analysis")
    st.caption(f"Version {APP_VERSION}")
    st.session_state["_software_record"] = software_record(Path(__file__).resolve().parent)
    st.caption(
        "Load Excel files from server path OR upload from your local PC, choose rows/columns, aggregate, and plot across files."
    )


    if "loaded_tables" not in st.session_state:
        st.session_state.loaded_tables = {}
    if "load_errors" not in st.session_state:
        st.session_state.load_errors = []
    if "analysis_ready" not in st.session_state:
        st.session_state.analysis_ready = False
    if "analysis_presets" not in st.session_state:
        st.session_state.analysis_presets = {}

    loaded_tables: dict[str, pd.DataFrame] = dict(st.session_state.loaded_tables)
    errors: list[str] = list(st.session_state.load_errors)

    with st.sidebar:
        st.header("1) Data source")
        source_mode = st.radio("Source", ["Server folder path", "Local PC upload"], index=1)

        with st.expander("Import options"):
            sheet_index = int(st.number_input("Worksheet number (1 = first)", min_value=1, value=1))
            sheet_label = st.text_input("Worksheet name (optional; overrides number)")
            header_row = int(st.number_input("Header row (1 = first)", min_value=1, value=1))
            plate_format = st.selectbox("Expected plate", ["Auto / partial table", "6 (2 x 3)", "12 (3 x 4)",
                                        "24 (4 x 6)", "48 (6 x 8)", "96 (8 x 12)",
                                        "384 (16 x 24)", "1536 (32 x 48)"])
            allow_custom = st.checkbox("Allow custom table labels (review possible metadata)", value=False)
            st.caption("These choices apply to every file in this load. After changing them, load the files again. "
                       "Auto accepts partial plate matrices; it does not verify a complete plate size.")
        shapes = {"6 (2 x 3)": (2, 3), "12 (3 x 4)": (3, 4), "24 (4 x 6)": (4, 6),
                  "48 (6 x 8)": (6, 8), "96 (8 x 12)": (8, 12),
                  "384 (16 x 24)": (16, 24), "1536 (32 x 48)": (32, 48)}
        import_options = dict(sheet_name=sheet_label.strip() or sheet_index - 1, header=header_row - 1,
                              allow_custom=allow_custom, plate_shape=shapes.get(plate_format))

        if source_mode == "Server folder path":
            folder_path = st.text_input("Server folder path with Excel files", value="")
            load_folder_clicked = st.button("Load folder")
            if load_folder_clicked:
                loaded_tables = {}
                records = {}
                errors = []
                for file in discover_excel_files(folder_path):
                    try:
                        content = file.read_bytes()
                        loaded_tables[file.name] = load_table_from_bytes(content, file.name, **import_options)
                        records[file.name] = input_record(file.name, content, loaded_tables[file.name], _extract_time_from_filename(file.name))
                    except Exception as exc:
                        errors.append(f"{file.name}: {exc}")
                st.session_state.input_records = records
                st.session_state.loaded_tables = loaded_tables
                st.session_state.load_errors = errors
                st.session_state.analysis_ready = False

        else:
            uploaded_files = st.file_uploader(
                "Upload Excel files from your local PC",
                type=["xlsx", "xls", "xlsm"],
                accept_multiple_files=True,
            )
            uploaded_zip = st.file_uploader(
                "Or upload ONE ZIP with many Excel files",
                type=["zip"],
                accept_multiple_files=False,
            )
            load_uploads_clicked = st.button("Load uploaded files")

            if load_uploads_clicked:
                loaded_tables = {}
                records = {}
                errors = []
                for uploaded in uploaded_files or []:
                    try:
                        if uploaded.name in loaded_tables:
                            raise ValueError("Duplicate input filename; rename the file before loading it")
                        content = uploaded.getvalue()
                        loaded_tables[uploaded.name] = load_table_from_bytes(content, uploaded.name, **import_options)
                        records[uploaded.name] = input_record(uploaded.name, content, loaded_tables[uploaded.name], _extract_time_from_filename(uploaded.name))
                    except Exception as exc:
                        errors.append(f"{uploaded.name}: {exc}")

                if uploaded_zip is not None:
                    try:
                        with zipfile.ZipFile(BytesIO(uploaded_zip.getvalue())) as archive:
                            for member in archive.namelist():
                                lower = member.lower()
                                if not lower.endswith(EXCEL_EXTENSIONS) or member.endswith("/"):
                                    continue
                                try:
                                    name = Path(member).name
                                    if name in loaded_tables:
                                        raise ValueError("Duplicate input filename; rename the file before loading it")
                                    content = archive.read(member)
                                    loaded_tables[name] = load_table_from_bytes(content, member, **import_options)
                                    records[name] = input_record(name, content, loaded_tables[name], _extract_time_from_filename(name))
                                except Exception as exc:
                                    errors.append(f"{member}: {exc}")
                    except Exception as exc:
                        errors.append(f"ZIP error: {exc}")

                st.session_state.input_records = records
                st.session_state.loaded_tables = loaded_tables
                st.session_state.load_errors = errors
                st.session_state.analysis_ready = False

    if errors:
        st.warning("Some files were skipped:\n- " + "\n- ".join(errors))

    if not loaded_tables:
        st.info("Choose data source and click a load button (**Load folder** or **Load uploaded files**).")
        return

    color_scale_mode = st.session_state.get("color_scale_mode", "Global scale")
    color_palette = st.session_state.get("color_palette", "Rose")

    st.success(f"Loaded {len(loaded_tables)} Excel file(s)")

    st.header("Input preview")
    for name, table in loaded_tables.items():
        for message in table.attrs.get("import_warnings", []):
            st.warning(f"{name}: {message}")
    _prepare_selectbox("preview_file", sorted(loaded_tables))
    preview_name = st.selectbox("Preview file", sorted(loaded_tables), key="preview_file")
    preview = loaded_tables[preview_name]
    st.caption(f"{len(preview)} rows x {len(preview.columns) - 1} measurement columns. "
               f"Import: {preview.attrs.get('import_options', {})}. Check that these are the intended measurements.")
    st.dataframe(_gradient_style(preview, color_scale_mode, color_palette), width="stretch")

    all_rows = sorted({row for table in loaded_tables.values() for row in table["row_id"].astype(str).tolist()})
    all_columns = _sort_columns_natural(sorted({str(c) for table in loaded_tables.values() for c in table.columns if c != "row_id"}))
    _apply_pending_config(all_rows, all_columns)
    for key, value in st.session_state.get("_last_settings", {}).items():
        if key not in st.session_state and value is not None:
            st.session_state[key] = deepcopy(value)
    active_method = st.session_state.get("outlier_method", "None")
    previous_settings = st.session_state.get("_last_settings", {})
    if active_method != previous_settings.get("outlier_method", "None"):
        threshold_key = {"IQR": "iqr_multiplier", "Z-score": "zscore_threshold", "MAD": "mad_threshold"}.get(active_method)
        if threshold_key:
            st.session_state[threshold_key] = previous_settings.get(threshold_key, DEFAULT_SETTINGS[threshold_key])
    # Switching data/layout can leave old widget values behind.
    for key, options in ENUMS.items():
        _prepare_selectbox(key, options)
    if st.session_state.get("agg_name", "Mean") != "Mean" and st.session_state.get("error_method") in ("SEM", "95% CI"):
        st.session_state.error_method = "None"
    for key, options in (("selected_rows", all_rows), ("selected_columns", all_columns)):
        if key in st.session_state:
            st.session_state[key] = [v for v in st.session_state[key] if v in options]
    if "_settings_error" in st.session_state:
        st.error(st.session_state.pop("_settings_error"))
    if "_settings_message" in st.session_state:
        st.success(st.session_state.pop("_settings_message"))
    if "_replay_reference" in st.session_state:
        differences = replay_warnings(st.session_state["_replay_reference"], _input_manifest(), st.session_state["_software_record"])
        if differences:
            st.warning("These settings are being reused with different or unverifiable inputs/environment. This is not an exact replay:\n- " + "\n- ".join(differences))
        else:
            st.info("Saved input checksums, parsed data and numerical software versions match.")


    with st.sidebar:
        st.header("2) Selection")
        layout_mode = st.selectbox(
            "Aggregation layout",
            options=LAYOUTS,
            index=0,
            key="agg_layout_mode",
        )
        selected_rows = st.multiselect("Rows", options=all_rows, default=all_rows if "selected_rows" not in st.session_state else None, key="selected_rows")
        selected_columns = st.multiselect("Columns", options=all_columns, default=all_columns if "selected_columns" not in st.session_state else None, key="selected_columns")
        selected_columns = _sort_columns_natural(selected_columns)
        agg_options = list(AGGREGATIONS.keys())
        agg_default_index = agg_options.index("Mean") if "Mean" in agg_options else 0
        agg_name = st.selectbox("Aggregation", options=agg_options, index=agg_default_index, key="agg_name")

        st.subheader("Outlier filtering")
        outlier_method = st.selectbox("Method", options=["None", "IQR", "Z-score", "MAD"], index=0, key="outlier_method")
        outlier_strictness = 1.5
        if outlier_method == "IQR":
            outlier_strictness = st.slider("IQR multiplier", min_value=0.5, max_value=5.0, value=1.5, step=0.1, key="iqr_multiplier")
        elif outlier_method == "Z-score":
            outlier_strictness = st.slider("Z-score threshold", min_value=1.0, max_value=5.0, value=3.0, step=0.1, key="zscore_threshold")
        elif outlier_method == "MAD":
            outlier_strictness = st.slider("MAD threshold", min_value=1.0, max_value=8.0, value=3.5, step=0.1, key="mad_threshold")

        color_controls_left, color_controls_right = st.columns(2)
        with color_controls_left:
            st.selectbox(
                "Color scale mode",
                options=["Global scale", "Per-column scale", "Per row scale"],
                key="color_scale_mode",
            )
        with color_controls_right:
            st.selectbox(
                "Color palette",
                options=["Rose", "Pink-Violet", "Purple", "Blue", "Green", "Orange", "Teal", "Gray"],
                key="color_palette",
            )
        color_scale_mode = st.session_state.color_scale_mode
        color_palette = st.session_state.color_palette

        st.subheader("Presets")
        preset_name = st.text_input("Preset name", value="", key="preset_name")
        preset_names = sorted(st.session_state.analysis_presets.keys())
        preset_to_load = st.selectbox(
            "Load preset",
            options=[""] + preset_names,
            format_func=lambda x: "Select preset..." if x == "" else x,
            key="preset_to_load",
        )
        preset_col1, preset_col3 = st.columns([3, 2])
        with preset_col1:
            st.button("Save preset now", on_click=_save_named_preset, width="stretch")
        with preset_col3:
            st.button("Apply preset", on_click=_queue_named_preset, width="stretch")
        st.file_uploader("Import analysis settings (.json)", type=["json"], key="settings_upload")
        settings_apply_col, settings_download_col = st.columns(2)
        with settings_apply_col:
            st.button(
                "Apply imported settings",
                on_click=_queue_settings_upload,
                width="stretch",
            )
        with settings_download_col:
            settings_download_slot = st.empty()
        st.caption("Settings include the plate groups and input checksums. Keep the original spreadsheets separately.")

        run_clicked = st.button("Run")
        if run_clicked:
            st.session_state.analysis_ready = True

    try:
        if not selected_rows or not selected_columns:
            st.warning("Please select at least one row and one column")
            return

        group_map: dict[str, list[tuple[str, str]]] = {}
        metric_labels = selected_columns
        if layout_mode == "By original rows":
            metric_labels = selected_rows
        elif layout_mode == "By plate groups (template)":
            st.header("Plate template grouping")
            st.info(
                "If you already have a plate template, upload its CSV file and click **Apply plate template**. "
                "If you do not have one, fill in the template grid below manually. You can then download the "
                "completed template and reuse it in future analyses."
            )
            template_upload_col, template_apply_col = st.columns([3, 1])
            with template_upload_col:
                st.file_uploader("Upload plate template (.csv)", type=["csv"], key="plate_template_upload")
            with template_apply_col:
                st.write("")
                st.button(
                    "Apply plate template",
                    key="apply_plate_template",
                    on_click=_apply_plate_template_upload,
                    args=(selected_rows, selected_columns),
                    width="stretch",
                )
            if "_plate_template_error" in st.session_state:
                st.error(st.session_state.pop("_plate_template_error"))
            if "_plate_template_message" in st.session_state:
                st.success(st.session_state.pop("_plate_template_message"))

            st.caption("Enter a group name in each required well. Wells with the same name are analysed as one group.")
            group_map = _edit_plate_groups(selected_rows, selected_columns)
            template_csv = _plate_template_csv(
                selected_rows, selected_columns, st.session_state.get("_plate_groups", [])
            )
            st.download_button(
                "Download plate template.csv",
                data=template_csv,
                file_name="plate_master_template.csv",
                mime="text/csv",
            )
            st.caption("The CSV contains every selected well. Wells without a group remain blank.")

            if not group_map:
                st.warning("Please define at least one group in plate template")
                return

            metric_labels = sorted(group_map.keys())

        if not st.session_state.analysis_ready:
            st.info("Press **Run** to start analysis with current settings.")
            return

        agg_func = AGGREGATIONS[agg_name]
        aggregated_tables: dict[str, pd.DataFrame] = {}
        removed_outliers_log: list[dict[str, float | str | None]] = []
        total_points_per_column: dict[str, int] = {c: 0 for c in metric_labels}
        total_points_per_time: dict[float, int] = {}
        error_lookup: dict[tuple[str, str], dict[str, float]] = {}
        source_value_lookup: dict[tuple[str, str], list[float]] = {}

        analysis_errors = []
        for name, table in loaded_tables.items():
            try:
                filtered_for_stats = table[table["row_id"].isin(selected_rows)]
                file_time = _extract_time_from_filename(name)
                file_total_points = 0

                if layout_mode == "By plate groups (template)":
                    table_indexed = table.set_index("row_id", drop=False)
                    for group_name, members in group_map.items():
                        count = 0
                        for row_id, col in members:
                            if row_id in table_indexed.index and col in table_indexed.columns:
                                v = pd.to_numeric(table_indexed.at[row_id, col], errors="coerce")
                                if pd.notna(v):
                                    count += 1
                        total_points_per_column[group_name] = total_points_per_column.get(group_name, 0) + count
                        file_total_points += count

                    avg_table, removed_rows, column_stats, source_values = aggregate_by_groups(
                        table, group_map, agg_func, outlier_method, outlier_strictness
                    )
                elif layout_mode == "By original rows":
                    row_filtered = table[table["row_id"].isin(selected_rows)]
                    for _, row in row_filtered.iterrows():
                        row_id = str(row["row_id"])
                        count = int(pd.to_numeric(row[selected_columns], errors="coerce").notna().sum())
                        total_points_per_column[row_id] = total_points_per_column.get(row_id, 0) + count
                        file_total_points += count

                    avg_table, removed_rows, column_stats, source_values = aggregate_columns_by_row(
                        table, selected_rows, selected_columns, agg_func, outlier_method, outlier_strictness
                    )
                else:
                    for col in selected_columns:
                        if col in filtered_for_stats.columns:
                            count = int(filtered_for_stats[col].notna().sum())
                            total_points_per_column[col] = total_points_per_column.get(col, 0) + count
                            file_total_points += count

                    avg_table, removed_rows, column_stats, source_values = aggregate_rows(
                        table, selected_rows, selected_columns, agg_func, outlier_method, outlier_strictness
                    )

                if file_time is not None:
                    total_points_per_time[file_time] = total_points_per_time.get(file_time, 0) + file_total_points

                aggregated_tables[name] = avg_table
                for col_name, stats in column_stats.items():
                    error_lookup[(name, str(col_name))] = stats
                for col_name, vals in source_values.items():
                    source_value_lookup[(name, str(col_name))] = vals
                for item in removed_rows:
                    removed_outliers_log.append({"Time": file_time, "file": name, **item})
            except Exception as exc:
                analysis_errors.append(f"{name}: {exc}")

        if analysis_errors:
            st.error("Analysis stopped; no partial results will be exported. Fix these files or selections:\n- " +
                     "\n- ".join(analysis_errors))
            return

        if not aggregated_tables:
            st.error("No aggregated tables could be produced for the current selection")
            return

        combined_rows: list[dict[str, float | str | None]] = []
        for name, avg_table in aggregated_tables.items():
            row = {k: float(v) for k, v in avg_table.iloc[0].to_dict().items()}
            row["Time"] = _extract_time_from_filename(name)
            row["File"] = name
            combined_rows.append(row)

        combined_df = pd.DataFrame(combined_rows)
        combined_df = combined_df.sort_values(by="Time", na_position="last").reset_index(drop=True)
        metric_columns = [c for c in _sort_columns_natural(metric_labels) if c in combined_df.columns]
        display_columns = ["Time", *metric_columns]

        st.header("Aggregated table (all files)")
        combined_display_df = combined_df[display_columns].copy()
        combined_display_df["Time"] = combined_display_df["Time"].map(
            lambda v: _format_time_label(v) if isinstance(v, (float, int, np.floating, np.integer)) else str(v)
        )
        st.dataframe(_gradient_style(combined_display_df, color_scale_mode, color_palette), width="stretch")

        st.header("Statistical analysis")
        stat_col1, stat_col2 = st.columns(2)

        unique_times = sorted(combined_df["Time"].dropna().unique().tolist())
        if not unique_times:
            st.warning("No Time values available for statistical analysis.")
            return

        for key in ("stat_time", "ttest_time_a", "ttest_time_b"):
            _prepare_selectbox(key, unique_times)
        for key in ("stat_metric", "ttest_metric_a", "ttest_metric_b"):
            _prepare_selectbox(key, metric_columns)
        st.caption("Tests use independent observations. Technical wells are not automatically independent biological replicates. P-values are unadjusted for multiple comparisons.")
        with stat_col1:
            selected_stat_time = st.selectbox(
                "Row (from Aggregated table)",
                options=unique_times,
                format_func=_format_time_label,
                key="stat_time",
            )

        with stat_col2:
            selected_stat_metric = st.selectbox(
                "Column (from Aggregated table)",
                options=[c for c in _sort_columns_natural(metric_labels) if c in combined_df.columns],
                key="stat_metric",
            )

        selected_files = combined_df.loc[combined_df["Time"] == selected_stat_time, "File"].astype(str).tolist()
        stat_values: list[float] = []
        for file_name in selected_files:
            stat_values.extend(source_value_lookup.get((file_name, str(selected_stat_metric)), []))
        stat_table = _basic_stat_metrics(stat_values)
        st.dataframe(stat_table.astype({"Value": "string"}), width="stretch")

        st.subheader("T-test analysis")
        ttest_col1, ttest_col2, ttest_col3, ttest_col4 = st.columns(4)
        with ttest_col1:
            ttest_time_a = st.selectbox("Group A", options=unique_times, format_func=_format_time_label, key="ttest_time_a")
        with ttest_col2:
            ttest_metric_a = st.selectbox("Group A Column", options=metric_columns, key="ttest_metric_a")

        with ttest_col3:
            ttest_time_b = st.selectbox(
                "Group B",
                options=unique_times,
                format_func=_format_time_label,
                key="ttest_time_b",
            )
        with ttest_col4:
            ttest_metric_b = st.selectbox("Group B Column", options=metric_columns, key="ttest_metric_b")

        files_a = combined_df.loc[combined_df["Time"] == ttest_time_a, "File"].astype(str).tolist()
        files_b = combined_df.loc[combined_df["Time"] == ttest_time_b, "File"].astype(str).tolist()
        values_a: list[float] = []
        values_b: list[float] = []
        for file_name in files_a:
            values_a.extend(source_value_lookup.get((file_name, str(ttest_metric_a)), []))
        for file_name in files_b:
            values_b.extend(source_value_lookup.get((file_name, str(ttest_metric_b)), []))

        ttest_df = _welch_t_test(values_a, values_b)
        st.dataframe(ttest_df.astype({"Value": "string"}), width="stretch")

        st.subheader("Optional pairwise comparisons")
        pairwise_matrix_df = pd.DataFrame()
        pairwise_enabled = st.checkbox("Enable pairwise comparisons", key="pairwise_enabled")
        if pairwise_enabled:
            st.warning("These p-values are unadjusted for multiple comparisons. "
                       "Interpreting many comparisons increases the risk of false-positive findings.")
            # Only build vectors for selected groups, not for every time/metric pair.
            group_options = {f"{float(tm)} | {metric}": (tm, metric)
                             for tm in unique_times for metric in metric_columns}
            saved_groups = st.session_state.get("pairwise_groups", [])
            available_groups = [label for label in saved_groups if label in group_options]
            if saved_groups != available_groups:
                st.warning("Unavailable pairwise groups were removed; review the selection.")
                st.session_state["pairwise_groups"] = available_groups
            chosen_groups = st.multiselect("Groups to compare (Time | metric)", options=list(group_options),
                                          key="pairwise_groups", max_selections=MAX_PAIRWISE_GROUPS)
            st.caption(f"Select 2–{MAX_PAIRWISE_GROUPS} groups. Each pair is calculated once; "
                       "the diagonal and upper triangle are blank. Values below 0.005 are shown as p < 0.005; other values are rounded to three decimal places. N/A means the test is undefined.")
            if 2 <= len(chosen_groups) <= MAX_PAIRWISE_GROUPS:
                group_values = {}
                for label in chosen_groups:
                    tm, metric = group_options[label]
                    files_tm = combined_df.loc[combined_df["Time"] == tm, "File"].astype(str)
                    group_values[label] = [value for name in files_tm
                                           for value in source_value_lookup.get((name, str(metric)), [])]
                pairwise_matrix_df = _pairwise_p_values(group_values)
                st.dataframe(pairwise_matrix_df, width="stretch")
            else:
                st.info(f"Select between 2 and {MAX_PAIRWISE_GROUPS} groups to calculate pairwise comparisons.")
        else:
            st.caption("Pairwise comparisons are off. Enable them only when needed.")

        removed_df_export = pd.DataFrame(removed_outliers_log)
        if not removed_df_export.empty:
            removed_df_export = removed_df_export.sort_values(by=["Time", "column", "row_id"], na_position="last")

        stats_rows: list[dict[str, float | int | str]] = []
        if outlier_method != "None":
            removed_total = int(len(removed_df_export))
            total_points = int(sum(total_points_per_column.values()))
            overall_pct = (removed_total / total_points * 100) if total_points else 0.0
            stats_rows.append({"Scope": "Overall", "Key": "All", "Total points": total_points, "Removed": removed_total, "Removed %": round(overall_pct, 2)})

            for col in _sort_columns_natural(metric_labels):
                col_total = int(total_points_per_column.get(col, 0))
                col_removed = int((removed_df_export["column"] == col).sum()) if not removed_df_export.empty else 0
                col_pct = (col_removed / col_total * 100) if col_total else 0.0
                stats_rows.append({"Scope": "Column", "Key": str(col), "Total points": col_total, "Removed": col_removed, "Removed %": round(col_pct, 2)})

            for tm in sorted(total_points_per_time.keys()):
                tm_total = int(total_points_per_time[tm])
                tm_removed = int((removed_df_export["Time"] == tm).sum()) if not removed_df_export.empty else 0
                tm_pct = (tm_removed / tm_total * 100) if tm_total else 0.0
                stats_rows.append({"Scope": "Time", "Key": str(tm), "Total points": tm_total, "Removed": tm_removed, "Removed %": round(tm_pct, 2)})

        outlier_stats_df = pd.DataFrame(stats_rows)

        st.header("Outlier filtering log")
        left_col, right_col = st.columns(2)

        with left_col:
            st.subheader("Removed values by file/row/column")
            if outlier_method == "None":
                st.info("Outlier filtering is disabled.")
            elif not removed_outliers_log:
                st.info("No outliers were removed with current method/settings.")
            else:
                st.dataframe(removed_df_export, width="stretch")

        with right_col:
            st.subheader("Outlier statistics")
            if outlier_method == "None":
                st.info("Outlier filtering is disabled.")
            else:
                st.dataframe(outlier_stats_df, width="stretch")

        st.header("Cross-file plot")
        if "plot_columns" not in st.session_state:
            st.session_state.plot_columns = list(metric_labels)
        available_saved_plot_columns = [c for c in st.session_state.plot_columns if c in metric_labels]
        st.session_state.plot_columns = available_saved_plot_columns or list(metric_labels)

        plot_columns = st.multiselect(
            "Columns to display on one chart (different colors)",
            options=metric_columns,
            key="plot_columns",
        )

        controls_col1, controls_col2 = st.columns(2)
        with controls_col1:
            plot_type = st.selectbox(
                "Plot type",
                options=PLOT_TYPES,
                index=0,
                key="plot_type",
            )
        with controls_col2:
            error_method = st.selectbox(
                "Error bar method",
                options=["None", "Std", "SEM", "95% CI", "IQR/2", "MAD"] if agg_name == "Mean" else ["None", "Std", "IQR/2", "MAD"],
                index=0,
                key="error_method",
            )

        plot_columns = [c for c in plot_columns if c in metric_labels]
        if not plot_columns:
            st.warning("Please choose at least one column for chart")
            return

        plot_df = combined_df[["File", "Time", *[c for c in _sort_columns_natural(metric_labels) if c in combined_df.columns]]].copy()
        available_plot_columns = [c for c in plot_columns if c in plot_df.columns]
        if not available_plot_columns:
            st.warning("No data for selected chart columns")
            return

        chart_long = plot_df[["File", "Time", *available_plot_columns]].melt(
            id_vars=["File", "Time"], var_name="Column", value_name="Value"
        )
        chart_long = chart_long.dropna(subset=["Value", "Time"]).copy()
        sorted_cols = _sort_columns_natural(available_plot_columns)
        chart_long["Column"] = pd.Categorical(chart_long["Column"], categories=sorted_cols, ordered=True)
        chart_long["TimeLabel"] = chart_long["Time"].map(_format_time_label)
        time_sort_order = [_format_time_label(v) for v in sorted(chart_long["Time"].dropna().unique())]

        error_field_map = {
            "Std": "std",
            "SEM": "sem",
            "95% CI": "ci95",
            "IQR/2": "iqr_half",
            "MAD": "mad",
        }
        if error_method != "None":
            field = error_field_map[error_method]
            chart_long["Error"] = chart_long.apply(
                lambda r: float(error_lookup.get((str(r["File"]), str(r["Column"])), {}).get(field, float("nan"))), axis=1
            )
        else:
            chart_long["Error"] = 0.0

        chart_long["y_low"] = chart_long["Value"] - chart_long["Error"]
        chart_long["y_high"] = chart_long["Value"] + chart_long["Error"]

        y_candidates = chart_long["Value"]
        if error_method != "None" and plot_type in {"Line + points", "Line + points (по Columns)", "Scatter", "Bar plot (по Time)", "Bar plot (по Columns)"}:
            y_candidates = pd.concat([chart_long["Value"], chart_long["y_low"], chart_long["y_high"]], ignore_index=True)
        y_min = float(y_candidates.min()) if not y_candidates.empty else 0.0
        y_max = float(y_candidates.max()) if not y_candidates.empty else 1.0
        pad = (y_max - y_min) * 0.08 if y_max > y_min else 0.1
        y_scale = alt.Scale(domain=[y_min - pad, y_max + pad], nice=True)

        x_min = float(chart_long["Time"].min()) if not chart_long.empty else 0.0
        x_max = float(chart_long["Time"].max()) if not chart_long.empty else 1.0
        x_pad = (x_max - x_min) * 0.04 if x_max > x_min else 0.5
        x_scale = alt.Scale(domain=[x_min - x_pad, x_max + x_pad], nice=True)
        cap_half_width = max((x_max - x_min) * 0.006, 0.02)
        chart_long["x_left"] = chart_long["Time"] - cap_half_width
        chart_long["x_right"] = chart_long["Time"] + cap_half_width

        if plot_type == "Scatter":
            chart = (
                alt.Chart(chart_long)
                .mark_point(filled=True, size=130, opacity=0.95)
                .encode(
                    x=alt.X("Time:Q", title="Time", scale=x_scale),
                    y=alt.Y("Value:Q", title="Aggregated value", scale=y_scale),
                    color=alt.Color("Column:N", title="Column"),
                    tooltip=["Time", "File", "Column", "Value"],
                )
                .interactive()
                .properties(height=450)
            )
        elif plot_type == "Bar plot (по Time)":
            chart = (
                alt.Chart(chart_long)
                .mark_bar(opacity=0.88)
                .encode(
                    x=alt.X("TimeLabel:N", title="Time", sort=time_sort_order),
                    y=alt.Y("Value:Q", title="Aggregated value", scale=y_scale, stack=None),
                    xOffset=alt.XOffset("Column:N"),
                    color=alt.Color("Column:N", title="Column"),
                    tooltip=["Time", "File", "Column", "Value"],
                )
                .properties(height=450)
            )
        elif plot_type == "Bar plot (по Columns)":
            chart = (
                alt.Chart(chart_long)
                .mark_bar(opacity=0.88)
                .encode(
                    x=alt.X("Column:N", title="Column", sort=sorted_cols),
                    y=alt.Y("Value:Q", title="Aggregated value", scale=y_scale, stack=None),
                    xOffset=alt.XOffset("TimeLabel:N"),
                    color=alt.Color("TimeLabel:N", title="Time"),
                    tooltip=["Time", "File", "Column", "Value"],
                )
                .properties(height=450)
            )
        elif plot_type == "Box plot (по Time)":
            chart = (
                alt.Chart(chart_long)
                .mark_boxplot(extent="min-max", size=24)
                .encode(
                    x=alt.X("TimeLabel:N", title="Time", sort=time_sort_order),
                    y=alt.Y("Value:Q", title="Aggregated value", scale=y_scale),
                    color=alt.Color("TimeLabel:N", title="Time"),
                    tooltip=["Time", "Value"],
                )
                .properties(height=460)
            )
        elif plot_type == "Box plot (по Columns)":
            chart = (
                alt.Chart(chart_long)
                .mark_boxplot(extent="min-max", size=24)
                .encode(
                    x=alt.X("Column:N", title="Column", sort=sorted_cols),
                    y=alt.Y("Value:Q", title="Aggregated value", scale=y_scale),
                    color=alt.Color("Column:N", title="Column"),
                    tooltip=["Column", "Value"],
                )
                .properties(height=460)
            )
        elif plot_type == "Violin plot (по Time)":
            if chart_long.empty:
                st.warning("No data available for violin plot")
                return

            violin_base = (
                alt.Chart(chart_long)
                .transform_density(
                    "Value",
                    as_=["Value", "Density"],
                    groupby=["TimeLabel"],
                )
                .mark_area(orient="horizontal", opacity=0.62, color="#ff6ab6")
                .encode(
                    y=alt.Y("Value:Q", title="Aggregated value", scale=y_scale),
                    x=alt.X("Density:Q", stack="center", axis=None),
                )
            )
            chart = violin_base.facet(
                column=alt.Column("TimeLabel:N", title="Time", sort=time_sort_order, header=alt.Header(labelAngle=0)),
                columns=4,
            ).properties(height=180, width=80)
        elif plot_type == "Violin plot (по Columns)":
            chart = (
                alt.Chart(chart_long)
                .transform_density(
                    "Value",
                    as_=["Value", "Density"],
                    groupby=["Column"],
                )
                .mark_area(orient="horizontal", opacity=0.62)
                .encode(
                    y=alt.Y("Value:Q", title="Aggregated value", scale=y_scale),
                    x=alt.X("Density:Q", stack="center", axis=None),
                    color=alt.Color("Column:N", title="Column"),
                    column=alt.Column("Column:N", title="Column", sort=sorted_cols, spacing=10),
                )
                .properties(height=440)
            )
        elif plot_type == "Line + points (по Columns)":
            by_col = chart_long.copy()
            by_col["Column"] = by_col["Column"].astype(str)

            base = alt.Chart(by_col).encode(
                x=alt.X("Column:N", title="Column", sort=sorted_cols),
                y=alt.Y("Value:Q", title="Aggregated value", scale=y_scale),
                color=alt.Color("TimeLabel:N", title="Time"),
                detail=alt.Detail("File:N"),
                tooltip=["TimeLabel", "File", "Column", "Value", "Error"],
            )

            main_chart = base.mark_line(strokeWidth=3.0, point=alt.OverlayMarkDef(size=95, filled=True))
            use_error = error_method != "None"

            if use_error:
                error_chart = alt.Chart(by_col).mark_rule(opacity=0.55, strokeWidth=1.6).encode(
                    x=alt.X("Column:N", sort=sorted_cols),
                    y=alt.Y("y_low:Q"),
                    y2=alt.Y2("y_high:Q"),
                    color=alt.Color("TimeLabel:N"),
                    detail=alt.Detail("File:N"),
                )
                chart = (error_chart + main_chart).interactive().properties(height=450)
            else:
                chart = main_chart.interactive().properties(height=450)
        else:
            base = alt.Chart(chart_long).encode(
                x=alt.X("Time:Q", title="Time", scale=x_scale),
                y=alt.Y("Value:Q", title="Aggregated value", scale=y_scale),
                color=alt.Color("Column:N", title="Column"),
                tooltip=["Time", "File", "Column", "Value", "Error"],
            )

            main_chart = base.mark_line(strokeWidth=3.0, point=alt.OverlayMarkDef(size=95, filled=True))
            use_error = error_method != "None"

            if use_error:
                error_chart = alt.Chart(chart_long).mark_rule(opacity=0.55, strokeWidth=1.6).encode(
                    x=alt.X("Time:Q", scale=x_scale),
                    y=alt.Y("y_low:Q"),
                    y2=alt.Y2("y_high:Q"),
                    color=alt.Color("Column:N"),
                )
                whisker_low = alt.Chart(chart_long).mark_rule(opacity=0.7, strokeWidth=1.4).encode(
                    x=alt.X("x_left:Q", scale=x_scale),
                    x2=alt.X2("x_right:Q"),
                    y=alt.Y("y_low:Q"),
                    color=alt.Color("Column:N"),
                )
                whisker_high = alt.Chart(chart_long).mark_rule(opacity=0.7, strokeWidth=1.4).encode(
                    x=alt.X("x_left:Q", scale=x_scale),
                    x2=alt.X2("x_right:Q"),
                    y=alt.Y("y_high:Q"),
                    color=alt.Color("Column:N"),
                )
                chart = (error_chart + whisker_low + whisker_high + main_chart).interactive().properties(height=450)
            else:
                chart = main_chart.interactive().properties(height=450)

        st.altair_chart(chart, use_container_width=True)
        plot_display_df = plot_df[["Time", *available_plot_columns]].copy()
        plot_display_df["Time"] = plot_display_df["Time"].map(_format_time_label)
        st.dataframe(_gradient_style(plot_display_df, color_scale_mode, color_palette), width="stretch")

        st.subheader("Export")
        report_sheets = {
            "Aggregated table": combined_display_df,
            "Statistical analysis": stat_table,
            "T-test analysis": ttest_df,
            "Outlier log": removed_df_export,
            "Outlier statistics": outlier_stats_df,
            "Plot data": plot_display_df,
        }
        if not pairwise_matrix_df.empty:
            report_sheets["Pairwise T-test p-values"] = pairwise_matrix_df.reset_index()
            report_sheets["Pairwise notes"] = pd.DataFrame({"Note": [
                "Two-sided Welch t-tests; p-values are unadjusted for multiple comparisons.",
                "Only selected time/metric groups are compared. Each pair is calculated once.",
                "The diagonal and upper triangle are intentionally blank; p-values below 0.005 are shown as p < 0.005, other values are rounded to three decimals, and N/A denotes an undefined test.",
            ]})
        config = _capture_analysis_config()
        report_sheets.update(provenance_sheets(config))
        report_bytes = _build_excel_report(report_sheets)
        st.download_button(
            "Export report.xlsx",
            data=report_bytes,
            file_name="plate_master_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    finally:
        _render_settings_download(settings_download_slot)



def main() -> None:
    """Launch Streamlit when run as a script, or render the dashboard inside Streamlit."""
    if get_script_run_ctx is not None and get_script_run_ctx() is None:
        try:
            os.execvp("streamlit", ["streamlit", "run", str(Path(__file__).resolve())])
        except FileNotFoundError:
            print("Streamlit is not installed in PATH. Run: pip install -r requirements.txt")
        return
    run_dashboard()


if __name__ == "__main__":
    main()
