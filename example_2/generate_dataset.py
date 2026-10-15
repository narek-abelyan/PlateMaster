"""Regenerate the CC0 PlateMaster validation dataset and its reference files."""
from __future__ import annotations

from copy import deepcopy
from datetime import datetime
from pathlib import Path
import json
import math
import re
import sys
import zipfile

import pandas as pd
from scipy import stats

REPOSITORY = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPOSITORY))

from analysis_settings import (
    APP_VERSION, DEFAULT_SETTINGS, dumps_config, input_record, make_config,
)


ROOT = Path(__file__).resolve().parent
DATA = ROOT / "data"
ROWS = list("ABCDEFGH")
COLUMNS = [str(column) for column in range(1, 13)]
GROUPS = ["Control", *[f"Treatment_{index:02d}" for index in range(1, 12)]]
COLUMN_GROUPS = dict(zip(COLUMNS, GROUPS))
TIMES = [0, 5, 10, 15, 20]
BASE_VALUES = [float(index) for index in range(1, 9)]


def normalize_xlsx(path: Path) -> None:
    """Give generated XLSX ZIP members fixed timestamps for reproducible bytes."""
    temporary = path.with_suffix(".normalized.xlsx")
    with zipfile.ZipFile(path, "r") as source, zipfile.ZipFile(
        temporary, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9
    ) as target:
        for name in sorted(source.namelist()):
            info = zipfile.ZipInfo(name, date_time=(2020, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = source.getinfo(name).external_attr
            content = source.read(name)
            if name == "docProps/core.xml":
                content = re.sub(
                    rb"<dcterms:modified[^>]*>.*?</dcterms:modified>",
                    b'<dcterms:modified xmlns:dcterms="http://purl.org/dc/terms/" '
                    b'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" '
                    b'xsi:type="dcterms:W3CDTF">2020-01-01T00:00:00Z</dcterms:modified>',
                    content,
                )
            target.writestr(info, content)
    temporary.replace(path)


def plate_frame(index: int) -> pd.DataFrame:
    """Return one 8×12 plate with 12 groups and eight replicates per group."""
    data: dict[str, list[float] | list[str]] = {"Well": ROWS}
    for group_index, column in enumerate(COLUMNS):
        time_increment = index * (1 + group_index % 3)
        data[column] = [value + 2 * group_index + time_increment for value in BASE_VALUES]
    return pd.DataFrame(data)


def group_values(frame: pd.DataFrame, columns: list[str]) -> list[float]:
    """Flatten group columns in row order, matching PlateMaster plate groups."""
    return [float(value) for _, row in frame.iterrows() for value in row[columns]]


def expected_row(time: int, group: str, values: list[float]) -> dict:
    """Calculate documented reference statistics from their standard formulas."""
    series = pd.Series(values, dtype=float)
    n = len(series)
    sample_sd = float(series.std(ddof=1))
    sem = sample_sd / math.sqrt(n)
    ci_half = float(stats.t.ppf(0.975, n - 1) * sem)
    one_sample = stats.ttest_1samp(series.to_numpy(), popmean=0.0)
    return {
        "time_minutes": time,
        "group": group,
        "n": n,
        "mean": float(series.mean()),
        "median": float(series.median()),
        "sample_sd": sample_sd,
        "sem": sem,
        "ci95_half_width": ci_half,
        "ci95_lower": float(series.mean() - ci_half),
        "ci95_upper": float(series.mean() + ci_half),
        "minimum": float(series.min()),
        "maximum": float(series.max()),
        "one_sample_t": float(one_sample.statistic),
        "one_sample_p_two_sided": float(one_sample.pvalue),
    }


def main() -> None:
    DATA.mkdir(exist_ok=True)
    frames: dict[int, pd.DataFrame] = {}
    input_records = []
    expected_rows = []

    for index, time in enumerate(TIMES):
        frame = plate_frame(index)
        frames[time] = frame
        path = DATA / f"validation_plate_{time}.xlsx"
        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            frame.to_excel(writer, sheet_name="Plate", index=False)
            writer.book.properties.creator = "PlateMaster validation dataset"
            writer.book.properties.created = datetime(2020, 1, 1)
            writer.book.properties.modified = datetime(2020, 1, 1)
        normalize_xlsx(path)

        # Fingerprint the values as they are actually read back from the XLSX,
        # because Excel may preserve integer-valued synthetic cells as integers.
        parsed = pd.read_excel(path, sheet_name=0, header=0)
        parsed.columns = ["row_id", *[str(column).strip() for column in parsed.columns[1:]]]
        parsed["row_id"] = parsed["row_id"].astype(str).str.strip()
        parsed.attrs["import_options"] = {
            "sheet": 0,
            "header_row": 1,
            "allow_custom": False,
            "plate_shape": None,
        }
        input_records.append(input_record(path.name, path.read_bytes(), parsed, float(time)))
        for column, group in COLUMN_GROUPS.items():
            expected_rows.append(expected_row(time, group, group_values(frame, [column])))

    pd.DataFrame(expected_rows).to_csv(ROOT / "expected_group_results.csv", index=False, float_format="%.15g")

    a = group_values(frames[0], ["1"])
    b = group_values(frames[20], ["12"])
    welch = stats.ttest_ind(a, b, equal_var=False)
    pd.DataFrame(
        [{
            "group_a": "Control",
            "time_a_minutes": 0,
            "group_b": "Treatment_11",
            "time_b_minutes": 20,
            "n_a": len(a),
            "n_b": len(b),
            "mean_a": sum(a) / len(a),
            "mean_b": sum(b) / len(b),
            "welch_t": float(welch.statistic),
            "welch_df": float(welch.df),
            "p_two_sided": float(welch.pvalue),
        }]
    ).to_csv(ROOT / "expected_welch_comparison.csv", index=False, float_format="%.15g")

    plate_map = [
        {"row_id": row, "column": column, "group": COLUMN_GROUPS[column]}
        for row in ROWS
        for column in COLUMNS
    ]
    pd.DataFrame(plate_map).to_csv(ROOT / "plate_map.csv", index=False)

    settings = deepcopy(DEFAULT_SETTINGS)
    settings.update(
        agg_layout_mode="By plate groups (template)",
        selected_rows=ROWS,
        selected_columns=COLUMNS,
        agg_name="Mean",
        outlier_method="None",
        plot_columns=GROUPS,
        error_method="95% CI",
        stat_time=10.0,
        stat_metric="Treatment_05",
        ttest_time_a=0.0,
        ttest_metric_a="Control",
        ttest_time_b=20.0,
        ttest_metric_b="Treatment_11",
        pairwise_enabled=False,
        pairwise_groups=[],
    )
    software = {
        "application_version": APP_VERSION,
        "dataset": "PlateMaster example_2 synthetic validation set",
        "reference_method": "Student-t intervals and SciPy Welch t-test",
        "license": "CC0-1.0",
    }
    config = make_config(settings, plate_map, input_records, software)
    (ROOT / "analysis_settings.json").write_bytes(dumps_config(config))

    manifest = {
        "format": "platemaster.validation-manifest",
        "version": 1,
        "license": "CC0-1.0",
        "time_unit": "minutes",
        "files": [{"file": item["file"], "sha256": item["sha256"]} for item in input_records],
    }
    (ROOT / "manifest.json").write_text(json.dumps(manifest, indent=2) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
